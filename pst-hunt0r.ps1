<#
.SYNOPSIS
    pst-hunt0r by Benjamin Iheukumere | b.iheukumere@safelink-it.com
    Scans ZIP archives for Outlook PST files, searches emails for a target domain,
    supports resume/checkpointing, and optionally exports matching emails.

.DESCRIPTION
    - Recursively scans a ZIP root folder
    - Pre-analyzes ZIP archives for contained PST entries
    - Supports resume on PST level via checkpoint.json
    - Extracts ZIP archives into temporary folders
    - Opens PST files through Outlook/MAPI
    - Recursively traverses all folders and mail items
    - Searches sender and recipient addresses for a target domain
    - Commits results per PST after successful processing
    - Optionally exports matching emails into found_mails as .msg or .eml
    - Writes log output and shows progress with elapsed time and ETA

.NOTES
    - PowerShell 5.1 compatible
    - Outlook Desktop with a working profile is required
    - ZIP archives protected by passwords are not supported
    - Resume/checkpointing is done on PST level, not on single-mail level
#>

[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ------------------------------------------------------------
# Configuration - edit these values before running the script
# ------------------------------------------------------------
$Config = [ordered]@{
    ZipRoot            = 'Z:\path\to\your\zippes\pst\files'
    TargetDomain       = '@exmaple.com'

    TempRoot           = 'F:\path\to\temp\folder'
    OutputCsv          = ''   # If empty: <TempRoot>\pst_hunt0r_results.csv
    StateRoot          = ''   # If empty: <TempRoot>\state
    FoundMailsRoot     = ''   # If empty: <TempRoot>\found_mails

    ExportMatchedMails = $true
    ExportFormat       = 'msg'   # msg | eml

    EnableResume       = $true
    StrictStateSafety  = $true
    StartFresh         = $false

    KeepExtractedFiles = $false

    DeleteRetryCount   = 10
    DeleteRetryDelayMs = 750
}

# ------------------------------------------------------------
# Derived paths
# ------------------------------------------------------------
if ([string]::IsNullOrWhiteSpace($Config.StateRoot)) {
    $Config.StateRoot = Join-Path -Path $Config.TempRoot -ChildPath 'state'
}

if ([string]::IsNullOrWhiteSpace($Config.OutputCsv)) {
    $Config.OutputCsv = Join-Path -Path $Config.TempRoot -ChildPath 'pst_hunt0r_results.csv'
}

if ([string]::IsNullOrWhiteSpace($Config.FoundMailsRoot)) {
    $Config.FoundMailsRoot = Join-Path -Path $Config.TempRoot -ChildPath 'found_mails'
}

$script:StageRoot       = Join-Path -Path $Config.StateRoot -ChildPath 'stage'
$script:CheckpointPath  = Join-Path -Path $Config.StateRoot -ChildPath 'checkpoint.json'
$script:LogPath         = Join-Path -Path $Config.StateRoot -ChildPath 'pst_hunt0r.log'

# ------------------------------------------------------------
# Script state
# ------------------------------------------------------------
$script:RunWatch                = [System.Diagnostics.Stopwatch]::StartNew()
$script:TotalZipCount           = 0
$script:ProcessedZipCount       = 0
$script:OverallTotalPst         = 0
$script:OverallProcessedPst     = 0
$script:CommittedHitCount       = 0

$script:CurrentZipName          = ''
$script:CurrentZipIndex         = 0
$script:CurrentZipTotalPst      = 0
$script:CurrentZipProcessedPst  = 0
$script:CurrentZipWatch         = $null
$script:CurrentPstName          = ''
$script:CurrentPhase            = 'Initialization'

$script:TargetDomainNormalized  = $null
$script:PrSmtpAddressUri        = "http://schemas.microsoft.com/mapi/proptag/0x39FE001F"

$script:CheckpointState         = $null

$script:CsvColumns = @(
    'ZipFile',
    'PstRelativePath',
    'FolderPath',
    'Subject',
    'Sender',
    'Recipients',
    'SentOn',
    'ReceivedTime',
    'SenderHit',
    'RecipientHit',
    'ExportedFilePath',
    'ExportStatus',
    'WorkKey'
)

# ------------------------------------------------------------
# Utility functions
# ------------------------------------------------------------

function Ensure-Directory {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -Path $Path -ItemType Directory -Force | Out-Null
    }
}

function Write-Log {
    param(
        [ValidateSet('INFO','WARN','ERROR')]
        [string]$Level,

        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    $line = "[{0}] [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message

    try {
        $logParent = Split-Path -Path $script:LogPath -Parent
        if (-not [string]::IsNullOrWhiteSpace($logParent)) {
            Ensure-Directory -Path $logParent
        }
        Add-Content -LiteralPath $script:LogPath -Value $line -Encoding UTF8
    }
    catch {
        # Logging must never stop the script
    }

    switch ($Level) {
        'INFO'  { Write-Host $line -ForegroundColor Cyan }
        'WARN'  { Write-Warning $Message }
        'ERROR' { Write-Error $Message }
    }
}

function Safe-String {
    param($Value)

    if ($null -eq $Value) {
        return ''
    }

    return [string]$Value
}

function Normalize-Email {
    param([AllowNull()][string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $null
    }

    $v = $Value.Trim().Trim('<', '>').ToLowerInvariant()

    if ($v.StartsWith('smtp:')) {
        $v = $v.Substring(5)
    }

    return $v
}

function Get-DateString {
    param($Value)

    try {
        if ($null -eq $Value) {
            return ''
        }

        return ([datetime]$Value).ToString('yyyy-MM-dd HH:mm:ss')
    }
    catch {
        return ''
    }
}

function Convert-ToCsvField {
    param($Value)

    $text = Safe-String $Value
    $text = $text.Replace('"', '""')
    return '"' + $text + '"'
}

function Get-ExpectedCsvHeaderLine {
    $quotedHeaderFields = $script:CsvColumns | ForEach-Object { Convert-ToCsvField $_ }
    return ($quotedHeaderFields -join ';')
}

function Write-CsvHeaderToWriter {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.StreamWriter]$Writer
    )

    $Writer.WriteLine((Get-ExpectedCsvHeaderLine))
}

function Initialize-MasterCsv {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $parent = Split-Path -Path $Path -Parent
    if (-not [string]::IsNullOrWhiteSpace($parent)) {
        Ensure-Directory -Path $parent
    }

    if (-not (Test-Path -LiteralPath $Path)) {
        $utf8Bom = New-Object System.Text.UTF8Encoding($true)
        $writer = New-Object System.IO.StreamWriter($Path, $false, $utf8Bom)
        try {
            Write-CsvHeaderToWriter -Writer $writer
        }
        finally {
            $writer.Flush()
            $writer.Dispose()
        }
        return
    }

    $firstLine = ''
    try {
        $firstLine = Get-Content -LiteralPath $Path -TotalCount 1 -ErrorAction Stop
    }
    catch {
        $firstLine = ''
    }

    if ([string]::IsNullOrWhiteSpace($firstLine)) {
        $utf8Bom = New-Object System.Text.UTF8Encoding($true)
        $writer = New-Object System.IO.StreamWriter($Path, $false, $utf8Bom)
        try {
            Write-CsvHeaderToWriter -Writer $writer
        }
        finally {
            $writer.Flush()
            $writer.Dispose()
        }
        return
    }

    $expected = Get-ExpectedCsvHeaderLine
    if ($firstLine -ne $expected) {
        throw "Master CSV header does not match the expected header. Refusing to continue with a potentially incompatible file: $Path"
    }
}

function Assert-StateConsistency {
    param(
        [Parameter(Mandatory = $true)]
        [string]$OutputCsv,

        [Parameter(Mandatory = $true)]
        [string]$CheckpointPath,

        [bool]$EnableResume,
        [bool]$StrictStateSafety
    )

    if (-not $EnableResume) {
        return
    }

    $csvExists = Test-Path -LiteralPath $OutputCsv
    $checkpointExists = Test-Path -LiteralPath $CheckpointPath

    if ($StrictStateSafety) {
        if ($csvExists -and -not $checkpointExists) {
            throw "State safety check failed: master CSV exists but checkpoint.json is missing. This may lead to duplicate results. Either restore the checkpoint or start fresh."
        }

        if ($checkpointExists -and -not $csvExists) {
            throw "State safety check failed: checkpoint.json exists but master CSV is missing. This indicates inconsistent resume state. Either restore the CSV or start fresh."
        }
    }
    else {
        if ($csvExists -and -not $checkpointExists) {
            Write-Log -Level 'WARN' -Message "Master CSV exists but checkpoint.json is missing. Resume safety is reduced and duplicates may occur."
        }

        if ($checkpointExists -and -not $csvExists) {
            Write-Log -Level 'WARN' -Message "checkpoint.json exists but master CSV is missing. Resume state is inconsistent."
        }
    }
}

function Load-CheckpointState {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $state = @{
        Completed = @{}
    }

    if (-not (Test-Path -LiteralPath $Path)) {
        return $state
    }

    try {
        $raw = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
        if ([string]::IsNullOrWhiteSpace($raw)) {
            return $state
        }

        $json = $raw | ConvertFrom-Json -ErrorAction Stop

        if ($null -ne $json.completedPstKeys) {
            foreach ($key in $json.completedPstKeys) {
                if (-not [string]::IsNullOrWhiteSpace([string]$key)) {
                    $state.Completed[[string]$key] = $true
                }
            }
        }
    }
    catch {
        throw "Failed to load checkpoint state from '$Path'. Refusing to continue because resume integrity would be uncertain. Error: $($_.Exception.Message)"
    }

    return $state
}

function Save-CheckpointState {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [hashtable]$State
    )

    $parent = Split-Path -Path $Path -Parent
    if (-not [string]::IsNullOrWhiteSpace($parent)) {
        Ensure-Directory -Path $parent
    }

    $payload = [ordered]@{
        updatedUtc       = (Get-Date).ToUniversalTime().ToString('o')
        completedCount   = $State.Completed.Count
        completedPstKeys = @($State.Completed.Keys | Sort-Object)
    }

    $json = $payload | ConvertTo-Json -Depth 5
    $tmp  = "$Path.tmp"

    try {
        Set-Content -LiteralPath $tmp -Value $json -Encoding UTF8
        Move-Item -LiteralPath $tmp -Destination $Path -Force
    }
    finally {
        if (Test-Path -LiteralPath $tmp) {
            Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
        }
    }
}

function New-ShortHash {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Text,

        [int]$Length = 8
    )

    $sha1 = [System.Security.Cryptography.SHA1]::Create()
    try {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($Text)
        $hash = $sha1.ComputeHash($bytes)
        $hex = ([System.BitConverter]::ToString($hash)).Replace('-', '').ToLowerInvariant()

        if ($Length -gt 0 -and $Length -lt $hex.Length) {
            return $hex.Substring(0, $Length)
        }

        return $hex
    }
    finally {
        $sha1.Dispose()
    }
}

function Normalize-RelativePath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Value
    )

    $v = $Value -replace '/', '\'
    $v = $v.Trim()
    $v = $v.Trim('\')
    return $v
}

function Get-RelativePathSafe {
    param(
        [Parameter(Mandatory = $true)]
        [string]$BasePath,

        [Parameter(Mandatory = $true)]
        [string]$ChildPath
    )

    $baseResolved  = (Resolve-Path -LiteralPath $BasePath).Path
    $childResolved = (Resolve-Path -LiteralPath $ChildPath).Path

    if (-not $baseResolved.EndsWith('\')) {
        $baseResolved += '\'
    }

    $baseUri  = New-Object System.Uri($baseResolved)
    $childUri = New-Object System.Uri($childResolved)
    $relative = $baseUri.MakeRelativeUri($childUri).ToString()

    return (Normalize-RelativePath -Value ([System.Uri]::UnescapeDataString($relative)))
}

function New-SafeFileName {
    param(
        [string]$Name,
        [int]$MaxLength = 120
    )

    if ([string]::IsNullOrWhiteSpace($Name)) {
        $Name = 'unnamed'
    }

    foreach ($c in [System.IO.Path]::GetInvalidFileNameChars()) {
        $Name = $Name.Replace($c, '_')
    }

    $Name = ($Name -replace '\s+', ' ').Trim().Trim('.')

    if ($Name.Length -gt $MaxLength) {
        $Name = $Name.Substring(0, $MaxLength)
    }

    if ([string]::IsNullOrWhiteSpace($Name)) {
        $Name = 'unnamed'
    }

    return $Name
}

function Release-ComObject {
    param($Obj)

    if ($null -ne $Obj -and [System.Runtime.InteropServices.Marshal]::IsComObject($Obj)) {
        try {
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($Obj)
        }
        catch {
            # Intentionally ignored
        }
    }
}

function Invoke-ComReleaseCycle {
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

function Remove-DirectoryRobust {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [int]$RetryCount = 10,
        [int]$RetryDelayMs = 750,

        [switch]$Quiet
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        return $true
    }

    $lastError = $null

    for ($attempt = 1; $attempt -le $RetryCount; $attempt++) {
        try {
            Remove-Item -LiteralPath $Path -Recurse -Force -ErrorAction Stop

            if (-not (Test-Path -LiteralPath $Path)) {
                return $true
            }
        }
        catch {
            $lastError = $_.Exception.Message
        }

        Invoke-ComReleaseCycle
        Start-Sleep -Milliseconds $RetryDelayMs
    }

    if (-not $Quiet) {
        Write-Log -Level 'WARN' -Message ("Temporary directory could not be deleted after {0} attempts: {1}. Last error: {2}" -f $RetryCount, $Path, $lastError)
    }

    return $false
}

function Remove-FileRobust {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [int]$RetryCount = 10,
        [int]$RetryDelayMs = 500
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        return $true
    }

    $lastError = $null

    for ($attempt = 1; $attempt -le $RetryCount; $attempt++) {
        try {
            Remove-Item -LiteralPath $Path -Force -ErrorAction Stop

            if (-not (Test-Path -LiteralPath $Path)) {
                return $true
            }
        }
        catch {
            $lastError = $_.Exception.Message
        }

        Invoke-ComReleaseCycle
        Start-Sleep -Milliseconds $RetryDelayMs
    }

    Write-Log -Level 'WARN' -Message ("File could not be deleted after {0} attempts: {1}. Last error: {2}" -f $RetryCount, $Path, $lastError)
    return $false
}

function Format-Duration {
    param($Duration)

    if ($null -eq $Duration) {
        return 'n/a'
    }

    if ($Duration -isnot [TimeSpan]) {
        return 'n/a'
    }

    if ($Duration.TotalSeconds -lt 0) {
        return 'n/a'
    }

    return ('{0:00}:{1:00}:{2:00}' -f [math]::Floor($Duration.TotalHours), $Duration.Minutes, $Duration.Seconds)
}

function Get-EtaByUnits {
    param(
        [double]$CompletedUnits,
        [double]$TotalUnits,
        [TimeSpan]$Elapsed
    )

    if ($TotalUnits -le 0 -or $CompletedUnits -le 0) {
        return $null
    }

    if ($CompletedUnits -ge $TotalUnits) {
        return [TimeSpan]::Zero
    }

    $secondsPerUnit = $Elapsed.TotalSeconds / $CompletedUnits
    $remainingUnits = $TotalUnits - $CompletedUnits

    return [TimeSpan]::FromSeconds([math]::Round($secondsPerUnit * $remainingUnits))
}

function Update-ProgressBars {
    param(
        [string]$OverallActivity = 'Overall Progress',
        [string]$ZipActivity = 'Current ZIP File'
    )

    $elapsedOverall = $script:RunWatch.Elapsed
    $overallEta     = $null
    $overallPercent = 0

    if ($script:OverallTotalPst -gt 0) {
        $overallPercent = [int][math]::Min(100, [math]::Round(($script:OverallProcessedPst / $script:OverallTotalPst) * 100, 0))
        $overallEta = Get-EtaByUnits -CompletedUnits $script:OverallProcessedPst -TotalUnits $script:OverallTotalPst -Elapsed $elapsedOverall
    }
    elseif ($script:TotalZipCount -gt 0) {
        $overallPercent = [int][math]::Min(100, [math]::Round(($script:ProcessedZipCount / $script:TotalZipCount) * 100, 0))
        $overallEta = Get-EtaByUnits -CompletedUnits $script:ProcessedZipCount -TotalUnits $script:TotalZipCount -Elapsed $elapsedOverall
    }

    $overallStatus = if ($script:OverallTotalPst -gt 0) {
        "ZIP $($script:ProcessedZipCount)/$($script:TotalZipCount) | Pending PST $($script:OverallProcessedPst)/$($script:OverallTotalPst) | Committed hits $($script:CommittedHitCount) | Elapsed $(Format-Duration $elapsedOverall) | ETA $(Format-Duration $overallEta)"
    }
    else {
        "ZIP $($script:ProcessedZipCount)/$($script:TotalZipCount) | Committed hits $($script:CommittedHitCount) | Elapsed $(Format-Duration $elapsedOverall) | ETA $(Format-Duration $overallEta)"
    }

    $overallOperation = if ($script:CurrentZipName) {
        "$($script:CurrentPhase) | $($script:CurrentZipName)"
    }
    else {
        $script:CurrentPhase
    }

    Write-Progress -Id 0 -Activity $OverallActivity -Status $overallStatus -CurrentOperation $overallOperation -PercentComplete $overallPercent

    if ($script:CurrentZipName) {
        $zipElapsed = if ($null -ne $script:CurrentZipWatch) { $script:CurrentZipWatch.Elapsed } else { [TimeSpan]::Zero }
        $zipEta     = $null
        $zipPercent = 0

        if ($script:CurrentZipTotalPst -gt 0) {
            $zipPercent = [int][math]::Min(100, [math]::Round(($script:CurrentZipProcessedPst / $script:CurrentZipTotalPst) * 100, 0))
            $zipEta = Get-EtaByUnits -CompletedUnits $script:CurrentZipProcessedPst -TotalUnits $script:CurrentZipTotalPst -Elapsed $zipElapsed
        }

        $zipStatus = if ($script:CurrentZipTotalPst -gt 0) {
            "ZIP $($script:CurrentZipIndex)/$($script:TotalZipCount) | Pending PST $($script:CurrentZipProcessedPst)/$($script:CurrentZipTotalPst) | Committed hits $($script:CommittedHitCount) | Elapsed $(Format-Duration $zipElapsed) | ETA $(Format-Duration $zipEta)"
        }
        else {
            "ZIP $($script:CurrentZipIndex)/$($script:TotalZipCount) | No pending PST files | Elapsed $(Format-Duration $zipElapsed)"
        }

        $zipOperation = if ($script:CurrentPstName) {
            "$($script:CurrentPhase) | $($script:CurrentPstName)"
        }
        else {
            $script:CurrentPhase
        }

        Write-Progress -Id 1 -ParentId 0 -Activity $ZipActivity -Status $zipStatus -CurrentOperation $zipOperation -PercentComplete $zipPercent
    }
    else {
        Write-Progress -Id 1 -ParentId 0 -Activity $ZipActivity -Completed
    }
}

function Complete-ProgressBars {
    Write-Progress -Id 1 -ParentId 0 -Activity 'Current ZIP File' -Completed
    Write-Progress -Id 0 -Activity 'Overall Progress' -Completed
}

function Test-IsTargetDomain {
    param([AllowNull()][string]$Value)

    $v = Normalize-Email -Value $Value
    if (-not $v) {
        return $false
    }

    return $v.EndsWith($script:TargetDomainNormalized)
}

function Get-SmtpFromAddressEntry {
    param($AddressEntry)

    if ($null -eq $AddressEntry) {
        return $null
    }

    try {
        $pa = $AddressEntry.PropertyAccessor
        $smtp = $pa.GetProperty($script:PrSmtpAddressUri)
        if ($smtp) {
            return (Normalize-Email -Value $smtp)
        }
    }
    catch {
        # Next fallback
    }

    try {
        $exUser = $AddressEntry.GetExchangeUser()
        if ($exUser -and $exUser.PrimarySmtpAddress) {
            return (Normalize-Email -Value $exUser.PrimarySmtpAddress)
        }
    }
    catch {
        # Next fallback
    }

    try {
        if ($AddressEntry.Address) {
            return (Normalize-Email -Value $AddressEntry.Address)
        }
    }
    catch {
        # Ignored
    }

    return $null
}

function Get-SenderSmtp {
    param($MailItem)

    try {
        $sender = Normalize-Email -Value $MailItem.SenderEmailAddress
        if ($sender -and $sender.Contains('@')) {
            return $sender
        }
    }
    catch {
        # Next fallback
    }

    try {
        return (Get-SmtpFromAddressEntry -AddressEntry $MailItem.Sender)
    }
    catch {
        return $null
    }
}

function Get-RecipientSmtps {
    param($MailItem)

    $list = New-Object 'System.Collections.Generic.List[string]'

    try {
        $count = $MailItem.Recipients.Count
    }
    catch {
        return @()
    }

    for ($i = 1; $i -le $count; $i++) {
        $recipient = $null

        try {
            $recipient = $MailItem.Recipients.Item($i)

            $address = $null
            try {
                $address = Get-SmtpFromAddressEntry -AddressEntry $recipient.AddressEntry
            }
            catch {
                # Fallback below
            }

            if (-not $address) {
                try {
                    $address = Normalize-Email -Value $recipient.Address
                }
                catch {
                    $address = $null
                }
            }

            if ($address -and -not $list.Contains($address)) {
                [void]$list.Add($address)
            }
        }
        catch {
            # Recipient issues are not fatal
        }
        finally {
            Release-ComObject -Obj $recipient
        }
    }

    return $list.ToArray()
}

function Get-FolderPathSafe {
    param($Folder)

    try {
        return [string]$Folder.FolderPath
    }
    catch {
        try {
            return [string]$Folder.Name
        }
        catch {
            return ''
        }
    }
}

function Get-StoreByFilePath {
    param(
        $Namespace,
        [string]$PstPath
    )

    $stores = $null
    try {
        $stores = $Namespace.Stores
        $count = $stores.Count

        for ($i = 1; $i -le $count; $i++) {
            $store   = $null
            $isMatch = $false

            try {
                $store = $stores.Item($i)
                if ($store -and $store.FilePath -and ([string]::Equals($store.FilePath, $PstPath, [System.StringComparison]::OrdinalIgnoreCase))) {
                    $isMatch = $true
                    return $store
                }
            }
            catch {
                # Ignored
            }
            finally {
                if ($null -ne $store -and -not $isMatch) {
                    Release-ComObject -Obj $store
                }
            }
        }
    }
    finally {
        Release-ComObject -Obj $stores
    }

    return $null
}

function Get-PstWorkKey {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$ZipFile,

        [Parameter(Mandatory = $true)]
        [string]$PstRelativePath
    )

    $raw = "{0}|{1}|{2}|{3}" -f `
        $ZipFile.FullName.ToLowerInvariant(), `
        $ZipFile.Length, `
        $ZipFile.LastWriteTimeUtc.Ticks, `
        (Normalize-RelativePath -Value $PstRelativePath).ToLowerInvariant()

    return (New-ShortHash -Text $raw -Length 40)
}

function Get-ZipExportFolderName {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$ZipFile
    )

    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($ZipFile.Name)
    $safeBase = New-SafeFileName -Name $baseName -MaxLength 70
    $suffix   = New-ShortHash -Text $ZipFile.FullName -Length 8

    return ("{0}_{1}" -f $safeBase, $suffix)
}

function Get-PstExportFolderName {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PstRelativePath,

        [Parameter(Mandatory = $true)]
        [string]$WorkKey
    )

    $safeBase = New-SafeFileName -Name ((Normalize-RelativePath -Value $PstRelativePath) -replace '\\', '__') -MaxLength 80
    $suffix   = $WorkKey.Substring(0, 8)

    return ("{0}_{1}" -f $safeBase, $suffix)
}

function Get-UniqueFilePath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Directory,

        [Parameter(Mandatory = $true)]
        [string]$BaseName,

        [Parameter(Mandatory = $true)]
        [string]$Extension
    )

    $candidate = Join-Path -Path $Directory -ChildPath ($BaseName + $Extension)
    $counter = 1

    while (Test-Path -LiteralPath $candidate) {
        $candidate = Join-Path -Path $Directory -ChildPath ("{0}_{1}{2}" -f $BaseName, $counter, $Extension)
        $counter++
    }

    return $candidate
}

function Initialize-StageContext {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$ZipFile,

        [Parameter(Mandatory = $true)]
        [string]$PstRelativePath,

        [Parameter(Mandatory = $true)]
        [string]$WorkKey
    )

    $stagePath = Join-Path -Path $script:StageRoot -ChildPath $WorkKey

    if (Test-Path -LiteralPath $stagePath) {
        Write-Log -Level 'WARN' -Message "Found stale stage directory for work key $WorkKey. Removing it before reprocessing."
        [void](Remove-DirectoryRobust -Path $stagePath -RetryCount $Config.DeleteRetryCount -RetryDelayMs $Config.DeleteRetryDelayMs)
    }

    Ensure-Directory -Path $stagePath

    $stageCsvPath = Join-Path -Path $stagePath -ChildPath 'stage_results.csv'
    $stageMailRoot = Join-Path -Path $stagePath -ChildPath 'mails'

    $utf8Bom = New-Object System.Text.UTF8Encoding($true)
    $writer  = New-Object System.IO.StreamWriter($stageCsvPath, $false, $utf8Bom)
    $writer.AutoFlush = $true
    Write-CsvHeaderToWriter -Writer $writer

    $zipExportFolder = Get-ZipExportFolderName -ZipFile $ZipFile
    $pstExportFolder = Get-PstExportFolderName -PstRelativePath $PstRelativePath -WorkKey $WorkKey
    $relativeMailRoot = Join-Path -Path $zipExportFolder -ChildPath $pstExportFolder

    if ($Config.ExportMatchedMails) {
        Ensure-Directory -Path (Join-Path -Path $stageMailRoot -ChildPath $relativeMailRoot)
    }

    return @{
        WorkKey            = $WorkKey
        StagePath          = $stagePath
        StageCsvPath       = $stageCsvPath
        StageMailRoot      = $stageMailRoot
        RelativeMailRoot   = $relativeMailRoot
        CsvWriter          = $writer
        HitCount           = 0
        MailCounter        = 0
    }
}

function Close-StageContext {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$StageContext
    )

    if ($null -ne $StageContext.CsvWriter) {
        try {
            $StageContext.CsvWriter.Flush()
            $StageContext.CsvWriter.Dispose()
        }
        catch {
            # Intentionally ignored
        }
        finally {
            $StageContext.CsvWriter = $null
        }
    }
}

function Write-StageCsvRow {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$StageContext,

        [Parameter(Mandatory = $true)]
        [hashtable]$Row
    )

    $line = @(
        Convert-ToCsvField $Row.ZipFile
        Convert-ToCsvField $Row.PstRelativePath
        Convert-ToCsvField $Row.FolderPath
        Convert-ToCsvField $Row.Subject
        Convert-ToCsvField $Row.Sender
        Convert-ToCsvField $Row.Recipients
        Convert-ToCsvField $Row.SentOn
        Convert-ToCsvField $Row.ReceivedTime
        Convert-ToCsvField ([string]$Row.SenderHit)
        Convert-ToCsvField ([string]$Row.RecipientHit)
        Convert-ToCsvField $Row.ExportedFilePath
        Convert-ToCsvField $Row.ExportStatus
        Convert-ToCsvField $Row.WorkKey
    ) -join ';'

    $StageContext.CsvWriter.WriteLine($line)
    $StageContext.HitCount++
}

function Append-StageCsvToMaster {
    param(
        [Parameter(Mandatory = $true)]
        [string]$StageCsvPath,

        [Parameter(Mandatory = $true)]
        [string]$MasterCsvPath
    )

    if (-not (Test-Path -LiteralPath $StageCsvPath)) {
        return
    }

    $utf8Bom = New-Object System.Text.UTF8Encoding($true)

    $reader = New-Object System.IO.StreamReader($StageCsvPath, $true)
    try {
        $writer = New-Object System.IO.StreamWriter($MasterCsvPath, $true, $utf8Bom)
        try {
            $isFirstLine = $true
            while (($line = $reader.ReadLine()) -ne $null) {
                if ($isFirstLine) {
                    $isFirstLine = $false
                    continue
                }

                $writer.WriteLine($line)
            }
        }
        finally {
            $writer.Flush()
            $writer.Dispose()
        }
    }
    finally {
        $reader.Dispose()
    }
}

function Save-FoundMailToStage {
    param(
        [Parameter(Mandatory = $true)]
        $MailItem,

        [Parameter(Mandatory = $true)]
        [hashtable]$StageContext
    )

    if (-not $Config.ExportMatchedMails) {
        return @{
            RelativePath = ''
            Status       = 'mail_export_disabled'
        }
    }

    $StageContext.MailCounter++

    $targetDir = Join-Path -Path $StageContext.StageMailRoot -ChildPath $StageContext.RelativeMailRoot
    Ensure-Directory -Path $targetDir

    try {
        $timestamp = ([datetime]$MailItem.ReceivedTime).ToString('yyyyMMdd_HHmmss')
    }
    catch {
        try {
            $timestamp = ([datetime]$MailItem.SentOn).ToString('yyyyMMdd_HHmmss')
        }
        catch {
            $timestamp = (Get-Date).ToString('yyyyMMdd_HHmmss')
        }
    }

    $subjectSafe = New-SafeFileName -Name $MailItem.Subject -MaxLength 80
    $baseName    = "{0}_{1}_{2:000000}" -f $timestamp, $subjectSafe, $StageContext.MailCounter

    switch ($Config.ExportFormat.ToLowerInvariant()) {
        'eml' {
            $extension = '.eml'
            $saveType  = 102 # olRFC822
        }
        default {
            $extension = '.msg'
            $saveType  = 9   # olMSGUnicode
        }
    }

    $filePath = Get-UniqueFilePath -Directory $targetDir -BaseName $baseName -Extension $extension

    try {
        $MailItem.SaveAs($filePath, $saveType)
        $relative = Get-RelativePathSafe -BasePath $StageContext.StageMailRoot -ChildPath $filePath

        return @{
            RelativePath = $relative
            Status       = 'saved'
        }
    }
    catch {
        return @{
            RelativePath = ''
            Status       = ('save_failed: ' + $_.Exception.Message)
        }
    }
}

function Move-StagedMailFilesToFinal {
    param(
        [Parameter(Mandatory = $true)]
        [string]$StageMailRoot,

        [Parameter(Mandatory = $true)]
        [string]$FinalRoot
    )

    if (-not (Test-Path -LiteralPath $StageMailRoot)) {
        return
    }

    $files = Get-ChildItem -Path $StageMailRoot -Recurse -File -ErrorAction Stop

    foreach ($file in $files) {
        $relative = Get-RelativePathSafe -BasePath $StageMailRoot -ChildPath $file.FullName
        $destinationPath = Join-Path -Path $FinalRoot -ChildPath $relative
        $destinationDir  = Split-Path -Path $destinationPath -Parent

        Ensure-Directory -Path $destinationDir

        if (Test-Path -LiteralPath $destinationPath) {
            $destinationPath = Get-UniqueFilePath `
                -Directory $destinationDir `
                -BaseName ([System.IO.Path]::GetFileNameWithoutExtension($destinationPath)) `
                -Extension ([System.IO.Path]::GetExtension($destinationPath))
        }

        Move-Item -LiteralPath $file.FullName -Destination $destinationPath -Force
    }
}

function Commit-PstStage {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$StageContext,

        [Parameter(Mandatory = $true)]
        [string]$MasterCsvPath,

        [Parameter(Mandatory = $true)]
        [string]$FoundMailsRoot,

        [Parameter(Mandatory = $true)]
        [hashtable]$CheckpointState,

        [Parameter(Mandatory = $true)]
        [string]$CheckpointPath
    )

    Close-StageContext -StageContext $StageContext

    if ($StageContext.HitCount -gt 0) {
        Append-StageCsvToMaster -StageCsvPath $StageContext.StageCsvPath -MasterCsvPath $MasterCsvPath

        if ($Config.ExportMatchedMails) {
            Move-StagedMailFilesToFinal -StageMailRoot $StageContext.StageMailRoot -FinalRoot $FoundMailsRoot
        }

        $script:CommittedHitCount += $StageContext.HitCount
    }

    $CheckpointState.Completed[$StageContext.WorkKey] = $true
    Save-CheckpointState -Path $CheckpointPath -State $CheckpointState

    [void](Remove-DirectoryRobust -Path $StageContext.StagePath -RetryCount $Config.DeleteRetryCount -RetryDelayMs $Config.DeleteRetryDelayMs -Quiet)
}

function Abandon-PstStage {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$StageContext
    )

    Close-StageContext -StageContext $StageContext
    [void](Remove-DirectoryRobust -Path $StageContext.StagePath -RetryCount $Config.DeleteRetryCount -RetryDelayMs $Config.DeleteRetryDelayMs -Quiet)
}

function Get-ZipInventory {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$ZipFile,

        [Parameter(Mandatory = $true)]
        [hashtable]$CheckpointState
    )

    $archive = $null
    $pendingEntries = New-Object 'System.Collections.Generic.List[object]'
    $pendingLookup  = @{}
    $totalPstEntries = 0
    $completedSkipped = 0

    try {
        $archive = [System.IO.Compression.ZipFile]::OpenRead($ZipFile.FullName)

        foreach ($entry in $archive.Entries) {
            if ([string]::IsNullOrWhiteSpace($entry.FullName)) {
                continue
            }

            if ($entry.FullName -notmatch '\.pst$') {
                continue
            }

            $relativePath = Normalize-RelativePath -Value $entry.FullName
            $workKey = Get-PstWorkKey -ZipFile $ZipFile -PstRelativePath $relativePath
            $totalPstEntries++

            if ($CheckpointState.Completed.ContainsKey($workKey)) {
                $completedSkipped++
                continue
            }

            $entryObject = [PSCustomObject]@{
                RelativePath = $relativePath
                WorkKey      = $workKey
            }

            [void]$pendingEntries.Add($entryObject)
            $pendingLookup[$workKey] = $entryObject
        }
    }
    finally {
        if ($null -ne $archive) {
            $archive.Dispose()
        }
    }

    return [PSCustomObject]@{
        ZipFile           = $ZipFile
        TotalPstEntries   = $totalPstEntries
        CompletedSkipped  = $completedSkipped
        PendingEntries    = $pendingEntries
        PendingLookup     = $pendingLookup
        PendingCount      = $pendingEntries.Count
    }
}

function Search-MailFolderRecursive {
    param(
        $Folder,

        [Parameter(Mandatory = $true)]
        [string]$ZipFilePath,

        [Parameter(Mandatory = $true)]
        [string]$PstRelativePath,

        [Parameter(Mandatory = $true)]
        [hashtable]$StageContext
    )

    $folderPath = Get-FolderPathSafe -Folder $Folder

    $items = $null
    try {
        $items = $Folder.Items
        $itemCount = $items.Count

        for ($i = 1; $i -le $itemCount; $i++) {
            $item = $null

            try {
                $item = $items.Item($i)

                $isMail = $false
                try {
                    $isMail = ($item.Class -eq 43) # olMail
                }
                catch {
                    $isMail = $false
                }

                if (-not $isMail) {
                    continue
                }

                $sender     = Get-SenderSmtp -MailItem $item
                $recipients = Get-RecipientSmtps -MailItem $item

                $senderHit    = Test-IsTargetDomain -Value $sender
                $recipientHit = $false

                foreach ($r in $recipients) {
                    if (Test-IsTargetDomain -Value $r) {
                        $recipientHit = $true
                        break
                    }
                }

                if ($senderHit -or $recipientHit) {
                    $exportResult = Save-FoundMailToStage -MailItem $item -StageContext $StageContext

                    $row = @{
                        ZipFile          = $ZipFilePath
                        PstRelativePath  = $PstRelativePath
                        FolderPath       = $folderPath
                        Subject          = (Safe-String $item.Subject)
                        Sender           = (Safe-String $sender)
                        Recipients       = (($recipients | Where-Object { $_ }) -join '; ')
                        SentOn           = (Get-DateString $item.SentOn)
                        ReceivedTime     = (Get-DateString $item.ReceivedTime)
                        SenderHit        = $senderHit
                        RecipientHit     = $recipientHit
                        ExportedFilePath = $exportResult.RelativePath
                        ExportStatus     = $exportResult.Status
                        WorkKey          = $StageContext.WorkKey
                    }

                    Write-StageCsvRow -StageContext $StageContext -Row $row
                }
            }
            catch {
                Write-Log -Level 'WARN' -Message ("Error reading an item in folder '{0}' from PST '{1}': {2}" -f $folderPath, $PstRelativePath, $_.Exception.Message)
            }
            finally {
                Release-ComObject -Obj $item
            }
        }
    }
    catch {
        Write-Log -Level 'WARN' -Message ("Error reading items in folder '{0}' from PST '{1}': {2}" -f $folderPath, $PstRelativePath, $_.Exception.Message)
    }
    finally {
        Release-ComObject -Obj $items
    }

    $subFolders = $null
    try {
        $subFolders = $Folder.Folders
        $subCount = $subFolders.Count

        for ($j = 1; $j -le $subCount; $j++) {
            $sub = $null

            try {
                $sub = $subFolders.Item($j)
                Search-MailFolderRecursive -Folder $sub -ZipFilePath $ZipFilePath -PstRelativePath $PstRelativePath -StageContext $StageContext
            }
            catch {
                Write-Log -Level 'WARN' -Message ("Error entering a subfolder under '{0}' in PST '{1}': {2}" -f $folderPath, $PstRelativePath, $_.Exception.Message)
            }
            finally {
                Release-ComObject -Obj $sub
            }
        }
    }
    catch {
        Write-Log -Level 'WARN' -Message ("Error reading subfolders in '{0}' from PST '{1}': {2}" -f $folderPath, $PstRelativePath, $_.Exception.Message)
    }
    finally {
        Release-ComObject -Obj $subFolders
    }
}

function Process-PstFile {
    param(
        $Namespace,

        [Parameter(Mandatory = $true)]
        [string]$PstPath,

        [Parameter(Mandatory = $true)]
        [string]$ZipFilePath,

        [Parameter(Mandatory = $true)]
        [string]$PstRelativePath,

        [Parameter(Mandatory = $true)]
        [hashtable]$StageContext
    )

    Write-Log -Level 'INFO' -Message ("Opening PST: {0}" -f $PstPath)

    $store      = $null
    $rootFolder = $null
    $storeAdded = $false

    try {
        $Namespace.AddStore($PstPath)
        $storeAdded = $true

        $store = Get-StoreByFilePath -Namespace $Namespace -PstPath $PstPath
        if (-not $store) {
            throw "Could not find the opened PST in the Outlook profile."
        }

        $rootFolder = $store.GetRootFolder()
        Search-MailFolderRecursive -Folder $rootFolder -ZipFilePath $ZipFilePath -PstRelativePath $PstRelativePath -StageContext $StageContext
    }
    finally {
        if ($storeAdded -and $rootFolder) {
            try {
                $Namespace.RemoveStore($rootFolder)
            }
            catch {
                Write-Log -Level 'WARN' -Message ("PST could not be cleanly removed from Outlook: {0}" -f $PstPath)
            }
        }

        Release-ComObject -Obj $rootFolder
        Release-ComObject -Obj $store

        Invoke-ComReleaseCycle
        Start-Sleep -Milliseconds 500
    }
}

function Cleanup-StaleExtractDirs {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TempRoot
    )

    if (-not (Test-Path -LiteralPath $TempRoot)) {
        return
    }

    $dirs = Get-ChildItem -Path $TempRoot -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -like 'extract_*' }

    foreach ($dir in $dirs) {
        Write-Log -Level 'WARN' -Message ("Removing stale extraction directory from a previous run: {0}" -f $dir.FullName)
        [void](Remove-DirectoryRobust -Path $dir.FullName -RetryCount $Config.DeleteRetryCount -RetryDelayMs $Config.DeleteRetryDelayMs -Quiet)
    }
}

function Cleanup-StaleStageDirs {
    param(
        [Parameter(Mandatory = $true)]
        [string]$StageRoot
    )

    if (-not (Test-Path -LiteralPath $StageRoot)) {
        return
    }

    $dirs = Get-ChildItem -Path $StageRoot -Directory -ErrorAction SilentlyContinue

    foreach ($dir in $dirs) {
        Write-Log -Level 'WARN' -Message ("Removing stale stage directory from a previous interrupted run: {0}" -f $dir.FullName)
        [void](Remove-DirectoryRobust -Path $dir.FullName -RetryCount $Config.DeleteRetryCount -RetryDelayMs $Config.DeleteRetryDelayMs -Quiet)
    }
}

function Process-ZipFile {
    param(
        [Parameter(Mandatory = $true)]
        $Namespace,

        [Parameter(Mandatory = $true)]
        [pscustomobject]$ZipInventory
    )

    $zip = $ZipInventory.ZipFile

    if ($ZipInventory.PendingCount -le 0) {
        Write-Log -Level 'INFO' -Message ("Skipping ZIP because all PST entries are already completed according to checkpoint: {0}" -f $zip.FullName)
        return
    }

    $extractDir = Join-Path -Path $Config.TempRoot -ChildPath ("extract_{0}_{1}" -f ([System.IO.Path]::GetFileNameWithoutExtension($zip.Name)), [guid]::NewGuid().ToString('N'))

    $script:CurrentZipName         = $zip.Name
    $script:CurrentZipTotalPst     = $ZipInventory.PendingCount
    $script:CurrentZipProcessedPst = 0
    $script:CurrentZipWatch        = [System.Diagnostics.Stopwatch]::StartNew()
    $script:CurrentPstName         = ''
    $script:CurrentPhase           = 'Extracting ZIP'

    Update-ProgressBars

    Ensure-Directory -Path $extractDir
    Write-Log -Level 'INFO' -Message ("Extracting ZIP: {0}" -f $zip.FullName)

    try {
        Expand-Archive -LiteralPath $zip.FullName -DestinationPath $extractDir -Force
    }
    catch {
        Write-Log -Level 'WARN' -Message ("ZIP could not be extracted: {0}. Error: {1}" -f $zip.FullName, $_.Exception.Message)
        return
    }

    try {
        $script:CurrentPhase = 'Searching for PST files'
        Update-ProgressBars

        $pstFiles = Get-ChildItem -Path $extractDir -Recurse -File -Filter '*.pst' -ErrorAction Stop
        $seenPending = @{}

        foreach ($pst in $pstFiles) {
            $relativePath = $null
            try {
                $relativePath = Get-RelativePathSafe -BasePath $extractDir -ChildPath $pst.FullName
            }
            catch {
                Write-Log -Level 'WARN' -Message ("Could not calculate relative PST path for '{0}'. Skipping file." -f $pst.FullName)
                continue
            }

            $workKey = Get-PstWorkKey -ZipFile $zip -PstRelativePath $relativePath

            if (-not $ZipInventory.PendingLookup.ContainsKey($workKey)) {
                continue
            }

            $seenPending[$workKey] = $true

            $script:CurrentPstName = [System.IO.Path]::GetFileName($relativePath)
            $script:CurrentPhase   = 'Processing PST'
            Update-ProgressBars

            $stageContext = Initialize-StageContext -ZipFile $zip -PstRelativePath $relativePath -WorkKey $workKey
            $pstCompleted = $false

            try {
                Process-PstFile `
                    -Namespace $Namespace `
                    -PstPath $pst.FullName `
                    -ZipFilePath $zip.FullName `
                    -PstRelativePath $relativePath `
                    -StageContext $stageContext

                $pstCompleted = $true
            }
            catch {
                Write-Log -Level 'WARN' -Message ("Fatal PST-level error while processing '{0}' from ZIP '{1}': {2}" -f $relativePath, $zip.FullName, $_.Exception.Message)
            }
            finally {
                if ($pstCompleted) {
                    try {
                        Commit-PstStage `
                            -StageContext $stageContext `
                            -MasterCsvPath $Config.OutputCsv `
                            -FoundMailsRoot $Config.FoundMailsRoot `
                            -CheckpointState $script:CheckpointState `
                            -CheckpointPath $script:CheckpointPath

                        Write-Log -Level 'INFO' -Message ("Committed PST successfully: {0} | Hits committed from this PST: {1}" -f $relativePath, $stageContext.HitCount)
                    }
                    catch {
                        Write-Log -Level 'WARN' -Message ("Failed to commit staged results for PST '{0}': {1}" -f $relativePath, $_.Exception.Message)
                        Abandon-PstStage -StageContext $stageContext
                    }
                }
                else {
                    Abandon-PstStage -StageContext $stageContext
                }

                $script:CurrentZipProcessedPst++
                $script:OverallProcessedPst++
                Update-ProgressBars
            }
        }

        foreach ($pendingEntry in $ZipInventory.PendingEntries) {
            if (-not $seenPending.ContainsKey($pendingEntry.WorkKey)) {
                Write-Log -Level 'WARN' -Message ("Expected pending PST entry was not found after extraction. ZIP: {0} | PST entry: {1}" -f $zip.FullName, $pendingEntry.RelativePath)
            }
        }

        $script:CurrentPhase = 'ZIP completed'
        Update-ProgressBars
    }
    finally {
        if ($null -ne $script:CurrentZipWatch) {
            $script:CurrentZipWatch.Stop()
        }

        Invoke-ComReleaseCycle
        Start-Sleep -Milliseconds 500

        if (-not $Config.KeepExtractedFiles) {
            [void](Remove-DirectoryRobust -Path $extractDir -RetryCount $Config.DeleteRetryCount -RetryDelayMs $Config.DeleteRetryDelayMs)
        }
    }
}

# ------------------------------------------------------------
# Main
# ------------------------------------------------------------
try {
    Add-Type -AssemblyName System.IO.Compression.FileSystem | Out-Null
}
catch {
    # Usually harmless if already loaded
}

# Normalize target domain
$script:TargetDomainNormalized = Normalize-Email -Value $Config.TargetDomain
if (-not $script:TargetDomainNormalized) {
    throw "TargetDomain is empty or invalid."
}
if (-not $script:TargetDomainNormalized.StartsWith('@')) {
    $script:TargetDomainNormalized = '@' + $script:TargetDomainNormalized
}

# Basic validation
if ([string]::IsNullOrWhiteSpace($Config.ZipRoot)) {
    throw "ZipRoot is empty. Please set it in the configuration block."
}
if (-not (Test-Path -LiteralPath $Config.ZipRoot)) {
    throw "ZipRoot does not exist: $($Config.ZipRoot)"
}
if ($Config.ExportFormat.ToLowerInvariant() -notin @('msg','eml')) {
    throw "ExportFormat must be 'msg' or 'eml'."
}

# Prepare directories
Ensure-Directory -Path $Config.TempRoot
Ensure-Directory -Path $Config.StateRoot
Ensure-Directory -Path $script:StageRoot

if ($Config.ExportMatchedMails) {
    Ensure-Directory -Path $Config.FoundMailsRoot
}

# Start fresh if requested
if ($Config.StartFresh) {
    Write-Log -Level 'WARN' -Message "StartFresh is enabled. Existing output/state for this run will be removed."

    [void](Remove-FileRobust -Path $Config.OutputCsv -RetryCount $Config.DeleteRetryCount -RetryDelayMs $Config.DeleteRetryDelayMs)
    [void](Remove-FileRobust -Path $script:CheckpointPath -RetryCount $Config.DeleteRetryCount -RetryDelayMs $Config.DeleteRetryDelayMs)
    [void](Remove-DirectoryRobust -Path $script:StageRoot -RetryCount $Config.DeleteRetryCount -RetryDelayMs $Config.DeleteRetryDelayMs -Quiet)
    [void](Remove-DirectoryRobust -Path $Config.FoundMailsRoot -RetryCount $Config.DeleteRetryCount -RetryDelayMs $Config.DeleteRetryDelayMs -Quiet)

    Ensure-Directory -Path $script:StageRoot
    if ($Config.ExportMatchedMails) {
        Ensure-Directory -Path $Config.FoundMailsRoot
    }
}

# Cleanup leftovers from earlier interrupted runs
Cleanup-StaleExtractDirs -TempRoot $Config.TempRoot
Cleanup-StaleStageDirs -StageRoot $script:StageRoot

# State checks
Assert-StateConsistency `
    -OutputCsv $Config.OutputCsv `
    -CheckpointPath $script:CheckpointPath `
    -EnableResume $Config.EnableResume `
    -StrictStateSafety $Config.StrictStateSafety

# Load state
if ($Config.EnableResume) {
    $script:CheckpointState = Load-CheckpointState -Path $script:CheckpointPath
}
else {
    $script:CheckpointState = @{
        Completed = @{}
    }
}

# Initialize master CSV
Initialize-MasterCsv -Path $Config.OutputCsv

$outlook   = $null
$namespace = $null

try {
    $script:CurrentPhase = 'Searching for ZIP files'
    Update-ProgressBars

    Write-Log -Level 'INFO' -Message ("Searching for ZIP files under: {0}" -f $Config.ZipRoot)
    $zipFiles = Get-ChildItem -Path $Config.ZipRoot -Recurse -File -Filter '*.zip' -ErrorAction Stop

    if (-not $zipFiles -or $zipFiles.Count -eq 0) {
        Complete-ProgressBars
        Write-Log -Level 'INFO' -Message "No ZIP files found."
        return
    }

    $script:TotalZipCount = $zipFiles.Count
    Write-Log -Level 'INFO' -Message ("Found {0} ZIP file(s)." -f $script:TotalZipCount)

    $zipInventoryList = New-Object 'System.Collections.Generic.List[object]'

    for ($i = 0; $i -lt $zipFiles.Count; $i++) {
        $zip = $zipFiles[$i]

        $percent = [int][math]::Min(100, [math]::Round((($i + 1) / $zipFiles.Count) * 100, 0))
        Write-Progress -Id 0 -Activity 'Initialization' -Status ("Analyzing ZIP contents {0}/{1}" -f ($i + 1), $zipFiles.Count) -CurrentOperation $zip.Name -PercentComplete $percent

        try {
            $inventory = Get-ZipInventory -ZipFile $zip -CheckpointState $script:CheckpointState
            [void]$zipInventoryList.Add($inventory)
        }
        catch {
            Write-Log -Level 'WARN' -Message ("ZIP inventory could not be analyzed: {0}. Error: {1}" -f $zip.FullName, $_.Exception.Message)
        }
    }

    $pendingCount = 0
    foreach ($item in $zipInventoryList) {
        $pendingCount += [int]$item.PendingCount
    }

    $script:OverallTotalPst = $pendingCount

    Write-Log -Level 'INFO' -Message ("Pending PST files to process in this run: {0}" -f $script:OverallTotalPst)

    if ($script:OverallTotalPst -le 0) {
        Complete-ProgressBars
        Write-Log -Level 'INFO' -Message "All PST files are already completed according to the checkpoint. Nothing to do."
        return
    }

    $script:CurrentPhase = 'Starting Outlook COM'
    Update-ProgressBars

    Write-Log -Level 'INFO' -Message "Starting Outlook COM"
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")

    for ($z = 0; $z -lt $zipInventoryList.Count; $z++) {
        $zipInfo = $zipInventoryList[$z]
        $script:CurrentZipIndex = $z + 1

        try {
            Process-ZipFile -Namespace $namespace -ZipInventory $zipInfo
        }
        catch {
            Write-Log -Level 'WARN' -Message ("Unhandled ZIP-level error for '{0}': {1}" -f $zipInfo.ZipFile.FullName, $_.Exception.Message)
        }
        finally {
            $script:ProcessedZipCount++
            $script:CurrentPstName = ''
            $script:CurrentPhase = 'Switching to next ZIP'
            Update-ProgressBars
        }
    }

    $script:CurrentZipName = ''
    $script:CurrentPstName = ''
    $script:CurrentPhase   = 'Completed'
    Update-ProgressBars

    Write-Log -Level 'INFO' -Message ("Run completed. Pending PST processed this run: {0}/{1}" -f $script:OverallProcessedPst, $script:OverallTotalPst)
    Write-Log -Level 'INFO' -Message ("Committed hits written in this run: {0}" -f $script:CommittedHitCount)
    Write-Log -Level 'INFO' -Message ("Master CSV: {0}" -f $Config.OutputCsv)

    if ($Config.ExportMatchedMails) {
        Write-Log -Level 'INFO' -Message ("Found mails root: {0}" -f $Config.FoundMailsRoot)
    }
}
finally {
    Complete-ProgressBars

    Release-ComObject -Obj $namespace
    Release-ComObject -Obj $outlook

    Invoke-ComReleaseCycle

    if ($null -ne $script:CurrentZipWatch -and $script:CurrentZipWatch.IsRunning) {
        $script:CurrentZipWatch.Stop()
    }

    $script:RunWatch.Stop()
}
