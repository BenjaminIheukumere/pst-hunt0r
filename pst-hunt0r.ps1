<#
.SYNOPSIS
    pst-hunt0r by Benjamin Iheukumere | b.iheukumere@safelink-it.com
    Scans ZIP files on a network share for PST files and searches them
    for emails where the sender or any recipient ends with @example.com.

.DESCRIPTION
    - Recursively scans a root folder for *.zip files
    - Pre-analyzes ZIP files and counts contained PST files
    - Extracts each ZIP into a temporary directory
    - Recursively finds *.pst files inside
    - Opens each PST via Outlook/MAPI
    - Recursively traverses all folders and mail items
    - Checks sender and all recipients against the target domain
    - Writes all hits to a CSV file
    - Shows progress per ZIP and overall, including ETA and elapsed time

.REQUIREMENTS
    - Windows
    - Outlook Desktop installed
    - A working Outlook profile on the machine
    - Read permissions to the network share
    - Enough local temp storage

.NOTES
    Not suitable for password-protected ZIP files.
    The time left for scanning is more of a guessing game...
#>

[CmdletBinding()]
param(
    [Parameter()]
    [string]$ZipRoot = 'Z:\path\to\your\zipped\pst\files',

    [Parameter()]
    [string]$TargetDomain = '@example.com',

    [Parameter()]
    [string]$TempRoot = 'F:\path\for\temp\extraction\of\zipped\files',

    [Parameter()]
    [string]$OutputCsv = '',

    [Parameter()]
    [switch]$KeepExtractedFiles
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ----------------------------------------
# Script-scope state for progress tracking
# ----------------------------------------
$script:RunWatch = [System.Diagnostics.Stopwatch]::StartNew()
$script:TotalZipCount = 0
$script:ProcessedZipCount = 0

$script:OverallTotalPst = 0
$script:OverallProcessedPst = 0

$script:CurrentZipName = ''
$script:CurrentZipIndex = 0
$script:CurrentZipTotalPst = 0
$script:CurrentZipProcessedPst = 0
$script:CurrentZipWatch = $null
$script:CurrentPstName = ''
$script:CurrentPhase = 'Initialization'

$script:ResultsRef = $null
$script:TargetDomainNormalized = $null
$script:PrSmtpAddressUri = "http://schemas.microsoft.com/mapi/proptag/0x39FE001F"

# ----------------------------------------
# Helper functions
# ----------------------------------------

function Write-Info {
    param([string]$Message)
    Write-Host "[INFO] $Message" -ForegroundColor Cyan
}

function Write-WarnMsg {
    param([string]$Message)
    Write-Warning $Message
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

function Test-IsTargetDomain {
    param([AllowNull()][string]$Value)

    $v = Normalize-Email -Value $Value
    if (-not $v) {
        return $false
    }

    return $v.EndsWith($script:TargetDomainNormalized)
}

function Safe-String {
    param($Value)
    if ($null -eq $Value) {
        return ''
    }
    return [string]$Value
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
    $overallEta = $null
    $overallPercent = 0

    if ($script:OverallTotalPst -gt 0) {
        $overallPercent = [int][math]::Min(100, [math]::Round(($script:OverallProcessedPst / $script:OverallTotalPst) * 100, 0))
        $overallEta = Get-EtaByUnits -CompletedUnits $script:OverallProcessedPst -TotalUnits $script:OverallTotalPst -Elapsed $elapsedOverall
    }
    elseif ($script:TotalZipCount -gt 0) {
        $overallPercent = [int][math]::Min(100, [math]::Round(($script:ProcessedZipCount / $script:TotalZipCount) * 100, 0))
        $overallEta = Get-EtaByUnits -CompletedUnits $script:ProcessedZipCount -TotalUnits $script:TotalZipCount -Elapsed $elapsedOverall
    }

    $hitCount = 0
    if ($null -ne $script:ResultsRef) {
        $hitCount = $script:ResultsRef.Count
    }

    $overallStatus = if ($script:OverallTotalPst -gt 0) {
        "ZIP $($script:ProcessedZipCount)/$($script:TotalZipCount) | PST $($script:OverallProcessedPst)/$($script:OverallTotalPst) | Hits $hitCount | Elapsed $(Format-Duration $elapsedOverall) | ETA $(Format-Duration $overallEta)"
    }
    else {
        "ZIP $($script:ProcessedZipCount)/$($script:TotalZipCount) | Hits $hitCount | Elapsed $(Format-Duration $elapsedOverall) | ETA $(Format-Duration $overallEta)"
    }

    $overallCurrentOperation = if ($script:CurrentZipName) {
        "$($script:CurrentPhase) | $($script:CurrentZipName)"
    }
    else {
        $script:CurrentPhase
    }

    Write-Progress -Id 0 -Activity $OverallActivity -Status $overallStatus -CurrentOperation $overallCurrentOperation -PercentComplete $overallPercent

    if ($script:CurrentZipName) {
        $zipElapsed = if ($null -ne $script:CurrentZipWatch) { $script:CurrentZipWatch.Elapsed } else { [TimeSpan]::Zero }
        $zipEta = $null
        $zipPercent = 0

        if ($script:CurrentZipTotalPst -gt 0) {
            $zipPercent = [int][math]::Min(100, [math]::Round(($script:CurrentZipProcessedPst / $script:CurrentZipTotalPst) * 100, 0))
            $zipEta = Get-EtaByUnits -CompletedUnits $script:CurrentZipProcessedPst -TotalUnits $script:CurrentZipTotalPst -Elapsed $zipElapsed
        }

        $zipStatus = if ($script:CurrentZipTotalPst -gt 0) {
            "ZIP $($script:CurrentZipIndex)/$($script:TotalZipCount) | PST $($script:CurrentZipProcessedPst)/$($script:CurrentZipTotalPst) | Hits $hitCount | Elapsed $(Format-Duration $zipElapsed) | ETA $(Format-Duration $zipEta)"
        }
        else {
            "ZIP $($script:CurrentZipIndex)/$($script:TotalZipCount) | No PST files detected | Hits $hitCount | Elapsed $(Format-Duration $zipElapsed)"
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

function Get-ZipPstEntryCount {
    param([System.IO.FileInfo]$ZipFile)

    $archive = $null

    try {
        $archive = [System.IO.Compression.ZipFile]::OpenRead($ZipFile.FullName)
        $count = 0

        foreach ($entry in $archive.Entries) {
            if ($entry.FullName -match '\.pst$') {
                $count++
            }
        }

        return $count
    }
    finally {
        if ($null -ne $archive) {
            $archive.Dispose()
        }
    }
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

    $list = [System.Collections.Generic.List[string]]::new()

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
                $list.Add($address)
            }
        }
        catch {
            # Recipient issue is not fatal
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
            $store = $null
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

function Search-MailFolderRecursive {
    param(
        $Folder,
        [string]$ZipFile,
        [string]$PstFile,
        [System.Collections.Generic.List[object]]$Results
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
                    $isMail = ($item.Class -eq 43)
                }
                catch {
                    $isMail = $false
                }

                if (-not $isMail) {
                    continue
                }

                $sender = Get-SenderSmtp -MailItem $item
                $recipients = Get-RecipientSmtps -MailItem $item

                $senderHit = Test-IsTargetDomain -Value $sender
                $recipientHit = $false

                foreach ($r in $recipients) {
                    if (Test-IsTargetDomain -Value $r) {
                        $recipientHit = $true
                        break
                    }
                }

                if ($senderHit -or $recipientHit) {
                    $Results.Add([PSCustomObject]@{
                        ZipFile      = $ZipFile
                        PstFile      = $PstFile
                        FolderPath   = $folderPath
                        Subject      = (Safe-String $item.Subject)
                        Sender       = (Safe-String $sender)
                        Recipients   = (($recipients | Where-Object { $_ }) -join '; ')
                        SentOn       = (Get-DateString $item.SentOn)
                        ReceivedTime = (Get-DateString $item.ReceivedTime)
                        SenderHit    = $senderHit
                        RecipientHit = $recipientHit
                    })
                }
            }
            catch {
                Write-WarnMsg "Error reading an item in folder '$folderPath' from PST '$PstFile': $($_.Exception.Message)"
            }
            finally {
                Release-ComObject -Obj $item
            }
        }
    }
    catch {
        Write-WarnMsg "Error reading items in folder '$folderPath' from PST '$PstFile': $($_.Exception.Message)"
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
                Search-MailFolderRecursive -Folder $sub -ZipFile $ZipFile -PstFile $PstFile -Results $Results
            }
            catch {
                Write-WarnMsg "Error entering a subfolder under '$folderPath' in PST '$PstFile': $($_.Exception.Message)"
            }
            finally {
                Release-ComObject -Obj $sub
            }
        }
    }
    catch {
        Write-WarnMsg "Error reading subfolders in '$folderPath' from PST '$PstFile': $($_.Exception.Message)"
    }
    finally {
        Release-ComObject -Obj $subFolders
    }
}

function Process-PstFile {
    param(
        $Namespace,
        [string]$PstPath,
        [string]$ZipFile,
        [System.Collections.Generic.List[object]]$Results
    )

    Write-Info "Opening PST: $PstPath"

    $store = $null
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
        Search-MailFolderRecursive -Folder $rootFolder -ZipFile $ZipFile -PstFile $PstPath -Results $Results
    }
    finally {
        if ($storeAdded -and $rootFolder) {
            try {
                $Namespace.RemoveStore($rootFolder)
            }
            catch {
                Write-WarnMsg "PST could not be cleanly removed from Outlook: $PstPath"
            }
        }

        Release-ComObject -Obj $rootFolder
        Release-ComObject -Obj $store
    }
}

function Process-ZipFile {
    param(
        [pscustomobject]$ZipInventory,
        $Namespace,
        [System.Collections.Generic.List[object]]$Results
    )

    $zip = $ZipInventory.ZipFile
    $extractDir = Join-Path -Path $TempRoot -ChildPath ("{0}_{1}" -f $zip.BaseName, [guid]::NewGuid().ToString('N'))

    $script:CurrentZipName = $zip.Name
    $script:CurrentZipTotalPst = [int]$ZipInventory.PstEntryCount
    $script:CurrentZipProcessedPst = 0
    $script:CurrentZipWatch = [System.Diagnostics.Stopwatch]::StartNew()
    $script:CurrentPstName = ''
    $script:CurrentPhase = 'Extracting ZIP'
    Update-ProgressBars

    Write-Info "Extracting ZIP: $($zip.FullName)"
    New-Item -Path $extractDir -ItemType Directory -Force | Out-Null

    try {
        Expand-Archive -LiteralPath $zip.FullName -DestinationPath $extractDir -Force
    }
    catch {
        Write-WarnMsg "ZIP could not be extracted: $($zip.FullName) - $($_.Exception.Message)"
        return
    }

    try {
        $script:CurrentPhase = 'Searching for PST files'
        Update-ProgressBars

        $pstFiles = Get-ChildItem -Path $extractDir -Recurse -File -Filter '*.pst' -ErrorAction Stop

        $actualPstCount = @($pstFiles).Count

        if ($actualPstCount -ne [int]$ZipInventory.PstEntryCount) {
            $script:OverallTotalPst = $script:OverallTotalPst - [int]$ZipInventory.PstEntryCount + $actualPstCount
            $script:CurrentZipTotalPst = $actualPstCount
            Update-ProgressBars
        }

        if (-not $pstFiles -or $actualPstCount -eq 0) {
            Write-Info "No PST files found in ZIP: $($zip.FullName)"
            $script:CurrentPhase = 'No PST files found'
            Update-ProgressBars
            return
        }

        foreach ($pst in $pstFiles) {
            $script:CurrentPstName = $pst.Name
            $script:CurrentPhase = 'Processing PST'
            Update-ProgressBars

            try {
                Process-PstFile -Namespace $Namespace -PstPath $pst.FullName -ZipFile $zip.FullName -Results $Results
            }
            catch {
                Write-WarnMsg "Error processing PST '$($pst.FullName)' from ZIP '$($zip.FullName)': $($_.Exception.Message)"
            }
            finally {
                $script:CurrentZipProcessedPst++
                $script:OverallProcessedPst++
                Update-ProgressBars
            }
        }

        $script:CurrentPhase = 'ZIP completed'
        Update-ProgressBars
    }
    finally {
        if (-not $KeepExtractedFiles) {
            try {
                Remove-Item -LiteralPath $extractDir -Recurse -Force -ErrorAction Stop
            }
            catch {
                Write-WarnMsg "Temporary directory could not be deleted: $extractDir"
            }
        }

        if ($null -ne $script:CurrentZipWatch) {
            $script:CurrentZipWatch.Stop()
        }
    }
}

# ----------------------------------------
# Main
# ----------------------------------------

try {
    Add-Type -AssemblyName System.IO.Compression.FileSystem | Out-Null
}
catch {
    # Usually harmless if already loaded
}

if (-not (Test-Path -LiteralPath $ZipRoot)) {
    throw "ZipRoot does not exist: $ZipRoot"
}

New-Item -Path $TempRoot -ItemType Directory -Force | Out-Null

if ([string]::IsNullOrWhiteSpace($OutputCsv)) {
    $OutputCsv = Join-Path -Path $TempRoot -ChildPath ("PstFarfetchHits_{0}.csv" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
}

$script:TargetDomainNormalized = (Normalize-Email -Value $TargetDomain)
if (-not $script:TargetDomainNormalized) {
    throw "TargetDomain is empty or invalid."
}

if (-not ($script:TargetDomainNormalized.StartsWith('@'))) {
    $script:TargetDomainNormalized = '@' + $script:TargetDomainNormalized
}

$outlook = $null
$namespace = $null
$results = [System.Collections.Generic.List[object]]::new()
$script:ResultsRef = $results

try {
    $script:CurrentPhase = 'Searching for ZIP files'
    Update-ProgressBars

    Write-Info "Searching for ZIP files under: $ZipRoot"
    $zipFiles = Get-ChildItem -Path $ZipRoot -Recurse -File -Filter '*.zip' -ErrorAction Stop

    if (-not $zipFiles -or $zipFiles.Count -eq 0) {
        Complete-ProgressBars
        Write-Info "No ZIP files found."
        return
    }

    $script:TotalZipCount = $zipFiles.Count
    Write-Info ("Found {0} ZIP file(s)." -f $script:TotalZipCount)

    $zipInventory = [System.Collections.Generic.List[object]]::new()

    for ($i = 0; $i -lt $zipFiles.Count; $i++) {
        $zip = $zipFiles[$i]
        $percent = [int][math]::Min(100, [math]::Round((($i + 1) / $zipFiles.Count) * 100, 0))
        Write-Progress -Id 0 -Activity 'Initialization' -Status "Analyzing ZIP contents $($i + 1)/$($zipFiles.Count)" -CurrentOperation $zip.Name -PercentComplete $percent

        $pstEntryCount = 0
        $scanError = $null

        try {
            $pstEntryCount = Get-ZipPstEntryCount -ZipFile $zip
        }
        catch {
            $scanError = $_.Exception.Message
            Write-WarnMsg "ZIP contents could not be pre-analyzed: $($zip.FullName) - $scanError"
        }

        $zipInventory.Add([PSCustomObject]@{
            ZipFile       = $zip
            PstEntryCount = $pstEntryCount
            ScanError     = $scanError
        })
    }

    $script:OverallTotalPst = ($zipInventory | Measure-Object -Property PstEntryCount -Sum).Sum
    if ($null -eq $script:OverallTotalPst) {
        $script:OverallTotalPst = 0
    }

    Write-Info ("Total PST files detected in pre-scan: {0}" -f $script:OverallTotalPst)

    $script:CurrentPhase = 'Starting Outlook COM'
    Update-ProgressBars

    Write-Info "Starting Outlook COM"
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")

    for ($z = 0; $z -lt $zipInventory.Count; $z++) {
        $zipInfo = $zipInventory[$z]
        $script:CurrentZipIndex = $z + 1

        try {
            Process-ZipFile -ZipInventory $zipInfo -Namespace $namespace -Results $results
        }
        catch {
            Write-WarnMsg "Error processing ZIP '$($zipInfo.ZipFile.FullName)': $($_.Exception.Message)"
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
    $script:CurrentPhase = 'Writing CSV'
    Update-ProgressBars

    Write-Info ("Total hits: {0}" -f $results.Count)

    $results |
        Sort-Object ZipFile, PstFile, FolderPath, SentOn, Subject |
        Export-Csv -Path $OutputCsv -Delimiter ';' -NoTypeInformation -Encoding UTF8

    Write-Info "CSV written: $OutputCsv"

    $script:CurrentPhase = 'Completed'
    Update-ProgressBars
}
finally {
    Complete-ProgressBars

    Release-ComObject -Obj $namespace
    Release-ComObject -Obj $outlook

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    if ($null -ne $script:CurrentZipWatch -and $script:CurrentZipWatch.IsRunning) {
        $script:CurrentZipWatch.Stop()
    }

    $script:RunWatch.Stop()
}
