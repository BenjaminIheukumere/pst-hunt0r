# PST Hunt0r

PST Hunt0r is a PowerShell hunting utility for digging through ZIP archives that contain Outlook PST files and identifying emails tied to a target domain.

It recursively scans ZIP collections, extracts embedded PST files, opens them via Outlook/MAPI, walks every mail folder, and exports all matching results to a clean CSV file. Progress bars, elapsed time, and ETA are included, so you can watch the beast chew through archives instead of staring into the void.

* * *
## Screenshot of pst-hunt0r in action
![Screenshot of pst-hunt0r in action](https://github.com/BenjaminIheukumere/pst-hunt0r/blob/main/pst_hunt0r.png)
* * *

## Features

- Recursively scans a folder for ZIP archives
- Pre-analyzes ZIP files to estimate contained PST files
- Extracts ZIP archives to a temporary working directory
- Recursively finds Outlook PST files inside extracted content
- Opens PST files through Outlook/MAPI
- Walks all folders and mail items recursively
- Checks both sender and recipient addresses against a target domain
- Resolves SMTP addresses properly, including Exchange/MAPI edge cases
- Exports all hits to a CSV file
- Shows progress for:
  - overall processing
  - current ZIP file
  - elapsed time
  - estimated remaining time
- Continues processing even if individual ZIPs or PSTs cause issues

* * *

## Requirements

- Windows
- Outlook Desktop installed
- A working Outlook profile on the machine
- PowerShell 5.1 or later
- Read access to the archive location
- Enough free disk space for temporary extraction

* * *

## Installation

Clone the repository:

    git clone https://github.com/BenjaminIheukumere/pst-hunt0r.git
    cd pst-hunt0r

Make sure Outlook Desktop is installed and configured on the machine that will run the script.

* * *

## Configuration

This script is meant to be configured directly inside the script before execution.

Set the variables in the script to match your own environment, for example:

- `$ZipRoot`
- `$TargetDomain`
- `$TempRoot`
- `$OutputCsv`

This keeps the repository generic and reusable. No customer-specific paths, no hardcoded case data, no baked-in target domain nonsense.

* * *

## Usage

Run the script in PowerShell:

    .\pst-hunt0r.ps1

Before running it, edit the variables inside the script so they match your environment.

The script will then:

1. Search for ZIP files under the configured root folder
2. Estimate how many PST files are contained in total
3. Extract each ZIP archive
4. Open each PST file through Outlook
5. Search all emails for the configured target domain
6. Export all matches to a CSV file

* * *

## Output

Matching emails are written to a CSV file.

The CSV includes fields such as:

- ZIP file path
- PST file path
- folder path inside the PST
- subject
- sender
- recipients
- sent time
- received time
- sender hit
- recipient hit

The script also prints status messages to the console and shows live progress bars while running.

* * *

## Notes

- Not intended for password-protected ZIP archives
- Outlook is required because PST processing relies on Outlook/MAPI
- SMTP resolution includes fallback handling for Exchange-style address entries
- Very large PST files may take time to process
- ETA is an estimate based on progress so far and may shift if archive sizes vary heavily

* * *

## Disclaimer

This tool is intended for authorized investigations, incident response, eDiscovery, and administrative analysis only.

Use it only on data sets and environments where you have explicit permission to do so. The author is not responsible for misuse, damage, or awkward conversations with legal.

* * *

## About

PST Hunt0r is a practical PowerShell utility for investigators, administrators, incident responders, and security teams who need to search large collections of archived Outlook PST files stored inside ZIP archives.

It is built to be simple, adaptable, and effective: edit a few variables in the script, launch it, and let it hunt.
