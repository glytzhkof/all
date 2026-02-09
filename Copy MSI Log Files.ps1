###############################################################################
# Copy all auto-generated MSI log files from the system's temp folder to a 
# folder on the desktop.
#
# Created: 3.Oct.2025
# Author: installdude.com
# Version: 1.0
#
###############################################################################
#
# When automatic MSI logging is enabled via policy, you can accumulate a lot
# of MSI log files in the system's temp folder. This little script will copy
# all auto-generated MSI log files from the temp folder to the specified
# folder on the user's desktop (for testing purposes and to get a copy of
# MSI log files before cleaning them from TMP).
#
# Read about MSI logging: https://stackoverflow.com/a/54458890/129130
#

$temppath = $env:TEMP
$counter = 0

try {    

    # Output path for the MSI log files on the desktop
    $desktoppath = Join-Path -Path $([Environment]::GetFolderPath("Desktop")) -ChildPath "LogArchive"
    New-Item -ItemType Directory -Path $desktoppath -Force | Out-Null

    # RegEx: First match "MSI" literal string for the start of the file name 
    #        followed by 4 or 5 characters before the file name ends with .log
    Get-ChildItem -Path $temppath | Where-Object { $_.Name -match '^MSI.{4,5}\.log$'} | ForEach-Object {
        
        # The full output path on the desktop for current log file
        $fullpath = Join-Path -Path $desktoppath -ChildPath $_.Name

        # We copy the log files to the desktop location
        Copy-Item -Path $_.FullName -Destination $fullpath

        # We keep track of how many files we have processed (for status messages)
        $counter++
    }

    Write-Host "Successfully processed: $temppath - $counter MSI log files copied." -ForegroundColor Green
}
catch {
    Write-Host "Error processing $temppath : $($_.Exception.Message)" -ForegroundColor Red
}

# Prevent Powershell window from closing
Read-Host -Prompt "Press Enter to exit"
