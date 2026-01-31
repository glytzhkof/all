###############################################################################
# Delete all auto-generated MSI log files from the system's temp folder.
#
# Created: 3.Oct.2025
# Author: installdude.com
# Version: 1.0
#
###############################################################################
#
# When automatic MSI logging is enabled via policy, you can accumulate a lot
# of MSI log files in the system's temp folder. This little script will remove
# all auto-generated MSI log files.
#
# Read about MSI logging: https://stackoverflow.com/a/54458890/129130
#

$temppath = $env:TEMP
$counter = 0

try {    

    # RegEx: First match "MSI" literal string for the start of the file name 
    #        followed by 4 or 5 characters before the file name ends with .log
    Get-ChildItem -Path $temppath | Where-Object { $_.Name -match '^MSI.{4,5}\.log$'} | ForEach-Object {
        # Remove-Item options: -Confirm, -Force, -Recurse, -WhatIf, -Verbose
        Remove-Item -Path $_.FullName -Verbose
        $counter++
    }

    Write-Host "Successfully processed: $temppath - $counter MSI log files deleted." -ForegroundColor Green
}
catch {
    Write-Host "Error processing $temppath : $($_.Exception.Message)" -ForegroundColor Red
}

# Prevent Powershell window from closing
Read-Host -Prompt "Press Enter to exit"
