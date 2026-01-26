###############################################################################
# Generates a HTML report listing all MSI packages on the local machine. 
#
# Created: 3.Oct.2025
# Author: Stein-Inge Åsmul
# Version: 1.0
#
###############################################################################
#
# DISCLAIMER:
# 
# This simplified script does not export the MSI upgrade code and it should 
# not trigger any special, unwanted event or logging "noise".
#
# If you need the MSI upgrade code, try this more elaborate script instead: 
# MsiPackageReport-Mini-SearchEngine.ps1 - make sure to read the disclaimer
# section in this script before using it! This is particularly important
# if you are a system administrator.
#

$installer = New-Object -ComObject WindowsInstaller.Installer
$p = 1

# A class and a generic list to store MSI package info
class MsiPackage {
    [string]$Counter; [string]$ProductName; [string]$Version; [string]$PackageCode; [string]$ProductCode; [string]$Language
}

$MsiPackages = New-Object System.Collections.Generic.List[MsiPackage]

# Retrieve MSI package information
$installer.ProductsEx("", "", 7) | ForEach-Object {
    #Write-Host "Processing Package: $p " + $($_.ProductCode())
    $MsiPackages.Add([MsiPackage]@{Counter=$p;ProductName=$($_.InstallProperty('ProductName'));Version=$($_.InstallProperty('VersionString'));PackageCode=$($_.InstallProperty("PackageCode"));ProductCode=$($_.ProductCode());Language=$($_.InstallProperty("Language"))})
    $p += 1
}

# We need an output name and location for the HTML report (unique file name on the desktop)
$desktoppath = [Environment]::GetFolderPath('Desktop')
$filename = "MsiInfoBasic_$($env:COMPUTERNAME)_$((Get-Date).Day).$((Get-Date).Month)(month).$((Get-Date).Year)_$((Get-Date).Hour)-$((Get-Date).Minute)-$((Get-Date).Second).html"
$outputpath = Join-Path -Path $desktoppath -ChildPath $filename

# Use standard HTML export with some simple styling
$MsiPackages | ConvertTo-Html -Title "MSI Package Report" -Head "<title>Package Report</title><style type='text/css'>
    body { font: 14px Calibri; }
    table, td { border: 1px solid black; border-collapse: collapse; padding: 0.3em; border-top: none;}
    th { font: bold 18px Calibri; background-color: purple; text-align: left; color: white; }
    table th { position: sticky;top: 0px;}
</style>" | Out-File $outputpath -Encoding utf8

# Open the exported HTML MSI package list 
Start-Process $outputpath

# Release our single COM object
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($installer) | Out-Null
