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
# This script will generate some entries in the system's event log ("Program").
# These entries are harmless, but may constitute some "noise" for system 
# administrators - especially since they are generated every time the script
# is run.
#
# If you have automatic MSI logging enabled, you will also see one log file
# created per MSI package in the system's %TEMP% folder (per run). This 
# happens only if the custom MSI logging policy is enabled.
#
# Check this SO answer for details on MSI auto logging (scroll down a bit):
# https://stackoverflow.com/a/54458890/129130
#
# Technically: it is the invokation of a Session object per MSI package which 
# causes these logging issues. The session object is used to retrieve the MSI
# package's upgrade code.
#
# ALTERNATIVE SCRIPT:
#
# Try the alternative, simpler script: MsiPackageReport-Micro-StandardHtml.ps1
# if you don't want any logging or event "noise". Be aware that this simpler
# script does not export the MSI upgrade code.
#

# Clean the Powershell console window
#Clear-Host

# Windows Installer COM object (MSI is old)
$installer = New-Object -ComObject WindowsInstaller.Installer

$msiUILevelNone = 2 # Show no GUI for activated MSI Session objects
$p = 1

#$ErrorActionPreference = 'SilentlyContinue' # Just continue with next package on error

# Construct the HTML header for output file (using here-string)
$htmloutput = @"
<!DOCTYPE html>
<html lang='en'><head><title>MSI Package Estate Information:</title><meta charset='utf-8'>
<script>function init() { try { document.querySelectorAll('td').forEach(link => { link.addEventListener('mouseenter', function (event) {var range = document.createRange(); range.selectNodeContents(this); var sel = window.getSelection(); sel.removeAllRanges(); sel.addRange(range);});});} catch (error) { console.log(error); }}
function filterTable(filter) { var row; var rows = document.querySelectorAll('table tbody tr'); var rowcount = rows.length; var hiddenrows = 0; for (row = 0; row < rowcount; row++) { if (rows[row].textContent.toUpperCase().indexOf(filter.toUpperCase()) > -1) { rows[row].style.display = '';} else { rows[row].style.display = 'none'; hiddenrows++;}}}
function reset() {document.getElementById('search-box').value = '';filterTable('');}</script>
<style>body {font: 12px Calibri;}a {color: lightgrey;} a:hover {background-color: black;}
table, td {border: 1px solid black;border-collapse: collapse;padding: 0.3em;vertical-align: text-top;border-top: none;}
table>*>tr>td:nth-child(2) { max-width: 300px;}
th {font: bold 18px Calibri;background-color: purple;text-align: left;color: white;}
table th {position: sticky;top: -1px;}</style>
</head><body onload='init()'>
<h1>MSI Package Report</h1><input id='search-box' type='text' onemptied='reset()' autocomplete='off' oninput='filterTable(this.value)' title='Filter table by keyword search' placeholder='Filter by...'>
<button onclick='reset()'>x</button><h2>Use your browser's zoom setting to make text more readable.</h2>
<table><thead><tr>
<th>#</th><th>Product Name</th><th>Version</th><th>Package Code</th><th>Product Code</th><th>Upgrade Code</th><th  title='Product codes that share the same upgrade code.'>Related Product Codes</th><th>Scope</th><th><a href='https://msdn.microsoft.com/en-us/library/ms912047(v=winembedded.10).aspx' target='_blank'>LCID</a></th>
</tr></thead><tbody>`r`n
"@

# Get all installed MSI packages and prepare to initiate session object in no-GUI mode
$products = $installer.ProductsEx("", "", 7)
$totalpackages = $products.Count()
$installer.UILevel = $msiUILevelNone #[Type]::GetType("Microsoft.Deployment.WindowsInstaller.InstallerUILevel").GetField("None").GetValue($null)

# Empty array to hold product codes that share the same upgrade code (related products)
$relatedproductcodes = @()

# Status update
Write-Host "Starting package retrieval..."

# Now process each MSI package in sequence
foreach ($product in $products) {

    $productcode = $product.ProductCode() # Crucial: must add () at end even if it is a property in the object model
    $productname = $product.InstallProperty('ProductName')
    $versionstring = $product.InstallProperty('VersionString')
    $packagecode = $product.InstallProperty('PackageCode')
    $scope = $product.InstallProperty("AssignmentType")
    $lcid = $product.InstallProperty("Language")
    $upgradecode = "" # Will be retrieved later

    switch ($scope) {
        0 { $assignment = "User" }
        1 { $assignment = "Machine" }
        default { $assignment = "Unknown" }
    }

    try {
        # Get upgrade code via MSI session object (reads cached MSI database with applied transforms - apparently)
        $session = $installer.OpenProduct($productcode)

        # So far so good, we have our session object, but upgrade code can be missing 
        $upgradecode = $session.ProductProperty("UpgradeCode")

        # Don't pass empty string to RelatedProducts, a runtime error will result
        if ($upgradecode -ne "") {
            # RelatedProducts lists products that share the same upgrade code (they are related)
            $upgrades = $installer.RelatedProducts($upgradecode)
            foreach ($u in $upgrades) {
                $relatedproductcodes += $u
            }
        }
    }
    catch {
       # Our whole session object failed to instantiate, report error in export
       $upgradecode = "Error Accessing Data: $($_.Exception.Source), 0x$([Convert]::ToString($_.Exception.HResult,16))"
    }
    finally {
        # Crucial: Always release the session object in order to be able to continue with
        #          the next package regardless if there was an error or not (hence finally)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($session) | Out-Null
    }

    # Create html element listing all related product codes (if more than one)
    if ($relatedproductcodes.Count -gt 0) {
        $allupgrades = $relatedproductcodes -join "<br>"
    }
    
    # The MSI package details we want to output for this product in HTML format
    $htmloutput += "<tr><td>$p</td><td>$productname</td><td>$versionstring</td><td>$packagecode</td><td>$productcode</td><td>$upgradecode</td><td>$allupgrades</td><td>$assignment</td><td>$lcid</td></tr>`r`n"
 
    # Clean up things for next package
    $relatedproductcodes = @()
    $upgradecode = ""
    $allupgrades = ""

    # Show a progress bar for the package retrieval process
    $progress = [math]::Floor(($p / $totalpackages) * 100)
    Write-Progress -Activity "Package retrieval:" "$progress % Complete:" -percentComplete $progress;

    $p++
}

# Remove progress bar
Write-Progress -Activity "Package retrieval:" -Completed

# Status update
Write-Host "End of package retrieval..."

# Release Windows Installer COM object as early as possible
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($installer) | Out-Null

# Finalize the custom HTML output file content
$htmloutput += "</tbody></table></body></html>"

# Build output filename with computer name, date and time embedded in filename for custom HTML export
$filename = "MsiInfo_$($env:COMPUTERNAME)_$((Get-Date).Day).$((Get-Date).Month)(month).$((Get-Date).Year)_$((Get-Date).Hour)-$((Get-Date).Minute)-$((Get-Date).Second).html"

# Create HTML output file
Write-Host "Generating output file..."
$outputpath = Join-Path -Path $PSScriptRoot -ChildPath $filename
$Utf8BomEncoding = New-Object System.Text.UTF8Encoding(1) # Using Utf8 with BOM - for now...
[System.IO.File]::WriteAllLines($outputpath, $htmloutput, $Utf8BomEncoding)

# Open the custom, exported HTML output file in default browser
Start-Process $outputpath

# The script has completed
Write-Host "Execution complete."
Write-Host "The exported MSI package information will show in your default browser."

# Prevent Powershell window from closing
Read-Host -Prompt "Press Enter to exit"
