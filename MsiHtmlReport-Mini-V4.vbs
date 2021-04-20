' This script could create a lot of log files in the TEMP folder if your system has MSI logging enabled by default (MSI log policy enabled)
' MSI logging policy: HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\Installer
'
' Note: Right click and select "Copy" for highlighted table cell content - CTRL + C with focus won't work.
'
' On MSI logging:
' - http://www.installsite.org/pages/en/msifaq/a/1022.htm
' - https://stackoverflow.com/a/54458890/129130

Const msiUILevelNone = 2 : p = 1
Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
Dim installer : Set installer = CreateObject("WindowsInstaller.Installer")

On Error Resume Next

' See alternative code line just below:
Set htmloutput = fso.CreateTextFile("msiinfo.html", True)

' Change to this for machines with Unicode characters in product name:
'Set htmloutput = fso.CreateTextFile("msiinfo.html", True, True)

htmloutput.writeline ("<!DOCTYPE html>")
htmloutput.writeline ("<html lang='en'><head><title>MSI Package Estate Information:</title><meta charset='windows-1252'>")
htmloutput.writeline "<script>function init() { try { document.querySelectorAll('td').forEach(link => { link.addEventListener('mouseenter', function (event) {var range = document.createRange(); range.selectNodeContents(this); var sel = window.getSelection(); sel.removeAllRanges(); sel.addRange(range);});});} catch (error) { console.log(error); }}</script>"
htmloutput.writeline ("<style>body {font: 12px Calibri;}")
htmloutput.writeline ("table, td {border: 1px solid black;border-collapse: collapse;padding: 0.3em;vertical-align: text-top;}")
htmloutput.writeline ("table>*>tr>td:nth-child(2) { max-width: 300px;}")
htmloutput.writeline ("th {font: bold 18px Calibri;background-color: purple;border: 1px solid black;text-align: left;color: white;}")
htmloutput.writeline ("table th {position: sticky;top: -1px;}</style>") : htmloutput.WriteLine ("")
htmloutput.writeline ("</head><body  onload='init()'>")    
htmloutput.writeline ("<table><thead><tr>")
htmloutput.writeline ("<th>#</th><th>Product Name</th><th>Version</th><th>Package Code</th><th>Product Code</th><th>Upgrade Code</th><th>Related Product Codes</th><th>Scope</th>" )
htmloutput.writeline ("</tr></thead><tbody>")

MsgBox "This export may take quite some time to complete." + vbNewLine + vbNewLine + "Please click OK in this dialog and then wait for the completion message box before opening the report found on the desktop.", vbOKOnly + vbSystemModal, "MSI Info Export Starting"

Set products = installer.ProductsEx("", "", 7) 
installer.UILevel = msiUILevelNone

ReDim relatedproductcodes(-1)
   
For Each product In products
   productcode = product.ProductCode 
   name = product.InstallProperty("ProductName") 
   version = product.InstallProperty("VersionString")
   scope = product.InstallProperty("AssignmentType")

   Select Case scope
     Case 0 assignment = "User"
     Case 1 assignment = "Machine"
     Case Else assignment = "Unknown"
   End Select

   ' Get upgrade code via MSI session object (reads cached MSI database with applied transforms - apparently)
   Err.Clear : Set session = installer.OpenProduct(productcode) ' Can fail to apply transforms, then we just report error in export
   If Err.Number = 0 Then
       ' So far so good, we have our session object, but upgrade code can be missing      
       upgradecode = session.ProductProperty("UpgradeCode")
       ' Don't pass empty string to RelatedProducts, a runtime error will result
       If upgradecode <> "" Then
         Set upgrades = installer.RelatedProducts(upgradecode) 
         For Each u In upgrades
            ReDim Preserve relatedproductcodes(UBound(relatedproductcodes) + 1) : relatedproductcodes(UBound(relatedproductcodes)) = u
         Next
       End If      
     Else
      ' Our whole session object failed to instantiate, report error in export, clear error and continue with next package
      upgradecode = "Error Accessing Data: " & Err.Source & ", " & Hex(Err.Number) : Err.Clear
   End If
   Set session = Nothing ' Important
   
   If UBound(relatedproductcodes) > -1 Then allupgrades = Join(relatedproductcodes, "<br />")
   ReDim relatedproductcodes(-1)
   
   htmloutput.writeline ("<tr><td>" & p & "</td><td>" & _
                         product.InstallProperty("ProductName") & _
                         "</td><td>" & product.InstallProperty("VersionString") & "</td><td>" & product.InstallProperty("PackageCode") & "</td><td>" & product.ProductCode & "</td><td>" & _
                         upgradecode & "</td><td>" & allupgrades & "</td><td>" & assignment & "</td></tr>")
   
   p = p + 1

Next

On Error GoTo 0

htmloutput.writeline ("</tbody></table></body></html>")
htmloutput.Close

   MsgBox "Export done, please open msiinfo.html (press F5 if file is already open in browser)", vbOKOnly + vbSystemModal, "MSI Info Export Complete"
