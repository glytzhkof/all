' Run as admin.

Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
Dim installer : Set installer = CreateObject("WindowsInstaller.Installer")

'On Error Resume Next

Set htmloutput = fso.CreateTextFile("MsiInstallLocation.html", True) ': CheckCOMError

htmloutput.writeline ("<!DOCTYPE html>")
htmloutput.writeline ("<html lang='en'><head><title>MSI Package Estate Information:</title><meta charset='windows-1252'>")
htmloutput.writeline ("<style>body {font: 12px Calibri;}")
htmloutput.writeline ("table, td {border: 1px solid black;border-collapse: collapse;padding: 0.3em;vertical-align: text-top;}</style>")
htmloutput.writeline ("<table><thead><tr>")
htmloutput.writeline ("<th>Product Code</th><th>Product Name</th><th>Install Location</th>" )
htmloutput.writeline ("</tr></thead><tbody>")

Set products = installer.ProductsEx("", "", 7) 
   
For Each product In products
   productcode = product.ProductCode 
   name = product.InstallProperty("ProductName") 
   location = product.InstallProperty("InstallLocation")
   
   If (location = "") Then
      location = "(Not Found)"
   End If
   
   htmloutput.writeline ( "<tr><td>" & productcode & "</td>" & _
                          "<td>" & name & "</td>" & _
                          "<td>" & location & "</td></tr>")
Next

htmloutput.writeline ("</tbody></table></body></html>")
htmloutput.Close

On Error GoTo 0

MsgBox "Export done, please open MsiInstallLocation.html", vbOKOnly + vbSystemModal, "MSI Info Export Complete"

