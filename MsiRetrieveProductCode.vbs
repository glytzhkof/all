' Retrieve all ProductCodes (with ProductName and ProductVersion)
' Put on desktop and run from there. Output: "msiinfo.csv" on desktop itself

Set fso = CreateObject("Scripting.FileSystemObject")
Set output = fso.CreateTextFile("msiinfo.csv", True, True)
Set installer = CreateObject("WindowsInstaller.Installer")

On Error Resume Next ' we ignore all errors

For Each product In installer.ProductsEx("", "", 7)
   productcode = product.ProductCode
   name = product.InstallProperty("ProductName")
   version=product.InstallProperty("VersionString")
   output.writeline (productcode & ", " & name & ", " & version)
Next

output.Close
