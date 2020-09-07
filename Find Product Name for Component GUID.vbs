Dim installer : Set installer = CreateObject("WindowsInstaller.Installer")
Dim counter : counter = 1

' Get component GUID from user
componentguid = Trim(InputBox("Please specify component GUID to look up (sample provided, please replace):", "Component GUID:","{4AC30CE3-6D22-5D84-972C-81C5A4775C3D}"))
If componentguid = vbCancel Or Trim(componentguid) = "" Then
   WScript.Quit(0) ' User aborted
End If

' Get list of products that share the component specified (if any)
Set componentclients = installer.ComponentClients(componentguid)
If (Err.number <> 0) Then
   MsgBox "Invalid component GUID?", vbOKOnly, "An error occurred:"
   WScript.Quit(2) ' Critical error, abort
End If

' Show the products
For Each productcode in componentclients
   productname = installer.productinfo (productcode, "InstalledProductName")
   productlist = productlist & counter & " - Product Code: " & productcode & vbNewLine & "Product Name: " & productname & vbNewLine & vbNewLine
   counter = counter + 1
Next

message = "The below products share component GUID: " & componentguid & vbNewLine & vbNewLine

MsgBox message & productlist, vbOKOnly, "Products sharing the component GUID: "