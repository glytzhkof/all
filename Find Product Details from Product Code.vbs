'On Error Resume Next ' we ignore all errors
Set installer = CreateObject("WindowsInstaller.Installer")
search = RemoveWhiteSpace(InputBox("Please paste or type in the product code you want to look up details for:", _
              "Find Product Details (test GUID provided):", "{7DC387B8-E6A2-480C-8EF9-A6E51AE81C19}"))
If search = vbCancel Or Trim(search) = "" Then
   WScript.Quit(0)
End If

For Each product In installer.ProductsEx("", "", 7)
   If (product.ProductCode = search) Then
   'If (Trim(LCase(product.ProductCode)) = Trim(LCase(search))) Then
      MsgBox "Product Code: " & product.ProductCode & vbNewLine & _
             "Product Name: " & product.InstallProperty("ProductName") & vbNewLine & _
             "Product Version: " & product.InstallProperty("VersionString"), vbOKOnly, "Match Found:"
      Exit For
   End If
Next

MsgBox "Completed product scan.", vbOKOnly, "Scan Complete"

' Remove not only spaces but tabs and control characters
Function RemoveWhiteSpace(str)

	Set re = New RegExp
	re.Pattern = "\s+"
	re.Global  = True
	
	RemoveWhiteSpace = Trim(re.Replace(str, " "))

End Function