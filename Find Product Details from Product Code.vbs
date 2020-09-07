'On Error Resume Next ' we ignore all errors
Set installer = CreateObject("WindowsInstaller.Installer")
search = Trim(InputBox("Please paste or type in the product code you want to look up details for:", _
              "Find Product Details (test GUID provided):", "{8BC4D6BF-C0CF-48EB-A229-FC692208DFF0}"))
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