' https://stackoverflow.com/questions/17543132/how-can-i-resolve-msi-paths-in-vbscript
' On Error resume Next

Set installer = CreateObject("WindowsInstaller.Installer")
const READONLY = 0
Dim DirList

productcode = Trim(InputBox("Please paste or type in the product code you want to look up details for:", _
              "Find Product Details (test GUID provided):", "{766AD270-A684-43D6-AF9A-74165C9B5796}"))
If search = vbCancel Or Trim(productcode) = "" Then
   WScript.Quit(0)
End If

Set session = installer.OpenProduct(productcode)
session.DoAction("CostInitialize")
session.DoAction("CostFinalize")

set view = session.Database.OpenView("SELECT * FROM Directory")
view.Execute
set record = view.Fetch

Do until record is Nothing

    If Err.number <> 0 Then
       MsgBox Err.Description 
    End If
    
    ResolvedDir = session.Property(record.StringData(1))
    DirList = DirList + record.StringData(1) + " => " + ResolvedDir + vbCrLf
    set record = view.Fetch

Loop

WScript.Echo DirList
