On Error Resume Next
Const msiOpenDatabaseModeReadOnly = 0

Dim installer : Set installer = Nothing
Set installer = Wscript.CreateObject("WindowsInstaller.Installer")
Dim counter : counter = 0

' Verify incoming drag and drop arguments
If WScript.Arguments.Count = 0 Then MsgBox "Drag and drop an MSI file onto the VBScript" End If
filename = Wscript.Arguments(0)
If (Right (LCase(filename),3) <> "msi") Then 
   WScript.Quit
End If

Dim database : Set database = installer.OpenDatabase(filename, msiOpenDatabaseModeReadOnly)
Dim table, view, record

Set view = database.OpenView("SELECT `Name` FROM _Tables")
view.Execute

Do
    Set record = view.Fetch
    If record Is Nothing Then Exit Do
    table = record.StringData(1)
    tables = tables + table + vbNewLine
    counter = counter + 1
Loop

MsgBox "Number of tables: " + CStr(counter) + vbNewLine + vbNewLine + tables

Set view = Nothing
