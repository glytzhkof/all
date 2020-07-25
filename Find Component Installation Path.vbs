Set i = CreateObject("WindowsInstaller.Installer")
' Microsoft Visual C++ 2012 x86 Minimum Runtime - 11.0.50727
MsgBox i.ComponentPath("{2F73A7B2-E50E-39A6-9ABC-EF89E4C62E36}","{F5CBD6DC-5C9C-430E-83A7-179BA49988CD}")

' https://stackoverflow.com/a/63083079/129130
