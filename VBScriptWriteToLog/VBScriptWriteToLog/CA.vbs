MsgBox "Hi!" ' Verify that the script runs

Const msiMessageTypeInfo = &H04000000

' create the message record
Set msgrec = Installer.CreateRecord(1)

' field 0 is the template
msgrec.StringData(0) = "Log: [1]"

' field 1, to be placed in [1] placeholder
msgrec.StringData(1) = "Calling LoggingTestVBS..."

' send message to running installer
Session.Message msiMessageTypeInfo, msgrec
