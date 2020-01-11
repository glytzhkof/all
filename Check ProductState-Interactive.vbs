Dim installer : Set installer = CreateObject("WindowsInstaller.Installer")
' Get product GUID from user
productguid = Trim(InputBox("Please specify product GUID (sample provided, please replace):", "Product GUID:","{FDB30193-FDA0-3DAA-ACCA-A75EEFE53607}"))
If productguid = vbCancel Or Trim(productguid) = "" Then
   WScript.Quit(0) ' User aborted
End If

Dim msg
Select Case installer.ProductState(productguid)
    Case -2
        msg = "Invalid GUID?"
    Case -1
        msg = "The product is neither advertised or installed."
    Case 1
        msg = "The product is advertised but not installed."
    Case 2
        msg = "The product is installed for a different user."
    Case 5 ' Normal installed status
        msg = "The product is installed for the current user."
    Case Else
        msg = "Unknown state."
End Select

MsgBox msg, vbOKOnly, "Product Installation Status:"

' Full list of states:

'Enum MsiInstallState
'Const msiInstallStateAbsent = 2
'Const msiInstallStateAdvertised = 1
'Const msiInstallStateBadConfig = -6
'Const msiInstallStateBroken = 0
'Const msiInstallStateDefault = 5
'Const msiInstallStateIncomplete = -5
'Const msiInstallStateInvalidArg = -2
'Const msiInstallStateLocal = 3
'Const msiInstallStateNotUsed = -7
'Const msiInstallStateRemoved = 1
'Const msiInstallStateSource = 4
'Const msiInstallStateSourceAbsent = -4
'Const msiInstallStateUnknown = -1
