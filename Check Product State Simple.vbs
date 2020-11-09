Dim installer : Set installer = CreateObject("WindowsInstaller.Installer")
MsgBox installer.ProductState("{6C961B30-A670-8A05-3BFE-3947E84DD4E4}")

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