On Error Resume Next

Public cmdline

' Sample Product Codes:
  ' Microsoft Visual C++ 2008 Redistributable - x86 9.0.30729.17: {9A25302D-30C0-39D9-BD6F-21E6EC160475}

productcode = InputBox("ProductCode for your MSI:", "ProductCode:","{9A25302D-30C0-39D9-BD6F-21E6EC160475}")
If productcode = vbCancel Or Trim(productcode) = "" Then
   WScript.Quit(0)
End If

' Arrays of current feature states
ReDim ADDLOCAL(-1), ADDSOURCE(-1), ADVERTISE(-1), REMOVE(-1)

Set installer = CreateObject("WindowsInstaller.Installer")
Set productfeatures = installer.Features(productcode)
If (Err.number <> 0) Then
   MsgBox "Failed to open MSI package. Invalid product code?", vbCritical, "Fatal error. Aborting:"
   WScript.Quit(2)
End If

' Spin over all product features detecting installation states
For Each feature In productfeatures

    featurestate = installer.FeatureState(productcode, feature)

    ' Using crazy VBScript arrays
    Select Case featurestate
       Case 1 ReDim Preserve ADVERTISE(UBound(ADVERTISE) + 1) : ADVERTISE(UBound(ADVERTISE)) = feature
       Case 2 ReDim Preserve REMOVE(UBound(REMOVE) + 1) : REMOVE(UBound(REMOVE)) = feature
       Case 3 ReDim Preserve ADDLOCAL(UBound(ADDLOCAL) + 1) : ADDLOCAL(UBound(ADDLOCAL)) = feature
       Case 4 ReDim Preserve ADDSOURCE(UBound(ADDSOURCE) + 1) : ADDSOURCE(UBound(ADDSOURCE)) = feature
       Case Else ' Errorstate MsgBox "Error for feature: " + feature
    End Select

Next

' Now add whatever feature you need to ADDLOCAL, here is just a sample:
ReDim Preserve ADDLOCAL(UBound(ADDLOCAL) + 1) : ADDLOCAL(UBound(ADDLOCAL)) = "MyNewFeature"

' Flatten arrays
If UBound(ADDLOCAL) > -1 Then cmdline = chr(34) + "ADDLOCAL=" + Join(ADDLOCAL, ",") + chr(34)
If UBound(REMOVE) > -1 Then cmdline = cmdline + + " " + chr(34) + "REMOVE=" + Join(REMOVE, ",") + chr(34)
If UBound(ADVERTISE) > -1 Then cmdline = cmdline + + " " + chr(34) + "ADVERTISE=" + Join(ADVERTISE, ",") + chr(34)
If UBound(ADDSOURCE) > -1 Then cmdline = cmdline + + " " + chr(34) + "ADDSOURCE=" + Join(ADDSOURCE, ",") + chr(34)

' Your current feature installstate translated to msiexec.exe command line parameters
Wscript.Echo cmdline ' MsgBox has 1024 character limit
