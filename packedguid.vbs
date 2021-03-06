'# Convert msi product code to Registry Product key and vice-versa.


Do

    strInput = InputBox("Please enter either Product Code or Registry Product Key:" & vbCrlf & vbCrlf _
               & "(Enter codes with or without '{ }' and hyphens)" & vbCrlf & vbCrlf & strXtraText, "Enter Code", strOutput)

    If strInput = "" Then WScript.Quit


    strOutput = fProdCodeConvert(strInput)


    strXtraText = "Below is a Conversion from: " & vbCrlf & strInput

Loop



Function fProdCodeConvert(sCode)

    Dim sInput, sOutput

    sInput = sCode

    sInput = Replace(sInput, "{", "")

    sInput = Replace(sInput, "}", "")

    sInput = Replace(sInput, "-", "")

    sOutput = StrReverse(Mid(sInput, 1, 8)) & StrReverse(Mid(sInput, 9, 4)) & StrReverse(Mid(sInput, 13, 4)) & _
                StrReverse(Mid(sInput, 17, 2)) & StrReverse(Mid(sInput, 19, 2)) & StrReverse(Mid(sInput, 21, 2)) & _
                StrReverse(Mid(sInput, 23, 2)) & StrReverse(Mid(sInput, 25, 2)) & StrReverse(Mid(sInput, 27, 2)) & _
                StrReverse(Mid(sInput, 29, 2)) & StrReverse(Mid(sInput, 31, 2))



    If InStr(sCode, "{") = 0 And InStr(sCode, "}") = 0 And InStr(sCode, "-") = 0 Then

        sOutput = "{" & Left(sOutput, 8) & "-" & Mid(sOutput, 9,4) & "-" & Mid(sOutput, 13,4) & "-" _
                    & Mid(sOutput, 17,4) & "-" & Mid(sOutput, 21) & "}"

    End If

    fProdCodeConvert = sOutput

End Function