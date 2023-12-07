Attribute VB_Name = "modDataConversion"
' =-=-=-=-=-=-=-=-=-=-=-==-=-=-=-=-=-=-=-=-=-=-==-=-=-=-=-=-=-=-=-=-=-==-=-=-=-=-=-=-=-=-=-=-=
' Class       : modDataConversion
' Description : Data Conversion-related routines
' Source      : Total Visual SourceBook 2000
'
' Procedures  :
'              fnASCIIByteToEBCDICByte(ByVal bytByte As Byte) As Byte
'              fnAtLeast(ByVal varCheckVal As Variant, ByVal varMinVal As Variant) _
'              fnAtMost(ByVal varCheckVal As Variant, ByVal varMaxVal As Variant) _
'              fnBinaryStringToHexString(strBinary As String) As String
'              fnConvertLongToBinaryString(ByVal lngNumber As Long, _
'              fnConvertVarToRoundLong(varIn As Variant) As Long
'              fnCurrencyToText(ByVal dblAmount As Double, _
'              fnDelimitSendKeys(ByVal strIn As String) As String
'              fnEBCDICByteToASCIIByte(ByVal bytByte As Byte) As Byte
'              fnNullIfZero(ByVal dblTest As Double) As Variant
'              fnNullIfZeroOrAll(ByVal varTest As Variant) As Variant
'              fnNullIfZLS(ByVal varIn As Variant, _
'              fnNullIfZLSOrAll(ByVal varIn As Variant, _
'              fnNumberToRoman(ByVal intIn As Integer) As String
'              fnOctalStringToDecimal(ByVal strOctal As String) As Long
'              fnOverpunchedStringToNumber(ByVal strNum As String, _
'              fnPhoneLetterToDigit(ByVal chrIn As String) As Integer
'              fnRGBToHTMLColor(ByVal lngRGB As Long) As String
'              fnVarToCurrency(ByVal varIn As Variant) As Currency
'              fnVarToDouble(varIn As Variant) As Double
'              fnVarToInteger(ByVal varIn As Variant) As Integer
'              fnVarToLong(ByVal varIn As Variant) As Long
'              fnVarToString(varIn As Variant) As String
'              fnZeroIfNull(ByVal varTest As Variant) As Double
'              fnZLSIfNull(ByVal varTest As Variant) As String'
'
' Modified:
'
'   Version Date     Who   What
'   ------- -------- ---   -------------------------------------------------------------------
'   1.0     04/04/02 BAW   (Phase2C) Added to project, updated comments & error handlers.
' =-=-=-=-=-=-=-=-=-=-=-==-=-=-=-=-=-=-=-=-=-=-==-=-=-=-=-=-=-=-=-=-=-==-=-=-=-=-=-=-=-=-=-=-=

Option Explicit
Option Compare Binary
Private Const mcstrName As String = "modDataConversion."


'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Function fnASCIIByteToEBCDICByte(ByVal bytByte As Byte) As Byte
    ' Comments  : Converts between ASCII character values and
    '             EBCDIC character values
    ' Parameters: bytByte - The ASCII value to convert
    ' Returns   : The converted value
    ' Source    : Total Visual SourceBook 2000
    '
    Dim abytLookup(255) As Byte
    Const cstrCurrentProc As String = "fnASCIIByteToEBCDICByte"

    On Error GoTo PROC_ERR

    ' There is no clean formula for iteratively calculating the byte
    ' values for this conversion. Because of this, we build the
    ' conversion lookup table into an array.
    abytLookup(0) = 0
    abytLookup(1) = 1
    abytLookup(2) = 2
    abytLookup(3) = 3
    abytLookup(4) = 55
    abytLookup(5) = 45
    abytLookup(6) = 46
    abytLookup(7) = 47
    abytLookup(8) = 22
    abytLookup(9) = 5
    abytLookup(10) = 37
    abytLookup(11) = 11
    abytLookup(12) = 12
    abytLookup(13) = 13
    abytLookup(14) = 14
    abytLookup(15) = 15
    abytLookup(16) = 16
    abytLookup(17) = 17
    abytLookup(18) = 18
    abytLookup(19) = 19
    abytLookup(20) = 60
    abytLookup(21) = 61
    abytLookup(22) = 50
    abytLookup(23) = 38
    abytLookup(24) = 24
    abytLookup(25) = 25
    abytLookup(26) = 63
    abytLookup(27) = 39
    abytLookup(28) = 28
    abytLookup(29) = 29
    abytLookup(30) = 30
    abytLookup(31) = 31
    abytLookup(32) = 64
    abytLookup(33) = 79
    abytLookup(34) = 127
    abytLookup(35) = 123
    abytLookup(36) = 91
    abytLookup(37) = 108
    abytLookup(38) = 80
    abytLookup(39) = 125
    abytLookup(40) = 77
    abytLookup(41) = 93
    abytLookup(42) = 92
    abytLookup(43) = 78
    abytLookup(44) = 107
    abytLookup(45) = 96
    abytLookup(46) = 75
    abytLookup(47) = 97
    abytLookup(48) = 240
    abytLookup(49) = 241
    abytLookup(50) = 242
    abytLookup(51) = 243
    abytLookup(52) = 244
    abytLookup(53) = 245
    abytLookup(54) = 246
    abytLookup(55) = 247
    abytLookup(56) = 248
    abytLookup(57) = 249
    abytLookup(58) = 122
    abytLookup(59) = 94
    abytLookup(60) = 76
    abytLookup(61) = 126
    abytLookup(62) = 110
    abytLookup(63) = 111
    abytLookup(64) = 124
    abytLookup(65) = 193
    abytLookup(66) = 194
    abytLookup(67) = 195
    abytLookup(68) = 196
    abytLookup(69) = 197
    abytLookup(70) = 198
    abytLookup(71) = 199
    abytLookup(72) = 200
    abytLookup(73) = 201
    abytLookup(74) = 209
    abytLookup(75) = 210
    abytLookup(76) = 211
    abytLookup(77) = 212
    abytLookup(78) = 213
    abytLookup(79) = 214
    abytLookup(80) = 215
    abytLookup(81) = 216
    abytLookup(82) = 217
    abytLookup(83) = 226
    abytLookup(84) = 227
    abytLookup(85) = 228
    abytLookup(86) = 229
    abytLookup(87) = 230
    abytLookup(88) = 231
    abytLookup(89) = 232
    abytLookup(90) = 233
    abytLookup(91) = 74
    abytLookup(92) = 224
    abytLookup(93) = 90
    abytLookup(94) = 95
    abytLookup(95) = 109
    abytLookup(96) = 121
    abytLookup(97) = 129
    abytLookup(98) = 130
    abytLookup(99) = 131
    abytLookup(100) = 132
    abytLookup(101) = 133
    abytLookup(102) = 134
    abytLookup(103) = 135
    abytLookup(104) = 136
    abytLookup(105) = 137
    abytLookup(106) = 145
    abytLookup(107) = 146
    abytLookup(108) = 147
    abytLookup(109) = 148
    abytLookup(110) = 149
    abytLookup(111) = 150
    abytLookup(112) = 151
    abytLookup(113) = 152
    abytLookup(114) = 153
    abytLookup(115) = 162
    abytLookup(116) = 163
    abytLookup(117) = 164
    abytLookup(118) = 165
    abytLookup(119) = 166
    abytLookup(120) = 167
    abytLookup(121) = 168
    abytLookup(122) = 169
    abytLookup(123) = 192
    abytLookup(124) = 106
    abytLookup(125) = 208
    abytLookup(126) = 161
    abytLookup(127) = 7
    abytLookup(128) = 32
    abytLookup(129) = 33
    abytLookup(130) = 34
    abytLookup(131) = 35
    abytLookup(132) = 36
    abytLookup(133) = 21
    abytLookup(134) = 6
    abytLookup(135) = 23
    abytLookup(136) = 40
    abytLookup(137) = 41
    abytLookup(138) = 42
    abytLookup(139) = 43
    abytLookup(140) = 44
    abytLookup(141) = 9
    abytLookup(142) = 10
    abytLookup(143) = 27
    abytLookup(144) = 48
    abytLookup(145) = 49
    abytLookup(146) = 26
    abytLookup(147) = 51
    abytLookup(148) = 52
    abytLookup(149) = 53
    abytLookup(150) = 54
    abytLookup(151) = 8
    abytLookup(152) = 56
    abytLookup(153) = 57
    abytLookup(154) = 58
    abytLookup(155) = 59
    abytLookup(156) = 4
    abytLookup(157) = 20
    abytLookup(158) = 62
    abytLookup(159) = 225
    abytLookup(160) = 65
    abytLookup(161) = 66
    abytLookup(162) = 67
    abytLookup(163) = 68
    abytLookup(164) = 69
    abytLookup(165) = 70
    abytLookup(166) = 71
    abytLookup(167) = 72
    abytLookup(168) = 73
    abytLookup(169) = 81
    abytLookup(170) = 82
    abytLookup(171) = 83
    abytLookup(172) = 84
    abytLookup(173) = 85
    abytLookup(174) = 86
    abytLookup(175) = 87
    abytLookup(176) = 88
    abytLookup(177) = 89
    abytLookup(178) = 98
    abytLookup(179) = 99
    abytLookup(180) = 100
    abytLookup(181) = 101
    abytLookup(182) = 102
    abytLookup(183) = 103
    abytLookup(184) = 104
    abytLookup(185) = 105
    abytLookup(186) = 112
    abytLookup(187) = 113
    abytLookup(188) = 114
    abytLookup(189) = 115
    abytLookup(190) = 116
    abytLookup(191) = 117
    abytLookup(192) = 118
    abytLookup(193) = 119
    abytLookup(194) = 120
    abytLookup(195) = 128
    abytLookup(196) = 138
    abytLookup(197) = 139
    abytLookup(198) = 140
    abytLookup(199) = 141
    abytLookup(200) = 142
    abytLookup(201) = 143
    abytLookup(202) = 144
    abytLookup(203) = 154
    abytLookup(204) = 155
    abytLookup(205) = 156
    abytLookup(206) = 157
    abytLookup(207) = 158
    abytLookup(208) = 159
    abytLookup(209) = 160
    abytLookup(210) = 170
    abytLookup(211) = 171
    abytLookup(212) = 172
    abytLookup(213) = 173
    abytLookup(214) = 174
    abytLookup(215) = 175
    abytLookup(216) = 176
    abytLookup(217) = 177
    abytLookup(218) = 178
    abytLookup(219) = 179
    abytLookup(220) = 180
    abytLookup(221) = 181
    abytLookup(222) = 182
    abytLookup(223) = 183
    abytLookup(224) = 184
    abytLookup(225) = 185
    abytLookup(226) = 186
    abytLookup(227) = 187
    abytLookup(228) = 188
    abytLookup(229) = 189
    abytLookup(230) = 190
    abytLookup(231) = 191
    abytLookup(232) = 202
    abytLookup(233) = 203
    abytLookup(234) = 204
    abytLookup(235) = 205
    abytLookup(236) = 206
    abytLookup(237) = 207
    abytLookup(238) = 218
    abytLookup(239) = 219
    abytLookup(240) = 220
    abytLookup(241) = 221
    abytLookup(242) = 222
    abytLookup(243) = 223
    abytLookup(244) = 234
    abytLookup(245) = 235
    abytLookup(246) = 236
    abytLookup(247) = 237
    abytLookup(248) = 238
    abytLookup(249) = 239
    abytLookup(250) = 250
    abytLookup(251) = 251
    abytLookup(252) = 252
    abytLookup(253) = 253
    abytLookup(254) = 254
    abytLookup(255) = 255

    ' After the table is initialized, the return value is a simple lookup
    fnASCIIByteToEBCDICByte = abytLookup(bytByte)
PROC_EXIT:
    On Error GoTo 0     ' disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Function
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Function



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Function fnAtLeast(ByVal varCheckVal As Variant, ByVal varMinVal As Variant) _
    As Variant
    ' Comments  : Returns the greater of two values
    ' Parameters: varCheckVal - value to check
    '             varMinVal - minimum allowed value
    ' Returns   : If the value is less than the minimum,
    '             the minimum is returned. Otherwise the
    '             original value is returned
    ' Source    : Total Visual SourceBook 2000
    '
    Const cstrCurrentProc As String = "fnAtLeast"
    On Error GoTo PROC_ERR

    If varCheckVal < varMinVal Then
        fnAtLeast = varMinVal
    Else
        fnAtLeast = varCheckVal
    End If
PROC_EXIT:
    On Error GoTo 0     ' disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Function
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Function




'////////////////////////////////////////////////////////////////////////////////////////////////
Public Function fnAtMost(ByVal varCheckVal As Variant, ByVal varMaxVal As Variant) _
    As Variant
    ' Comments  : Returns the lesser of two values
    ' Parameters: varCheckVal - value to check
    '             varMaxVal - maximum allowed value
    ' Returns   : If the value is greater than the maximum,
    '             the maximum is returned. Otherwise the
    '             original value is returned
    ' Source    : Total Visual SourceBook 2000
    '
    Const cstrCurrentProc As String = "fnAtMost"
    On Error GoTo PROC_ERR

    If varCheckVal > varMaxVal Then
        fnAtMost = varMaxVal
    Else
        fnAtMost = varCheckVal
    End If
PROC_EXIT:
    On Error GoTo 0     ' disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Function
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Function



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Function fnBinaryStringToHexString(ByVal strBinary As String) As String
    ' Comments : Converts a string representation of a binary number
    ' to a string representation of a hexadecimal number
    ' Parameters: strBinary - string representation of the binary number
    ' Returns : string representation of the hexadecimal equivalent
    ' Source : Total Visual SourceBook 2000
    '
    Const cstrCurrentProc   As String = "fnBinaryStringToHexString"
    Dim strResult           As String
    Dim lngTmp              As Long
    Dim strTmp              As String
    Dim lngMask             As Long
    Dim lngCounter          As Long

    strTmp = Right$(String$(16, "0") + strBinary, 16)

    If Left$(strTmp, 1) = "1" Then
        lngTmp = &H8000
    End If

    lngMask = &H4000

    For lngCounter = 2 To 16
        If Mid$(strTmp, lngCounter, 1) = "1" Then
            lngTmp = lngTmp Or lngMask
        End If

        lngMask = lngMask \ 2

    Next lngCounter

    strResult = Hex$(lngTmp)
    fnBinaryStringToHexString = strResult
PROC_EXIT:
    On Error GoTo 0     ' disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Function
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Function



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Function fnConvertLongToBinaryString(ByVal lngNumber As Long, _
    Optional intDigits As Integer = 0) As String
    ' Comments  : Converts a long number to a string representation
    '             of a binary number
    ' Parameters: lngNumber - A long number to convert to binary
    '             intDigits - The minimum number of digits to return
    ' Returns   : String representation of a binary number
    ' Source    : Total Visual SourceBook 2000
    '
    Const cstrCurrentProc As String = "fnConvertLongToBinaryString"
    Dim lngCounter As Long
    Dim strTemp As String

    On Error GoTo PROC_ERR

    ' Check to see if number is negative
    If lngNumber < 0 Then
        strTemp = strTemp & "1"
    Else
        strTemp = strTemp & "0"
    End If

    ' Convert each bit into "0" or "1"
    For lngCounter = 30 To 0 Step -1
        If lngNumber And (2 ^ lngCounter) Then
            strTemp = strTemp & "1"
        Else
            strTemp = strTemp & "0"
        End If
    Next

    ' Return the result
    fnConvertLongToBinaryString = strTemp
PROC_EXIT:
    On Error GoTo 0     ' disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Function
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Function



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Function fnConvertVarToRoundLong(ByVal varIn As Variant) As Long
    ' Comments   : Converts the passed variant to a long integer and
    '              rounds it using arithmetic rounding. Nulls are returned
    '              as 0.
    ' Parameters : varIn - number to convert/round
    ' Returns    : Long integer
    ' Source     : Total Visual SourceBook 2000
    '
    Const cstrCurrentProc   As String = "fnConvertVarToRoundLong"
    Dim dbl                 As Double

    On Error GoTo PROC_ERR

    If IsNull(varIn) Then
        fnConvertVarToRoundLong = 0
    Else
        If Not IsNumeric(varIn) Then
            fnConvertVarToRoundLong = 0
        Else
            dbl = varIn + 0.5
            fnConvertVarToRoundLong = CLng(dbl)
        End If
    End If
PROC_EXIT:
    On Error GoTo 0     ' disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Function
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Function



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Function fnCurrencyToText(ByVal dblAmount As Double, _
    ByVal intLength As Integer) As String
    ' Comments  : Converts a number to spelled out text with padding/length
    '             options.
    ' Parameters: dblAmount - Dollar amount to convert
    '             intLength - Length of string to create (pads with trailing
    '             asterisks to fill to length). If text exceeds the given
    '             length, the numeric representation is given. If the
    '             intLength parameter is set to 0, the full string is
    '             returned without padding.
    ' Returns   : String representation of the amount
    ' Source    : Total Visual SourceBook 2000
    '
    Const cstrCurrentProc As String = "fnCurrencyToText"
    Dim strText As String
    Dim strCurrFormat As String
    Dim intLowDigit As Integer
    Dim strCents As String
    Dim intGroup As Integer
    Dim intDigit1 As Integer
    Dim intDigit2 As Integer
    Dim intDigit3 As Integer
    Dim strSubText As String
    Dim intGroupCounter As Integer
    Dim intCounter As Integer

    On Error GoTo PROC_ERR

    ' Create the arrays
    ReDim astrOnes(1 To 19) As String
    ReDim astrTens(1 To 9) As String
    ReDim astrGroup(1 To 3) As String

    ' Fill the arrays.
    astrOnes(1) = "ONE"
    astrOnes(2) = "TWO"
    astrOnes(3) = "THREE"
    astrOnes(4) = "FOUR"
    astrOnes(5) = "FIVE"
    astrOnes(6) = "SIX"
    astrOnes(7) = "SEVEN"
    astrOnes(8) = "EIGHT"
    astrOnes(9) = "NINE"
    astrOnes(10) = "TEN"
    astrOnes(11) = "ELEVEN"
    astrOnes(12) = "TWELVE"
    astrOnes(13) = "THIRTEEN"
    astrOnes(14) = "FOURTEEN"
    astrOnes(15) = "FIFTEEN"
    astrOnes(16) = "SIXTEEN"
    astrOnes(17) = "SEVENTEEN"
    astrOnes(18) = "EIGHTTEEN"
    astrOnes(19) = "NINETEEN"
    astrTens(1) = "TEN"
    astrTens(2) = "TWENTY"
    astrTens(3) = "THIRTY"
    astrTens(4) = "FORTY"
    astrTens(5) = "FIFTY"
    astrTens(6) = "SIXTY"
    astrTens(7) = "SEVENTY"
    astrTens(8) = "EIGHTY"
    astrTens(9) = "NINETY"
    astrGroup(1) = "THOUSAND"
    astrGroup(2) = "MILLION"
    astrGroup(3) = "BILLION"

    ' Prepare the temp variable
    strText = ""

    ' Ensure amount is greater than zero
    If dblAmount > 0 Then

        ' Format the string
        strCurrFormat = Format$(dblAmount, "#,###.00")

        ' Get the lower digit part
        intLowDigit = InStr(strCurrFormat, ".") - 1

        ' Get the cents
        strCents = Mid$(strCurrFormat, intLowDigit + 2, 2)

        intGroup = 0

        ' Loop through lower digit part
        While intLowDigit > 0

            intDigit3 = CInt(Mid$(strCurrFormat, intLowDigit, 1))

            If intLowDigit > 1 Then
                intDigit2 = CInt(Mid$(strCurrFormat, intLowDigit - 1, 1))
            Else
                intDigit2 = 0
            End If

            If intLowDigit > 2 Then
                intDigit1 = CInt(Mid$(strCurrFormat, intLowDigit - 2, 1))
            Else
                intDigit1 = 0
            End If

            strSubText = vbNullString

            ' Get the hundreds
            If intDigit1 > 0 Then
                strSubText = astrOnes(intDigit1) & " HUNDRED "
            End If

            If intDigit2 > 0 Then

                ' Get the ones
                If intDigit2 = 1 Then
                    strSubText = strSubText & astrOnes(intDigit3 + 10) & " "
                Else
                    strSubText = strSubText & astrTens(intDigit2)

                    If intDigit3 > 0 Then
                        strSubText = strSubText & "-" & astrOnes(intDigit3)
                    End If

                    strSubText = strSubText & " "
                End If

            Else

                If intDigit3 > 0 Then
                    strSubText = strSubText & astrOnes(intDigit3) & " "
                End If

            End If

            ' Get the grouping
            If strSubText <> "" And intGroupCounter <> 0 Then
                strSubText = strSubText & astrGroup(intGroupCounter) & " "
            End If

            ' Concatenate the temp vars
            strText = strSubText & strText

            ' Move back through the number
            intLowDigit = intLowDigit - 4

            ' Increment the counter
            intGroupCounter = intGroupCounter + 1

        Wend

        ' Finalize the text
        strText = strText + "& " + strCents + "/100"

        ' Replace the place holder with "NO" cents string
        If Left$(strText, 1) = "&" Then
            strText = "NO " + strText
        End If

        ' Cleanup and pad
        If intLength > 0 Then
            If Len(strText) > intLength Then
                strText = strCurrFormat
            Else
                For intCounter = 1 To (intLength - Len(strText))
                    strText = strText + "*"
                Next intCounter
            End If
        End If
    End If

    ' Return the result
    fnCurrencyToText = strText
PROC_EXIT:
    On Error GoTo 0     ' disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Function
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Function



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Function fnDelimitSendKeys(ByVal strIn As String) As String
    ' Comments  : Fixes sendkeys statements by delimiting magic characters
    ' Parameters: strIn - string to fix
    ' Returns   : Fixed string
    ' Source    : Total Visual SourceBook 2000
    '
    Const cstrCurrentProc As String = "fnDelimitSendKeys"
    Dim strTmp As String
    Dim intCounter As Integer
    Dim chrTmp As String * 1

    On Error GoTo PROC_ERR

    For intCounter = 1 To Len(strIn)

        chrTmp = Mid$(strIn, intCounter, 1)

        If chrTmp = "+" Or chrTmp = "^" Or chrTmp = "%" Or chrTmp = "~" Then
            strTmp = strTmp & "{" & chrTmp & "}"
        Else
            strTmp = strTmp & chrTmp
        End If

    Next intCounter

    fnDelimitSendKeys = strTmp
PROC_EXIT:
    On Error GoTo 0     ' disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Function
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Function



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Function fnEBCDICByteToASCIIByte(ByVal bytByte As Byte) As Byte
    ' Comments  : Converts between EBCDIC character values and
    '             ASCII character values
    ' Parameters: bytByte - The EBCDIC value to convert
    ' Returns   : The converted value
    ' Source    : Total Visual SourceBook 2000
    '
    ' This is the lookup table used to do the conversion
    Const cstrCurrentProc As String = "fnEBCDICByteToASCIIByte"
    Dim abytLookup(255) As Byte

    On Error GoTo PROC_ERR

    ' There is no clean formula for iteratively calculating the byte
    ' values for this conversion. Because of this, we build the
    ' conversion lookup table into an array.
    abytLookup(0) = 0
    abytLookup(1) = 1
    abytLookup(2) = 2
    abytLookup(3) = 3
    abytLookup(4) = 156
    abytLookup(5) = 9
    abytLookup(6) = 134
    abytLookup(7) = 127
    abytLookup(8) = 151
    abytLookup(9) = 141
    abytLookup(10) = 142
    abytLookup(11) = 11
    abytLookup(12) = 12
    abytLookup(13) = 13
    abytLookup(14) = 14
    abytLookup(15) = 15
    abytLookup(16) = 16
    abytLookup(17) = 17
    abytLookup(18) = 18
    abytLookup(19) = 19
    abytLookup(20) = 157
    abytLookup(21) = 133
    abytLookup(22) = 8
    abytLookup(23) = 135
    abytLookup(24) = 24
    abytLookup(25) = 25
    abytLookup(26) = 146
    abytLookup(27) = 143
    abytLookup(28) = 28
    abytLookup(29) = 29
    abytLookup(30) = 30
    abytLookup(31) = 31
    abytLookup(32) = 128
    abytLookup(33) = 129
    abytLookup(34) = 130
    abytLookup(35) = 131
    abytLookup(36) = 132
    abytLookup(37) = 10
    abytLookup(38) = 23
    abytLookup(39) = 27
    abytLookup(40) = 136
    abytLookup(41) = 137
    abytLookup(42) = 138
    abytLookup(43) = 139
    abytLookup(44) = 140
    abytLookup(45) = 5
    abytLookup(46) = 6
    abytLookup(47) = 7
    abytLookup(48) = 144
    abytLookup(49) = 145
    abytLookup(50) = 22
    abytLookup(51) = 147
    abytLookup(52) = 148
    abytLookup(53) = 149
    abytLookup(54) = 150
    abytLookup(55) = 4
    abytLookup(56) = 152
    abytLookup(57) = 153
    abytLookup(58) = 154
    abytLookup(59) = 155
    abytLookup(60) = 20
    abytLookup(61) = 21
    abytLookup(62) = 158
    abytLookup(63) = 26
    abytLookup(64) = 32
    abytLookup(65) = 160
    abytLookup(66) = 161
    abytLookup(67) = 162
    abytLookup(68) = 163
    abytLookup(69) = 164
    abytLookup(70) = 165
    abytLookup(71) = 166
    abytLookup(72) = 167
    abytLookup(73) = 168
    abytLookup(74) = 91
    abytLookup(75) = 46
    abytLookup(76) = 60
    abytLookup(77) = 40
    abytLookup(78) = 43
    abytLookup(79) = 33
    abytLookup(80) = 38
    abytLookup(81) = 169
    abytLookup(82) = 170
    abytLookup(83) = 171
    abytLookup(84) = 172
    abytLookup(85) = 173
    abytLookup(86) = 174
    abytLookup(87) = 175
    abytLookup(88) = 176
    abytLookup(89) = 177
    abytLookup(90) = 93
    abytLookup(91) = 36
    abytLookup(92) = 42
    abytLookup(93) = 41
    abytLookup(94) = 59
    abytLookup(95) = 94
    abytLookup(96) = 45
    abytLookup(97) = 47
    abytLookup(98) = 178
    abytLookup(99) = 179
    abytLookup(100) = 180
    abytLookup(101) = 181
    abytLookup(102) = 182
    abytLookup(103) = 183
    abytLookup(104) = 184
    abytLookup(105) = 185
    abytLookup(106) = 124
    abytLookup(107) = 44
    abytLookup(108) = 37
    abytLookup(109) = 95
    abytLookup(110) = 62
    abytLookup(111) = 63
    abytLookup(112) = 186
    abytLookup(113) = 187
    abytLookup(114) = 188
    abytLookup(115) = 189
    abytLookup(116) = 190
    abytLookup(117) = 191
    abytLookup(118) = 192
    abytLookup(119) = 193
    abytLookup(120) = 194
    abytLookup(121) = 96
    abytLookup(122) = 58
    abytLookup(123) = 35
    abytLookup(124) = 64
    abytLookup(125) = 39
    abytLookup(126) = 61
    abytLookup(127) = 34
    abytLookup(128) = 195
    abytLookup(129) = 97
    abytLookup(130) = 98
    abytLookup(131) = 99
    abytLookup(132) = 100
    abytLookup(133) = 101
    abytLookup(134) = 102
    abytLookup(135) = 103
    abytLookup(136) = 104
    abytLookup(137) = 105
    abytLookup(138) = 196
    abytLookup(139) = 197
    abytLookup(140) = 198
    abytLookup(141) = 199
    abytLookup(142) = 200
    abytLookup(143) = 201
    abytLookup(144) = 202
    abytLookup(145) = 106
    abytLookup(146) = 107
    abytLookup(147) = 108
    abytLookup(148) = 109
    abytLookup(149) = 110
    abytLookup(150) = 111
    abytLookup(151) = 112
    abytLookup(152) = 113
    abytLookup(153) = 114
    abytLookup(154) = 203
    abytLookup(155) = 204
    abytLookup(156) = 205
    abytLookup(157) = 206
    abytLookup(158) = 207
    abytLookup(159) = 208
    abytLookup(160) = 209
    abytLookup(161) = 126
    abytLookup(162) = 115
    abytLookup(163) = 116
    abytLookup(164) = 117
    abytLookup(165) = 118
    abytLookup(166) = 119
    abytLookup(167) = 120
    abytLookup(168) = 121
    abytLookup(169) = 122
    abytLookup(170) = 210
    abytLookup(171) = 211
    abytLookup(172) = 212
    abytLookup(173) = 213
    abytLookup(174) = 214
    abytLookup(175) = 215
    abytLookup(176) = 216
    abytLookup(177) = 217
    abytLookup(178) = 218
    abytLookup(179) = 219
    abytLookup(180) = 220
    abytLookup(181) = 221
    abytLookup(182) = 222
    abytLookup(183) = 223
    abytLookup(184) = 224
    abytLookup(185) = 225
    abytLookup(186) = 226
    abytLookup(187) = 227
    abytLookup(188) = 228
    abytLookup(189) = 229
    abytLookup(190) = 230
    abytLookup(191) = 231
    abytLookup(192) = 123
    abytLookup(193) = 65
    abytLookup(194) = 66
    abytLookup(195) = 67
    abytLookup(196) = 68
    abytLookup(197) = 69
    abytLookup(198) = 70
    abytLookup(199) = 71
    abytLookup(200) = 72
    abytLookup(201) = 73
    abytLookup(202) = 232
    abytLookup(203) = 233
    abytLookup(204) = 234
    abytLookup(205) = 235
    abytLookup(206) = 236
    abytLookup(207) = 237
    abytLookup(208) = 125
    abytLookup(209) = 74
    abytLookup(210) = 75
    abytLookup(211) = 76
    abytLookup(212) = 77
    abytLookup(213) = 78
    abytLookup(214) = 79
    abytLookup(215) = 80
    abytLookup(216) = 81
    abytLookup(217) = 82
    abytLookup(218) = 238
    abytLookup(219) = 239
    abytLookup(220) = 240
    abytLookup(221) = 241
    abytLookup(222) = 242
    abytLookup(223) = 243
    abytLookup(224) = 92
    abytLookup(225) = 159
    abytLookup(226) = 83
    abytLookup(227) = 84
    abytLookup(228) = 85
    abytLookup(229) = 86
    abytLookup(230) = 87
    abytLookup(231) = 88
    abytLookup(232) = 89
    abytLookup(233) = 90
    abytLookup(234) = 244
    abytLookup(235) = 245
    abytLookup(236) = 246
    abytLookup(237) = 247
    abytLookup(238) = 248
    abytLookup(239) = 249
    abytLookup(240) = 48
    abytLookup(241) = 49
    abytLookup(242) = 50
    abytLookup(243) = 51
    abytLookup(244) = 52
    abytLookup(245) = 53
    abytLookup(246) = 54
    abytLookup(247) = 55
    abytLookup(248) = 56
    abytLookup(249) = 57
    abytLookup(250) = 250
    abytLookup(251) = 251
    abytLookup(252) = 252
    abytLookup(253) = 253
    abytLookup(254) = 254
    abytLookup(255) = 255

    ' After the table is initialized, the return value is a simple lookup
    fnEBCDICByteToASCIIByte = abytLookup(bytByte)
PROC_EXIT:
    On Error GoTo 0     ' disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Function
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Function



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Function fnNullIfZero(ByVal dblTest As Double) As Variant
    ' Comments  : Returns Null if the passed value is zero, otherwise returns
    '             the passed value.
    ' Parameters: lngTest - Value to test
    ' Returns   : Null or passed value
    ' Source    : Total Visual SourceBook 2000
    '
    Const cstrCurrentProc As String = "fnNullIfZero"
    On Error GoTo PROC_ERR

    ' CMP Modified this since it wasn't working.
    'If Len(dblTest & "") = 0 Then
    If dblTest = 0 Then
        fnNullIfZero = Null
    Else
        fnNullIfZero = dblTest
    End If

PROC_EXIT:
    On Error GoTo 0     ' disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Function
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Function



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Function fnNullIfZeroOrAll(ByVal varTest As Variant) As Variant
    ' Comments  : Returns Null if the passed value is zero or "--All--", otherwise returns
    '             the passed value.
    ' Parameters: lngTest - Value to test
    ' Returns   : Null or passed value
    ' Source    : Total Visual SourceBook 2000
    '
    Const cstrCurrentProc As String = "fnNullIfZeroOrAll"
    Dim dblResult As Double
    On Error GoTo PROC_ERR

    ' CMP modified this 6/4/02 since the old way wasn't working.
    If varTest = 0 Or (varTest = gcstrAllEntry) Then
        fnNullIfZeroOrAll = Null
    Else
        fnNullIfZeroOrAll = CDbl(varTest)
    End If

PROC_EXIT:
    On Error GoTo 0     ' disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Function
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Function



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Function fnNullIfZLS(ByVal varIn As Variant, _
    Optional ByVal bHandleEmbeddedQuotes As Boolean = False) As Variant
    ' Comments  : Returns Null if the passed value is a zero-length
    '             string (""), otherwise returns the passed value.
    '
    '             NOTE: If working with data to send to SQL Server to
    '                   do an Insert or Update, then you should use
    '                   fnQuotedOrNull( ) in modGeneral.bas.
    '
    ' Parameters:
    '       varIn (in)                 - Value to test
    '       bHandleEmbeddedQuotes (in) - Indicates whether to replace
    '                                    single quotes with two single quotes.
    '
    ' Returns   : Null or passed value
    ' Source    : Total Visual SourceBook 2000
    '
    Const cstrCurrentProc As String = "fnNullIfZLS"
    Dim varResult As Variant

    On Error GoTo PROC_ERR

    If Len(varIn & "") = 0 Then
        varResult = Null
    Else
        If bHandleEmbeddedQuotes Then
            varResult = Replace(varIn, "'", "''")
        End If
        varResult = varIn
    End If

    fnNullIfZLS = varResult
PROC_EXIT:
    On Error GoTo 0     ' disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Function
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Function



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Function fnNullIfZLSOrAll(ByVal varIn As Variant, _
    Optional ByVal bHandleEmbeddedQuotes As Boolean = False) As Variant
    ' Comments  : Returns Null if the passed value is a zero-length
    '             string ("") or "--All--", otherwise returns the passed value.
    '
    '             NOTE: If working with data to send to SQL Server to
    '                   do an Insert or Update, then you should use
    '                   fnQuotedOrNull( ) in modGeneral.bas.
    '
    ' Parameters:
    '       varIn (in)                 - Value to test
    '       bHandleEmbeddedQuotes (in) - Indicates whether to replace
    '                                    single quotes with two single quotes.
    '
    ' Returns   : Null or passed value
    ' Source    : Total Visual SourceBook 2000
    '
    Const cstrCurrentProc As String = "fnNullIfZLSOrAll"
    Dim varResult As Variant

    On Error GoTo PROC_ERR

    If Len(varIn & "") = 0 Or (varIn = gcstrAllEntry) Then
        varResult = Null
    Else
        If bHandleEmbeddedQuotes Then
            varResult = Replace(varIn, "'", "''")
        End If
        varResult = varIn
    End If

    fnNullIfZLSOrAll = varResult
PROC_EXIT:
    On Error GoTo 0     ' disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Function
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Function


'////////////////////////////////////////////////////////////////////////////////////////////////
Public Function fnNumberToRoman(ByVal intIn As Integer) As String
    ' Comments  : Converts the passed integer to Roman numerals
    ' Parameters: intIn - Value to convert
    ' Returns   : String
    ' Source    : Total Visual SourceBook 2000
    '
    Const cstrCurrentProc As String = "fnNumberToRoman"
    Dim intCounter As Integer
    Dim intDigit As Integer
    Dim strTmp As String
    Const cstrDigits As String = "IVXLCDM"

    On Error GoTo PROC_ERR

    intCounter = 1

    ' Loop through values in input value
    Do While intIn > 0

        ' Get  the current digit
        intDigit = intIn Mod 10

        intIn = intIn \ 10

        ' Build the temp string
        Select Case intDigit

            Case 1
                strTmp = Mid$(cstrDigits, intCounter, 1) & strTmp

            Case 2
                strTmp = Mid$(cstrDigits, intCounter, 1) & _
                    Mid$(cstrDigits, intCounter, 1) & strTmp

            Case 3
                strTmp = Mid$(cstrDigits, intCounter, 1) & _
                    Mid$(cstrDigits, intCounter, 1) & _
                    Mid$(cstrDigits, intCounter, 1) & strTmp

            Case 4
                strTmp = Mid$(cstrDigits, intCounter, 2) & strTmp

            Case 5
                strTmp = Mid$(cstrDigits, intCounter + 1, 1) & strTmp

            Case 6
                strTmp = Mid$(cstrDigits, intCounter + 1, 1) & _
                    Mid$(cstrDigits, intCounter, 1) & strTmp

            Case 7
                strTmp = Mid$(cstrDigits, intCounter + 1, 1) & _
                    Mid$(cstrDigits, intCounter, 1) & _
                    Mid$(cstrDigits, intCounter, 1) & strTmp

            Case 8
                strTmp = Mid$(cstrDigits, intCounter + 1, 1) & _
                    Mid$(cstrDigits, intCounter, 1) & _
                    Mid$(cstrDigits, intCounter, 1) & _
                    Mid$(cstrDigits, intCounter, 1) & strTmp

            Case 9
                strTmp = Mid$(cstrDigits, intCounter, 1) & _
                    Mid$(cstrDigits, intCounter + 2, 1) & strTmp

        End Select
        intCounter = intCounter + 2
    Loop

    fnNumberToRoman = strTmp
PROC_EXIT:
    On Error GoTo 0     ' disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Function
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Function



'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Function fnOctalStringToDecimal(ByVal strOctal As String) As Long
    ' Comments   : Converts the passed string representation of an octal
    '              number to a decimal long integer.
    ' Parameters : strOctal - String representation of octal number
    ' Returns    : Decimal value
    ' Source     : Total Visual SourceBook 2000
    '
    Const cstrCurrentProc As String = "fnOctalStringToDecimal"
    On Error GoTo PROC_ERR

    fnOctalStringToDecimal = Val("&O" & strOctal)
PROC_EXIT:
    On Error GoTo 0     ' disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Function
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Function



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Function fnOverpunchedStringToNumber(ByVal strNum As String, _
    ByVal intDecimals As Integer) As Variant
    ' Comments  : Converts a "zoned overpunch" number to a regular number
    ' Parameters: strNum - Zoned overpunch value to convert
    '             intDecimals - Number of decimal places
    ' Returns   : Converted number
    ' Source    : Total Visual SourceBook 2000
    '
    Const cstrCurrentProc As String = "fnOverpunchedStringToNumber"
    Dim intLen As Integer
    Dim intSign As Integer
    Dim dblOut As Double
    Dim strLast As String * 1
    Dim intLast As Integer

    On Error GoTo PROC_ERR

    ' Get the length of the string
    intLen = Len(Trim$(strNum))

    ' Get the last character
    strLast = Mid$(strNum, intLen, 1)

    ' Decide how to convert the last character
    Select Case strLast
        Case "A" To "I"
            intSign = 1
            intLast = Asc(strLast) - 65 + 1
        Case "J" To "R"
            intSign = -1
            intLast = Asc(strLast) - 74 + 1
        Case "{"
            intSign = 1
        Case "}"
            intSign = -1
        Case Else
            intSign = 1
            intLast = 9
            strNum = "9999999999999"
    End Select

    dblOut = Val(Left$(strNum, intLen - 1) & intLast) * intSign
    fnOverpunchedStringToNumber = dblOut * (10 ^ -(intDecimals))
PROC_EXIT:
    On Error GoTo 0     ' disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Function
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Function



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Function fnPhoneLetterToDigit(ByVal chrIn As String) As Integer
    ' Comments  : Converts a phone number letter to a number
    ' Parameters: chrIn - Letter to check. Must be in the range a-p
    '             or r-y. Q and Z are not valid phone letters.
    ' Returns   : Integer number
    ' Source    : Total Visual SourceBook 2000
    '
    Const cstrCurrentProc As String = "fnPhoneLetterToDigit"
    Dim intDigit As Integer
    Dim chrTmp As String * 1

    On Error GoTo PROC_ERR

    If chrIn <> "" Then

        ' Trim any excess characters
        chrTmp = LCase$(Left$(chrIn, 1))

        ' Make sure its a letter
        If chrTmp >= "a" And chrTmp <= "z" Then

            ' For historical reasons, Q is not a valid letter on a phone.
            ' Z is also left out.
            If chrTmp <> "q" And chrTmp <> "z" Then

                intDigit = Asc(UCase$(chrTmp))

                If intDigit > Asc("Q") Then
                    intDigit = intDigit - 1
                End If

                intDigit = (intDigit - Asc("A")) \ 3 + 2

                fnPhoneLetterToDigit = Trim$(CStr(intDigit))
            End If
        End If
    End If
PROC_EXIT:
    On Error GoTo 0     ' disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Function
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Function



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Function fnRGBToHTMLColor(ByVal lngRGB As Long) As String
    ' Comments  : Formats an RGB value into the hex format standard used
    '             in HTML.
    ' Parameters: lngRGB - the RGB value, or a VB-defined constant such
    '             as 'vbRed' that evaluates to an RGB value
    ' Returns   : The formatted hex value of the RGB color
    ' Source    : Total Visual SourceBook 2000
    '
    Const cstrCurrentProc As String = "fnRGBToHTMLColor"
    Dim strValue As String

    On Error GoTo PROC_ERR

    ' Break out individual color portions of the RGB value, and then
    ' get the hex value in the format HTML expects (rrggbb)
    strValue = Hex$( _
        (lngRGB And &HFF&) * &H10000 Or _
        (lngRGB And &HFF00&) Or _
        (lngRGB And &HFF0000) \ &H10000)

    ' Force leading zeroes, which VB's hex function drops
    strValue = String(6 - Len(strValue), "0") & strValue

    fnRGBToHTMLColor = strValue
PROC_EXIT:
    On Error GoTo 0     ' disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Function
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Function



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Function fnVarToCurrency(ByVal varIn As Variant) As Currency
    ' Comments   : Converts the passed variant to a currency value,
    '              0 if the passed value is Null.
    ' Parameters : varIn - Value to convert
    ' Returns    : Currency
    ' Source     : Total Visual SourceBook 2000
    '
    Const cstrCurrentProc As String = "fnVarToCurrency"
    On Error GoTo PROC_ERR

    If IsNull(varIn) Then
        fnVarToCurrency = 0
    Else
        If Not IsNumeric(varIn) Then
            fnVarToCurrency = 0
        Else
            fnVarToCurrency = CCur(varIn)
        End If
    End If
PROC_EXIT:
    On Error GoTo 0     ' disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Function
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Function



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Function fnVarToDouble(varIn As Variant) As Double
    ' Comments   : Converts the passed variant to a double, returning
    '              0 if the passed value is Null.
    ' Parameters : varIn - Value to convert
    ' Returns    : Double
    ' Source     : Total Visual SourceBook 2000
    '
    Const cstrCurrentProc As String = "fnVarToDouble"
    On Error GoTo PROC_ERR

    If IsNull(varIn) Then
        fnVarToDouble = 0
    Else
        If Not IsNumeric(varIn) Then
            fnVarToDouble = 0
        Else
            fnVarToDouble = varIn
        End If
    End If
PROC_EXIT:
    On Error GoTo 0     ' disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Function
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Function



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Function fnVarToInteger(ByVal varIn As Variant) As Integer
    ' Comments   : Converts the passed variant to an integer, returning
    '              0 if the passed value is Null
    ' Parameters : varIn - Value to convert
    ' Returns    : Integer
    ' Source     : Total Visual SourceBook 2000
    '
    Const cstrCurrentProc As String = "fnVarToInteger"
    On Error GoTo PROC_ERR

    If IsNull(varIn) Then
        fnVarToInteger = 0
    Else
        If Not IsNumeric(varIn) Then
            fnVarToInteger = 0
        Else
            fnVarToInteger = varIn
        End If
    End If
PROC_EXIT:
    On Error GoTo 0     ' disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Function
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Function



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Function fnVarToLong(ByVal varIn As Variant) As Long
    ' Comments   : Converts the passed variant to a long integer, returning
    '              0 if the passed value is Null
    ' Parameters : varIn - Value to convert
    ' Returns    : Long integer
    ' Source     : Total Visual SourceBook 2000
    '
    Const cstrCurrentProc As String = "fnVarToLong"
    On Error GoTo PROC_ERR

    If IsNull(varIn) Then
        fnVarToLong = 0
    Else
        If Not IsNumeric(varIn) Then
            fnVarToLong = 0
        Else
            fnVarToLong = varIn
        End If
    End If
PROC_EXIT:
    On Error GoTo 0     ' disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Function
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Function



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Function fnVarToString(varIn As Variant) As String
    ' Comments  : Converts the supplied variant to a string. Nulls are returned
    '             as a zero-length string ("")
    ' Parameters: varIn - Variant to convert
    ' Returns   : String
    ' Source    : Total Visual SourceBook 2000
    '
    Const cstrCurrentProc As String = "fnVarToString"
    On Error GoTo PROC_ERR

    If IsNull(varIn) Then
        fnVarToString = vbNullString
    Else
        fnVarToString = varIn
    End If
PROC_EXIT:
    On Error GoTo 0     ' disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Function
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Function



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Function fnZeroIfNull(ByVal varTest As Variant) As Double
    ' Comments  : Returns zero if Null is passed, otherwise returns the
    '             passed value.
    ' Parameters: varTest - Value to test
    ' Returns   : Zero or Null
    ' Source    : Total Visual SourceBook 2000
    '
    Const cstrCurrentProc As String = "fnZeroIfNull"
    Dim dblResult As Double

    On Error GoTo PROC_ERR

    If IsNull(varTest) Then
        dblResult = 0
    Else
        dblResult = varTest
    End If

    fnZeroIfNull = dblResult
PROC_EXIT:
    On Error GoTo 0     ' disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Function
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Function



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Function fnZLSIfNull(ByVal varTest As Variant) As String
    ' Comments  : Returns a zero-length string ("") if Null is passed,
    '             otherwise returns the passed value.
    ' Parameters: varTest - Value to test
    ' Returns   : If the value is Null, it returns a zero-length string,
    '             otherwise it returns the passed value.
    ' Source    : Total Visual SourceBook 2000
    '
    Const cstrCurrentProc As String = "fnZLSIfNull"
    Dim varResult As Variant

    On Error GoTo PROC_ERR

    If IsNull(varTest) Then
        varResult = vbNullString
    Else
        varResult = varTest
    End If

    fnZLSIfNull = varResult
PROC_EXIT:
    On Error GoTo 0     ' disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Function
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Function
