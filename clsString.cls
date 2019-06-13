Option Explicit
Private Declare Function GetTickCount Lib "Kernel32" () As Long

Private siLeft_Last As String, siRight_Last As String

Private Function CountChar(strText As String, strChar As String) As Long
    Dim lngPosition As Long
    If Len(strChar) = 0 Then
        CountChar = 0
        Exit Function
    End If
    lngPosition = InStr(1, strText, strChar, vbTextCompare)
    If lngPosition <> 0 Then
        CountChar = CountChar + 1
        Do
            lngPosition = InStr(lngPosition + 1, strText, strChar, vbTextCompare)
            If lngPosition <> 0 Then
                CountChar = CountChar + 1
            End If
        Loop Until lngPosition = 0
    End If
End Function
Private Function Evaluate_Expression(strExpression As String) As String
    Dim strText              As String, lngOpenParenthesis As Long, strAnswer As String
    Dim lngClosedParenthesis As Long, strLastText As String
    Dim lngParenthesis       As Long
    strText = strExpression
    lngOpenParenthesis = 0
    lngClosedParenthesis = 0
    Do
        lngOpenParenthesis = InStr(lngOpenParenthesis + 1, strText, "(", vbTextCompare)
        lngClosedParenthesis = InStr(lngOpenParenthesis + 1, strText, ")", vbTextCompare)
        lngParenthesis = InStr(lngOpenParenthesis + 1, strText, "(", vbTextCompare)
        If lngParenthesis < lngClosedParenthesis Then
            Do
                If InStr(lngParenthesis + 1, strText, "(", vbTextCompare) < lngClosedParenthesis And InStr(lngParenthesis + 1, strText, "(", vbTextCompare) <> 0 Then
                    lngParenthesis = InStr(lngParenthesis + 1, strText, "(", vbTextCompare)
                Else
                    Exit Do
                End If
            Loop While lngParenthesis < lngClosedParenthesis
            If lngParenthesis <> 0 Then
                lngOpenParenthesis = lngParenthesis
            End If
        End If
        If lngOpenParenthesis <> 0 And lngClosedParenthesis <> 0 Then
            strAnswer = Evaluate_Expression(Mid(strText, lngOpenParenthesis + 1, lngClosedParenthesis - (lngOpenParenthesis + 1)))
            strText = MidReplace(strText, lngOpenParenthesis, lngClosedParenthesis - lngOpenParenthesis + 1, strAnswer)
        End If
    Loop While lngOpenParenthesis <> 0 Or lngClosedParenthesis <> 0
    Do
        strLastText = strText
        strText = Replace_DivisionAndMultiplication(strText)
    Loop While strText <> strLastText
    Do
        strLastText = strText
        strText = Replace_Addition(strText)
    Loop While strText <> strLastText
    Evaluate_Expression = strText
End Function
Private Function MidReplace(strString As String, lngStart As Long, lngLength As Long, strReplace As String) As String
    Dim strFirst As String, strSecond As String
    strFirst = Mid(strString, 1, lngStart - 1)
    strSecond = Mid(strString, lngStart + lngLength, Len(strString) - (lngStart + lngLength) + 1)
    MidReplace = strFirst & strReplace & strSecond
End Function
Private Function RegularExpression_Find(strExpression As String, strPattern As String, lngOperation As Long)
    Dim regExpression As New RegExp, mtchMatch As Match, mtchMatchCol As MatchCollection
    Dim strArray()    As String, sngAnswer As Single
    On Error GoTo Die
    With regExpression
        .IgnoreCase = True
        .Global = True
        .Pattern = strPattern
        Set mtchMatchCol = .Execute(strExpression)
        For Each mtchMatch In mtchMatchCol
            If mtchMatch.Value <> "" Then
                Select Case lngOperation
                Case 0
                    strArray = Split(mtchMatch.Value, "/", -1, vbTextCompare)
                    sngAnswer = CSng(strArray(0)) / CSng(strArray(1))
                    strExpression = Strings.Replace(strExpression, mtchMatch.Value, CStr(sngAnswer), 1, 1)
                Case 1
                    strArray = Split(mtchMatch.Value, "*", -1, vbTextCompare)
                    sngAnswer = CSng(strArray(0)) * CSng(strArray(1))
                    strExpression = Strings.Replace(strExpression, mtchMatch.Value, CStr(sngAnswer), 1, 1)
                Case 2
                    strArray = Split(mtchMatch.Value, "+", -1, vbTextCompare)
                    sngAnswer = CSng(strArray(0)) + CSng(strArray(1))
                    strExpression = Strings.Replace(strExpression, mtchMatch.Value, CStr(sngAnswer), 1, 1)
                Case 3
                    strArray = Split(mtchMatch.Value, "-", -1, vbTextCompare)
                    sngAnswer = CSng(strArray(0)) - CSng(strArray(1))
                    strExpression = Strings.Replace(strExpression, mtchMatch.Value, CStr(sngAnswer), 1, 1)
                Case 4
                    strExpression = Strings.Replace(strExpression, mtchMatch.Value, Mid(mtchMatch.Value, 1, Len(mtchMatch.Value) - 1) & "*(", 1, 1, vbTextCompare)
                Case 5
                    strExpression = Strings.Replace(strExpression, mtchMatch.Value, ")*" & Mid(mtchMatch.Value, 2, Len(mtchMatch.Value) - 1), 1, 1, vbTextCompare)
                End Select
            End If
        Next
        Set mtchMatchCol = Nothing
    End With
    RegularExpression_Find = strExpression
    Exit Function
Die:
    RegularExpression_Find = strExpression
    On Error GoTo 0
End Function
Private Function Replace_Addition(strExpression As String) As String
    strExpression = RegularExpression_Find(strExpression, "([\d\.-])+\+([\d\.-])+", 2)
    Replace_Addition = strExpression
End Function
Private Function Replace_DivisionAndMultiplication(strExpression As String) As String
    strExpression = RegularExpression_Find(strExpression, "([\d\.-])+/([\d\.-])+", 0)
    strExpression = RegularExpression_Find(strExpression, "([\d\.-])+\*([\d\.-])+", 1)
    Replace_DivisionAndMultiplication = strExpression
End Function

Public Property Get iLeftLast(ByVal KeyString As String) As String: iLeftLast = siLeft_Last: End Property
Public Property Get iRightLast(ByVal KeyString As String) As String: iRightLast = siRight_Last: End Property

Public Function AddBegin(ByVal OriginString As String, ByVal ValueString As String) As String: AddBegin = (ValueString & OriginString): End Function
Public Function AddEnd(ByVal OriginString As String, ByVal ValueString As String) As String: AddEnd = (OriginString & ValueString): End Function
Public Function Buff(ByVal Number As Long, Optional ByVal Character As String = " ") As String: Buff = String(Number, Character): End Function
Public Function Collapse(ByVal OriginString As String) As String: Collapse = Strings.Replace(Strings.Replace(Strings.Replace(OriginString, " ", ""), vbCrLf, ""), vbTab, ""): End Function
Public Function Conv(ByVal OriginString As String, ByVal Conversion As VbStrConv) As String: Conv = StrConv(OriginString, Conversion): End Function
Public Function CountWithLimit(ByVal cStart As Integer, ByVal cAdd As Integer, ByVal cMax As Integer, Optional ByVal cMin As Integer = 0) As Integer
    CountWithLimit = cStart + cAdd
    If CountWithLimit < cMin Then
        CountWithLimit = cMax
    ElseIf CountWithLimit > cMax Then
        CountWithLimit = cMin
    End If
End Function
Public Function Crypt(ByVal StringValue As String, Optional ByVal Enable As Boolean = True, Optional ByVal StringKey As String = "", Optional ByVal StringTrimBy As String = " ") As String
    Dim i As Integer, i2 As Integer
    Dim RandAddKeyHex As String, RandAddKey As String
    Dim RandRandKeyHex As String, RandRandKey As String
    Dim RandKey As String, RandKeyHex As String
    Dim RandLengthKey As String, RandLengthKeyHex As String
    Dim ChrValue As Long, ChrKey As Long, ChrChar As String
    If Enable Then
        If StringValue = "" Then Exit Function
        If StringKey = "" Then StringKey = KeySeed
        RandRandKey = Rand(255)
        RandRandKeyHex = UCase(Hex(RandRandKey))
        If Len(RandRandKeyHex) = 1 Then RandRandKeyHex = "0" & RandRandKeyHex
        RandAddKey = Rand(255)
        RandAddKeyHex = UCase(Hex(RandAddKey))
        If Len(RandAddKeyHex) = 1 Then RandAddKeyHex = "0" & RandAddKeyHex
        RandKey = Rand(255)
        RandKeyHex = UCase(Hex(RandKey))
        If Len(RandKeyHex) = 1 Then RandKeyHex = "0" & RandKeyHex
        RandKey = (RandKey Xor RandRandKey)
        RandLengthKey = Len(StringValue)
        While (RandLengthKey >= 255)
            RandLengthKey = RandLengthKey / 255
        Wend
        RandLengthKey = (RandLengthKey Xor RandRandKey)
        RandLengthKeyHex = UCase(Hex(RandLengthKey))
        If Len(RandLengthKeyHex) = 1 Then RandLengthKeyHex = "0" & RandLengthKeyHex
        If Len(StringValue) >= 50000 Then Exit Function
        For i = 1 To Len(StringValue) Step 1
            If i2 < Len(StringKey) Then i2 = i2 + 1 Else i2 = 1
            ChrValue = Asc(Mid(StringValue, i, 1))
            ChrValue = (ChrValue Xor RandKey)
            ChrKey = Asc(Mid(StringKey, i2, 1))
            ChrKey = ChrKey + (Len(StringKey) - (i + i2))
            ChrChar = CStr(UCase(Hex(ChrValue Xor ChrKey)))
            If Len(ChrChar) = 1 Then ChrChar = "0" & ChrChar
            Crypt = Crypt & ChrChar
            If i <> Len(StringValue) Then Crypt = Crypt & StringTrimBy
        Next i
        'FINISH
        Crypt = RandRandKeyHex & StringTrimBy & Crypt
        Crypt = Crypt & StringTrimBy & RandLengthKeyHex
        Crypt = Crypt & StringTrimBy & RandKeyHex
        Crypt = Crypt & StringTrimBy & RandAddKeyHex
    Else
        If StringValue = "" Then Exit Function
        If Not InString(StringValue, StringTrimBy) Then Exit Function
        If StringKey = "" Then StringKey = KeySeed
        On Error GoTo Die
        'ADD
        RandAddKeyHex = Right(StringValue, 2)
        RandAddKey = CLng("&H" & RandAddKeyHex)
        StringValue = Left(StringValue, Len(StringValue) - Len(StringTrimBy & "xx"))
        'RAND BEGIN
        RandRandKeyHex = Left(StringValue, 2)
        RandRandKey = CLng("&H" & RandRandKeyHex)
        StringValue = Right(StringValue, Len(StringValue) - Len(StringTrimBy & "xx"))
        'RAND END
        RandKeyHex = Right(StringValue, 2)
        RandKey = CLng("&H" & RandKeyHex)
        StringValue = Left(StringValue, Len(StringValue) - Len(StringTrimBy & "xx"))
        RandKey = (RandKey Xor RandRandKey)
        'CHECK LENGTH
        RandLengthKeyHex = Right(StringValue, 2)
        RandLengthKey = CLng("&H" & RandLengthKeyHex)
        StringValue = Left(StringValue, Len(StringValue) - Len(StringTrimBy & "xx"))
        RandLengthKey = (RandLengthKey Xor RandRandKey)
        'DECRYPT
        While StringValue <> ""
            i = i + 1
            If InStr(1, StringValue, StringTrimBy) Then
                ChrValue = CLng("&H" & Left(StringValue, InStr(1, StringValue, StringTrimBy) - 1))
                StringValue = Right(StringValue, Len(StringValue) - (InStr(1, StringValue, StringTrimBy)))
            Else
                ChrValue = CLng("&H" & StringValue)
                StringValue = ""
            End If
            If i2 < Len(StringKey) Then i2 = i2 + 1 Else i2 = 1
            ChrKey = Asc(Mid(StringKey, i2, 1))
            ChrKey = ChrKey + (Len(StringKey) - (i + i2))
            ChrValue = ((ChrValue) Xor (ChrKey))
            ChrChar = Chr(ChrValue Xor RandKey)
            Crypt = Crypt & ChrChar
        Wend
        'FINISH
        While (RandLengthKey >= 255)
            RandLengthKey = RandLengthKey / 255
        Wend
        If RandLengthKey <> Len(Crypt) Then Crypt = ""
        On Error GoTo 0
        Exit Function
Die:
        On Error GoTo 0
        'HndlErr Err.Number, Err.description, "WHEXUnCrypt." & StringValue
        Crypt = ""
    End If
End Function
Public Function CutBegin(ByVal OriginString As String, ByVal iLen As Long) As String: CutBegin = Left(OriginString, iLen): End Function
Public Function CutEnd(ByVal OriginString As String, ByVal iLen As Long) As String: CutEnd = Right(OriginString, iLen): End Function
Public Function Format(ByVal OriginString As String, ByVal Pattern As String) As String: Format = Strings.Format(OriginString, Pattern): End Function
Public Function FormatColor(ByVal tStr As String) As String
    Dim cRed, cGreen, cBlue As Long
    If iLeft(tStr, "#") Then
        tStr = Right(tStr, Len(tStr) - Len("#"))
        cRed = CLng("&H" & Left(tStr, 2))
        cGreen = CLng("&H" & Mid(tStr, 3, 2))
        cBlue = CLng("&H" & Right(tStr, 2))
        tStr = RGB(cRed, cGreen, cBlue)
    ElseIf InString(tStr, ", ") Then
        cRed = Left(tStr, InStr(1, tStr, ", ") - 1)
        tStr = Right(tStr, Len(tStr) - (Len(cRed) + Len(", ")))
        cGreen = Left(tStr, InStr(1, tStr, ", ") - 1)
        tStr = Right(tStr, Len(tStr) - (Len(cGreen) + Len(", ")))
        cBlue = tStr
        tStr = RGB(cRed, cGreen, cBlue)
    End If
    FormatColor = tStr
End Function
Public Function HexToLong(ByVal HexValue As String) As Long
    If iLeft(HexValue, "0x") Then HexValue = "&H" & iLeft_Last
    If Not iLeft(HexValue, "&H") Then HexValue = "&H" & HexValue
    If iRight(HexValue, "&") Then HexValue = iRight_Last
    HexToLong = CLng(HexValue)
End Function
Public Function IIf(ByVal TrueValue As String, FalseValue As String, ByVal fValue As Variant, Optional sValue As Variant, Optional Method As String = "Boolean") As String
    Select Case UCase$(Method)
    Case "BOOLEAN": IIf = Interaction.IIf(fValue, TrueValue, FalseValue)
    Case "<": IIf = Interaction.IIf(fValue < sValue, TrueValue, FalseValue)
    Case ">": IIf = Interaction.IIf(fValue > sValue, TrueValue, FalseValue)
    Case "<>", "!=": IIf = Interaction.IIf(fValue <> sValue, TrueValue, FalseValue)
    Case "=": IIf = Interaction.IIf(fValue = sValue, TrueValue, FalseValue)
    Case "OR", "||": IIf = Interaction.IIf(fValue Or sValue, TrueValue, FalseValue)
    Case "AND", "&&": IIf = Interaction.IIf(fValue Or sValue, TrueValue, FalseValue)
    End Select
End Function

Public Function iLeft(ByVal StringWhere As String, ByVal StringSearch As String, Optional RespectCase As Boolean = True) As Boolean
    On Error Resume Next
    iLeft = (Strings.Left(StringWhere, Len(StringSearch)) = StringSearch)
    If Not RespectCase And Not iLeft Then iLeft = (Strings.Left(UCase(StringWhere), Len(StringSearch)) = UCase(StringSearch))
    If iLeft Then siLeft_Last = Right(StringWhere, Len(StringWhere) - Len(StringSearch))
    If Err.Number <> 0 Then iLeft = False
    On Error GoTo 0
End Function
Public Function InString(ByVal WhereSearch As String, ByVal WhoSearch As String, Optional StrictUcase As Boolean = False) As Boolean
    If Not StrictUcase Then _
       If InStr(1, UCase(WhereSearch), UCase(WhoSearch)) <> 0 Then InString = True Else InString = False _
   Else If InStr(1, WhereSearch, WhoSearch) <> 0 Then InString = True Else InString = False
End Function
Public Function iRight(ByVal StringWhere As String, ByVal StringSearch As String, Optional RespectCase As Boolean = True) As Boolean
    On Error Resume Next
    iRight = (Strings.Right(StringWhere, Len(StringSearch)) = StringSearch)
    If Not RespectCase Then iRight = (Strings.Right(UCase(StringWhere), Len(StringSearch)) = UCase(StringSearch))
    If iRight Then siRight_Last = Left(StringWhere, Len(StringWhere) - Len(StringSearch))
    If Err.Number <> 0 Then iRight = False
    On Error GoTo 0
End Function
Public Function isUnicode(ByVal OriginString As String) As Boolean: isUnicode = Not (Len(OriginString) = LenB(OriginString)): End Function
Public Function LongColorFromString(ByVal OriginString As String): LongColorFromString = FormatColor(OriginString): End Function
Public Function LongToHexColor(ByVal lngColor As Long) As String
    Dim hColor As String
    hColor = Right$("000000" & Hex(lngColor), 6)
    LongToHexColor = Mid$(hColor, 5, 2) & Mid$(hColor, 3, 2) & Mid$(hColor, 1, 2)
End Function
Public Function Low(ByVal OriginString As String) As String: Low = LCase(OriginString): End Function
Public Function Math(strExpression As String) As String
    Dim strText            As String, strLastText As String, regExpression As New RegExp
    Dim mtchMatch          As Match, mtchMatchCol As MatchCollection, lngCount As Long
    Dim lngOpenParenthesis As Long, lngClosedParenthesis As Long
    strText = strExpression
    strText = Strings.Replace(strExpression, " ", "", 1, -1, vbTextCompare)
    lngOpenParenthesis = CountChar(strText, "(")
    lngClosedParenthesis = CountChar(strText, ")")
    If lngOpenParenthesis <> lngClosedParenthesis Then
        Math = "error"
        Exit Function
    End If
    Do
        strLastText = strText
        strText = RegularExpression_Find(strText, "([\d\.\)])+\(", 4)
    Loop While strText <> strLastText
    Do
        strLastText = strText
        strText = RegularExpression_Find(strText, "\)([\d\.])+", 5)
    Loop While strText <> strLastText
    With regExpression
        .Global = True
        .IgnoreCase = True
        .MultiLine = True
        .Pattern = "[^\+\(]-"
        Set mtchMatchCol = .Execute(strText)
        lngCount = 1
        For Each mtchMatch In mtchMatchCol
            If mtchMatch.Value <> "" Then
                strText = MidReplace(strText, mtchMatch.FirstIndex + lngCount, mtchMatch.Length, Mid(mtchMatch.Value, 1, 1) & "+-")
                lngCount = lngCount + 1
            End If
        Next
        Set mtchMatchCol = Nothing
    End With
    Math = Evaluate_Expression(strText)
End Function
Public Function Ptr(ByVal OriginString As String) As String: Ptr = StrPtr(OriginString): End Function
Public Function Rand(Optional ByVal tMaxValue As Long = 1, Optional ByVal tStartValue As Long = 0, Optional ByVal tFreq As Integer = 2, Optional Reset As Boolean = True) As Long
    Dim tStr       As Double
    Dim tStr2      As Double
    Static tRandNE As Long
    If Reset Then
        If tFreq > 20 Then tFreq = 20
        If tRandNE > 255 Then tRandNE = 0
        tRandNE = (tRandNE * tFreq) + tFreq
    End If
    If Reset Then Randomize
    tStr2 = CStr(GetTickCount + tRandNE)
    tStr2 = Right$(tStr2, tFreq)
    tStr2 = (100 * (tFreq - 1)) - tStr2
    tStr = tStr2 / 100
    Dim tSysR As Long, GlR As Long
    tSysR = ((Rnd() * (tMaxValue - tStartValue)) + tStartValue)
    GlR = (tStr * (tMaxValue - tStartValue)) + tStartValue
    If tSysR > GlR Then
        Rand = tSysR - GlR
    ElseIf tSysR < GlR Then
        Rand = GlR - tSysR
    Else
        Rand = tStartValue
    End If
End Function
Public Function Remove(ByVal OriginString As String, ByVal ValueString As String, Optional ByVal Count As Long = 0) As String
    If Count = 0 Then
        Remove = Strings.Replace(OriginString, ValueString, "")
    ElseIf Count > 0 Then
        Remove = Strings.Replace(OriginString, ValueString, "", 1, Count)
    ElseIf Count < 0 Then
        Remove = StrReverse(Strings.Replace(StrReverse(OriginString), StrReverse(ValueString), "", , (Count * -1)))
    End If
End Function
Public Function RemoveBegin(ByVal OriginString As String, ByVal iLen As Long) As String: RemoveBegin = Right(OriginString, Len(OriginString) - iLen): End Function
Public Function RemoveEnd(ByVal OriginString As String, ByVal iLen As Long) As String: RemoveEnd = Left(OriginString, Len(OriginString) - iLen): End Function
Public Function Replace(ByVal OriginString As String, ByVal SourceString As String, ByVal ByString As String, Optional ByVal Count As Long = 0) As String
    If Count = 0 Then
        Replace = Strings.Replace(OriginString, SourceString, ByString)
    ElseIf Count > 0 Then
        Replace = Strings.Replace(OriginString, SourceString, ByString, 1, Count)
    ElseIf Count < 0 Then
        Replace = StrReverse(Strings.Replace(StrReverse(OriginString), StrReverse(SourceString), StrReverse(ByString), , (Count * -1)))
    End If
End Function
Public Function Reverse(ByVal OriginString As String) As String: Reverse = StrReverse(OriginString): End Function
Public Function Trim(ByVal StringChain As String, Optional ByVal RespectCase As Boolean = False, Optional WithNumbers As Boolean = True) As String
    Dim i     As Integer
    Dim tStr  As String
    Dim tChar As String
    If Not RespectCase Then tStr = LCase(StringChain)
    For i = 0 To Len(tStr)
        Select Case Left$(tStr, 1)
        Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0"
            If WithNumbers Then tChar = Left$(tStr, 1)
        Case "?", "?", "?", "?", "?": tChar = "a"
        Case "?", "?", "?", "?": tChar = "e"
        Case "?", "?", "?", "?": tChar = "i"
        Case "?", "?", "?", "?": tChar = "u"
        Case "?": tChar = "n"
        Case "?": tChar = "c"
        Case " ": tChar = ""
        Case "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", _
           "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z"
            tChar = Left$(tStr, 1)
        End Select
        If tChar <> "" Then Trim = Trim + tChar
        If tStr <> "" Then tStr = Right$(tStr, Len(tStr) - 1)
        tChar = ""
    Next i
End Function
Public Function Unicode(ByVal OriginString As String) As String
    Dim i As Long, sChar As String
    For i = 1 To Len(OriginString)
        sChar = Mid(OriginString, i, 1)
        If (AscW(Mid(OriginString, i, 1)) > 255) Then Mid(OriginString, i, 1) = "?"
    Next i
End Function
Public Function Upp(ByVal OriginString As String) As String: Upp = UCase(OriginString): End Function
Public Function URLToServerName(ByVal URLString As String) As String
    If InString(URLString, ":///") Then URLString = Right(URLString, Len(URLString) - (InStr(1, URLString, ":///") + Len("://")))
    If InString(URLString, "://") Then URLString = Right(URLString, Len(URLString) - (InStr(1, URLString, "://") + Len(":/")))
    If InString(URLString, "/") Then URLString = Left(URLString, InStr(1, URLString, "/") - 1)
    If InString(URLString, ":") Then URLString = Left(URLString, InStr(1, URLString, ":") - 1)
    While (Len(URLString) - Len(Strings.Replace(URLString, ".", ""))) > 1
        URLString = Right(URLString, Len(URLString) - (InStr(1, URLString, ".")))
    Wend
    URLToServerName = URLString
End Function


