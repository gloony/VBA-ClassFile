Option Explicit
Private Declare Sub keybd_event Lib "user32.dll" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Public Sub Down(ByVal vkCode As Byte): Call keybd_event(vkCode, 0, 0, 0): End Sub
Public Sub Press(ByVal vkCode As Byte): Call keybd_event(vkCode, 0, 0, 0): Call keybd_event(vkCode, 0, &H2, 0): End Sub
Public Sub Send(ByVal TextString As String)
    If Left(TextString, 1) = "{" And Right(TextString, 1) = "}" Then
        Dim tStr As String
        tStr = Right(TextString, Len(TextString) - 1)
        tStr = Left(tStr, Len(tStr) - 1)
        Press CInt(tStr)
    ElseIf iLeft(TextString, "+") Then
        SendKeys "+" & iLeft_Last
    ElseIf iLeft(TextString, "^") Then
        SendKeys "^" & iLeft_Last
    ElseIf iLeft(TextString, "%") Then
        SendKeys "%" & iLeft_Last
    Else
        Dim i As Integer
        For i = 1 To Len(TextString) Step 1
            On Error Resume Next
            Press Asc(Mid(UCase(TextString), i, 1))
            On Error GoTo 0
        Next i
    End If
End Sub
Public Sub SendKeys(ByVal Text As String, Optional ByVal Wait As Boolean = False)
    On Error Resume Next
    Dim WshShell As Object
    Set WshShell = CreateObject("wscript.shell")
    WshShell.SendKeys Text, Wait
    Set WshShell = Nothing
    On Error GoTo 0
End Sub
Public Sub Up(Optional ByVal vkCode As Byte)
    If vkCode = 0 Then
        Dim i As Integer
        For i = 1 To 249 Step 1
            If GetKeyState(i) < 0 Then Call keybd_event(i, 0, &H2, 0)
        Next i
   Else
        keybd_event vkCode, 0, &H2, 0
    End If
End Sub
Public Sub Copy(Optional ByVal Wait As Boolean = False)
    Me.SendKeys "^C", Wait
End Sub
Public Sub Cut(Optional ByVal Wait As Boolean = False)
    Me.SendKeys "^X", Wait
End Sub
Public Sub Paste(Optional ByVal Wait As Boolean = False)
    Me.SendKeys "^V", Wait
End Sub

Public Function KeyCodeToString(ByVal vkCode As Long) As String
    Select Case vkCode
    Case 0: KeyCodeToString = "N/A"
    Case 1: KeyCodeToString = "mouse left"
    Case 2: KeyCodeToString = "mouse right"
    Case vbKeyA: KeyCodeToString = "a"
    Case vbKeyB: KeyCodeToString = "b"
    Case vbKeyC: KeyCodeToString = "c"
    Case vbKeyD: KeyCodeToString = "d"
    Case vbKeyE: KeyCodeToString = "e"
    Case vbKeyF: KeyCodeToString = "f"
    Case vbKeyG: KeyCodeToString = "g"
    Case vbKeyH: KeyCodeToString = "h"
    Case vbKeyI: KeyCodeToString = "i"
    Case vbKeyJ: KeyCodeToString = "j"
    Case vbKeyK: KeyCodeToString = "k"
    Case vbKeyL: KeyCodeToString = "l"
    Case vbKeyM: KeyCodeToString = "m"
    Case vbKeyN: KeyCodeToString = "n"
    Case vbKeyO: KeyCodeToString = "o"
    Case vbKeyP: KeyCodeToString = "p"
    Case vbKeyQ: KeyCodeToString = "q"
    Case vbKeyR: KeyCodeToString = "r"
    Case vbKeyS: KeyCodeToString = "s"
    Case vbKeyT: KeyCodeToString = "t"
    Case vbKeyU: KeyCodeToString = "u"
    Case vbKeyV: KeyCodeToString = "v"
    Case vbKeyW: KeyCodeToString = "w"
    Case vbKeyX: KeyCodeToString = "x"
    Case vbKeyY: KeyCodeToString = "y"
    Case vbKeyZ: KeyCodeToString = "z"
    Case vbKeyShift: KeyCodeToString = "shift"
    Case vbKeySpace: KeyCodeToString = "space"
    Case vbKeyTab: KeyCodeToString = "tab"
    Case vbKeyReturn: KeyCodeToString = "enter"
    Case vbKeyUp: KeyCodeToString = "up"
    Case vbKeyDown: KeyCodeToString = "down"
    Case vbKeyLeft: KeyCodeToString = "left"
    Case vbKeyRight: KeyCodeToString = "right"
    Case vbKey0: KeyCodeToString = "0"
    Case vbKey1: KeyCodeToString = "1"
    Case vbKey2: KeyCodeToString = "2"
    Case vbKey3: KeyCodeToString = "3"
    Case vbKey4: KeyCodeToString = "4"
    Case vbKey5: KeyCodeToString = "5"
    Case vbKey6: KeyCodeToString = "6"
    Case vbKey7: KeyCodeToString = "7"
    Case vbKey8: KeyCodeToString = "8"
    Case vbKey9: KeyCodeToString = "9"
    Case vbKeyNumpad0: KeyCodeToString = "Numpad0"
    Case vbKeyNumpad1: KeyCodeToString = "Numpad1"
    Case vbKeyNumpad2: KeyCodeToString = "Numpad2"
    Case vbKeyNumpad3: KeyCodeToString = "Numpad3"
    Case vbKeyNumpad4: KeyCodeToString = "Numpad4"
    Case vbKeyNumpad5: KeyCodeToString = "Numpad5"
    Case vbKeyNumpad6: KeyCodeToString = "Numpad6"
    Case vbKeyNumpad7: KeyCodeToString = "Numpad7"
    Case vbKeyNumpad8: KeyCodeToString = "Numpad8"
    Case vbKeyNumpad9: KeyCodeToString = "Numpad9"
    Case vbKeyEscape: KeyCodeToString = "esc"
    Case vbKeyDelete: KeyCodeToString = "del"
    Case vbKeyBack: KeyCodeToString = "backspace"
    Case vbKeyControl, 162, 163: KeyCodeToString = "control"
    Case vbKeyCapital: KeyCodeToString = "verr. maj"
    Case 160, 161: KeyCodeToString = "maj"
    Case vbKeyHome: KeyCodeToString = "home"
    Case vbKeyEnd: KeyCodeToString = "end"
    Case 18, 164: KeyCodeToString = "alt"
    Case 165: KeyCodeToString = "alt gr"
    Case 91, 92: KeyCodeToString = "windows"
    Case vbKeyInsert: KeyCodeToString = "insert"
    Case vbKeyPageDown: KeyCodeToString = "page down"
    Case vbKeyPageUp: KeyCodeToString = "page up"
    'Case vbKeyScrollLock: KeyCodeToString = "scroll lock"
    Case vbKeyPause: KeyCodeToString = "pause"
    Case vbKeySnapshot, &H2C: KeyCodeToString = "prnt scrn"
    Case vbKeyF1: KeyCodeToString = "f1"
    Case vbKeyF2: KeyCodeToString = "f2"
    Case vbKeyF3: KeyCodeToString = "f3"
    Case vbKeyF4: KeyCodeToString = "f4"
    Case vbKeyF5: KeyCodeToString = "f5"
    Case vbKeyF6: KeyCodeToString = "f6"
    Case vbKeyF7: KeyCodeToString = "f7"
    Case vbKeyF8: KeyCodeToString = "f8"
    Case vbKeyF9: KeyCodeToString = "f9"
    Case vbKeyF10: KeyCodeToString = "f10"
    Case vbKeyF11: KeyCodeToString = "f11"
    Case vbKeyF12: KeyCodeToString = "f12"
    Case vbKeyNumlock: KeyCodeToString = "num lock"
    Case 191: KeyCodeToString = "?"
    Case 221: KeyCodeToString = "^"
    Case 219: KeyCodeToString = "'"
    Case 93: KeyCodeToString = "menu"
    Case 223: KeyCodeToString = "$"
    Case 220: KeyCodeToString = "?"
    Case 222: KeyCodeToString = "?"
    Case 186: KeyCodeToString = "?"
    Case 192: KeyCodeToString = "?"
    Case 110, 190: KeyCodeToString = "."
    Case 188: KeyCodeToString = ","
    Case 109, 189: KeyCodeToString = "-"
    Case 107: KeyCodeToString = "+"
    Case 106: KeyCodeToString = "*"
    Case 111: KeyCodeToString = "/"
    Case Else: KeyCodeToString = "Key" & vkCode
    End Select
End Function
Public Function Keys(Optional ByVal SeparationString As String = ";") As String
    Dim i As Integer
    For i = 1 To 249 Step 1
        If GetKeyState(i) < 0 Then If Keys <> "" Then Keys = Keys & SeparationString & CStr(i) Else Keys = CStr(i)
    Next i
End Function
Public Function KeyState(ByVal vkCode As Long) As Boolean: KeyState = (GetKeyState(vkCode) < 0): End Function
Public Function StringToKeyCode(ByVal Character As String) As Long
    Select Case LCase(Character)
    Case "a": StringToKeyCode = vbKeyA
    Case "b": StringToKeyCode = vbKeyB
    Case "c": StringToKeyCode = vbKeyC
    Case "d": StringToKeyCode = vbKeyD
    Case "e": StringToKeyCode = vbKeyE
    Case "f": StringToKeyCode = vbKeyF
    Case "g": StringToKeyCode = vbKeyG
    Case "h": StringToKeyCode = vbKeyH
    Case "i": StringToKeyCode = vbKeyI
    Case "j": StringToKeyCode = vbKeyJ
    Case "k": StringToKeyCode = vbKeyK
    Case "l": StringToKeyCode = vbKeyL
    Case "m": StringToKeyCode = vbKeyM
    Case "n": StringToKeyCode = vbKeyN
    Case "o": StringToKeyCode = vbKeyO
    Case "p": StringToKeyCode = vbKeyP
    Case "q": StringToKeyCode = vbKeyQ
    Case "r": StringToKeyCode = vbKeyR
    Case "s": StringToKeyCode = vbKeyS
    Case "t": StringToKeyCode = vbKeyT
    Case "u": StringToKeyCode = vbKeyU
    Case "v": StringToKeyCode = vbKeyV
    Case "w": StringToKeyCode = vbKeyW
    Case "x": StringToKeyCode = vbKeyX
    Case "y": StringToKeyCode = vbKeyY
    Case "z": StringToKeyCode = vbKeyZ
    Case "shift": StringToKeyCode = vbKeyShift
    Case "space": StringToKeyCode = vbKeySpace
    Case "tab": StringToKeyCode = vbKeyTab
    Case "return": StringToKeyCode = vbKeyReturn
    Case "up": StringToKeyCode = vbKeyUp
    Case "down": StringToKeyCode = vbKeyDown
    Case "left": StringToKeyCode = vbKeyLeft
    Case "right": StringToKeyCode = vbKeyRight
    Case "0": StringToKeyCode = vbKey0
    Case "1": StringToKeyCode = vbKey1
    Case "2": StringToKeyCode = vbKey2
    Case "3": StringToKeyCode = vbKey3
    Case "4": StringToKeyCode = vbKey4
    Case "5": StringToKeyCode = vbKey5
    Case "6": StringToKeyCode = vbKey6
    Case "7": StringToKeyCode = vbKey7
    Case "8": StringToKeyCode = vbKey8
    Case "9": StringToKeyCode = vbKey9
    Case "numpad0": StringToKeyCode = vbKeyNumpad0
    Case "numpad1": StringToKeyCode = vbKeyNumpad1
    Case "numpad2": StringToKeyCode = vbKeyNumpad2
    Case "numpad3": StringToKeyCode = vbKeyNumpad3
    Case "numpad4": StringToKeyCode = vbKeyNumpad4
    Case "numpad5": StringToKeyCode = vbKeyNumpad5
    Case "numpad6": StringToKeyCode = vbKeyNumpad6
    Case "numpad7": StringToKeyCode = vbKeyNumpad7
    Case "numpad8": StringToKeyCode = vbKeyNumpad8
    Case "numpad9": StringToKeyCode = vbKeyNumpad9
    Case "escape": StringToKeyCode = vbKeyEscape
    Case "del": StringToKeyCode = vbKeyDelete
    Case "backspace": StringToKeyCode = vbKeyBack
    Case "control": StringToKeyCode = vbKeyControl
    Case "capital": StringToKeyCode = vbKeyCapital
    Case "maj": StringToKeyCode = 160
    Case "home": StringToKeyCode = vbKeyHome
    Case "end": StringToKeyCode = vbKeyEnd
    Case "alt": StringToKeyCode = 164
    Case "alt gr": StringToKeyCode = 165
    Case "windows": StringToKeyCode = 91
    Case "insert": StringToKeyCode = vbKeyInsert
    Case "page down": StringToKeyCode = vbKeyPageDown
    Case "page up": StringToKeyCode = vbKeyPageUp
    'Case "scroll lock": StringToKeyCode = vbKeyScrollLock
    Case "pause": StringToKeyCode = vbKeyPause
    Case "snapshot": StringToKeyCode = vbKeySnapshot
    Case "f1": StringToKeyCode = vbKeyF1
    Case "f2": StringToKeyCode = vbKeyF2
    Case "f3": StringToKeyCode = vbKeyF3
    Case "f4": StringToKeyCode = vbKeyF4
    Case "f5": StringToKeyCode = vbKeyF5
    Case "f6": StringToKeyCode = vbKeyF6
    Case "f7": StringToKeyCode = vbKeyF7
    Case "f8": StringToKeyCode = vbKeyF8
    Case "f9": StringToKeyCode = vbKeyF9
    Case "f10": StringToKeyCode = vbKeyF10
    Case "f11": StringToKeyCode = vbKeyF11
    Case "f12": StringToKeyCode = vbKeyF12
    Case "num lock": StringToKeyCode = vbKeyNumlock
    Case "?": StringToKeyCode = 191
    Case "^": StringToKeyCode = 221
    Case "'": StringToKeyCode = 219
    Case "menu": StringToKeyCode = 93
    Case "$": StringToKeyCode = 223
    Case "?": StringToKeyCode = 220
    Case "?": StringToKeyCode = 222
    Case "?": StringToKeyCode = 186
    Case "?": StringToKeyCode = 192
    Case ".": StringToKeyCode = 110
    Case ",": StringToKeyCode = 188
    Case "-": StringToKeyCode = 109
    Case "+": StringToKeyCode = 107
    Case "*": StringToKeyCode = 106
    Case "/": StringToKeyCode = 111
    Case Else: StringToKeyCode = Asc(Character)
    End Select
End Function

