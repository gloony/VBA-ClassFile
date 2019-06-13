Option Explicit
Private Declare Function joyGetNumDevs Lib "winmm.dll" () As Long
Private Declare Function joyGetPosEx Lib "winmm.dll" (ByVal uJoyID As Long, pji As Joypad) As Long

Private Type Joypad
    dwSize As Long
    dwFlags As Long
    dwXpos As Long
    dwYpos As Long
    dwZpos As Long
    dwRpos As Long
    dwUpos As Long
    dwVpos As Long
    dwButtons As Long
    dwButtonNumber As Long
    dwPOV As Long
    dwReserved1 As Long
    dwReserved2 As Long
End Type

Private Function GetTouch(Optional nPad As Integer = 0) As Joypad
    On Error Resume Next
    GetTouch.dwSize = 64
    GetTouch.dwFlags = &HFFF
    joyGetPosEx 0, GetTouch
    On Error GoTo 0
End Function

Public Function Button(nButton As Integer, Optional nPad As Integer = 0) As Boolean
    On Error Resume Next
    If (GetTouch(nPad).dwButtons And (2 ^ nButton)) > 0 Then Button = True
    If Err.Number <> 0 Then Button = False
    On Error GoTo 0
End Function
Public Function Count() As Long
    On Error Resume Next
    Count = joyGetNumDevs
    On Error GoTo 0
End Function
Public Function Stick(nButton As String, Optional nPad As Integer = 0) As Long
    Dim nJoy As Joypad: nJoy = GetTouch(nPad)
    Select Case UCase(nButton)
    Case "POV": Stick = nJoy.dwPOV
    Case "R": Stick = nJoy.dwRpos
    Case "U": Stick = nJoy.dwUpos
    Case "V": Stick = nJoy.dwVpos
    Case "X": Stick = nJoy.dwXpos
    Case "Y": Stick = nJoy.dwYpos
    Case "Z": Stick = nJoy.dwZpos
    End Select
End Function
