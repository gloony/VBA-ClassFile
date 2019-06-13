Option Explicit
'Ceci n'est pas utilisé dans le code, mais c'est fun de montrer comment faire du sub classing sur vba :P
Private Const WH_JOURNALRECORD = 0
Private Const WH_JOURNALPLAYBACK = 1
Private Const WH_KEYBOARD = 2
Private Const WH_GETMESSAGE = 3
Private Const WH_CALLWNDPROC = 4
Private Const WH_CBT = 5
Private Const WH_SYSMSGFILTER = 6
Private Const WH_MOUSE = 7
Private Const WH_HARDWARE = 8
Private Const WH_DEBUG = 9
Private Const WH_SHELL = 10
Private Const WH_FOREGROUNDIDLE = 11
Private Const WH_CALLWNDPROCRET = 12
Private Const WH_KEYBOARD_LL = 13
Private Const WH_MOUSE_LL = 14

'Keyboard Constants, Types and Functions
Private Const vbKeyAlt = 18
Private Const vbKeyWindows = 91
Private Const WM_KEYDOWN = &H100
Private Const WM_KEYUP = &H101
Private Const WM_SYSKEYDOWN = &H104
Private Const WM_SYSKEYUP = &H105
Private Type KBDLLHOOKSTRUCT
    vkCode As Long
    scanCode As Long
    Flags As Long
    time As Long
    dwExtraInfo As Long
End Type
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

'Mouse Constants and Types
Private Type POINT
    X As Long
    Y As Long
End Type
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_MOUSEMOVE = &H200
Private Const WM_MOUSEWHEEL = &H20A
Private Const WM_MOUSEHWHEEL = &H20E
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Type MSLLHOOKSTRUCT
    pt As POINT
    scanCode As Long
    Flags As Long
    time As Long
    dwExtraInfo As Long
End Type

Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Function BoolGetKeyState(ByVal nVirtKey As Long) As Boolean
    BoolGetKeyState = (GetKeyState(nVirtKey) < 0)
End Function

Public Sub HookCreateKeyboardLL()
    If shHookKeyboard.Range("A1").Value <> "" Then HookKillKeyboardLL
    shHookKeyboard.Range("A1").Value = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf HookProcKeyboardLL, Application.hInstance, 0)
    If shHookKeyboard.Range("A1").Value = 0 Then shHookKeyboard.Range("A1").Value = ""
End Sub
Public Sub HookKillKeyboardLL()
    If shHookKeyboard.Range("A1").Value = "" Then Exit Sub
    UnhookWindowsHookEx shHookKeyboard.Range("A1").Value
    shHookKeyboard.Range("A1").Value = ""
End Sub
Private Function HookProcKeyboardLL(ByVal uCode As Long, ByVal wParam As Long, lParam As KBDLLHOOKSTRUCT) As Long
    If uCode >= 0 Then
        CallNextHookEx WH_KEYBOARD_LL, uCode, wParam, lParam.scanCode
        Dim Modifiers As Integer: Modifiers = 0
        Dim alertTime
        Select Case wParam
        Case WM_KEYDOWN, WM_SYSKEYDOWN
            If BoolGetKeyState(vbKeyAlt) Then Modifiers = Modifiers + 1
            If BoolGetKeyState(vbKeyControl) Then Modifiers = Modifiers + 2
            If BoolGetKeyState(vbKeyShift) Then Modifiers = Modifiers + 4
            If BoolGetKeyState(vbKeyWindows) Then Modifiers = Modifiers + 8
            If shHookKeyboard.Cells(lParam.vkCode, Modifiers + 2).Value <> "" And Not shHookKeyboard.Cells(lParam.vkCode, Modifiers + 2).Font.Bold Then
                If shHookKeyboard.Cells(lParam.vkCode, Modifiers + 2).Value <> "void" Then
                    shHookKeyboard.Cells(lParam.vkCode, Modifiers + 2).Interior.Color = vbYellow
                    alertTime = Now + TimeValue("00:00:00")
                    Application.OnTime alertTime, "'HookCallBackKeyboard """ & Modifiers & """, """ & lParam.vkCode & """'"
                End If
                If shHookKeyboard.Cells(lParam.vkCode, Modifiers + 2).Font.Color <> vbRed Then HookProcKeyboardLL = 1
            End If
        Case WM_KEYUP, WM_SYSKEYUP
            If BoolGetKeyState(vbKeyAlt) Then Modifiers = Modifiers + 1
            If BoolGetKeyState(vbKeyControl) Then Modifiers = Modifiers + 2
            If BoolGetKeyState(vbKeyShift) Then Modifiers = Modifiers + 4
            If BoolGetKeyState(vbKeyWindows) Then Modifiers = Modifiers + 8
            If shHookKeyboard.Cells(lParam.vkCode, Modifiers + 2).Value <> "" And shHookKeyboard.Cells(lParam.vkCode, Modifiers + 2).Font.Bold Then
                If shHookKeyboard.Cells(lParam.vkCode, Modifiers + 2).Value <> "void" Then
                    shHookKeyboard.Cells(lParam.vkCode, Modifiers + 2).Interior.Color = vbYellow
                    alertTime = Now + TimeValue("00:00:00")
                    Application.OnTime alertTime, "'HookCallBackKeyboard """ & Modifiers & """, """ & lParam.vkCode & """'"
                End If
                If shHookKeyboard.Cells(lParam.vkCode, Modifiers + 2).Font.Color <> vbRed Then HookProcKeyboardLL = 1
            End If
        Case Else
            HookProcKeyboardLL = 0
        End Select
    Else
        HookProcKeyboardLL = CallNextHookEx(WH_KEYBOARD_LL, uCode, wParam, lParam.scanCode)
    End If
End Function
Public Sub HookCallBackKeyboard(ByVal Modifiers As Long, ByVal xCode As Long)
    shHookKeyboard.Cells(xCode, Modifiers + 2).Interior.Color = vbGreen
    Init
    Dim ScriptString As String: ScriptString = shHookKeyboard.Cells(xCode, Modifiers + 2).Value
    If Right(ScriptString, "4") = ".vbs" Or Right(ScriptString, "3") = ".js" Then
        Dim tPath As String
        tPath = ScriptString
        If Not File.Exist(ScriptString) Then ScriptString = ThisWorkbook.Path & "\scripts\hotkey\" & tPath
        If Not File.Exist(ScriptString) Then ScriptString = ThisWorkbook.Path & "\scripts\" & tPath
        If File.Exist(ScriptString) Then Script.Execute ScriptString
    Else
        Script.AddCode ScriptString
    End If
    shHookKeyboard.Cells(xCode, Modifiers + 2).Interior.Color = xlNone
End Sub

Sub ladkjsas()
    HookCallBackKeyboard 8, 67
End Sub

Public Sub HookCreateMouseLL()
    If shGUI.Range("C1").Value <> "" Then HookKillMouseLL
    shGUI.Range("C1").Value = SetWindowsHookEx(WH_MOUSE_LL, AddressOf HookProcMouseLL, Application.hInstance, 0)
    If shGUI.Range("C1").Value = 0 Then shGUI.Range("C1").Value = ""
End Sub
Public Sub HookKillMouseLL()
    If shGUI.Range("C1").Value = "" Then Exit Sub
    UnhookWindowsHookEx shGUI.Range("C1").Value
    shGUI.Range("C1").Value = ""
End Sub
Private Function HookProcMouseLL(ByVal uCode As Long, ByVal wParam As Long, lParam As MSLLHOOKSTRUCT) As Long
    If uCode >= 0 Then
        CallNextHookEx WH_MOUSE_LL, uCode, wParam, lParam.scanCode
        Select Case wParam
        Case WM_RBUTTONUP
            If BoolGetKeyState(vbKeyWindows) Then
                Dim alertTime: alertTime = Now + TimeValue("00:00:01")
                Application.OnTime alertTime, "'HookCallBackMouse """ & wParam & """'"
                HookProcMouseLL = 1
            End If
        Case Else
            HookProcMouseLL = 0
        End Select
    Else
        HookProcMouseLL = CallNextHookEx(WH_MOUSE_LL, uCode, wParam, lParam.scanCode)
    End If
End Function
