Option Explicit

Public OldWindowProc As Long
Public ActualID As Long

Public Enum fsModifiers
    Alt = &H1
    Control = &H2
    Shift = &H4
    Win = &H8 'Reserved for system
    NoRepeat = &H4000
End Enum

Private Const HTFS_ALT = &H1
Private Const HTFS_CONTROL = &H2
Private Const HTFS_SHIFT = &H4
Private Const HTFS_WIN = &H8 'Reserved for system
Private Const HTFS_NOREPEAT = &H4000

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function RegisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal id As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
Public Declare Function UnregisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal id As Long) As Long

Public Sub ToggleHotKey()
    If OldWindowProc = 0 Then
        ActivateHotKey
    Else
        RemoveHotKey
    End If
End Sub

Public Sub ActivateHotKey()
    If OldWindowProc <> 0 Then RemoveHotKey
    Init
    Dim i As Long: i = 4
    While shGUI.Range("B" & i).Value <> ""
        Dim Modifiers As String: Modifiers = shGUI.Range("B" & i).Value
        Select Case UCase(Modifiers)
            Case "ALT": Modifiers = HTFS_ALT
            Case "CTRL": Modifiers = HTFS_CONTROL
            Case "SHIFT": Modifiers = HTFS_SHIFT
            Case "ALT + CTRL": Modifiers = HTFS_ALT Or HTFS_CONTROL
            Case "CTRL + ALT": Modifiers = HTFS_ALT Or HTFS_CONTROL
            Case "ALT + SHIFT": Modifiers = HTFS_ALT Or HTFS_SHIFT
            Case "SHIFT + ALT": Modifiers = HTFS_ALT Or HTFS_SHIFT
            Case "CTRL + SHIFT": Modifiers = HTFS_CONTROL Or HTFS_SHIFT
            Case "SHIFT + CTRL": Modifiers = HTFS_CONTROL Or HTFS_SHIFT
            Case "ALT + CTRL + SHIFT": Modifiers = HTFS_ALT Or HTFS_CONTROL Or HTFS_SHIFT
            Case "CTRL + ALT + SHIFT": Modifiers = HTFS_ALT Or HTFS_CONTROL Or HTFS_SHIFT
            Case "SHIFT + CTRL + ALT": Modifiers = HTFS_ALT Or HTFS_CONTROL Or HTFS_SHIFT
            Case "SHIFT + ALT + CTRL": Modifiers = HTFS_ALT Or HTFS_CONTROL Or HTFS_SHIFT
            Case "NR + ALT": Modifiers = HTFS_ALT Or HTFS_NOREPEAT
            Case "NR + CTRL": Modifiers = HTFS_CONTROL Or HTFS_NOREPEAT
            Case "NR + SHIFT": Modifiers = HTFS_SHIFT Or HTFS_NOREPEAT
            Case "NR + ALT + CTRL": Modifiers = HTFS_ALT Or HTFS_CONTROL Or HTFS_NOREPEAT
            Case "NR + CTRL + ALT": Modifiers = HTFS_ALT Or HTFS_CONTROL Or HTFS_NOREPEAT
            Case "NR + ALT + SHIFT": Modifiers = HTFS_ALT Or HTFS_SHIFT Or HTFS_NOREPEAT
            Case "NR + SHIFT + ALT": Modifiers = HTFS_ALT Or HTFS_SHIFT Or HTFS_NOREPEAT
            Case "NR + CTRL + SHIFT": Modifiers = HTFS_CONTROL Or HTFS_SHIFT Or HTFS_NOREPEAT
            Case "NR + SHIFT + CTRL": Modifiers = HTFS_CONTROL Or HTFS_SHIFT Or HTFS_NOREPEAT
            Case "NR + ALT + CTRL + SHIFT": Modifiers = HTFS_ALT Or HTFS_CONTROL Or HTFS_SHIFT Or HTFS_NOREPEAT
            Case "NR + CTRL + ALT + SHIFT": Modifiers = HTFS_ALT Or HTFS_CONTROL Or HTFS_SHIFT Or HTFS_NOREPEAT
            Case "NR + SHIFT + CTRL + ALT": Modifiers = HTFS_ALT Or HTFS_CONTROL Or HTFS_SHIFT Or HTFS_NOREPEAT
            Case "NR + SHIFT + ALT + CTRL": Modifiers = HTFS_ALT Or HTFS_CONTROL Or HTFS_SHIFT Or HTFS_NOREPEAT
        End Select
        Dim KeyCode As String: KeyCode = shGUI.Range("C" & i).Value
        If Not IsNumeric(KeyCode) Then KeyCode = Keyboard.StringToKeyCode(KeyCode)
        RegisterHotKey Application.hwnd, (i - 3), CLng(Modifiers), KeyCode
        i = i + 1
    Wend
    OldWindowProc = SetWindowLong(Application.hwnd, (-4), AddressOf NewWindowProc)
    shGUI.Range("B2:D2").Interior.Color = vbGreen
End Sub

Public Sub RemoveHotKey()
    If OldWindowProc = 0 Then Exit Sub
    Dim i As Long: i = 4
    While shGUI.Range("B" & i).Value <> ""
        UnregisterHotKey Application.hwnd, (i - 3)
        i = i + 1
    Wend
    UnregisterHotKey Application.hwnd, 3
    SetWindowLong Application.hwnd, (-4), OldWindowProc
    OldWindowProc = 0
    shGUI.Range("B2:D2").Interior.Color = vbWhite
End Sub

Public Sub CreateHotKey(ByVal PrimaryKey As Long, ByVal ForeignKey As Long, Optional Index As Long = -1)
    RegisterHotKey Application.hwnd, Index, PrimaryKey, ForeignKey
End Sub

Public Function NewWindowProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Const WM_NCDESTROY = &H82
    Const WM_HOTKEY = &H312
    If Msg = WM_NCDESTROY Then RemoveHotKey
    If Msg = WM_HOTKEY Then Hotkey wParam
    NewWindowProc = CallWindowProc(OldWindowProc, hwnd, Msg, wParam, lParam)
End Function

Private Sub Hotkey(ByVal HookID As Long)
    shGUI.Range("D" & (HookID + 3)).Interior.Color = vbYellow
    Init
    Dim ScriptString As String: ScriptString = shGUI.Range("D" & (HookID + 3)).Value
    If Right(ScriptString, "4") = ".vbs" Or Right(ScriptString, "3") = ".js" Then
        Dim tPath As String
        tPath = ScriptString
        If Not File.Exist(ScriptString) Then ScriptString = ThisWorkbook.Path & "\scripts\hotkey\" & tPath
        If Not File.Exist(ScriptString) Then ScriptString = ThisWorkbook.Path & "\scripts\" & tPath
        If File.Exist(ScriptString) Then Script.Execute ScriptString
    Else
        Script.AddCode ScriptString
    End If
    shGUI.Range("D" & (HookID + 3)).Interior.Color = vbWhite
End Sub
