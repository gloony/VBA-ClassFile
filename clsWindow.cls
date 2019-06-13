Option Explicit
Private Declare Function AttachThreadInput Lib "user32.dll" (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As XY) As Long
Private Declare Function GetCurrentThreadId Lib "kernel32.dll" () As Long
Private Declare Function GetWindowThreadProcessId Lib "user32.dll" (ByVal hwnd As Long, ByRef lpdwProcessId As Long) As Long
Private Declare Function MapVirtualKey Lib "user32.dll" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long
Private Declare Function GetWindowModuleFileName Lib "user32.dll" Alias "GetWindowModuleFileNameA" (ByVal hwnd As Long, ByVal pszFileName As String, ByVal cchFileNameMax As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As XY) As Long
Private Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function ChildWindowFromPointEx Lib "user32" (ByVal Parent As Long, ByVal X As Long, ByVal Y As Long, ByVal Flags As Long) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function MoveWindowA Lib "user32" Alias "MoveWindow" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function FindWindowExA Lib "user32" (ByVal hWndParent As Long, ByVal hWndChildAfter As Long, ByVal lpszClassName As String, ByVal lpszWindowName As String) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function IsWindowA Lib "user32.dll" Alias "IsWindow" (ByVal hwnd As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function IsZoomed Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetKeyboardState Lib "user32.dll" (ByRef pbKeyState As Byte) As Long
Private Declare Function SetKeyboardState Lib "user32.dll" (ByRef lppbKeyState As Byte) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function PostMessageA Lib "user32" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessageA Lib "user32" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageS Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function SetParentS Lib "user32.dll" Alias "SetParent" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function OpenProcess Lib "Kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessID As Long) As Long
Private Declare Function CloseHandle Lib "Kernel32" (ByVal hObject As Long) As Long
Private Declare Function TerminateProcess Lib "Kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long

Private Type XY
    X As Long
    Y As Long
End Type
Private Type RECT
    Left            As Long
    Top             As Long
    Right           As Long
    Bottom          As Long
End Type

Private fHWnd As Long
Private wDir  As String

Private Function HWndOrCurrent(ByVal hwnd As Long) As Long
    If hwnd = -1 Then
        If fHWnd <> -1 Then HWndOrCurrent = fHWnd _
        Else HWndOrCurrent = Me.Foreground
    Else
        HWndOrCurrent = hwnd
    End If
End Function

Property Get All(ByVal Index As Long)
    All = UserWindowVar(Index).hwnd
End Property
Public Property Get Alls() As Long
    If IsWindowVarInitialized Then Alls = UBound(UserWindowVar) Else Alls = -1
End Property
Property Get Child(ByVal Index As Long, Optional ByVal ByHWnd As Long = -1)
    ByHWnd = HWndOrCurrent(ByHWnd)
    Child = 0
    Dim wh   As Long
    Dim LvlH As Long
    LvlH = 0
    wh = GetWindow(ByHWnd, 5)
    Do While wh <> 0
        If LvlH = Index Then
            Child = wh
            Exit Do
        End If
        wh = GetWindow(wh, 2)
        LvlH = LvlH + 1
    Loop
End Property
Public Property Get Childs(Optional ByVal ByHWnd As Long = -1) As Long
    ByHWnd = HWndOrCurrent(ByHWnd)
    Childs = 0
    Dim wh As Long
    wh = GetWindow(ByHWnd, 5)
    Do While wh <> 0
        wh = GetWindow(wh, 2)
        Childs = Childs + 1
    Loop
End Property
Public Property Get CurHWnd() As Long: CurHWnd = fHWnd: End Property
Public Property Let CurHWnd(ByVal ByHWnd As Long): fHWnd = ByHWnd: End Property
Public Sub EnableAll(Optional ByVal ByHWnd As Long = -1)
    ByHWnd = HWndOrCurrent(ByHWnd)
    Dim wh As Long: wh = GetWindow(ByHWnd, 5)
    Do While wh <> 0
        If Not IsWindowEnabled(wh) Then EnableWindow wh, True
        Me.EnableAll wh
        wh = GetWindow(wh, 2)
    Loop
End Sub
Public Property Get Enable(Optional ByVal ByHWnd As Long = -1) As Boolean: ByHWnd = HWndOrCurrent(ByHWnd): Enable = (IsWindowEnabled(ByHWnd) = 1): End Property
Public Property Let Enable(Optional ByVal ByHWnd As Long = -1, ByVal Enabled As Boolean)
    ByHWnd = HWndOrCurrent(ByHWnd)
    Dim B As Integer
    If Enabled Then B = 1 Else B = 0
    EnableWindow ByHWnd, B
End Property
Public Property Get Foreground() As Long
    If CurrentHwnd <> 0 Then Foreground = CurrentHwnd _
    Else Foreground = GetForegroundWindow
End Property
Public Property Let Foreground(ByVal vData As Long): Call Focus(vData): End Property
Public Property Get Height(Optional ByVal ByHWnd As Long = -1) As Long
    ByHWnd = HWndOrCurrent(ByHWnd)
    Dim REC As RECT
    GetWindowRect ByHWnd, REC
    Height = REC.Bottom - REC.Top
End Property
Public Property Let Height(Optional ByVal ByHWnd As Long = -1, ByVal vData As Long)
    ByHWnd = HWndOrCurrent(ByHWnd)
    Dim intTopMost As Integer
    Dim REC        As RECT
    GetWindowRect ByHWnd, REC
    If OnTop(ByHWnd) Then intTopMost = -1 Else intTopMost = -2
    SetWindowPos ByHWnd, intTopMost, REC.Left, REC.Top, REC.Right - REC.Left, vData, &H40
End Property
Public Property Get Hide(Optional ByVal ByHWnd As Long = -1) As Boolean: ByHWnd = HWndOrCurrent(ByHWnd): Hide = UserWindowVar(GetWindowVar(ByHWnd).id).Hide: End Property
Public Property Let Hide(Optional ByVal ByHWnd As Long = -1, ByVal Enable As Boolean)
    ByHWnd = HWndOrCurrent(ByHWnd)
    UserWindowVar(GetWindowVar(ByHWnd).id).Hide = Enable
    If Enable Then ShowWindow ByHWnd, 0 _
   Else ShowWindow ByHWnd, 5
End Property
Public Property Get Left(Optional ByVal ByHWnd As Long = -1) As Long
    ByHWnd = HWndOrCurrent(ByHWnd)
    Dim REC As RECT: GetWindowRect ByHWnd, REC
    Left = REC.Left
End Property
Public Property Let Left(Optional ByVal ByHWnd As Long = -1, ByVal vData As Long)
    ByHWnd = HWndOrCurrent(ByHWnd)
    Dim intTopMost As Integer
    Dim REC        As RECT
    GetWindowRect ByHWnd, REC
    If OnTop(ByHWnd) Then intTopMost = -1 Else intTopMost = -2
    SetWindowPos ByHWnd, intTopMost, vData, REC.Top, REC.Right - REC.Left, REC.Bottom - REC.Top, &H40
End Property
Public Property Get Maximize(Optional ByVal ByHWnd As Long = -1) As Boolean: ByHWnd = HWndOrCurrent(ByHWnd): Maximize = IsZoomed(ByHWnd): End Property
Public Property Let Maximize(Optional ByVal ByHWnd As Long = -1, ByVal Enable As Boolean)
    ByHWnd = HWndOrCurrent(ByHWnd)
    If Enable Then ShowWindow ByHWnd, 3 Else ShowWindow ByHWnd, 5
End Property
Public Property Get Minimize(Optional ByVal ByHWnd As Long = -1) As Boolean: ByHWnd = HWndOrCurrent(ByHWnd): Minimize = IsIconic(ByHWnd): End Property
Public Property Let Minimize(Optional ByVal ByHWnd As Long = -1, ByVal Enable As Boolean)
    ByHWnd = HWndOrCurrent(ByHWnd)
    If Enable Then ShowWindow ByHWnd, 6 Else ShowWindow ByHWnd, 5
End Property
Public Property Get OnTop(Optional ByVal ByHWnd As Long = -1) As Boolean: ByHWnd = HWndOrCurrent(ByHWnd): OnTop = (GetWindowLong(ByHWnd, (-20)) And &H8&) <> 0: End Property
Public Property Let OnTop(Optional ByVal ByHWnd As Long = -1, ByVal Enable As Boolean)
    ByHWnd = HWndOrCurrent(ByHWnd)
    If Enable Then SetWindowPos ByHWnd, -1, 0, 0, 0, 0, &H2 Or &H1 Else SetWindowPos ByHWnd, -2, 0, 0, 0, 0, &H10 Or &H40 Or &H2 Or &H1
End Property
Property Get Parent(Optional ByVal Index As Long = -1, Optional ByVal ByHWnd As Long = -1)
    ByHWnd = HWndOrCurrent(ByHWnd)
    If Index = -1 Then
        Parent = Father(ByHWnd)
        Exit Property
    Else
        Parent = 0
    End If
    Dim tnHook As Long, oHook As Long
    Dim LvlH   As Long
    LvlH = 0
    oHook = Parent
    Do While oHook <> 0
        tnHook = oHook
        oHook = GetParent(oHook)
        If oHook <> 0 Then
            If LvlH = Index Then
                Parent = tnHook
                Exit Do
            End If
            LvlH = LvlH + 1
        End If
    Loop
    If Index = -1 Then Parent = tnHook
End Property
Public Property Get Parents(Optional ByVal ByHWnd As Long = -1) As Long
    Parents = HWndOrCurrent(ByHWnd)
    Dim tnHook As Long
    tnHook = ByHWnd
    While tnHook <> 0
        tnHook = GetParent(tnHook)
        If tnHook <> 0 Then Parents = Parents + 1
    Wend
End Property
Public Property Get ScriptDir() As String: ScriptDir = wDir: End Property
Public Property Let ScriptDir(ByVal PathString As String): wDir = PathString: End Property
Public Property Get Style(ByVal StyleConst As Long, Optional ByVal ByHWnd As Long = -1) As Boolean: ByHWnd = HWndOrCurrent(ByHWnd): Style = ((ByHWnd And StyleConst) = StyleConst): End Property
Public Property Let Style(ByVal StyleConst As Long, Optional ByVal ByHWnd As Long = -1, ByVal Enable As Boolean)
    ByHWnd = HWndOrCurrent(ByHWnd)
    Dim dwCurStyle As Long
    Dim fAdd As Boolean: fAdd = True
    Dim fRedraw As Boolean: fRedraw = True
    dwCurStyle = GetWindowLong(ByHWnd, -16)
    If Err.LastDllError = 0 Then
        If fAdd And (dwCurStyle And StyleConst) = 0 Then
            dwCurStyle = dwCurStyle Or StyleConst
        ElseIf (Not fAdd) And (dwCurStyle And StyleConst) Then
            dwCurStyle = dwCurStyle And (Not StyleConst)
        End If
        Call SetWindowLong(ByHWnd, -16, dwCurStyle)
        If fRedraw Then Call SetWindowPos(ByHWnd, 0, 0, 0, 0, 0, &H4 Or &H2 Or &H1 Or &H20)
    End If
End Property
Public Property Get Top(Optional ByVal ByHWnd As Long = -1) As Long
    ByHWnd = HWndOrCurrent(ByHWnd)
    Dim REC As RECT
    GetWindowRect ByHWnd, REC
    Top = REC.Top
End Property
Public Property Let Top(Optional ByVal ByHWnd As Long = -1, ByVal vData As Long)
    ByHWnd = HWndOrCurrent(ByHWnd)
    Dim intTopMost As Integer
    Dim REC        As RECT
    GetWindowRect ByHWnd, REC
    If Window.OnTop(ByHWnd) Then intTopMost = -1 Else intTopMost = -2
    SetWindowPos ByHWnd, intTopMost, REC.Left, vData, REC.Right - REC.Left, REC.Bottom - REC.Top, &H40
End Property
Public Property Get Width(Optional ByVal ByHWnd As Long = -1) As Long
    ByHWnd = HWndOrCurrent(ByHWnd)
    Dim REC As RECT
    GetWindowRect ByHWnd, REC
    Width = REC.Right - REC.Left
End Property
Public Property Let Width(Optional ByVal ByHWnd As Long = -1, ByVal vData As Long)
    ByHWnd = HWndOrCurrent(ByHWnd)
    Dim intTopMost As Integer
    Dim REC        As RECT
    GetWindowRect ByHWnd, REC
    If Window.OnTop(ByHWnd) Then intTopMost = -1 Else intTopMost = -2
    SetWindowPos ByHWnd, intTopMost, REC.Left, REC.Top, vData, REC.Bottom - REC.Top, &H40
End Property

Public Sub BlendColor(ByVal ColorString As String, Optional ByVal Opacity As Integer = 100, Optional ByVal ByHWnd As Long = -1)
    ByHWnd = HWndOrCurrent(ByHWnd)
    Dim ExStyle As Long, tTemp   As Long
    If Opacity >= 100 Then
        ExStyle = GetWindowLong(ByHWnd, -20)
        If ExStyle <> (GetWindowLong(ByHWnd, -20) Or &H80000) Then
            ExStyle = (GetWindowLong(ByHWnd, -20) Or &H80000)
            Call SetWindowLong(ByHWnd, -20, ExStyle)
        End If
        tTemp = (SetLayeredWindowAttributes(ByHWnd, ColorString, 255, IIf(False, &H2, &H1)) <> 0)
   Else
        ExStyle = GetWindowLong(ByHWnd, -20)
        If ExStyle <> (GetWindowLong(ByHWnd, -20) Or &H80000) Then
            ExStyle = (GetWindowLong(ByHWnd, -20) Or &H80000)
            Call SetWindowLong(ByHWnd, -20, ExStyle)
        End If
        tTemp = (SetLayeredWindowAttributes(ByHWnd, ColorString, Opacity, &H3) <> 0)
    End If
End Sub
Public Function ChildWindowFromPoint(X As Long, Y As Long, Optional ByHWnd As Long = -1) As Long
    Dim hwnd As Long: hwnd = HWndOrCurrent(ByHWnd)
    X = X - Window.Left(hwnd)
    Y = Y - Window.Top(hwnd)
    Dim lHWnd As Long
    lHWnd = ChildWindowFromPointEx(hwnd, X, Y, 0)
    If lHWnd = 0 Then
        ChildWindowFromPoint = hwnd
        Exit Function
    ElseIf lHWnd <> hwnd Then
        Dim lPoint As XY
        lPoint.X = X
        lPoint.Y = Y
        Call ClientToScreen(hwnd, lPoint)
        Call ScreenToClient(lHWnd, lPoint)
        Dim lHWndRecurse As Long
        lHWndRecurse = ChildWindowFromPoint(X, Y, hwnd)
        If lHWndRecurse <> 0 Then lHWnd = lHWndRecurse
    End If
    ChildWindowFromPoint = lHWnd
End Function
Public Function Class(Optional ByVal ByHWnd As Long = -1) As String
    ByHWnd = HWndOrCurrent(ByHWnd)
    Const MAX_LEN As Byte = 255
    Dim strBuff   As String, intLen As Integer
    strBuff = String(MAX_LEN, vbNullChar)
    intLen = GetClassName(ByHWnd, strBuff, MAX_LEN)
    Class = Strings.Left(strBuff, intLen)
End Function
Public Sub Destroy(Optional ByVal ByHWnd As Long = -1): ByHWnd = HWndOrCurrent(ByHWnd): Call DestroyWindow(ByHWnd): End Sub
Public Function Father(Optional ByVal ByHWnd As Long = -1) As Long
    ByHWnd = HWndOrCurrent(ByHWnd)
    Dim tnHook As Long, oHook As Long
    tnHook = ByHWnd
    oHook = ByHWnd
    While oHook <> 0
        tnHook = oHook
        oHook = GetParent(oHook)
    Wend
    Father = tnHook
End Function
Public Function FindWindow(Optional ByVal ClassString As String = "", Optional ByVal NameString As String = "", Optional Method As String = "STRICT") As Long
    If ClassString = "" Then ClassString = vbNullString
    FindWindow = FindWindowA(ClassString, NameString)
    If FindWindow <> 0 Then Exit Function
    Select Case UCase(Method)
    Case "LIKE"
        Dim tHWnd As Long
        tHWnd = HwndFirstLoaded
        Do
            If ClassString = "" And NameString = "" Then
                Exit Function
            ElseIf ClassString = "" Then
                If Name(tHWnd) Like NameString Then
                    FindWindow = tHWnd
                    Exit Function
                End If
            ElseIf NameString = "" Then
                If Class(tHWnd) Like ClassString Then
                    FindWindow = tHWnd
                    Exit Function
                End If
            ElseIf Name(tHWnd) Like NameString And Class(tHWnd) = ClassString Then
                FindWindow = tHWnd
                Exit Function
            End If
            tHWnd = GetWindow(tHWnd, 2)
        Loop While tHWnd <> 0
    End Select
End Function
Public Sub KillProcess(Optional ByVal ByHWnd As Long = -1)
    ByHWnd = HWndOrCurrent(ByHWnd)
    Dim IDProc: IDProc = Me.IDProc(ByHWnd)
    Dim hProcess As Long
    hProcess = OpenProcess(1048576 Or 1, ByVal 0&, IDProc)
    If hProcess = 0 Then Exit Sub
    TerminateProcess hProcess, 0&
    CloseHandle hProcess
End Sub
Public Function FindWindowEx(Optional ByVal hWndParent As Long = -1, Optional ByVal hWndChildAfter As Long = 0, Optional ByVal lpszClassName As String = vbNullString, Optional ByVal lpszWindowName As String = vbNullString) As Long
    hWndParent = HWndOrCurrent(hWndParent)
    FindWindowEx = FindWindowExA(hWndParent, hWndChildAfter, lpszClassName, lpszWindowName)
End Function
Public Sub Flash(Optional ByVal Enable = True, Optional ByVal ByHWnd As Long = -1)
    ByHWnd = HWndOrCurrent(ByHWnd)
    If Enable Then FlashWindow ByHWnd, 0 Else FlashWindow ByHWnd, 1
End Sub
Public Sub Focus(Optional ByVal ByHWnd As Long = -1)
    ByHWnd = HWndOrCurrent(ByHWnd)
    Dim ThreadID1 As Long, ThreadID2 As Long, nRet As Long
    ThreadID1 = GetWindowThreadProcessId(Foreground, ByVal 0&)
    ThreadID2 = GetWindowThreadProcessId(ByHWnd, ByVal 0&)
    If ThreadID1 <> ThreadID2 Then
        Call AttachThreadInput(ThreadID1, ThreadID2, True)
        nRet = SetForegroundWindow(ByHWnd)
        Call AttachThreadInput(ThreadID1, ThreadID2, False)
    Else
        nRet = SetForegroundWindow(ByHWnd)
    End If
    If IsIconic(ByHWnd) Then Call ShowWindow(ByHWnd, 9) _
    Else Call ShowWindow(ByHWnd, 5)
End Sub
Public Sub Ghost(Optional ByVal Enable As Boolean = True, Optional ByVal ByHWnd As Long = -1)
    ByHWnd = HWndOrCurrent(ByHWnd)
    If Enable = True Then SetWindowLong ByHWnd, (-20), GetWindowLong(ByHWnd, (-20)) Or &H20& _
    Else SetWindowLong ByHWnd, (-20), GetWindowLong(ByHWnd, (-20)) And Not &H20&
End Sub
Public Function GetWnd(Optional ByVal wCmd As Long = 6, Optional ByVal ByHWnd As Long = -1) As Long
    ByHWnd = HWndOrCurrent(ByHWnd)
    GetWnd = GetWindow(ByHWnd, wCmd)
End Function
Public Function IDProc(Optional ByVal ByHWnd As Long = -1) As Long
    ByHWnd = HWndOrCurrent(ByHWnd)
    Dim tProc As Long: GetWindowThreadProcessId ByHWnd, tProc
    IDProc = tProc
End Function
Public Function IsWindow(Optional ByVal ByHWnd As Long = -1)
    ByHWnd = HWndOrCurrent(ByHWnd)
    IsWindow = IsWindowA(ByHWnd)
End Function

Public Function ModuleBinary(Optional ByHWnd As Long = -1) As String
    ByHWnd = HWndOrCurrent(ByHWnd)
    Dim Str As String, StrLen As Long
    StrLen = 300
    Str = Space(StrLen)
    StrLen = GetWindowModuleFileName(ByHWnd, Str, StrLen)
    ModuleBinary = Strings.Left(Str, StrLen)
End Function
Public Sub Move(ByVal X As Integer, ByVal Y As Integer, Optional ByVal ByHWnd As Long = -1)
    ByHWnd = HWndOrCurrent(ByHWnd)
    Dim REC As RECT: GetWindowRect ByHWnd, REC
    Dim intTopMost As Integer
    If OnTop(ByHWnd) Then intTopMost = -1 Else intTopMost = -2
    SetWindowPos ByHWnd, intTopMost, REC.Left + X, REC.Top + Y, REC.Right - REC.Left, REC.Bottom - REC.Top, &H40
End Sub
Public Sub MoveWindow(ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, Optional ByVal ByHWnd As Long = -1)
    ByHWnd = HWndOrCurrent(ByHWnd)
    MoveWindowA ByHWnd, X, Y, nWidth, nHeight, 1
End Sub
Public Function Name(Optional ByVal ByHWnd As Long = -1) As String
    ByHWnd = HWndOrCurrent(ByHWnd)
    Dim Length As Long
    Dim sItem  As String
    Length = GetWindowTextLength(ByHWnd)
    sItem = Space$(Length + 1)
    Length = GetWindowText(ByHWnd, sItem, Length + 1)
    Name = Mid(sItem, 1, Len(sItem) - 1)
End Function
Public Function NextHWnd(Optional ByVal NearHWnd As Long = -1) As Long
    NearHWnd = HWndOrCurrent(NearHWnd)
    NextHWnd = GetWindow(NearHWnd, 2 Or 4)
End Function
Public Sub Opacity(ByVal vData As Integer, Optional ByVal ByHWnd As Long = -1)
    ByHWnd = HWndOrCurrent(ByHWnd)
    Dim ExStyle As Long, tTemp As Long
    ExStyle = GetWindowLong(ByHWnd, -20)
    If ExStyle <> (GetWindowLong(ByHWnd, -20) Or &H80000) Then
        ExStyle = (GetWindowLong(ByHWnd, -20) Or &H80000)
        Call SetWindowLong(ByHWnd, -20, ExStyle)
    End If
    tTemp = (SetLayeredWindowAttributes(ByHWnd, vbGreen, vData, IIf(True, &H2, &H1)) <> 0)
End Sub
Public Sub Position(ByVal X As Long, ByVal Y As Long, Optional ByVal ByHWnd As Long = -1)
    ByHWnd = HWndOrCurrent(ByHWnd)
    Dim REC As RECT: GetWindowRect ByHWnd, REC
    Dim intTopMost As Integer
    If OnTop(ByHWnd) Then intTopMost = -1 Else intTopMost = -2
    SetWindowPos ByHWnd, intTopMost, X, Y, REC.Right - REC.Left, REC.Bottom - REC.Top, &H40
End Sub
Public Function PostMessage(ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, Optional ByVal ByHWnd As Long = -1) As Long: ByHWnd = HWndOrCurrent(ByHWnd): PostMessage = PostMessageA(ByHWnd, wMsg, wParam, lParam): End Function
Public Sub Quit(Optional ByVal ByHWnd As Long = -1)
    ByHWnd = HWndOrCurrent(ByHWnd)
    If PostMessage(18, 0, 0, ByHWnd) <> 0 Then Call PostMessage(16, 0, 0, ByHWnd)
End Sub
Public Sub ReadStop(Optional ByVal Enable As Boolean = True, Optional ByVal ByHWnd As Long = -1)
    ByHWnd = HWndOrCurrent(ByHWnd)
    GetWindowVar(fHWnd).Once = Enable
End Sub
Public Sub Refresh()
    Static shWnd As String, lHWnd As Long
    fHWnd = -1
    GetCurrentWindow
    If lHWnd <> CurrentHwnd Then
        lHWnd = CurrentHwnd
        If Not GetWindowVar(Foreground).Once Then fEnumWindowsCallBack CurrentHwnd, 0
    End If
    If GetWindowVar(Foreground).id = -1 Then SetWindowVar Foreground
End Sub
Public Sub ResetHide(): ResetWindowHideVar: End Sub
Public Sub Resize(Optional ByVal Enable As Boolean = True, Optional ByVal ByHWnd As Long = -1)
    ByHWnd = HWndOrCurrent(ByHWnd)
    If Enable <> True Then
        Call SetWindowLong(ByHWnd, (-16&), GetWindowLong(ByHWnd, (-16&)) And Not (&H40000 Or &H20000 Or &H10000))
        Call SetWindowPos(ByHWnd, 0&, 0&, 0&, 0&, 0&, 2 Or 1 Or &H4 And Not &H20)
    Else
        Call SetWindowLong(ByHWnd, (-16&), GetWindowLong(ByHWnd, (-16&)) Or (&H40000 Or &H20000 Or &H10000))
        Call SetWindowPos(ByHWnd, 0&, 0&, 0&, 0&, 0&, 2 Or 1 Or &H4 Or &H20)
    End If
End Sub
Public Sub Restore(Optional ByVal ByHWnd As Long = -1): ByHWnd = HWndOrCurrent(ByHWnd): Call ShowWindow(ByHWnd, 9): End Sub
Public Sub ScriptAll(): fEnumWindows: End Sub
Public Sub ScriptChild(ByVal PathString As String, Optional ByVal ByHWnd As Long = -1)
    ByHWnd = HWndOrCurrent(ByHWnd)
    Dim wh As Long: wh = GetWindow(ByHWnd, 5)
    Do While wh <> 0
        Script.Win PathString, wh
        wh = GetWindow(wh, 2)
        SetWindowVar wh
    Loop
End Sub
Public Sub SendKey(ByVal veKey As Integer, Optional ByVal veShift As Integer = 0, Optional ByVal vbExtendedKey As Boolean = False, Optional ByVal ByHWnd As Long = -1)
    ByHWnd = HWndOrCurrent(ByHWnd)
    Dim vhTargetWnd As Long: vhTargetWnd = HWndOrCurrent(ByHWnd)
    Dim xbMemBuffer(255) As Byte
    Dim xbNewBuffer(255) As Byte
    Dim nKeyParam        As Long
    Dim nTargetThreadID  As Long
    nTargetThreadID = GetWindowThreadProcessId(vhTargetWnd, ByVal 0&)
    AttachThreadInput GetCurrentThreadId, nTargetThreadID, 1&
    nKeyParam = MapVirtualKey(veKey, 0) * &H10000
    If vbExtendedKey Then nKeyParam = nKeyParam Or &H1000000  '# bit 24
    GetKeyboardState xbMemBuffer(0)
    If (veShift And vbShiftMask) <> 0 Then xbNewBuffer(vbKeyShift) = &H80
    If (veShift And vbCtrlMask) <> 0 Then xbNewBuffer(vbKeyControl) = &H80
    SetKeyboardState xbNewBuffer(0)
    Window.PostMessage &H100, veKey, nKeyParam, vhTargetWnd
    Window.PostMessage &H101, veKey, nKeyParam Or &HC0000000, vhTargetWnd
    wV.DoSleep 10
    SetKeyboardState xbMemBuffer(0)
    AttachThreadInput GetCurrentThreadId, nTargetThreadID, 0
End Sub
Public Function SendMessage(ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam, Optional ByVal ByHWnd As Long = -1) As Long: ByHWnd = HWndOrCurrent(ByHWnd): SendMessage = SendMessageA(ByHWnd, wMsg, wParam, lParam): End Function
Public Function SendMessageString(ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam, Optional ByVal ByHWnd As Long = -1) As Long: ByHWnd = HWndOrCurrent(ByHWnd): SendMessageString = SendMessageS(ByHWnd, wMsg, wParam, lParam): End Function
Public Sub SendString(ByVal TextString As String, Optional ByVal ByHWnd As Long = -1): ByHWnd = HWndOrCurrent(ByHWnd): Call SendMessageS(ByHWnd, 12, 256, TextString): End Sub
Public Sub SetParent(ByVal hwnd As Long, Optional ByVal ByHWnd As Long = -1): ByHWnd = HWndOrCurrent(ByHWnd): Call SetParentS(ByHWnd, hwnd): End Sub
Public Sub Show(Optional ByVal nCmdShow As Integer = 0, Optional ByVal ByHWnd As Long = -1)
    ByHWnd = HWndOrCurrent(ByHWnd)
    ShowWindow ByHWnd, 5
    SetForegroundWindow ByHWnd
    Focus ByHWnd
End Sub
Public Sub Size(ByVal nWidth As Integer, ByVal nHeight As Integer, Optional ByVal ByHWnd As Long = -1)
    ByHWnd = HWndOrCurrent(ByHWnd)
    Dim REC        As RECT: GetWindowRect ByHWnd, REC
    Dim intTopMost As Integer
    If OnTop(ByHWnd) Then intTopMost = -1 Else intTopMost = -2
    SetWindowPos ByHWnd, intTopMost, REC.Left, REC.Top, REC.Left + nWidth, REC.Top + nHeight, &H40
End Sub
Public Function Text(Optional ByVal ByHWnd As Long = -1) As String
    On Error Resume Next
    ByHWnd = HWndOrCurrent(ByHWnd)
    Dim txtlen As Long, txt As String
    txtlen = Window.SendMessage(&HE, 0, 0, ByHWnd)
    If txtlen = 0 Then Exit Function
    txtlen = txtlen + 1
    txt = Space$(txtlen)
    txtlen = Window.SendMessageString(&HD, txtlen, ByVal txt, ByHWnd)
    Text = Strings.Left$(txt, txtlen)
End Function
Public Sub Unload(Optional ByVal ByHWnd As Long = -1): ByHWnd = HWndOrCurrent(ByHWnd): Call SendMessageA(ByHWnd, &H10, 0, 0): End Sub
