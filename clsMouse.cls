Option Explicit
Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As Any) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function ChildWindowFromPointEx Lib "user32" (ByVal Parent As Long, ByVal X As Long, ByVal Y As Long, ByVal Flags As Long) As Long
Private Declare Function GetCursorPosA Lib "user32" Alias "GetCursorPos" (lpPoint As POINTAPI) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dX As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Function ChildWindowFormPoint(ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal Flags As Long) As Long
    X = X - Window.Left(hwnd)
    Y = Y - Window.Top(hwnd)
    Dim lHWnd As Long
    lHWnd = ChildWindowFromPointEx(hwnd, X, Y, Flags)
    If lHWnd = 0 Then
        ChildWindowFormPoint = hwnd
        Exit Function
    ElseIf lHWnd <> hwnd Then
        Dim lPoint As POINTAPI
        lPoint.X = X
        lPoint.Y = Y
        Call ClientToScreen(hwnd, lPoint)
        Call ScreenToClient(lHWnd, lPoint)
        Dim lHWndRecurse As Long
        lHWndRecurse = ChildWindowFormPoint(lHWnd, X, Y, Flags)
        If lHWndRecurse <> 0 Then lHWnd = lHWndRecurse
    End If
    ChildWindowFormPoint = lHWnd
End Function

Public Property Get hwnd(Optional ByVal Flags As Long = -1) As Long
    Dim tXY As POINTAPI: GetCursorPosA tXY
    hwnd = WindowFromPoint(tXY.X, tXY.Y)
    If Flags <> -1 Then hwnd = ChildWindowFormPoint(hwnd, X, Y, Flags)
End Property
Public Property Let hwnd(Optional ByVal Flags As Long = -1, ByVal vData As Long)
    'Dim lPoint As POINTAPI
    'Call ScreenToClient(vData, lPoint)
    'Cursor_Position lPoint.X, lPoint.Y
End Property
Public Property Get X() As Long
    Dim tXY As POINTAPI: GetCursorPosA tXY
    X = tXY.X + 1
End Property
Public Property Let X(ByVal vData As Long)
    SetCursorPos vData, Y
End Property
Public Property Get Y() As Long
    Dim tXY As POINTAPI: GetCursorPosA tXY
    Y = tXY.Y + 1
End Property
Public Property Let Y(ByVal vData As Long)
    SetCursorPos X, vData
End Property

Public Sub ButtonDown(Optional wMode As Integer = 0)
    Select Case wMode
    Case 0
        mouse_event &H2, 0, 0, 1, 0
    Case 1
        mouse_event &H8, 0, 1, 0, 0
    Case 2
        mouse_event &H20, 0, 1, 0, 0
    End Select
End Sub
Public Sub ButtonUp(Optional wMode As Integer = 0)
    Select Case wMode
    Case 0
        mouse_event &H4, 0, 0, 1, 0
    Case 1
        mouse_event &H10, 0, 1, 0, 0
    Case 2
        mouse_event &H40, 0, 1, 0, 0
    End Select
End Sub
Public Sub Click(Optional wMode As Integer = 0)
    Select Case wMode
    Case 0
        mouse_event &H2, 0, 0, 1, 0
        mouse_event &H4, 0, 0, 1, 0
    Case 1
        mouse_event &H8, 0, 1, 0, 0
        mouse_event &H10, 0, 1, 0, 0
    Case 2
        mouse_event &H20, 0, 1, 0, 0
        mouse_event &H40, 0, 1, 0, 0
    End Select
End Sub
Public Sub Move(Optional ByVal X As Integer = 0, Optional ByVal Y As Integer = 0)
    Dim mPosition As POINTAPI
    GetCursorPosA mPosition
    SetCursorPos mPosition.X + X, mPosition.Y + Y
End Sub
Public Sub ScrollH(Optional ByVal Notches As Long = 1)
    Call mouse_event(&H1000, 0, 0, Notches * 120, 0)
End Sub
Public Sub ScrollV(Optional ByVal Notches As Long = 1)
    Call mouse_event(&H800, 0, 0, Notches * -120, 0)
End Sub

Public Function Color() As String
    Dim DC As Long: DC = CreateDC("DISPLAY", vbNullString, vbNullString, 0&)
    Color = Format$(GetPixel(DC, X, Y), "00000000") 'long
    DeleteDC DC
End Function
