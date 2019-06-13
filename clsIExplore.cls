Option Explicit
Private sClass As String, objIE As SHDocVw.InternetExplorer

Private Sub Class_Initialize()
    sClass = "IEFrame"
End Sub
Private Sub Class_Terminate()
    Set objIE = Nothing
End Sub

Public Property Let Class(ByVal ClassString As String)
    sClass = ClassString
End Property
Public Property Get Class() As String
    Class = sClass
End Property
Public Function GetObject(Optional ByVal lHWnd As Long = 0, Optional ByVal NameString As String = "", Optional ByVal ForceReset As Boolean = True) As SHDocVw.InternetExplorer
    If Not (objIE Is Nothing) And Not ForceReset Then
        Set GetObject = objIE
        Exit Function
    ElseIf Not (objIE Is Nothing) And ForceReset Then
        Set objIE = Nothing
    End If
    If lHWnd = 0 Then lHWnd = Window.Foreground
    If NameString <> "" Then lHWnd = Window.FindWindow("", NameString)
    If Window.Class(lHWnd) <> sClass Then Exit Function
    Dim sw As SHDocVw.ShellWindows
    Set sw = New SHDocVw.ShellWindows
    For Each objIE In sw
        If objIE Is Nothing Then
        ElseIf Not (UCase(objIE.FullName) Like "C:\PROGRAM FILES*\INTERNET EXPLORER\IEXPLORE.EXE") Then
        ElseIf UCase(objIE.Path) <> "C:\PROGRAM FILES\INTERNET EXPLORER\" And UCase(objIE.Path) <> "C:\PROGRAM FILES (X86)\INTERNET EXPLORER\" Then
        Else
            Dim tBool As Boolean
            If Not lHWnd = objIE.hwnd Then
            ElseIf iLeft(Str.Trim(Window.Name(lHWnd)), Str.Trim(objIE.LocationName)) Then
                tBool = True
            ElseIf iLeft(Str.Trim(Window.Name(lHWnd)), Str.Trim(objIE.Document.Title)) Then
                tBool = True
            ElseIf iLeft(Str.Trim(Window.Name(lHWnd)), Str.Trim(objIE.Document.nameProp)) Then
                tBool = True
            End If
            If tBool Then
                Set GetObject = objIE
                Exit Function
            End If
        End If
    Next
End Function
Public Function SetObject(ByVal nObjIE As SHDocVw.InternetExplorer)
    Set objIE = nObjIE
End Function

Public Property Let URL(ByVal URLString As String)
    GetObject , , False
    If objIE Is Nothing Then
        Create URLString
    Else
        objIE.Navigate URLString
    End If
End Property
Public Property Get URL() As String
    GetObject , , False
    URL = objIE.LocationURL
End Property
Public Property Let Value(ByVal ValueName As String, ByVal TextString As String)
    GetObject , , False
    If Not (objIE.Document.All(ValueName) Is Nothing) Then objIE.Document.All(ValueName).Value = TextString
End Property
Public Property Get Value(ByVal ValueName As String) As String
    GetObject , , False
    If Not (objIE.Document.All(ValueName) Is Nothing) Then Value = objIE.Document.All(ValueName).Value
End Property

Public Function Busy() As Boolean
    GetObject , , False
    Busy = objIE.Busy
End Function
Public Function hwnd() As Long
    GetObject , , False
    hwnd = objIE.hwnd
End Function

Public Sub Cancel()
    GetObject , , False
    objIE.Stop
End Sub
Public Sub Click(ByVal NameString As String)
    GetObject , , False
    If Not (objIE.Document.All(NameString) Is Nothing) Then objIE.Document.All(NameString).Click
End Sub
Public Sub Create(Optional ByVal URLString As String, Optional ByVal bTab As Boolean = False)
    If Not bTab Then
        Dim NewBrowser
        Set NewBrowser = CreateObject("InternetExplorer.Application")
        With NewBrowser
            .FullScreen = False
            .StatusBar = True
            .Visible = True
            'ExternalWindow_Flash (.hWnd)
            Window.Focus .hwnd
            If URL <> "" Then .Navigate URL
        End With
        Set NewBrowser = Nothing
    Else
        GetObject , , False
        objIE.Navigate2 URLString, 2048
    End If
End Sub
Public Sub Javascript(ByVal ScriptString As String)
    GetObject , , False
    objIE.Navigate "javascript: " & ScriptString
End Sub
Public Sub Quit()
    GetObject , , False
    objIE.Quit
    Set objIE = Nothing
End Sub
Public Sub Submit(ByVal FormString As String)
    GetObject , , False
    If Not (objIE.Document.All(FormString) Is Nothing) Then objIE.Document.All(FormString).Submit
End Sub
