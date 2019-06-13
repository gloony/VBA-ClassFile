Option Explicit
Private sClass As String, objEx As SHDocVw.InternetExplorer

Private Sub Class_Initialize()
    sClass = "CabinetWClass"
End Sub
Private Sub Class_Terminate()
    Set objEx = Nothing
End Sub

Public Property Let Class(ByVal ClassString As String)
    sClass = ClassString
End Property
Public Property Get Class() As String
    Class = sClass
End Property
Public Function GetObject(Optional ByVal lHWnd As Long = 0, Optional ByVal NameString As String = "", Optional ByVal ForceReset As Boolean = True) As SHDocVw.InternetExplorer
    If lHWnd = 0 Then lHWnd = Window.Foreground
    If Window.Class(Window.Foreground) <> sClass Then
        If Not (objEx Is Nothing) Then Set objEx = Nothing
        Exit Function
    End If
    If Not (objEx Is Nothing) And Not ForceReset Then
        Set GetObject = objEx
        Exit Function
    ElseIf Not (objEx Is Nothing) And ForceReset Then
        Set objEx = Nothing
    End If
    If NameString <> "" Then lHWnd = Window.FindWindow("", NameString)
    Dim sw As SHDocVw.ShellWindows
    Set sw = New SHDocVw.ShellWindows
    For Each objEx In sw
        If objEx Is Nothing Then
        ElseIf Not (UCase(objEx.FullName) Like "C:\WINDOWS*\EXPLORER.EXE") Then
        ElseIf UCase(objEx.Path) <> "C:\WINDOWS\" Then
        Else
            Dim tBool As Boolean
            If Not lHWnd = objEx.hwnd Then
            ElseIf iLeft(Str.Trim(Window.Name(lHWnd)), Str.Trim(objEx.LocationName)) Then
                tBool = True
            ElseIf iLeft(Str.Trim(Window.Name(lHWnd)), Str.Trim(objEx.LocationURL)) Then
                tBool = True
            ElseIf iLeft(Str.Trim("file:///" & Window.Name(lHWnd)), Str.Trim(objEx.LocationURL)) Then
                tBool = True
            End If
            If tBool Then
                Set GetObject = objEx
                Exit Function
            End If
        End If
    Next
End Function

Public Property Let Path(ByVal PathString As String)
    GetObject , , False
    If objEx Is Nothing Then Explorer.Create PathString _
   Else objEx.Navigate PathString
End Property
Public Property Get Path() As String
    GetObject , , False
    Path = objEx.LocationURL
End Property

Public Function hwnd() As Long
    GetObject , , False
    hwnd = objEx.hwnd
End Function

Public Function FocusedFile() As String
    GetObject , , False
    If Not (objEx Is Nothing) Then FocusedFile = objEx.Document.FocusedItem.Path
End Function

Public Sub Create(Optional ByVal PathString As String)
    If File.Exist(PathString, vbDirectory) Then File.ShellExecute "explorer", Chr(34) & PathString & Chr(34), 5 _
   Else: Shell "explorer.exe"
End Sub
Public Sub Quit()
    GetObject , , False
    objEx.Quit
    Set objEx = Nothing
End Sub
