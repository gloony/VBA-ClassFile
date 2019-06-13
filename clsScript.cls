Option Explicit
Private MainJScriptControl As ScriptControl, MainVBScriptControl As ScriptControl
Private nScript As clsScript
Private lCommand As String

Private Sub BetterErrHandler(ByVal ErrObject As Object, Optional ByVal ScriptString As String = "")
    Dim Msg As frmMsg
    Set Msg = New frmMsg
    Msg.Caption = "WDE [" & ErrObject.Source & "]"
    Msg.lblMsg.Caption = "Error: " & ErrObject.Number & vbCrLf
    Msg.lblMsg.Caption = Msg.lblMsg.Caption & "Description: " & ErrObject.Description & vbCrLf
    If ScriptString <> "" Then Msg.lblMsg.Caption = Msg.lblMsg.Caption & "Script: " & ScriptString & vbCrLf
    If ErrObject.Line <> 1 Then Msg.lblMsg.Caption = Msg.lblMsg.Caption & "Line: " & ErrObject.Line
    If ErrObject.Column <> 0 Then Msg.lblMsg.Caption = Msg.lblMsg.Caption & " / Column: " & ErrObject.Column & vbCrLf
    If ErrObject.Text <> "" Then Msg.lblMsg.Caption = Msg.lblMsg.Caption & "Text: " & ErrObject.Text
    Msg.Show
    Err.Clear
End Sub
Private Sub PrepareScriptControl(ByRef scObject As ScriptControl, ByVal ScriptType As String, Optional ByVal Args As Variant = "")
    If scObject.Language <> "" Then scObject.Reset
    With scObject
        Select Case ScriptType
        Case "JScript"
            .AllowUI = True
            .Language = "JScript"
            .AddCode "function alert(m){ Computer.Msg(m); }"
            .AddCode "function dump(obj){ var out = ''; for (var i in obj) out += i + ': ' + obj[i] + ""\n""; alert(out); }"
            .AddCode "function prompt(m){ return Computer.Prompt(m); }"
            .AddCode "function startsWith(str, suffix){ return str.indexOf(suffix, suffix.length) !== -1; }"
            .AddCode "function insWith(str, suffix){ return str.indexOf(suffix) !== -1; }"
            .AddCode "function endsWith(str, suffix){ return str.indexOf(suffix, str.length - suffix.length) !== -1; }"
            .AddCode "var Return = """";"
            If Args <> "" Then .AddCode "var Args = """ & Args & """; alert(Args);"
        Case "VBScript"
            .AllowUI = True
            .Language = "VBScript"
            .AddCode "Const MF_CHECKED = &H8&"
            .AddCode "Const MF_APPEND = &H100&"
            .AddCode "Const TPM_LEFTALIGN = &H0&"
            .AddCode "Const MF_DISABLED = &H2&"
            .AddCode "Const MF_GRAYED = &H1&"
            .AddCode "Const MF_SEPARATOR = &H800&"
            .AddCode "Const MF_STRING = &H0&"
            .AddCode "Const MF_POPUP = &H10&"
            .AddCode "Const TPM_RETURNCMD = &H100&"
            .AddCode "Const TPM_RIGHTBUTTON = &H2&"
            .AddCode "Dim Return"
            If Args <> "" Then .AddCode "Const Args = """ & Args & """"
        End Select
        .AddObject "Application", Application
        .AddObject "shGUI", shGUI
        .AddObject "Clip", Clip
        .AddObject "Computer", Computer
        .AddObject "Explorer", Explorer
        .AddObject "File", File
        .AddObject "iExplore", iExplore
        '.AddObject "inet", iNet
        .AddObject "Keyboard", Keyboard
        .AddObject "Menu", Menu
        .AddObject "Mouse", Mouse
        .AddObject "Joypad", Joypad
        .AddObject "Script", Me
        .AddObject "Str", Str
        .AddObject "ThisWorkbook", ThisWorkbook
        .AddObject "V", Var
        .AddObject "Window", Window
    End With
End Sub

Private Sub Class_Terminate()
    Set MainJScriptControl = Nothing
    Set MainVBScriptControl = Nothing
End Sub

Public Property Get LastCommand() As String
    LastCommand = lCommand
End Property

Public Function AddCode(ByVal CommandString As String, Optional ByVal getReturn As Boolean = False, Optional ByVal ScriptType As String = "Auto") As String
    lCommand = CommandString
    If ScriptType = "Auto" Then If iRight(CommandString, ";") Or iRight(CommandString, "}") Then ScriptType = "JScript" Else ScriptType = "VBScript"
    On Error Resume Next
    Select Case UCase(ScriptType)
    Case "JSCRIPT"
        If MainJScriptControl Is Nothing Then Reset
        MainJScriptControl.AddCode CommandString
        If Err.Number <> 0 Then BetterErrHandler MainJScriptControl.Error
        If getReturn Then If Not MainJScriptControl.CodeObject.Return Is Nothing Then AddCode = MainJScriptControl.CodeObject.Return
    Case "VBSCRIPT"
        If MainVBScriptControl Is Nothing Then Reset
        MainVBScriptControl.AddCode CommandString
        If Err.Number <> 0 Then BetterErrHandler MainVBScriptControl.Error
        If getReturn Then If Not MainVBScriptControl.CodeObject.Return Is Nothing Then AddCode = MainVBScriptControl.CodeObject.Return
    End Select
    On Error GoTo 0
End Function
Public Function Execute(ByVal FilePath As String, Optional ByVal Args As Variant = "", Optional ByVal ByLine As Boolean = False, Optional ByVal getReturn As Boolean = False, Optional ByVal ScriptType As String = "Auto", Optional ByVal StandAlone As Boolean = False) As String
    If StandAlone Then
        Dim nScript As clsScript
        Set nScript = New clsScript
        Execute = nScript.Execute(FilePath, Args, ByLine, getReturn, ScriptType)
        Set nScript = Nothing
        Exit Function
    End If
    Dim tPath As String
    tPath = FilePath
    If Not File.Exist(FilePath) Then FilePath = ThisWorkbook.Path & "\scripts\" & tPath
    If Not File.Exist(FilePath) Then FilePath = ThisWorkbook.Path & "\scripts\" & tPath & ".js"
    If Not File.Exist(FilePath) Then FilePath = ThisWorkbook.Path & "\scripts\" & tPath & ".vbs"
    If Not File.Exist(FilePath) Then FilePath = ThisWorkbook.Path & "\scripts\" & tPath & ".bat"
    If Not File.Exist(FilePath) Then
        Computer.Form "Fichier " & tPath & vbCrLf & "Introuvable", "WDE [Script.Execute]"
        Exit Function
    End If
    If ScriptType <> "Auto" Then
    ElseIf iRight(FilePath, ".js") Then
        ScriptType = "JScript"
    ElseIf iRight(FilePath, ".vbs") Then
        ScriptType = "VBScript"
    ElseIf iRight(FilePath, ".bat") Then
        Shell "CMD /k " & FilePath, vbNormalFocus
    Else
        Exit Function
    End If
    If Dir(FilePath) = "" Then Exit Function
    Dim ScriptScriptControl As ScriptControl
    Set ScriptScriptControl = New ScriptControl
    PrepareScriptControl ScriptScriptControl, ScriptType, Args
    On Error Resume Next
    Dim nIF As Integer
    nIF = FreeFile
    Dim StringLine As String, PosLine As Long
    Dim StringAll  As String
    Open FilePath For Input As #nIF
    PosLine = 0
    While Not EOF(nIF)
        Line Input #nIF, StringLine
        If ByLine Then
            PosLine = PosLine + 1
            AddCode StringLine, , ScriptType
            If Err.Number <> 0 Then
                MsgBox "Error: " & Err.Number & vbCrLf & _
                    "Script: " & FilePath & vbCrLf & _
                    "Line: " & PosLine & vbCrLf & _
                    "Command: " & StringLine & vbCrLf & _
                    Err.Description, vbCritical, "wV [" & Err.Source & "]"
                Err.Clear
            End If
            lCommand = StringLine
        Else
            If UCase(ScriptType) = "VBSCRIPT" Then
                If UCase(Left(StringLine, 4)) = "DIM " And InStr(1, StringLine, " = ", vbTextCompare) > 0 Then
                    Dim sVar As String
                    sVar = Right(StringLine, Len(StringLine) - 4)
                    sVar = Left(sVar, InStr(1, sVar, " = ") - 1)
                    If InStr(1, StringLine, ": ", vbTextCompare) > 0 Then If InStr(1, StringLine, ": ", vbTextCompare) < InStr(1, StringLine, " = ", vbTextCompare) Then sVar = ""
                    If sVar <> "" Then
                        StringLine = Right(StringLine, Len(StringLine) - 4)
                        StringLine = "Dim " & sVar & ": " & StringLine
                    End If
                End If
            End If
            StringAll = StringAll & vbCrLf & StringLine
            lCommand = StringAll
        End If
    Wend
    Close #nIF
    If Not ByLine Then ScriptScriptControl.AddCode StringAll
    If Err.Number <> 0 Then BetterErrHandler ScriptScriptControl.Error, FilePath
    If getReturn Then If Not ScriptScriptControl.CodeObject.Return Is Nothing Then Execute = ScriptScriptControl.CodeObject.Return
    On Error GoTo 0
    ScriptScriptControl.Reset
    Set ScriptScriptControl = Nothing
End Function
Public Function Func(ByVal FilePath As String, Optional Args As String = "", Optional ByVal ByLine As Boolean = False, Optional ByVal ScriptType As String = "Auto", Optional ByVal StandAlone As Boolean = False) As String
    Dim tPath As String
    tPath = FilePath
    Func = Execute(FilePath, Args, ByLine, True, ScriptType, StandAlone)
End Function
Public Sub Reset()
    If Not (MainJScriptControl Is Nothing) Then Set MainJScriptControl = Nothing
    If Not (MainVBScriptControl Is Nothing) Then Set MainVBScriptControl = Nothing
    Set MainJScriptControl = New ScriptControl
    Set MainVBScriptControl = New ScriptControl
    PrepareScriptControl MainJScriptControl, "JScript"
    PrepareScriptControl MainVBScriptControl, "VBScript"
End Sub
Public Sub Win(ByVal StringPath As String, ByVal hwnd As Long)
    Dim tFile As String
    tFile = StringPath
    If Not File.Exist(StringPath) Then StringPath = Window.ScriptDir & tFile
    If Not File.Exist(StringPath) Then StringPath = Window.ScriptDir & tFile & ".js"
    If Not File.Exist(StringPath) Then StringPath = Window.ScriptDir & tFile & ".vbs"
    If Not File.Exist(StringPath) Then Exit Sub
    Dim lHWnd As Long
    lHWnd = Window.CurHWnd
    Window.CurHWnd = hwnd
    Execute StringPath, , False
    Window.CurHWnd = -1
End Sub
