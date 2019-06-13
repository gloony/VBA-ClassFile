Option Explicit
Private Declare Function CreateFile Lib "Kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "Kernel32" (ByVal hObject As Long) As Long
Private Declare Function fCreateShellLink Lib "STKIT432.DLL" (ByVal lpstrFolderName As String, ByVal lpstrLinkName As String, ByVal lpstrLinkPath As String, ByVal lpstrLinkArgs As String) As Long
Private Declare Function LocalFileTimeToFileTime Lib "Kernel32" (lpLocalFileTime As FILETIME, lpFileTime As FILETIME) As Long
Private Declare Function SetFileTime Lib "Kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Declare Function SystemTimeToFileTime Lib "Kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long
Private Declare Function ShellExecuteA Lib "shell32.dll" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Private Declare Function WinExecS Lib "Kernel32" Alias "WinExec" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long

Private Type FILETIME
    dwLowDateTime               As Long
    dwHighDateTime              As Long
End Type
Private Type SYSTEMTIME
    wYear                       As Integer
    wMonth                      As Integer
    wDayOfWeek                  As Integer
    wDay                        As Integer
    wHour                       As Integer
    wMinute                     As Integer
    wSecond                     As Integer
    wMilliseconds               As Integer
End Type

Public Property Get TimeStamp(ByVal PathString As String) As String: TimeStamp = FileDateTime(PathString): End Property
Public Property Let TimeStamp(ByVal PathString As String, ByVal sDateTime As String)
    Dim sDate As String: sDate = Left(sDateTime, Len("##.##.####"))
    Dim sTime As String: sTime = Right(sDateTime, Len("##:##:##"))
    Dim dDate         As Date
    Dim udtFileTime   As FILETIME
    Dim udtLocalTime  As FILETIME
    Dim udtSystemTime As SYSTEMTIME
    Dim lFileHandle   As Long
    dDate = Format(sDate & " " & sTime, "DD-MM-YYYY HH:MM:SS")
    udtSystemTime.wYear = Year(dDate)
    udtSystemTime.wMonth = Month(dDate)
    udtSystemTime.wDay = Day(dDate)
    udtSystemTime.wDayOfWeek = Weekday(dDate) - 1
    udtSystemTime.wHour = Hour(dDate)
    udtSystemTime.wMinute = Minute(dDate)
    udtSystemTime.wSecond = Second(dDate)
    udtSystemTime.wMilliseconds = 0
    SystemTimeToFileTime udtSystemTime, udtLocalTime
    LocalFileTimeToFileTime udtLocalTime, udtFileTime
    lFileHandle = CreateFile(PathString, &H40000000, &H1 Or &H2, ByVal 0&, 3, 0, 0)
    If lFileHandle >= 0 Then SetFileTime lFileHandle, udtFileTime, udtFileTime, udtFileTime
    CloseHandle lFileHandle
End Property

Public Sub Append(ByVal PathString As String, ByVal ValueString As String)
    On Error GoTo Die
    Dim nIF As Integer: nIF = FreeFile
    Open PathString For Append As #nIF
    Print #nIF, ValueString
    Close #nIF
Die:
End Sub
Public Sub Copy(ByVal PathString As String, ByVal DestinationString As String, Optional ByVal Mode As String = "Auto")
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If UCase(Mode) = "AUTO" Then
        If Exist(PathString, vbDirectory) Then Mode = "DIR"
        If Exist(PathString) Then Mode = "FILE"
    End If
    Select Case UCase(Mode)
    Case "Dir": objFSO.CopyFolder PathString, DestinationString
    Case "FILE": objFSO.CopyFile PathString, DestinationString
    End Select
    Set objFSO = Nothing
End Sub
Public Function CreateShortcut(ByVal strFolder As String, ByVal strShortcutName As String, ByVal strShortcutTarget As String, ByVal strShortcutCommandLineParams As String) As Boolean
    Select Case UCase(Right(strShortcutName, 4))
    Case ".LNK", ".PIF", ".URL"
        strShortcutName = Left(strShortcutName, Len(strShortcutName) - 4)
    End Select
    Call fCreateShellLink(strFolder, strShortcutName, strShortcutTarget, strShortcutCommandLineParams)
End Function
Public Function Dialog(Optional Owner As Long = 0&, Optional FileName As String, Optional Filter As String = "All (*.*)| *.*", Optional MultiSelect As Boolean = False, Optional FileMustExist As Boolean = True, Optional ReadOnly As Boolean = False, Optional HideReadOnly As Boolean = False, Optional FilterIndex As Long = 1, Optional InitDir As String, Optional DefaultExt As String = "", Optional FileTitle As String, Optional DlgTitle As String) As String
    Dim Dlg As clsDlg
    Set Dlg = New clsDlg
    Dialog = Dlg.ShowOpen(Owner, FileName, Filter, MultiSelect, FileMustExist, ReadOnly, HideReadOnly, FilterIndex, InitDir, DefaultExt, FileTitle)
    Set Dlg = Nothing
End Function
Public Sub Download(ByVal URLString As String, ByVal PathString As String): Call URLDownloadToFile(0, URLString, PathString, 0, 0): End Sub
Public Sub Execute(ByVal PathString As String): Call Shell(PathString, vbNormalFocus): End Sub
Public Sub Kill(ByVal PathString As String)
    If File.Exist(iLeft_Last) Then Kill PathString
End Sub
Public Sub MkDir(ByVal PathString As String)
    If Not File.Exist(iLeft_Last, vbDirectory) Then MkDir PathString
End Sub
Public Sub MkFile(ByVal PathString As String)
    If Not File.Exist(iLeft_Last) Then File.Output PathString, ""
End Sub
Public Sub Move(ByVal PathString As String, ByVal DestinationString As String, Optional ByVal Mode As String = "Auto")
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If UCase(Mode) = "AUTO" Then
        If Exist(PathString, vbDirectory) Then Mode = "DIR"
        If Exist(PathString) Then Mode = "FILE"
    End If
    Select Case UCase(Mode)
    Case "Dir": objFSO.MoveFolder PathString, DestinationString
    Case "FILE": objFSO.MoveFile PathString, DestinationString
    End Select
    Set objFSO = Nothing
End Sub
Public Sub Output(ByVal PathString As String, ByVal ValueString As String)
    On Error GoTo Die
    Dim nIF As Integer: nIF = FreeFile
    Open PathString For Output As #nIF
        Print #nIF, ValueString
    Close #nIF
Die:
End Sub
Public Sub ShellExecute(ByVal StringPath As String, Optional ByVal ArgString As String = vbNullString, Optional nShowCmd As Long)
     If ArgString = vbNullString Then ShellExecuteA 0, "Open", StringPath, vbNullString, "", nShowCmd _
    Else ShellExecuteA 0, "Open", StringPath, Chr(34) & ArgString & Chr(34), "", nShowCmd
End Sub
Public Sub ShellEecuteFolder(ByVal StringPath As String)
    On Error GoTo Die
    Dim obj, RepP, F, F1
    Dim Ext As String, Chem As String
    Set obj = CreateObject("Scripting.FileSystemObject")
    If Not File.Exist(StringPath, vbDirectory) Then Exit Sub
    Set RepP = obj.GetFolder(StringPath)
    Chem = StringPath
    If Right(Chem, 1) <> "\" Then Chem = Chem & "\"
    Set F = RepP.Files
    For Each F1 In F
        File.Execute F1.ParentFolder & "\" & F1.Name
    Next F1
Die:
    On Error GoTo 0
    Set obj = Nothing
    Set F = Nothing
End Sub
Public Sub WinExec(ByVal StringPath As String): Call WinExecS(StringPath, 1): End Sub

Public Function Exist(ByVal PathString As String, Optional ByVal fAttribute As VbFileAttribute = vbNormal) As Boolean
    If Len(PathString) = 3 Then If Right(PathString, 2) = ":\" Then GoTo TestDrive
    On Error Resume Next
    If PathString = "" Or PathString = "*" Then Exit Function
    Exist = Dir(PathString, fAttribute) <> vbNullString
    If Err.Number <> 0 Then Exist = False
    On Error GoTo 0
    Exit Function
TestDrive:
    Dim objFSO As Object
    Exist = objFSO.DriveExists(Left(PathString, 1))
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFSO = Nothing
End Function
Public Function Locked(ByVal PathString As String) As Boolean
    Locked = True
    On Error GoTo NotFree
    Dim nIF As Long: nIF = FreeFile
    Open PathString For Append As #nIF: Locked = False: Close #nIF
NotFree:
    Locked = True
End Function
Public Function Read(ByVal PathString As String, Optional ByVal AvoidLines As Long = 0) As String
    On Error GoTo Die
    If PathString = "" Then Exit Function
    If File.Exist(PathString) Then
        Dim nIF As Integer, tStr As String, i As Long
        nIF = FreeFile
        Open PathString For Input As #nIF
        On Error GoTo CloseFile
        While Not EOF(nIF)
            Line Input #nIF, tStr
            If i >= AvoidLines Then If Read = "" Then Read = tStr Else Read = Read & vbCrLf & tStr
            i = i + 1
        Wend
CloseFile:
        Close #nIF
    End If
Die:
    On Error GoTo 0
End Function
