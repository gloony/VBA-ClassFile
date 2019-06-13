Option Explicit
Private Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)

Private objShell As Object

Private Sub Class_Initialize(): Set objShell = CreateObject("Shell.Application"): End Sub
Private Sub Class_Terminate(): Set objShell = Nothing: End Sub

Public Sub AddFilesToZip(sFileNames() As String, ByVal zipPath As String, Optional ByVal FlushOld As Boolean = True)
    Dim i As Long, iCount As Long
    If FlushOld Then CreateEmptyZip zipPath
    On Error Resume Next
    For i = LBound(sFileNames) To UBound(sFileNames)
        objShell.Namespace("" & zipPath).CopyHere "" & sFileNames(i), &H14
        iCount = objShell.Namespace("" & zipPath).Items.Count
        Do Until iCount > i
            Sleep 100
            iCount = objShell.Namespace("" & zipPath).Items.Count
        Loop
    Next
    On Error GoTo 0
End Sub
Public Sub AddFileToZip(ByVal sFileName As String, ByVal zipPath As String, Optional FlushOld As Boolean = False)
    Dim tFile(0) As String: tFile(0) = sFileName
    AddFilesToZip tFile, zipPath, FlushOld
End Sub
Public Sub CreateEmptyZip(ByVal zipPath As String)
    Dim strZIPHeader As String: strZIPHeader = Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, vbNullChar)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CreateTextFile(zipPath).Write strZIPHeader
    Set fso = Nothing
End Sub
Public Function UnzipToFolder(ByVal sFolderName As String, ByVal zipPath As String) As Boolean: Call objShell.Namespace("" & sFolderName).CopyHere(objShell.Namespace("" & zipPath).Items): End Function
