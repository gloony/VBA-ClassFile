Option Explicit
'Ceci n'est pas utilisé dans le code, mais c'est fun de montrer comment faire du multi thread sur vba :P
'Spoiler: C'est pas stable en VBA xD
Private Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function CreateThread Lib "Kernel32" (lpThreadAttributes As Any, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadID As Long) As Long
Private Declare Function TerminateThread Lib "Kernel32" (ByVal hThread As Long, ByVal dwExitCode As Long) As Long
Private Declare Function CloseHandle Lib "Kernel32" (ByVal hObject As Long) As Long
Private hThread As Long, hThreadID As Long

Public Sub AsyncThread()
    Sleep 2000
    MsgBox "alkjsd"
    hThread = 0
End Sub
Private Sub CreateAsync()
    hThread = CreateThread(ByVal 0&, ByVal 0&, AddressOf AsyncThread, ByVal 0&, ByVal 0&, hThreadID)
    CloseHandle hThread
End Sub
Private Sub TerminateAsync()
    If hThread <> 0 Then TerminateThread hThread, 0
End Sub
