Option Explicit
'Ceci n'est pas utilisé dans le code, mais c'est fun de montrer comment faire des tâches plannifié sur vba :P
Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long

Public Function CronStart(ByVal nIDEvent As Long) As Boolean
    If nIDEvent <= 2 Then Exit Function
    If ThisWorkbook.Sheets("Cron").Range("B" & nIDEvent).Value <> 0 Then CronRemove nIDEvent
    CronStart = (SetTimer(Application.hwnd, nIDEvent, ThisWorkbook.Sheets("Cron").Range("D" & nIDEvent).Value, AddressOf CronProc) <> 0)
    If CronStart Then ThisWorkbook.Sheets("Cron").Range("B" & nIDEvent).Value = Application.hwnd
    ThisWorkbook.Sheets("Cron").Range("C" & nIDEvent).Value = 0
End Function
Private Function CronProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case Msg
    Case 275
        ThisWorkbook.Sheets("Cron").Range("C" & wParam).Value = GetTickCount
        ExecuteCode ThisWorkbook.Sheets("Cron").Range("E" & wParam).Value
    End Select
End Function
Public Sub CronRemove(Optional ByVal nIDEvent As Long = -1)
    If nIDEvent = -1 Then
        Dim i As Long
        For i = 3 To ThisWorkbook.Sheets("Cron").Range("D" & rows.Count).End(xlUp).Row Step 1
            If ThisWorkbook.Sheets("Cron").Range("B" & i).Value <> 0 Then CronRemove i
        Next i
    ElseIf ThisWorkbook.Sheets("Cron").Range("B" & nIDEvent).Value <> "" Then
        KillTimer ThisWorkbook.Sheets("Cron").Range("B" & nIDEvent).Value, nIDEvent
        ThisWorkbook.Sheets("Cron").Range("B" & nIDEvent).Value = ""
        ThisWorkbook.Sheets("Cron").Range("C" & nIDEvent).Value = ""
    End If
End Sub
