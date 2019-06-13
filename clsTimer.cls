Option Explicit
Private hTimers() As Long
Private dTimers() As String
Private sTimers() As String
Private rTimers() As Boolean
Private pTimers() As Long

Public Function IsTimersInitialized() As Boolean
    On Error GoTo Die
    Dim i As Integer
    i = UBound(hTimers)
Die:
    If Err.Number = 0 Then IsTimersInitialized = True
End Function

Public Sub CreateFromRange(ByVal Row As Long)
    Me.Create shGUI.Range("G" & Row).Value, shGUI.Range("H" & Row).Value, shGUI.Range("F" & Row).Value, Row
End Sub

Public Sub DisableFromRange(ByVal Row As Long)
    If Not IsTimersInitialized Then Exit Sub
    Dim i As Long
    For i = 0 To UBound(rTimers) Step 1
        If pTimers(i) = Row And rTimers(i) Then
            Me.Disable i
            shGUI.Range("H" & Row).Interior.Color = vbRed
        End If
    Next i
End Sub

Public Function Create(ByVal Delay As String, ByVal ScriptString As String, Optional ByVal ExecuteNTime As Long = 0, Optional Row As Long = -1) As Long
    If Not IsTimersInitialized Then
        ReDim hTimers(0)
        ReDim dTimers(0)
        ReDim sTimers(0)
        ReDim rTimers(0)
        ReDim pTimers(0)
    Else
        ReDim Preserve hTimers(UBound(hTimers) + 1)
        ReDim Preserve dTimers(UBound(hTimers))
        ReDim Preserve sTimers(UBound(hTimers))
        ReDim Preserve rTimers(UBound(hTimers))
        ReDim Preserve pTimers(UBound(hTimers))
    End If
    hTimers(UBound(hTimers)) = ExecuteNTime
    dTimers(UBound(hTimers)) = Delay
    sTimers(UBound(hTimers)) = ScriptString
    rTimers(UBound(hTimers)) = True
    pTimers(UBound(hTimers)) = Row
    BeginLoop UBound(hTimers)
    Create = UBound(hTimers)
End Function
Public Sub Disable(Optional ByVal Index As Long = -1)
    If Not IsTimersInitialized Then Exit Sub
    If Index = -1 Then
        Dim i As Long
        For i = 0 To UBound(hTimers) Step 1
            rTimers(i) = False
        Next i
    Else
        rTimers(Index) = False
    End If
End Sub

Public Sub BeginLoop(ByVal Index As Long)
    If Not IsTimersInitialized Then Exit Sub
    Dim alertTime
    If hTimers(Index) = -1 Then
        alertTime = CDate(Date & " " & dTimers(Index))
        Application.OnTime alertTime, "'TimerCallBackLoop """ & Index & """'"
        If pTimers(Index) <> -1 Then shGUI.Range("H" & pTimers(Index)).Interior.Color = vbGreen
    ElseIf hTimers(Index) >= 0 Then
        alertTime = Now + TimeValue(dTimers(Index))
        Application.OnTime alertTime, "'TimerCallBackLoop """ & Index & """'"
        If pTimers(Index) <> -1 Then shGUI.Range("H" & pTimers(Index)).Interior.Color = vbGreen
    End If
End Sub

Public Sub DoLoop(ByVal Index As Long)
    If rTimers(Index) Then
        Init
        If pTimers(Index) <> -1 Then shGUI.Range("H" & pTimers(Index)).Interior.Color = vbYellow
        Script.AddCode sTimers(Index)
        If hTimers(Index) = -1 Or hTimers(Index) = 1 Then
            rTimers(Index) = False
            If pTimers(Index) <> -1 Then shGUI.Range("H" & pTimers(Index)).Interior.Color = vbWhite
            Exit Sub
        ElseIf hTimers(Index) > 0 Then
            hTimers(Index) = hTimers(Index) - 1
            If pTimers(Index) <> -1 Then shGUI.Range("H" & pTimers(Index)).Interior.Color = vbGreen
        End If
        BeginLoop Index
    Else
        If pTimers(Index) <> -1 Then
            Dim i As Long
            For i = 0 To UBound(pTimers) Step 1
                If rTimers(i) Then
                    shGUI.Range("H" & pTimers(Index)).Interior.Color = vbGreen
                    Exit Sub
                End If
            Next i
            shGUI.Range("H" & pTimers(Index)).Interior.Color = vbWhite
        End If
    End If
End Sub
