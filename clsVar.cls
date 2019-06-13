Option Explicit
Private DynWDVar()  As WDVar
Private Type WDVar
    Name As String
    Value As String
End Type

Private Function IsInit() As Boolean
    On Error GoTo Die
    Dim i As Integer
    i = UBound(DynWDVar)
Die:
    If Err.Number = 0 Then IsInit = True
End Function
Private Function GetVar(ByVal StringName As String) As String
    Dim i As Integer
    If Not IsInit Then Exit Function
    For i = 0 To UBound(DynWDVar) Step 1
        If DynWDVar(i).Name = StringName Then
            GetVar = DynWDVar(i).Value
            Exit For
        End If
    Next i
End Function
Private Sub SetVar(ByVal StringName As String, ByVal StringValue As String)
    Dim i As Integer, IsExist As Boolean
    IsExist = False
    If IsInit Then
        For i = 0 To UBound(DynWDVar) Step 1
            If DynWDVar(i).Name = StringName Then
                DynWDVar(i).Value = StringValue
                IsExist = True
                Exit For
            End If
        Next i
        If Not IsExist Then ReDim Preserve DynWDVar(UBound(DynWDVar) + 1)
    Else
        ReDim DynWDVar(0)
    End If
    If Not IsExist Then
        DynWDVar(UBound(DynWDVar)).Name = StringName
        DynWDVar(UBound(DynWDVar)).Value = StringValue
    End If
End Sub

Public Property Get rValue(ByVal Index As String) As String
    Dim i As Long: i = 4
    While shGUI.Range("J" & i).Value <> ""
        Dim j As Long: j = 4
        Dim Multiple As Boolean: Multiple = False
        While shGUI.Range("J" & j).Value <> "" And Not Multiple
            If shGUI.Range("J" & i).Value = shGUI.Range("J" & j).Value Then Multiple = True
            j = j + 1
        Wend
        If Multiple Then
            If shGUI.Range("J" & i).Interior.Color = vbGreen Then
                rValue = shGUI.Range("K" & i).Value
            End If
        Else
            rValue = shGUI.Range("K" & i).Value
        End If
        i = i + 1
    Wend
End Property

Public Property Get Value(ByVal Index As String) As String: Value = GetVar(Index): End Property
Public Property Let Value(ByVal Index As String, ByVal ValueString As String): Call SetVar(Index, ValueString): End Property

Public Property Get lValue(ByVal Index As String) As Long: lValue = CLng(GetVar(Index)): End Property
