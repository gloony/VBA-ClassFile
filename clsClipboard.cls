Option Explicit
Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
Private Declare Function GlobalAlloc Lib "Kernel32" (ByVal wFlags&, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "Kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "Kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalSize Lib "Kernel32" (ByVal hMem As Long) As Long
Private Declare Function lstrcpy Lib "Kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long

Public Property Get Value() As String
Dim hClipMemory As Long
Dim lpClipMemory As Long
Dim MyString As String
Dim RetVal As Long
   If OpenClipboard(0&) = 0 Then
      MsgBox "Cannot open Clipboard. Another app. may have it open"
      Exit Property
   End If
On Error GoTo OutOfHere
   hClipMemory = GetClipboardData(1)
   If IsNull(hClipMemory) Then
      MsgBox "Could not allocate memory"
      GoTo OutOfHere
   End If
   lpClipMemory = GlobalLock(hClipMemory)
   If Not IsNull(lpClipMemory) Then
      MyString = Space$(4096)
      RetVal = lstrcpy(MyString, lpClipMemory)
      RetVal = GlobalUnlock(hClipMemory)
      MyString = Mid(MyString, 1, InStr(1, MyString, Chr$(0), 0) - 1)
   Else
      MsgBox "Could not lock memory to copy string from."
   End If
OutOfHere:
   RetVal = CloseClipboard()
   Value = MyString
End Property
Public Property Let Value(ByVal ValueString As String)
Dim MSForms_DataObject As Object: Set MSForms_DataObject = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    MSForms_DataObject.SetText ValueString
    MSForms_DataObject.PutInClipboard
    Set MSForms_DataObject = Nothing
End Property
