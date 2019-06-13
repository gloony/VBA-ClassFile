Option Explicit
Private Declare Function InternetCheckConnection Lib "wininet.dll" Alias "InternetCheckConnectionA" (ByVal lpszUrl As String, ByVal dwFlags As Long, ByVal dwReserved As Long) As Long
Private Declare Function InternetOpen Lib "wininet" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function InternetOpenUrl Lib "wininet" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal lpszUrl As String, ByVal lpszHeaders As String, ByVal dwHeadersLength As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Private Declare Function InternetCloseHandle Lib "wininet" (ByVal hInet As Long) As Integer
Private Declare Function InternetReadFile Lib "wininet" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer

Private Const INTERNET_OPEN_TYPE_PRECONFIG As Long = 0
Private Const INTERNET_OPEN_TYPE_DIRECT    As Long = 1
Private Const INTERNET_OPEN_TYPE_PROXY     As Long = 3
Private Const INTERNET_FLAG_RELOAD         As Long = &H80000000

Private scUserAgent As String
Private ErrorFlag    As Boolean
Private CloseRequest As Boolean
Private Downloading  As Boolean
Private dlSize       As Long

Private Sub Class_Initialize(): scUserAgent = "WorlDesk [WDL V6.0b - wininet.dll]": End Sub

Property Get UserAgent() As Long: UserAgent = scUserAgent: End Property
Property Let UserAgent(ByVal vData As Long): scUserAgent = vData: End Property

Private Function ShellURL(ByVal sUrl As String) As String
    Dim hOpen As Long, hFile As Long, sBuffer As String, ret As Long, sRead As String
    sBuffer = Space(1024)
    hOpen = InternetOpen(scUserAgent, INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
    hFile = InternetOpenUrl(hOpen, sUrl, vbNullString, ByVal 0&, INTERNET_FLAG_RELOAD, ByVal 0&)
    Do
        InternetReadFile hFile, sBuffer, 1024, ret
        sRead = sRead & Left(sBuffer, ret)
        If ret = 0 Then Exit Do
    Loop
    ShellURL = sRead
    InternetCloseHandle hFile
    InternetCloseHandle hOpen
End Function

Public Sub Download(URI As String, SaveAs As String) ' As Boolean
    Dim lNumberOfBytesRead As Long, KbCounter As Long, hOpen As Long, hOpenUrl As Long
    Dim KbCounterGranularity As Integer, GranularityCounter As Integer
    Dim OriginalText As String, sBuffer As String, ret As String
    Dim bDoLoop As Boolean, bRet As Boolean
    Dim sReadBuffer As String * 2048
    ErrorFlag = False
    KbCounterGranularity = 5
    On Error GoTo Oops
    'Downloading = True
    DoEvents
    hOpen = InternetOpen(scUserAgent, INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
    hOpenUrl = InternetOpenUrl(hOpen, URI, vbNullString, 0, INTERNET_FLAG_RELOAD, 0)
    KbCounter = 0
    bDoLoop = True
    GranularityCounter = 0
    While bDoLoop
        sReadBuffer = vbNullString
        bRet = InternetReadFile(hOpenUrl, sReadBuffer, Len(sReadBuffer), lNumberOfBytesRead)
        sBuffer = sBuffer & Left$(sReadBuffer, lNumberOfBytesRead)
        GranularityCounter = GranularityCounter + 1
        KbCounter = KbCounter + CLng((lNumberOfBytesRead / 1024))
        If GranularityCounter = KbCounterGranularity Then
            GranularityCounter = 0
            DoEvents
        End If
        If (Not CBool(lNumberOfBytesRead)) Or CloseRequest Then bDoLoop = False
    Wend
    On Error GoTo 0
    If Not CloseRequest Then
        Open SaveAs For Binary Access Write As #1
        Put #1, , sBuffer
        Close #1
    End If
    If hOpenUrl <> 0 Then InternetCloseHandle (hOpenUrl)
    If hOpen <> 0 Then InternetCloseHandle (hOpen)
    'Downloading = False
    Exit Sub
Oops:
    If hOpenUrl <> 0 Then InternetCloseHandle (hOpenUrl)
    If hOpen <> 0 Then InternetCloseHandle (hOpen)
End Sub
Public Function GetContent(URI As String, Optional Default As String = "") As String
    Dim sRet As String: sRet = ShellURL(URI)
    GetContent = IIf(sRet = vbNullChar, Default, sRet)
End Function
Public Function isConnected() As Boolean
    Dim hInet As Long, hFile As Long
    hInet = InternetOpen(scUserAgent, INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
    hFile = InternetOpenUrl(hInet, wV.URL, vbNullString, ByVal 0&, INTERNET_FLAG_RELOAD, ByVal 0&)
    isConnected = (hFile <> 0)
    If hFile <> 0 Then InternetCloseHandle hFile
    If hInet <> 0 Then InternetCloseHandle hInet
End Function
