Option Explicit
Private Declare Sub RtlMoveMemory Lib "Kernel32" (pDst As Any, pSrc As Any, ByVal dlen As Long)
Private Declare Function AdjustTokenPrivileges Lib "advapi32" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetCurrentProcess Lib "Kernel32" () As Long
Private Declare Function GetModuleHandle Lib "kernel32.dll" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetSystemPowerStatus Lib "Kernel32" (lpSystemPowerStatus As SYSTEM_POWER_STATUS) As Long
Private Declare Function GetVersionExA Lib "Kernel32" (LpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function GetWindowsDirectory Lib "Kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetTickCountK32 Lib "Kernel32" Alias "GetTickCount" () As Long
Private Declare Function GetVersionEx Lib "Kernel32" Alias "GetVersionExA" (LpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function GetLastInputInfo Lib "user32.dll" (inputStructure As inputInfo) As Boolean
Private Declare Function IIDFromString Lib "ole32" (ByVal lpszIID As Long, iid As Any) As Long
Private Declare Function IsUserAnAdmin Lib "Shell32" Alias "#680" () As Integer
Private Declare Function IsWow64Process Lib "Kernel32" (ByVal hProc As Long, ByRef bWow64Process As Boolean) As Long
Private Declare Function LockWorkStation Lib "user32.dll" () As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
Private Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function SendMessageA Lib "user32" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hWndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function CoCreateInstance Lib "ole32" (rclsid As Any, ByVal pUnkOuter As Long, ByVal dwClsContext As Long, riid As Any, ByVal ppv As Long) As Long
Private Declare Function CallWindowProcA Lib "user32" (ByVal addr As Long, ByVal p1 As Long, ByVal p2 As Long, ByVal p3 As Long, ByVal p4 As Long) As Long
Private Declare Function SystemParametersInfoA Lib "user32" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Sub SleepA Lib "Kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)

Public Enum AD_APPLY
    AD_APPLY_SAVE = &H1
    AD_APPLY_HTMLGEN = &H2
    AD_APPLY_REFRESH = &H4
    AD_APPLY_ALL = &H7
    AD_APPLY_FORCE = &H8
    AD_APPLY_BUFFERED_REFRESH = &H10
    AD_APPLY_DYNAMICREFRESH = &H20
End Enum

Private Type GUID
    Data1                       As Long
    Data2                       As Integer
    Data3                       As Integer
    Data4(7)                    As Byte
End Type
Private Type LUID
    UsedPart                    As Long
    IgnoredForNowHigh32BitPart  As Long
End Type
Private Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    TheLuid As LUID
    Attributes As Long
End Type
Private Type OSVERSIONINFO
    OSVSize         As Long
    dwVerMajor      As Long
    dwVerMinor      As Long
    dwBuildNumber   As Long
    PlatformID      As Long
    szCSDVersion    As String * 128
    dwOSVersionInfoSize As String * 128
    dwPlatformId As Long
End Type
Private Type inputInfo
    structSize As Long
    tickCount As Long
End Type
Private Type IActiveDesktop
    QueryInterface              As Long
    AddRef                      As Long
    Release                     As Long
    ApplyChanges                As Long
    GetWallpaper                As Long
    SetWallpaper                As Long
    GetWallpaperOptions         As Long
    SetWallpaperOptions         As Long
    GetPattern                  As Long
    SetPattern                  As Long
    GetDesktopItemOptions       As Long
    SetDesktopItemOptions       As Long
    AddDesktopItem              As Long
    AddDesktopItemWithUI        As Long
    ModifyDesktopItem           As Long
    RemoveDesktopItem           As Long
    GetDesktopItemCount         As Long
    GetDesktopItem              As Long
    GetDesktopItemByID          As Long
    GenerateDesktopItemHtml     As Long
    AddUrl                      As Long
    GetDesktopItemBySource      As Long
End Type
Private Type SHITEMID
    cb                            As Long
    abID                          As Byte
End Type
Private Type ITEMIDLIST
    mkID As SHITEMID
End Type
Private Type SYSTEM_POWER_STATUS
    ACLineStatus As Byte
    BatteryFlag As Byte
    BatteryLifePercent As Byte
    Reserved1 As Byte
    BatteryLifeTime As Long
    BatteryFullLifeTime As Long
End Type
Private bStatus As SYSTEM_POWER_STATUS

Private Sub AddByte(pASM As Long, bt As Byte)
    RtlMoveMemory ByVal pASM, bt, 1
    pASM = pASM + 1
End Sub
Private Sub AddCall(pASM As Long, addr As Long)
    AddByte pASM, &HE8
    AddLong pASM, addr - pASM - 4
End Sub
Private Sub AddLong(pASM As Long, lng As Long)
    RtlMoveMemory ByVal pASM, lng, 4
    pASM = pASM + 4
End Sub
Private Sub AddPush(pASM As Long, lng As Long)
    AddByte pASM, &H68
    AddLong pASM, lng
End Sub
Private Sub AdjustToken()
    Dim hdlProcessHandle As Long
    Dim hdlTokenHandle As Long
    Dim tmpLuid As LUID
    Dim tkp As TOKEN_PRIVILEGES
    Dim tkpNewButIgnored As TOKEN_PRIVILEGES
    Dim lBufferNeeded    As Long
    hdlProcessHandle = GetCurrentProcess()
    OpenProcessToken hdlProcessHandle, (&H20 Or &H8), hdlTokenHandle
    LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
    tkp.PrivilegeCount = 1
    tkp.TheLuid = tmpLuid
    tkp.Attributes = &H2
    AdjustTokenPrivileges hdlTokenHandle, False, tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
End Sub
Private Sub ChangeActiveDesktopWallpaper(ByVal strFile As String)
    Const CLSID_ActiveDesktop  As String = "{75048700-EF1F-11D0-9888-006097DEACF9}"
    Const IID_ActiveDesktop    As String = "{F490EB00-1240-11D1-9888-006097DEACF9}"
    Dim vtbl    As IActiveDesktop
    Dim vtblptr As Long
    Dim classid As GUID
    Dim iid     As GUID
    Dim obj     As Long
    Dim hRes    As Long
    hRes = IIDFromString(StrPtr(CLSID_ActiveDesktop), classid)
    If hRes <> 0 Then Exit Sub
    hRes = IIDFromString(StrPtr(IID_ActiveDesktop), iid)
    If hRes <> 0 Then Exit Sub
    hRes = CoCreateInstance(classid, 0, 1&, iid, VarPtr(obj))
    If hRes <> 0 Then Exit Sub
    RtlMoveMemory vtblptr, ByVal obj, 4
    RtlMoveMemory vtbl, ByVal vtblptr, Len(vtbl)
    hRes = CallPointer(vtbl.SetWallpaper, obj, StrPtr(strFile), 0)
    hRes = CallPointer(vtbl.ApplyChanges, obj, AD_APPLY_ALL Or AD_APPLY_FORCE)
    CallPointer vtbl.Release, obj
End Sub
Private Sub ChangeDekstopWallpaper(ByVal PathFile As String)
    If SystemParametersInfoA(20, 0, PathFile, &H2 Or &H1) = 0 Then Call ChangeActiveDesktopWallpaper(PathFile)
End Sub

Private Function CallPointer(ByVal fnc As Long, ParamArray params()) As Long
    Dim btASM(&HEC00& - 1) As Byte
    Dim pASM               As Long
    Dim i                  As Integer
    pASM = VarPtr(btASM(0))
    AddByte pASM, &H58
    AddByte pASM, &H59
    AddByte pASM, &H59
    AddByte pASM, &H59
    AddByte pASM, &H59
    AddByte pASM, &H50
    For i = UBound(params) To 0 Step -1
        AddPush pASM, CLng(params(i))
    Next
    AddCall pASM, fnc
    AddByte pASM, &HC3
    CallPointer = CallWindowProcA(VarPtr(btASM(0)), 0, 0, 0, 0)
End Function

Public Property Get Reg(ByVal KeyString As String) As String
    On Error GoTo Die
    Dim RegObj, RegKey As String
    Set RegObj = CreateObject("WScript.Shell")
    RegKey = RegObj.RegRead(KeyString)
    Reg = RegKey
Die:
    Set RegObj = Nothing
End Property
Public Property Let Reg(ByVal KeyString As String, ByVal ValueString As String)
    On Error GoTo Error
    Dim RegObj
    Set RegObj = CreateObject("WScript.Shell")
    If ValueString = vbNullString Then Call RegObj.Delete(KeyString) _
    Else: Call RegObj.RegWrite(KeyString, ValueString)
Error:
    Set RegObj = Nothing
End Property
Public Property Get Settings(ByVal Appname As String, ByVal Section As String, ByVal Key As String) As String
    Settings = Interaction.GetSetting(Appname, Section, Key, "")
End Property
Public Property Let Settings(ByVal Appname As String, ByVal Section As String, ByVal Key As String, ByVal Settings As String)
    Interaction.SaveSetting Appname, Section, Key, Settings
End Property

Public Sub AppActivate(ByVal TitleString As String): Call Interaction.AppActivate(TitleString): End Sub
Public Sub CMD(ByVal CommandString As String): Call Shell(CommandString, vbNormalFocus): End Sub
Public Sub Hibernate(): Call Shell("rundll32.exe powrprof.dll,SetSuspendState 1,1,0"): End Sub
Public Sub LockScreen(): LockWorkStation: End Sub
Public Sub LogOff(): AdjustToken: Call ExitWindowsEx(0, 0): End Sub
Public Sub Reboot(): AdjustToken: Call ExitWindowsEx(2, 0): End Sub
Public Sub ScreenSaver(): SendMessageA 0, &H112&, &HF140&, 0&: End Sub
Public Sub ShutDown(): AdjustToken: Call ExitWindowsEx(1, 0): End Sub
Public Sub Taskbar(ByVal Visible As Boolean)
    If Visible Then Call SetWindowPos(FindWindow("Shell_traywnd", ""), 0, 0, 0, 0, 0, &H80) _
    Else Call SetWindowPos(FindWindow("Shell_traywnd", ""), 0, 0, 0, 0, 0, &H40)
End Sub
Public Sub TaskKill(ByVal ProcessString As String, Optional ByVal ForceExe As Boolean = True)
    If ForceExe Then If Right(ProcessString, Len(".exe")) <> ".exe" Then ProcessString = ProcessString & ".exe"
    On Error GoTo Die
    Dim objWMIService, colProcessList, objProcess As Object
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set colProcessList = objWMIService.ExecQuery("Select * from Win32_Process Where Name = '" & ProcessString & "'")
    For Each objProcess In colProcessList
        objProcess.Terminate
    Next
    Set colProcessList = Nothing
    Set objWMIService = Nothing
    Exit Sub
Die:
    Shell "taskkill /im " & ProcessString
    Set colProcessList = Nothing
    Set objWMIService = Nothing
End Sub
Public Sub WaitForUserAction()
    Dim firstTick As Long, LastTick As Long
    Dim info As inputInfo
    info.structSize = Len(info)
    GetLastInputInfo info
    firstTick = info.tickCount
    LastTick = info.tickCount
    Do While firstTick = LastTick
        info.structSize = Len(info)
        GetLastInputInfo info
        LastTick = info.tickCount
    Loop
End Sub
Public Sub Wallpaper(ByVal PathString As String): Call ChangeDekstopWallpaper(PathString): End Sub

Public Function ACLineStatus() As Boolean: Call GetSystemPowerStatus(bStatus): ACLineStatus = bStatus.ACLineStatus: End Function
Public Function BatteryLife() As Integer: Call GetSystemPowerStatus(bStatus): BatteryLife = bStatus.BatteryLifePercent: End Function
Public Function Environ(ByVal Expression As String) As String: Environ = Interaction.Environ(Expression): End Function
Public Function GetSpecialFolder(ByVal FolderToReturn As String) As String
    On Error Resume Next
    Dim ReturnCode As Long
    Dim ReturnPath As String
    Dim ID_List    As ITEMIDLIST
    ReturnCode = -1
    Select Case LCase(FolderToReturn)
    Case "bitbucket": ReturnCode = SHGetSpecialFolderLocation(0, 10, ID_List)
    Case "controls": ReturnCode = SHGetSpecialFolderLocation(0, 3, ID_List)
    Case "desktop": ReturnCode = SHGetSpecialFolderLocation(0, 0, ID_List)
    Case "desktopdirectory": ReturnCode = SHGetSpecialFolderLocation(0, 16, ID_List)
    Case "drives": ReturnCode = SHGetSpecialFolderLocation(0, 17, ID_List)
    Case "fonts": ReturnCode = SHGetSpecialFolderLocation(0, 20, ID_List)
    Case "nethood": ReturnCode = SHGetSpecialFolderLocation(0, 19, ID_List)
    Case "network": ReturnCode = SHGetSpecialFolderLocation(0, 18, ID_List)
    Case "personal": ReturnCode = SHGetSpecialFolderLocation(0, 5, ID_List)
    Case "printers": ReturnCode = SHGetSpecialFolderLocation(0, 4, ID_List)
    Case "programs": ReturnCode = SHGetSpecialFolderLocation(0, 2, ID_List)
    Case "recent": ReturnCode = SHGetSpecialFolderLocation(0, 8, ID_List)
    Case "sendto": ReturnCode = SHGetSpecialFolderLocation(0, 9, ID_List)
    Case "startmenu": ReturnCode = SHGetSpecialFolderLocation(0, 11, ID_List)
    Case "startup": ReturnCode = SHGetSpecialFolderLocation(0, 7, ID_List)
    Case "templates": ReturnCode = SHGetSpecialFolderLocation(0, 21, ID_List)
    Case "altstartup": ReturnCode = SHGetSpecialFolderLocation(0, 29, ID_List)
    Case "appdata": ReturnCode = SHGetSpecialFolderLocation(0, 26, ID_List)
    Case "common_altstartup": ReturnCode = SHGetSpecialFolderLocation(0, 30, ID_List)
    Case "common_desktopdirectory": ReturnCode = SHGetSpecialFolderLocation(0, 25, ID_List)
    Case "common_favorites": ReturnCode = SHGetSpecialFolderLocation(0, 31, ID_List)
    Case "common_programs": ReturnCode = SHGetSpecialFolderLocation(0, 23, ID_List)
    Case "common_startmenu": ReturnCode = SHGetSpecialFolderLocation(0, 22, ID_List)
    Case "common_startup": ReturnCode = SHGetSpecialFolderLocation(0, 24, ID_List)
    Case "cookies": ReturnCode = SHGetSpecialFolderLocation(0, 33, ID_List)
    Case "favorites": ReturnCode = SHGetSpecialFolderLocation(0, 6, ID_List)
    Case "history": ReturnCode = SHGetSpecialFolderLocation(0, 34, ID_List)
    Case "internet": ReturnCode = SHGetSpecialFolderLocation(0, 1, ID_List)
    Case "internet_cache": ReturnCode = SHGetSpecialFolderLocation(0, 32, ID_List)
    Case "printhood": ReturnCode = SHGetSpecialFolderLocation(0, 27, ID_List)
    Case "admintools": ReturnCode = SHGetSpecialFolderLocation(0, &H30, ID_List)
    Case "common_admintools": ReturnCode = SHGetSpecialFolderLocation(0, &H2F, ID_List)
    Case "mypictures": ReturnCode = SHGetSpecialFolderLocation(0, &H27, ID_List)
    Case "program_files": ReturnCode = SHGetSpecialFolderLocation(0, &H26, ID_List)
    Case "program_files_commonx86": ReturnCode = SHGetSpecialFolderLocation(0, &H2C, ID_List)
    Case "program_files_common": ReturnCode = SHGetSpecialFolderLocation(0, &H2B, ID_List)
    Case "system": ReturnCode = SHGetSpecialFolderLocation(0, &H25, ID_List)
    Case "systemx86": ReturnCode = SHGetSpecialFolderLocation(0, &H29, ID_List)
    Case "windows": ReturnCode = SHGetSpecialFolderLocation(0, &H24, ID_List)
    End Select
    If ReturnCode = 0 Then
        ReturnPath = Space(512)
        ReturnCode = SHGetPathFromIDList(ByVal ID_List.mkID.cb, ByVal ReturnPath)
        GetSpecialFolder = Left(ReturnPath, InStr(ReturnPath, Chr(0)) - 1)
        Exit Function
    Else
        GetSpecialFolder = ""
    End If
End Function
Public Function GetTickCount() As Long: GetTickCount = GetTickCountK32: End Function
Public Function GetSystemDrive() As String
    GetSystemDrive = Space(1000)
    Call GetWindowsDirectory(GetSystemDrive, Len(GetSystemDrive))
    GetSystemDrive = Left$(GetSystemDrive, 3)
End Function
Public Function GetWinVer()
    Dim osv As OSVERSIONINFO
    osv.OSVSize = Len(osv)
    GetVersionEx osv
    GetWinVer = osv.dwPlatformId & "." & osv.dwVerMajor & "." & osv.dwVerMinor
End Function
Public Function LastInput() As Long
    Dim info As inputInfo
    info.structSize = Len(info)
    GetLastInputInfo info
    LastInput = info.tickCount
End Function
Public Function Is64Bit() As Boolean
    Dim Handle  As Long
    Dim IsHost64Bit As Boolean
    IsHost64Bit = False
    Handle = GetProcAddress(GetModuleHandle("kernel32"), "IsWow64Process")
    If Handle <> 0 Then IsWow64Process GetCurrentProcess(), IsHost64Bit
    Is64Bit = IsHost64Bit
End Function
Public Function IsAdmin() As Boolean: IsAdmin = IsUserAnAdmin: End Function
Public Function MacID() As String
    Dim Computer As String
    Dim wmi      As Variant
    Dim query    As Variant
    Dim Mac      As Variant
    Dim mac_ids  As String
    Computer = "."
    Set wmi = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & Computer & "\root\cimv2")
    Set query = wmi.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration where IPEnabled = true")
    For Each Mac In query
        mac_ids = mac_ids & ", " & Mac.MacAddress
    Next Mac
    If Len(mac_ids) > 0 Then mac_ids = Mid$(mac_ids, 3)
    MacID = mac_ids
End Function
Public Function Msg(ByVal Prompt As String, Optional ByVal Style As Integer = vbOKOnly, Optional ByVal Title As String = "") As Integer
    If Title = "" Then Title = ThisWorkbook.Name
    Msg = MsgBox(Prompt, Style, Title)
End Function
Public Function Prompt(ByVal PromptString As String, Optional ByVal Default As String = "", Optional ByVal Title As String = "") As String
    If Title = "" Then Title = ThisWorkbook.Name
    Prompt = InputBox(PromptString, Default, Title)
End Function
Public Sub Sleep(ByVal dwMilliseconds As Long)
    SleepA dwMilliseconds
End Sub
