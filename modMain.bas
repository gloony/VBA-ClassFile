Option Explicit
Dim cWindow As clsWindow

Public Clip As clsClipboard
Public Computer As clsComputer
Public Explorer As clsExplorer
Public File As clsFile
Public iExplore As clsIExplore
Public iNet As clsINet
Public Joypad As clsJoyPad
Public Keyboard As clsKeyboard
Public Menu As clsMenu
Public Mouse As clsMouse
Public Script As clsScript
Public Str As clsString
Public Timer As clsTimer
Public Var As clsVar
Public Window As clsWindow
Public Zip As clsZip

Public Sub Init()
    If Clip Is Nothing Then Set Clip = New clsClipboard
    If Computer Is Nothing Then Set Computer = New clsComputer
    If Explorer Is Nothing Then Set Explorer = New clsExplorer
    If File Is Nothing Then Set File = New clsFile
    If iExplore Is Nothing Then Set iExplore = New clsIExplore
    If iNet Is Nothing Then Set iNet = New clsINet
    If Joypad Is Nothing Then Set Joypad = New clsJoyPad
    If Keyboard Is Nothing Then Set Keyboard = New clsKeyboard
    If Menu Is Nothing Then Set Menu = New clsMenu
    If Mouse Is Nothing Then Set Mouse = New clsMouse
    If Script Is Nothing Then Set Script = New clsScript
    If Str Is Nothing Then Set Str = New clsString
    If Timer Is Nothing Then Set Timer = New clsTimer
    If Var Is Nothing Then Set Var = New clsVar
    If Window Is Nothing Then Set Window = New clsWindow
    If Zip Is Nothing Then Set Zip = New clsZip
End Sub

Public Sub Reset()
    Set Clip = New clsClipboard
    Set Computer = New clsComputer
    Set Explorer = New clsExplorer
    Set File = New clsFile
    Set Keyboard = New clsKeyboard
    If Menu Is Nothing Then Set Menu = New clsMenu
    Set Mouse = New clsMouse
    Set Script = New clsScript
    Set Str = New clsString
    If Timer Is Nothing Then Set Timer = New clsTimer
    Set Var = New clsVar
    Set Window = New clsWindow
    Set Zip = New clsZip
End Sub
