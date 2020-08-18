Attribute VB_Name = "modStartMain"
'  http://www.mvps.org/vb
' *********************************************************************
'  You are free to use this code within your own applications, but you
'  are expressly forbidden from selling or otherwise distributing this
'  source code without prior written consent.
' *********************************************************************
Option Explicit

Public Type NOTIFYICONDATA
   cbSize As Long
   hWnd As Long
   uID As Long
   uFlags As Long
   uCallbackMessage As Long
   hIcon As Long
   szTip As String * 64
End Type

Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2

Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4

Public Declare Function ShellNotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Public Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long

Public Const TaskbarCreatedString As String = "TaskbarCreated"

Public Const WM_MOUSEFIRST = &H200
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_MBUTTONDOWN = &H207
Public Const WM_MBUTTONUP = &H208
Public Const WM_MBUTTONDBLCLK = &H209
Public Const WM_MOUSELAST = &H209

Public Const VER_PLATFORM_WIN32s As Long = 0
Public Const VER_PLATFORM_WIN32_WINDOWS As Long = 1
Public Const VER_PLATFORM_WIN32_NT As Long = 2

Public Const WM_USER As Long = &H400
Public Const WM_NULL As Long = &H0

Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long

Public gfrmCommStatus As frmCommStatus
Public gsMainAppExeName As String

Private Property Get msClassName() As String
    msClassName = "modStartMain"
End Property

Public Sub Main()
    On Error GoTo EH
    Dim sMess As String
    
    'If we already have this running then Bail
    If App.PrevInstance Then
        sMess = App.EXEName & " is already running. " & vbCrLf & vbCrLf
        sMess = sMess & "If you can't see " & App.EXEName & " in your system tray(next to your clock), "
        sMess = sMess & "that means Windows had trouble ending it's process.  If this is true, "
        sMess = sMess & "please manually end task on the previous " & App.EXEName & " session using the task manager." & vbCrLf & vbCrLf
        sMess = sMess & "Thank You!"
        MsgBox sMess, vbInformation
        SaveSetting App.EXEName, "MSG", "COMMAND", "SHOW_FTP"
        End
        Exit Sub
    End If
    
    'Set Public Objects Here
    Set goUtil = New V2ECKeyBoard.clsUtil
    
    goUtil.gsAppEXEName = App.EXEName 'Application Name
    goUtil.gsCommandString = Command$
    gsMainAppExeName = GetSetting("ECS", "WEB_SECURITY", "MAIN_APP_EXE_NAME", "EasyClaim")
    goUtil.gsMainAppExeName = gsMainAppExeName
    goUtil.gsCarPrefix = GetSetting(gsMainAppExeName, "COMPANY", "CAR_PREFIX", "V2ECcar")
    'Intialize the DB_PASSWORD Property
    goUtil.INIT_DB_PASSWORD = " r91223-i3q j12223-t1e z02223-n2d q81223-i4x"
    
    SaveSetting goUtil.gsAppEXEName, "COMPANY", "CAR_PREFIX", goUtil.gsCarPrefix
    
    Set gfrmCommStatus = New frmCommStatus
    
    Set goUtil.gfrmCommStatus = gfrmCommStatus
    goUtil.SetUtilObject goUtil

    Load gfrmCommStatus
    
    SaveSetting App.EXEName, "MSG", "COMMAND", "SHOW_FTP"
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub Main"
End Sub

