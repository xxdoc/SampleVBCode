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

Public Type OSVERSIONINFO
   dwOSVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long
   szCSDVersion As String * 128      '  Maintenance string for PSS usage
End Type

Public Const VER_PLATFORM_WIN32s As Long = 0
Public Const VER_PLATFORM_WIN32_WINDOWS As Long = 1
Public Const VER_PLATFORM_WIN32_NT As Long = 2

Public Const WM_USER As Long = &H400
Public Const WM_NULL As Long = &H0

Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Sub Main()
    On Error GoTo EH
    Dim sMess As String
    Dim os As OSVERSIONINFO
    Dim frm As frmProcessTokin
    Dim vAryCommand As Variant
    Dim lCount As Long
    Dim sUserName As String
    Dim lHwnd As Long
    Dim sUserFolderPath As String
    Dim sDelMess As String
    
    'Split up the Command string sent by Autoimport...
    'This Command string will contain The userName Connecting
    If InStr(1, Command$, "RunAsDepOfAutoImport", vbTextCompare) > 0 Then
        vAryCommand = Split(Command$, "|")
        If IsArray(vAryCommand) Then
            sUserName = vAryCommand(1)
            sUserFolderPath = vAryCommand(2)
        Else
            Exit Sub
        End If
        GoTo RUN_SERVICE
    Else
        MsgBox App.EXEName & " may only be run as a dependant of Auto Import.", vbExclamation
        End
        Exit Sub
    End If
RUN_SERVICE:
    ' Until the new shell arrives, this program
    ' only runs in Win95 (not in NT 3.5x).
    os.dwOSVersionInfoSize = Len(os)
    Call GetVersionEx(os)
    If os.dwMajorVersion >= 4 Then
        SetUtilObject
        'Delete the batch file that created this process
        sDelMess = goUtil.utDeleteFile(sUserFolderPath & sUserName & ".bat")
                
        'See if there already is a Process for this username...
        lHwnd = goUtil.utFindWindowPartial("Process User Tokin (" & sUserName & ")", FwpContains, False, False)
        If lHwnd = 0 Then
            ' Go ahead and load.
            Set frm = New frmProcessTokin
            'Set the username
            frm.UserName = sUserName
            frm.UserFolderPath = sUserFolderPath
            Load frm
        Else
            'Need to clean up and exit
            goUtil.CLEANUP
            Exit Sub
        End If
    Else
        MsgBox "This program only runs under NewShell", vbCritical, "Program Ending"
    End If
    Exit Sub
EH:
    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
    sMess = sMess & "Public Sub Main" & vbCrLf
    sMess = sMess & "ERROR # " & Err.Number & " " & Now & vbCrLf
    sMess = sMess & Err.Description & vbCrLf
    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
    ErrorLog sMess
End Sub




