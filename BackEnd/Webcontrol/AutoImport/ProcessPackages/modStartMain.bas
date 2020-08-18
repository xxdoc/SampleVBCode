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
    Dim frm As frmProcessPackages
    Dim vAryCommand As Variant
    Dim lCount As Long
    'Package Process params
    Dim sUserName As String
    Dim sPassWord As String
    Dim sAdjUserName As String
    Dim sAssignmentsID As String
    Dim sPackageID As String
    Dim sCarListClassName As String
    Dim sPackageEmailQueueID As String
    'End Package Process params
    Dim sFTPSitePath As String
    Dim sAssignmentsPath As String
    Dim lHwnd As Long
    Dim sBuildAssgnPackPath As String
    Dim sDelMess As String
    Dim bCreateSinglePdfOnly As Boolean
    
    'There can only be one instance of the Package Process
    'This one process will be shelled again when needed.
    If App.PrevInstance Then
        End
        Exit Sub
    End If
    
    sFTPSitePath = GetSetting("V2WebControl", "Dir", "FTPSitePath", vbNullString)
    sAssignmentsPath = Replace(sFTPSitePath, "\Upload\", "\", , , vbTextCompare)
    
    SetUtilObject
    
    'Split up the Command string sent by Autoimport...
    'This Command string will contain The userName Connecting
    If InStr(1, Command$, "RunAsDepOfAutoImport", vbTextCompare) > 0 Then
        vAryCommand = Split(Command$, "|")
        If IsArray(vAryCommand) Then
            sUserName = vAryCommand(1)
            sPassWord = vAryCommand(2)
            sAdjUserName = vAryCommand(3)
            sAssignmentsID = vAryCommand(4)
            sPackageID = vAryCommand(5)
            sCarListClassName = vAryCommand(6)
            sPackageEmailQueueID = vAryCommand(7)
            sBuildAssgnPackPath = sAssignmentsPath & "USER_FOLDERS\" & sAdjUserName
            sBuildAssgnPackPath = sBuildAssgnPackPath & "\PROCESS_PACKAGES\"
            'Delete the batch File
            goUtil.utDeleteFile sBuildAssgnPackPath & sAdjUserName & ".bat"
            sBuildAssgnPackPath = sBuildAssgnPackPath & "BUILD\ASSIGNMENTS\"
            sBuildAssgnPackPath = sBuildAssgnPackPath & sAssignmentsID & "\PACKAGES\"
            sBuildAssgnPackPath = sBuildAssgnPackPath & sPackageID & "\"
            
            'Determine if this Process is a Single PDF File Email Only...
            'that is...
            'This package is being sent outside the normal Schedule to deliver process
            'Which means...
            'a User selected certain Items from a package to be sent to whom ever they
            'wish...
            'Which means the normal updating of Sent Dates the date each package item and
            'the entire Package, whatever is selected to be sent does not apply.
            'So... When the Package Status has [USER SENDING SINGLE PDF FILE EMAIL]
            'Need to Set the mbCreateSinglePdfOnly Flag
            If sPackageEmailQueueID <> vbNullString Then
                bCreateSinglePdfOnly = True
            End If
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
        ' Go ahead and load.
        Set frm = New frmProcessPackages
        'Set the username
        frm.UserName = sUserName
        frm.PassWord = sPassWord
        frm.AdjUserName = sAdjUserName
        frm.AssignmentsID = sAssignmentsID
        frm.PackageID = sPackageID
        frm.CarListClassName = sCarListClassName
        frm.CreateSinglePdfOnly = bCreateSinglePdfOnly
        frm.PackageEmailQueueID = sPackageEmailQueueID
        frm.BuildAssgnPackPath = sBuildAssgnPackPath
        Load frm
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




