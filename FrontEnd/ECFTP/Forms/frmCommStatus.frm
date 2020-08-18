VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{307C5043-76B3-11CE-BF00-0080AD0EF894}#1.0#0"; "MsgHoo32.OCX"
Object = "{B71A484A-57D1-11D2-821F-000086075197}#1.0#0"; "FTPX.OCX"
Begin VB.Form frmCommStatus 
   AutoRedraw      =   -1  'True
   Caption         =   "Communications Status"
   ClientHeight    =   7140
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9480
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCommStatus.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7140
   ScaleWidth      =   9480
   Begin VB.CommandButton cmdCancelPhotoAttachUL 
      Caption         =   "&Finish Later"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7800
      Picture         =   "frmCommStatus.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   5900
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "&Connect"
      Height          =   375
      Left            =   7800
      TabIndex        =   22
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton cmdViewHistory 
      Caption         =   "&View History"
      Height          =   360
      Left            =   120
      TabIndex        =   0
      Top             =   40
      Width           =   1400
   End
   Begin VB.Frame framProcess 
      Height          =   4845
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   9255
      Begin FtpXCtl.FtpXCtl EC_FTP 
         Left            =   8520
         Top             =   600
         Blocking        =   -1  'True
         DebugMode       =   1
         Directory       =   ""
         DstFilename     =   ""
         Host            =   ""
         LogonPassword   =   ""
         Pattern         =   ""
         SrcFilename     =   ""
         Type            =   1
         LogonName       =   ""
         Account         =   ""
         Timeout         =   0
         Port            =   21
         DisablePasv     =   0   'False
         DirItemPattern  =   ""
         LibraryName     =   "WSOCK32.DLL"
         BlockingMode    =   0
         FirewallType    =   0
         FirewallHost    =   ""
         FirewallPort    =   0
         FirewallLogonName=   ""
         FirewallPassword=   ""
      End
      Begin VB.Timer Timer_Resize 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   7440
         Top             =   600
      End
      Begin VB.Timer TimerMsg 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   7920
         Top             =   600
      End
      Begin MsghookLib.Msghook Msghook 
         Left            =   8760
         Top             =   1200
         _Version        =   65536
         _ExtentX        =   847
         _ExtentY        =   847
         _StockProps     =   0
      End
      Begin MSComctlLib.ImageList imgList 
         Left            =   8160
         Top             =   1200
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         UseMaskColor    =   0   'False
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCommStatus.frx":0884
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCommStatus.frx":0CD6
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCommStatus.frx":1128
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.ListBox lstProcess 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         CausesValidation=   0   'False
         Height          =   4110
         ItemData        =   "frmCommStatus.frx":157A
         Left            =   120
         List            =   "frmCommStatus.frx":157C
         TabIndex        =   3
         Top             =   600
         Width           =   9015
      End
      Begin VB.Image imgEncrypt 
         Height          =   480
         Left            =   8640
         Picture         =   "frmCommStatus.frx":157E
         ToolTipText     =   "Encrypted Data"
         Top             =   120
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label lblProcess 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   9015
      End
   End
   Begin VB.Frame framCommands 
      Height          =   1815
      Left            =   120
      TabIndex        =   4
      Top             =   5160
      Width           =   9255
      Begin VB.CommandButton cmdFireWall 
         Caption         =   "__"
         Height          =   375
         Left            =   7080
         Picture         =   "frmCommStatus.frx":19C0
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Fire Wall Settings (Optional)"
         Top             =   740
         Width           =   495
      End
      Begin VB.CheckBox chkForceDownload 
         Caption         =   "Overwrite Photo Attachments"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5640
         TabIndex        =   19
         ToolTipText     =   "Download and overwrite  Photos and Attachments even if they already exisit on your computer, when restoring files."
         Top             =   1200
         Width           =   1575
      End
      Begin VB.CheckBox chkLogAllStatus 
         Caption         =   "Log ALL Status Details"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5640
         TabIndex        =   18
         Top             =   720
         Width           =   1575
      End
      Begin VB.CommandButton cmdPassiveInfo 
         Caption         =   "<<"
         Height          =   375
         Left            =   7080
         TabIndex        =   20
         Top             =   240
         Width           =   495
      End
      Begin VB.CheckBox chkUsePasiveConnection 
         Caption         =   "Use Passive Connection"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5640
         TabIndex        =   17
         Top             =   190
         Width           =   1575
      End
      Begin MSComctlLib.ProgressBar PBDownLoad 
         Height          =   255
         Left            =   2500
         TabIndex        =   6
         Top             =   240
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComctlLib.ProgressBar PBUpLoad 
         Height          =   255
         Left            =   2500
         TabIndex        =   8
         Top             =   480
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComctlLib.ProgressBar PBPhotoDownLoad 
         Height          =   255
         Left            =   2500
         TabIndex        =   10
         Top             =   720
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComctlLib.ProgressBar PBPhoto 
         Height          =   255
         Left            =   2500
         TabIndex        =   12
         Top             =   960
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComctlLib.ProgressBar PBAttachDownload 
         Height          =   255
         Left            =   2500
         TabIndex        =   14
         Top             =   1200
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComctlLib.ProgressBar PBAttach 
         Height          =   255
         Left            =   2500
         TabIndex        =   16
         Top             =   1440
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.CommandButton cmdHide 
         Cancel          =   -1  'True
         Caption         =   "&Hide"
         Height          =   375
         Left            =   7680
         TabIndex        =   24
         Top             =   1350
         Width           =   1455
      End
      Begin VB.Label lblDownloadAttachment 
         Caption         =   "Attachment Download"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1200
         Width           =   2300
      End
      Begin VB.Label lblDownloadPhoto 
         Caption         =   "Photo Download"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   2300
      End
      Begin VB.Label lblUploadAttachment 
         Caption         =   "Attachment Upload"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1440
         Width           =   2300
      End
      Begin VB.Label lblUploadPhoto 
         Caption         =   "Photo Upload"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   2300
      End
      Begin VB.Label lblUpLoadFiles 
         Caption         =   "File Upload"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   2300
      End
      Begin VB.Label lblDownLoad 
         Caption         =   "File Download"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   2300
      End
   End
   Begin VB.Label lblMess 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1680
      TabIndex        =   26
      Top             =   90
      Width           =   3495
   End
   Begin VB.Label lblTitleBytesTransferred 
      Caption         =   "Bytes Transferred:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   27
      Top             =   90
      Width           =   1815
   End
   Begin VB.Label lblBytesTransferred 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7200
      TabIndex        =   25
      Top             =   90
      Width           =   1815
   End
   Begin VB.Menu mPopUp 
      Caption         =   "Pop"
      Visible         =   0   'False
      Begin VB.Menu mPop 
         Caption         =   "&Show"
         Index           =   0
      End
      Begin VB.Menu mPop 
         Caption         =   "&Hide"
         Index           =   1
      End
      Begin VB.Menu mPop 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mPop 
         Caption         =   "Connect"
         Index           =   3
      End
      Begin VB.Menu mPop 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mPop 
         Caption         =   "&Exit"
         Index           =   5
      End
   End
End
Attribute VB_Name = "frmCommStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Enum PicList
    FTP01 = 1
    FTP02
    FTP03
End Enum

Public Enum MenuList
    Show = 0
    Hide
    BarConnect
    Connect
    BarExit
    ExitApp
End Enum

Private Const SD = vbNullChar

' User defined constant values
Private Const cbNotify As Long = &H4000
Private Const uID As Long = 61860

'Control Size in Ref to Form Diff Constants
Private Const FORM_W As Long = 9600
Private Const FORM_H As Long = 7545
'framProcess
Private Const framProcess_W As Long = 345
Private Const framProcess_H As Long = 2700
Private Const lblProcess_W As Long = 585
Private Const imgEncrypt_L As Long = 960
Private Const lstProcess_W As Long = 585
Private Const lstProcess_H As Long = 3405

'framCommands
Private Const framCommands_W As Long = 345
Private Const framCommands_T As Long = 2385
Private Const cmdConnect_T As Long = 2145
Private Const cmdConnect_L As Long = 1800
Private Const cmdCancelPhotoAttachUL_T = 1665
Private Const cmdCancelPhotoAttachUL_L = 1800
Private Const cmdHide_L As Long = 1920

'user Folder
Private Const USER_FOLDERS As String = " a71223-t5k l81223-i4c w61223-A6q p70223-s51s e21223-s01k a60223-s61q g11223-m11u y80223-241b e90223-V31t b01223-n21h q02223-E2v d41223-g8z y12223-C1r i91223-n3h c31223-_9x i51223-e7p" 'ECV2_Assignments
Private Const SP_FOLDERS As String = " e71223-S5c l02223-22b m51223-E7j f91223- 3y u81223-P4j p41223-C8r l12223-V1a y61223-_6i" 'ECV2_SP

' Member variables
Private m_NID As NOTIFYICONDATA
Private m_TaskbarCreated As Long
Private mdicRemoteFiles As Scripting.Dictionary
Private msCurrentFile As String
'timed Falg
Private mbShutDownFTP As Boolean
Private mbResize As Boolean
Private mbUpdateDB As Boolean
Private msIBSuffix As String 'BGS 11.20.2001 used to store IB number
Private mbCancelPhotoAttach As Boolean
Private mbConnected As Boolean
Private mbSingleFileProcess As Boolean
Private msCar As String
Private msCat As String
Private mbLoadingRegForm As Boolean
Private msLastTime As String
Private msFTPDLPath As String 'Download Path
Private msFTPULPath As String 'Upload Path
Private msFTPLogPath As String 'FTP Log Path History
Private msAttachReposPath As String
Private msPhotoReposPath As String
Private msSPPath As String 'Software Package path
Private mbDatabaseUpgrade As Boolean
Private msDBSPName As String
Private msMainUtilSPName As String
Private msMainARVSPName As String
Private msMainEXESPName As String
Private msMainFTPEXESPName As String
Private msUserName As String
Private msUsersID As String
Private mlProgressBytesTransferred As Long
Private msTokName As String
Private msTokPath As String
Private msUserFolders As String
Private msSPFolders As String

Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Property Get msClassName() As String
    msClassName = Me.Name
End Property

Public Property Let Cat(psName As String)
    msCat = psName
End Property
Public Property Get Cat() As String
    Cat = msCat
End Property

Public Property Let Car(psName As String)
    msCar = psName
End Property
Public Property Get Car() As String
    Car = msCar
End Property

Public Property Let FlagShutDownFTP(pbFlag As Boolean)
    mbShutDownFTP = pbFlag
End Property
Public Property Get FlagShutDownFTP() As Boolean
    FlagShutDownFTP = mbShutDownFTP
End Property

Private Sub chkForceDownload_Click()
    On Error GoTo EH
    Dim bForceDownload As Boolean
    
    If chkForceDownload.Value = vbChecked Then
        bForceDownload = True
    Else
        bForceDownload = False
    End If
    
    SaveSetting App.EXEName, "CONNECTION_SETTINGS", "ForceDownload", bForceDownload
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub chkForceDownload_Click"
End Sub

Private Sub chkLogAllStatus_Click()
    On Error GoTo EH
    Dim bLogAllStatus As Boolean
    Dim sMess As String
    
    If chkLogAllStatus.Value = vbChecked Then
        sMess = "Use ""Log ALL Status Details"" to debug connection issues ONLY!" & vbCrLf & vbCrLf
        sMess = sMess & "Click ""YES"" if you are sure you want to Log all status." & vbCrLf
        sMess = sMess & "Click ""NO"" to uncheck this option." & vbCrLf
        If MsgBox(sMess, vbExclamation + vbYesNo) = vbYes Then
            bLogAllStatus = True
        Else
            chkLogAllStatus.Value = vbUnchecked
            Exit Sub
        End If
    Else
        bLogAllStatus = False
    End If
    
    SaveSetting App.EXEName, "CONNECTION_SETTINGS", "LOG_ALL_STATUS", bLogAllStatus
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub chkLogAllStatus_Click"
End Sub

Private Sub chkUsePasiveConnection_Click()
    On Error GoTo EH
    Dim bDisablePasv As Boolean
    
    If chkUsePasiveConnection.Value = vbChecked Then
        bDisablePasv = False
        chkUsePasiveConnection.Caption = "Use Passive Connection"
    Else
        bDisablePasv = True
        chkUsePasiveConnection.Caption = "Use Active Connection"
    End If
    
    SaveSetting App.EXEName, "CONNECTION_SETTINGS", "DisablePasv", bDisablePasv
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub chkUsePasiveConnection_Click"
End Sub

Private Sub cmdCancelPhotoAttachUL_Click()
    mbCancelPhotoAttach = True
End Sub

Private Sub cmdConnect_Click()
    mbUpdateDB = True
    EnableCommandFrame False
    cmdConnect.Enabled = False
    cmdViewHistory.Enabled = False
    Start_Comm
End Sub

Private Sub cmdFireWall_Click()
    On Error Resume Next
    Load frmFireWall
    frmFireWall.Show vbModal
End Sub

'Private Sub cmdDelHistory_Click()
'    On Error GoTo EH
'    Dim lRet As VbMsgBoxResult
'    If cboHistory.Text <> vbNullString Then
'        lRet = MsgBox("Are you sure you want to delete " & cboHistory.Text & " ?", vbQuestion + vbYesNo, "Delete History File")
'        If lRet = vbYes Then
'            goUtil.utDeleteFile goUtil.gsInstallDir & "\" & cboHistory.Text
'            LoadHistory
'        End If
'    End If
'
'    Exit Sub
'EH:
'   goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdDelHistory_Click"
'End Sub

Private Sub cmdHide_Click()
    On Error Resume Next
    HideMe
End Sub

Private Sub cmdPassiveInfo_Click()
    On Error GoTo EH
    Dim sMess As String
    
    'Give Information about the Passive version Active Connections...
    sMess = "FTP protocol uses two sockets when performing transfers " & vbCrLf
    sMess = sMess & "-- one for sending commands (called the control socket) " & vbCrLf
    sMess = sMess & "and one for transferring the data (communication socket). " & vbCrLf & vbCrLf
    
    sMess = sMess & "The PORT command specifies that the ""client"" will determine which port to use " & vbCrLf
    sMess = sMess & "for the communication socket. This is called an ""active"" connection." & vbCrLf
    sMess = sMess & "Uncheck ""Use Passive Connection"", and then re-connect." & vbCrLf & vbCrLf
    
    sMess = sMess & "The PASV command specifies that the ""server"" will determine which port to use " & vbCrLf
    sMess = sMess & "for the communication. This is called a ""passive"" connection. " & vbCrLf
    sMess = sMess & "Check ""Use Passive Connection"", and then re-connect." & vbCrLf & vbCrLf
    
    sMess = sMess & "***Note***" & vbCrLf
    sMess = sMess & "Although the FTP RFC allows both modes, some servers or firewalls may only allow passive connections " & vbCrLf
    sMess = sMess & "while others only allow active connections. Since there is no way to know ahead of time, you should first " & vbCrLf
    sMess = sMess & "try one method and if it fails or times out then you will need to Disconnect and re-connect using the alternate mode."

    MsgBox sMess, vbInformation + vbOKOnly, "FTP Active / Passive Connections"

    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdPassiveInfo_Click"
End Sub


 

Private Sub cmdViewHistory_Click()
    On Error GoTo EH
    Dim sPath As String
    Dim sSelFile As String
    Dim lFileSize As Long
    Dim sMyFilter As String
    Dim sFile As String
    Dim itmX As ListItem
    Dim lRet As Long
    Dim sFileName As String
    
    sFileName = "*_FTP.Log"

    sPath = msFTPLogPath

    sMyFilter = sMyFilter & "FTP LOG File" & " (" & sFileName & ")" & SD & sFileName & SD

    sPath = goUtil.utGetPath(App.EXEName, "FTP LOG HISTORY", "Manage FTP Log Files" & sFileName, "", sPath, Me.hWnd, sMyFilter, sSelFile)
    
    If sSelFile <> vbNullString Then
        sPath = msFTPLogPath & "\" & sSelFile
    End If
    
    If goUtil.utFileExists(sPath) Then
        lRet = goUtil.utShellExecute(GetDesktopWindow, "OPEN", sPath, vbNullString, App.Path, vbNormalFocus, False, False, True)
    End If

    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdViewHistory_Click"
End Sub


Private Sub EC_FTP_Progress(ByVal BytesTransfered As Long)
    On Error Resume Next
    mlProgressBytesTransferred = BytesTransfered
    lblBytesTransferred.Caption = mlProgressBytesTransferred
    lblBytesTransferred.Refresh
End Sub

Private Sub EC_FTP_StateChanged(ByVal NewState As FtpXCtl.StatesEnum, ByVal OldState As FtpXCtl.StatesEnum)
    On Error Resume Next
    If chkLogAllStatus.Value = vbChecked Then
        lstProcess.AddItem EC_FTP.StateString & " " & Now()
        lstProcess.ListIndex = lstProcess.NewIndex
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo EH
    Dim sMess As String
    Dim bDisablePasv As Boolean
    Dim bLogAllStatus As Boolean
    Dim bForceDownload As Boolean
    
    msIBSuffix = vbNullString
    
    
    
    ' Don't want to be visible initially!
    HideMe
    Me.Caption = "Communications Status"
    App.Title = Me.Caption
    
    goUtil.utFormWinRegPos goUtil.gsMainAppExeName, Me, , , , , True
    
    'Set the Install dir here !
    goUtil.gsInstallDir = GetSetting(goUtil.gsMainAppExeName, "Dir", "INSTALL_DIR", App.Path)
    'Need to save it in case this is the first time running (Using the App.path as default)
    SaveSetting goUtil.gsMainAppExeName, "Dir", "INSTALL_DIR", goUtil.gsInstallDir
    
    'Check the path to see if the install dir is different from app path
    If StrComp(goUtil.gsInstallDir, App.Path, vbTextCompare) <> 0 Then
        If InStr(1, Command$, "DEBUG", vbTextCompare) > 0 Then
            MsgBox "DEBUG: " & goUtil.gsInstallDir & vbCrLf & vbCrLf & "In: " & App.Path
        Else
            sMess = "Running " & App.EXEName & " from the following directory..." & vbCrLf
            sMess = sMess & App.Path & vbCrLf & vbCrLf
            sMess = sMess & App.EXEName & " was originally installed in this directory..." & vbCrLf
            sMess = sMess & goUtil.gsInstallDir & vbCrLf & vbCrLf
            sMess = sMess & "Click ""OK"" to network " & App.EXEName & " from this new directory."
            If MsgBox(sMess, vbExclamation + vbOKCancel, App.EXEName & " Directory Has Changed!") = vbCancel Then
                SaveSetting App.EXEName, "MSG", "COMMAND", "SHUT_DOWN_FTP"
                Exit Sub
            End If
            goUtil.gsInstallDir = App.Path
            Me.Caption = Me.Caption & " (Networked)"
            'Don't Save this new setting into the registry.  Want to give them the
            'same message every time they go into a directory other than the Installed one
            'So they know they are Networking to another Easy Claim
        End If
        
    End If
    
    'Make the Errorlog folder be in the same directory as the Install dir
    SaveSetting "ECS", "Dir", "ERRORLOG_DIR", goUtil.gsInstallDir
    
    'System Tray Icon...
    ' Retrieve broadcast message sent by
    ' Windows when taskbar is created.
    m_TaskbarCreated = RegisterWindowMessage(TaskbarCreatedString)
    
    ' Setup MsgHook
    Msghook.HwndHook = Me.hWnd
    Msghook.Message(cbNotify) = True
    ' Msghook only accepts Integer-ranged values
    If m_TaskbarCreated > &H7FFF& Then
      Msghook.Message(m_TaskbarCreated - &H10000) = True
    Else
      Msghook.Message(m_TaskbarCreated) = True
    End If
    
    ' Setup icon notification from shell
    AddTrayIcon
    
    'Show the Registration Form
    mbLoadingRegForm = True
    If Trim(Command$) <> "EZas123" & Format(Now, "DDYYMM") Then
        ShowRegForm
        'If this is the very first time running the Lic then
        'Unload until they put in the 10 day Lic.
        'Once this is done then any further Lic renewel will be
        'accomplished by connecting to the Server.
        If goUtil.gbInitLic And Not goUtil.gbValidLic Then
            MsgBox "Invalid TempCode! " & vbCrLf & vbCrLf & App.EXEName & " will exit.", vbExclamation, "Invalid Code"
            FlagShutDownFTP = True
            Exit Sub
        Else
            'Still show the Navigator tree even if Lic is expired.
            'Adjuster will only be allowed to connect to renew Lic if approved
            If Not goUtil.gbValidLic Then
                Me.Caption = Me.Caption & " (License Expired!)"
            End If
        End If
    Else
        goUtil.gbValidLic = True
        Me.Caption = Me.Caption & " {DEMO MODE}"
    End If
    mbLoadingRegForm = False
    
    '----------------Build Directory Paths--------------------
    goUtil.BuildPathsEasyClaim
    msFTPDLPath = goUtil.FTPDLPath
    msFTPULPath = goUtil.FTPULPath
    msFTPLogPath = goUtil.gsInstallDir & "\FTP\FTPLog"
    msSPPath = goUtil.SPPath
    msAttachReposPath = goUtil.AttachReposPath
    msPhotoReposPath = goUtil.PhotoReposPath
    '----------------END Build Directory Paths--------------------
    
    bDisablePasv = GetSetting(App.EXEName, "CONNECTION_SETTINGS", "DisablePasv", False)
    If Not bDisablePasv Then
        chkUsePasiveConnection.Value = vbChecked
        chkUsePasiveConnection.Caption = "Use Passive Connection"
    Else
        chkUsePasiveConnection.Value = vbUnchecked
        chkUsePasiveConnection.Caption = "Use Active Connection"
    End If
    
    bLogAllStatus = GetSetting(App.EXEName, "CONNECTION_SETTINGS", "LOG_ALL_STATUS", False)
    If bLogAllStatus Then
        chkLogAllStatus.Value = vbChecked
    Else
        chkLogAllStatus.Value = vbUnchecked
    End If
    
    bForceDownload = GetSetting(App.EXEName, "CONNECTION_SETTINGS", "ForceDownload", False)
    If bForceDownload Then
        chkForceDownload.Value = vbChecked
    Else
        chkForceDownload.Value = vbUnchecked
    End If
    
    'Start Listening to messages
    TimerMsg.Enabled = True
    
    'Connect when loading for the first time
    mbUpdateDB = True
    EnableCommandFrame False
    cmdConnect.Enabled = False
    cmdViewHistory.Enabled = False
    
    'Dcrypt USER FOLDERS
    msUserFolders = goUtil.Decode(USER_FOLDERS)
    msSPFolders = goUtil.Decode(SP_FOLDERS)
    
'    If MsgBox("Do you want to connect now?", vbQuestion + vbYesNo) = vbYes Then
'        Start_Comm
'    Else
        'Enable commands
        EnableCommandFrame True
        cmdConnect.Enabled = True
        cmdViewHistory.Enabled = True
        lblMess.Caption = "CLICK CONNECT!"
'    End If
    
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_Load"
End Sub

Public Sub ShowRegForm()
    On Error GoTo EH
    Dim sMess As String
    Dim frmReg As frmRegForm
    
    Set frmReg = New frmRegForm
    Load frmReg
    frmReg.Show
    Do Until Not frmReg.Visible
        DoEvents
        Sleep 100
    Loop
    Unload frmReg
    Set frmReg = Nothing

    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_Load"
End Sub


Public Function Start_Comm() As Boolean
    On Error GoTo EH
    Dim dicFiles As Scripting.Dictionary
    Dim vFile As Variant
    Dim sFile As String
    Dim sData As String
   
    If Not mbShutDownFTP And Not mbConnected Then
        'Need to Get rid of any Download and upload files
        'That never got processed in the event of an error
        'durring the previous connection.  These files
        'Will be created again the next time.
        '
        Set dicFiles = New Scripting.Dictionary
        'Check Download Files
        sFile = Dir(msFTPDLPath & "*.*", vbNormal)
        Do Until sFile = vbNullString
            dicFiles.Add msFTPDLPath & sFile, sFile
            sFile = Dir
        Loop
        sFile = Dir(msFTPULPath & "*.*", vbNormal)
        Do Until sFile = vbNullString
            dicFiles.Add msFTPULPath & sFile, sFile
            sFile = Dir
        Loop
        
        For Each vFile In dicFiles
            sFile = vFile
            goUtil.utDeleteFile sFile
        Next
        
        SaveAndClearHistory lstProcess

        lstProcess.AddItem Now() & " Starting Communications"
        lstProcess.AddItem "Please wait..."
        SaveSetting App.EXEName, "MSG", "CONNECTED", True
        Start_Comm = True
        lblMess.Caption = vbNullString
        ConnectNow
    Else
        Start_Comm = False
    End If
    
    Set dicFiles = Nothing
    Exit Function
EH:
    Start_Comm = False
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Function Start_Comm"
End Function

Private Sub EC_FTP_AsyncError(ByVal ErrorNum As Long, ByVal ErrorMsg As String)
  lstProcess.AddItem "Communication Timeout"
End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo EH
    If UnloadMode = vbFormControlMenu Then
        ' Just hide form if user presses Closes by X
        HideMe
        Cancel = True
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_QueryUnload"
End Sub

Private Sub Form_Resize()
    On Error GoTo EH
    If Not mbResize Then
'        VisibleFrames False
'        DoEvents
'        Sleep 100
        Timer_Resize.Enabled = True
    End If
    
    Exit Sub
EH:
    Err.Clear
End Sub

Public Sub VisibleFrames(pbVisible As Boolean)
    On Error GoTo EH
    framProcess.Visible = pbVisible
    framCommands.Visible = pbVisible
    cmdConnect.Visible = pbVisible
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub VisibleFrames"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo EH
    'First Save and clear any history
    SaveAndClearHistory lstProcess
    goUtil.utFormWinRegPos goUtil.gsMainAppExeName, Me, True, , , , True
    Call ShellNotifyIcon(NIM_DELETE, m_NID)
   Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_Unload"
End Sub

Private Sub lstProcess_Click()
    lstProcess.ToolTipText = lstProcess.Text
End Sub

Private Sub EC_FTP_DirItem(ByVal Item As String)
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    If Right(Item, 2) = vbCrLf Then
        Item = Left(Item, InStrRev(Item, vbCrLf, , vbBinaryCompare) - 1)
    End If
    If mdicRemoteFiles Is Nothing Then
        Set mdicRemoteFiles = New Scripting.Dictionary
    End If
    mdicRemoteFiles.Add LCase(Item), Item
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub EC_FTP_DirItem", False
    lstProcess.AddItem sErrDesc
End Sub

Public Sub ProcessFiles()
    On Error GoTo EH
    Dim vFile As Variant
    Dim lRetryCount As Long
    Dim sFilePath As String
    Dim sData As String

    For Each vFile In mdicRemoteFiles
        msCurrentFile = vFile
        If msCurrentFile = vbNullString Then
            Exit For
        End If

        If Trim(msCurrentFile) <> vbNullString Then
RETRY:
            On Error Resume Next
            sFilePath = goUtil.gsInstallDir & "\FTP\DownLoad\" & Trim(msCurrentFile)
            DoEvents
            Sleep 100
            EC_FTP.GetFile Trim(msCurrentFile), sFilePath
            'Check the current file just downloaded to see if it is 0 len
            If Not goUtil.utFileExists(sFilePath) Then
                lRetryCount = lRetryCount + 1
                If lRetryCount <= 10 Then
                    Sleep 500
                    GoTo RETRY
                End If
            ElseIf goUtil.utFileExists(sFilePath) Then
                sData = goUtil.utGetFileData(sFilePath)
                If Trim(sData) = vbNullString Then
                    lRetryCount = lRetryCount + 1
                    If lRetryCount <= 10 Then
                        Sleep 500
                        GoTo RETRY
                    End If
                End If
            End If
            If Err.Number <> 0 Then
                lstProcess.AddItem "Error # " & Err.Number & " " & Err.Description
                goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function ProcessFiles", False
                lblMess.Caption = "ERROR CONNECT AGAIN!"
                Err.Clear
            End If
            On Error GoTo EH
        End If
    Next

  Exit Sub
EH:
    Err.Raise Err.Number, , Err.Description & vbCrLf & msClassName & vbCrLf & "Public Sub ProcessFiles"
End Sub

Private Sub EC_FTP_Done(ByVal LastMethod As FtpXCtl.MethodsEnum, ByVal ErrorCode As Integer, ByVal ErrorString As String)
On Error GoTo EH
Dim sFilePath As String
Dim sData As String

  If LastMethod = FtpActionGetFilenameList And Not mbSingleFileProcess Then
    If Not mdicRemoteFiles Is Nothing Then
        'BGS 11.15.2001 Update the max value on the Bar here
        PBDownLoad.Value = 0
        PBDownLoad.Max = mdicRemoteFiles.Count
        ProcessFiles
    End If
  End If

    If LastMethod = FtpActionGetFile And Not mbSingleFileProcess Then
        sFilePath = goUtil.gsInstallDir & "\FTP\DownLoad\" & Trim(msCurrentFile)
        If goUtil.utFileExists(sFilePath) Then
            sData = goUtil.utGetFileData(sFilePath)
            If Trim(sData) = vbNullString Then
                'Do not delete 0 length files from the
                'server... That means ECFTP had a problem downloading it
                'Leav it on the server to be removed when appropriate.
                GoTo SHOW_ERROR
            End If
            If chkLogAllStatus.Value = vbChecked Then
                lstProcess.AddItem "Retrieved File " & Trim(msCurrentFile)
            End If

            On Error Resume Next
            PBDownLoad.Value = PBDownLoad.Value + 1
            PBDownLoad.Refresh
            On Error GoTo 0
            On Error Resume Next
            EC_FTP.Delete Trim(msCurrentFile)
            On Error GoTo 0
        Else
SHOW_ERROR:
            lstProcess.AddItem "Error getting file " & Trim(msCurrentFile)
        End If
    End If
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub EC_FTP_Done", False
End Sub

Private Function GetIBSuffix() As String
    On Error GoTo EH
    Dim lNext As Long
    'BGS 11.20.2001 Since we are entering in multiple claims all at the same time
    'the IBID has to increment starting with the initial time Hack
    If msIBSuffix = vbNullString Then
        msIBSuffix = Format(Now, "YYMMDDHHMMSS")
        GetIBSuffix = msIBSuffix
    Else 'BGS if we already have the first suffix need to increment the next one
        lNext = CLng(Right(msIBSuffix, 6))
        lNext = lNext + 1
        msIBSuffix = Left(msIBSuffix, 6) & CStr(lNext)
        GetIBSuffix = msIBSuffix
    End If

    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Function GetIBSuffix"
End Function


Private Sub ConnectNow()
    On Error GoTo EH
    'Verify User Password / Licensing
    Dim vAryToken(0 To 100) As Variant
    Dim sUserName As String
    Dim sUserNameDecoded As String
    Dim sSSN As String
    Dim sPass As String
    Dim sOldPass As String
    Dim bResetPass As Boolean
    Dim sLicDaysLeft As String
    Dim sAppVSInfo As String
    Dim sEmail As String
    Dim sTeamLeader As String
    Dim sContactPhone As String
    Dim sFName As String
    Dim sLName As String
    Dim sCarrier As String
    '---
    Dim sFTPLogonName As String
    Dim sFTPLogonPass As String
    Dim lDLFileTypeCount As Long
    Dim lCount As Long
    'Download FIles
    Dim dicDLFiles As Scripting.Dictionary
    Dim vDLFile As Variant
    Dim sDLFile As String
    'upload FIles
    Dim dicULFiles As Scripting.Dictionary
    Dim vULFile As Variant
    Dim sULFile As String
    
    Dim oXZip As V2ECKeyBoard.clsXZip
    Dim bDownloadDBSPSuccess As Boolean
    Dim sEasyClaimMsg As String
    
    'DB Backup
    Dim sTemp As String
    Dim sMainDBFullPath As String
    Dim sBackupFullPath As String
    
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim lRetry As Long
    
    If mbConnected Or mbShutDownFTP Then
        Exit Sub
    End If

    'BGS set all the Security variables here
    vAryToken(SecurityToken.AppVSInfo) = Replace(goUtil.utGetAppVSInfo(gsMainAppExeName, goUtil.gsInstallDir), vbCrLf, F_VBCRLF)
    vAryToken(SecurityToken.Carrier) = GetSetting(gsMainAppExeName, "GENERAL", "CURRENT_CAR", vbNullString)
    vAryToken(SecurityToken.Company) = GetSetting(gsMainAppExeName, "GENERAL", "CURRENT_COMPANY", vbNullString)
    vAryToken(SecurityToken.ContactPhone) = GetSetting(gsMainAppExeName, "GENERAL", "ADJ_CONTACT_PHONE", vbNullString)
    vAryToken(SecurityToken.Email) = GetSetting(gsMainAppExeName, "GENERAL", "ADJ_EMAIL", vbNullString)
    vAryToken(SecurityToken.FName) = GetSetting(gsMainAppExeName, "GENERAL", "ADJUSTOR_FIRST_NAME", vbNullString)
    vAryToken(SecurityToken.IBPrefix) = goUtil.utGetIBPREFIX
    vAryToken(SecurityToken.iZip) = GetSetting(gsMainAppExeName, "GENERAL", "ADJ_ZIP", vbNullString)
    vAryToken(SecurityToken.iZip4) = GetSetting(gsMainAppExeName, "GENERAL", "ADJ_ZIP4", vbNullString)
    vAryToken(SecurityToken.LicDaysLeft) = GetSetting("ECS", "WEB_SECURITY", "LIC", vbNullString)
    vAryToken(SecurityToken.LName) = GetSetting(gsMainAppExeName, "GENERAL", "ADJUSTOR_LAST_NAME", vbNullString)
    vAryToken(SecurityToken.OldPass) = GetSetting("ECS", "WEB_SECURITY", "OLD_PASSWORD", vbNullString)
    vAryToken(SecurityToken.Pass) = GetSetting("ECS", "WEB_SECURITY", "PASSWORD", vbNullString)
    vAryToken(SecurityToken.ResetPass) = GetSetting("ECS", "WEB_SECURITY", "RESET_PASSWORD", False)
    vAryToken(SecurityToken.sAddress) = GetSetting(gsMainAppExeName, "GENERAL", "ADJ_ADDRESS", vbNullString)
    vAryToken(SecurityToken.sCity) = GetSetting(gsMainAppExeName, "GENERAL", "ADJ_CITY", vbNullString)
    vAryToken(SecurityToken.sEmergencyPhone) = GetSetting(gsMainAppExeName, "GENERAL", "ADJ_EMERGENCY_PHONE", vbNullString)
    vAryToken(SecurityToken.sOtherPostCode) = GetSetting(gsMainAppExeName, "GENERAL", "ADJ_OTHER_POSTCODE", vbNullString)
    vAryToken(SecurityToken.SSN) = GetSetting("ECS", "WEB_SECURITY", "SSN", vbNullString)
    vAryToken(SecurityToken.sState) = GetSetting(gsMainAppExeName, "GENERAL", "ADJ_STATE", vbNullString)
    vAryToken(SecurityToken.TeamLeader) = GetSetting(gsMainAppExeName, "GENERAL", "TEAM_LEADER", vbNullString)
    vAryToken(SecurityToken.TokenType) = TokenType.Security
    vAryToken(SecurityToken.UserName) = GetSetting("ECS", "WEB_SECURITY", "USER_NAME", vbNullString)

    
    sUserName = vAryToken(SecurityToken.UserName)
    sUserNameDecoded = sUserName
    sUserNameDecoded = goUtil.Decode(sUserNameDecoded)
    sSSN = vAryToken(SecurityToken.SSN)
    sPass = vAryToken(SecurityToken.Pass)
    sOldPass = vAryToken(SecurityToken.OldPass)
    bResetPass = vAryToken(SecurityToken.ResetPass)
    sLicDaysLeft = vAryToken(SecurityToken.LicDaysLeft)
    sAppVSInfo = vAryToken(SecurityToken.AppVSInfo)
    sEmail = vAryToken(SecurityToken.Email)
    sTeamLeader = vAryToken(SecurityToken.TeamLeader)
    sContactPhone = vAryToken(SecurityToken.ContactPhone)
    sFName = vAryToken(SecurityToken.FName)
    sLName = vAryToken(SecurityToken.LName)
    sCarrier = vAryToken(SecurityToken.Carrier)
    'BGS 11.20.2001 Need to Check for UserName before we can allow Upload
    If sUserName = vbNullString Then
        MsgBox "You must provide User Name before uploading any billing information!", vbExclamation + vbOKOnly, "User Name Required !"
        gfrmECTray.EasyClaimCommand "GlobalPref"
        mbUpdateDB = False
        EnableCommandFrame True
        cmdConnect.Enabled = True
        cmdViewHistory.Enabled = True
        Exit Sub
    End If

    'BGS 11.20.2001 Need to Check for SSN before we can allow Upload
    If sSSN = vbNullString Then
        MsgBox "You must provide SSN before uploading any billing information!", vbExclamation + vbOKOnly, "SSN Required !"
        SaveSetting goUtil.gsMainAppExeName, "MSG", "FTP_COMMAND", "GLOBAL_PREF"
        lstProcess.AddItem "You must provide SSN before uploading any billing information!"
        lstProcess.AddItem Now() & " Connection Aborted."
        m_NID.hIcon = imgList.ListImages(PicList.FTP03).Picture
        Call ShellNotifyIcon(NIM_MODIFY, m_NID)
        mbUpdateDB = False
        EnableCommandFrame True
        cmdConnect.Enabled = True
        cmdViewHistory.Enabled = True
        Exit Sub
    End If

    'BGS 6.3.2002 Need to Check for Password before we can allow Upload
    If sPass = vbNullString Then
        MsgBox "You must provide a Password before uploading any information!", vbExclamation + vbOKOnly, "Password Required !"
        SaveSetting goUtil.gsMainAppExeName, "MSG", "COMMAND", "GLOBAL_PREF"
        lstProcess.AddItem "You must provide a Password before uploading any information!"
        lstProcess.AddItem Now() & " Connection Aborted."
        m_NID.hIcon = imgList.ListImages(PicList.FTP03).Picture
        Call ShellNotifyIcon(NIM_MODIFY, m_NID)
        mbUpdateDB = False
        EnableCommandFrame True
        cmdConnect.Enabled = True
        cmdViewHistory.Enabled = True
        Exit Sub
    End If
    
    mlProgressBytesTransferred = 0

    lblProcess.Caption = "Connecting..."
    lblProcess.Refresh
    EC_FTP.LicenseKey = "5TBQ-XKBNZ2C88UL1"
    EC_FTP.Host = GetSetting("ECS", "WEB_SECURITY", "WEB_HOST", "www.eberls.net") '"206.196.148.98"

    'BGS 6.12.2001 Use this new User name and Password to connect
    '12.16.2002 If a new User name and Password for FTP logon is needed...
    'A service pack would have to update the registry settings with the new User and Password.
    On Error Resume Next
GET_FTP_LOGON_NAME:
    sFTPLogonName = GetSetting("ECS", "WEB_SECURITY", "FTP_LOGON_NAME", " u61223-e6o h02223-E2i o91223-s3h d51223- 7a n81223-Z4k d12223- 1t x41223-U8h e71223-r5m")
    sFTPLogonName = goUtil.Decode(sFTPLogonName)
    If Err.Number <> 0 Then
        Err.Clear
        'EZUser (Default FTP User Name)
        SaveSetting "ECS", "WEB_SECURITY", "FTP_LOGON_NAME", " u61223-e6o h02223-E2i o91223-s3h d51223- 7a n81223-Z4k d12223- 1t x41223-U8h e71223-r5m"
        On Error GoTo EH
        GoTo GET_FTP_LOGON_NAME
    End If

    On Error Resume Next
GET_FTP_LOGON_PASS:
    sFTPLogonPass = GetSetting("ECS", "WEB_SECURITY", "FTP_LOGON_PASS", " w61223-s6q k51223-a7z m91223- 3q b71223-E5v m41223-18z s81223-34u q02223-Z2k s12223-21g")
    sFTPLogonPass = goUtil.Decode(sFTPLogonPass)
    If Err.Number <> 0 Then
        Err.Clear
        'EZas123
        SaveSetting "ECS", "WEB_SECURITY", "FTP_LOGON_PASS", " w61223-s6q k51223-a7z m91223- 3q b71223-E5v m41223-18z s81223-34u q02223-Z2k s12223-21g"
        On Error GoTo EH
        GoTo GET_FTP_LOGON_PASS
    End If
On Error GoTo EH
    EC_FTP.LogonName = sFTPLogonName
    EC_FTP.LogonPassword = sFTPLogonPass
    'Check for Active or Passive Connection
    If chkUsePasiveConnection.Value = vbChecked Then
        EC_FTP.DisablePasv = False   ' Use Passive Connection
    Else
        EC_FTP.DisablePasv = True    'Use Active Connection
    End If
    
    SetFireWallSettings
    
    EC_FTP.Connect
    mbConnected = True
    
'<-------------  check for any Error messages , and upload them to the Server.----------->
If goUtil.utFileExists(goUtil.gsInstallDir & "\ErrorLog", True) Then
    Set dicULFiles = New Scripting.Dictionary
    sULFile = Dir(goUtil.gsInstallDir & "\ErrorLog\" & "*.*", vbNormal)
    Do Until sULFile = vbNullString
        dicULFiles.Add sULFile, sULFile
        sULFile = Dir
    Loop
    'notify of progress
    If dicULFiles.Count > 0 Then
        lblProcess.Caption = "Uploading Error Report to server... Please Wait!"
        lstProcess.AddItem lblProcess.Caption
        PBUpLoad.Max = dicULFiles.Count
        mbSingleFileProcess = True
        On Error Resume Next
        EC_FTP.ChangeDir "\.\"
        EC_FTP.ChangeDir msUserFolders & "\USER_FOLDERS\" & sUserNameDecoded & "\"
        'see if the Userfolder Exisits if not create it
        If Err.Number <> 0 Then
            Err.Clear
            On Error Resume Next
            EC_FTP.CreateDir msUserFolders & "\USER_FOLDERS\" & sUserNameDecoded & "\"
            EC_FTP.ChangeDir "\.\"
            EC_FTP.ChangeDir msUserFolders & "\USER_FOLDERS\" & sUserNameDecoded & "\"
        End If
        mbSingleFileProcess = False
    End If
    For Each vULFile In dicULFiles
        lCount = lCount + 1
        sULFile = vULFile
        mbSingleFileProcess = True
        On Error Resume Next
        lRetry = 0
RETRY:
        DoEvents
        Sleep 100
        EC_FTP.PutFile goUtil.gsInstallDir & "\ErrorLog\" & sULFile, sULFile
        If Err.Number <> 0 Then
            lRetry = lRetry + 1
            If lRetry <= 10 Then
                Sleep 500
                Err.Clear
                On Error Resume Next
                GoTo RETRY
            End If
            lstProcess.AddItem "ERROR #" & Err.Number & " File(" & sULFile & ") "
            lstProcess.AddItem Err.Description
        Else
            'remove the Error Log file only if there was not an error
            'trying to upload the file.
            goUtil.utDeleteFile goUtil.gsInstallDir & "\ErrorLog\" & sULFile
        End If
        mbSingleFileProcess = False
        PBUpLoad.Value = lCount
        DoEvents
        Sleep 10
    Next
    PBUpLoad.Value = 0
    'cleanup
    Set dicULFiles = Nothing
    mbSingleFileProcess = True
    EC_FTP.ChangeDir "\.\"
    mbSingleFileProcess = False
    On Error GoTo EH
End If
'<----------  END check for any Error messages , and upload them to the Server.-------->

'>-----------------------INSERT SECURITY CHECK HERE----------------------<

'Set the Main DB
goUtil.SetMainDB App.EXEName, goUtil.gsInstallDir & "\ECMain.mdb", , True
If Not VerifySecurity(vAryToken) Then
    lblProcess.Caption = Now() & " Verify User/Password Failed"
    lstProcess.AddItem lblProcess.Caption
    m_NID.hIcon = imgList.ListImages(PicList.FTP03).Picture
    Call ShellNotifyIcon(NIM_MODIFY, m_NID)
    mbUpdateDB = False
    EnableCommandFrame True
    cmdConnect.Enabled = True
    cmdViewHistory.Enabled = True
    mbUpdateDB = False
    EC_FTP.Disconnect
    mbConnected = False
    SaveSetting App.EXEName, "MSG", "CONNECTED", False
    lblMess.Caption = "ERROR CONNECT AGAIN!"
    Exit Sub
Else
    If Not goUtil.gbValidLic Then
        SaveSetting gsMainAppExeName, "MSG", "FTP_COMMAND", "SHOW_ABOUT"
        mbUpdateDB = False
        EC_FTP.Disconnect
        mbConnected = False
        SaveSetting App.EXEName, "MSG", "CONNECTED", False
        Unload Me
        Exit Sub
    End If
End If
'>-----------------------END SECURITY CHECK HERE----------------------<

'<---12/1/2003 DownLoad Tables From Server---->
lstProcess.AddItem "Connected to Eberls.com at " & Now()
lblProcess.Caption = ""
lblProcess.Caption = "Downloading Records From Server.  Please wait!!!"
lstProcess.AddItem lblProcess.Caption
lblProcess.Refresh

'Need zip utility to uncompress download files
Set oXZip = New V2ECKeyBoard.clsXZip
Sleep 500


'Need to Synchronize UL Data
'Do this here just in case there were problems
'on the previous connection.  This call should not process anything
'unless there were problems durring previous connection.
If Not SynchronizeULData() Then
    lstProcess.AddItem "Synchronize Data Halted.  Please try again later."
    lblMess.Caption = "ERROR CONNECT AGAIN!"
    GoTo CLEANUP
End If

'1 Do Look Up Tables First
'2 Do DownLoad Tables
'3 Check DB Version
For lDLFileTypeCount = 1 To 3
    Set mdicRemoteFiles = Nothing
    If lDLFileTypeCount = 1 Then
        lblProcess.Caption = "Getting DB_VERSION..."
        'Give the server a chance to create the DB Version File
        Sleep 2000
        DoEvents
    ElseIf lDLFileTypeCount = 2 Then
        lblProcess.Caption = "Getting Look up Tables List..."
    ElseIf lDLFileTypeCount = 3 Then
        lblProcess.Caption = "Getting Down Load Tables List..."
    End If
   
    lblProcess.Refresh
    lstProcess.AddItem lblProcess.Caption
    'Add the pattern to que up files to download from server
    'Be sure that the patters is LCASE if it is not
    'then the return will miss anything in the patter that
    'is lcase. Making the Pattern lcase will include anything
    'on the server that is UCASE.
    
    If lDLFileTypeCount = 1 Then
        EC_FTP.Pattern = LCase("*.dllu")
    ElseIf lDLFileTypeCount = 2 Then
        EC_FTP.Pattern = LCase("*.zdllu")
     ElseIf lDLFileTypeCount = 3 Then
        EC_FTP.Pattern = LCase("*.zdl")
    End If
    
    'This Process will get the File list
    'From server... As Well it will fire the
    'EC_FTP_DirItem event.  inside this event
    'anything the on the server that fits the pattern
    'will be populated into mdicRemoteFiles dictionary
    'Then the EC_FTP_Done event will fire and will Download
    'and then remove the files in the mdicRemoteFiles dictionary object.
    EC_FTP.GetFilenameList
    
    
    'Once the files have been downloaded, they need to be unzipped...
    
    '1. Unzip All the DL files
    Set dicDLFiles = New Scripting.Dictionary
    If lDLFileTypeCount = 1 Then
        sDLFile = Dir(msFTPDLPath & "*.dllu", vbNormal)
    ElseIf lDLFileTypeCount = 2 Then
        sDLFile = Dir(msFTPDLPath & "*.zdllu", vbNormal)
    ElseIf lDLFileTypeCount = 3 Then
        sDLFile = Dir(msFTPDLPath & "*.zdl", vbNormal)
    End If
    
    Do Until sDLFile = vbNullString
        dicDLFiles.Add sDLFile, sDLFile
        sDLFile = Dir
    Loop
    
    For Each vDLFile In dicDLFiles
        sDLFile = vDLFile
        '6.15.2004 BGS Only unzip compressed Files
        If lDLFileTypeCount = 1 Then
            '1.
            'Check For Data base Version Changes
            'False return means a Version Upgrade Required
            If Not CheckDBVerison(bDownloadDBSPSuccess, sUserNameDecoded) Then
                'if the SP was successfully downloaded then shell it
                If bDownloadDBSPSuccess Then
                    MsgBox "Required Database Update!" & vbCrLf & vbCrLf & "Please save your work and then click OK", vbExclamation + vbOKOnly, "Required Database Update"
                    mbDatabaseUpgrade = True
                Else
                    lstProcess.AddItem "Problems downloading Data Base Service Pack.  Please try again later."
                End If
                GoTo CLEANUP
            Else
                'Need to inform Server to continue with creating Lookup and Download Files
                UploadClientFlag sUserNameDecoded, "DBUpToDate.flag", True
                'Check the Security Again to give the server a chance to create
                'Download files before FTP tries to download them
                If Not VerifySecurity(vAryToken, True, "Waiting for Lookup Files ", 120) Then
                    lstProcess.AddItem "Server not responding!  Please try again later. "
                    GoTo CLEANUP
                End If
            End If
        Else
            oXZip.UNZipFiles msFTPDLPath, msFTPDLPath & sDLFile, False
            goUtil.utDeleteFile msFTPDLPath & sDLFile
        End If
    Next
    If lDLFileTypeCount = 2 Then
        '2.
        'Update Lookup Info
        'These are the tables that contain all the look up information concerning
        'What Cat, what software, what Cat ID, What Adjuster ID , Fee Schedule, Class Of Loss, Type Of Loss
        'etc.  Basically the Who, What, Where When, Why, and How of Easy Claim.
        If Not UpdateLookUpInfo() Then
            'Need to inform Server that software update is required if this fails
            'so Process Tokins will end on server.
            UploadClientFlag sUserNameDecoded, "SoftwareUpToDate.flag", False
            lstProcess.AddItem "Lookup Info Update Halted.  Please try again later."
            GoTo CLEANUP
        End If
        
        '3.
        'After updating the Lookup info,  It is now possible to look at the Software Situation.
        'Update Software changes that need to be accomplished to get the User Sofware up to speed--->
        'This May Require the Application to Restart.
        'This function MUST RETURN TRUE BEFORE AND SUBSEQUENT FUNCTIONS CAN BE ALLOWED.
        If UpdateSoftwareChanges(sUserNameDecoded) Then
            lstProcess.AddItem "Software Changes Required... Please Exit."
            GoTo CLEANUP
        Else
            'Need to inform Server to continue with creating Lookup and Download Files
            UploadClientFlag sUserNameDecoded, "SoftwareUpToDate.flag", True
            'Check the Security Again to give the server a chance to create
            'Download files before FTP tries to download them
            If Not VerifySecurity(vAryToken, True, "Waiting for Download Files ", 120) Then
                lstProcess.AddItem "Server not responding!  Please try again later. "
                GoTo CLEANUP
            End If
        End If
    End If
Next

'After downloading and unzipping the files from the server, need to process them...

'4.
'Update Download Info
'This will include All the Production Data that is Sitting Server Side that
'Needs to be updated on the Client Side.
If Not UpdateDownLoadInfo() Then
    lstProcess.AddItem "Download Info Halted.  Please try again later."
    lblMess.Caption = "ERROR CONNECT AGAIN!"
    GoTo CLEANUP
End If
'5.Update Photos and Attachments that need to be Downloaded from server
If Not UpdatePhotoAttatchmentDownload() Then
    lblMess.Caption = "ERROR CONNECT AGAIN!"
    GoTo CLEANUP
End If

'6.
'Update Upload Info
'This Will Include all the Production Data sitting on the CLient Side that needs
'to be upated on the Server Side.
If Not UpdateUpLoadInfo(vAryToken) Then
    lstProcess.AddItem "Upload Info Halted.  Please try again later."
    lblMess.Caption = "ERROR CONNECT AGAIN!"
    GoTo CLEANUP
End If

'7. Update Photos and Attachments that need to be uploaded to server
If Not UpdatePhotoAttatchmentUpload() Then
    lstProcess.AddItem "Photos and Attachment Upload Halted.  Please try again later."
    lblMess.Caption = "ERROR CONNECT AGAIN!"
    GoTo CLEANUP
End If

CLEANUP:
    'Instead of compacting and repairing , just do a straight copy of the existing DB
    'to the backup.  This way the backup will have the most current and synchronized Data.
    'Update Message
    lblProcess.Caption = "Creating Backup... " & Now()
    lstProcess.AddItem lblProcess.Caption
    lblProcess.Refresh
    
    sMainDBFullPath = goUtil.gMainDB.Name
    sTemp = Left(sMainDBFullPath, InStrRev(sMainDBFullPath, "\", , vbBinaryCompare))
    sBackupFullPath = sTemp
    sTemp = "ECMain_BackUp.db"
    sBackupFullPath = sBackupFullPath & sTemp
    
    sTemp = goUtil.utCopyFile(sMainDBFullPath, sBackupFullPath)
    
    If sTemp = vbNullString Then
        lblProcess.Caption = "Backup Complete... " & Now()
    Else
        lblProcess.Caption = "Problems creating backup... " & sTemp & " " & Now()
    End If
    lstProcess.AddItem lblProcess.Caption
    lblProcess.Refresh
    
    'Close the Main DB
    goUtil.CloseMainDB
    'Inform Easy Claim to reload tree after FTP Connection
    sEasyClaimMsg = GetSetting("EasyClaim", "MSG", "COMMAND", vbNullString) '"SHUT_DOWN_COMPLETE"
    If StrComp(sEasyClaimMsg, "SHUT_DOWN_COMPLETE", vbTextCompare) <> 0 Then
        SaveSetting "EasyClaim", "MSG", "COMMAND", "LOAD_TREE"
    End If
    
    '6. Send client flag telling server it is ok to stop tokin process
    UploadClientFlag sUserNameDecoded, "ShutDownTokinProcess.flag", True
    
    'Clean up
    Set dicDLFiles = Nothing
    Set dicULFiles = Nothing
    Set oXZip = Nothing
    Set mdicRemoteFiles = Nothing
    msCurrentFile = vbNullString
    
    'Disconnect FTP
    EC_FTP.Disconnect
    mbConnected = False
    
    'Update Message of Disconnect
    SaveSetting App.EXEName, "MSG", "CONNECTED", False
    
    'Update Icon
    m_NID.hIcon = imgList.ListImages(PicList.FTP01).Picture
    Call ShellNotifyIcon(NIM_MODIFY, m_NID)
    
    'Update Message
    lstProcess.AddItem "Processing Complete...Disconnected " & Now()
    lblProcess.Caption = "Process Complete"
    lblProcess.Refresh
    
    ShowNotePadMess
    
    'Enable commands
    EnableCommandFrame True
    cmdConnect.Enabled = True
    cmdViewHistory.Enabled = True
  Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    
    'Log the error but don't show it since its already shown above
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub ConnectNow", False
    
    'Use message box instead of Error handler
    lstProcess.AddItem Now() & " Problems during connection... "
    lstProcess.AddItem "Is your Internet Service Provider (ISP) connected to the Internet?"
    lstProcess.AddItem "If you are connected, try disconnecting and reconnecting to the internet."
    lstProcess.AddItem "If you still have problems please call technical support."
    lstProcess.AddItem "Thank you."
    lstProcess.AddItem vbNullString
    lstProcess.AddItem "Technical Support Data: "
    lstProcess.AddItem "<--------------------------BEGIN ERROR REPORT-------------------------->"
    lstProcess.AddItem "Error # " & lErrNum
    lstProcess.AddItem "Description: " & sErrDesc
    lstProcess.AddItem "<---------------------------END ERROR REPORT--------------------------->"
    
    
    'Close the Main DB
    goUtil.CloseMainDB
    
    'Update Icon
    m_NID.hIcon = imgList.ListImages(PicList.FTP03).Picture
    Call ShellNotifyIcon(NIM_MODIFY, m_NID)
    
    'Enable Commands
    EnableCommandFrame True
    cmdConnect.Enabled = True
    cmdViewHistory.Enabled = True
    'Clean up
    Set dicDLFiles = Nothing
    Set dicULFiles = Nothing
    Set oXZip = Nothing
    Set mdicRemoteFiles = Nothing
    msCurrentFile = vbNullString
    
    cmdCancelPhotoAttachUL.Enabled = False
    ' Send client flag telling server it is ok to stop tokin process
    UploadClientFlag sUserNameDecoded, "ShutDownTokinProcess.flag", True
    EC_FTP.Disconnect
    mbConnected = False
    SaveSetting App.EXEName, "MSG", "CONNECTED", False
    m_NID.hIcon = imgList.ListImages(PicList.FTP01).Picture
    Call ShellNotifyIcon(NIM_MODIFY, m_NID)
    mbSingleFileProcess = False
    ShowNotePadMess
End Sub

Private Sub UploadFiles(psFileNamePattern As String)
    On Error GoTo EH
    Dim dicULFiles As Scripting.Dictionary
    Dim lCount As Long
    Dim sULFile As String
    Dim vULFile As Variant
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim lRetry As Long
    
    Set dicULFiles = New Scripting.Dictionary
    
    sULFile = Dir(msFTPULPath & psFileNamePattern, vbNormal)
    Do Until sULFile = vbNullString
        dicULFiles.Add sULFile, sULFile
        sULFile = Dir
    Loop
    
    mbSingleFileProcess = True
    On Error Resume Next
    EC_FTP.ChangeDir "\.\"
    EC_FTP.ChangeDir msUserFolders & "\USER_FOLDERS\" & msUserName & "\"
    PBUpLoad.Max = dicULFiles.Count
    For Each vULFile In dicULFiles.Items
        sULFile = vULFile
        lstProcess.AddItem "Uploading " & sULFile & " " & Now()
        lRetry = 0
RETRY:
        DoEvents
        Sleep 100
        EC_FTP.PutFile msFTPULPath & sULFile, sULFile
        If Err.Number <> 0 Then
            lRetry = lRetry + 1
            If lRetry <= 10 Then
                Sleep 500
                Err.Clear
                On Error Resume Next
                GoTo RETRY
            End If
            lstProcess.AddItem "ERROR #" & Err.Number & " File(" & sULFile & ") "
            lstProcess.AddItem Err.Description
            lblProcess.Caption = "ERROR CONNECT AGAIN!"
        End If
        goUtil.utDeleteFile msFTPULPath & sULFile
        lCount = lCount + 1
        PBUpLoad.Value = lCount
    Next
    
    EC_FTP.ChangeDir "\.\"
    mbSingleFileProcess = False
    
    'cleanup
    Set dicULFiles = Nothing
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    'Log the error but don't show it since its already shown above
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub UploadFiles", False
    
    lstProcess.AddItem "ERROR #" & lErrNum
    lstProcess.AddItem sErrDesc
End Sub


Private Sub UploadClientFlag(psUserNameDecoded As String, psFlagName As String, pbFlag As Boolean)
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim lRetry As Long
    
    'Need to inform Server to continue with creating Lookup and Download Files
    goUtil.utSaveFileData msFTPULPath & psFlagName, CStr(pbFlag)
    'Change FTP Dr and Update the Server Side UPdate
    mbSingleFileProcess = True
    On Error Resume Next
    EC_FTP.ChangeDir "\.\"
    EC_FTP.ChangeDir msUserFolders & "\USER_FOLDERS\" & psUserNameDecoded & "\"
    lRetry = 0
RETRY:
    DoEvents
    Sleep 100
    EC_FTP.PutFile msFTPULPath & psFlagName, psFlagName
    If Err.Number <> 0 Then
        lRetry = lRetry + 1
        If lRetry <= 10 Then
            Sleep 500
            Err.Clear
            On Error Resume Next
            GoTo RETRY
        End If
        lstProcess.AddItem "ERROR #" & Err.Number & " File(" & psFlagName & ") "
        lstProcess.AddItem Err.Description
        lblProcess.Caption = "ERRROR CONNECT AGAIN!"
    End If
    goUtil.utDeleteFile msFTPULPath & psFlagName
    EC_FTP.ChangeDir "\.\"
    mbSingleFileProcess = False
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    'Log the error but don't show it since its already shown above
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub UploadClientFlag", False
    
    lstProcess.AddItem "ERROR #" & lErrNum
    lstProcess.AddItem sErrDesc
    
End Sub

Private Function CheckDBVerison(pbDownloadDBSPSuccess As Boolean, psUserNameDecoded As String) As Boolean
    'This function retruns true if Database Version info is up to date
    'Returns False if DataBase Upgrade Required
    On Error GoTo EH
    Dim sData As String
    Dim saryRecords() As String
    Dim lRecordsCount As Long
    Dim saryFields() As String
    Dim lFieldsCount As Long
    Dim sFieldValue As String
    Dim RS As ADODB.Recordset
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    Dim oField As ADODB.Field
    Dim lVersion As Long
    Dim sSPName As String
    Dim sInstallFileLocation As String
    Dim dtVersionDate As Date
    Dim lOldVersion As Long
    Dim sOldSPName As String
    Dim sOldInstallFileLocation As String
    Dim sComments As String
    Dim sOldComments As String
    Dim dtOldVersionDate As Date
    'Main Util
    Dim sMainUtilSPName As String
    Dim sOldMainUtilSPName As String
    Dim sMainUtilInstallFileLocation As String
    Dim sOldMainUtilInstallFileLocation As String
    'Main ARV
    Dim sMainARVSPName As String
    Dim sOldMainARVSPName As String
    Dim sMainARVInstallFileLocation As String
    Dim sOldMainARVInstallFileLocation As String
    'Main EXE
    Dim sMainEXESPName As String
    Dim sOldMainEXESPName As String
    Dim sMainEXEInstallFileLocation As String
    Dim sOldMainEXEInstallFileLocation As String
    'Main FTP EXE
    Dim sMainFTPEXESPName As String
    Dim sOldMainFTPEXESPName As String
    Dim sMainFTPEXEInstallFileLocation As String
    Dim sOldMainFTPEXEInstallFileLocation As String
    Dim lUpdateByUserID As Long
    Dim lOldUpdateByUserID As Long
    
    Set oConn = New ADODB.Connection
    Set RS = New ADODB.Recordset
    
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    'set table records to array
    sData = goUtil.utGetFileData(msFTPDLPath & "DB_VERSION.dllu")
    'Now that we got it need to remove the .dl file
    goUtil.utDeleteFile msFTPDLPath & "DB_VERSION.dllu"
    
    saryRecords() = Split(sData, RECORD_DELIM, , vbBinaryCompare)
    'Only one rcord in this table !
    sData = saryRecords(0)
    saryFields() = Split(sData, COLUMN_DELIM, , vbBinaryCompare)
    
    sSQL = "SELECT TOP 1 * FROM DB_VERSION "
    
    RS.CursorLocation = adUseClient
    RS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set RS.ActiveConnection = Nothing
    
    lFieldsCount = 0
    
    For Each oField In RS.Fields
        sFieldValue = saryFields(lFieldsCount)
        Select Case UCase(oField.Name)
            Case UCase("Version")
                lOldVersion = IIf(IsNull(oField.Value), 0, oField.Value)
                lVersion = CLng(sFieldValue)
            Case UCase("Comments")
                sOldComments = IIf(IsNull(oField.Value), 0, oField.Value)
                sComments = sFieldValue
            Case UCase("InstallFileLocation")
                sOldInstallFileLocation = IIf(IsNull(oField.Value), vbNullString, oField.Value)
                sOldInstallFileLocation = Replace(sOldInstallFileLocation, "{InstallDir}", goUtil.gsInstallDir, , , vbTextCompare)
                sInstallFileLocation = sFieldValue
                sInstallFileLocation = Replace(sInstallFileLocation, "{InstallDir}", goUtil.gsInstallDir, , , vbTextCompare)
            Case UCase("SPName")
                sOldSPName = IIf(IsNull(oField.Value), vbNullString, oField.Value)
                sSPName = sFieldValue
            Case UCase("MainUtilInstallFileLocation")
                sOldMainUtilInstallFileLocation = IIf(IsNull(oField.Value), vbNullString, oField.Value)
                sOldMainUtilInstallFileLocation = Replace(sOldMainUtilInstallFileLocation, "{SystemDir}", goUtil.utGetSystemDir, , , vbTextCompare)
                sMainUtilInstallFileLocation = sFieldValue
                sMainUtilInstallFileLocation = Replace(sMainUtilInstallFileLocation, "{SystemDir}", goUtil.utGetSystemDir, , , vbTextCompare)
            Case UCase("MainUtilSPName")
                sOldMainUtilSPName = IIf(IsNull(oField.Value), vbNullString, oField.Value)
                sMainUtilSPName = sFieldValue
            Case UCase("MainARVInstallFileLocation")
                sOldMainARVInstallFileLocation = IIf(IsNull(oField.Value), vbNullString, oField.Value)
                sOldMainARVInstallFileLocation = Replace(sOldMainARVInstallFileLocation, "{SystemDir}", goUtil.utGetSystemDir, , , vbTextCompare)
                sMainARVInstallFileLocation = sFieldValue
                sMainARVInstallFileLocation = Replace(sMainARVInstallFileLocation, "{SystemDir}", goUtil.utGetSystemDir, , , vbTextCompare)
            Case UCase("MainARVSPName")
                sOldMainARVSPName = IIf(IsNull(oField.Value), vbNullString, oField.Value)
                sMainARVSPName = sFieldValue
            Case UCase("MainEXEInstallFileLocation")
                sOldMainEXEInstallFileLocation = IIf(IsNull(oField.Value), vbNullString, oField.Value)
                sOldMainEXEInstallFileLocation = Replace(sOldMainEXEInstallFileLocation, "{InstallDir}", goUtil.gsInstallDir, , , vbTextCompare)
                sMainEXEInstallFileLocation = sFieldValue
                sMainEXEInstallFileLocation = Replace(sMainEXEInstallFileLocation, "{InstallDir}", goUtil.gsInstallDir, , , vbTextCompare)
            Case UCase("MainEXESPName")
                sOldMainEXESPName = IIf(IsNull(oField.Value), vbNullString, oField.Value)
                sMainEXESPName = sFieldValue
            Case UCase("MainFTPEXEInstallFileLocation")
                sOldMainFTPEXEInstallFileLocation = IIf(IsNull(oField.Value), vbNullString, oField.Value)
                sOldMainFTPEXEInstallFileLocation = Replace(sOldMainFTPEXEInstallFileLocation, "{InstallDir}", goUtil.gsInstallDir, , , vbTextCompare)
                sMainFTPEXEInstallFileLocation = sFieldValue
                sMainFTPEXEInstallFileLocation = Replace(sMainFTPEXEInstallFileLocation, "{InstallDir}", goUtil.gsInstallDir, , , vbTextCompare)
            Case UCase("MainFTPEXESPName")
                sOldMainFTPEXESPName = IIf(IsNull(oField.Value), vbNullString, oField.Value)
                sMainFTPEXESPName = sFieldValue
            Case UCase("DateLastUpdated")
                dtOldVersionDate = IIf(IsNull(oField.Value), NULL_DATE, oField.Value)
                If IsDate(sFieldValue) Then
                    dtVersionDate = CDate(sFieldValue)
                End If
            Case UCase("UpdateByUserID")
                lOldUpdateByUserID = IIf(IsNull(oField.Value), 0, oField.Value)
                lUpdateByUserID = CLng(sFieldValue)
        End Select
        lFieldsCount = lFieldsCount + 1
    Next

    'Check to see if the Version has changed Or the SP For the DB is Missing
    If (lOldVersion <> lVersion) Or (Not goUtil.utFileExists(msSPPath & "DataBase\SP\" & sSPName)) Then
        CheckDBVerison = False
        'Need to inform Server That Database Update is required,
        'So server can end the Process Tokin
        UploadClientFlag psUserNameDecoded, "DBUpToDate.flag", False
        
        If lOldVersion <> lVersion Then
            lstProcess.AddItem "Database Version Changed from VS " & CStr(lOldVersion) & " (" & CStr(dtOldVersionDate) & ") To VS " & CStr(lVersion) & " (" & CStr(dtVersionDate) & ")"
        ElseIf Not goUtil.utFileExists(msSPPath & "DataBase\SP\" & sSPName) Then
            lstProcess.AddItem "Database Service Pack Missing!"
        End If
        lstProcess.AddItem "Downloading Database Service Pack!  Please Wait..."
        mbSingleFileProcess = True
        'Get Database SP
        EC_FTP.ChangeDir "\.\"
        EC_FTP.ChangeDir msSPFolders & "\DataBase\SP\"
        EC_FTP.GetFile sSPName, msSPPath & "DataBase\SP\" & sSPName
        'The reason the Main Util object is included in the DataBase SP
        'Is because the Version information is Corroborated inside the Main
        'utility object inside clsUtil |Public Property Get ECMainDBVersion
        'Get Main Util SP
        EC_FTP.ChangeDir "\.\"
        EC_FTP.ChangeDir msSPFolders & "\Application\SP\"
        EC_FTP.GetFile sMainUtilSPName, msSPPath & "Application\SP\" & sMainUtilSPName
        'The Reason why the ARV Object, Main EXE and FTP SP are needed with the Database SP
        'Is to be sure that these EXE match Compatibility with the Main Utility Object.
        'If it is necessary to break compatibility then these EXE's must be installed
        'directly after the Main Utility installation.
        'get Main ARV Object SP
        EC_FTP.ChangeDir "\.\"
        EC_FTP.ChangeDir msSPFolders & "\Application\SP\"
        EC_FTP.GetFile sMainARVSPName, msSPPath & "Application\SP\" & sMainARVSPName
        'Get Main EXE SP
        EC_FTP.ChangeDir "\.\"
        EC_FTP.ChangeDir msSPFolders & "\Application\SP\"
        EC_FTP.GetFile sMainEXESPName, msSPPath & "Application\SP\" & sMainEXESPName
        'Get Main FTP EXE SP
        EC_FTP.ChangeDir "\.\"
        EC_FTP.ChangeDir msSPFolders & "\Application\SP\"
        EC_FTP.GetFile sMainFTPEXESPName, msSPPath & "Application\SP\" & sMainFTPEXESPName
        
        mbSingleFileProcess = False
        msDBSPName = sSPName
        msMainUtilSPName = sMainUtilSPName
        msMainARVSPName = sMainARVSPName
        msMainEXESPName = sMainEXESPName
        msMainFTPEXESPName = sMainFTPEXESPName
        pbDownloadDBSPSuccess = True
    Else
        CheckDBVerison = True
    End If

    'cleanup
    Set RS = Nothing
    Set oConn = Nothing
    Set oField = Nothing
    
    Exit Function
EH:
    mbSingleFileProcess = False
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Function CheckDBVerison", False
End Function


Public Function UpdateLookUpInfo() As Boolean
    'This function retruns true if SuccessFull Update of Look Up Info
    On Error GoTo EH
    Dim sData As String
    Dim saryRecords() As String
    Dim lRecordsCount As Long
    Dim saryFields() As String
    Dim lFieldsCount As Long
    Dim sFieldValue As String
    Dim sSQL As String
    Dim oField As DAO.Field
    Dim dicDLFiles As Scripting.Dictionary
    Dim vDLFile As Variant
    Dim sDLFile As String
    Dim sTableName As String
    Dim oTableDef As DAO.TableDef
    Dim oConn As ADODB.Connection
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    Set oConn = New ADODB.Connection
    
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    'Get List Of Down Load Look up files
    
    Set dicDLFiles = New Scripting.Dictionary
    sDLFile = Dir(msFTPDLPath & "*.dllu", vbNormal)
    
    Do Until sDLFile = vbNullString
        dicDLFiles.Add sDLFile, sDLFile
        sDLFile = Dir
    Loop
    lblProcess.Caption = "Updating Look Up Tables. Please Wait! ..."
    lblProcess.Refresh
    lstProcess.AddItem lblProcess.Caption & " " & Now()
    'Need to inform Easy Claim Application that FTP is Updating Tables
    'this will show a modal Screen that will not allow the
    'user to make any changes while the data update is taking place
    Sleep 1000
    SaveSetting "EasyClaim", "MSG", "COMMAND", "SHOW_FTP_IS_UPDATING_DATA"
'    Me.Visible = True
    Me.Refresh
    DoEvents
    Sleep 1000
    'Update Each Table
    For Each vDLFile In dicDLFiles
        sDLFile = vDLFile
        sTableName = Replace(sDLFile, ".dllu", vbNullString, , , vbTextCompare)
        'Set The Table Def
        Set oTableDef = goUtil.gMainDB.TableDefs(sTableName)
        'First Need to Remove All Records From Table
        sSQL = "DELETE * FROM " & sTableName & " "
        
        
        oConn.Execute sSQL
        
        'set table records to array
        sData = goUtil.utGetFileData(msFTPDLPath & sDLFile)
        'Now that we got it need to remove the .dl file
        goUtil.utDeleteFile msFTPDLPath & sDLFile
        
        saryRecords() = Split(sData, RECORD_DELIM, , vbBinaryCompare)
        
        For lRecordsCount = LBound(saryRecords, 1) To UBound(saryRecords, 1)
            sData = saryRecords(lRecordsCount)
            If sData <> vbNullString Then
                saryFields() = Split(sData, COLUMN_DELIM, , vbBinaryCompare)
                lFieldsCount = 0
                
                'Build Update SQL
                sSQL = "INSERT INTO " & sTableName & " "
                sSQL = sSQL & "SELECT "
                For Each oField In oTableDef.Fields
                    If lFieldsCount > 0 Then
                        'Add Comma to start new Column
                        sSQL = sSQL & ", "
                    End If
                    sFieldValue = saryFields(lFieldsCount)
                    sFieldValue = Replace(sFieldValue, COLUMN_DELIM_REP, COLUMN_DELIM, , , vbBinaryCompare)
                    sFieldValue = Replace(sFieldValue, RECORD_DELIM_REP, RECORD_DELIM, , , vbBinaryCompare)
                                 
                    'Snag the UserName and Users ID from Users Table Update
                    If StrComp(sTableName, "USERS", vbTextCompare) = 0 Then
                        If msUserName = vbNullString Or msUsersID = vbNullString Then
                            If StrComp(oField.Name, "UsersID", vbTextCompare) = 0 Then
                                msUsersID = sFieldValue
                            End If
                            If StrComp(oField.Name, "UserName", vbTextCompare) = 0 Then
                                msUserName = sFieldValue
                            End If
                        End If
                    End If
                    
                    If StrComp(sFieldValue, "IS_NULL", vbTextCompare) = 0 Then
                        sSQL = sSQL & "null "
                    Else
                         Select Case oField.Type
                            Case DAO.DataTypeEnum.dbBigInt
                                sSQL = sSQL & sFieldValue
                            Case DAO.DataTypeEnum.dbBinary
                                sFieldValue = Replace(sFieldValue, "'", "''", , , vbBinaryCompare)
                                If sFieldValue = vbNullString Then
                                    sFieldValue = " "
                                End If
                                sSQL = sSQL & "'" & sFieldValue & "'"
                            Case DAO.DataTypeEnum.dbBoolean
                                sSQL = sSQL & sFieldValue
                            Case DAO.DataTypeEnum.dbByte
                                sSQL = sSQL & sFieldValue
                            Case DAO.DataTypeEnum.dbChar
                                sFieldValue = Replace(sFieldValue, "'", "''", , , vbBinaryCompare)
                                If sFieldValue = vbNullString Then
                                    sFieldValue = " "
                                End If
                                sSQL = sSQL & "'" & sFieldValue & "'"
                            Case DAO.DataTypeEnum.dbCurrency
                                sSQL = sSQL & sFieldValue
                            Case DAO.DataTypeEnum.dbDate
                                sSQL = sSQL & "#" & sFieldValue & "#"
                            Case DAO.DataTypeEnum.dbDecimal
                                sSQL = sSQL & sFieldValue
                            Case DAO.DataTypeEnum.dbDouble
                                sSQL = sSQL & sFieldValue
                            Case DAO.DataTypeEnum.dbFloat
                                sSQL = sSQL & sFieldValue
                            Case DAO.DataTypeEnum.dbGUID
                                sFieldValue = Replace(sFieldValue, "'", "''", , , vbBinaryCompare)
                                If sFieldValue = vbNullString Then
                                    sFieldValue = " "
                                End If
                                sSQL = sSQL & "'" & sFieldValue & "'"
                            Case DAO.DataTypeEnum.dbInteger
                                sSQL = sSQL & sFieldValue
                            Case DAO.DataTypeEnum.dbLong
                                sSQL = sSQL & sFieldValue
                            Case DAO.DataTypeEnum.dbLongBinary
                                sSQL = sSQL & sFieldValue
                            Case DAO.DataTypeEnum.dbMemo
                                sFieldValue = Replace(sFieldValue, "'", "''", , , vbBinaryCompare)
                                If sFieldValue = vbNullString Then
                                    sFieldValue = " "
                                End If
                                sSQL = sSQL & "'" & sFieldValue & "'"
                            Case DAO.DataTypeEnum.dbNumeric
                                sSQL = sSQL & sFieldValue
                            Case DAO.DataTypeEnum.dbSingle
                                sSQL = sSQL & sFieldValue
                            Case DAO.DataTypeEnum.dbText
                                sFieldValue = Replace(sFieldValue, "'", "''", , , vbBinaryCompare)
                                If sFieldValue = vbNullString Then
                                    sFieldValue = " "
                                End If
                                sSQL = sSQL & "'" & sFieldValue & "'"
                            Case DAO.DataTypeEnum.dbTime
                                sSQL = sSQL & "#" & sFieldValue & "#"
                            Case DAO.DataTypeEnum.dbTimeStamp
                                sSQL = sSQL & "#" & sFieldValue & "#"
                            Case DAO.DataTypeEnum.dbVarBinary
                                sFieldValue = Replace(sFieldValue, "'", "''", , , vbBinaryCompare)
                                If sFieldValue = vbNullString Then
                                    sFieldValue = " "
                                End If
                                sSQL = sSQL & "'" & sFieldValue & "'"
                        End Select
                    End If
                    sSQL = sSQL & " As [" & oField.Name & "] "
                    lFieldsCount = lFieldsCount + 1
                Next 'oField In oTableDef.Fields
                'Insert this Record Into the Table
                oConn.Execute sSQL
            End If
            lblProcess.Caption = "Updated " & sTableName & " Table... Record (" & lRecordsCount & ") Of (" & UBound(saryRecords, 1) & ")..."
            lblProcess.Refresh
        Next 'lRecordsCount
    Next 'vDLFile In dicDLFiles
    
    'Once the Data is Finished Updating Need to remove the Update Data
    'Screen From Easy Claim
    Sleep 1000
    SaveSetting "EasyClaim", "MSG", "COMMAND", "HIDE_FTP_IS_UPDATING_DATA"
    Sleep 1000
    
    'cleanup
    Set dicDLFiles = Nothing
    Set oTableDef = Nothing
    Set oField = Nothing
    Set oConn = Nothing
    
    
    UpdateLookUpInfo = True
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    'Log the error but don't show it since its already shown above
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function UpdateLookUpInfo", False
    
    lstProcess.AddItem "ERROR #" & lErrNum
    lstProcess.AddItem sErrDesc
    
End Function

Public Function UpdateSoftwareChanges(psUserNameDecoded As String) As Boolean
    'This function retruns true if ...
    '1. There are Software Updates
    'This function returns False if
    '2. There are no Software Updates Required
    
    On Error GoTo EH
    Dim sApplicationSPPath As String
    Dim sDocumentSPPath As String
    Dim sRegSettingSPPath As String
    Dim sECSPLexePath As String
    Dim sECFTPListPath As String
    Dim sECFTPListData As String  'used to build the data for ECFTP.lst File
    Dim sMyCommandStr As String
    Dim sSQL As String
    Dim RS As ADODB.Recordset
    Dim oConn As ADODB.Connection
    Dim sInstallSPPath As String
    Dim sInstallLogPath As String
    Dim sInstallLogData As String
    Dim bFlagInstall As Boolean
    Dim sDescription As String
    Dim sSPName As String
    Dim bFirstRecord As Boolean
    Dim bSentClientFlag As Boolean
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    Set oConn = New ADODB.Connection
    Set RS = New ADODB.Recordset
    
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    
    'Build the SP Directory Paths
    sApplicationSPPath = msSPPath & "Application\SP\"
    sDocumentSPPath = msSPPath & "Document\SP\"
    sRegSettingSPPath = msSPPath & "RegSetting\SP\"
    sECFTPListPath = goUtil.gsInstallDir & "\" & goUtil.utGetTickCount & "_ECFTP.lst"
    sECSPLexePath = goUtil.gsInstallDir & "\ECSPL.exe"
    sMyCommandStr = " " & sECFTPListPath
    
    'init the string
    sECFTPListData = vbNullString
    
    'Update Message
    lblProcess.Caption = "Checking For Software Updates. Please Wait! ..."
    lblProcess.Refresh
    lstProcess.AddItem lblProcess.Caption & " " & Now()
    
    'Need to Look Up All Applicable Software , Application, Document and RegSettings
    'Build the List And Execute the List using ECSPL.exe
    'Get the RegSettings
    sSQL = "SELECT SPName, "
    sSQL = sSQL & "Description, "
    sSQL = sSQL & "VersionDate "
    sSQL = sSQL & "FROM     RegSetting "
    sSQL = sSQL & "WHERE    RegSettingID IN ( "
                            sSQL = sSQL & "SELECT   RegSettingID "
                            sSQL = sSQL & "FROM     SoftwarePackageRegSetting "
                            sSQL = sSQL & "WHERE    SoftWarePackageID IN ( "
                                                            sSQL = sSQL & "SELECT   SoftWarePackageID "
                                                            sSQL = sSQL & "FROM     SoftWarePackage "
                                                            sSQL = sSQL & "WHERE    ClientCompanyID IN ( "
                                                                                        sSQL = sSQL & "SELECT   ClientCompanyID "
                                                                                        sSQL = sSQL & "FROM     ClientCompanyUsersCat "
                                                                                        sSQL = sSQL & ") "
                                                            sSQL = sSQL & "AND      CATID IN ( "
                                                                            sSQL = sSQL & "SELECT   CATID "
                                                                            sSQL = sSQL & "FROM     ClientCompanyUsersCat "
                                                                            sSQL = sSQL & ") "
                                                            sSQL = sSQL & "AND IsDeleted = 0 "
                                                        sSQL = sSQL & ") "
                            sSQL = sSQL & "AND IsDeleted = 0 "
                            sSQL = sSQL & ") "
    sSQL = sSQL & "AND IsDeleted = 0 "
    
    
    'Use Disconnected Record Set on asUseClient Cusor ONLY !
    RS.CursorLocation = adUseClient
    RS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set RS.ActiveConnection = Nothing
    
    If Not RS.EOF Then
        RS.MoveFirst
    End If
    bFlagInstall = False
    bFirstRecord = True
    Do Until RS.EOF
        If Not IsNull(RS!SPName) Then
            sSPName = RS!SPName
            bFlagInstall = False
            'Need to Check for Previous Install log for Reg Setting
            sInstallSPPath = sRegSettingSPPath & RS!SPName
            sInstallLogPath = StrReverse(sRegSettingSPPath & RS!SPName)
            sInstallLogPath = Replace(sInstallLogPath, "exe.", "gol.", , 1, vbTextCompare)
            sInstallLogPath = StrReverse(sInstallLogPath)
            If goUtil.utFileExists(sInstallLogPath) And goUtil.utFileExists(sInstallSPPath) Then
                sInstallLogData = goUtil.utGetFileData(sInstallLogPath)
                If Not IsNull(RS!VersionDate) Then
                    If Format(CDate(sInstallLogData), "YY/MM/DD HH:MM") <> Format(CDate(RS!VersionDate), "YY/MM/DD HH:MM") Then
                        bFlagInstall = True
                    End If
                End If
            Else
                bFlagInstall = True
            End If
            If sECFTPListData <> vbNullString Then
                 If bFlagInstall Then
                    sECFTPListData = sECFTPListData & vbCrLf
                End If
            End If
            If bFlagInstall Then
                If Not bSentClientFlag Then
                    'Need to inform Server that Software Updates Required
                    'And to end the Process Tokin
                    UploadClientFlag psUserNameDecoded, "SoftwareUpToDate.flag", False
                    bSentClientFlag = True
                End If
                'Need to Get the Latest Version from Server
                sDescription = vbNullString
                If Not IsNull(RS!Description) Then
                    sDescription = RS!Description
                End If
                lblProcess.Caption = "Software Update Found. " & "(" & sDescription & ") Please Wait! ..."
                lblProcess.Refresh
                lstProcess.AddItem lblProcess.Caption & " " & Now()
                mbSingleFileProcess = True
                EC_FTP.ChangeDir "\.\"
                EC_FTP.ChangeDir msSPFolders & "\RegSetting\SP\"
                EC_FTP.GetFile sSPName, sRegSettingSPPath & sSPName
                mbSingleFileProcess = False
                sECFTPListData = sECFTPListData & sRegSettingSPPath
                sECFTPListData = sECFTPListData & RS!SPName
            End If
        End If
        bFirstRecord = False
        RS.MoveNext
    Loop
    
    
    'Get Documents
    sSQL = "SELECT SPName, "
    sSQL = sSQL & "Description, "
    sSQL = sSQL & "VersionDate "
    sSQL = sSQL & "FROM     Document "
    sSQL = sSQL & "WHERE    DocumentID IN ( "
                            sSQL = sSQL & "SELECT   DocumentID "
                            sSQL = sSQL & "FROM     SoftwarePackageDocument "
                            sSQL = sSQL & "WHERE    SoftWarePackageID IN ( "
                                                            sSQL = sSQL & "SELECT   SoftWarePackageID "
                                                            sSQL = sSQL & "FROM     SoftWarePackage "
                                                            sSQL = sSQL & "WHERE    ClientCompanyID IN ( "
                                                                                        sSQL = sSQL & "SELECT   ClientCompanyID "
                                                                                        sSQL = sSQL & "FROM     ClientCompanyUsersCat "
                                                                                        sSQL = sSQL & ") "
                                                            sSQL = sSQL & "AND      CATID IN ( "
                                                                            sSQL = sSQL & "SELECT   CATID "
                                                                            sSQL = sSQL & "FROM     ClientCompanyUsersCat "
                                                                            sSQL = sSQL & ") "
                                                            sSQL = sSQL & "AND IsDeleted = 0 "
                                                        sSQL = sSQL & ") "
                            sSQL = sSQL & "AND IsDeleted = 0 "
                            sSQL = sSQL & ") "
    sSQL = sSQL & "AND IsDeleted = 0 "
    
    Set RS = Nothing
    Set RS = New ADODB.Recordset
     'Use Disconnected Record Set on asUseClient Cusor ONLY !
    RS.CursorLocation = adUseClient
    RS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set RS.ActiveConnection = Nothing
    
    If Not RS.EOF Then
        RS.MoveFirst
    End If
    bFlagInstall = False
    bFirstRecord = True
    Do Until RS.EOF
        If Not IsNull(RS!SPName) Then
            sSPName = RS!SPName
            bFlagInstall = False
            'Need to Check for Previous Install log for Document
            sInstallSPPath = sDocumentSPPath & RS!SPName
            sInstallLogPath = StrReverse(sDocumentSPPath & RS!SPName)
            sInstallLogPath = Replace(sInstallLogPath, "exe.", "gol.", , 1, vbTextCompare)
            sInstallLogPath = StrReverse(sInstallLogPath)
            If goUtil.utFileExists(sInstallLogPath) And goUtil.utFileExists(sInstallSPPath) Then
                sInstallLogData = goUtil.utGetFileData(sInstallLogPath)
                If Not IsNull(RS!VersionDate) Then
                    If Format(CDate(sInstallLogData), "YY/MM/DD HH:MM") <> Format(CDate(RS!VersionDate), "YY/MM/DD HH:MM") Then
                        bFlagInstall = True
                    End If
                End If
            Else
                bFlagInstall = True
            End If
            If sECFTPListData <> vbNullString Then
                If bFlagInstall Then
                    sECFTPListData = sECFTPListData & vbCrLf
                End If
            End If
            If bFlagInstall Then
                If Not bSentClientFlag Then
                    'Need to inform Server that Software Updates Required
                    'And to end the Process Tokin
                    UploadClientFlag psUserNameDecoded, "SoftwareUpToDate.flag", False
                    bSentClientFlag = True
                End If
                'Need to Get the Latest Version from Server
                sDescription = vbNullString
                If Not IsNull(RS!Description) Then
                    sDescription = RS!Description
                End If
                lblProcess.Caption = "Software Update Found. " & "(" & sDescription & ") Please Wait! ..."
                lblProcess.Refresh
                lstProcess.AddItem lblProcess.Caption & " " & Now()
                mbSingleFileProcess = True
                EC_FTP.ChangeDir "\.\"
                EC_FTP.ChangeDir msSPFolders & "\Document\SP\"
                EC_FTP.GetFile sSPName, sDocumentSPPath & sSPName
                mbSingleFileProcess = False
                If bFirstRecord And sECFTPListData <> vbNullString Then
                    sECFTPListData = sECFTPListData & vbCrLf
                End If
                sECFTPListData = sECFTPListData & sDocumentSPPath
                sECFTPListData = sECFTPListData & RS!SPName
            End If
        End If
        bFirstRecord = False
        RS.MoveNext
    Loop
    
    'Get Application
    sSQL = "SELECT SPName, "
    sSQL = sSQL & "Description, "
    sSQL = sSQL & "VersionDate "
    sSQL = sSQL & "FROM     Application "
    sSQL = sSQL & "WHERE    ApplicationID IN ( "
                            sSQL = sSQL & "SELECT   ApplicationID "
                            sSQL = sSQL & "FROM     SoftwarePackageApplication "
                            sSQL = sSQL & "WHERE    SoftWarePackageID IN ( "
                                                            sSQL = sSQL & "SELECT   SoftWarePackageID "
                                                            sSQL = sSQL & "FROM     SoftWarePackage "
                                                            sSQL = sSQL & "WHERE    ClientCompanyID IN ( "
                                                                                        sSQL = sSQL & "SELECT   ClientCompanyID "
                                                                                        sSQL = sSQL & "FROM     ClientCompanyUsersCat "
                                                                                        sSQL = sSQL & ") "
                                                            sSQL = sSQL & "AND      CATID IN ( "
                                                                            sSQL = sSQL & "SELECT   CATID "
                                                                            sSQL = sSQL & "FROM     ClientCompanyUsersCat "
                                                                            sSQL = sSQL & ") "
                                                            sSQL = sSQL & "AND IsDeleted = 0 "
                                                        sSQL = sSQL & ") "
                            sSQL = sSQL & "AND IsDeleted = 0 "
                            sSQL = sSQL & ") "
    sSQL = sSQL & "AND IsDeleted = 0 "
    
    
    Set RS = Nothing
    Set RS = New ADODB.Recordset
    'Use Disconnected Record Set on asUseClient Cusor ONLY !
    RS.CursorLocation = adUseClient
    RS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set RS.ActiveConnection = Nothing
    
    If Not RS.EOF Then
        RS.MoveFirst
    End If
    bFlagInstall = False
    bFirstRecord = True
    Do Until RS.EOF
        If Not IsNull(RS!SPName) Then
            sSPName = RS!SPName
            bFlagInstall = False
            'Need to Check for Previous Install log for Application
            sInstallSPPath = sApplicationSPPath & RS!SPName
            sInstallLogPath = StrReverse(sApplicationSPPath & RS!SPName)
            sInstallLogPath = Replace(sInstallLogPath, "exe.", "gol.", , 1, vbTextCompare)
            sInstallLogPath = StrReverse(sInstallLogPath)
            If goUtil.utFileExists(sInstallLogPath) And goUtil.utFileExists(sInstallSPPath) Then
                sInstallLogData = goUtil.utGetFileData(sInstallLogPath)
                If Not IsNull(RS!VersionDate) Then
                    If Format(CDate(sInstallLogData), "YY/MM/DD HH:MM") <> Format(CDate(RS!VersionDate), "YY/MM/DD HH:MM") Then
                        bFlagInstall = True
                    End If
                End If
            Else
                bFlagInstall = True
            End If
            If sECFTPListData <> vbNullString Then
                If bFlagInstall Then
                    sECFTPListData = sECFTPListData & vbCrLf
                End If
            End If
            If bFlagInstall Then
                'Do not flag Server to shut off process tokins if Installing ECSPL.exe
                If Not bSentClientFlag And StrComp(sSPName, "ECSPL.exe_V1.exe", vbTextCompare) <> 0 Then
                    'Need to inform Server that Software Updates Required
                    'And to end the Process Tokin
                    UploadClientFlag psUserNameDecoded, "SoftwareUpToDate.flag", False
                    bSentClientFlag = True
                End If
                'Need to Get the Latest Version from Server
                sDescription = vbNullString
                If Not IsNull(RS!Description) Then
                    sDescription = RS!Description
                End If
                lblProcess.Caption = "Software Update Found. " & "(" & sDescription & ") Please Wait! ..."
                lblProcess.Refresh
                lstProcess.AddItem lblProcess.Caption & " " & Now()
                mbSingleFileProcess = True
                EC_FTP.ChangeDir "\.\"
                EC_FTP.ChangeDir msSPFolders & "\Application\SP\"
                EC_FTP.GetFile sSPName, sApplicationSPPath & sSPName
                mbSingleFileProcess = False
                If StrComp(sSPName, "ECSPL.exe_V1.exe", vbTextCompare) = 0 Then
                    'Need to install the ECSPL.EXE Separate and First if it has been
                    'Updated.  This is the Software Package List App that installs
                    'List of updates
                    Shell sApplicationSPPath & sSPName, vbNormalFocus
                    Sleep 1000
                    bFlagInstall = False
                Else
                    If bFirstRecord And sECFTPListData <> vbNullString Then
                        sECFTPListData = sECFTPListData & vbCrLf
                    End If
                    sECFTPListData = sECFTPListData & sApplicationSPPath
                    sECFTPListData = sECFTPListData & RS!SPName
                End If
            End If
        End If
        bFirstRecord = False
        RS.MoveNext
    Loop
    
    If sECFTPListData <> vbNullString Then
        'Need to Save the List Data to Temp Files
        goUtil.utSaveFileData sECFTPListPath, sECFTPListData
        'Execute the List
        'don't worry about cleanup of this file
        'As ECSPL.exe will take care of it
        Shell sECSPLexePath & sMyCommandStr, vbNormalFocus
        UpdateSoftwareChanges = True
    Else
        UpdateSoftwareChanges = False
    End If
    
    'CleanUp
    Set RS = Nothing
    Set oConn = Nothing
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    'Log the error but don't show it since its already shown above
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function UpdateSoftwareChanges", False
    
    lstProcess.AddItem "ERROR #" & lErrNum
    lstProcess.AddItem sErrDesc
    
End Function

Public Function UpdateDownLoadInfo() As Boolean
    'This function retruns true if SuccessFull Update of Download Info
    On Error GoTo EH
    Dim sData As String
    Dim saryRecords() As String
    Dim lRecordsCount As Long
    Dim saryFields() As String
    Dim lFieldsCount As Long
    Dim sFieldValue As String
    Dim sSQL As String
    Dim oField As DAO.Field
    Dim dicDLFiles As Scripting.Dictionary
    Dim vDLFile As Variant
    Dim sDLFile As String
    Dim sTableName As String
    Dim oTableDef As DAO.TableDef
    'Need to Check for Certain Fields in
    'Assignments Table When Updating
    Dim bSkipThisRecord As Boolean
    Dim lPosReassigned As Long
    Dim lPosIsLocked As Long
    Dim lPosIsDeleted As Long
    
    Dim RS As ADODB.Recordset
    Dim oConn As ADODB.Connection
    Dim sSQLWHERE As String
    
    'Update SQL Server Flags and Assignments Status
    Dim oXZip As New V2ECKeyBoard.clsXZip
    Dim sTickCount As String
    Dim sFileName As String
    Dim sFileNameZip As String
    Dim sSQLUPServer As String
    Dim bDoUpdateDownload As Boolean
    Dim sLOCKED_AssignmentsID As String
    Dim sApprovalRequestStatusID As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim lRetry As Long
    
    
    sApprovalRequestStatusID = GetStatusID("APPROVALRequest")
    sLOCKED_AssignmentsID = GetSetting("ECFTP", "MSG", "LOCKED_AssignmentsID", vbNullString)
    
    Set oConn = New ADODB.Connection
    Set RS = New ADODB.Recordset
    
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    
    oXZip.SetUtilObject goUtil
    
    'Get List Of Down Load Look up files
    
    Set dicDLFiles = New Scripting.Dictionary
    sDLFile = Dir(msFTPDLPath & "*.dl", vbNormal)
    
    Do Until sDLFile = vbNullString
        dicDLFiles.Add sDLFile, sDLFile
        sDLFile = Dir
    Loop
    lblProcess.Caption = "Updating Data Tables. Please Wait! ..."
    lblProcess.Refresh
    lstProcess.AddItem lblProcess.Caption & " " & Now()
    'Need to inform Easy Claim Application that FTP is Updating Tables
    'this will show a modal Screen that will not allow the
    'user to make any changes while the data update is taking place
    Sleep 1000
    SaveSetting "EasyClaim", "MSG", "COMMAND", "SHOW_FTP_IS_UPDATING_DATA"
'    Me.Visible = True
    Me.Refresh
    DoEvents
    Sleep 1000
    'Update Each Table
    For Each vDLFile In dicDLFiles
        sDLFile = vDLFile
        sTableName = Left(sDLFile, InStrRev(sDLFile, "_", , vbBinaryCompare) - 1)
        'Set The Table Def
        Set oTableDef = goUtil.gMainDB.TableDefs(sTableName)
        
        'set table records to array
        sData = goUtil.utGetFileData(msFTPDLPath & sDLFile)
        'Now that we got it need to remove the .dl file
        goUtil.utDeleteFile msFTPDLPath & sDLFile
        
        saryRecords() = Split(sData, RECORD_DELIM, , vbBinaryCompare)
        
        'If the Table is Assignments, Need to Get the Pos for
        If StrComp(sTableName, "Assignments", vbTextCompare) = 0 Then
            '1. lPosReassigned
            '2. lPosIsLocked
            '3. lPosIsDeleted
            lFieldsCount = 0
            For Each oField In oTableDef.Fields
                If StrComp(oField.Name, "Reassigned", vbTextCompare) = 0 Then
                    lPosReassigned = lFieldsCount
                ElseIf StrComp(oField.Name, "IsLocked", vbTextCompare) = 0 Then
                    lPosIsLocked = lFieldsCount
                ElseIf StrComp(oField.Name, "IsDeleted", vbTextCompare) = 0 Then
                    lPosIsDeleted = lFieldsCount
                End If
                If lPosReassigned > 0 And lPosIsLocked > 0 And lPosIsDeleted > 0 Then
                    Exit For
                End If
                lFieldsCount = lFieldsCount + 1
            Next
        End If
        'SQL Server Update Statement for Flags and Status(Assignments Table only)
        'Do one update Statement per Table
        sSQLUPServer = "UPDATE " & sTableName & " SET DownLoadMe = 0, "
        Select Case UCase(sTableName)
            Case UCase("Assignments")
                'If the Assignment inserted or being editied is in Pending Status...
                'need to change the status to NEW
                sSQLUPServer = sSQLUPServer & "StatusID = ( "
                sSQLUPServer = sSQLUPServer & "CASE "
                sSQLUPServer = sSQLUPServer & "WHEN StatusID = " & AssgnStatus.iAssignmentsStatus_PENDING & " "
                sSQLUPServer = sSQLUPServer & "THEN " & AssgnStatus.iAssignmentsStatus_NEW & " "
                sSQLUPServer = sSQLUPServer & "ELSE StatusID "
                sSQLUPServer = sSQLUPServer & "END), "
                sSQLUPServer = sSQLUPServer & "DownLoadAll = 0, "
        End Select
        sSQLUPServer = sSQLUPServer & "DateLastUpdated = GetDate(), "
        sSQLUPServer = sSQLUPServer & "UpdateByUserID = " & msUsersID & " "
        sSQLUPServer = sSQLUPServer & "WHERE DownLoadMe = 1 "
        'Check for a locked assignment
        If sLOCKED_AssignmentsID <> vbNullString Then
            sSQLUPServer = sSQLUPServer & "AND [AssignmentsID] <> " & sLOCKED_AssignmentsID & "  "
        End If
        Select Case UCase(sTableName)
            Case UCase("Assignments")
                sSQLUPServer = sSQLUPServer & "AND AdjusterSpecID IN( "
                                        sSQLUPServer = sSQLUPServer & "SELECT   ClientCoAdjusterSpecID "
                                        sSQLUPServer = sSQLUPServer & "FROM     ClientCoAdjusterSpec "
                                        sSQLUPServer = sSQLUPServer & "WHERE    UsersID = " & msUsersID & " "
                                        sSQLUPServer = sSQLUPServer & ") "
            Case Else
                sSQLUPServer = sSQLUPServer & "AND AssignmentsID IN ("
                                        sSQLUPServer = sSQLUPServer & "SELECT   AssignmentsID "
                                        sSQLUPServer = sSQLUPServer & "FROM     Assignments "
                                        sSQLUPServer = sSQLUPServer & "WHERE    AdjusterSpecID IN( "
                                                                sSQLUPServer = sSQLUPServer & "SELECT   ClientCoAdjusterSpecID "
                                                                sSQLUPServer = sSQLUPServer & "FROM     ClientCoAdjusterSpec "
                                                                sSQLUPServer = sSQLUPServer & "WHERE    UsersID = " & msUsersID & " "
                                                                sSQLUPServer = sSQLUPServer & ") "
                                         sSQLUPServer = sSQLUPServer & ") "
        End Select
        
        For lRecordsCount = LBound(saryRecords, 1) To UBound(saryRecords, 1)
            sData = saryRecords(lRecordsCount)
            
            If sData = vbNullString Then
                GoTo SKIP_THIS_RECORD
            End If
            
            saryFields() = Split(sData, COLUMN_DELIM, , vbBinaryCompare)
            lFieldsCount = 0
            
            'Build Update SQL
            sSQL = "SELECT * FROM " & sTableName & " "
            sSQLWHERE = "WHERE "
            For Each oField In oTableDef.Fields
                'Get the Unique ID... All tables have the Unique ID in the very
                'First Column.  In the future this could Change per table
                'So the Code is set up incase of this eventuality
                sFieldValue = saryFields(lFieldsCount)
                sFieldValue = Replace(sFieldValue, COLUMN_DELIM_REP, COLUMN_DELIM, , , vbBinaryCompare)
                sFieldValue = Replace(sFieldValue, RECORD_DELIM_REP, RECORD_DELIM, , , vbBinaryCompare)
                Select Case UCase(sTableName)
                    Case UCase("")
                    Case Else
                        If lFieldsCount > 0 Then
                            Exit For
                        End If
                        sSQLWHERE = sSQLWHERE & "[" & oField.Name & "] = " & sFieldValue & " "
                End Select
                lFieldsCount = lFieldsCount + 1
             Next
  
            sSQL = sSQL & sSQLWHERE
            
            Set RS = Nothing
            Set RS = New ADODB.Recordset
            'Use Disconnected Record Set on asUseClient Cusor ONLY !
            RS.CursorLocation = adUseClient
            RS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
            Set RS.ActiveConnection = Nothing
             
                       
             'If the record ID does not exist , then Insert the Download Record
             If RS.RecordCount = 0 Then
                lFieldsCount = 0
                'Build Update SQL
                sSQL = "INSERT INTO " & sTableName & " "
                sSQL = sSQL & "SELECT "
                
                For Each oField In oTableDef.Fields
                    If lFieldsCount > 0 Then
                        'Add Comma to start new Column
                        sSQL = sSQL & ", "
                    End If
                    
                    sFieldValue = saryFields(lFieldsCount)
                    sFieldValue = Replace(sFieldValue, COLUMN_DELIM_REP, COLUMN_DELIM, , , vbBinaryCompare)
                    sFieldValue = Replace(sFieldValue, RECORD_DELIM_REP, RECORD_DELIM, , , vbBinaryCompare)
                    
                    'Check for Locked Assignment
                    If sLOCKED_AssignmentsID <> vbNullString Then
                        If StrComp(oField.Name, "AssignmentsID", vbTextCompare) = 0 Then
                            If StrComp(sFieldValue, sLOCKED_AssignmentsID, vbTextCompare) = 0 Then
                                GoTo SKIP_THIS_RECORD
                            End If
                        End If
                    End If
                    
                    'Modify Field values Here if applicable
                    Select Case UCase(sTableName)
                        Case UCase("Assignments")
                            Select Case UCase(oField.Name)
                                Case UCase("StatusID")
                                    If IsNumeric(sFieldValue) Then
                                        Select Case CLng(sFieldValue)
                                            'If the Assignment inserted is in Pending Status...
                                            'need to change the status to NEW
                                            Case AssgnStatus.iAssignmentsStatus_PENDING
                                                sFieldValue = AssgnStatus.iAssignmentsStatus_NEW
                                        End Select
                                    End If
                            End Select
                    End Select
                    
                    If StrComp(sFieldValue, "IS_NULL", vbTextCompare) = 0 Then
                        sSQL = sSQL & "null "
                    Else
                         Select Case oField.Type
                            Case DAO.DataTypeEnum.dbBigInt
                                sSQL = sSQL & sFieldValue
                            Case DAO.DataTypeEnum.dbBinary
                                sFieldValue = Replace(sFieldValue, "'", "''", , , vbBinaryCompare)
                                If sFieldValue = vbNullString Then
                                    sFieldValue = " "
                                End If
                                sSQL = sSQL & "'" & sFieldValue & "'"
                            Case DAO.DataTypeEnum.dbBoolean
                                'Reset DownLoadMe Flags
                                If StrComp(oField.Name, "DownLoadMe", vbTextCompare) = 0 Then
                                    sSQL = sSQL & " 0"
                                ElseIf StrComp(oField.Name, "DownLoadAll", vbTextCompare) = 0 Then
                                    sSQL = sSQL & " 0"
                                Else
                                    sSQL = sSQL & sFieldValue
                                End If
                            Case DAO.DataTypeEnum.dbByte
                                sSQL = sSQL & sFieldValue
                            Case DAO.DataTypeEnum.dbChar
                                sFieldValue = Replace(sFieldValue, "'", "''", , , vbBinaryCompare)
                                If sFieldValue = vbNullString Then
                                    sFieldValue = " "
                                End If
                                sSQL = sSQL & "'" & sFieldValue & "'"
                            Case DAO.DataTypeEnum.dbCurrency
                                sSQL = sSQL & sFieldValue
                            Case DAO.DataTypeEnum.dbDate
                                sSQL = sSQL & "#" & sFieldValue & "#"
                            Case DAO.DataTypeEnum.dbDecimal
                                sSQL = sSQL & sFieldValue
                            Case DAO.DataTypeEnum.dbDouble
                                sSQL = sSQL & sFieldValue
                            Case DAO.DataTypeEnum.dbFloat
                                sSQL = sSQL & sFieldValue
                            Case DAO.DataTypeEnum.dbGUID
                                sFieldValue = Replace(sFieldValue, "'", "''", , , vbBinaryCompare)
                                If sFieldValue = vbNullString Then
                                    sFieldValue = " "
                                End If
                                sSQL = sSQL & "'" & sFieldValue & "'"
                            Case DAO.DataTypeEnum.dbInteger
                                sSQL = sSQL & sFieldValue
                            Case DAO.DataTypeEnum.dbLong
                                sSQL = sSQL & sFieldValue
                            Case DAO.DataTypeEnum.dbLongBinary
                                sSQL = sSQL & sFieldValue
                            Case DAO.DataTypeEnum.dbMemo
                                sFieldValue = Replace(sFieldValue, "'", "''", , , vbBinaryCompare)
                                If sFieldValue = vbNullString Then
                                    sFieldValue = " "
                                End If
                                sSQL = sSQL & "'" & sFieldValue & "'"
                            Case DAO.DataTypeEnum.dbNumeric
                                sSQL = sSQL & sFieldValue
                            Case DAO.DataTypeEnum.dbSingle
                                sSQL = sSQL & sFieldValue
                            Case DAO.DataTypeEnum.dbText
                                sFieldValue = Replace(sFieldValue, "'", "''", , , vbBinaryCompare)
                                If sFieldValue = vbNullString Then
                                    sFieldValue = " "
                                End If
                                sSQL = sSQL & "'" & sFieldValue & "'"
                            Case DAO.DataTypeEnum.dbTime
                                sSQL = sSQL & "#" & sFieldValue & "#"
                            Case DAO.DataTypeEnum.dbTimeStamp
                                sSQL = sSQL & "#" & sFieldValue & "#"
                            Case DAO.DataTypeEnum.dbVarBinary
                                sFieldValue = Replace(sFieldValue, "'", "''", , , vbBinaryCompare)
                                If sFieldValue = vbNullString Then
                                    sFieldValue = " "
                                End If
                                sSQL = sSQL & "'" & sFieldValue & "'"

                            Case Else
                                sSQL = sSQL & sFieldValue
                        End Select
                    End If
                    sSQL = sSQL & " As [" & oField.Name & "] "
                    lFieldsCount = lFieldsCount + 1
                Next 'oField In oTableDef.Fields
                'Insert this Record Into the Table
                
                oConn.Execute sSQL
                
             End If
             
             'If the Record Does Exists then Update the Download record
             If RS.RecordCount = 1 Then
                
                If CBool(RS!UpLoadMe) Then
                    'Do not Allow Updates to Records that are marked For Upload.
                    'If the client (Access Record) is currently working a record and made changes to it,
                    'The client (Access Record) Wins.
                    '1. An exception to this rule is If this is the Assignments Table....
                    'And this Reocord is Marked for Deletion, Locking, Or Reassignment.
                    '10.6.2005 And low and behold one more cock sucker makes an exception...
                    'OR IF THE CURRENT STATUS ON THE ASSIGNMETS TABLE IS APPROVAL REQUEST
                    bSkipThisRecord = True
                    Select Case UCase(sTableName)
                        Case UCase("Assignments")
                            'Reassigned
                            sFieldValue = saryFields(lPosReassigned)
                            If StrComp(sFieldValue, "IS_NULL", vbTextCompare) <> 0 Then
                                If CBool(sFieldValue) Then
                                    bSkipThisRecord = False
                                End If
                            End If
                            'IsLocked
                            sFieldValue = saryFields(lPosIsLocked)
                            If StrComp(sFieldValue, "IS_NULL", vbTextCompare) <> 0 Then
                                If CBool(sFieldValue) Then
                                    bSkipThisRecord = False
                                End If
                            End If
                            'IsDeleted
                            sFieldValue = saryFields(lPosIsDeleted)
                            If StrComp(sFieldValue, "IS_NULL", vbTextCompare) <> 0 Then
                                If CBool(sFieldValue) Then
                                    bSkipThisRecord = False
                                End If
                            End If
                            
                            '10.6.2005 Check the current Assignmentrecord status
                            'If it is Approval Request then need to allow this updae from the server
                            If CStr(RS!StatusID) = sApprovalRequestStatusID Then
                                bSkipThisRecord = False
                            End If
                        Case Else
                            bSkipThisRecord = True
                    End Select
                    'Skip it if flagged
                    If bSkipThisRecord Then
                        GoTo SKIP_THIS_RECORD
                    End If
                End If
                lFieldsCount = 0
                'Build Update SQL
                sSQL = "UPDATE " & sTableName & " SET "
                For Each oField In oTableDef.Fields
                    If lFieldsCount > 0 Then
                        'Add Comma to start new Column
                        sSQL = sSQL & ", "
                    End If
                    sSQL = sSQL & "[" & oField.Name & "] = "
                    sFieldValue = saryFields(lFieldsCount)
                    sFieldValue = Replace(sFieldValue, COLUMN_DELIM_REP, COLUMN_DELIM, , , vbBinaryCompare)
                    sFieldValue = Replace(sFieldValue, RECORD_DELIM_REP, RECORD_DELIM, , , vbBinaryCompare)
                    
                    'Check for Locked Assignment
                    If sLOCKED_AssignmentsID <> vbNullString Then
                        If StrComp(oField.Name, "AssignmentsID", vbTextCompare) = 0 Then
                            If StrComp(sFieldValue, sLOCKED_AssignmentsID, vbTextCompare) = 0 Then
                                GoTo SKIP_THIS_RECORD
                            End If
                        End If
                    End If
                    
                    'Modify Field values Here if applicable
                    Select Case UCase(sTableName)
                        Case UCase("Assignments")
                            Select Case UCase(oField.Name)
                                Case UCase("StatusID")
                                    If IsNumeric(sFieldValue) Then
                                        Select Case CLng(sFieldValue)
                                            'If the Assignment being editied is in Pending Status...
                                            'need to change the status to NEW
                                            Case AssgnStatus.iAssignmentsStatus_PENDING
                                                sFieldValue = AssgnStatus.iAssignmentsStatus_NEW
                                        End Select
                                    End If
                            End Select
                    End Select
                    
                    If StrComp(sFieldValue, "IS_NULL", vbTextCompare) = 0 Then
                        sSQL = sSQL & "null"
                    Else
                         Select Case oField.Type
                            Case DAO.DataTypeEnum.dbBigInt
                                sSQL = sSQL & sFieldValue
                            Case DAO.DataTypeEnum.dbBinary
                                sFieldValue = Replace(sFieldValue, "'", "''", , , vbBinaryCompare)
                                If sFieldValue = vbNullString Then
                                    sFieldValue = " "
                                End If
                                sSQL = sSQL & "'" & sFieldValue & "'"
                            Case DAO.DataTypeEnum.dbBoolean
                                'Reset DownLoadMe Flags
                                If StrComp(oField.Name, "DownLoadMe", vbTextCompare) = 0 Then
                                    sSQL = sSQL & " 0"
                                ElseIf StrComp(oField.Name, "DownLoadAll", vbTextCompare) = 0 Then
                                    sSQL = sSQL & " 0"
                                Else
                                    sSQL = sSQL & sFieldValue
                                End If
                            Case DAO.DataTypeEnum.dbByte
                                sSQL = sSQL & sFieldValue
                            Case DAO.DataTypeEnum.dbChar
                                sFieldValue = Replace(sFieldValue, "'", "''", , , vbBinaryCompare)
                                If sFieldValue = vbNullString Then
                                    sFieldValue = " "
                                End If
                                sSQL = sSQL & "'" & sFieldValue & "'"
                            Case DAO.DataTypeEnum.dbCurrency
                                sSQL = sSQL & sFieldValue
                            Case DAO.DataTypeEnum.dbDate
                                sSQL = sSQL & "#" & sFieldValue & "#"
                            Case DAO.DataTypeEnum.dbDecimal
                                sSQL = sSQL & sFieldValue
                            Case DAO.DataTypeEnum.dbDouble
                                sSQL = sSQL & sFieldValue
                            Case DAO.DataTypeEnum.dbFloat
                                sSQL = sSQL & sFieldValue
                            Case DAO.DataTypeEnum.dbGUID
                                sFieldValue = Replace(sFieldValue, "'", "''", , , vbBinaryCompare)
                                If sFieldValue = vbNullString Then
                                    sFieldValue = " "
                                End If
                                sSQL = sSQL & "'" & sFieldValue & "'"
                            Case DAO.DataTypeEnum.dbInteger
                                sSQL = sSQL & sFieldValue
                            Case DAO.DataTypeEnum.dbLong
                                sSQL = sSQL & sFieldValue
                            Case DAO.DataTypeEnum.dbLongBinary
                                sSQL = sSQL & sFieldValue
                            Case DAO.DataTypeEnum.dbMemo
                                sFieldValue = Replace(sFieldValue, "'", "''", , , vbBinaryCompare)
                                If sFieldValue = vbNullString Then
                                    sFieldValue = " "
                                End If
                                sSQL = sSQL & "'" & sFieldValue & "'"
                            Case DAO.DataTypeEnum.dbNumeric
                                sSQL = sSQL & sFieldValue
                            Case DAO.DataTypeEnum.dbSingle
                                sSQL = sSQL & sFieldValue
                            Case DAO.DataTypeEnum.dbText
                                sFieldValue = Replace(sFieldValue, "'", "''", , , vbBinaryCompare)
                                If sFieldValue = vbNullString Then
                                    sFieldValue = " "
                                End If
                                sSQL = sSQL & "'" & sFieldValue & "'"
                            Case DAO.DataTypeEnum.dbTime
                                sSQL = sSQL & "#" & sFieldValue & "#"
                            Case DAO.DataTypeEnum.dbTimeStamp
                                sSQL = sSQL & "#" & sFieldValue & "#"
                            Case DAO.DataTypeEnum.dbVarBinary
                                sFieldValue = Replace(sFieldValue, "'", "''", , , vbBinaryCompare)
                                If sFieldValue = vbNullString Then
                                    sFieldValue = " "
                                End If
                                sSQL = sSQL & "'" & sFieldValue & "'"
                            Case Else
                                sSQL = sSQL & sFieldValue
                        End Select
                    End If
                    lFieldsCount = lFieldsCount + 1
                Next 'oField In oTableDef.Fields
                'Insert this Record Into the Table
                sSQL = sSQL & " " & sSQLWHERE
                
                oConn.Execute sSQL

             End If

SKIP_THIS_RECORD:
        lblProcess.Caption = "Updated " & sTableName & " Table... Record (" & lRecordsCount & ") Of (" & UBound(saryRecords, 1) & ")..."
        lblProcess.Refresh
        Next 'lRecordsCount
        '-------------------Update SQL SERVER FLAGS and STATUS------
        'Save the Data to Upload Folder
        sTickCount = goUtil.utGetTickCount
        sFileName = "UPDATE_" & sTableName & "_" & sTickCount & ".ulud"
        sFileNameZip = "UPDATE_" & sTableName & "_" & sTickCount & ".zulud"

        goUtil.utSaveFileData msFTPULPath & sFileName, sSQLUPServer
        oXZip.SaveZIPFiles msFTPULPath, sFileNameZip, sFileName, goUtil.DB_PASSWORD("1")
        SetAttr msFTPULPath & sFileNameZip, vbNormal
        bDoUpdateDownload = True
        'Change FTP Dr and Update the Server Side UPdate
        mbSingleFileProcess = True
        On Error Resume Next
        EC_FTP.ChangeDir "\.\"
        EC_FTP.ChangeDir msUserFolders & "\USER_FOLDERS\" & msUserName & "\"
        lRetry = 0
RETRY:
        DoEvents
        Sleep 100
        EC_FTP.PutFile msFTPULPath & sFileNameZip, sFileNameZip
        If Err.Number <> 0 Then
            lRetry = lRetry + 1
            If lRetry <= 10 Then
                Sleep 500
                Err.Clear
                On Error Resume Next
                GoTo RETRY
            End If
            lstProcess.AddItem "ERROR #" & Err.Number & " File(" & sFileNameZip & ") "
            lstProcess.AddItem Err.Description
            Err.Clear
        End If
        goUtil.utDeleteFile msFTPULPath & sFileNameZip
        mbSingleFileProcess = False
        On Error GoTo EH
        '---------------------Update SQL SERVER FLAGS and STATUS------
    Next 'vDLFile In dicDLFiles
    

    
    'cleanup
    Set oXZip = Nothing
    Set dicDLFiles = Nothing
    Set oTableDef = Nothing
    Set oField = Nothing
    Set RS = Nothing
    Set oConn = Nothing
    
    'Download Attahcs and Photos that are flaged for Download
    If bDoUpdateDownload Then
        UploadClientFlag msUserName, "UploadUpdateReady.flag", True
        Sleep 1000
    End If
    
    UpdateDownLoadInfo = True
    
    'Once the Data is Finished Updating Need to remove the Update Data
    'Screen From Easy Claim
    Sleep 1000
    SaveSetting "EasyClaim", "MSG", "COMMAND", "HIDE_FTP_IS_UPDATING_DATA"
    Sleep 1000
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    'Log the error but don't show it since its already shown above
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function UpdateDownLoadInfo", False
    
    lstProcess.AddItem "ERROR #" & lErrNum
    lstProcess.AddItem sErrDesc
    
End Function

Public Function UpdateUpLoadInfo(pvAryToken As Variant) As Boolean
    'This function returns true if SuccessFull Update of Upload Info
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    Dim sMess As String
    Dim sProdDSN As String
    Dim sTableRSData As String
    Dim sTickCount As String
    Dim sFileName As String
    Dim sFileNameZip As String
    Dim sULFilePath As String
    Dim sTableName As String
    Dim saryDBLookUpTables(1 To 200) As String
    Dim lCountDBLookupTables As Long
    Dim saryDBTables(1 To 500) As String
    Dim lCountDBTables As Long
    Dim oXZip As V2ECKeyBoard.clsXZip
    Dim sDLFile As String
    Dim sListDLFile As String
    Dim lCheckCount As Long
    Dim bDoUpload As Boolean
    Dim sLOCKED_AssignmentsID As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    sLOCKED_AssignmentsID = GetSetting("ECFTP", "MSG", "LOCKED_AssignmentsID", vbNullString)
    
    
    lblProcess.Caption = "Checking for Data to Upload..."
    lblProcess.Refresh
    lstProcess.AddItem lblProcess.Caption & " " & Now()
    'Need to inform Easy Claim Application that FTP is Updating Tables
    'this will show a modal Screen that will not allow the
    'user to make any changes while the data update is taking place
    Sleep 1000
    SaveSetting "EasyClaim", "MSG", "COMMAND", "SHOW_FTP_IS_UPDATING_DATA"
'    Me.Visible = True
    Me.Refresh
    DoEvents
    Sleep 1000
    
    Set oConn = New ADODB.Connection
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    Set oXZip = New V2ECKeyBoard.clsXZip
    
    'B. Need to Create DL files that Easy Claim will Download to Update Look up
    'Info specific to User and User Assigned Cat Info.  This will include Software Information.
    'Add the Look up tables to the lookup table array
    saryDBTables(1) = "Assignments"
    saryDBTables(2) = "BillingCount"
    saryDBTables(3) = "PolicyLimits"
    saryDBTables(4) = "RTIB"
    saryDBTables(5) = "RTIBFee"
    saryDBTables(6) = "IB"
    saryDBTables(7) = "IBFee"
    saryDBTables(8) = "RTChecks"
    saryDBTables(9) = "RTIndemnity"
    saryDBTables(10) = "RTActivityLog"
    saryDBTables(11) = "RTActivityLogInfo"
    saryDBTables(12) = "RTPhotoReport"
    saryDBTables(13) = "RTPhotoLog"
    saryDBTables(14) = "RTWSDiagram"
    saryDBTables(15) = "RTAttachments"
    saryDBTables(16) = "Package"
    saryDBTables(17) = "PackageItem"
    '2/8/2004 MiscReportParam , and MiscReportParam01 to MiscReportParam30
    saryDBTables(18) = "MiscReportParam"
    saryDBTables(19) = "MiscReportParam01"
    saryDBTables(20) = "MiscReportParam02"
    saryDBTables(21) = "MiscReportParam03"
    saryDBTables(22) = "MiscReportParam04"
    saryDBTables(23) = "MiscReportParam05"
    saryDBTables(24) = "MiscReportParam06"
    saryDBTables(25) = "MiscReportParam07"
    saryDBTables(26) = "MiscReportParam08"
    saryDBTables(27) = "MiscReportParam09"
    saryDBTables(28) = "MiscReportParam10"
    saryDBTables(29) = "MiscReportParam11"
    saryDBTables(30) = "MiscReportParam12"
    saryDBTables(31) = "MiscReportParam13"
    saryDBTables(32) = "MiscReportParam14"
    saryDBTables(33) = "MiscReportParam15"
    saryDBTables(34) = "MiscReportParam16"
    saryDBTables(35) = "MiscReportParam17"
    saryDBTables(36) = "MiscReportParam18"
    saryDBTables(37) = "MiscReportParam19"
    saryDBTables(38) = "MiscReportParam20"
    saryDBTables(39) = "MiscReportParam21"
    saryDBTables(40) = "MiscReportParam22"
    saryDBTables(41) = "MiscReportParam23"
    saryDBTables(42) = "MiscReportParam24"
    saryDBTables(43) = "MiscReportParam25"
    saryDBTables(44) = "MiscReportParam26"
    saryDBTables(45) = "MiscReportParam27"
    saryDBTables(46) = "MiscReportParam28"
    saryDBTables(47) = "MiscReportParam29"
    saryDBTables(48) = "MiscReportParam30"
    'B. Get DataBase Updates Here
     
    'Use Tick Count on Non Look up info
    sTickCount = goUtil.utGetTickCount
    For lCountDBTables = LBound(saryDBTables, 1) To UBound(saryDBTables, 1)
        sTableName = saryDBTables(lCountDBTables)
        
        If sTableName = vbNullString Then
            Exit For
        End If
        
        sSQL = "SELECT * FROM " & sTableName & " "
        
        'Build the Where Statement
        Select Case sTableName
            Case "Assignments"
                sSQL = sSQL & "WHERE AdjusterSpecID IN( "
                                        sSQL = sSQL & "SELECT   ClientCoAdjusterSpecID "
                                        sSQL = sSQL & "FROM     ClientCoAdjusterSpec "
                                        sSQL = sSQL & "WHERE    UsersID = " & msUsersID & " "
                                        sSQL = sSQL & ") "
                sSQL = sSQL & "AND UpLoadMe = True "
                
            Case Else
                '"RTIB", "RTIBFee", "IB", "IBFee", "BillingCount",
                '"PolicyLimits", "Package", "PackageItem",
                '"RTChecks", "RTIndemnity", "RTActivityLog",
                '"RTActivityLogInfo", "RTPhotoLog", "RTPhotoReport",
                '"RTWSDiagram", "RTAttachments", "RTFarmerNCC",
                '2/8/2004 MiscReportParam , and MiscReportParam01 to MiscReportParam30
                '"MiscReportParam"
                sSQL = sSQL & "WHERE AssignmentsID IN ("
                                        sSQL = sSQL & "SELECT   AssignmentsID "
                                        sSQL = sSQL & "FROM     Assignments "
                                        sSQL = sSQL & "WHERE    AdjusterSpecID IN( "
                                                                sSQL = sSQL & "SELECT   ClientCoAdjusterSpecID "
                                                                sSQL = sSQL & "FROM     ClientCoAdjusterSpec "
                                                                sSQL = sSQL & "WHERE    UsersID = " & msUsersID & " "
                                                                sSQL = sSQL & ") "
                                         sSQL = sSQL & ") "
                sSQL = sSQL & "AND UpLoadMe = True "
                
        End Select
        
        'Check for a locked assignment
        If sLOCKED_AssignmentsID <> vbNullString Then
             sSQL = sSQL & "AND [AssignmentsID] <> " & sLOCKED_AssignmentsID & "  "
        End If
        
        sFileName = Format(lCountDBTables, "000") & "_" & sTableName & "_" & sTickCount & ".ul"
        sFileNameZip = Format(lCountDBTables, "000") & "_" & sTableName & "_" & sTickCount & ".zul"
        sULFilePath = msFTPULPath
        sTableRSData = GetExportULTableRS(sSQL, oConn)
        
        'If there is No data don't create Download
        If sTableRSData <> vbNullString Then
            bDoUpload = True
            lblProcess.Caption = "Upload Data found: " & sTableName & " Table(" & lCountDBTables & ") "
            lblProcess.Refresh
            lstProcess.AddItem lblProcess.Caption & " " & Now()
            goUtil.utSaveFileData sULFilePath & sFileName, sTableRSData
            oXZip.SaveZIPFiles sULFilePath, sFileNameZip, sFileName, goUtil.DB_PASSWORD("1")
            SetAttr sULFilePath & sFileNameZip, vbNormal
        End If
    Next
    
    If bDoUpload Then
        lblProcess.Caption = "Uploading Data Please Wait!!! "
        lblProcess.Refresh
        lstProcess.AddItem lblProcess.Caption & " " & Now()
        'Need FTP the Upload files
        UploadFiles "*_" & sTickCount & ".zul"
        UploadClientFlag msUserName, "UploadReady.flag", True
        Sleep 1000
        'Synchronize Unique IDs for all tables
        If Not VerifySecurity(pvAryToken, True, "Waiting for Synchronize Data file ", 120) Then
            lstProcess.AddItem "Server not responding!  Please try again later. "
            lblMess.Caption = "ERROR CONNECT AGAIN!"
            GoTo CLEAN_UP
        End If
        
        'Need to Synchronize data
        If SynchronizeULData() Then
            UpdateUpLoadInfo = True
        End If
    Else
        lblProcess.Caption = "No Data found to upload. "
        lblProcess.Refresh
        lstProcess.AddItem lblProcess.Caption & " " & Now()
        UpdateUpLoadInfo = True
    End If
    
CLEAN_UP:
    'cleanup
    Set oConn = Nothing
    Set oXZip = Nothing
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    'Log the error but don't show it since its already shown above
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function UpdateUpLoadInfo", False
    
    lstProcess.AddItem "ERROR #" & lErrNum
    lstProcess.AddItem sErrDesc
    
End Function

Public Function SynchronizeULData() As Boolean
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim dicDLFiles As Scripting.Dictionary
    Dim oXZip As V2ECKeyBoard.clsXZip
    Dim sDLFile As String
    Dim vDLFile As Variant
    Dim sSQL As String
    Dim sUpdateData As String
    Dim saryUpdateData() As String
    Dim lPos As Long
    Dim lRecordsAffected As Long
    Dim lRecordsAffectedTotal As Long
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    Set mdicRemoteFiles = Nothing
    'This Process will get the File list
    'From server... As Well it will fire the
    'EC_FTP_DirItem event.  inside this event
    'anything the on the server that fits the pattern
    'will be populated into mdicRemoteFiles dictionary
    'Then the EC_FTP_Done event will fire and will Download
    'and then remove the files in the mdicRemoteFiles dictionary object.
    EC_FTP.Pattern = "SYNCH_*.zsyc"
    EC_FTP.GetFilenameList
    
    'Once the files have been downloaded, they need to be unzipped...
    
    '1. Unzip All the DL files
    sDLFile = Dir(msFTPDLPath & "SYNCH_*.zsyc", vbNormal)
    
    If sDLFile = vbNullString Then
        GoTo CLEAN_UP
    End If
    
    lblProcess.Caption = "Synchronizing Data Please Wait!!!"
    lblProcess.Refresh
    lstProcess.AddItem lblProcess.Caption & " " & Now()
    
    Set dicDLFiles = New Scripting.Dictionary
    
    Do Until sDLFile = vbNullString
        dicDLFiles.Add sDLFile, sDLFile
        sDLFile = Dir
    Loop
    'Need to loop through the zipped Synch FIles
    'and unzip multiple docs inside
    Set oXZip = New V2ECKeyBoard.clsXZip
    For Each vDLFile In dicDLFiles
        sDLFile = vDLFile
        oXZip.UNZipFiles msFTPDLPath, msFTPDLPath & sDLFile, False
        Sleep 100
        goUtil.utDeleteFile msFTPDLPath & sDLFile
    Next
    
    'Now that all the files have been unzipped need to process each one
    
    Set dicDLFiles = New Scripting.Dictionary
    sDLFile = Dir(msFTPDLPath & "SYNCH_*.syc", vbNormal)
    
    Do Until sDLFile = vbNullString
        dicDLFiles.Add sDLFile, sDLFile
        sDLFile = Dir
    Loop
    
    'Need to loop throug and process Synchronize Update Statements
    Set oConn = New ADODB.Connection
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    For Each vDLFile In dicDLFiles
        sDLFile = vDLFile
        sUpdateData = goUtil.utGetFileData(msFTPDLPath & sDLFile)
        'Remove the File from server
        goUtil.utDeleteFile msFTPDLPath & sDLFile
        Erase saryUpdateData()
        saryUpdateData() = Split(sUpdateData, RECORD_DELIM, , vbBinaryCompare)
        lRecordsAffectedTotal = 0
        For lPos = LBound(saryUpdateData, 1) To UBound(saryUpdateData, 1)
            sSQL = saryUpdateData(lPos)
            If sSQL <> vbNullString Then
                oConn.Execute sSQL, lRecordsAffected
                lRecordsAffectedTotal = lRecordsAffectedTotal + lRecordsAffected
            End If
        Next
        lstProcess.AddItem sDLFile & " " & lRecordsAffectedTotal & " Record(s) Affected. " & Now()
        lstProcess.Refresh
        Me.Refresh
    Next
    
    
CLEAN_UP:

    SynchronizeULData = True
    Set oXZip = Nothing
    Set dicDLFiles = Nothing
    Set oConn = Nothing
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
     'Log the error but don't show it since its already shown above
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function SynchronizeULData", False
    
    lstProcess.AddItem "ERROR #" & lErrNum
    lstProcess.AddItem sErrDesc
   
End Function

Private Function GetExportULTableRS(psSQL As String, poConn As ADODB.Connection) As String
    On Error GoTo EH
    Dim sMess As String
    Dim RS As ADODB.Recordset
    Dim oField As ADODB.Field
    Dim sFieldData As String
    Dim sRecords As String
    Dim sTempData As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    Set RS = New ADODB.Recordset
    'Use Disconnected Record Set on asUseClient Cusor ONLY !
    
    RS.CursorLocation = adUseClient
    RS.Open psSQL, poConn, adOpenForwardOnly, adLockReadOnly
    Set RS.ActiveConnection = Nothing
    
    If Not RS.EOF Then
        RS.MoveFirst
        sRecords = vbNullString
        Do Until RS.EOF
            sFieldData = vbNullString
            sTempData = vbNullString
            For Each oField In RS.Fields
                'Clear previous Value
                If IsNull(oField.Value) Then
                    sFieldData = sFieldData & "IS_NULL"
                Else
                    Select Case oField.Type
                        Case ADODB.DataTypeEnum.adWChar, ADODB.DataTypeEnum.adVarWChar, ADODB.DataTypeEnum.adVarChar, ADODB.DataTypeEnum.adLongVarWChar, ADODB.DataTypeEnum.adLongVarChar, ADODB.DataTypeEnum.adChar, ADODB.DataTypeEnum.adBSTR
                            sTempData = CStr(oField.Value)
                            sTempData = Replace(sTempData, COLUMN_DELIM, COLUMN_DELIM_REP, , , vbBinaryCompare)
                            sTempData = Replace(sTempData, RECORD_DELIM, RECORD_DELIM_REP, , , vbBinaryCompare)
                            'Check for Nullstring
                            If sTempData = vbNullString Then
                                sTempData = " "
                            End If
                            sFieldData = sFieldData & sTempData
                        Case Else
                            sFieldData = sFieldData & CStr(oField.Value)
                    End Select
                End If
                sFieldData = sFieldData & COLUMN_DELIM
            Next
            RS.MoveNext
            sFieldData = sFieldData & RECORD_DELIM
            sRecords = sRecords & sFieldData
        Loop
    End If
    
    GetExportULTableRS = sRecords
    
    'cleanup
    Set RS = Nothing
    
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    'Log the error but don't show it since its already shown above
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Function GetExportULTableRS", False
    
    lstProcess.AddItem "ERROR #" & lErrNum
    lstProcess.AddItem sErrDesc
    
End Function


Private Function UpdatePhotoAttatchmentDownload() As Boolean
    On Error GoTo EH
    Dim bForceDownload As Boolean
    Dim sMess As String
    Dim lWaitAWhile As Long
    Dim lHwnd As Long
    Dim dblRet As Double
    Dim RS As ADODB.Recordset
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    Dim sSQLWHERE As String
    Dim sSQLAndNotIN As String
    Dim sSQLBuildNotIn As String
    Dim sTableName As String
    Dim saryDBTables(1 To 4) As String
    Dim lCountDBTables As Long
    Dim sLRFormat As String
    Dim sDLFile As String
    Dim lDLFileCount As Long
    Dim sYYMM As String
    Dim sDD As String
    Dim saryTemp() As String
    'Photos
    Dim sPhotoName As String
    Dim sPhotoHighResName As String
    Dim sPhotoThumbName As String
    'For Rebuilding Thumb if it does not exist on the Server
    'When flagged for download
    Dim sCmdStr As String
    Dim sIrfanViewPath As String
    Dim Optimal_H As Long
    Dim Optimal_W As Long
    'Used by Photos or Diagram Photo
    Dim sDownloadFlag As String
    
    'Update Server to reset Download Flags
    Dim sTickCount As String
    Dim bRecordsFound As Boolean
    Dim sResetDLFlagsData As String
    Dim sFileName As String
    Dim sFileNameZip As String
    Dim oXZip As V2ECKeyBoard.clsXZip
    Dim bDoUpdateDownload  As Boolean
    Dim sLOCKED_AssignmentsID As String
    'If Is Deleted don't try to download it
    Dim bIsDeleted As Boolean
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim lRetry As Long
    
    sLOCKED_AssignmentsID = GetSetting("ECFTP", "MSG", "LOCKED_AssignmentsID", vbNullString)
    
    'Flaged if Attachments and Photos will be overwitten even if they already
    'Exist on the client machine.
    bForceDownload = GetSetting(App.EXEName, "CONNECTION_SETTINGS", "ForceDownload", False)
    
    Set RS = New ADODB.Recordset
    Set oConn = New ADODB.Connection
    
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    saryDBTables(1) = "Assignments"
    saryDBTables(2) = "RTPhotoLog"
    saryDBTables(3) = "RTAttachments"
    saryDBTables(4) = "RTWSDiagram"
    
    sIrfanViewPath = goUtil.utGetSystemDir & "\ECS\DLL\i_view32.exe"
    Optimal_W = V2ECKeyBoard.PhotoSettings.Optimal_W
    Optimal_H = V2ECKeyBoard.PhotoSettings.Optimal_H
    Optimal_W = goUtil.ConvertTwipsToPixels(Optimal_W)
    Optimal_H = goUtil.ConvertTwipsToPixels(Optimal_H)
    
    For lCountDBTables = LBound(saryDBTables, 1) To UBound(saryDBTables, 1)
        sTableName = saryDBTables(lCountDBTables)
        If sTableName <> vbNullString Then
            sSQL = "SELECT * FROM " & sTableName & " "
            sSQLWHERE = "WHERE "
            Select Case UCase(sTableName)
                Case UCase("Assignments")
                    sSQLWHERE = sSQLWHERE & "DownLoadLossReport = True "
                    sSQLWHERE = sSQLWHERE & "AND AdjusterSpecID IN( "
                                            sSQLWHERE = sSQLWHERE & "SELECT   ClientCoAdjusterSpecID "
                                            sSQLWHERE = sSQLWHERE & "FROM     ClientCoAdjusterSpec "
                                            sSQLWHERE = sSQLWHERE & "WHERE    UsersID = " & msUsersID & " "
                                            sSQLWHERE = sSQLWHERE & ") "
                    sSQLAndNotIN = "AND [AssignmentsID] Not IN (***) "
                    sSQLBuildNotIn = vbNullString
                Case UCase("RTPhotoLog")
                    sSQLWHERE = sSQLWHERE & "( "
                    sSQLWHERE = sSQLWHERE & "DownloadPhoto = True "
                    sSQLWHERE = sSQLWHERE & "OR "
                    sSQLWHERE = sSQLWHERE & "DownloadPhotoThumb = True "
                    sSQLWHERE = sSQLWHERE & "OR "
                    sSQLWHERE = sSQLWHERE & "DownloadPhotoHighRes = True "
                    sSQLWHERE = sSQLWHERE & ") "
                    sSQLWHERE = sSQLWHERE & "AND AssignmentsID IN ("
                                            sSQLWHERE = sSQLWHERE & "SELECT   AssignmentsID "
                                            sSQLWHERE = sSQLWHERE & "FROM     Assignments "
                                            sSQLWHERE = sSQLWHERE & "WHERE    AdjusterSpecID IN( "
                                                                    sSQLWHERE = sSQLWHERE & "SELECT   ClientCoAdjusterSpecID "
                                                                    sSQLWHERE = sSQLWHERE & "FROM     ClientCoAdjusterSpec "
                                                                    sSQLWHERE = sSQLWHERE & "WHERE    UsersID = " & msUsersID & " "
                                                                    sSQLWHERE = sSQLWHERE & ") "
                                             sSQLWHERE = sSQLWHERE & ") "
                    sSQLAndNotIN = "AND [RTPhotoLogID] Not IN (***) "
                    sSQLBuildNotIn = vbNullString
                Case UCase("RTAttachments")
                    sSQLWHERE = sSQLWHERE & "DownloadAttachment = True "
                    sSQLWHERE = sSQLWHERE & "AND AssignmentsID IN ("
                                            sSQLWHERE = sSQLWHERE & "SELECT   AssignmentsID "
                                            sSQLWHERE = sSQLWHERE & "FROM     Assignments "
                                            sSQLWHERE = sSQLWHERE & "WHERE    AdjusterSpecID IN( "
                                                                    sSQLWHERE = sSQLWHERE & "SELECT   ClientCoAdjusterSpecID "
                                                                    sSQLWHERE = sSQLWHERE & "FROM     ClientCoAdjusterSpec "
                                                                    sSQLWHERE = sSQLWHERE & "WHERE    UsersID = " & msUsersID & " "
                                                                    sSQLWHERE = sSQLWHERE & ") "
                                             sSQLWHERE = sSQLWHERE & ") "
                    sSQLAndNotIN = "AND [RTAttachmentsID] Not IN (***) "
                    sSQLBuildNotIn = vbNullString
                Case UCase("RTWSDiagram")
                    sSQLWHERE = sSQLWHERE & "DownloadDiagramPhoto = True "
                    sSQLWHERE = sSQLWHERE & "AND AssignmentsID IN ("
                                            sSQLWHERE = sSQLWHERE & "SELECT   AssignmentsID "
                                            sSQLWHERE = sSQLWHERE & "FROM     Assignments "
                                            sSQLWHERE = sSQLWHERE & "WHERE    AdjusterSpecID IN( "
                                                                    sSQLWHERE = sSQLWHERE & "SELECT   ClientCoAdjusterSpecID "
                                                                    sSQLWHERE = sSQLWHERE & "FROM     ClientCoAdjusterSpec "
                                                                    sSQLWHERE = sSQLWHERE & "WHERE    UsersID = " & msUsersID & " "
                                                                    sSQLWHERE = sSQLWHERE & ") "
                                             sSQLWHERE = sSQLWHERE & ") "
                    sSQLAndNotIN = "AND [RTWSDiagramID] Not IN (***) "
                    sSQLBuildNotIn = vbNullString
            End Select
            
            'Check for a locked assignment
            If sLOCKED_AssignmentsID <> vbNullString Then
                sSQLWHERE = sSQLWHERE & "AND [AssignmentsID] <> " & sLOCKED_AssignmentsID & "  "
            End If
            
            'Add the Where
            sSQL = sSQL & sSQLWHERE
            
            Set RS = New ADODB.Recordset
            'Use Disconnected Record Set on asUseClient Cusor ONLY !
            RS.CursorLocation = adUseClient
            RS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
            Set RS.ActiveConnection = Nothing
            
            lDLFileCount = 0
            PBAttachDownload.Value = 0
            PBPhotoDownLoad.Value = 0
            bRecordsFound = False
            If Not RS.EOF Then
                'Unhide the cmdCancelPhotoAttachUL button "Finis Later"
                cmdCancelPhotoAttachUL.Visible = True
                cmdCancelPhotoAttachUL.Enabled = True
                'Reset the Cancel Flag
                mbCancelPhotoAttach = False
                
                bRecordsFound = True
                Select Case UCase(sTableName)
                    Case UCase("Assignments"), UCase("RTAttachments")
                        PBAttachDownload.Max = RS.RecordCount
                    Case UCase("RTPhotoLog"), UCase("RTWSDiagram")
                        PBPhotoDownLoad.Max = RS.RecordCount
                End Select
                Do Until RS.EOF
                    lDLFileCount = lDLFileCount + 1
                    Select Case UCase(sTableName)
                        Case UCase("Assignments"), UCase("RTAttachments")
                            PBAttachDownload.Value = lDLFileCount
                            If mbCancelPhotoAttach Then
                                lblProcess.Caption = "Canceling Download Attachments... (" & PBAttachDownload.Max & ") Item(s)  Please Wait! ...  Item (" & PBAttachDownload.Value & ") Of (" & PBAttachDownload.Max & ") "
                            Else
                                lblProcess.Caption = "Downloading Attachments... (" & PBAttachDownload.Max & ") Item(s)  Please Wait! ...  Item (" & PBAttachDownload.Value & ") Of (" & PBAttachDownload.Max & ") "
                            End If
                            lblProcess.Refresh
                        Case UCase("RTPhotoLog"), UCase("RTWSDiagram")
                            PBPhotoDownLoad.Value = lDLFileCount
                            If mbCancelPhotoAttach Then
                                lblProcess.Caption = "Canceling  Download Photos... (" & PBPhotoDownLoad.Max & ") Item(s)  Please Wait! ...  Item (" & PBPhotoDownLoad.Value & ") Of (" & PBPhotoDownLoad.Max & ") "
                            Else
                                lblProcess.Caption = "Downloading Photos... (" & PBPhotoDownLoad.Max & ") Item(s)  Please Wait! ...  Item (" & PBPhotoDownLoad.Value & ") Of (" & PBPhotoDownLoad.Max & ") "
                            End If
                            lblProcess.Refresh
                    End Select
                    bIsDeleted = IIf(IsNull(RS!IsDeleted), False, RS!IsDeleted)
                    If bIsDeleted Then
                        GoTo SKIP_DELETED
                    End If
                    Select Case UCase(sTableName)
                        Case UCase("Assignments")
                            sLRFormat = vbNullString
                            sDLFile = vbNullString
                            sLRFormat = IIf(IsNull(RS!LRFormat), vbNullString, RS!LRFormat)
                            sDLFile = IIf(IsNull(RS!LossReport), vbNullString, RS!LossReport)
                            sDLFile = Trim(sDLFile)
                            If InStr(1, sLRFormat, "OLEType_pdf", vbTextCompare) > 0 And sDLFile <> vbNullString Then
                                'Have to Build the Path to the File
                                'Accounting for YYMM and DD Folder Structure
                                'This is How the Files are stored on the Server.
                                'the Client will just store the files in
                                'AttachRepos
                                'PhotoRepos
                                'Without the YYMM and DD Sub Folders
                                saryTemp = Split(sDLFile, "_")
                                sYYMM = saryTemp(1)
                                sDD = saryTemp(1)
                                sYYMM = Left(sYYMM, 4)
                                sDD = Mid(sDD, 5, 2)
                                
                                'If there is an error Downloading this File Log it
                                'But Continue On to other files
                                'first check to see if the file already exists
                                'on the client. if it does then don't bother
                                'downloading it again.
                                If Not goUtil.utFileExists(msAttachReposPath & sDLFile) Or bForceDownload Then
                                    If mbCancelPhotoAttach Then
                                        GoTo CANCEL_DOWNLOAD1
                                    End If
                                    On Error Resume Next
                                    mbSingleFileProcess = True
                                    EC_FTP.ChangeDir "\.\"
                                    EC_FTP.ChangeDir msUserFolders & "\Upload\AttachRepos\" & sYYMM & "\" & sDD & "\"
                                    If Err.Number = 0 Then
                                        EC_FTP.GetFile sDLFile, msAttachReposPath & sDLFile
                                        DoEvents
                                        Sleep 10
                                        If goUtil.utGetFileData(msAttachReposPath & sDLFile) = vbNullString Then
                                            For lWaitAWhile = 1 To 10
                                                If goUtil.utGetFileData(msAttachReposPath & sDLFile) = vbNullString Then
                                                    DoEvents
                                                    Sleep 500
                                                    EC_FTP.GetFile sDLFile, msAttachReposPath & sDLFile
                                                Else
                                                    Exit For
                                                End If
                                            Next
                                        End If
                                    End If
                                    mbSingleFileProcess = False
                                    'Check for Errors
                                    If Err.Number <> 0 Then
                                        lstProcess.AddItem "ERROR #" & Err.Number & " File(" & sDLFile & ") "
                                        lstProcess.AddItem Err.Description
                                        goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Function UpdatePhotoAttatchmentDownload", False
                                        Err.Clear
CANCEL_DOWNLOAD1:
                                        If sSQLBuildNotIn = vbNullString Then
                                            sSQLBuildNotIn = sSQLBuildNotIn & RS.Fields("AssignmentsID")
                                        Else
                                            sSQLBuildNotIn = sSQLBuildNotIn & ", " & RS.Fields("AssignmentsID")
                                        End If
                                    End If
                                    On Error GoTo 0
                                    On Error GoTo EH
                                End If
                            End If
                            
                        Case UCase("RTAttachments")
                            sDLFile = vbNullString
                            sDLFile = IIf(IsNull(RS!Attachment), vbNullString, RS!Attachment)
                            sDLFile = Trim(sDLFile)
                            If sDLFile <> vbNullString Then
                                'Have to Build the Path to the File
                                'Accounting for YYMM and DD Folder Structure
                                'This is How the Files are stored on the Server.
                                'the Client will just store the files in
                                'AttachRepos
                                'PhotoRepos
                                'Without the YYMM and DD Sub Folders
                                saryTemp = Split(sDLFile, "_")
                                sYYMM = saryTemp(1)
                                sDD = saryTemp(1)
                                sYYMM = Left(sYYMM, 4)
                                sDD = Mid(sDD, 5, 2)
                                
                                'If there is an error Downloading this File Log it
                                'But Continue On to other files
                                If Not goUtil.utFileExists(msAttachReposPath & sDLFile) Or bForceDownload Then
                                    If mbCancelPhotoAttach Then
                                        GoTo CANCEL_DOWNLOAD2
                                    End If
                                    On Error Resume Next
                                    mbSingleFileProcess = True
                                    EC_FTP.ChangeDir "\.\"
                                    EC_FTP.ChangeDir msUserFolders & "\Upload\AttachRepos\" & sYYMM & "\" & sDD & "\"
                                    If Err.Number = 0 Then
                                        EC_FTP.GetFile sDLFile, msAttachReposPath & sDLFile
                                        DoEvents
                                        Sleep 10
                                        If goUtil.utGetFileData(msAttachReposPath & sDLFile) = vbNullString Then
                                            For lWaitAWhile = 1 To 10
                                                If goUtil.utGetFileData(msAttachReposPath & sDLFile) = vbNullString Then
                                                    DoEvents
                                                    Sleep 500
                                                    EC_FTP.GetFile sDLFile, msAttachReposPath & sDLFile
                                                Else
                                                    Exit For
                                                End If
                                            Next
                                        End If
                                    End If
                                    mbSingleFileProcess = False
                                    'Check for Errors
                                    If Err.Number <> 0 Then
                                        lstProcess.AddItem "ERROR #" & Err.Number & " File(" & sDLFile & ") "
                                        lstProcess.AddItem Err.Description
                                        goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Function UpdatePhotoAttatchmentDownload", False
                                        Err.Clear
CANCEL_DOWNLOAD2:
                                        If sSQLBuildNotIn = vbNullString Then
                                            sSQLBuildNotIn = sSQLBuildNotIn & RS.Fields("RTAttachmentsID")
                                        Else
                                            sSQLBuildNotIn = sSQLBuildNotIn & ", " & RS.Fields("RTAttachmentsID")
                                        End If
                                    End If
                                    On Error GoTo 0
                                    On Error GoTo EH
                                End If
                            End If
                        Case UCase("RTPhotoLog"), UCase("RTWSDiagram")
                            sDLFile = vbNullString
                            If StrComp(sTableName, "RTPhotoLog", vbTextCompare) = 0 Then
                                sDLFile = IIf(IsNull(RS!PhotoName), vbNullString, RS!PhotoName)
                                sDownloadFlag = "DownloadPhoto"
                            ElseIf StrComp(sTableName, "RTWSDiagram", vbTextCompare) = 0 Then
                                sDLFile = IIf(IsNull(RS!DiagramPhotoName), vbNullString, RS!DiagramPhotoName)
                                sDownloadFlag = "DownloadDiagramPhoto"
                            End If
                            sDLFile = Trim(sDLFile)
                            sPhotoName = vbNullString
                            If sDLFile <> vbNullString Then
                                'Have to Build the Path to the File
                                'Accounting for YYMM and DD Folder Structure
                                'This is How the Files are stored on the Server.
                                'the Client will just store the files in
                                'AttachRepos
                                'PhotoRepos
                                'Without the YYMM and DD Sub Folders
                                saryTemp = Split(sDLFile, "_")
                                sYYMM = saryTemp(1)
                                sDD = saryTemp(1)
                                sYYMM = Left(sYYMM, 4)
                                sDD = Mid(sDD, 5, 2)
                                
                                If mbCancelPhotoAttach Then
                                    GoTo CANCEL_DOWNLOAD3
                                End If
                                
                                'If there is an error Downloading this File Log it
                                'But Continue On to other files
                                mbSingleFileProcess = True
                                'Check for Each Type OF Photo Download...
                                If CBool(RS.Fields(sDownloadFlag)) Then
                                     sPhotoName = sDLFile
                                    'DownloadPhoto
                                    If Not goUtil.utFileExists(msPhotoReposPath & sDLFile) Or bForceDownload Then
                                        On Error Resume Next
                                        EC_FTP.ChangeDir "\.\"
                                        EC_FTP.ChangeDir msUserFolders & "\Upload\PhotoRepos\" & sYYMM & "\" & sDD & "\"
                                        If Err.Number = 0 Then
                                            EC_FTP.GetFile sDLFile, msPhotoReposPath & sDLFile
                                            DoEvents
                                            Sleep 10
                                            If goUtil.utGetFileData(msPhotoReposPath & sDLFile) = vbNullString Then
                                                For lWaitAWhile = 1 To 10
                                                    If goUtil.utGetFileData(msPhotoReposPath & sDLFile) = vbNullString Then
                                                        DoEvents
                                                        Sleep 500
                                                        EC_FTP.GetFile sDLFile, msPhotoReposPath & sDLFile
                                                    Else
                                                        Exit For
                                                    End If
                                                Next
                                            End If
                                        End If
                                    End If
                                End If
                                If StrComp(sDownloadFlag, "DownloadDiagramPhoto", vbTextCompare) = 0 Then
                                    GoTo SKIP_DIAGRAM
                                End If
                                If Err.Number = 0 Then
                                    On Error GoTo 0
                                    On Error GoTo EH
                                    If CBool(RS!DownloadPhotoHighRes) Then
                                        sDLFile = Replace(sDLFile, "_1.jpg", "_0.jpg", , , vbTextCompare)
                                        sPhotoHighResName = sDLFile
                                        'DownloadPhotoHighRes
                                        If Not goUtil.utFileExists(msPhotoReposPath & sDLFile) Or bForceDownload Then
                                            On Error Resume Next
                                            EC_FTP.ChangeDir "\.\"
                                            EC_FTP.ChangeDir msUserFolders & "\Upload\PhotoRepos\" & sYYMM & "\" & sDD & "\"
                                            If Err.Number = 0 Then
                                                EC_FTP.GetFile sDLFile, msPhotoReposPath & sDLFile
                                                DoEvents
                                                Sleep 10
                                                If Err.Number <> 0 Or goUtil.utGetFileData(msPhotoReposPath & sDLFile) = vbNullString Then
                                                    Err.Clear
                                                    On Error GoTo 0
                                                    On Error GoTo EH
                                                    'Need to Rebuild the High Res Photo
                                                    'Check for Download of Photo Use the Low Res for High
                                                    If goUtil.utFileExists(msPhotoReposPath & sPhotoName) Then
                                                        'if the mian photo is zero length give it some time
                                                        'to be created. 5 seconds
                                                        For lWaitAWhile = 1 To 10
                                                            If goUtil.utGetFileData(msPhotoReposPath & sPhotoName) = vbNullString Then
                                                                DoEvents
                                                                Sleep 500
                                                            Else
                                                                Exit For
                                                            End If
                                                        Next
                                                        sMess = goUtil.utCopyFile(msPhotoReposPath & sPhotoName, msPhotoReposPath & sPhotoHighResName)
                                                        DoEvents
                                                        Sleep 10
                                                        If sMess <> vbNullString Then
                                                            lstProcess.AddItem "ERROR " & sMess & " File(" & sDLFile & ") "
                                                        End If
                                                        For lWaitAWhile = 1 To 10
                                                            If goUtil.utGetFileData(msPhotoReposPath & sPhotoHighResName) = vbNullString Then
                                                                DoEvents
                                                                Sleep 500
                                                            Else
                                                                Exit For
                                                            End If
                                                        Next
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                    'Be sure to create the Thumb last since it will be using the main
                                    'File.  don't want to Copy this file while processing it by irfanview
                                    If CBool(RS!DownloadPhotoThumb) Then
                                        If CBool(RS!DownloadPhotoHighRes) Then
                                            sDLFile = Replace(sDLFile, "_0.jpg", "_2.jpg", , , vbTextCompare)
                                        Else
                                            sDLFile = Replace(sDLFile, "_1.jpg", "_2.jpg", , , vbTextCompare)
                                        End If
                                        sPhotoThumbName = sDLFile
                                        'DownloadPhotoThumb
                                        If Not goUtil.utFileExists(msPhotoReposPath & sDLFile) Or bForceDownload Then
                                            On Error Resume Next
                                            EC_FTP.ChangeDir "\.\"
                                            EC_FTP.ChangeDir msUserFolders & "\Upload\PhotoRepos\" & sYYMM & "\" & sDD & "\"
                                            If Err.Number = 0 Then
                                                EC_FTP.GetFile sDLFile, msPhotoReposPath & sDLFile
                                                DoEvents
                                                Sleep 10
                                                If Err.Number <> 0 Or goUtil.utGetFileData(msPhotoReposPath & sDLFile) = vbNullString Then
                                                    'Need to Rebuild the Photo Thumb
                                                    Err.Clear
                                                    On Error GoTo 0
                                                    On Error GoTo EH
                                                    'if the main photo is zero length give it some time
                                                    'to be created. 5 seconds
                                                    For lWaitAWhile = 1 To 10
                                                        If goUtil.utGetFileData(msPhotoReposPath & sPhotoName) = vbNullString Then
                                                            DoEvents
                                                            Sleep 500
                                                        Else
                                                            Exit For
                                                        End If
                                                    Next
                                                    'if the irfanview window is found then wait until it is not
                                                    sCmdStr = """" & sIrfanViewPath & """ "
                                                    sCmdStr = sCmdStr & LCase(msPhotoReposPath & sPhotoName) & " "
                                                    'Make the Thumb 20% of the Optimal Pixels width
                                                    sCmdStr = sCmdStr & "/resize=(" & Optimal_W * 0.2 & "," & Optimal_H * 0.2 & ") /aspectratio /convert="
                                                    sCmdStr = sCmdStr & LCase(msPhotoReposPath & sPhotoThumbName)
                                                    dblRet = Shell(sCmdStr, vbHide)
                                                    DoEvents
                                                    Sleep 10
                                                    For lWaitAWhile = 1 To 10
                                                        If goUtil.utGetFileData(msPhotoReposPath & sPhotoThumbName) = vbNullString Then
                                                            DoEvents
                                                            Sleep 1000
                                                            If goUtil.utGetFileData(msPhotoReposPath & sPhotoThumbName) = vbNullString Then
                                                                'try it again
                                                                dblRet = Shell(sCmdStr, vbHide)
                                                            End If
                                                        Else
                                                            Exit For
                                                        End If
                                                    Next
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
SKIP_DIAGRAM:
                                mbSingleFileProcess = False
                                'Check for Errors
                                If Err.Number <> 0 Then
                                    lstProcess.AddItem "ERROR #" & Err.Number & " File(" & sDLFile & ") "
                                    lstProcess.AddItem Err.Description
                                    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Function UpdatePhotoAttatchmentDownload", False
                                    Err.Clear
CANCEL_DOWNLOAD3:
                                    If StrComp(sTableName, "RTPhotoLog", vbTextCompare) = 0 Then
                                        If sSQLBuildNotIn = vbNullString Then
                                            sSQLBuildNotIn = sSQLBuildNotIn & RS.Fields("RTPhotoLogID")
                                        Else
                                            sSQLBuildNotIn = sSQLBuildNotIn & ", " & RS.Fields("RTPhotoLogID")
                                        End If
                                    ElseIf StrComp(sTableName, "RTWSDiagram", vbTextCompare) = 0 Then
                                        If sSQLBuildNotIn = vbNullString Then
                                            sSQLBuildNotIn = sSQLBuildNotIn & RS.Fields("RTWSDiagramID")
                                        Else
                                            sSQLBuildNotIn = sSQLBuildNotIn & ", " & RS.Fields("RTWSDiagramID")
                                        End If
                                    End If
                                End If
                                On Error GoTo 0
                                On Error GoTo EH
                            End If
                    End Select
SKIP_DELETED:
                    RS.MoveNext
                    DoEvents
                    Sleep 10
                Loop
            Else
                bRecordsFound = False
            End If
            'If records were found then need to update Server side
            'to reset download flags
            If bRecordsFound Then
                If oXZip Is Nothing Then
                    Set oXZip = New V2ECKeyBoard.clsXZip
                End If
                sSQL = "UPDATE " & sTableName & " SET "
                'Need to reset other download upload flags
                Select Case UCase(sTableName)
                    Case UCase("Assignments")
                        sSQL = sSQL & "DownLoadLossReport = 0 "
                    Case UCase("RTPhotoLog")
                        sSQL = sSQL & "DownloadPhoto = 0, "
                        sSQL = sSQL & "DownloadPhotoThumb = 0, "
                        sSQL = sSQL & "DownloadPhotoHighRes = 0 "
                    Case UCase("RTAttachments")
                        sSQL = sSQL & "DownloadAttachment = 0 "
                    Case UCase("RTWSDiagram")
                        sSQL = sSQL & "DownloadDiagramPhoto = 0 "
                End Select
                
                'Add the Where
                sSQL = sSQL & sSQLWHERE & " "
                'Check for not in string...
                'The not in string will be populated with ID for those Photos
                'or Attachments that created errors while trying to upload
                'These items upload flags will not be reset and will thereby
                'be uploaded again.  This will continue until those items
                'have been successfully uploaded.
                If sSQLBuildNotIn <> vbNullString Then
                    sSQLAndNotIN = Replace(sSQLAndNotIN, "***", sSQLBuildNotIn, , , vbBinaryCompare)
                    sSQL = sSQL & sSQLAndNotIN & " "
                End If
                
                'Update Access DB
                oConn.Execute (sSQL)
                'Also need to Update Server to reset download flags
                sResetDLFlagsData = vbNullString
                'Replace the word True with 1 for TSQL Bit
                sResetDLFlagsData = Replace(sSQL, "= True", "= 1", , , vbTextCompare)
                'Save the Data to Upload Folder
                sTickCount = goUtil.utGetTickCount
                sFileName = "UPDATE_" & sTableName & "_" & sTickCount & ".ulud"
                sFileNameZip = "UPDATE_" & sTableName & "_" & sTickCount & ".zulud"

                goUtil.utSaveFileData msFTPULPath & sFileName, sResetDLFlagsData
                oXZip.SaveZIPFiles msFTPULPath, sFileNameZip, sFileName, goUtil.DB_PASSWORD("1")
                SetAttr msFTPULPath & sFileNameZip, vbNormal
                bDoUpdateDownload = True
                'Change FTP Dr and Update the Server Side UPdate
                mbSingleFileProcess = True
                On Error Resume Next
                EC_FTP.ChangeDir "\.\"
                EC_FTP.ChangeDir msUserFolders & "\USER_FOLDERS\" & msUserName & "\"
                lRetry = 0
RETRY:
                DoEvents
                Sleep 10
                EC_FTP.PutFile msFTPULPath & sFileNameZip, sFileNameZip
                If Err.Number <> 0 Then
                    lRetry = lRetry + 1
                    If lRetry <= 10 Then
                        Sleep 500
                        Err.Clear
                        On Error Resume Next
                        GoTo RETRY
                    End If
                    lstProcess.AddItem "ERROR #" & Err.Number & " File(" & sFileNameZip & ") "
                    lstProcess.AddItem Err.Description
                    Err.Clear
                    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Function UpdatePhotoAttatchmentDownload", False
                    On Error GoTo 0
                    On Error GoTo EH
                End If
                goUtil.utDeleteFile msFTPULPath & sFileNameZip
                mbSingleFileProcess = False
                On Error GoTo 0
                On Error GoTo EH
            End If
        End If
    Next
    
    If bDoUpdateDownload Then
        UploadClientFlag msUserName, "UploadUpdateReady.flag", True
        Sleep 1000
    End If
    
    If mbCancelPhotoAttach Then
        lstProcess.AddItem "Canceled Photo and Attachment Download... " & Now()
        lstProcess.AddItem "Please connect again as soon as possible to Download these files!!!"
        Me.Refresh
    End If
    
    cmdCancelPhotoAttachUL.Visible = False
    
    'cleanup
    Set RS = Nothing
    Set oConn = Nothing
    Set oXZip = Nothing
    
    UpdatePhotoAttatchmentDownload = True
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    'Log the error but don't show it since its already shown above
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Function UpdatePhotoAttatchmentDownload", False
    
    lstProcess.AddItem "ERROR #" & lErrNum
    lstProcess.AddItem sErrDesc
    
End Function

Private Function UpdatePhotoAttatchmentUpload() As Boolean
    On Error GoTo EH
    Dim sMess As String
    Dim lWaitAWhile As Long
    Dim lHwnd As Long
    Dim dblRet As Double
    Dim RS As ADODB.Recordset
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    Dim sSQLWHERE As String
    Dim sSQLAndNotIN As String
    Dim sSQLBuildNotIn As String
    Dim sTableName As String
    Dim saryDBTables(1 To 4) As String
    Dim lCountDBTables As Long
    Dim sLRFormat As String
    Dim sULFile As String
    Dim lDLFileCount As Long
    Dim sYYMM As String
    Dim sDD As String
    Dim saryTemp() As String
    'Photos
    Dim sPhotoName As String
    Dim sPhotoHighResName As String
    Dim sPhotoThumbName As String
    'For Rebuilding Thumb if it does not exist on the Server
    'When flagged for UPload
    Dim sCmdStr As String
    Dim sIrfanViewPath As String
    Dim Optimal_H As Long
    Dim Optimal_W As Long
    'Used by Photos or Diagram Photo
    Dim sUploadFlag As String
    
    'Update Server to reset Upload Flags
    Dim sTickCount As String
    Dim bRecordsFound As Boolean
    Dim sResetULFlagsData As String
    Dim sFileName As String
    Dim sFileNameZip As String
    Dim oXZip As V2ECKeyBoard.clsXZip
    Dim bDoUpdateUpload  As Boolean
    Dim lRecordsAffected As Long
    Dim sLOCKED_AssignmentsID As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim lRetry As Long
    
    sLOCKED_AssignmentsID = GetSetting("ECFTP", "MSG", "LOCKED_AssignmentsID", vbNullString)
    
    
    lblProcess.Caption = "Checking for Photo and Attachment Upload..."
    lblProcess.Refresh
    lstProcess.AddItem lblProcess.Caption & " " & Now()
    
    Set RS = New ADODB.Recordset
    Set oConn = New ADODB.Connection
    
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    saryDBTables(1) = "Assignments"
    saryDBTables(2) = "RTPhotoLog"
    saryDBTables(3) = "RTAttachments"
    saryDBTables(4) = "RTWSDiagram"
    
    sIrfanViewPath = goUtil.utGetSystemDir & "\ECS\DLL\i_view32.exe"
    Optimal_W = V2ECKeyBoard.PhotoSettings.Optimal_W
    Optimal_H = V2ECKeyBoard.PhotoSettings.Optimal_H
    Optimal_W = goUtil.ConvertTwipsToPixels(Optimal_W)
    Optimal_H = goUtil.ConvertTwipsToPixels(Optimal_H)
    
    For lCountDBTables = LBound(saryDBTables, 1) To UBound(saryDBTables, 1)
        sTableName = saryDBTables(lCountDBTables)
        If sTableName <> vbNullString Then
            sSQL = "SELECT * FROM " & sTableName & " "
            sSQLWHERE = "WHERE "
            Select Case UCase(sTableName)
                Case UCase("Assignments")
                    sSQLWHERE = sSQLWHERE & "UpLoadLossReport = True "
                    sSQLWHERE = sSQLWHERE & "AND AdjusterSpecID IN( "
                                            sSQLWHERE = sSQLWHERE & "SELECT   ClientCoAdjusterSpecID "
                                            sSQLWHERE = sSQLWHERE & "FROM     ClientCoAdjusterSpec "
                                            sSQLWHERE = sSQLWHERE & "WHERE    UsersID = " & msUsersID & " "
                                            sSQLWHERE = sSQLWHERE & ") "
                    sSQLAndNotIN = "AND [AssignmentsID] Not IN (***) "
                    sSQLBuildNotIn = vbNullString
                Case UCase("RTPhotoLog")
                    sSQLWHERE = sSQLWHERE & "( "
                    sSQLWHERE = sSQLWHERE & "UploadPhoto = True "
                    sSQLWHERE = sSQLWHERE & "OR "
                    sSQLWHERE = sSQLWHERE & "UploadPhotoThumb = True "
                    sSQLWHERE = sSQLWHERE & "OR "
                    sSQLWHERE = sSQLWHERE & "UploadPhotoHighRes = True "
                    sSQLWHERE = sSQLWHERE & ") "
                    sSQLWHERE = sSQLWHERE & "AND AssignmentsID IN ("
                                            sSQLWHERE = sSQLWHERE & "SELECT   AssignmentsID "
                                            sSQLWHERE = sSQLWHERE & "FROM     Assignments "
                                            sSQLWHERE = sSQLWHERE & "WHERE    AdjusterSpecID IN( "
                                                                    sSQLWHERE = sSQLWHERE & "SELECT   ClientCoAdjusterSpecID "
                                                                    sSQLWHERE = sSQLWHERE & "FROM     ClientCoAdjusterSpec "
                                                                    sSQLWHERE = sSQLWHERE & "WHERE    UsersID = " & msUsersID & " "
                                                                    sSQLWHERE = sSQLWHERE & ") "
                                             sSQLWHERE = sSQLWHERE & ") "
                    sSQLAndNotIN = "AND [RTPhotoLogID] Not IN (***) "
                    sSQLBuildNotIn = vbNullString
                Case UCase("RTAttachments")
                    sSQLWHERE = sSQLWHERE & "UploadAttachment = True "
                    sSQLWHERE = sSQLWHERE & "AND AssignmentsID IN ("
                                            sSQLWHERE = sSQLWHERE & "SELECT   AssignmentsID "
                                            sSQLWHERE = sSQLWHERE & "FROM     Assignments "
                                            sSQLWHERE = sSQLWHERE & "WHERE    AdjusterSpecID IN( "
                                                                    sSQLWHERE = sSQLWHERE & "SELECT   ClientCoAdjusterSpecID "
                                                                    sSQLWHERE = sSQLWHERE & "FROM     ClientCoAdjusterSpec "
                                                                    sSQLWHERE = sSQLWHERE & "WHERE    UsersID = " & msUsersID & " "
                                                                    sSQLWHERE = sSQLWHERE & ") "
                                             sSQLWHERE = sSQLWHERE & ") "
                    sSQLAndNotIN = "AND [RTAttachmentsID] Not IN (***) "
                    sSQLBuildNotIn = vbNullString
                Case UCase("RTWSDiagram")
                    sSQLWHERE = sSQLWHERE & "UploadDiagramPhoto = True "
                    sSQLWHERE = sSQLWHERE & "AND AssignmentsID IN ("
                                            sSQLWHERE = sSQLWHERE & "SELECT   AssignmentsID "
                                            sSQLWHERE = sSQLWHERE & "FROM     Assignments "
                                            sSQLWHERE = sSQLWHERE & "WHERE    AdjusterSpecID IN( "
                                                                    sSQLWHERE = sSQLWHERE & "SELECT   ClientCoAdjusterSpecID "
                                                                    sSQLWHERE = sSQLWHERE & "FROM     ClientCoAdjusterSpec "
                                                                    sSQLWHERE = sSQLWHERE & "WHERE    UsersID = " & msUsersID & " "
                                                                    sSQLWHERE = sSQLWHERE & ") "
                                             sSQLWHERE = sSQLWHERE & ") "
                    sSQLAndNotIN = "AND [RTWSDiagramID] Not IN (***) "
                    sSQLBuildNotIn = vbNullString
            End Select
            
            'Check for a locked assignment
            If sLOCKED_AssignmentsID <> vbNullString Then
                sSQLWHERE = sSQLWHERE & "AND [AssignmentsID] <> " & sLOCKED_AssignmentsID & "  "
            End If
            
            'Add the Where
            sSQL = sSQL & sSQLWHERE
            
            Set RS = New ADODB.Recordset
            'Use Disconnected Record Set on asUseClient Cusor ONLY !
            RS.CursorLocation = adUseClient
            RS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
            Set RS.ActiveConnection = Nothing
            
            lDLFileCount = 0
            PBAttach.Value = 0
            PBPhoto.Value = 0
            bRecordsFound = False
            If Not RS.EOF Then
                'Unhide the cmdCancelPhotoAttachUL button "Finis Later"
                cmdCancelPhotoAttachUL.Visible = True
                cmdCancelPhotoAttachUL.Enabled = True
                'Reset the Cancel Flag
                mbCancelPhotoAttach = False
                
                bRecordsFound = True
                Select Case UCase(sTableName)
                    Case UCase("Assignments"), UCase("RTAttachments")
                        PBAttach.Max = RS.RecordCount
                    Case UCase("RTPhotoLog"), UCase("RTWSDiagram")
                        PBPhoto.Max = RS.RecordCount
                End Select
                Do Until RS.EOF
                    lDLFileCount = lDLFileCount + 1
                    Select Case UCase(sTableName)
                        Case UCase("Assignments"), UCase("RTAttachments")
                            PBAttach.Value = lDLFileCount
                            If mbCancelPhotoAttach Then
                                lblProcess.Caption = "Canceling Upload Attachments... (" & PBAttach.Max & ") Item(s)  Please Wait! ...  Item (" & PBAttach.Value & ") Of (" & PBAttach.Max & ") "
                            Else
                                lblProcess.Caption = "Uploading Attachments... (" & PBAttach.Max & ") Item(s)  Please Wait! ...  Item (" & PBAttach.Value & ") Of (" & PBAttach.Max & ") "
                            End If
                            lblProcess.Refresh
                        Case UCase("RTPhotoLog"), UCase("RTWSDiagram")
                            PBPhoto.Value = lDLFileCount
                            If mbCancelPhotoAttach Then
                                lblProcess.Caption = "Canceling Upload Photos... (" & PBPhoto.Max & ") Item(s)  Please Wait! ...  Item (" & PBPhoto.Value & ") Of (" & PBPhoto.Max & ") "
                            Else
                                lblProcess.Caption = "Uploading Photos... (" & PBPhoto.Max & ") Item(s)  Please Wait! ...  Item (" & PBPhoto.Value & ") Of (" & PBPhoto.Max & ") "
                            End If
                            lblProcess.Refresh
                    End Select
                    Select Case UCase(sTableName)
                        Case UCase("Assignments")
                            sLRFormat = vbNullString
                            sULFile = vbNullString
                            sLRFormat = IIf(IsNull(RS!LRFormat), vbNullString, RS!LRFormat)
                            sULFile = IIf(IsNull(RS!LossReport), vbNullString, RS!LossReport)
                            sULFile = Trim(sULFile)
                            If InStr(1, sLRFormat, "OLEType_pdf", vbTextCompare) > 0 And sULFile <> vbNullString Then
                                'Have to Build the Path to the File
                                'Accounting for YYMM and DD Folder Structure
                                'This is How the Files are stored on the Server.
                                'the Client will just store the files in
                                'AttachRepos
                                'PhotoRepos
                                'Without the YYMM and DD Sub Folders
                                saryTemp = Split(sULFile, "_")
                                sYYMM = saryTemp(1)
                                sDD = saryTemp(1)
                                sYYMM = Left(sYYMM, 4)
                                sDD = Mid(sDD, 5, 2)
                                
                                'If there is an error Uploading this File Log it
                                'But Continue On to other files
                                If goUtil.utFileExists(msAttachReposPath & sULFile) Then
                                    If mbCancelPhotoAttach Then
                                        GoTo CANCEL_UPLOAD1
                                    End If
                                    On Error Resume Next
                                    mbSingleFileProcess = True
                                    EC_FTP.ChangeDir "\.\"
                                    'Try to change directly to the dir first
                                    'If it errors then create it
                                    EC_FTP.ChangeDir msUserFolders & "\Upload\AttachRepos\" & sYYMM & "\" & sDD & "\"
                                    If Err.Number <> 0 Then
                                        Err.Clear
                                        EC_FTP.ChangeDir "\.\"
                                        EC_FTP.ChangeDir msUserFolders & "\Upload\AttachRepos\"
                                        EC_FTP.CreateDir sYYMM
                                        EC_FTP.ChangeDir "\.\"
                                        EC_FTP.ChangeDir msUserFolders & "\Upload\AttachRepos\" & sYYMM & "\"
                                        EC_FTP.CreateDir sDD
                                        EC_FTP.ChangeDir "\.\"
                                        EC_FTP.ChangeDir msUserFolders & "\Upload\AttachRepos\" & sYYMM & "\" & sDD & "\"
                                    End If
                                    If Err.Number = 0 Then
                                        lRetry = 0
RETRY:
                                        DoEvents
                                        Sleep 10
                                        EC_FTP.PutFile msAttachReposPath & sULFile, sULFile
                                        If Err.Number <> 0 Then
                                            lRetry = lRetry + 1
                                            If lRetry <= 10 Then
                                                Sleep 500
                                                Err.Clear
                                                On Error Resume Next
                                                GoTo RETRY
                                            End If
                                        End If
                                    End If
                                    mbSingleFileProcess = False
                                    'Check for Errors
                                    If Err.Number <> 0 Then
                                        lstProcess.AddItem "ERROR #" & Err.Number & " File(" & sULFile & ") "
                                        lstProcess.AddItem Err.Description
                                        goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Function UpdatePhotoAttatchmentUpload", False
                                        lblProcess.Caption = "ERROR CONNECT AGAIN!"
                                        Err.Clear
CANCEL_UPLOAD1:
                                        If sSQLBuildNotIn = vbNullString Then
                                            sSQLBuildNotIn = sSQLBuildNotIn & RS.Fields("AssignmentsID")
                                        Else
                                            sSQLBuildNotIn = sSQLBuildNotIn & ", " & RS.Fields("AssignmentsID")
                                        End If
                                    End If
                                    On Error GoTo 0
                                    On Error GoTo EH
                                End If
                            End If
                            
                        Case UCase("RTAttachments")
                            sULFile = vbNullString
                            sULFile = IIf(IsNull(RS!Attachment), vbNullString, RS!Attachment)
                            sULFile = Trim(sULFile)
                            If sULFile <> vbNullString Then
                                'Have to Build the Path to the File
                                'Accounting for YYMM and DD Folder Structure
                                'This is How the Files are stored on the Server.
                                'the Client will just store the files in
                                'AttachRepos
                                'PhotoRepos
                                'Without the YYMM and DD Sub Folders
                                saryTemp = Split(sULFile, "_")
                                sYYMM = saryTemp(1)
                                sDD = saryTemp(1)
                                sYYMM = Left(sYYMM, 4)
                                sDD = Mid(sDD, 5, 2)
                                
                                'If there is an error Uploading this File Log it
                                'But Continue On to other files
                                If goUtil.utFileExists(msAttachReposPath & sULFile) Then
                                    If mbCancelPhotoAttach Then
                                        GoTo CANCEL_UPLOAD2
                                    End If
                                    On Error Resume Next
                                    mbSingleFileProcess = True
                                    EC_FTP.ChangeDir "\.\"
                                    'Try to change directly to the dir first
                                    'If it errors then create it
                                    EC_FTP.ChangeDir msUserFolders & "\Upload\AttachRepos\" & sYYMM & "\" & sDD & "\"
                                    If Err.Number <> 0 Then
                                        Err.Clear
                                        EC_FTP.ChangeDir "\.\"
                                        EC_FTP.ChangeDir msUserFolders & "\Upload\AttachRepos\"
                                        EC_FTP.CreateDir sYYMM
                                        EC_FTP.ChangeDir "\.\"
                                        EC_FTP.ChangeDir msUserFolders & "\Upload\AttachRepos\" & sYYMM & "\"
                                        EC_FTP.CreateDir sDD
                                        EC_FTP.ChangeDir "\.\"
                                        EC_FTP.ChangeDir msUserFolders & "\Upload\AttachRepos\" & sYYMM & "\" & sDD & "\"
                                    End If
                                    If Err.Number = 0 Then
                                        lRetry = 0
RETRY2:
                                        DoEvents
                                        Sleep 10
                                        EC_FTP.PutFile msAttachReposPath & sULFile, sULFile
                                        If Err.Number <> 0 Then
                                            lRetry = lRetry + 1
                                            If lRetry <= 10 Then
                                                Sleep 500
                                                Err.Clear
                                                On Error Resume Next
                                                GoTo RETRY2
                                            End If
                                        End If
                                    End If
                                    mbSingleFileProcess = False
                                    'Check for Errors
                                    If Err.Number <> 0 Then
                                        lstProcess.AddItem "ERROR #" & Err.Number & " File(" & sULFile & ") "
                                        lstProcess.AddItem Err.Description
                                        goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Function UpdatePhotoAttatchmentUpload", False
                                        lblProcess.Caption = "ERROR CONNECT AGAIN!"
                                        Err.Clear
CANCEL_UPLOAD2:
                                        If sSQLBuildNotIn = vbNullString Then
                                            sSQLBuildNotIn = sSQLBuildNotIn & RS.Fields("RTAttachmentsID")
                                        Else
                                            sSQLBuildNotIn = sSQLBuildNotIn & ", " & RS.Fields("RTAttachmentsID")
                                        End If
                                    End If
                                    On Error GoTo 0
                                    On Error GoTo EH
                                End If
                            End If
                        Case UCase("RTPhotoLog"), UCase("RTWSDiagram")
                            sULFile = vbNullString
                            If StrComp(sTableName, "RTPhotoLog", vbTextCompare) = 0 Then
                                sULFile = IIf(IsNull(RS!PhotoName), vbNullString, RS!PhotoName)
                                sUploadFlag = "UploadPhoto"
                            ElseIf StrComp(sTableName, "RTWSDiagram", vbTextCompare) = 0 Then
                                sULFile = IIf(IsNull(RS!DiagramPhotoName), vbNullString, RS!DiagramPhotoName)
                                sUploadFlag = "UploadDiagramPhoto"
                            End If
                            sULFile = Trim(sULFile)
                            sPhotoName = vbNullString
                            If sULFile <> vbNullString Then
                                'Have to Build the Path to the File
                                'Accounting for YYMM and DD Folder Structure
                                'This is How the Files are stored on the Server.
                                'the Client will just store the files in
                                'AttachRepos
                                'PhotoRepos
                                'Without the YYMM and DD Sub Folders
                                saryTemp = Split(sULFile, "_")
                                sYYMM = saryTemp(1)
                                sDD = saryTemp(1)
                                sYYMM = Left(sYYMM, 4)
                                sDD = Mid(sDD, 5, 2)
                                
                                If mbCancelPhotoAttach Then
                                    GoTo CANCEL_UPLOAD3
                                End If
                                'If there is an error Uploading this File Log it
                                'But Continue On to other files
                                mbSingleFileProcess = True
                                'Check for Each Type OF Photo Upload...
                                If CBool(RS.Fields(sUploadFlag)) Then
                                     sPhotoName = sULFile
                                    'UploadPhoto
                                    If goUtil.utFileExists(msPhotoReposPath & sULFile) Then
                                        On Error Resume Next
                                        EC_FTP.ChangeDir "\.\"
                                        'Try to change directly to the dir first
                                        'If it errors then create it
                                        EC_FTP.ChangeDir msUserFolders & "\Upload\PhotoRepos\" & sYYMM & "\" & sDD & "\"
                                        If Err.Number <> 0 Then
                                            Err.Clear
                                            EC_FTP.ChangeDir "\.\"
                                            EC_FTP.ChangeDir msUserFolders & "\Upload\PhotoRepos\"
                                            EC_FTP.CreateDir sYYMM
                                            EC_FTP.ChangeDir "\.\"
                                            EC_FTP.ChangeDir msUserFolders & "\Upload\PhotoRepos\" & sYYMM & "\"
                                            EC_FTP.CreateDir sDD
                                            EC_FTP.ChangeDir "\.\"
                                            EC_FTP.ChangeDir msUserFolders & "\Upload\PhotoRepos\" & sYYMM & "\" & sDD & "\"
                                        End If
                                        If Err.Number = 0 Then
                                            lRetry = 0
RETRY3:
                                            DoEvents
                                            Sleep 10
                                            EC_FTP.PutFile msPhotoReposPath & sULFile, sULFile
                                            If Err.Number <> 0 Then
                                                lRetry = lRetry + 1
                                                If lRetry <= 10 Then
                                                    Sleep 500
                                                    Err.Clear
                                                    On Error Resume Next
                                                    GoTo RETRY3
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                                If Err.Number = 0 Then
                                    If StrComp(sUploadFlag, "UploadDiagramPhoto", vbTextCompare) = 0 Then
                                        GoTo SKIP_DIAGRAM
                                    End If
                                    '######################## Not uploading Highres or Thumbnails at this time#################
                                    '-------------------------------------------12/2/2004-------------------------------------
                                    '##########################################################################################
                                    GoTo SKIP_DIAGRAM
                                    '######################## Not uploading Highres or Thumbnails at this time#################
                                    '-------------------------------------------12/2/2004-------------------------------------
                                    '##########################################################################################
                                    On Error GoTo 0
                                    On Error GoTo EH
                                    If CBool(RS!UploadPhotoHighRes) Then
                                        sULFile = Replace(sULFile, "_1.jpg", "_0.jpg", , , vbTextCompare)
                                        sPhotoHighResName = sULFile
                                        'UploadPhotoHighRes
                                        If goUtil.utFileExists(msPhotoReposPath & sULFile) Then
                                            On Error Resume Next
                                            EC_FTP.ChangeDir "\.\"
                                            EC_FTP.ChangeDir msUserFolders & "\Upload\PhotoRepos\" & sYYMM & "\" & sDD & "\"
                                            If Err.Number = 0 Then
                                                EC_FTP.PutFile msPhotoReposPath & sULFile, sULFile
                                                DoEvents
                                                Sleep 10
                                            End If
                                        End If
                                    End If
                                    'Be sure to create the Thumb last since it will be using the main
                                    'File.  don't want to Copy this file while processing it by irfanview
                                    If CBool(RS!UploadPhotoThumb) Then
                                        If CBool(RS!UploadPhotoHighRes) Then
                                            sULFile = Replace(sULFile, "_0.jpg", "_2.jpg", , , vbTextCompare)
                                        Else
                                            sULFile = Replace(sULFile, "_1.jpg", "_2.jpg", , , vbTextCompare)
                                        End If
                                        
                                        sPhotoThumbName = sULFile
                                        'UploadPhotoThumb
                                        If goUtil.utFileExists(msPhotoReposPath & sULFile) Then
                                            On Error Resume Next
                                            EC_FTP.ChangeDir "\.\"
                                            EC_FTP.ChangeDir msUserFolders & "\Upload\PhotoRepos\" & sYYMM & "\" & sDD & "\"
                                            If Err.Number = 0 Then
                                                EC_FTP.PutFile msPhotoReposPath & sULFile, sULFile
                                                DoEvents
                                                Sleep 10
                                            End If
                                        End If
                                    End If
                                End If
SKIP_DIAGRAM:
                                mbSingleFileProcess = False
                                'Check for Errors
                                If Err.Number <> 0 Then
                                    lstProcess.AddItem "ERROR #" & Err.Number & " File(" & sULFile & ") "
                                    lstProcess.AddItem Err.Description
                                    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Function UpdatePhotoAttatchmentUpLoad", False
                                    lblProcess.Caption = "ERROR CONNECT AGAIN!"
                                    Err.Clear
CANCEL_UPLOAD3:
                                    If StrComp(sTableName, "RTPhotoLog", vbTextCompare) = 0 Then
                                        If sSQLBuildNotIn = vbNullString Then
                                            sSQLBuildNotIn = sSQLBuildNotIn & RS.Fields("RTPhotoLogID")
                                        Else
                                            sSQLBuildNotIn = sSQLBuildNotIn & ", " & RS.Fields("RTPhotoLogID")
                                        End If
                                    ElseIf StrComp(sTableName, "RTWSDiagram", vbTextCompare) = 0 Then
                                        If sSQLBuildNotIn = vbNullString Then
                                            sSQLBuildNotIn = sSQLBuildNotIn & RS.Fields("RTWSDiagramID")
                                        Else
                                            sSQLBuildNotIn = sSQLBuildNotIn & ", " & RS.Fields("RTWSDiagramID")
                                        End If
                                    End If
                                End If
                                On Error GoTo 0
                                On Error GoTo EH
                            End If
                    End Select
                    RS.MoveNext
                    DoEvents
                    Sleep 10
                Loop
            Else
                bRecordsFound = False
            End If
            'If records were found then need to update Server side
            'to reset Upload flags
            If bRecordsFound Then
                If oXZip Is Nothing Then
                    Set oXZip = New V2ECKeyBoard.clsXZip
                End If
                sSQL = "UPDATE " & sTableName & " SET "
                'Need to reset other Upload upload flags
                Select Case UCase(sTableName)
                    Case UCase("Assignments")
                        sSQL = sSQL & "UpLoadLossReport = 0 "
                    Case UCase("RTPhotoLog")
                        sSQL = sSQL & "UploadPhoto = 0, "
                        sSQL = sSQL & "UploadPhotoThumb = 0, "
                        sSQL = sSQL & "UploadPhotoHighRes = 0 "
                    Case UCase("RTAttachments")
                        sSQL = sSQL & "UploadAttachment = 0 "
                    Case UCase("RTWSDiagram")
                        sSQL = sSQL & "UploadDiagramPhoto = 0 "
                End Select
                
                'Add the Where
                sSQL = sSQL & sSQLWHERE & " "
                'Check for not in string...
                'The not in string will be populated with ID for those Photos
                'or Attachments that created errors while trying to upload
                'These items upload flags will not be reset and will thereby
                'be uploaded again.  This will continue until those items
                'have been successfully uploaded.
                If sSQLBuildNotIn <> vbNullString Then
                    sSQLAndNotIN = Replace(sSQLAndNotIN, "***", sSQLBuildNotIn, , , vbBinaryCompare)
                    sSQL = sSQL & sSQLAndNotIN & " "
                End If
                
                'Update Access DB
                oConn.Execute sSQL, lRecordsAffected
                'Also need to Update Server to reset Upload flags
                sResetULFlagsData = vbNullString
                'Replace the word True with 1 for TSQL Bit
                sResetULFlagsData = Replace(sSQL, "= True", "= 1", , , vbTextCompare)
                'Save the Data to Upload Folder
                sTickCount = goUtil.utGetTickCount
                sFileName = "UPDATE_" & sTableName & "_" & sTickCount & ".ulud"
                sFileNameZip = "UPDATE_" & sTableName & "_" & sTickCount & ".zulud"

                goUtil.utSaveFileData msFTPULPath & sFileName, sResetULFlagsData
                oXZip.SaveZIPFiles msFTPULPath, sFileNameZip, sFileName, goUtil.DB_PASSWORD("1")
                SetAttr msFTPULPath & sFileNameZip, vbNormal
                bDoUpdateUpload = True
                'Change FTP Dr and Update the Server Side UPdate
                mbSingleFileProcess = True
                On Error Resume Next
                EC_FTP.ChangeDir "\.\"
                EC_FTP.ChangeDir msUserFolders & "\USER_FOLDERS\" & msUserName & "\"
                lRetry = 0
RETRY4:
                DoEvents
                Sleep 10
                EC_FTP.PutFile msFTPULPath & sFileNameZip, sFileNameZip
                If Err.Number <> 0 Then
                    lRetry = lRetry + 1
                    If lRetry <= 10 Then
                        Sleep 500
                        Err.Clear
                        On Error Resume Next
                        GoTo RETRY4
                    End If
                    lstProcess.AddItem "ERROR #" & Err.Number & " File(" & sFileNameZip & ") "
                    lstProcess.AddItem Err.Description
                    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Function UpdatePhotoAttatchmentUpload", False
                    lblProcess.Caption = "ERROR CONNECT AGAIN!"
                    Err.Clear
                    On Error GoTo 0
                    On Error GoTo EH
                End If
                goUtil.utDeleteFile msFTPULPath & sFileNameZip
                mbSingleFileProcess = False
                On Error GoTo 0
                On Error GoTo EH
            End If
        End If
    Next
    
    If bDoUpdateUpload Then
        UploadClientFlag msUserName, "UploadUpdateReady.flag", True
        Sleep 1000
    End If
    
    If mbCancelPhotoAttach Then
        lstProcess.AddItem "Canceled Photo and Attachment Upload... " & Now()
        lstProcess.AddItem "Please connect again as soon as possible to upload these files!!!"
        Me.Refresh
    End If
    
    cmdCancelPhotoAttachUL.Visible = False
    
    
    'cleanup
    Set RS = Nothing
    Set oConn = Nothing
    Set oXZip = Nothing
    
    UpdatePhotoAttatchmentUpload = True
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    'Log the error but don't show it since its already shown above
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Function UpdatePhotoAttatchmentUpload", False
    
    lstProcess.AddItem "ERROR #" & lErrNum
    lstProcess.AddItem sErrDesc
    
    cmdCancelPhotoAttachUL.Visible = False
End Function

Private Function VerifySecurity(pvAryToken As Variant, _
                                Optional pbCheckAgain As Boolean, _
                                Optional psMessage, _
                                Optional plWaitSeconds As Long = 60) As Boolean
    On Error GoTo EH
    Dim sToken As String
    Dim sTokenOut As String
    Dim iSleepCount As Integer
    Dim bFoundToken As Boolean
    Dim saryTokOUT() As String
    Dim sUserName As String
    Dim sLicDaysLeft As String
    Dim sIBPrefix As String
    Dim bUserFolderSelected As Boolean
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    sUserName = pvAryToken(SecurityToken.UserName)
    sUserName = goUtil.Decode(sUserName)
    sLicDaysLeft = pvAryToken(SecurityToken.LicDaysLeft)
    sIBPrefix = pvAryToken(SecurityToken.IBPrefix)
    
    mbSingleFileProcess = True
    
    If Not pbCheckAgain Then
        lblProcess.Caption = "Verifying User/Password *"
        lblProcess.Refresh
    
        lstProcess.AddItem "Verifying User/Password (Encrypted Data Token)"
        lstProcess.Refresh
        imgEncrypt.Visible = True
        imgEncrypt.Refresh
    Else
        lblProcess.Caption = psMessage & " *"
        lblProcess.Refresh
    
        lstProcess.AddItem psMessage & " (Encrypted ZIP Files)"
        lstProcess.Refresh
        imgEncrypt.Visible = True
        imgEncrypt.Refresh
    End If
    
    EC_FTP.ChangeDir msUserFolders & "\"

'Build the Token...
    'Build the Security Token we will send to the server for verification
    sToken = Join(pvAryToken, vbCrLf)

    'build the temp file path
    If Not pbCheckAgain Then
        msTokName = sUserName & "_" & Format(Now(), "YYMMDDHHMMSS")
        msTokPath = App.Path & "\" & msTokName
        goUtil.utSaveFileData msTokPath & ".tokin", sToken
        EC_FTP.PutFile msTokPath & ".tokin", msTokName & ".tokin"
    End If
    
    EC_FTP.ChangeDir "USER_FOLDERS\"
    
    'Wait for 10 seconds to change to Directory to user Name under User_Folders
    For iSleepCount = 1 To 10
        DoEvents
        Sleep 1000
        lblProcess.Caption = lblProcess.Caption & "*"
        lblProcess.Refresh
        On Error Resume Next
        EC_FTP.ChangeDir sUserName & "\"
        If Err.Number = 0 Then
            bUserFolderSelected = True
            Exit For
        Else
            Err.Clear
        End If
    Next

    If Not pbCheckAgain Then
        'Now can get rid of it from the temp dir
        'since it has been uploaded to the server.
        SetAttr msTokPath & ".tokin", vbNormal
        Kill msTokPath & ".tokin"
    End If

    'Wait for up to 60 seconds for the .tokout file appears
    For iSleepCount = 1 To plWaitSeconds
        DoEvents
        Sleep 1000

        lblProcess.Caption = lblProcess.Caption & "*"
        lblProcess.Refresh

        On Error Resume Next
        EC_FTP.GetFile msTokName & ".tokout", msTokPath & ".tokout"
        If Err.Number <> 0 Then
            Err.Clear
        End If
        On Error GoTo EH
        If goUtil.utFileExists(msTokPath & ".tokout") Then
            'Reset the token to the .tokout produced by the server
            sTokenOut = goUtil.utGetFileData(msTokPath & ".tokout")
            If Len(sTokenOut) > 0 Then

                'Remove the .tokout from the server
                'the .tokin will be removed by the server
                On Error Resume Next
                EC_FTP.Delete msTokName & ".tokout"
                If Err.Number <> 0 Then
                    Err.Clear
                End If
                On Error GoTo EH

                'remove the .tokout from temp dir
                SetAttr msTokPath & ".tokout", vbNormal
                Kill msTokPath & ".tokout"

                bFoundToken = True
                'we can exit here since we found the out token
                Exit For
            Else
                SetAttr msTokPath & ".tokout", vbNormal
                Kill msTokPath & ".tokout"
            End If
        End If
    Next

    If bFoundToken Then
        'Check the TokenOut
        saryTokOUT = Split(sTokenOut, vbCrLf)

        'Check for Invalid Security Tokin
        'Means somebody has been messing around with Reg settings
        If InStr(1, saryTokOUT(0), "<--", vbTextCompare) > 0 Then
            lstProcess.AddItem saryTokOUT(0)
            lstProcess.AddItem "Please try to connect again."
            lstProcess.AddItem "If you are still unable to connect after several attempts,"
            lstProcess.AddItem "and receive the same error..."
            lstProcess.AddItem """" & saryTokOUT(0) & """"
            lstProcess.AddItem "Your security settings may need to be reset."

            GoTo CLEANUP
        End If

        Select Case CBool(saryTokOUT(SecurityToken.UserName)) & IIf(saryTokOUT(SecurityToken.AppVSInfo) = "NO_UPDATES", True, False)

            Case True & True, True & False
                'Passed verification for UserName and Version Info

                'If we get to here we can set the password reset to false
                SaveSetting "ECS", "WEB_SECURITY", "RESET_PASSWORD", False
                'See if Password was correct too
                If CBool(saryTokOUT(SecurityToken.Pass)) Then
                    VerifySecurity = True
                    If pbCheckAgain Then
                        GoTo CLEANUP
                    End If
                    'if the lic days more than 0 then make sure the Lic is saved
                    If saryTokOUT(SecurityToken.LicDaysLeft) > 0 Then
                        If goUtil.gbValidLic Then
                            If saryTokOUT(SecurityToken.LicDaysLeft) <> goUtil.Decode(sLicDaysLeft) Then
                                'Reset this to false to reset the regform with new Lic days
                                goUtil.gbValidLic = False
                            End If
                        End If
                    Else
                        goUtil.gbValidLic = False
                    End If
                    goUtil.utSaveECSCryptSetting "ECS", "WEB_SECURITY", "LIC", saryTokOUT(SecurityToken.LicDaysLeft)
                    '1.16.2003 do the same for IBPrefix
                    sIBPrefix = saryTokOUT(SecurityToken.IBPrefix)
                    If Len(Trim(sIBPrefix)) <= 3 Then
                        SaveSetting "ECS", "WEB_SECURITY", "IB_PREFIX", sIBPrefix
                    End If
                    
                    If Not goUtil.gbValidLic Then
                        ShowRegForm
                    End If
                Else
                    lstProcess.AddItem "Invalid UserName Or Password!"
                    SaveSetting goUtil.gsMainAppExeName, "MSG", "FTP_COMMAND", "GLOBAL_PREF"
                End If

            Case True & False
                'found UserName but Version info must be updated
                lstProcess.AddItem "Software Update "
                lstProcess.AddItem "Make sure your UserName and SSN are correct."
                SaveSetting goUtil.gsMainAppExeName, "MSG", "FTP_COMMAND", "GLOBAL_PREF"

        End Select
    Else
        'If there is no token then be sure to remove the .tokin file.
        'The server is currently down and not processing Tokens
        On Error Resume Next
        EC_FTP.ChangeDir "\.\"
        EC_FTP.ChangeDir msUserFolders & "\"
        EC_FTP.Delete msTokName & ".tokin"
        If Err.Number <> 0 Then
            Err.Clear
        End If
        On Error GoTo EH
        lstProcess.AddItem "Server not responding to Token request."
        lstProcess.AddItem "Please try again later."
        lstProcess.AddItem "If you still have problems please call technical support."
    End If
CLEANUP:
    mbSingleFileProcess = False
    imgEncrypt.Visible = False
    imgEncrypt.Refresh
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    'Log the error but don't show it since its already shown above
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Function VerifySecurity", False
    
    mbSingleFileProcess = False
    lstProcess.AddItem "ERROR #" & lErrNum
    lstProcess.AddItem sErrDesc
    
End Function

Private Sub Msghook_Message(ByVal Msg As Long, ByVal wp As Long, ByVal lp As Long, result As Long)
    On Error GoTo EH
    Dim Param As String
    Dim sMess As String
    Select Case Msg
        Case cbNotify
            If mbLoadingRegForm Then
                Exit Sub
            End If
            If wp = uID Then
                Select Case lp
                    Case WM_MOUSEMOVE
                    Case WM_LBUTTONDOWN
                    Case WM_LBUTTONUP
                    Case WM_LBUTTONDBLCLK
                        ' Show form
                        DoEvents
                        Sleep 100
                        Me.Visible = True
                        Me.WindowState = vbNormal
                        DoEvents
                        Sleep 100
                    Case WM_RBUTTONDOWN
                    Case WM_RBUTTONUP
                    ' Display context menu
                    ' Highlight default (Open)
                    Call SetForegroundWindow(Me.hWnd)
                    Me.PopupMenu mPopUp, vbPopupMenuRightButton, , , mPop(0)
        
                    Case WM_RBUTTONDBLCLK
                    Case WM_MBUTTONDOWN
                    Case WM_MBUTTONUP
                    Case WM_MBUTTONDBLCLK
                    Case Else
                        Param = "msg: " & Msg & ", wp: " & wp & ", lp: " & lp
'                        Debug.Print "Message unknown!" & Param
                End Select
            End If
        
        Case m_TaskbarCreated
            ' IE just (re)started the taskbar!
            Call AddTrayIcon
    End Select
   Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Msghook_Message"
End Sub

Private Sub mPop_Click(Index As Integer)
    On Error GoTo EH
    Dim sRet As String
    Dim sMess As String
    Dim sDefault As String
'<----------------------------VERY IMPORTANT NOTE---------------------------->
    'IMPORTANT !!! All Menu Tasks MUST GO IN HERE
    'Otherwise MESSAGE HOOK WILL MESS UP without this Post message Call
    ' Necessary to force task switch -- see Q135788
    Call PostMessage(Me.hWnd, WM_NULL, 0, 0)
'<----------------------------VERY IMPORTANT NOTE---------------------------->
    ' React to menu choice
    Select Case Index
        Case MenuList.Show  'Open (show form)
            Me.Visible = True
            Me.WindowState = vbNormal
        Case MenuList.Hide 'Hide
            HideMe
        Case MenuList.Connect
            mbUpdateDB = True
            EnableCommandFrame False
            cmdConnect.Enabled = False
            cmdViewHistory.Enabled = False
            Start_Comm
        Case MenuList.ExitApp
            FlagShutDownFTP = True
    End Select
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub mPop_Click"
End Sub

' *****************************************
'  Private Methods
' *****************************************
Private Sub AddTrayIcon()
    On Error GoTo EH
    Dim sMess As String
   ' Initialize NOTIFYICONDATA structure
   ' and add icon to tray.
   With m_NID
      .cbSize = Len(m_NID)
      .hWnd = Msghook.HwndHook
      .uID = uID
      .uFlags = NIF_MESSAGE Or NIF_TIP Or NIF_ICON
      .uCallbackMessage = cbNotify
      .hIcon = imgList.ListImages(PicList.FTP01).Picture
      .szTip = Me.Caption & Chr(0)
   End With
   Call ShellNotifyIcon(NIM_ADD, m_NID)
   Exit Sub
EH:
   goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub AddTrayIcon"
End Sub



Private Sub Timer_Resize_Timer()
    On Error GoTo EH
    Dim lH As Long
    Dim lW As Long
    Dim lNewWidth As Long
    
    Timer_Resize.Enabled = False
    mbResize = True
    If Me.Height < FORM_H Then
        Me.Height = FORM_H
    End If
    
    If Me.Width < FORM_W Then
        Me.Width = FORM_W
    End If
    
    lH = Me.Height
    lW = Me.Width
    
    'framProcess
    framProcess.Height = lH - framProcess_H
    framProcess.Width = lW - framProcess_W
    lblProcess.Width = lW - lblProcess_W
    imgEncrypt.Left = lW - imgEncrypt_L
    lstProcess.Width = lW - lstProcess_W
    lstProcess.Height = lH - lstProcess_H
    
    'framCommands
    framCommands.Top = lH - framCommands_T
    framCommands.Width = lW - framCommands_W
    cmdConnect.Top = lH - cmdConnect_T
    cmdConnect.Left = lW - cmdConnect_L
    cmdCancelPhotoAttachUL.Top = lH - cmdCancelPhotoAttachUL_T
    cmdCancelPhotoAttachUL.Left = lW - cmdCancelPhotoAttachUL_L
    cmdHide.Left = lW - cmdHide_L
    
    
    
    VisibleFrames True
    
    mbResize = False
    Exit Sub
EH:
    Err.Clear
    Resume Next
End Sub

Private Sub TimerMsg_Timer()
    On Error GoTo EH
    Dim sMsg As String
    Dim sScreenPos As String
    Dim sCurCar As String
    Dim sCurCat As String
    Dim sCurCatDir As String
    Dim sTime As String
    Dim bCloseCurrent As Boolean
    Dim bOpenCurrent As Boolean
    Dim bFTPCommandComplete As Boolean
    Dim lHwnd As Long
    
    'DataBase Upgrade Vars, SP installation
    Dim sECSPLexePath As String
    Dim sMyCommandStr As String
    Dim sECFTPListPath As String
    Dim sECFTPListData As String  'used to build the data for ECFTP.lst File
    
    Static lPic As Long
    
    'Check the Connection
    If mbConnected Then
        lPic = lPic + 1
        m_NID.hIcon = imgList.ListImages(lPic).Picture
        If lPic = 2 Then
            lPic = 0
        End If
        Call ShellNotifyIcon(NIM_MODIFY, m_NID)
    End If
    
    'Check for Commands sent to the Registry
    sMsg = GetSetting(App.EXEName, "MSG", "COMMAND", vbNullString)
    
    Select Case sMsg
        Case "SHUT_DOWN_FTP"
            FlagShutDownFTP = True
        Case "SET_FOCUS"
            sScreenPos = sMsg
        Case "SHOW_FTP"
            Me.Visible = True
            Me.WindowState = vbNormal
        Case "HIDE_FTP"
            Me.Visible = False
        Case "CLOSE_CURRENT"
            bCloseCurrent = True
        Case "OPEN_CURRENT"
            bOpenCurrent = True
            
    End Select
    
    'Check For DataBase Upgrade
    If mbDatabaseUpgrade Then
        FlagShutDownFTP = True
        Me.Visible = False
        'Build List of Installed Regesettings, Documents, and Applications
        sECFTPListPath = goUtil.gsInstallDir & "\" & goUtil.utGetTickCount & "_ECFTP.lst"
        sECSPLexePath = goUtil.gsInstallDir & "\ECSPL.exe"
        sMyCommandStr = " " & sECFTPListPath
        'DataBase
        sECFTPListData = sECFTPListData & goUtil.SPPath & "DataBase\SP\" & msDBSPName & vbCrLf
        'Main Util
        sECFTPListData = sECFTPListData & goUtil.SPPath & "Application\SP\" & msMainUtilSPName & vbCrLf
        'Main ARV
        sECFTPListData = sECFTPListData & goUtil.SPPath & "Application\SP\" & msMainARVSPName & vbCrLf
        'Main EXE
        sECFTPListData = sECFTPListData & goUtil.SPPath & "Application\SP\" & msMainEXESPName & vbCrLf
        'Main FTP EXE
        sECFTPListData = sECFTPListData & goUtil.SPPath & "Application\SP\" & msMainFTPEXESPName
        
        goUtil.utSaveFileData sECFTPListPath, sECFTPListData
        Shell sECSPLexePath & sMyCommandStr, vbNormalFocus

    End If
    
    'Clear the COMMAND
    SaveSetting App.EXEName, "MSG", "COMMAND", vbNullString
    
    'Connect Every hour on the half hour
    sTime = Format(Now(), "HH:MM")
    If InStr(1, sTime, ":30", vbTextCompare) > 0 Then
        If msLastTime <> sTime Then
            msLastTime = sTime
            Start_Comm
        End If
    End If
    
    'Check Pos
    If Me.Visible And Me.WindowState <> vbMinimized Then
        'If the user moves the navigator to left or right then reset the
        'the startup to left or right
        Select Case UCase(sScreenPos)
            Case "SET_FOCUS"
                On Error Resume Next
                Me.SetFocus
                If Err.Number <> 0 Then
                    Err.Clear
                End If
                On Error GoTo EH
        End Select
    End If
    
    'Only execute these commands if not currently in the middle of a connection
    If Not mbConnected Then
        If bCloseCurrent Then
            
            If Not goUtil.goCurCarList Is Nothing Then
                TimerMsg.Enabled = False
                goUtil.goCurCarList.CLEANUP
                Set goUtil.goCurCarList = Nothing
                goUtil.CloseCurDB
                SaveSetting App.EXEName, "MSG", "COMMAND_COMPLETE", True
                'Wait for EasyClaim to give OPEN_CURRENT Command
                Do
                    DoEvents
                    Sleep 100
                    sMsg = GetSetting(App.EXEName, "MSG", "COMMAND", vbNullString)
                    If sMsg = "OPEN_CURRENT" Then
                        bOpenCurrent = True
                    End If
                    lHwnd = goUtil.utFindWindowPartial("Easy Claim Navigator", FwpStartsWith, False, False)
                    If lHwnd = 0 Then
                        Exit Do
                    End If
                Loop Until bOpenCurrent
                TimerMsg.Enabled = True
                
            End If
        Else
            'Update Current Car and Cat
            'these Reg Settings are set in the Main App not in the FTP app ...Private Function ExecuteNode()
            sCurCar = GetSetting(goUtil.gsMainAppExeName, "GENERAL", "CURRENT_CAR", goUtil.gsCurCar)
            sCurCat = GetSetting(goUtil.gsMainAppExeName, "GENERAL", "CURRENT_CAT", goUtil.gsCurCat)
            sCurCatDir = GetSetting(goUtil.gsMainAppExeName, "DIR", "CURRENT_CAT_DIR", goUtil.gsCurCatDir)
            If (bOpenCurrent Or goUtil.gsCurCatDir <> sCurCatDir) Or sCurCatDir = vbNullString Then
                If Not goUtil.goCurCarList Is Nothing Then
                    goUtil.goCurCarList.CLEANUP
                    Set goUtil.goCurCarList = Nothing
                End If
                'See if the Current Cat and car have been closed
                If sCurCar = vbNullString Or sCurCat = vbNullString Or sCurCatDir = vbNullString Then
                    'Change the Caption to include Current info
'                    If Me.Caption <> "Communications Status " & "(Cat not Selected)" Then
'                        Me.Caption = "Communications Status " & "(Cat not Selected)"
'                        'Also let the Icon know it
'                        m_NID.szTip = Me.Caption & Chr(0)
'                        Call ShellNotifyIcon(NIM_MODIFY, m_NID)
'                    End If
                Else
'                    goUtil.gsCurCatDir = sCurCatDir
'                    SetCarrierGlobalObjects sCurCar
'                    goUtil.SetCurDB goUtil.gsMainAppExeName, sCurCar, goUtil.gsCurCatDir & "\" & sCurCar & ".db", goUtil.gsInstallDir
'                    goUtil.SetUtilObject goUtil
'                    'Change the Caption to include Current info
'                    Me.Caption = "Communications Status " & "(" & sCurCar & ", " & sCurCat & ")"
'                    'Also let the Icon know it
'                    m_NID.szTip = Me.Caption & Chr(0)
'                    Call ShellNotifyIcon(NIM_MODIFY, m_NID)
'                    If bOpenCurrent Then
'                        SaveSetting App.EXEName, "MSG", "COMMAND_COMPLETE", True
'                    End If
                End If
                
            End If
        End If
        
        'Shut Down must be last Since it will Terminate Util object
        If FlagShutDownFTP Then
            TimerMsg.Enabled = False
            ShutdownFTP
        End If
    End If
    
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub TimerMsg_Timer"
End Sub

Public Sub ShutdownFTP()
    On Error GoTo EH
    Dim lCount As Long
    
    Unload Me
    'CLean Utility Object HERE !!!
    'Make sure it is the last thing to go bye bye
    If Not goUtil Is Nothing Then
        goUtil.CLEANUP
        Set goUtil = Nothing
    End If
    SaveSetting App.EXEName, "MSG", "COMMAND", "SHUT_DOWN_COMPLETE"
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub ShutdownFTP"
End Sub

Public Function SetCarrierGlobalObjects(psCar As String) As Boolean
    On Error GoTo EH
    'Global Collection
    Dim colGlobalObjects As Collection
    
    
    Set goUtil.goCurCarList = CreateObject(goUtil.gsCarPrefix & psCar & ".clsLists")
    Set goUtil.goECKeyBoardList = New V2ECKeyBoard.clsLists
    Set goUtil.gARV = New V2ARViewer.clsARViewer
    Set goUtil.goProgForm = New V2ECKeyBoard.clsProgForm
    
    Set colGlobalObjects = New Collection
    colGlobalObjects.Add goUtil, "goUtil"
    
    
    goUtil.goCurCarList.SetGlobalObjects colGlobalObjects
    goUtil.gARV.SetGlobalObjects colGlobalObjects
    goUtil.SetGlobalObjects colGlobalObjects
    
    Set colGlobalObjects = Nothing
    
    SetCarrierGlobalObjects = True
    Exit Function
EH:
    SetCarrierGlobalObjects = False
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function SetCarrierGlobalObjects"
End Function

Public Function SaveAndClearHistory(poList As Object) As Boolean
    On Error GoTo EH
    
    Dim sHistory As String
    Dim FileName As String
    Dim oListBox As ListBox
    Dim lCount As Long
    Dim sMess As String
    
    If TypeOf poList Is ListBox Then
        Set oListBox = poList
        'Only save History if there is is stuff in the List
    
        If oListBox.ListCount > 0 Then
            For lCount = 0 To oListBox.ListCount - 1
                sHistory = sHistory & oListBox.List(lCount) & vbCrLf
            Next
            'Clear the items
            oListBox.Clear
            If Not goUtil.utFileExists(msFTPLogPath, True) Then
                sMess = goUtil.utMakeDir(msFTPLogPath)
            End If
            
            If sMess = vbNullString Then
                FileName = msFTPLogPath & "\" & Format(Now(), "YYMMDD") & "_FTP.Log"
                sHistory = goUtil.utGetFileData(FileName) & vbCrLf & Now() & String(90, "-") & vbCrLf & sHistory
                goUtil.utSaveFileData FileName, sHistory
            End If
        End If
    End If
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function SaveAndClearHistory"
End Function

Public Function HideMe() As Boolean
    On Error GoTo EH
    Dim lHwnd As Long
    SaveAndClearHistory lstProcess
    Me.Visible = False
    'If the Main app is there set focus to it
    lHwnd = goUtil.utFindWindowPartial("Easy Claim Navigator", FwpStartsWith, False, False)
    If lHwnd > 0 Then
        goUtil.utLookForWindow "Easy Claim Navigator", 1
    End If

    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function HideMe"
End Function

Public Function EnableCommandFrame(pbEnable As Boolean) As Boolean
    On Error GoTo EH
    Dim oControl As Control
    
    framCommands.Enabled = pbEnable
    
    For Each oControl In Me.Controls
        If Not TypeOf oControl Is ImageList _
                And Not TypeOf oControl Is Timer _
                And Not TypeOf oControl Is Msghook _
                And Not TypeOf oControl Is FtpXCtl.FtpXCtl _
                And Not TypeOf oControl Is Menu Then

            If oControl.Container.Name = framCommands.Name Then
                If Not TypeOf oControl Is ProgressBar And Not TypeOf oControl Is Label Then
                    oControl.Enabled = pbEnable
                End If
            End If
        End If
    Next
    EnableCommandFrame = True
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function EnableCommandFrame"
End Function

Public Function GetStatusID(psStatusName As String, _
                            Optional pbCheckAlias As Boolean, _
                            Optional psStatusAlias As String) As String
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim RS As ADODB.Recordset
    Dim sSQL As String
    
    Set oConn = New ADODB.Connection
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    Set RS = New ADODB.Recordset
    
    sSQL = "SELECT  [StatusID], "
    sSQL = sSQL & "[Status], "
    sSQL = sSQL & "[StatusAlias] "
    sSQL = sSQL & "FROM Status "
    sSQL = sSQL & "WHERE [IsDeleted] = 0 "
    sSQL = sSQL & "AND [Status] = '" & goUtil.utCleanSQLString(psStatusName) & "' "
    
    RS.CursorLocation = adUseClient
    RS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set RS.ActiveConnection = Nothing
    
    If RS.RecordCount = 1 Then
        If pbCheckAlias Then
            psStatusAlias = RS.Fields("StatusAlias").Value
        End If
        GetStatusID = RS.Fields("StatusID").Value
    End If
    
CLEAN_UP:
    Set RS = Nothing
    Set oConn = Nothing
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function GetStatusID"
End Function

Private Sub SetFireWallSettings()
    On Error GoTo EH
    Dim sRet As String
    Dim lRet As Long
    
    sRet = GetSetting(App.EXEName, "FIREWALL_SETTINGS", "FireWallType", "0")
    If Not IsNumeric(sRet) Then
        sRet = "0"
    End If
    lRet = CLng(sRet)
    If lRet < 0 Then
        lRet = 0
    End If
    EC_FTP.FirewallType = lRet
    sRet = GetSetting(App.EXEName, "FIREWALL_SETTINGS", "FireWallPort", "0")
    If Not IsNumeric(sRet) Then
        sRet = "0"
    End If
    lRet = CLng(sRet)
    If lRet < 0 Then
        lRet = 0
    End If
    EC_FTP.FirewallPort = lRet
    sRet = GetSetting(App.EXEName, "FIREWALL_SETTINGS", "FireWallHost", "")
    EC_FTP.FirewallHost = sRet
    sRet = GetSetting(App.EXEName, "FIREWALL_SETTINGS", "FireWallLogonName", "")
    EC_FTP.FirewallLogonName = sRet
    sRet = GetSetting(App.EXEName, "FIREWALL_SETTINGS", "FireWallPassword", "")
    EC_FTP.FirewallPassword = sRet
    
    Exit Sub
EH:
   goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub SetFireWallSettings"
End Sub

Private Sub ShowNotePadMess()
    On Error GoTo EH
    Dim sPath As String
    Dim sSelFile As String
    Dim lFileSize As Long
    Dim sMyFilter As String
    Dim sFile As String
    Dim itmX As ListItem
    Dim lRet As Long
    Dim sFileName As String
    
    sFileName = "ECFTP*.Log"

    sPath = goUtil.gsInstallDir & "\ErrorLog\"

    sFile = Dir(sPath & sFileName, vbNormal)
    
    If sFile = vbNullString Then
        lblMess.Caption = vbNullString
    Else
        lblMess.Caption = "ERROR CONNECT AGAIN!"
    End If
    
    If goUtil.utFileExists(sPath & sFile) Then
        lRet = goUtil.utShellExecute(GetDesktopWindow, "OPEN", sPath & sFile, vbNullString, App.Path, vbNormalFocus, False, False, True)
    End If

    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private SubShowNotePadMess"
End Sub

