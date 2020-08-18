VERSION 5.00
Object = "{307C5043-76B3-11CE-BF00-0080AD0EF894}#1.0#0"; "MsgHoo32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmECTray 
   AutoRedraw      =   -1  'True
   Caption         =   "Easy Claim Navigator"
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   600
   ClientWidth     =   3495
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmECTray.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   3495
   Begin MSComctlLib.ImageList imgExplor 
      Left            =   2880
      Top             =   5655
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   -2147483647
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   44
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmECTray.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmECTray.frx":075E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmECTray.frx":0BB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmECTray.frx":1006
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmECTray.frx":145A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmECTray.frx":18AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmECTray.frx":1D02
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmECTray.frx":229C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmECTray.frx":26F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmECTray.frx":2B42
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmECTray.frx":2F94
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmECTray.frx":33E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmECTray.frx":383A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmECTray.frx":3C8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmECTray.frx":3FA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmECTray.frx":42C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmECTray.frx":4712
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmECTray.frx":4B64
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmECTray.frx":4FB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmECTray.frx":5408
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmECTray.frx":585A
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmECTray.frx":5CAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmECTray.frx":77FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmECTray.frx":7958
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmECTray.frx":7DAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmECTray.frx":81FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmECTray.frx":864E
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmECTray.frx":8AA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmECTray.frx":8EF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmECTray.frx":9344
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmECTray.frx":965E
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmECTray.frx":9978
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmECTray.frx":9C92
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmECTray.frx":9FAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmECTray.frx":A2C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmECTray.frx":A5E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmECTray.frx":A8FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmECTray.frx":AC14
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmECTray.frx":AF2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmECTray.frx":B248
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmECTray.frx":B562
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmECTray.frx":B9B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmECTray.frx":BCCE
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmECTray.frx":BFE8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MsghookLib.Msghook Msghook 
      Left            =   2160
      Top             =   5655
      _Version        =   65536
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
   End
   Begin VB.Timer TimerMsg 
      Interval        =   500
      Left            =   3000
      Top             =   960
   End
   Begin VB.Timer Timer_SpinMe 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   2640
      Top             =   960
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   1560
      Top             =   5640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmECTray.frx":C582
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmECTray.frx":C89C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmECTray.frx":CBB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmECTray.frx":CED0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picCurCat 
      Height          =   855
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   3435
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   3495
      Begin VB.CommandButton cmdRefresh 
         Height          =   570
         Left            =   600
         Picture         =   "frmECTray.frx":D1EA
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Refresh"
         Top             =   240
         Width           =   615
      End
      Begin VB.PictureBox picCurCatTop 
         BackColor       =   &H8000000E&
         Height          =   280
         Left            =   0
         ScaleHeight     =   225
         ScaleWidth      =   3435
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   -45
         Width           =   3495
         Begin VB.Label lblCurCatTop 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000017&
            Height          =   225
            Left            =   0
            TabIndex        =   2
            Top             =   10
            Width           =   3495
            WordWrap        =   -1  'True
         End
      End
      Begin VB.CommandButton cmdConnect 
         Height          =   570
         Left            =   0
         Picture         =   "frmECTray.frx":D334
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Connect"
         Top             =   240
         Width           =   615
      End
      Begin VB.Image imgCurCat 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1275
         Stretch         =   -1  'True
         Top             =   390
         Width           =   285
      End
      Begin VB.Label lblCurCat 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000017&
         Height          =   465
         Left            =   1710
         TabIndex        =   5
         Top             =   300
         Width           =   1695
         WordWrap        =   -1  'True
      End
   End
   Begin MSComctlLib.StatusBar ECStatBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   6255
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   0
            Enabled         =   0   'False
            Object.Width           =   5654
            MinWidth        =   18
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.TreeView ECTree 
      Height          =   5400
      Left            =   0
      TabIndex        =   6
      Top             =   880
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   9525
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   0
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "imgExplor"
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuSetUpNewCat 
         Caption         =   "&Setup New Cat"
         Visible         =   0   'False
      End
      Begin VB.Menu BarSetUpNewCat 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSpellChecker 
         Caption         =   "Spell &Checker"
         Begin VB.Menu mnuEditCustDic 
            Caption         =   "&Edit Custom Dictionary"
         End
      End
      Begin VB.Menu BarSpellChecker 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintSetUp 
         Caption         =   "&Printer"
         Begin VB.Menu mnuSetPrinterManually 
            Caption         =   "Setup &Printer (Windows Default)"
         End
         Begin VB.Menu mnuUseWinDefaultPrinter 
            Caption         =   "&Use Windows Default Printer"
         End
      End
      Begin VB.Menu BarPrinter 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDataBaseUtility 
         Caption         =   "&Database Utility"
         Begin VB.Menu mnuCompactRepair 
            Caption         =   "&Compact Repair \ Backup Database"
         End
      End
      Begin VB.Menu BarDataBaseUtility 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPref 
         Caption         =   "Prefe&rences"
      End
      Begin VB.Menu BarPreferences 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHideAll 
         Caption         =   "&Hide All"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
      Begin VB.Menu mnuSupport 
         Caption         =   "&Support"
      End
      Begin VB.Menu BarAbout 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAboutEasyClaim 
         Caption         =   "&About EasyClaim"
      End
      Begin VB.Menu BarConcentration 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConcentration 
         Caption         =   "&Concentration"
      End
   End
   Begin VB.Menu mPopUp 
      Caption         =   "Pop"
      Visible         =   0   'False
      Begin VB.Menu mPop 
         Caption         =   "&Show"
         Index           =   0
      End
      Begin VB.Menu mPop 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mPop 
         Caption         =   "&Hide"
         Index           =   2
      End
      Begin VB.Menu mPop 
         Caption         =   "Hide &All"
         Index           =   3
      End
      Begin VB.Menu mPop 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mPop 
         Caption         =   "&Reset Form Positions"
         Index           =   5
      End
      Begin VB.Menu mPop 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mPop 
         Caption         =   "&Exit"
         Index           =   7
      End
   End
End
Attribute VB_Name = "frmECTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Enum PicList
   ECHurc = 1
   ECHurc2
End Enum

Public Enum MenuList
    Show = 0
    BarHide
    Hide
    Hide_All
    BarResetFormPos
    ResetFormPos
    BarExit
    ExitApp
End Enum


' User defined constant values
Private Const cbNotify As Long = &H4000
Private Const uID As Long = 61860
'Carrier List Command
Private Const CAR_LIST_COMMAND As String = "CarList+"

'More Pictures
Private Const PIC_TREE_IE_EXPLORE As Long = 44

' Member variables
Private m_NID As NOTIFYICONDATA
Private m_TaskbarCreated As Long
'Spinner object
Private moSpinner As Object
Private mbLoadingRegForm As Boolean
Private mbEndECTray As Boolean

'Timed Flags
Private mbShutDownEasyClaim As Boolean
Private mbCheckPrinter As Boolean

'Tree View
Private mNodX As MSComctlLib.Node

'Send to Xactimate
Private mbSendToXactimate As Boolean
'Loss Reports
Private mbLoadingLossReports As Boolean
'FTP Stuff
Private mbFTPConnected As Boolean
'paths For Easy Claim
Private msFTPDLPath As String 'Download Path
Private msFTPULPath As String 'Upload Path
Private msAttachReposPath As String
Private msPhotoReposPath As String
Private msSPPath As String 'Software Package path
Private mFrmCancel As frmCancel 'Send to Xactimate Cancel Form
Private mbDeleting As Boolean


Private Property Get msClassName() As String
    msClassName = Me.Name
End Property

Public Property Let FlagSendToXactimate(pbFlag As Boolean)
    mbSendToXactimate = pbFlag
End Property
Public Property Get FlagSendToXactimate() As Boolean
    FlagSendToXactimate = mbSendToXactimate
End Property

Public Property Let FlagLoadingLossReports(pbFlag As Boolean)
    mbLoadingLossReports = pbFlag
End Property
Public Property Get FlagLoadingLossReports() As Boolean
    FlagLoadingLossReports = mbLoadingLossReports
End Property

Public Property Let FlagShutDownEasyClaim(pbFlag As Boolean)
    mbShutDownEasyClaim = pbFlag
End Property
Public Property Get FlagShutDownEasyClaim() As Boolean
    FlagShutDownEasyClaim = mbShutDownEasyClaim
End Property

Public Property Let FlagCheckPrinter(pbFlag As Boolean)
    mbCheckPrinter = pbFlag
End Property
Public Property Get FlagCheckPrinter() As Boolean
    FlagCheckPrinter = mbCheckPrinter
End Property

Public Property Let EndECTray(pbFlag As Boolean)
    mbEndECTray = pbFlag
End Property
Public Property Get EndECTray() As Boolean
    EndECTray = mbEndECTray
End Property

Public Property Let Spinner(poSpinner As Object)
    Set moSpinner = poSpinner
End Property
Public Property Set Spinner(poSpinner As Object)
    Set moSpinner = poSpinner
End Property
Public Property Get Spinner() As Object
    Set Spinner = moSpinner
End Property

Private Sub cmdConnect_Click()
    On Error GoTo EH
    goUtil.utShowFTP goUtil.gsInstallDir & "\ECFTP.EXE " & Command$
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdConnect_Click"
End Sub

Private Sub cmdRefresh_Click()
    LoadTree
End Sub

Private Sub ECTree_Collapse(ByVal Node As MSComctlLib.Node)
    On Error GoTo EH
    
    If Node.Image = PicTree.A10_OpenFolder Then
        Node.Image = PicTree.A09_ClosedFolder
    End If
    
    If Node.Image = PicTree.A15_OpenHurc Then
       Node.Image = PicTree.A14_ClosedHurc
    End If
    
    'also check for Inactive Cat
     If Node.Image = PicTree.A43_OpenHurc_Inactive Then
       Node.Image = PicTree.A42_closedHurc_Inactive
    End If
    
    Exit Sub
EH:
    
End Sub

Private Sub ECTree_DblClick()
    On Error GoTo EH
    If goUtil Is Nothing Then
        Exit Sub
    End If
    If goUtil.gbValidLic Then
        ExecuteNode
    ElseIf StrComp(mNodX.Key, App.EXEName & "|GlobalPref", vbTextCompare) = 0 Then
        ExecuteNode
    ElseIf StrComp(mNodX.Key, App.EXEName & "|About", vbTextCompare) = 0 Then
        ExecuteNode
    ElseIf StrComp(mNodX.Key, App.EXEName & "|Support", vbTextCompare) = 0 Then
        ExecuteNode
    Else
        If StrComp(mNodX.Key, App.EXEName, vbTextCompare) <> 0 Then
            MsgBox "This option is unavailable.  License Expired!", vbExclamation, "License Expired!"
        End If
    End If
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub ECTree_DblClick"
End Sub

Private Sub ECTree_Expand(ByVal Node As MSComctlLib.Node)
    On Error GoTo EH
    If Node.Image = PicTree.A09_ClosedFolder Then
        Node.Image = PicTree.A10_OpenFolder
    End If
    
    If Node.Image = PicTree.A14_ClosedHurc Then
       Node.Image = PicTree.A15_OpenHurc
    End If
    
    'Also check for inactive Cat
    If Node.Image = PicTree.A42_closedHurc_Inactive Then
       Node.Image = PicTree.A43_OpenHurc_Inactive
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub ECTree_Expand"
End Sub

Private Sub ECTree_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo EH
    
    Select Case KeyCode + Shift
        Case vbKeyH + 1
            HideMe
        Case vbKeyEscape
            Me.WindowState = vbMinimized
        Case vbKeyReturn
            If goUtil.gbValidLic Then
                If Not ECTree.SelectedItem Is Nothing Then
                    Set mNodX = ECTree.SelectedItem
                    ExecuteNode
                End If
            Else
'                Load CommStatus
'                CommStatus.Show vbModeless
            End If
    End Select
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub ECTree_KeyDown"
End Sub

Private Sub ECTree_NodeClick(ByVal Node As MSComctlLib.Node)
    On Error GoTo EH
    
    Set mNodX = Node
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub ECTree_NodeClick"
End Sub

Private Sub Form_Activate()
    On Error GoTo EH
    
    SaveSetting App.EXEName, "MSG", "COMMAND", "SET_FOCUS"
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_Activate"
End Sub


Private Sub Form_Load()
    On Error GoTo EH
    Dim sMess As String
    Dim frmReg As frmRegForm
    Dim frmLog As frmLogOn
    Dim sUserName As String
    Dim sPassword As String
    
    ' Don't want to be visible initially!
    Me.Visible = False
    Me.Caption = App.EXEName & " Navigator"
    App.Title = Me.Caption
    If goUtil Is Nothing Then
        Exit Sub
    End If
    goUtil.utFormWinRegPos goUtil.gsMainAppEXEName, Me
    
    'Set the Install dir here !
    goUtil.gsInstallDir = GetSetting(App.EXEName, "Dir", "INSTALL_DIR", App.Path)
    'Need to save it in case this is the first time running (Using the App.path as default)
    SaveSetting App.EXEName, "Dir", "INSTALL_DIR", goUtil.gsInstallDir
    
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
                SaveSetting App.EXEName, "MSG", "COMMAND", "SHUT_DOWN_EASYCLAIM"
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
    Call AddTrayIcon
    
    '----------------Build Directory Paths--------------------
    goUtil.BuildPathsEasyClaim
    msFTPDLPath = goUtil.FTPDLPath
    msFTPULPath = goUtil.FTPULPath
    msSPPath = goUtil.SPPath
    msAttachReposPath = goUtil.AttachReposPath
    msPhotoReposPath = goUtil.PhotoReposPath
    '----------------END Build Directory Paths--------------------
    
    'Load the Tree
    LoadTree
    If mbShutDownEasyClaim Then
        Exit Sub
    End If
                
    'Show the Registration Form
    mbLoadingRegForm = True
    If Trim(Command$) <> "EZas123" & Format(Now, "DDYYMM") Then
        Set frmReg = New frmRegForm
        Load frmReg
        frmReg.Show
        Do Until Not frmReg.Visible
            DoEvents
            Sleep 100
        Loop
        Unload frmReg
        Set frmReg = Nothing
        'If this is the very first time running the Lic then
        'Unload until they put in the 10 day Lic.
        'Once this is done then any further Lic renewel will be
        'accomplished by connecting to the Server.
        sUserName = goUtil.utGetECSCryptSetting("ECS", "WEB_SECURITY", "USER_NAME")
        sPassword = goUtil.utGetECSCryptSetting("ECS", "WEB_SECURITY", "PASSWORD")
        If goUtil.gbInitLic And Not goUtil.gbValidLic Then
            MsgBox "Invalid TempCode! " & vbCrLf & vbCrLf & App.EXEName & " will exit.", vbExclamation, "Invalid Code"
            FlagShutDownEasyClaim = True
            Exit Sub
        ElseIf sUserName = vbNullString Or sPassword = vbNullString Then
            'If the first time running then show the Prefs screen
            MsgBox "Please complete the Preferences screen!", vbExclamation + vbOKOnly, "Preferences Set up"
            EasyClaimCommand "GlobalPref"
            goUtil.utAlwaysOnTop frmPreferences, True
            'Still show the Navigator tree even if Lic is expired.
            'Adjuster will only be allowed to connect to renew Lic if approved
            If Not goUtil.gbValidLic Then
                mnuFile.Enabled = False
                Me.Caption = Me.Caption & " (License Expired!)"
            End If
        ElseIf sUserName <> vbNullString And sPassword <> vbNullString Then
            Set frmLog = New frmLogOn
            Load frmLog
SHOW_LOGON:
            frmLog.Show vbModal
            If UCase(frmLog.txtUserName.Text) <> sUserName Or frmLog.txtPass.Text <> sPassword Then
                If frmLog.txtUserName.Text = vbNullString And frmLog.txtPass.Text = vbNullString Then
                    Unload frmLog
                    Set frmLog = Nothing
                    FlagShutDownEasyClaim = True
                    Exit Sub
                Else
                    MsgBox "Invalid User Name or Password! " & vbCrLf & vbCrLf & " Try Again.", vbExclamation, "Invalid"
                    GoTo SHOW_LOGON
                End If
            End If
            If Not goUtil.gbValidLic Then
                mnuFile.Enabled = False
                Me.Caption = Me.Caption & " (License Expired!)"
            End If
            Unload frmLog
            Set frmLog = Nothing
            ShowMe
            PosECTRAY
        Else
            'Still show the Navigator tree even if Lic is expired.
            'Adjuster will only be allowed to connect to renew Lic if approved
            If Not goUtil.gbValidLic Then
                mnuFile.Enabled = False
                Me.Caption = Me.Caption & " (License Expired!)"
            End If
            ShowMe
            PosECTRAY
        End If
    Else
        goUtil.gbValidLic = True
        Me.Caption = Me.Caption & " {DEMO MODE}"
    End If
    mbLoadingRegForm = False

    'Enable Printer Timer It monitors For User changing Default Windows Printer
    InitDefaultPrintMenu
    
   Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_Load"
End Sub

Private Sub Form_Paint()
    On Error GoTo EH
    Dim bAlwaysOnTop As Boolean
    
    bAlwaysOnTop = CBool(GetSetting(App.EXEName, "GENERAL", "ECTRAY_ALWAYS_ON_TOP", False))
    goUtil.utAlwaysOnTop Me, bAlwaysOnTop
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_Paint"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo EH
    Dim sMess As String
    If UnloadMode = vbFormControlMenu Then
        ' Just hide form if user presses Closes by X
        HideMe
        HideAll
        Cancel = True
    End If
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_QueryUnload"
End Sub

Private Sub Form_Resize()
    On Error GoTo EH
    Dim sMess As String
    If Me.Height - 1830 > 0 Then
        ECTree.Height = Me.Height - 1830
    End If
    ECTree.Width = Me.Width - 120
    picCurCatTop.Width = Me.Width - 120
    lblCurCatTop.Width = Me.Width - 120
    picCurCat.Width = Me.Width - 120
    lblCurCat.Width = Me.Width - 1920
    ShowWallPaper
    If Me.WindowState = vbNormal Then
        ShowAll
    End If
    Exit Sub
EH:
    If Err.Number > 0 Then
        Err.Clear
        Resume Next
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo EH
    Dim MyForm As Form
    If Me.Visible Then
        goUtil.utFormWinRegPos goUtil.gsMainAppEXEName, Me, True
    End If
    Call ShellNotifyIcon(NIM_DELETE, m_NID)
    
    'Stop chekcing for printer
    FlagCheckPrinter = False
    
    'Clear the Current Cat in Registry
    SaveSetting App.EXEName, "DIR", "CURRENT_CAT_DIR", vbNullString
    SaveSetting App.EXEName, "GENERAL", "CURRENT_CAR", vbNullString
    SaveSetting App.EXEName, "GENERAL", "CURRENT_CAT", vbNullString


    'Clean any forms that are still open
    'This will only close forms that are open within
    'Easy Claim Project. Any forms that are open in other DLLs
    'Will need to be closed in its CleanUp method
    'All Easy Claim Project Forms other than ECTRay
    For Each MyForm In Forms
        On Error Resume Next
        If MyForm.Name <> Me.Name Then
            Unload MyForm
            If Err.Number > 0 Then
                Err.Clear
            End If
            Set MyForm = Nothing
        End If
    Next
    
    If Not gfrmWallPaper Is Nothing Then
        Unload gfrmWallPaper
        Set gfrmWallPaper = Nothing
    End If
    
    EndECTray = True
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_Unload"
End Sub

Private Sub mnuAboutEasyClaim_Click()
    EasyClaimCommand "About"
End Sub

Private Sub mnuCompactRepair_Click()
    On Error GoTo EH
    Dim sMess As String
    Dim bFTPExists As Boolean
    
     'If the FTP Application is currently connected and download or Uploading
     'Data then can't run this at all !
    If goUtil.utFTPConnected Then
        sMess = "FTP connection is currently active." & vbCrLf & vbCrLf
        sMess = sMess & "You must wait until the current connection is finished " & vbCrLf
        sMess = sMess & "before running this utility."
        MsgBox sMess, vbExclamation + vbOKOnly, "FTP Connection Detected"
        Exit Sub
    End If
    
    'If the FTP Application is running then need to close it before
    'Running the Compact Repair.
    
    bFTPExists = goUtil.utFTPExists
    If bFTPExists Then
        sMess = "Closing FTP Application!" & vbCrLf & vbCrLf
        sMess = sMess & "FTP will restart after Compact Repair completes."
        MsgBox sMess, vbInformation + vbOKOnly, "FTP SHUTDOWN"
        goUtil.utShutDownFTP
        Sleep 2000 'Wait couple seconds
    End If
    'Running the Compact Repair.
    CompactAndRepairMainDB
    
    MsgBox "Compact Repair Complete", vbInformation + vbOKOnly, "Compact Repair"
    
    If bFTPExists Then
        goUtil.utShowFTP goUtil.gsInstallDir & "\ECFTP.EXE " & Command$
    End If
    
    Exit Sub
EH:
    Screen.MousePointer = vbNormal
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub mnuCompactRepair_Click"
End Sub

Private Sub mnuConcentration_Click()
    On Error GoTo EH
    Dim dProc As Double
    dProc = Shell(App.Path & "\Concentration.exe", vbNormalFocus)
    AppActivate dProc, True
    Exit Sub
EH:
    MsgBox "Concentration.exe not found!", vbExclamation, "Concentration Game Not Found!"
    Err.Clear
End Sub

Private Sub mnuEditCustDic_Click()
    On Error GoTo EH
    Dim pbUnloaded As Boolean
    
    goUtil.utLoadSP
    goUtil.goSP.ShowCustomDictionary vbModeless
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub mnuEditCustDic_Click"
End Sub

Private Sub mnuExit_Click()
    FlagShutDownEasyClaim = True
End Sub

Private Sub mnuHideAll_Click()
    HideMe
    HideAll
End Sub

Private Sub mnuPref_Click()
    EasyClaimCommand "GlobalPref"
End Sub

Private Sub mnuSetPrinterManually_Click()
    EasyClaimCommand "PrinterSetup"
End Sub

Private Sub mnuSetUpNewCat_Click()
    EasyClaimCommand "SetupNewCat"
End Sub

Private Sub mnuSupport_Click()
   EasyClaimCommand "Support"
End Sub

Private Sub mnuUseWinDefaultPrinter_Click()
    On Error GoTo EH
    
    If mnuUseWinDefaultPrinter.Checked Then
        mnuUseWinDefaultPrinter.Checked = False
        SaveSetting App.EXEName, "PRINTER", "INIT_WITH_WIN_DEFAULT_PRINTER", False
        FlagCheckPrinter = False
    Else
        mnuUseWinDefaultPrinter.Checked = True
        SaveSetting App.EXEName, "PRINTER", "INIT_WITH_WIN_DEFAULT_PRINTER", True
        FlagCheckPrinter = True
    End If
        
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub mnuUseWinDefaultPrinter_Click"
End Sub




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
                        ShowMe
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
                        Debug.Print "Message unknown!" & Param
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
            ShowMe
        Case MenuList.Hide 'Hide
            Me.WindowState = vbMinimized
        Case MenuList.Hide_All
            HideMe
            HideAll
        Case MenuList.ResetFormPos
            sMess = "Reset requires " & goUtil.gsAppEXEName & " to exit." & vbCrLf & vbCrLf
            sMess = sMess & "Before you click ""YES""..." & vbCrLf
            sMess = sMess & "Click ""NO"" and perform the following items:" & vbCrLf
            sMess = sMess & "1.  Save any work you have open for " & goUtil.gsAppEXEName & "." & vbCrLf
            sMess = sMess & "2.  Close any " & goUtil.gsAppEXEName & " items you see on your task bar ""Right-Click|Close""." & vbCrLf
            sMess = sMess & "3.  Run ""Reset Form Positions"" again and then click ""YES"""
            If MsgBox(sMess, vbYesNo + vbExclamation, "Reset Form Positions") = vbYes Then
                HideAll
                On Error Resume Next
                DeleteSetting goUtil.gsAppEXEName, "FORM_POSN"
                If Err.Number > 0 Then
                    Err.Clear
                End If
                On Error GoTo EH
                mPop_Click MenuList.ExitApp
            End If
        Case MenuList.ExitApp
'            sMess = "Are you sure you want to Exit " & App.EXEName & "?"
'            If MsgBox(sMess, vbQuestion + vbOKCancel, "Exit " & App.EXEName) = vbOK Then
                If Not goUtil.goProgForm Is Nothing Then
                    If Not goUtil.goProgForm.Object Is Nothing Then
                        goUtil.goProgForm.CancelMe = True
                    End If
                End If
                FlagShutDownEasyClaim = True
'            End If
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
      .hIcon = imgList.ListImages(PicList.ECHurc).Picture
      .szTip = Me.Caption & Chr(0)
   End With
   Call ShellNotifyIcon(NIM_ADD, m_NID)
   Exit Sub
EH:
   goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub AddTrayIcon"
End Sub

Public Sub ShutdownEasyClaim()
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
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub ShutdownEasyClaim"
End Sub

Private Sub Timer_SpinMe_Timer()
    On Error GoTo EH
    Static lPic As Long
    lPic = lPic + 1
    'Change the Tray Icon
    m_NID.hIcon = imgList.ListImages(lPic).Picture
    'check to see if we have a form we want to show spinning
    If Not moSpinner Is Nothing Then
        moSpinner.Picture = imgList.ListImages(lPic).Picture
    End If
    If lPic = 4 Then
        lPic = 0
    End If
    Call ShellNotifyIcon(NIM_MODIFY, m_NID)
    Exit Sub
EH:
    Err.Clear
    Timer_SpinMe.Enabled = False
End Sub

Private Sub TimerMsg_Timer()
    On Error GoTo EH
    Dim sMsg As String
    Dim sNavScreenPos As String
    
    'Check FTP Connection
    mbFTPConnected = GetSetting("ECFTP", "MSG", "CONNECTED", False)
    
    sNavScreenPos = GetSetting(App.EXEName, "GENERAL", "NAV_SCREEN_POS", "RIGHT")
    
    'Check for Commands sent to the Registry
    sMsg = GetSetting(App.EXEName, "MSG", "COMMAND", vbNullString)
    
    'Clear the COMMAND
    SaveSetting App.EXEName, "MSG", "COMMAND", vbNullString
    
    Select Case sMsg
        Case "SHOW"
            ShowMe False
        Case "SHOW_ALL"
            ShowMe
        Case "SHOW_ABOUT"
            EasyClaimCommand "About"
        Case "SHUT_DOWN_EASYCLAIM"
            FlagShutDownEasyClaim = True
        Case "SET_FOCUS"
            sNavScreenPos = sMsg
        Case "GLOBAL_PREF"
            EasyClaimCommand "GlobalPref"
        Case "LOAD_TREE"
            LoadTree
            If goUtil.gsCurCompany <> vbNullString And goUtil.gsCurCar <> vbNullString Then
                SetCarrierGlobalObjects goUtil.gsCurCompany, goUtil.gsCurCar, False
            End If
        Case "COMPACT_AND_REAPIR_MAIN_DB"
            CompactAndRepairMainDB
        Case "SHOW_FTP_IS_UPDATING_DATA"
            EasyClaimCommand "SHOW_FTP_IS_UPDATING_DATA"
        Case "HIDE_FTP_IS_UPDATING_DATA"
            EasyClaimCommand "HIDE_FTP_IS_UPDATING_DATA"
    End Select
    
    
    'Check for Commands sent from FTP
    sMsg = GetSetting(App.EXEName, "MSG", "FTP_COMMAND", vbNullString)
    
    Select Case sMsg
        Case "SHUT_DOWN_EASYCLAIM"
            FlagShutDownEasyClaim = True
        Case "GLOBAL_PREF"
            EasyClaimCommand "GlobalPref", vbModal
            
    End Select
    
    'Clear the FTP_COMMAND
    SaveSetting App.EXEName, "MSG", "FTP_COMMAND", vbNullString
    
    
    'Check Pos
    If Me.Visible And Me.WindowState <> vbMinimized Then
        'If the user moves the navigator to left or right then reset the
        'the startup to left or right
        Select Case UCase(sNavScreenPos)
            Case "RIGHT"
                If Me.left <= Screen.Width / 2 Then
                    SaveSetting App.EXEName, "GENERAL", "NAV_SCREEN_POS", "LEFT"
                    If Not gfrmWallPaper Is Nothing Then
                        gfrmWallPaper.picBack.left = 1500
                    End If
                    If goUtil.utFormExists(Forms, "frmPreferences") Then
                        frmPreferences.LoadNavPref
                    Else
                        PosECTRAY
                        gfrmECTray.ShowMe , False
                    End If
                    
                End If
            Case "LEFT"
                If Me.left >= Screen.Width / 2 Then
                    SaveSetting App.EXEName, "GENERAL", "NAV_SCREEN_POS", "RIGHT"
                    If Not gfrmWallPaper Is Nothing Then
                        gfrmWallPaper.picBack.left = 0
                    End If
                    If goUtil.utFormExists(Forms, "frmPreferences") Then
                        frmPreferences.LoadNavPref
                    Else
                        PosECTRAY
                        gfrmECTray.ShowMe , False
                    End If
                    
                End If
            Case "SET_FOCUS"
                On Error Resume Next
                goUtil.utLookForWindow Me.Caption, 1
                DoEvents
                Sleep 100
                If Me.ECTree.Visible Then
                    Me.ECTree.SetFocus
                End If
                If Err.Number > 0 Then
                    Err.Clear
                End If
                On Error GoTo EH
        End Select
        PosECTRAY
    End If
    
    'Also Check for Timed Flags
    If FlagCheckPrinter Then
        CheckPrinter
    End If
    
    'Shut Down must be last Since it will Terminate Util object
    If FlagShutDownEasyClaim Then
        'Check for other flags that will require the
        'ShutDown process to wait
        'if in the middle of deleting then have to wait
        If mbDeleting Then
            Exit Sub
        End If
        TimerMsg.Enabled = False
        ShutdownEasyClaim
    End If
    
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub TimerMsg_Timer"
End Sub

Private Sub CheckPrinter(Optional pbIgnoreDefault As Boolean)
    On Error GoTo EH
    Dim sDefaultPrinter As String
    Dim nret As Integer
    Dim sRet As String
    Dim lPos As Long
    Dim sTemp As String
    
    If pbIgnoreDefault Then
        sDefaultPrinter = GetSetting(App.EXEName, "PRINTER", "PRINTER_NAME", vbNullString)
    Else
        'Get Default Printer Name
        sDefaultPrinter = Space(255)
        nret = GetProfileString("Windows", ByVal "device", "", sDefaultPrinter, Len(sDefaultPrinter))
        'Trim it
        If nret Then
            sDefaultPrinter = left(sDefaultPrinter, InStr(sDefaultPrinter, ",") - 1)
        End If
    End If
    
    sTemp = mnuUseWinDefaultPrinter.Caption
    lPos = InStr(1, sTemp, "(", vbTextCompare)
    If lPos = 0 Then
        sTemp = sTemp & " (" & sDefaultPrinter & ")"
    Else
        sTemp = left(sTemp, lPos)
        sTemp = sTemp & sDefaultPrinter & ")"
    End If
    mnuUseWinDefaultPrinter.Caption = sTemp
    
    sRet = GetSetting(App.EXEName, "PRINTER", "PRINTER_NAME", vbNullString)
    
    If StrComp(sRet, sDefaultPrinter, vbTextCompare) = 0 Then
        Exit Sub
    Else
        goUtil.utSaveDefaultPrinterSettings App.EXEName, sDefaultPrinter
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub TimerPrinter_Timer"
End Sub

Private Sub InitDefaultPrintMenu()
    On Error GoTo EH
    Dim sUseDefault As String
    
    sUseDefault = GetSetting(App.EXEName, "PRINTER", "INIT_WITH_WIN_DEFAULT_PRINTER", vbNullString)
    
    If sUseDefault = vbNullString Then
        SaveSetting App.EXEName, "PRINTER", "INIT_WITH_WIN_DEFAULT_PRINTER", True
        mnuUseWinDefaultPrinter.Checked = True
        FlagCheckPrinter = True
    ElseIf CBool(sUseDefault) Then
        mnuUseWinDefaultPrinter.Checked = True
        FlagCheckPrinter = True
    Else
        mnuUseWinDefaultPrinter.Checked = False
        FlagCheckPrinter = False
        CheckPrinter True
    End If

    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub InitDefaultPrintMenu"
End Sub

Public Sub HideAll(Optional pbSkipHideME As Boolean)
    On Error GoTo EH
    Dim MyForm As Form
    
    goUtil.utHideAllForms Forms, Me.Name
    If Not pbSkipHideME Then
        If Me.WindowState = vbMinimized Then
            Me.Visible = False
            pbSkipHideME = True
        End If
    End If
    'Add Objects Here that may have forms that need to be Hidden
    goUtil.HideAllForms
    If Not goUtil.goCurCarList Is Nothing Then
        goUtil.goCurCarList.HideAllForms
    End If
    If Not goUtil.gARV Is Nothing Then
        goUtil.utHideAllForms goUtil.gARV.goForms
    End If
    
    If Not pbSkipHideME Then
        HideMe
    End If
    
    'Tell FTP to Hide
    SaveSetting "ECFTP", "MSG", "COMMAND", "HIDE_FTP"
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub HideAll"
End Sub

Private Sub HideMe()
    On Error GoTo EH
    Dim lCount As Long
    
    Timer_SpinMe.Enabled = True
    For lCount = 1 To 3
        DoEvents
        Sleep 100
    Next
    If Me.Visible Then
        goUtil.utFormWinRegPos goUtil.gsMainAppEXEName, Me, True
    End If
    Me.WindowState = vbMinimized
    Me.Visible = False
    Me.WindowState = vbNormal
    For lCount = 1 To 5
        DoEvents
        Sleep 100
    Next
    Timer_SpinMe.Enabled = False
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub HideMe"
End Sub

Public Sub ShowAll()
    On Error GoTo EH
    Dim MyForm As Form
    Dim sNavScreenPos As String
    
    sNavScreenPos = GetSetting(App.EXEName, "GENERAL", "NAV_SCREEN_POS", "RIGHT")
    
    goUtil.utShowAllForms Forms, Me, sNavScreenPos, Me.Name
    'Add Objects Here that may have forms that need to be shown
    goUtil.ShowAllForms Me, sNavScreenPos
    If Not goUtil.goCurCarList Is Nothing Then
        goUtil.goCurCarList.ShowAllForms Me, sNavScreenPos
    End If
    If Not goUtil.gARV Is Nothing Then
        goUtil.utShowAllForms goUtil.gARV.goForms, Me, sNavScreenPos
    End If
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub ShowAll"
End Sub

Public Sub ShowMe(Optional pbShowAll As Boolean = True, Optional pbSpinMe = True)
    On Error GoTo EH
    Dim lCount As Long
    
    If pbSpinMe Then
        Timer_SpinMe.Enabled = True
        For lCount = 1 To 3
            DoEvents
            Sleep 100
        Next
    End If
    If Not Me.Visible Then
        On Error Resume Next
        Me.WindowState = vbMinimized
        Me.Visible = True
        Me.WindowState = vbNormal
        Me.SetFocus
        If Err.Number > 0 Then
            Err.Clear
        End If
        On Error GoTo EH
        If pbSpinMe Then
            For lCount = 1 To 5
                DoEvents
                Sleep 100
            Next
        End If
        PosECTRAY
    Else
        Me.WindowState = vbNormal
        SaveSetting App.EXEName, "MSG", "COMMAND", "SET_FOCUS"
    End If
    If pbSpinMe Then
        Timer_SpinMe.Enabled = False
    End If
    If pbShowAll Then
        ShowAll
    End If
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub ShowMe"
End Sub

Public Sub PosECTRAY()
    On Error GoTo EH
    Dim sNavScreenPos As String
    
    sNavScreenPos = GetSetting(App.EXEName, "GENERAL", "NAV_SCREEN_POS", "RIGHT")
    If gfrmECTray.Visible And gfrmECTray.WindowState = vbNormal Then
        Select Case UCase(sNavScreenPos)
            Case "RIGHT"
                gfrmECTray.left = Screen.Width - gfrmECTray.Width
                gfrmECTray.top = 0
                gfrmECTray.Height = Screen.Height - goUtil.utGetTaskbarHeight  '(Account for TaskBar Height)
                
            Case "LEFT"
                gfrmECTray.left = 0
                gfrmECTray.top = 0
                gfrmECTray.Height = Screen.Height - goUtil.utGetTaskbarHeight '(Account for TaskBar Height)
        End Select
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub PosECTRAY"
End Sub

Public Sub ShowWallPaper()
    On Error GoTo EH
    Dim bUseWallPaper As Boolean
    Dim lTBHeight As Long
    
    lTBHeight = goUtil.utGetTaskbarHeight
    
    If gfrmWallPaper Is Nothing Then
        Exit Sub
    End If
    
    bUseWallPaper = CBool(GetSetting(App.EXEName, "GENERAL", "USE_WALL_PAPER", True))
    
    If bUseWallPaper Then
        gfrmWallPaper.top = 0
        gfrmWallPaper.left = 0
        gfrmWallPaper.Width = Screen.Width
        gfrmWallPaper.Height = Screen.Height - IIf(lTBHeight = 0, 10, lTBHeight)
        LoadWallPaperImage
        gfrmWallPaper.Visible = True
    Else
        gfrmWallPaper.Visible = False
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub ShowWallPaper"
End Sub

Public Sub LoadWallPaperImage()
    On Error GoTo EH
    Dim sWallPicPath As String
    
    sWallPicPath = GetSetting(App.EXEName, "DIR", "WALL_PAPER_IMAGE", vbNullString)
    
    If goUtil.utFileExists(sWallPicPath) Then
        If Not gfrmWallPaper Is Nothing Then
            If gfrmWallPaper.LastPicPath <> sWallPicPath Then
                gfrmWallPaper.imgWallPic.Visible = False
                gfrmWallPaper.imgWallPic.Picture = LoadPicture(sWallPicPath)
                POSWallPic
            End If
        End If
    Else
        If Not gfrmWallPaper Is Nothing Then
            gfrmWallPaper.imgWallPic.Picture = Nothing
        End If
    End If
    
    gfrmWallPaper.LastPicPath = sWallPicPath
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub LoadWallPaperImage"
End Sub

Public Sub POSWallPic()
    On Error GoTo EH
    gfrmWallPaper.imgWallPic.Visible = False
    gfrmWallPaper.imgWallPic.left = 0
    gfrmWallPaper.imgWallPic.top = 0
    gfrmWallPaper.imgWallPic.Width = gfrmWallPaper.Width
    gfrmWallPaper.imgWallPic.Height = gfrmWallPaper.Height
    gfrmWallPaper.imgWallPic.Visible = True
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub POSWallPic"
End Sub

Public Function LoadTree() As Boolean
    On Error GoTo EH
    Dim sMess As String
    Dim sDBSPName As String
    Dim lRet As Long
    Dim Nodx As Node
    Dim sDir As String
    Dim sTemplateDir As String
    Dim sValue As String
    Dim sTemp As String
    Dim sCompanyRelative As String
    Dim sCarRelative As String
    Dim sCatRelative As String
    Dim sRelative As String
    Dim sKey As String
    
    Dim colCompany As Collection
    Dim sCompany As String
    Dim sCompanyName As String
    Dim sCompanyDBFolderName As String
    Dim vCompany As Variant
    
    Dim colCar As Collection
    Dim sCar As String
    Dim sCarName As String
    Dim sCarDBName As String
    Dim vCar As Variant
    
    Dim colCat As Collection
    Dim sCat As String
    Dim sCatName As String
    Dim bActiveCat As Boolean
    Dim vCat As Variant
    
    'RECYCLEBIN
    Dim bAddedDeleteAllFromRecycleBin As Boolean
    Dim sAddedCompanyRecycleBin As String
    Dim sAddedCarRecycleBin As String
    
    Dim sSendCatToDisk As String
    Dim sGetCatFromDisk As String
    Dim sTreeLastKey As String
    Dim oCar As V2ECKeyBoard.clsCarLists
    'Used to get Main DB info
    Dim RS As DAO.Recordset
    Dim CompanyRS As DAO.Recordset
    Dim ClientCompanyRS As DAO.Recordset
    Dim ClientCompanyCatRS As DAO.Recordset
    
    Dim sSQL As String
    
    'Clear the Tree
    ECTree.Visible = False
    ECTree.Nodes.Clear
  
    'Add the Main Icons here
    Set Nodx = ECTree.Nodes.Add(, , goUtil.gsAppEXEName, goUtil.gsAppEXEName, PicTree.A10_OpenFolder)
    Nodx.Selected = True
'    Set Nodx = ECTree.Nodes.Add(goUtil.gsAppEXEName, tvwChild, goUtil.gsAppEXEName & "|SetupNewCat", "Setup New Cat", PicTree.A16_SetUpNewCat)
    Set Nodx = ECTree.Nodes.Add(goUtil.gsAppEXEName, tvwChild, goUtil.gsAppEXEName & "|Assignments", "Assignments", PIC_TREE_IE_EXPLORE)
    Set Nodx = ECTree.Nodes.Add(goUtil.gsAppEXEName, tvwChild, goUtil.gsAppEXEName & "|PrinterSetup", "Printer Setup", PicTree.A07_Printer)
    Set Nodx = ECTree.Nodes.Add(goUtil.gsAppEXEName, tvwChild, goUtil.gsAppEXEName & "|GlobalPref", "Preferences", PicTree.A12_GlobalPref)
    Set Nodx = ECTree.Nodes.Add(goUtil.gsAppEXEName, tvwChild, goUtil.gsAppEXEName & "|Support", "Support", PicTree.A11_Support)
    Set Nodx = ECTree.Nodes.Add(goUtil.gsAppEXEName, tvwChild, goUtil.gsAppEXEName & "|About", "About " & goUtil.gsAppEXEName, PicTree.A13_HandShake)
'    'Get Cat From Disk
'    sGetCatFromDisk = goUtil.gsAppEXEName & "|GetCatFromDisk"
'    Set Nodx = ECTree.Nodes.Add(goUtil.gsAppEXEName, tvwChild, sGetCatFromDisk, "Get Cat From Disk", PicTree.A27_GetCatFromDisk)
'    'Add Options to Get Cat From Disk
'    Set Nodx = ECTree.Nodes.Add(sGetCatFromDisk, tvwChild, sGetCatFromDisk & "|GetCatBackup", "Get Backup", PicTree.A28_GetCatBackup)
'    Set Nodx = ECTree.Nodes.Add(sGetCatFromDisk, tvwChild, sGetCatFromDisk & "|ImportCat", "Import", PicTree.A29_ImportCat)
    'Send to Xactimate from Export
    Set Nodx = ECTree.Nodes.Add(goUtil.gsAppEXEName, tvwChild, goUtil.gsAppEXEName & "|SendExportToXactimate", "Send Export File to Xactimate", PicTree.A22_SendCatToXactimate)
    'Add the Recycle Bin
    Set Nodx = ECTree.Nodes.Add(goUtil.gsAppEXEName, tvwChild, goUtil.gsAppEXEName & "|RecycleBin", "Recycle Bin", PicTree.A18_RecycleBinEmpty)
    Nodx.EnsureVisible
    
    'bgs 12.17.2003 Need to set the Main DB here to get lookup information...
    'Check version, Do NoCompact Repair
    'Check to See if the Main DB is Missing
    If Not goUtil.utFileExists(goUtil.gsInstallDir & "\ECMain.mdb") Then
RESTORE_DB:
        Err.Clear
        On Error GoTo EH
        sMess = "Database missing!" & vbCrLf
        sMess = sMess & "Could not find " & goUtil.gsInstallDir & "\ECMain.mdb" & vbCrLf & vbCrLf
        If goUtil.utFileExists(goUtil.gsInstallDir & "\ECMain_BackUp.db") Then
            sMess = sMess & App.EXEName & " Found " & goUtil.gsInstallDir & "\ECMain_BackUp.db" & vbCrLf
            sMess = sMess & goUtil.utGetAppVSInfo(vbNullString, goUtil.gsInstallDir & "\ECMain_BackUp.db") & vbCrLf & vbCrLf
            sMess = sMess & "Do you want to use this backup?"
            If MsgBox(sMess, vbExclamation + vbYesNo, "Restore from backup") = vbYes Then
                sMess = goUtil.utCopyFile(goUtil.gsInstallDir & "\ECMain_BackUp.db", goUtil.gsInstallDir & "\ECMain.mdb")
                'if there was an error smess will have error info in it
                If sMess <> vbNullString Then
                    MsgBox sMess, vbCritical + vbOKOnly, "Error"
                    Exit Function
                Else
                    mbShutDownEasyClaim = False
                    GoTo SET_DB
                End If
            End If
        End If
        'If the Backup was unavailable then try to use the Template
        If goUtil.utFileExists(goUtil.gsInstallDir & "\Templates\ECMain.mdb") Then
            sMess = "Database missing!" & vbCrLf
            sMess = sMess & App.EXEName & " can restore the database with an empty one!" & vbCrLf & vbCrLf
            sMess = sMess & "Click ""Yes"" to reset your database." & vbCrLf
            If MsgBox(sMess, vbExclamation + vbYesNo, "Reset Dastabase") = vbYes Then
                sMess = goUtil.utCopyFile(goUtil.gsInstallDir & "\Templates\ECMain.mdb", goUtil.gsInstallDir & "\ECMain.mdb")
                'if there was an error smess will have error info in it
                If sMess <> vbNullString Then
                    MsgBox sMess, vbCritical + vbOKOnly, "Error"
                Else
                    MsgBox App.EXEName & " Database reset success!" & vbCrLf & App.EXEName & " must now exit.", vbInformation + vbOKOnly, "Reset Dastabase"
                End If
                mbShutDownEasyClaim = True
                Exit Function
            End If
        End If
        'If the Template was unavailable then try to get the DB from the SP folder
        If goUtil.utFileExists(msSPPath & "\DataBase\SP", True) Then
            ' if the SP directory exists then look for the latest Version DB
            sDir = Dir(msSPPath & "\DataBase\SP\*.exe", vbNormal)
            Do Until sDir = vbNullString
                sDBSPName = sDir
                sDir = Dir
            Loop
            If sDBSPName <> vbNullString Then
                sMess = "Database missing!" & vbCrLf
                sMess = sMess & App.EXEName & " Found a DataBase Service Pack " & vbCrLf
                sMess = sMess & msSPPath & "DataBase\SP\" & sDBSPName & vbCrLf
                sMess = sMess & goUtil.utGetAppVSInfo(vbNullString, msSPPath & "\DataBase\SP\" & sDBSPName) & vbCrLf & vbCrLf
                sMess = sMess & "Do you want to install this database service pack?" & vbCrLf
                sMess = sMess & "This will also reset your database to an empty one!"
                If MsgBox(sMess, vbExclamation + vbYesNo, "Reinstall from Service Pack") = vbYes Then
                    Shell msSPPath & "\DataBase\SP\" & sDBSPName
                    Sleep 1000
                    sMess = goUtil.utCopyFile(goUtil.gsInstallDir & "\Templates\ECMain.mdb", goUtil.gsInstallDir & "\ECMain.mdb")
                    'if there was an error smess will have error info in it
                    If sMess <> vbNullString Then
                        MsgBox sMess, vbCritical + vbOKOnly, "Error"
                    Else
                        MsgBox App.EXEName & " Database Installed and Reset success!" & vbCrLf & App.EXEName & " must now exit.", vbInformation + vbOKOnly, "Restore from backup"
                    End If
                    mbShutDownEasyClaim = True
                    Exit Function
                End If
            End If
        End If
        
    End If
SET_DB:
    goUtil.SetMainDB App.EXEName, goUtil.gsInstallDir & "\ECMain.mdb", , True, False
    
    'Set the UsersID
    
    sSQL = "SELECT UsersID "
    sSQL = sSQL & "FROM Users "
    Set RS = goUtil.gMainDB.OpenRecordset(sSQL)
    If Not RS.EOF Then
        goUtil.gsCurUsersID = IIf(IsNull(RS!UsersID), 0, RS!UsersID)
    End If
    RS.Close
    
    'Need to Get a list of Current Companies for this user
    'and build the directory for them if none exists...
    'Folder structure will be created using Primary Key ID for
    'Companies and Client Companies, and their Cats...
    'That way the folder structure need not change if
    'the Label for that ID changes.
    
    sSQL = "SELECT * FROM Company "
    sSQL = sSQL & "WHERE IsClientOF is Null "
    sSQL = sSQL & "AND CompanyID IN ( "
                    sSQL = sSQL & "SELECT CompanyID "
                    sSQL = sSQL & "FROM CAT "
                    sSQL = sSQL & ") "
    Set CompanyRS = goUtil.gMainDB.OpenRecordset(sSQL)
        
    If CompanyRS.EOF Then
        GoTo SKIP_COMPANY
    End If
    CompanyRS.MoveFirst
    
    Do Until CompanyRS.EOF
               
        sCompany = CompanyRS!CompanyID
        sCompanyName = CompanyRS!Name
        sCompanyDBFolderName = CompanyRS!DBName
        
        'Need to Validate Cat Directory Structure in Registry
        If Not goUtil.ValidCatStructure(sCompany, sCompany) Then
            If Not goUtil.SaveCatStructure(sCompany, sCompany) Then
                GoTo SKIP_COMPANY
            End If
        End If
        
        'Add the Company folder to the tree
        sRelative = goUtil.gsAppEXEName
        sKey = "COMPANY|" & sCompany
        Set Nodx = ECTree.Nodes.Add(sRelative, tvwChild, sKey, sCompanyName, PicTree.A10_OpenFolder)
        Nodx.EnsureVisible
        sCompanyRelative = sKey
        
        'Get the list of Client Companies for this Company
        sSQL = "SELECT * FROM Company "
        sSQL = sSQL & "WHERE IsClientOF = " & sCompany & " "
        sSQL = sSQL & "AND CompanyID IN ( "
        sSQL = sSQL & "SELECT ClientCompanyID "
        sSQL = sSQL & "FROM ClientCompanyUsersCat "
        sSQL = sSQL & ") "
        
        Set ClientCompanyRS = goUtil.gMainDB.OpenRecordset(sSQL)
        
        If ClientCompanyRS.EOF Then
            GoTo SKIP_CLIENT_COMPANY
        End If
        ClientCompanyRS.MoveFirst
        
        Do Until ClientCompanyRS.EOF
            sCar = ClientCompanyRS!CompanyID
            sCarName = ClientCompanyRS!Name
            sCarDBName = ClientCompanyRS!DBName
            'Need to Validate Cat Directory Structure in Registry
            If Not goUtil.ValidCatStructure(sCompany & "\" & sCar, sCar) Then
                If Not goUtil.SaveCatStructure(sCompany & "\" & sCar, sCar) Then
                    GoTo SKIP_CLIENT_COMPANY
                End If
            End If
            
            'Add the Carrier folder to the tree
            sKey = "COMPANY|" & sCompany & "|CAR|" & sCar
            Set Nodx = ECTree.Nodes.Add(sCompanyRelative, tvwChild, sKey, sCarDBName, PicTree.A10_OpenFolder)
            Nodx.EnsureVisible
            sCarRelative = sKey
            
            'Get the list of Client Company Cats for this Client Company
            sSQL = "SELECT *, "
            sSQL = sSQL & "(    SELECT  Name "
            sSQL = sSQL & "     FROM CAT "
            sSQL = sSQL & "     WHERE CATID = CCC.CATID "
            sSQL = sSQL & ") As CATName, "
            sSQL = sSQL & "(    SELECT  Name "
            sSQL = sSQL & "     FROM Company "
            sSQL = sSQL & "     WHERE CompanyID = CCC.ClientCompanyID "
            sSQL = sSQL & ") As ClientCompanyName, "
            sSQL = sSQL & "(    SELECT  DBName "
            sSQL = sSQL & "     FROM Company "
            sSQL = sSQL & "     WHERE CompanyID = CCC.ClientCompanyID "
            sSQL = sSQL & ") As ClientCompanyDBName, "
            sSQL = sSQL & "(    SELECT  ScheduleName & '(' & Description & ')' "
            sSQL = sSQL & "     FROM FeeSchedule "
            sSQL = sSQL & "     WHERE FeeScheduleID = CCC.FeeScheduleID "
            sSQL = sSQL & ") As FeeScheduleName, "
            sSQL = sSQL & "(    SELECT  TypeOfLoss & '(' & Description & ')' "
            sSQL = sSQL & "     FROM TypeOfLoss "
            sSQL = sSQL & "     WHERE TypeOfLossID = CCC.TypeOfLossID "
            sSQL = sSQL & ") As TypeOfLossName, "
            sSQL = sSQL & "(    SELECT  Active "
            sSQL = sSQL & "     FROM ClientCompanyUsersCat "
            sSQL = sSQL & "     WHERE ClientCompanyID = CCC.ClientCompanyID "
            sSQL = sSQL & "     AND CATID = CCC.CATID "
            sSQL = sSQL & ") As UserActiveForCat "
            sSQL = sSQL & "FROM ClientCompanyCat CCC "
            sSQL = sSQL & "WHERE ClientCompanyID = " & sCar & " "
            sSQL = sSQL & "AND "
            sSQL = sSQL & "(    SELECT  Active "
            sSQL = sSQL & "     FROM ClientCompanyUsersCat "
            sSQL = sSQL & "     WHERE ClientCompanyID = CCC.ClientCompanyID "
            sSQL = sSQL & "     AND CATID = CCC.CATID "
            sSQL = sSQL & ") "
            'Only add the cat if there are assignments associated witht the Cat
            sSQL = sSQL & "AND "
            sSQL = sSQL & "( "
            sSQL = sSQL & "SELECT COUNT(Assignments.AssignmentsID) "
            sSQL = sSQL & "FROM Assignments "
            sSQL = sSQL & "WHERE ClientCompanyCatSpecID IN "
                                        sSQL = sSQL & "( "
                                        sSQL = sSQL & "SELECT   ClientCompanyCatSpecID "
                                        sSQL = sSQL & "FROM     ClientCompanyCatSpec  "
                                        sSQL = sSQL & "WHERE    ClientCompanyID = CCC.ClientCompanyID "
                                        sSQL = sSQL & "AND  CATID = CCC.CATID "
                                        sSQL = sSQL & ") "
            sSQL = sSQL & "AND AdjusterSpecID IN "
                                sSQL = sSQL & "( "
                                sSQL = sSQL & "SELECT   ClientCoAdjusterSpecID "
                                sSQL = sSQL & "FROM     ClientCoAdjusterSpec  "
                                sSQL = sSQL & "WHERE    ClientCompanyID = CCC.ClientCompanyID "
                                sSQL = sSQL & "AND USERSID = " & goUtil.gsCurUsersID & " "
                                sSQL = sSQL & ") "
            sSQL = sSQL & ") "
            
            Set ClientCompanyCatRS = goUtil.gMainDB.OpenRecordset(sSQL)
            
            If ClientCompanyCatRS.EOF Then
                GoTo SKIP_CLIENT_COMPANY_CAT
            End If
            ClientCompanyCatRS.MoveFirst
            
            Do Until ClientCompanyCatRS.EOF
                sCat = ClientCompanyCatRS!CATID
                sCatName = ClientCompanyCatRS!CatName
                If IsDate(ClientCompanyCatRS!InactiveDate) Then
                    bActiveCat = False
                Else
                    bActiveCat = True
                End If
                
                'Need to Validate Cat Directory Structure in Registry
                'but also need to see if It has been put in the Recycle Bin
                Set Nodx = ECTree.Nodes(goUtil.gsAppEXEName & "|RecycleBin")
                If Not goUtil.ValidCatStructure(sCompany & "\" & sCar & "\" & sCat, sCat, Nodx, ECTree, sCompany, sCompanyName, sCar, sCarDBName, sCat, sCatName, bAddedDeleteAllFromRecycleBin, sAddedCompanyRecycleBin, sAddedCarRecycleBin, bActiveCat) Then
                    If Nodx Is Nothing Then
                        If Not goUtil.SaveCatStructure(sCompany & "\" & sCar & "\" & sCat, sCat) Then
                            GoTo SKIP_CLIENT_COMPANY_CAT
                        End If
                    Else
                        'If Nodx has not been set to Nothing then
                        'this Cat is in the Recycle Bin
                        'Can skip adding it to this Car
                        Set Nodx = Nothing
                        GoTo SKIP_CLIENT_COMPANY_CAT
                    End If
                End If
                'Add the Carrier folder to the tree
                sKey = "COMPANY|" & sCompany & "|CAR|" & sCar & "|" & "CAT|" & sCat
                Set Nodx = ECTree.Nodes.Add(sCarRelative, tvwChild, sKey, sCatName, IIf(bActiveCat, PicTree.A14_ClosedHurc, PicTree.A42_closedHurc_Inactive))
                sCatRelative = sKey
                Nodx.EnsureVisible
        
                sKey = "COMPANY|" & sCompany & "|CAR|" & sCar & "|" & "CAT|" & sCat
                'Add Specific Items to the Cat
                Set Nodx = ECTree.Nodes.Add(sCatRelative, tvwChild, sKey & "|AddClaim", "Add Claim", PIC_TREE_IE_EXPLORE)
                Set Nodx = ECTree.Nodes.Add(sCatRelative, tvwChild, sKey & "|ClaimsListView", "Claims List View", PicTree.A01_ListView)
                Set Nodx = ECTree.Nodes.Add(sCatRelative, tvwChild, sKey & "|FeeSchedule", "Fee Schedule", PicTree.A04_FeeSched)
'                Set Nodx = ECTree.Nodes.Add(sCatRelative, tvwChild, sKey & "|CatPreferences", "Cat Preferences", PicTree.A05_CatPref)
'                Set Nodx = ECTree.Nodes.Add(sCatRelative, tvwChild, sKey & "|Communications", "Communications", PicTree.A06_Communications)
                Set Nodx = ECTree.Nodes.Add(sCatRelative, tvwChild, sKey & "|ViewLossReports", "View Loss Reports", PicTree.A41_ViewLossReports)
                Set Nodx = ECTree.Nodes.Add(sCatRelative, tvwChild, sKey & "|ViewAdjusterReports", "View Adjuster Reports", PicTree.A31_Item1)
'                Set Nodx = ECTree.Nodes.Add(sCatRelative, tvwChild, sKey & "|MakeCurrent", "Make Current", PicTree.A08_MakeCurrent)
                Set Nodx = ECTree.Nodes.Add(sCatRelative, tvwChild, sKey & "|CatMaintenance", "Cat Maintenance", PicTree.A09_ClosedFolder)
                'Add options to Cat Maintenance Folder
                Set Nodx = ECTree.Nodes.Add(sCatRelative & "|CatMaintenance", tvwChild, sKey & "|CatMaintenance" & "|SendCatToXactimate", "Send To Xactimate", PicTree.A22_SendCatToXactimate)
'                Set Nodx = ECTree.Nodes.Add(sCatRelative & "|CatMaintenance", tvwChild, sKey & "|CatMaintenance" & "|UpdateCat", "Update", PicTree.A23_UpdateCurrentCat)

'                'Send Cat To Disk
'                sSendCatToDisk = sTemp & "|CatMaintenance" & "|SendCatToDisk"
'                Set Nodx = ECTree.Nodes.Add(sTemp & "|CatMaintenance", tvwChild, sSendCatToDisk, "Send Cat To Disk", PicTree.A24_SendCatToDisk)
'                'Add Options to Send Cat To Disk
'                Set Nodx = ECTree.Nodes.Add(sSendCatToDisk, tvwChild, sSendCatToDisk & "|CreateCatBackup", "Create Backup", PicTree.A25_CreateCatBackup)
'                Set Nodx = ECTree.Nodes.Add(sSendCatToDisk, tvwChild, sSendCatToDisk & "|ExportCat", "Export", PicTree.A26_ExportCat)
                'Add Send to Recycle Bin
                Set Nodx = ECTree.Nodes.Add(sCatRelative & "|CatMaintenance", tvwChild, sKey & "|CatMaintenance" & "|SendCatToRecycleBin", "Send To Recycle Bin", PicTree.A18_RecycleBinEmpty)

                'Create Carrier Object and fill in Carrier Specific Tree Items
                Set oCar = CreateObject(goUtil.gsCarPrefix & sCarDBName & ".clsLists")
                oCar.PopulateECTree Nodx, ECTree, sTemp, sCar, sCat
        
SKIP_CLIENT_COMPANY_CAT:
                If Not ClientCompanyCatRS.EOF Then
                    ClientCompanyCatRS.MoveNext
                End If
            Loop

SKIP_CLIENT_COMPANY:
            If Not ClientCompanyRS.EOF Then
                ClientCompanyRS.MoveNext
            End If
        Loop


SKIP_COMPANY:
        If Not CompanyRS.EOF Then
            CompanyRS.MoveNext
        End If
    Loop
    
    'Select the Last Tree Key
    sTreeLastKey = GetSetting(App.EXEName, "GENERAL", "TREE_LAST_KEY", vbNullString)
    On Error Resume Next
    Set Nodx = ECTree.Nodes(sTreeLastKey)
    If Err.Number > 0 Then
        Err.Clear
        On Error GoTo EH
        Set Nodx = ECTree.Nodes(goUtil.gsAppEXEName)
        GoTo ENSURE_VISIBLE
    Else
ENSURE_VISIBLE:
        Nodx.EnsureVisible
        Nodx.Selected = True
        SaveSetting App.EXEName, "MSG", "COMMAND", "SET_FOCUS"
    End If
    
    LoadTree = True
    
    'CleanUp
    Set RS = Nothing
    Set CompanyRS = Nothing
    Set ClientCompanyRS = Nothing
    Set ClientCompanyCatRS = Nothing
    Set oCar = Nothing
    Set colCompany = Nothing
    Set colCar = Nothing
    Set colCat = Nothing
    Set Nodx = Nothing
    ECTree.Visible = True
    Exit Function
EH:
    If Err.Number = 3343 Or Err.Number = 3024 Then
        '3343 unrecognized DB Format
        '3024 Could not find file
        MsgBox "Error # " & Err.Number & vbCrLf & Err.Description, vbCritical + vbOKOnly, "Error"
        GoTo RESTORE_DB
    End If
    ECTree.Visible = True
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function LoadTree"
    'CleanUp
    Set RS = Nothing
    Set CompanyRS = Nothing
    Set ClientCompanyRS = Nothing
    Set ClientCompanyCatRS = Nothing
    Set oCar = Nothing
    Set colCompany = Nothing
    Set colCar = Nothing
    Set colCat = Nothing
    Set Nodx = Nothing
    ECTree.Visible = True
End Function

Private Function ExecuteNode() As Boolean
    On Error GoTo EH
    Dim Nodx As Node
    Dim sParent As String
    Dim sText As String
    Dim vKey As Variant
    Dim sKey As String
    Dim lCount As Long
    'EasyClaim Commands
    Dim sEasyClaimCommand As String
    Dim sEmptyRecycleBin As String
    Dim sRecBinCompany As String
    Dim sRecBinCar As String
    Dim sRecBinCat As String
    Dim sRecBinCatCommand As String
    'Company
    Dim sCompany As String
    Dim sCompanyName As String
    Dim sCar As String
    Dim sCarDBName As String
    Dim sCat As String
    Dim sCatName As String
    'Cat Commands
    Dim sCatCommand As String
    'Cat Maintenance Commands
    Dim sCatMaintCommand As String
    Dim sSendCatToDiskCommand As String
    Dim sGetCatFromDiskCommand As String
    Dim bCompRepair As Boolean
    Dim sMess As String
    
    'FTP Commands
    Dim bFTPCommandComplete As Boolean
    Dim lHwnd As Long
    Dim lSleep As Long
    
    
    If FlagSendToXactimate Then
        MessSendToXactimate App.EXEName
        GoTo CLEANUP
    End If
    
    If FlagLoadingLossReports Then
        GoTo CLEANUP
    End If
    
    
    If mNodX Is Nothing Then
        GoTo CLEANUP
    End If
    
    'Update the Screen Mouse pointer
    Screen.MousePointer = vbHourglass
    
    sKey = mNodX.Key
     
    'Remember the Last Tree Key
    SaveSetting App.EXEName, "GENERAL", "TREE_LAST_KEY", sKey
    
    'Update the Status Bar with current Tree Selection
    sText = mNodX.Text
    If mNodX.Parent Is Nothing Then
        ECStatBar.Panels(1).Text = sText
        ECStatBar.Panels(1).ToolTipText = ECStatBar.Panels(1).Text
        Screen.MousePointer = vbNormal
        GoTo CLEANUP
    Else
        sParent = mNodX.Parent
        ECStatBar.Panels(1).Text = sParent & " (" & sText & ")"
        ECStatBar.Panels(1).ToolTipText = ECStatBar.Panels(1).Text
        ECStatBar.Refresh
    End If
    
    'Check for Valid Tree Selection
   
    If InStr(1, sKey, "|", vbBinaryCompare) > 0 Then
        vKey = Split(sKey, "|")
    Else
        Screen.MousePointer = vbNormal
        GoTo CLEANUP
    End If
    
    For lCount = LBound(vKey, 1) To UBound(vKey, 1)
        'EasyClaim Commands
        sKey = vKey(lCount)
        If StrComp(sKey, goUtil.gsAppEXEName, vbTextCompare) = 0 Then
            sEasyClaimCommand = GetNextKeyValue(vKey, lCount)
        End If
        'Easy Claim Child Commands
        '1. Get Cat From Disk
        If StrComp(sEasyClaimCommand, "GetCatFromDisk", vbTextCompare) = 0 Then
            sGetCatFromDiskCommand = GetNextKeyValue(vKey, lCount)
            If sGetCatFromDiskCommand <> vbNullString Then
                GetCatFromDiskCommand sGetCatFromDiskCommand
            End If
        '2 RecycleBin Commands
        ElseIf StrComp(sEasyClaimCommand, "RecycleBin", vbTextCompare) = 0 Then
            sEmptyRecycleBin = GetNextKeyValue(vKey, lCount)
            'Empty all contents of Recycle Bin
            If StrComp(sEmptyRecycleBin, "Empty", vbTextCompare) = 0 Then
                If EmptyRecycleBin Then
'                    DoEvents
                    Sleep 1000
                    LoadTree
                End If
                Exit For
            Else
                'Move back to check for Recycle Bin Company
                lCount = lCount - 1
            End If
            'Check for Recycle Bin Cat Command
            sRecBinCompany = GetNextKeyValue(vKey, lCount)
            sRecBinCar = GetNextKeyValue(vKey, lCount)
            sRecBinCat = GetNextKeyValue(vKey, lCount)
            sRecBinCatCommand = GetNextKeyValue(vKey, lCount)
            If sRecBinCompany <> vbNullString And sRecBinCar <> vbNullString And sRecBinCat <> vbNullString Then
                If sRecBinCatCommand <> vbNullString Then
                    RecycleBinCatCommand sRecBinCompany, sRecBinCar, sRecBinCat, sRecBinCatCommand
                    Exit For
                End If
            End If
        Else
            If sEasyClaimCommand <> vbNullString Then
                EasyClaimCommand sEasyClaimCommand
            End If
        End If
        
        'Carriers
        sKey = vKey(lCount)
        If StrComp(sKey, "COMPANY", vbTextCompare) = 0 Then
            sCompany = GetNextKeyValue(vKey, lCount)
            Set Nodx = ECTree.Nodes("COMPANY|" & sCompany)
            sCompanyName = Nodx.Text
            sKey = GetNextKeyValue(vKey, lCount)
            If StrComp(sKey, "CAR", vbTextCompare) = 0 Then
                sCar = GetNextKeyValue(vKey, lCount)
                Set Nodx = ECTree.Nodes("COMPANY|" & sCompany & "|CAR|" & sCar)
                sCarDBName = Nodx.Text
                sKey = GetNextKeyValue(vKey, lCount)
            End If
        End If
        'Carrier CAT Commands
        If StrComp(sKey, "CAT", vbTextCompare) = 0 Then
            sCat = GetNextKeyValue(vKey, lCount)
            Set Nodx = ECTree.Nodes("COMPANY|" & sCompany & "|CAR|" & sCar & "|CAT|" & sCat)
            sCatName = Nodx.Text
            sCatCommand = GetNextKeyValue(vKey, lCount)
            'Now Set the Current Cat Directory If applicable
            If sCar <> vbNullString And sCat <> vbNullString Then
                'Check to see if we have already selected the same Car and Cat
                If StrComp(sCatCommand, "MakeCurrent", vbTextCompare) = 0 Then
                    bCompRepair = True
                    If Not CloseFTPDBConnection Then
                        GoTo CLEANUP
                    End If
                End If
                If goUtil.gsCurCatDir <> goUtil.gsInstallDir & "\Cats\" & sCompany & "\" & sCar & "\" & sCat Or bCompRepair Then
                    If Not goUtil.goCurCarList Is Nothing Then
                        goUtil.goCurCarList.CLEANUP
                        Set goUtil.goCurCarList = Nothing
                    End If
                    goUtil.gsCurCatDir = goUtil.gsInstallDir & "\Cats\" & sCompany & "\" & sCar & "\" & sCat
                    goUtil.gsCurCompany = sCompany
                    goUtil.gsCurCarDBName = sCarDBName
                    goUtil.gsCurCar = sCar
                    goUtil.gsCurCat = sCat
                    'Save these in reg setting so other Apps can get updated
                    SaveSetting App.EXEName, "DIR", "CURRENT_CAT_DIR", goUtil.gsCurCatDir
                    SaveSetting App.EXEName, "GENERAL", "CURRENT_COMPANY", sCompany
                    SaveSetting App.EXEName, "GENERAL", "CURRENT_COMPANY_NAME", sCompanyName
                    SaveSetting App.EXEName, "GENERAL", "CURRENT_CAR", sCar
                    SaveSetting App.EXEName, "GENERAL", "CURRENT_CAR_NAME", sCarDBName
                    SaveSetting App.EXEName, "GENERAL", "CURRENT_CAT", sCat
                    SaveSetting App.EXEName, "GENERAL", "CURRENT_CAT_NAME", sCatName
                End If
                
                'Set Current Carrier Object
                If goUtil.goCurCarList Is Nothing Then
                    GoTo SET_CARRIER_OBJECTS
                Else
                    If StrComp(goUtil.goCurCarList.ClassName, goUtil.gsCarPrefix & sCarDBName & ".clsLists", vbTextCompare) <> 0 Then
SET_CARRIER_OBJECTS:
                        SetCarrierGlobalObjects sCompany, sCar, bCompRepair
                    End If
                End If
                lblCurCatTop.Caption = mNodX.Parent.Parent.Text & "\" & mNodX.Parent.Text & "\" & mNodX.Text
                'Also Update the Tray Icon Caption
                m_NID.szTip = Me.Caption & " " & lblCurCatTop.Caption & vbNullChar
                Call ShellNotifyIcon(NIM_MODIFY, m_NID)
                imgCurCat.Picture = imgExplor.ListImages.Item(mNodX.Image).Picture
                imgCurCat.ToolTipText = mNodX.Text
                lblCurCat.Caption = imgCurCat.ToolTipText
                lblCurCat.ToolTipText = imgCurCat.ToolTipText
'                cmdConnect.Enabled = True
            End If
            
            
            'Cat Maintenance
            If StrComp(sCatCommand, "CatMaintenance", vbTextCompare) = 0 Then
                sCatMaintCommand = GetNextKeyValue(vKey, lCount)
                'Send Cat To Disk Commands
                If StrComp(sCatMaintCommand, "SendCatToDisk", vbTextCompare) = 0 Then
                    sSendCatToDiskCommand = GetNextKeyValue(vKey, lCount)
                    If sSendCatToDiskCommand <> vbNullString Then
                        CatMaintSendCatToDiskCommand sCar, sCat, sSendCatToDiskCommand
                    End If
                'Cat Maintenance Commands
                Else
                    If sCatMaintCommand <> vbNullString Then
                        CatMaintenanceCommand sCompany, sCar, sCat, sCatMaintCommand
                    End If
                End If
            'Cat Commands
            Else
                If sCatCommand <> vbNullString Then
                    '4.25.2005 BGS
                    'Check for inactive Cat
                    'Do not allow Cat Commands (but do allow Cat Maintenance Commands)
                    If Nodx.Image = PicTree.A42_closedHurc_Inactive Or Nodx.Image = PicTree.A43_OpenHurc_Inactive Then
                        If StrComp(imgCurCat.ToolTipText, "Add Claim", vbTextCompare) = 0 Then
                            lblCurCat.Caption = imgCurCat.ToolTipText
                            lblCurCat.Caption = lblCurCat.Caption & " (DISABLED)"
                        ElseIf StrComp(imgCurCat.ToolTipText, "Claims List View", vbTextCompare) = 0 Then
                            lblCurCat.Caption = imgCurCat.ToolTipText
                            lblCurCat.Caption = lblCurCat.Caption & " (DISABLED)"
                        Else
                            CatCommand sCompany, sCar, sCat, sCatCommand
                        End If
                    Else
                        CatCommand sCompany, sCar, sCat, sCatCommand
                    End If
                End If
            End If
        End If
    Next
CLEANUP:
   
    Set mNodX = Nothing
    ExecuteNode = True
    Screen.MousePointer = vbNormal
    Exit Function
EH:
    
    Screen.MousePointer = vbNormal
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Function ExecuteNode"
End Function

Public Function CloseFTPDBConnection() As Boolean
    On Error GoTo EH
    Dim lHwnd As Long
    Dim lSleep As Long
    Dim bFTPCommandComplete As Boolean
    
    CloseFTPDBConnection = True
    
    'IF FTP is Running..
    lHwnd = goUtil.utFindWindowPartial("Communications Status", FwpStartsWith, True, False)
    If lHwnd > 0 Then
        'Also need to tell FTP to Disconnect from Current Car
        'First Check to see if it is Currently Connected
        If Not mbFTPConnected Then
            SaveSetting "ECFTP", "MSG", "COMMAND", "CLOSE_CURRENT"
            For lSleep = 1 To 10
                Sleep 500
                bFTPCommandComplete = GetSetting("ECFTP", "MSG", "COMMAND_COMPLETE", False)
                If bFTPCommandComplete Then
                    SaveSetting "ECFTP", "MSG", "COMMAND_COMPLETE", False
                    Exit For
                End If
            Next
            If Not bFTPCommandComplete Then
                MsgBox "Communications Status not responding!", vbCritical + vbOKOnly, "Communications Status Close DB"
                CloseFTPDBConnection = False
            End If
        Else
            MsgBox "Please wait... Communication Status is Connected!", vbExclamation + vbOKOnly, "Communications Status Close DB"
            CloseFTPDBConnection = False
        End If
    End If
    Exit Function
EH:
    CloseFTPDBConnection = False
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function CloseFTPDBConnection"
End Function

Public Function OpenFTPDBConnection() As Boolean
    On Error GoTo EH
    Dim lHwnd As Long
    Dim lSleep As Long
    Dim bFTPCommandComplete As Boolean
    
    OpenFTPDBConnection = True
    
    'IF FTP is Running..
    lHwnd = goUtil.utFindWindowPartial("Communications Status", FwpStartsWith, True, False)
    If lHwnd > 0 Then
        'Also need to tell FTP to Disconnect from Current Car
        'First Check to see if it is Currently Connected
        If Not mbFTPConnected Then
            SaveSetting "ECFTP", "MSG", "COMMAND", "OPEN_CURRENT"
            For lSleep = 1 To 10
                Sleep 500
                bFTPCommandComplete = GetSetting("ECFTP", "MSG", "COMMAND_COMPLETE", False)
                If bFTPCommandComplete Then
                    SaveSetting "ECFTP", "MSG", "COMMAND_COMPLETE", False
                    Exit For
                End If
            Next
            If Not bFTPCommandComplete Then
                MsgBox "Communications Status not responding!", vbCritical + vbOKOnly, "Communications Status Open DB"
                OpenFTPDBConnection = False
            End If
        Else
            MsgBox "Please wait... Communication Status is Connected!", vbExclamation + vbOKOnly, "Communications Status Open DB"
            OpenFTPDBConnection = False
        End If
    End If
    Exit Function
EH:
    OpenFTPDBConnection = False
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function OpenFTPDBConnection"
End Function

Public Function SetCarrierGlobalObjects(psCompany As String, psCar As String, pbCompRepair As Boolean, Optional pbUpdateDB As Boolean = False) As Boolean
    On Error GoTo EH
    'Global Collection
    Dim colGlobalObjects As Collection
    Dim sSQL As String
    Dim RS As DAO.Recordset
    Dim sCompanyFolderName As String
    Dim sCarFolderName As String
    Dim sCarDBName As String
    Dim bCloseMainDB As Boolean
    Dim sCarPrefix As String
    
    'Set the main DB if its not already open
    If goUtil.gMainDB Is Nothing Then
       goUtil.SetMainDB App.EXEName, goUtil.gsInstallDir & "\ECMain.mdb", , True, False
    End If
    
    'Need to Get Company info to Create the Carrier Object
    sSQL = "SELECT *, "
    sSQL = sSQL & "( "
    sSQL = sSQL & "     SELECT DBName "
    sSQL = sSQL & "     FROM Company "
    sSQL = sSQL & "     WHERE CompanyID = C.IsClientOF "
    sSQL = sSQL & ") As CompanyFolderName "
    sSQL = sSQL & "FROM Company C "
    sSQL = sSQL & "WHERE CompanyID = " & psCar & " "
    sSQL = sSQL & "AND IsClientOF = " & psCompany & " "
    
    Set RS = goUtil.gMainDB.OpenRecordset(sSQL)
    
    If Not RS.EOF Then
        RS.MoveFirst
    Else
        Err.Raise -999, , "No Company data found!"
    End If
    
    sCompanyFolderName = RS!CompanyFolderName
    sCarDBName = RS!DBName
    sCarFolderName = Replace(sCarDBName, ".mdb", vbNullString, , , vbTextCompare)
    sCarPrefix = RS!CarrierPrefix
    goUtil.gsCarPrefix = sCarPrefix
    RS.Close
    
    Set goUtil.goCurCarList = CreateObject(goUtil.gsCarPrefix & sCarFolderName & ".clsLists")
    If Not goUtil.goECKeyBoardList Is Nothing Then
        goUtil.goECKeyBoardList.CLEANUP
        Set goUtil.goECKeyBoardList = Nothing
    End If
    Set goUtil.goECKeyBoardList = New V2ECKeyBoard.clsLists
    If Not goUtil.gARV Is Nothing Then
        goUtil.gARV.CLEANUP
        Set goUtil.gARV = Nothing
    End If
    Set goUtil.gARV = New V2ARViewer.clsARViewer
    If Not goUtil.goProgForm Is Nothing Then
        goUtil.goProgForm.CLEANUP
        Set goUtil.goProgForm = Nothing
    End If
    Set goUtil.goProgForm = New V2ECKeyBoard.clsProgForm
    
    Set colGlobalObjects = New Collection
    colGlobalObjects.Add goUtil, "goUtil"
    
    
    goUtil.goCurCarList.SetGlobalObjects colGlobalObjects
    goUtil.gARV.SetGlobalObjects colGlobalObjects
    goUtil.SetGlobalObjects colGlobalObjects
    
    Set colGlobalObjects = Nothing
    
   
    goUtil.SetUtilObject goUtil
        
    SetCarrierGlobalObjects = True
    'Clean up
    Set RS = Nothing
    Exit Function
EH:
    SetCarrierGlobalObjects = False
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function SetCarrierGlobalObjects"
End Function

Private Function GetNextKeyValue(pvKey As Variant, plCount As Long) As String
    On Error GoTo EH
    
    If plCount < UBound(pvKey, 1) Then
        plCount = plCount + 1
        GetNextKeyValue = pvKey(plCount)
    End If
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Function GetKeyValue"
End Function

Public Function EasyClaimCommand(psCommand As String, Optional iShowMode As VBRUN.FormShowConstants = vbModeless) As Boolean
    On Error GoTo EH
    
    'Used For Assignments
    Dim sHTML As String
    Dim oLR As V2ECKeyBoard.clsLossReports
    Dim sWebSite As String
    Dim sUserName As String
    Dim sPassword As String
    Dim bUseSSL As Boolean
    Dim frmReg As frmRegForm
    Dim bShowFTPUpdate As Boolean
    Dim oForm As Form
    
    'Execute Command
    If StrComp(psCommand, "SetupNewCat", vbTextCompare) = 0 Then
        frmSetUpNewCat.Show iShowMode
        frmSetUpNewCat.WindowState = vbNormal
    ElseIf StrComp(psCommand, "Assignments", vbTextCompare) = 0 Then
        Screen.MousePointer = vbNormal
        bUseSSL = CBool(GetSetting("ECS", "WEB_SECURITY", "USE_SSL", True))
        If bUseSSL Then
            sWebSite = "https://"
        Else
            sWebSite = "http://"
        End If
        sWebSite = sWebSite & GetSetting("ECS", "WEB_SECURITY", "WEB_HOST", "WWW.EBERLS.NET")
        sUserName = goUtil.utGetECSCryptSetting("ECS", "WEB_SECURITY", "USER_NAME", vbNullString)
        sPassword = goUtil.utGetECSCryptSetting("ECS", "WEB_SECURITY", "PASSWORD", vbNullString)
        sHTML = goUtil.utGetFileData(goUtil.gsInstallDir & "\Templates\AssignmentsLogon.html")
        
        'Need to Replace Marked Text with above vars
        If sHTML <> vbNullString Then
            sHTML = Replace(sHTML, "|ENTER_WEB_SITE|", sWebSite, , , vbTextCompare)
            sHTML = Replace(sHTML, "|ENTER_USER_NAME|", sUserName, , , vbTextCompare)
            sHTML = Replace(sHTML, "|ENTER_PASSWORD|", sPassword, , , vbTextCompare)
        Else
            MsgBox "Could Not Find " & goUtil.gsInstallDir & "\Templates\AssignmentsLogon.html", vbCritical + vbOKOnly, "Error"
            Exit Function
        End If
        
        Set oLR = New V2ECKeyBoard.clsLossReports
        oLR.SetUtilObject goUtil
        
        oLR.ShowHelpViewer sHTML, "Assignments Logon", True
        oLR.CLEANUP
        Set oLR = Nothing
    ElseIf StrComp(psCommand, "PrinterSetup", vbTextCompare) = 0 Then
        ShowPrinter Me.hWnd, False
        If EndECTray Then
            Exit Function
        End If
        mnuUseWinDefaultPrinter.Checked = True
        FlagCheckPrinter = True
    ElseIf StrComp(psCommand, "GlobalPref", vbTextCompare) = 0 Then
        frmPreferences.Show iShowMode
        If iShowMode <> vbModal Then
            frmPreferences.WindowState = vbNormal
        End If
    ElseIf StrComp(psCommand, "Support", vbTextCompare) = 0 Then
        frmSupport.Show iShowMode
        frmSupport.WindowState = vbNormal
    ElseIf StrComp(psCommand, "About", vbTextCompare) = 0 Then
        mbLoadingRegForm = True
        Set frmReg = New frmRegForm
        Load frmReg
        'Let the About Splash go for 6 seconds instead of 2
        frmReg.TimerUnloadSplash.Enabled = False
        frmReg.TimerUnloadSplash.Interval = 6000
        frmReg.TimerUnloadSplash.Enabled = True
        frmReg.Caption = "About " & App.EXEName
        frmReg.lblMess.Caption = "Copyright 2001 - Eberl's Claim Service, Inc.  All rights reserved." & vbCrLf & vbCrLf
        frmReg.lblMess.Caption = frmReg.lblMess.Caption & "This product contains SpellChecker from Polar Software - Copyright 2001 Polar Software. All Rights Reserved."
        DoEvents
        Sleep 100
        frmReg.Show vbModal
        Unload frmReg
        Set frmReg = Nothing
        mbLoadingRegForm = False
    ElseIf StrComp(psCommand, "SHOW_FTP_IS_UPDATING_DATA", vbTextCompare) = 0 Then
        mbLoadingRegForm = True
        Set frmReg = New frmRegForm
        Load frmReg
        frmReg.TimerUnloadSplash.Enabled = False
        frmReg.TimerUnloadSplash.Interval = 15000
        frmReg.TimerUnloadSplash.Enabled = True
        frmReg.cmdOK.Visible = False
        frmReg.cmdViewLic.Visible = False
        frmReg.Caption = "FTP IS UPDATING DATA!"
        frmReg.lblMess.Caption = "FTP IS UPDATING DATA!  Please Wait for this process." & vbCrLf & vbCrLf
        frmReg.lblMess.Caption = frmReg.lblMess.Caption & "This screen will close when finished."
        
        'Check to see if Navigator, Claims List or Claim form is open and visible
        'if any of them are, then set the flag to show the FTP Update modal
        If Not goUtil.goCurCarList Is Nothing Then
            If goUtil.utFindSetForm(goUtil.goCurCarList.goForms, "frmClaimsList", oForm) Then
                If oForm.Visible Then
                    bShowFTPUpdate = True
                End If
            End If
            If goUtil.utFindSetForm(goUtil.goCurCarList.goForms, "frmClaim", oForm) Then
                If oForm.Visible Then
                    bShowFTPUpdate = True
                End If
            End If
        End If
        If Me.Visible Then
            bShowFTPUpdate = True
        End If
        
        If bShowFTPUpdate Then
            frmReg.Show vbModal
        Else
            'If non of the above mentioned forms are visible...
            'need to loop until the unloadsplash timer is disabled
            Do Until Not frmReg.TimerUnloadSplash.Enabled
                DoEvents
                Sleep 100
            Loop
        End If
        
        Unload frmReg
        Set frmReg = Nothing
        Set oForm = Nothing
        mbLoadingRegForm = False
    ElseIf StrComp(psCommand, "SendExportToXactimate", vbTextCompare) = 0 Then
        ' Need to set Mouse back since going inside another object
            Screen.MousePointer = vbNormal
            FlagSendToXactimate = True
            SendExportToXactimate
            FlagSendToXactimate = False
            If Not goUtil Is Nothing Then
                If Not goUtil.goXact Is Nothing Then
                    goUtil.goXact.CLEANUP
                    Set goUtil.goXact = Nothing
                End If
            Else
                Exit Function
            End If
    End If
        
    EasyClaimCommand = True
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "EasyClaimCommand"
End Function

Public Sub ShowCommStatus(psCar As String, psCat As String)
    On Error GoTo EH
    
    frmSetUpNewCat.Show
    frmSetUpNewCat.WindowState = vbNormal
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub ShowCommStatus"
End Sub

Private Function EmptyRecycleBin() As Boolean
    On Error GoTo EH
    Dim sKey As String
    Dim vKey As Variant
    Dim Nodx As Node
    Dim sRecBinCompany As String
    Dim sRecBinCar As String
    Dim sRecBinCat As String
    Dim sRecBinCatCommand As String
    Dim sMess As String
    Dim lCount As Long
    Dim sPassword As String
    
    sMess = "Are you sure?" & vbCrLf & vbCrLf
    sMess = sMess & "ALL Recycle Bin Items " & vbCrLf
    sMess = sMess & "And their associated " & vbCrLf & vbCrLf
    sMess = sMess & "Assignments" & vbCrLf
    sMess = sMess & "Attachments" & vbCrLf
    sMess = sMess & "Photos" & vbCrLf & vbCrLf
    sMess = sMess & "Will be removed!"
    
    If MsgBox(sMess, vbQuestion + vbYesNo, "Empty Recycle Bin") = vbNo Then
        Exit Function
    Else
        'ask for Password to run this option just to be sure
        sPassword = InputBox("Enter Your Password", "Enter Password")
        If sPassword <> goUtil.utGetECSCryptSetting("ECS", "WEB_SECURITY", "PASSWORD") Then
            If sPassword <> vbNullString Then
                MsgBox "Invalid Password!", vbExclamation + vbOKOnly, "Invalid Password"
            End If
            Exit Function
        End If
    End If
    
    For Each Nodx In ECTree.Nodes
        sKey = Nodx.Key
        If InStrRev(sKey, "|DeleteCat", , vbTextCompare) > 0 Then
            vKey = Split(sKey, "|")
            For lCount = LBound(vKey, 1) To UBound(vKey, 1)
                sKey = vKey(lCount)
                If StrComp(sKey, "RecycleBin", vbTextCompare) = 0 Then
                    'Get Recycle Bin Delete Cat Command
                    sRecBinCompany = GetNextKeyValue(vKey, lCount)
                    sRecBinCar = GetNextKeyValue(vKey, lCount)
                    sRecBinCat = GetNextKeyValue(vKey, lCount)
                    sRecBinCatCommand = GetNextKeyValue(vKey, lCount)
                    If sRecBinCompany <> vbNullString And sRecBinCar <> vbNullString And sRecBinCat <> vbNullString And sRecBinCatCommand <> vbNullString Then
                        If Not RecycleBinCatCommand(sRecBinCompany, sRecBinCar, sRecBinCat, sRecBinCatCommand, False, False) Then
                            LoadTree
                            Exit Function
                        End If
                    End If
                End If
            Next
        End If
    Next
    LoadTree
    EmptyRecycleBin = True
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Function EmptyRecycleBin"
End Function

Private Function RecycleBinCatCommand(psRecBinCompany, _
                                    psRecBinCar As String, _
                                    psRecBinCat As String, _
                                    psCommand As String, _
                                    Optional pbShowMessage As Boolean = True, _
                                    Optional pbLoadTree As Boolean = True) As Boolean
    On Error GoTo EH
    Dim sMess As String
    Dim sPassword As String
    Dim sSQL As String
    Dim sDelSQL As String
    Dim lCount As Long
    Dim lCountMax As Long
    Dim IDAssignments As String
    Dim IDTable As String
    Dim lTableValue As Long
    Dim lTableValueMax As Long
    Dim lFileValue As Long
    Dim lFileValueMax As Long
    Dim sTable As Variant
    Dim bRemoveFiles As Boolean
    'Photos
    Dim sPhotoName As String
    Dim sPhotoHighResName As String
    Dim sPhotoThumbName As String
    'Attachments
    Dim sAttach As String
    'WSDiagram Photos
    Dim sDiagramPhotoName
    'Objects
    Dim colTables As Collection
    Dim oConn As New ADODB.Connection
    Dim RS As ADODB.Recordset
    Dim RSTable As ADODB.Recordset
    'FTP Check
    Dim bFTPExists As Boolean
    'Adjuster Reports
    Dim sAdjReportFilterName As String
    
    'Execute Command
    If StrComp(psCommand, "RestoreCat", vbTextCompare) = 0 Then
        SaveSetting App.EXEName, "CAT_STRUCTURE\" & psRecBinCompany & "\" & psRecBinCar & "\" & psRecBinCat, psRecBinCat, goUtil.Encode(psRecBinCat)
        LoadTree
    ElseIf StrComp(psCommand, "DeleteCat", vbTextCompare) = 0 Then
        If Not pbShowMessage Then
            GoTo DELETE_CAT
        End If
        sMess = "Are you sure you want to remove ALL " & vbCrLf & vbCrLf
        sMess = sMess & "Assignments" & vbCrLf
        sMess = sMess & "Attachments" & vbCrLf
        sMess = sMess & "Photos" & vbCrLf & vbCrLf
        sMess = sMess & "associated with this cat?"
        If MsgBox(sMess, vbQuestion + vbYesNo, "Delete CAT") = vbYes Then
             'ask for Password to run this option just to be sure
            sPassword = InputBox("Enter Your Password", "Enter Password")
            If sPassword <> goUtil.utGetECSCryptSetting("ECS", "WEB_SECURITY", "PASSWORD") Then
                If sPassword <> vbNullString Then
                    MsgBox "Invalid Password!", vbExclamation + vbOKOnly, "Invalid Password"
                End If
                Exit Function
            End If
DELETE_CAT:
            'If the FTP Application is currently connected and download or Uploading
            'Data then can't run this at all !
            If goUtil.utFTPConnected Then
                sMess = "FTP connection is currently active." & vbCrLf & vbCrLf
                sMess = sMess & "You must wait until the current connection is finished " & vbCrLf
                sMess = sMess & "before running this utility."
                MsgBox sMess, vbExclamation + vbOKOnly, "FTP Connection Detected"
                Exit Function
            End If
            
            'If the FTP Application is running then need to close it before
            'Running the Compact Repair.
            
            bFTPExists = goUtil.utFTPExists
            If bFTPExists Then
                sMess = "Closing FTP Application!"
                MsgBox sMess, vbInformation + vbOKOnly, "FTP SHUTDOWN"
                goUtil.utShutDownFTP
                Sleep 2000 'Wait couple seconds
            End If
            
            'Remove any Adjuster Reports saved under Repos...
            sAdjReportFilterName = goUtil.gsCurCarDBName & "_" & GetCurCatName & "_*.zip"
            
            goUtil.utDeleteFile msAttachReposPath & "\" & sAdjReportFilterName

            sSQL = "SELECT A.ID "
            sSQL = sSQL & "FROM Assignments A "
            sSQL = sSQL & "INNER JOIN CLIENTCOMPANYCATSPEC CCCS ON A.ClientCompanyCatSpecID = CCCS.ClientCompanyCatSpecID "
            'If they are in a specific claim then only get this one claim to be sent to
            'xactimate. Otherwise we are sending all projects to xactimate.

            sSQL = sSQL & "WHERE A.ClientCompanyCatSpecID IN "
            sSQL = sSQL & "( "
            sSQL = sSQL & "SELECT   ClientCompanyCatSpecID "
            sSQL = sSQL & "FROM     ClientCompanyCatSpec "
            sSQL = sSQL & "WHERE    ClientCompanyID = " & psRecBinCar & " "
            sSQL = sSQL & "AND      CATID = " & psRecBinCat & " "
            sSQL = sSQL & ") "
            sSQL = sSQL & "AND A.AdjusterSpecID IN "
            sSQL = sSQL & "( "
            sSQL = sSQL & "SELECT   ClientCoAdjusterSpecID "
            sSQL = sSQL & "FROM     ClientCoAdjusterSpec "
            sSQL = sSQL & "Where    ClientCompanyID = " & psRecBinCar & " "
            sSQL = sSQL & "AND      UsersID = " & goUtil.gsCurUsersID & " "
            sSQL = sSQL & ") "
            
            
            Set oConn = New ADODB.Connection
            goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
            Set RS = New ADODB.Recordset
            
            
            'Use Disconnected Record Set on asUseClient Cusor ONLY !
            RS.CursorLocation = adUseClient
            RS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
            Set RS.ActiveConnection = Nothing
            If RS.EOF Then
                Exit Function
            End If
            
            Set colTables = New Collection
            If goUtil.goProgForm Is Nothing Then
                Set goUtil.goProgForm = New V2ECKeyBoard.clsProgForm
            End If
            'Set the Progress Form
             With goUtil.goProgForm
                .LoadForm
                .Caption = "Remove Data"
                .cmdCancelEnable = True
                .ShowForm True
                .SetFocus
            End With
            'Remove stuff from the following Tables
            'MiscReportParam
            '2/8/2004 MiscReportParam , and MiscReportParam01 to MiscReportParam30
            colTables.Add "MiscReportParam", "MiscReportParam"
            colTables.Add "MiscReportParam01", "MiscReportParam01"
            colTables.Add "MiscReportParam02", "MiscReportParam02"
            colTables.Add "MiscReportParam03", "MiscReportParam03"
            colTables.Add "MiscReportParam04", "MiscReportParam04"
            colTables.Add "MiscReportParam05", "MiscReportParam05"
            colTables.Add "MiscReportParam06", "MiscReportParam06"
            colTables.Add "MiscReportParam07", "MiscReportParam07"
            colTables.Add "MiscReportParam08", "MiscReportParam08"
            colTables.Add "MiscReportParam09", "MiscReportParam09"
            colTables.Add "MiscReportParam10", "MiscReportParam10"
            colTables.Add "MiscReportParam11", "MiscReportParam11"
            colTables.Add "MiscReportParam12", "MiscReportParam12"
            colTables.Add "MiscReportParam13", "MiscReportParam13"
            colTables.Add "MiscReportParam14", "MiscReportParam14"
            colTables.Add "MiscReportParam15", "MiscReportParam15"
            colTables.Add "MiscReportParam16", "MiscReportParam16"
            colTables.Add "MiscReportParam17", "MiscReportParam17"
            colTables.Add "MiscReportParam18", "MiscReportParam18"
            colTables.Add "MiscReportParam19", "MiscReportParam19"
            colTables.Add "MiscReportParam20", "MiscReportParam20"
            colTables.Add "MiscReportParam21", "MiscReportParam21"
            colTables.Add "MiscReportParam22", "MiscReportParam22"
            colTables.Add "MiscReportParam23", "MiscReportParam23"
            colTables.Add "MiscReportParam24", "MiscReportParam24"
            colTables.Add "MiscReportParam25", "MiscReportParam25"
            colTables.Add "MiscReportParam26", "MiscReportParam26"
            colTables.Add "MiscReportParam27", "MiscReportParam27"
            colTables.Add "MiscReportParam28", "MiscReportParam28"
            colTables.Add "MiscReportParam29", "MiscReportParam29"
            colTables.Add "MiscReportParam30", "MiscReportParam30"
            'PackageItem
            colTables.Add "PackageItem", "PackageItem"
            'Package
            colTables.Add "Package", "Package"
            'RTPhotoReport
            colTables.Add "RTPhotoReport", "RTPhotoReport"
            'RTPhotoLog (Remove Main Photo, Highres Photo, and Thumnail)
            colTables.Add "RTPhotoLog", "RTPhotoLog"
            'RTIndemnity
            colTables.Add "RTIndemnity", "RTIndemnity"
            'RTIB
            colTables.Add "RTIB", "RTIB"
            'RTIBFee
            colTables.Add "RTIBFee", "RTIBFee"
            'RTChecks
            colTables.Add "RTChecks", "RTChecks"
            'RTAttachments (Remove File Attahcments)
            colTables.Add "RTAttachments", "RTAttachments"
            'RTWSDiagram  (Remove Diagram Photos)
            colTables.Add "RTWSDiagram", "RTWSDiagram"
            'RTActivityLogInfo
            colTables.Add "RTActivityLogInfo", "RTActivityLogInfo"
            'RTActivityLog
            colTables.Add "RTActivityLog", "RTActivityLog"
            'PolicyLimits
            colTables.Add "PolicyLimits", "PolicyLimits"
            'IB
            colTables.Add "IB", "IB"
            'IB
            colTables.Add "IBFee", "IBFee"
            'BillingCount
            colTables.Add "BillingCount", "BillingCount"
            'Assignments (Remove Loss Report Attachments)
            colTables.Add "Assignments", "Assignments"
            
            RS.MoveFirst
            'This flag will let Application know it is Currently
            'in the middle of deleting stuff
            mbDeleting = True
            goUtil.goProgForm.PBarTable.Max = colTables.Count
            lCountMax = RS.RecordCount
            goUtil.goProgForm.PBarRecord.Max = lCountMax
            Do Until RS.EOF
                lCount = lCount + 1
                IDAssignments = RS!ID
                bRemoveFiles = False
                goUtil.goProgForm.lblFieldText = "Removing Assignments (" & lCount & ") Of (" & lCountMax & ") "
                goUtil.goProgForm.PBarRecord.Value = lCount
                goUtil.goProgForm.RefreshMe
                DoEvents
                Sleep 10
                'Check for user cancel
                If Not goUtil.goProgForm Is Nothing Then
                    If goUtil.goProgForm.CancelMe Then
                        GoTo CLEAN_UP
                    End If
                Else
                    GoTo CLEAN_UP
                End If
                
                For lTableValue = 1 To colTables.Count
                    sTable = colTables(lTableValue)
                    goUtil.goProgForm.lblTableText = sTable
                    goUtil.goProgForm.PBarTable.Value = lTableValue
                    goUtil.goProgForm.RefreshMe
                    DoEvents
                    Sleep 10
                    'Check for user cancel
                    If Not goUtil.goProgForm Is Nothing Then
                        If goUtil.goProgForm.CancelMe Then
                            GoTo CLEAN_UP
                        End If
                    Else
                        GoTo CLEAN_UP
                    End If
                    Select Case UCase(sTable)
                        Case UCase("RTPhotoLog")
                            'Need to Remove any Photos
                            bRemoveFiles = True
                            sSQL = "SELECT IDAssignments, "
                            sSQL = sSQL & "PhotoName "
                            sSQL = sSQL & "FROM RTPhotoLog "
                            sSQL = sSQL & "WHERE IDAssignments = " & IDAssignments & " "
                            sDelSQL = "DELETE * "
                            sDelSQL = sDelSQL & "FROM " & sTable & " "
                            sDelSQL = sDelSQL & "WHERE IDAssignments = " & IDAssignments & " "
                        Case UCase("RTAttachments")
                            'Need to remove any attachments
                            bRemoveFiles = True
                            sSQL = "SELECT IDAssignments, "
                            sSQL = sSQL & "Attachment "
                            sSQL = sSQL & "FROM RTAttachments "
                            sSQL = sSQL & "WHERE IDAssignments = " & IDAssignments & " "
                            sDelSQL = "DELETE * "
                            sDelSQL = sDelSQL & "FROM " & sTable & " "
                            sDelSQL = sDelSQL & "WHERE IDAssignments = " & IDAssignments & " "
                        Case UCase("RTWSDiagram")
                            'Need to remove any Diagram Photos
                            bRemoveFiles = True
                            sSQL = "SELECT IDAssignments, "
                            sSQL = sSQL & "DiagramPhotoName "
                            sSQL = sSQL & "FROM RTWSDiagram "
                            sSQL = sSQL & "WHERE IDAssignments = " & IDAssignments & " "
                            sDelSQL = "DELETE * "
                            sDelSQL = sDelSQL & "FROM " & sTable & " "
                            sDelSQL = sDelSQL & "WHERE IDAssignments = " & IDAssignments & " "
                        Case UCase("Assignments")
                            'Need to remove any attached Loss Reports
                            bRemoveFiles = True
                            sSQL = "SELECT ID, "
                            sSQL = sSQL & "IBNUM, "
                            sSQL = sSQL & "LossReport, "
                            sSQL = sSQL & "LRFormat "
                            sSQL = sSQL & "FROM Assignments "
                            sSQL = sSQL & "WHERE ID = " & IDAssignments & " "
                            sDelSQL = "DELETE * "
                            sDelSQL = sDelSQL & "FROM " & sTable & " "
                            sDelSQL = sDelSQL & "WHERE ID = " & IDAssignments & " "
                        Case Else
                            bRemoveFiles = False
                            sDelSQL = "DELETE * "
                            sDelSQL = sDelSQL & "FROM " & sTable & " "
                            sDelSQL = sDelSQL & "WHERE IDAssignments = " & IDAssignments & " "
                    End Select
                    
                    'Check for Removing of files first
                    If bRemoveFiles Then
                        'Use Disconnected Record Set on asUseClient Cusor ONLY !
                        Set RSTable = New ADODB.Recordset
                        RSTable.CursorLocation = adUseClient
                        RSTable.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
                        Set RSTable.ActiveConnection = Nothing
                        
                        If Not RSTable.EOF Then
                            lFileValue = 0
                            lFileValueMax = RSTable.RecordCount
                            goUtil.goProgForm.PBarFile.Max = lFileValueMax
                            Do Until RSTable.EOF
                                lFileValue = lFileValue + 1
                                goUtil.goProgForm.PBarFile.Value = lFileValue
                                Select Case UCase(sTable)
                                    Case UCase("RTPhotoLog")
                                        sPhotoName = Trim(RSTable!PhotoName)
                                        sPhotoThumbName = Replace(sPhotoName, "_1.jpg", "_2.jpg", , , vbTextCompare)
                                        sPhotoHighResName = Replace(sPhotoName, "_1.jpg", "_0.jpg", , , vbTextCompare)
                                        'Delete each photo
                                        goUtil.goProgForm.lblFileText = sPhotoName
                                        goUtil.goProgForm.RefreshMe
                                        DoEvents
                                        Sleep 10
                                        'Check for user cancel
                                        If Not goUtil.goProgForm Is Nothing Then
                                            If goUtil.goProgForm.CancelMe Then
                                                GoTo CLEAN_UP
                                            End If
                                        Else
                                            GoTo CLEAN_UP
                                        End If
                                        sPhotoName = left(sPhotoName, InStr(1, sPhotoName, "_", vbBinaryCompare)) & "*"
                                        goUtil.utDeleteFile msPhotoReposPath & "\" & sPhotoName
                                        goUtil.goProgForm.lblFileText = sPhotoThumbName
                                        goUtil.goProgForm.RefreshMe
                                        DoEvents
                                        Sleep 10
                                        'Check for user cancel
                                        If Not goUtil.goProgForm Is Nothing Then
                                            If goUtil.goProgForm.CancelMe Then
                                                GoTo CLEAN_UP
                                            End If
                                        Else
                                            GoTo CLEAN_UP
                                        End If
                                        sPhotoThumbName = left(sPhotoThumbName, InStr(1, sPhotoThumbName, "_", vbBinaryCompare)) & "*"
                                        goUtil.utDeleteFile msPhotoReposPath & "\" & sPhotoThumbName
                                        goUtil.goProgForm.lblFileText = sPhotoHighResName
                                        goUtil.goProgForm.RefreshMe
                                        DoEvents
                                        Sleep 10
                                        'Check for user cancel
                                        If Not goUtil.goProgForm Is Nothing Then
                                            If goUtil.goProgForm.CancelMe Then
                                                GoTo CLEAN_UP
                                            End If
                                        Else
                                            GoTo CLEAN_UP
                                        End If
                                        sPhotoHighResName = left(sPhotoHighResName, InStr(1, sPhotoHighResName, "_", vbBinaryCompare)) & "*"
                                        goUtil.utDeleteFile msPhotoReposPath & "\" & sPhotoHighResName
                                    Case UCase("RTAttachments")
                                        sAttach = Trim(RSTable!Attachment)
                                        'Delete Attachment
                                        goUtil.goProgForm.lblFileText = sAttach
                                        goUtil.goProgForm.RefreshMe
                                        DoEvents
                                        Sleep 10
                                        'Check for user cancel
                                        If Not goUtil.goProgForm Is Nothing Then
                                            If goUtil.goProgForm.CancelMe Then
                                                GoTo CLEAN_UP
                                            End If
                                        Else
                                            GoTo CLEAN_UP
                                        End If
                                        sAttach = left(sAttach, InStr(1, sAttach, "_", vbBinaryCompare)) & "*"
                                        goUtil.utDeleteFile msAttachReposPath & "\" & sAttach
                                    Case UCase("RTWSDiagram")
                                        sDiagramPhotoName = Trim(RSTable!DiagramPhotoName)
                                        'Delete each photo
                                        goUtil.goProgForm.lblFileText = sDiagramPhotoName
                                        goUtil.goProgForm.RefreshMe
                                        DoEvents
                                        Sleep 10
                                        'Check for user cancel
                                        If Not goUtil.goProgForm Is Nothing Then
                                            If goUtil.goProgForm.CancelMe Then
                                                GoTo CLEAN_UP
                                            End If
                                        Else
                                            GoTo CLEAN_UP
                                        End If
                                        sDiagramPhotoName = left(sDiagramPhotoName, InStr(1, sDiagramPhotoName, "_", vbBinaryCompare)) & "*"
                                        goUtil.utDeleteFile msPhotoReposPath & "\" & sDiagramPhotoName
                                    Case UCase("Assignments")
                                        sAttach = left(Trim(RSTable!LRFormat), 11)
                                        If StrComp(sAttach, "OLEType_pdf", vbTextCompare) = 0 Then
                                            sAttach = Trim(RSTable!LossReport)
                                            'Delete Loss Report
                                            goUtil.goProgForm.lblFileText = sAttach
                                            goUtil.goProgForm.RefreshMe
                                            DoEvents
                                            Sleep 10
                                            'Check for user cancel
                                            If Not goUtil.goProgForm Is Nothing Then
                                                If goUtil.goProgForm.CancelMe Then
                                                    GoTo CLEAN_UP
                                                End If
                                            Else
                                                GoTo CLEAN_UP
                                            End If
                                            sAttach = left(sAttach, InStr(1, sAttach, "_", vbBinaryCompare)) & "*"
                                            goUtil.utDeleteFile msAttachReposPath & "\" & sAttach
                                        End If
                                        'Also Delete any Word Docuemnts(.doc) or Excel documents(.xls)
                                        sAttach = Trim(RSTable!IBNUM) & "_*.doc"
                                        goUtil.utDeleteFile msAttachReposPath & "\" & sAttach
                                        sAttach = Trim(RSTable!IBNUM) & "_*.xls"
                                        goUtil.utDeleteFile msAttachReposPath & "\" & sAttach
                                End Select
                                RSTable.MoveNext
                            Loop
                        End If
                    End If
                    
                    'Delete Records from table
                    oConn.Execute sDelSQL
                    
                Next
                RS.MoveNext
            Loop
        Else 'If MsgBox(sMess, vbQuestion + vbYesNo, "Delete CAT") = vbNo
            Exit Function
        End If
    End If
    
    mbDeleting = False
    RecycleBinCatCommand = True
CLEAN_UP:
    'cleanup
    If Not goUtil.goProgForm Is Nothing Then
        goUtil.goProgForm.CLEANUP
        Set goUtil.goProgForm = Nothing
    End If
    Set colTables = Nothing
    Set RS = Nothing
    Set RSTable = Nothing
    Set oConn = Nothing
    'If being called from Empty Recycle bin then this shold be False
    'since It will be called by Empty Recycle bin
    If pbLoadTree Then
        LoadTree
    End If
    Exit Function
EH:
    mbDeleting = False
    Set colTables = Nothing
    Set RS = Nothing
    Set RSTable = Nothing
    Set oConn = Nothing
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Function RecycleBinCatCommand"
End Function

Private Function GetCatFromDiskCommand(psCommand As String) As Boolean
    On Error GoTo EH
    Dim sMess As String
    'Execute Command
    If StrComp(psCommand, "GetCatBackup", vbTextCompare) = 0 Then
        If Not goUtil.goCurCarList Is Nothing Then
            goUtil.goCurCarList.CLEANUP
            Set goUtil.goCurCarList = Nothing
        End If
        'Close Currnet DB
        goUtil.CloseCurDB
        goUtil.gsCurCatDir = vbNullString
'        cmdConnect.Enabled = False
'        cmdPrint.Enabled = False
        If goUtil.SendGetCATFromDisk(App.EXEName, False, True) Then
            sMess = "Get Cat Backup Complete!"
        End If
    ElseIf StrComp(psCommand, "ImportCat", vbTextCompare) = 0 Then
        If goUtil.SendGetCATFromDisk(App.EXEName, False, False) Then
            sMess = "Import Cat Complete!"
        End If
    End If
    If sMess <> vbNullString Then
        LoadTree
        MsgBox sMess, vbInformation + vbOKOnly, "Get Cat From Disk"
    End If
    GetCatFromDiskCommand = True
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Function GetCatFromDiskCommand"
End Function

Private Function CatCommand(psCompany As String, psCar As String, psCat As String, psCommand As String) As Boolean
    On Error GoTo EH
    '--------------Add Claim-------------------
    Dim sHTML As String
    Dim oLR As V2ECKeyBoard.clsLossReports
    Dim sWebSite As String
    Dim sUserName As String
    Dim sPassword As String
    Dim bUseSSL As Boolean
    Dim sSelCompany As String
    Dim sSelClientCompany As String
    Dim sSelAssignmentType As String
    Dim sselCat As String
    Dim sselClientCompanyUser As String
    Dim oConn As New ADODB.Connection   'Needed to get the correct assignment type
    Dim RS As ADODB.Recordset           'Needed to get the correct assignment type
    Dim sSQL As String
    '--------------End Add Claim-------------------
    
    If goUtil.goCurCarList Is Nothing Then
        Exit Function
    End If
    
    ' Need to set Mouse back since going inside another object
    Screen.MousePointer = vbNormal
    
    'Check for Carrier List Command
    'A Carrier List command is an item added to the Tree
    'that is Carrier specific, (IE the command only applies to that Specific Carrier.)
    If left(psCommand, Len(CAR_LIST_COMMAND)) = CAR_LIST_COMMAND Then
        goUtil.goCurCarList.CarListCommand psCar, psCat, psCommand
    ElseIf StrComp(psCommand, "AddClaim", vbTextCompare) = 0 Then
        Screen.MousePointer = vbNormal
        bUseSSL = CBool(GetSetting("ECS", "WEB_SECURITY", "USE_SSL", True))
        If bUseSSL Then
            sWebSite = "https://"
        Else
            sWebSite = "http://"
        End If
        sWebSite = sWebSite & GetSetting("ECS", "WEB_SECURITY", "WEB_HOST", "WWW.EBERLS.NET")
        sUserName = goUtil.utGetECSCryptSetting("ECS", "WEB_SECURITY", "USER_NAME", vbNullString)
        sPassword = goUtil.utGetECSCryptSetting("ECS", "WEB_SECURITY", "PASSWORD", vbNullString)
        sHTML = goUtil.utGetFileData(goUtil.gsInstallDir & "\Templates\AddClaimLogon.html")
        sSelCompany = goUtil.gsCurCompany
        sSelClientCompany = goUtil.gsCurCar
        sselCat = goUtil.gsCurCat
        'need to get Correct Assignment type
        Set oConn = New ADODB.Connection
        goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
        Set RS = New ADODB.Recordset
        RS.CursorLocation = adUseClient
        sSQL = "SELECT  [AssignmentTypeID] "
        sSQL = sSQL & "FROM     CAT "
        sSQL = sSQL & "WHERE    [CATID] = " & goUtil.gsCurCat & " "
        sSQL = sSQL & "AND      [CompanyID] = " & goUtil.gsCurCompany & " "
        
        RS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
        Set RS.ActiveConnection = Nothing
        RS.MoveFirst
        sSelAssignmentType = goUtil.IsNullIsVbNullString(RS.Fields("AssignmentTypeID"))
        
        Set RS = Nothing
        Set oConn = Nothing
        
        sselClientCompanyUser = goUtil.gsCurUsersID
        'Need to Replace Marked Text with above vars
        If sHTML <> vbNullString Then
            sHTML = Replace(sHTML, "|ENTER_WEB_SITE|", sWebSite, , , vbTextCompare)
            sHTML = Replace(sHTML, "|ENTER_USER_NAME|", sUserName, , , vbTextCompare)
            sHTML = Replace(sHTML, "|ENTER_PASSWORD|", sPassword, , , vbTextCompare)
            sHTML = Replace(sHTML, "|ENTER_SelCompany|", sSelCompany, , , vbTextCompare)
            sHTML = Replace(sHTML, "|ENTER_SelClientCompany|", sSelClientCompany, , , vbTextCompare)
            sHTML = Replace(sHTML, "|ENTER_SelAssignmentType|", sSelAssignmentType, , , vbTextCompare)
            sHTML = Replace(sHTML, "|ENTER_selCat|", sselCat, , , vbTextCompare)
            sHTML = Replace(sHTML, "|ENTER_selClientCompanyUser|", sselClientCompanyUser, , , vbTextCompare)
        Else
            MsgBox "Could Not Find " & goUtil.gsInstallDir & "\Templates\AssignmentsLogon.html", vbCritical + vbOKOnly, "Error"
            Exit Function
        End If
        
        Set oLR = New V2ECKeyBoard.clsLossReports
        oLR.SetUtilObject goUtil
        
        oLR.ShowHelpViewer sHTML, "Add Claim Logon", True
        oLR.CLEANUP
        Set oLR = Nothing
    Else
        goUtil.goCurCarList.CatCommand psCar, psCat, psCommand
    End If
    
    
    CatCommand = True
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Function CatCommand"
End Function

Private Function CatMaintenanceCommand(psCompany As String, psCar As String, psCat As String, psCommand As String) As Boolean
    On Error GoTo EH
    
    'Execute Command
    If StrComp(psCommand, "SendCatToXactimate", vbTextCompare) = 0 Then
        If Not goUtil.goCurCarList Is Nothing Then
            ' Need to set Mouse back since going inside another object
            Screen.MousePointer = vbNormal
            FlagSendToXactimate = True
            goUtil.goCurCarList.SendToXactimate
            FlagSendToXactimate = False
            If Not goUtil Is Nothing Then
                If Not goUtil.goXact Is Nothing Then
                    goUtil.goXact.CLEANUP
                    Set goUtil.goXact = Nothing
                End If
            Else
                Exit Function
            End If
        End If
    ElseIf StrComp(psCommand, "UpdateCat", vbTextCompare) = 0 Then
        If Not CloseFTPDBConnection Then
            Exit Function
        End If
        If Not goUtil.goCurCarList Is Nothing Then
            goUtil.goCurCarList.CLEANUP
            Set goUtil.goCurCarList = Nothing
        End If
        
        SetCarrierGlobalObjects psCompany, psCar, True, True
        
    ElseIf StrComp(psCommand, "SendCatToRecycleBin", vbTextCompare) = 0 Then
        If MsgBox("Are You sure?", vbQuestion + vbYesNo, "Send To Recycle Bin") = vbYes Then
            If Not goUtil.goCurCarList Is Nothing Then
                goUtil.goCurCarList.CLEANUP
                Set goUtil.goCurCarList = Nothing
            End If
            SaveSetting App.EXEName, "CAT_STRUCTURE\" & psCompany & "\" & psCar & "\" & psCat, psCat, goUtil.Encode("RECYCLEBIN")
            goUtil.gsCurCatDir = vbNullString
            LoadTree
        End If
    End If
    
    CatMaintenanceCommand = True
    Exit Function
EH:
    FlagSendToXactimate = False
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Function CatMaintenanceCommand"
End Function

Private Function CatMaintSendCatToDiskCommand(psCar As String, psCat As String, psCommand As String) As Boolean
    On Error GoTo EH
    Dim sMess As String
    
    'Execute Command
    If StrComp(psCommand, "CreateCatBackup", vbTextCompare) = 0 Then
        If goUtil.SendGetCATFromDisk(App.EXEName, True, True, psCar, psCat) Then
            sMess = "Backup Complete!"
        End If
    ElseIf StrComp(psCommand, "ExportCat", vbTextCompare) = 0 Then
        If goUtil.SendGetCATFromDisk(App.EXEName, True, False, psCar, psCat) Then
            sMess = "Export Complete!"
        End If
    End If
    If sMess <> vbNullString Then
        LoadTree
        MsgBox sMess, vbInformation + vbOKOnly, "Send Cat To Disk"
    End If
    CatMaintSendCatToDiskCommand = True
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Function CatMaintSendCatToDiskCommand"
End Function

Public Sub MessSendToXactimate(psTitle As String)
    On Error GoTo EH
    Dim sMess As String
    
    sMess = "Please wait!  Send To Xactimate is still Active." & vbCrLf & vbCrLf
    sMess = sMess & "Press ""OK"" to continue to wait." & vbCrLf
    sMess = sMess & "Press ""Cancel"" to stop Send To Xactimate."
    
    If MsgBox(sMess, vbExclamation + vbOKCancel, psTitle) = vbCancel Then
        If Not goUtil Is Nothing Then
            If Not goUtil.goXact Is Nothing Then
                goUtil.goXact.CancelSendToXactimate = True
            End If
        End If
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub MessSendToXactimate"
End Sub

Public Function SendExportToXactimate() As Boolean
    On Error GoTo EH
    Dim colGlobalObjects As Collection
    
    Set colGlobalObjects = New Collection
    colGlobalObjects.Add goUtil, "goUtil"
        
    If goUtil.goECKeyBoardList Is Nothing Then
       Set goUtil.goECKeyBoardList = New V2ECKeyBoard.clsLists
       goUtil.goECKeyBoardList.SetUtilObject goUtil
       goUtil.SetGlobalObjects colGlobalObjects
    End If
    
    If goUtil.gARV Is Nothing Then
       Set goUtil.gARV = New V2ARViewer.clsARViewer
       goUtil.gARV.SetGlobalObjects colGlobalObjects
    End If

    
    If Not goUtil.goXact Is Nothing Then
        goUtil.goXact.CLEANUP
        Set goXact = Nothing
    End If
    
    Set goUtil.goXact = New V2ECKeyBoard.clsXact
    goUtil.goXact.StartHwnd = goUtil.gfrmECTray.hWnd
    
    If Not goUtil.goXact.LookUpLoaded Then
        goUtil.goXact.CLEANUP
        Set goUtil.goXact = Nothing
        SendExportToXactimate = True
        GoTo SKIPXACT
    End If
    
    If Not goUtil Is Nothing Then
        If goUtil.goXact Is Nothing Then
            GoTo SKIPXACT
        End If
    Else
        GoTo SKIPXACT
    End If
    
    'BGS we still have not launched Xactimate yet,
    'after ValidateXactprojects Passes then we update
    'Easy Claim DB then check to see if we need to launch xactimate
    goUtil.goXact.GetFromExport = True
    If goUtil.goXact.ValidateXactProjects() Then
        If Not goUtil Is Nothing Then
            If goUtil.goXact Is Nothing Then
               GoTo SKIPXACT
            End If
        Else
            GoTo SKIPXACT
        End If
        
        'If we are not skiping all then we need to SendToXactimate
        'SendToXact will launch Xactimate if it isn't already loaded and set focus to it
        If Not goUtil.goXact.SkipAll Then
            'Display the Cancel Form
            Set mFrmCancel = New frmCancel
            Set mFrmCancel.Util = goUtil
            Load mFrmCancel
            mFrmCancel.Visible = True
            mFrmCancel.lblMess.Caption = "Please wait! Sending files..."
            If goUtil.goXact.SendToXact() Then
                'OK we need a check here incase user tries to unload in middle of sending
                If Not goUtil Is Nothing Then
                    If goUtil.goXact Is Nothing Then
                       GoTo SKIPXACT
                    End If
                Else
                    GoTo SKIPXACT
                End If
                
                goUtil.goXact.LookForWindow left(goUtil.gfrmECTray.Caption, 10)
                If goUtil.goXact.SendToExport Then
                    MsgBox "Project(s) Exported to: " & vbCrLf & vbCrLf & goUtil.goXact.ExportFilePath, vbInformation + vbOKOnly, "Send To Xactimate Export File"
                Else
                    MsgBox "Project(s) Sent to Xactimate From Export File!", vbInformation + vbOKOnly, "Send To Xactimate From Export File"
                End If
                
            Else
                'OK we need a check here incase user tries to unload in middle of sending
                If Not goUtil Is Nothing Then
                    If goUtil.goXact Is Nothing Then
                        GoTo SKIPXACT
                    End If
                Else
                    GoTo SKIPXACT
                End If
                
                goUtil.goXact.LookForWindow left(goUtil.gfrmECTray.Caption, 10)
                
                If goUtil.goXact.SendToExport Then
                    MsgBox "Project(s) Failed to Export to: " & vbCrLf & vbCrLf & goUtil.goXact.ExportFilePath, vbExclamation + vbOKOnly, "Send To Xactimate Export File"
                Else
                    MsgBox "Project(s) Not Sent to Xactimate From Export File!", vbExclamation + vbOKOnly, "Send To Xactimate From Export File"
                End If
                
            End If
        End If
    End If
    SendExportToXactimate = True
SKIPXACT:
    If Not mFrmCancel Is Nothing Then
        Unload mFrmCancel
        Set mFrmCancel = Nothing
    End If
    If Not goUtil.goXact Is Nothing Then
        goUtil.goXact.CLEANUP
    End If
    Set goUtil.goXact = Nothing
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function SendExportToXactimate"
    SendExportToXactimate = False
    If Not goUtil.goXact Is Nothing Then
        goUtil.goXact.CLEANUP
    End If
    Set goUtil.goXact = Nothing
    If Not mFrmCancel Is Nothing Then
        Unload mFrmCancel
        Set mFrmCancel = Nothing
    End If
    
End Function

Public Function CompactAndRepairMainDB() As Boolean
    On Error GoTo EH
    
    'Running the Compact Repair.
    Screen.MousePointer = vbHourglass
    goUtil.CloseMainDB
    goUtil.SetMainDB App.EXEName, goUtil.gsInstallDir & "\ECMain.mdb", , True, True
    Screen.MousePointer = vbNormal
    
    CompactAndRepairMainDB = True
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function CompactAndRepairMainDB"
End Function

Public Function GetCurCatName() As String
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim RS As ADODB.Recordset
    Dim sSQL As String
    
    
    'need to get Correct Assignment type
    Set oConn = New ADODB.Connection
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    Set RS = New ADODB.Recordset
    RS.CursorLocation = adUseClient
    sSQL = "SELECT  [Name] "
    sSQL = sSQL & "FROM     CAT "
    sSQL = sSQL & "WHERE    [CATID] = " & goUtil.gsCurCat & " "
    sSQL = sSQL & "AND      [CompanyID] = " & goUtil.gsCurCompany & " "
    
    RS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set RS.ActiveConnection = Nothing
    RS.MoveFirst
    
    GetCurCatName = goUtil.IsNullIsVbNullString(RS.Fields("Name"))
    
    'cleanup
    Set RS = Nothing
    Set oConn = Nothing
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function GetCurCatName"
End Function
