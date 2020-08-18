VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPrint 
   Caption         =   "File"
   ClientHeight    =   6705
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11820
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6705
   ScaleWidth      =   11820
   StartUpPosition =   3  'Windows Default
   Tag             =   "File"
   Begin VB.Timer Timer_MaximizeMe 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4440
      Top             =   4560
   End
   Begin VB.CommandButton cmdAddMultiReport 
      Caption         =   "&Add"
      Height          =   375
      Left            =   8280
      TabIndex        =   20
      Top             =   600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ComboBox cboPackage 
      Height          =   360
      ItemData        =   "frmPrint.frx":0000
      Left            =   9240
      List            =   "frmPrint.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   600
      Width           =   2250
   End
   Begin VB.Frame framEditPackage 
      Caption         =   "Export Claim File Options"
      Enabled         =   0   'False
      Height          =   1455
      Left            =   5360
      TabIndex        =   11
      Top             =   360
      Width           =   6240
      Begin VB.CheckBox chkXMLOnly 
         Caption         =   "Export XML Only"
         Height          =   285
         Left            =   2040
         TabIndex        =   19
         ToolTipText     =   "Loss Report and Attachments will not have XML export files!"
         Top             =   1080
         Width           =   2055
      End
      Begin VB.CheckBox chkXMLExport 
         Caption         =   "Include XML Export"
         Height          =   285
         Left            =   2040
         TabIndex        =   18
         ToolTipText     =   "Loss Report and Attachments will not have XML export files!"
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox txtPassWord 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   120
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   13
         Top             =   400
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtPassWord 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   1080
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   15
         Top             =   400
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmdSavePackage 
         Caption         =   "<< P&rint to File"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         Picture         =   "frmPrint.frx":0004
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   720
         Width           =   1830
      End
      Begin VB.CheckBox chkUsePass 
         Caption         =   "Use password"
         Height          =   285
         Left            =   2040
         TabIndex        =   17
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label lblPass 
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   200
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblPass 
         Caption         =   "Confirm"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   14
         Top             =   200
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin VB.Frame framAdjSavedPackages 
      Caption         =   "Saved ZIP Files (Not uploaded to server!)"
      Height          =   1455
      Left            =   210
      TabIndex        =   1
      Top             =   360
      Width           =   5040
      Begin VB.CommandButton cmdMailTo 
         Caption         =   "Mail To:"
         Height          =   375
         Left            =   1848
         TabIndex        =   3
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton cmdViewFile 
         Caption         =   "View &File"
         Height          =   375
         Left            =   3575
         TabIndex        =   4
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdDelItem 
         Caption         =   "&Delete File"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
      Begin MSComctlLib.ListView lvwSavedPackages 
         Height          =   690
         Left            =   120
         TabIndex        =   5
         Tag             =   "Enable"
         Top             =   675
         Width           =   4790
         _ExtentX        =   8440
         _ExtentY        =   1217
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "imgVarDoc"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.Frame framPackage 
      Caption         =   "Claim File"
      Enabled         =   0   'False
      Height          =   5175
      Left            =   75
      TabIndex        =   0
      Top             =   120
      Width           =   11655
      Begin VB.Frame framPackageItemMaint 
         Caption         =   "Claim File Maintenance"
         Enabled         =   0   'False
         Height          =   3375
         Left            =   5280
         TabIndex        =   22
         Top             =   1680
         Width           =   6255
         Begin VB.CommandButton cmdSendMe 
            Caption         =   "Res&end Item(s) to Client"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   2760
            MaskColor       =   &H00000000&
            Picture         =   "frmPrint.frx":014E
            Style           =   1  'Graphical
            TabIndex        =   30
            ToolTipText     =   "Exit"
            Top             =   240
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CommandButton cmdFindNext 
            Caption         =   "Find &Next"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   720
            TabIndex        =   28
            Top             =   960
            Width           =   1100
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "F&ind"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   720
            TabIndex        =   27
            Top             =   720
            Width           =   1100
         End
         Begin VB.CommandButton cmdSelAll 
            Caption         =   "&Select A&ll"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   720
            TabIndex        =   26
            Top             =   480
            Width           =   1100
         End
         Begin VB.CommandButton cmdSaveSort 
            Caption         =   "&Save Sort"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   720
            TabIndex        =   25
            Top             =   240
            Width           =   1100
         End
         Begin VB.CheckBox chkPrintPreview 
            Alignment       =   1  'Right Justify
            Caption         =   "Print Preview"
            Height          =   375
            Left            =   4530
            TabIndex        =   31
            Top             =   240
            Width           =   1575
         End
         Begin VB.CommandButton cmdPrint 
            Caption         =   "&Print"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5490
            TabIndex        =   34
            Top             =   840
            Width           =   615
         End
         Begin VB.ComboBox cboCopy 
            Height          =   360
            ItemData        =   "frmPrint.frx":0298
            Left            =   3600
            List            =   "frmPrint.frx":029A
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   840
            Width           =   1815
         End
         Begin VB.CommandButton cmdDown 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   120
            Picture         =   "frmPrint.frx":029C
            Style           =   1  'Graphical
            TabIndex        =   24
            ToolTipText     =   "Move Selected Item DOWN"
            Top             =   720
            Width           =   480
         End
         Begin VB.CommandButton cmdUp 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   120
            Picture         =   "frmPrint.frx":06DE
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Move Selected Item UP"
            Top             =   240
            Width           =   480
         End
         Begin VB.CommandButton cmdDelPackageItem 
            Caption         =   "&<< Remove"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   1800
            Picture         =   "frmPrint.frx":0B20
            Style           =   1  'Graphical
            TabIndex        =   29
            ToolTipText     =   "Click to mark selected item(s) as deleted"
            Top             =   240
            Width           =   945
         End
         Begin MSComctlLib.ListView lvwPackageItem 
            Height          =   2055
            Left            =   120
            TabIndex        =   35
            Tag             =   "Enable"
            Top             =   1200
            Width           =   6015
            _ExtentX        =   10610
            _ExtentY        =   3625
            View            =   3
            Arrange         =   2
            LabelEdit       =   1
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            OLEDragMode     =   1
            FullRowSelect   =   -1  'True
            _Version        =   393217
            Icons           =   "imgPackageStatus"
            SmallIcons      =   "imgPackageStatus"
            ColHdrIcons     =   "imgPackageStatus"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            OLEDragMode     =   1
            NumItems        =   0
         End
         Begin VB.Label lblCopy 
            Caption         =   "Copies apply to IB only."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3600
            TabIndex        =   32
            Top             =   645
            Width           =   1815
         End
      End
      Begin VB.Frame framAvailRptDoc 
         Caption         =   "Document Items that were removed by you"
         Enabled         =   0   'False
         Height          =   3375
         Left            =   120
         TabIndex        =   6
         Top             =   1680
         Width           =   5055
         Begin VB.ListBox cboMainReports 
            Height          =   2700
            ItemData        =   "frmPrint.frx":0C6A
            Left            =   120
            List            =   "frmPrint.frx":0C6C
            Sorted          =   -1  'True
            TabIndex        =   9
            Top             =   480
            Width           =   3975
         End
         Begin VB.CommandButton cmdAddAvailRptDoc 
            Caption         =   "Add&>>"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   4080
            TabIndex        =   10
            Top             =   480
            Width           =   855
         End
         Begin VB.CheckBox chkAddAllAvailRptDoc 
            Alignment       =   1  'Right Justify
            Caption         =   "(Uncheck to undelete Documents one at a time.)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   840
            TabIndex        =   8
            Top             =   240
            Width           =   4095
         End
         Begin VB.Label lblMainReports 
            Caption         =   "Reports:"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   2655
         End
      End
   End
   Begin VB.Frame framCommands 
      Height          =   1215
      Left            =   75
      TabIndex        =   36
      Top             =   5280
      Width           =   11655
      Begin VB.TextBox txtAdminComments 
         BackColor       =   &H80000018&
         Height          =   855
         Left            =   5400
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   39
         ToolTipText     =   "Double Click to view text"
         Top             =   240
         Width           =   3975
      End
      Begin VB.CommandButton cmdPrintList 
         Caption         =   "Sho&w Item List"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1080
         MaskColor       =   &H00000000&
         Picture         =   "frmPrint.frx":0C6E
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Exit"
         Top             =   240
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkMaxPackageMaint 
         Alignment       =   1  'Right Justify
         Caption         =   "&Claim File Options"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   240
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Sa&ve"
         Enabled         =   0   'False
         Height          =   855
         Left            =   9480
         MaskColor       =   &H00000000&
         Picture         =   "frmPrint.frx":10B0
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Exit"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   855
         Left            =   10560
         MaskColor       =   &H00000000&
         Picture         =   "frmPrint.frx":11FA
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Exit"
         Top             =   240
         Width           =   975
      End
   End
   Begin MSComctlLib.ImageList imgPackageStatus 
      Left            =   5640
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrint.frx":1504
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrint.frx":1958
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrint.frx":1DAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrint.frx":2198
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrint.frx":25EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrint.frx":2A40
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Const MAIN_REPORT As Long = 1
Private Const CARSPEC_REPORT As Long = 2
Private Const IB_REPORT As Long = 3
Private Const PAYMENTS_REPORT As Long = 4
Private Const ATTACHMENTS_REPORT As Long = 5
Private Const DELETED_SORTORDER As Long = -999

Private mfrmClaim As frmClaim
Private mbUnloadMe As Boolean
Private mbEditMode As Boolean
Private msAssignmentsID As String
Private moGUI As V2ECKeyBoard.clsCarGUI
Private mbLoading As Boolean 'Loading Form
Private mbLoadingMe As Boolean 'Just loading data
Private msPackageID As String
Private mlLastFindIndex As Long
Private msFindText As String
Private mbPrintPreview As Boolean
Private mbXMLExport As Boolean
Private mbXMLOnly As Boolean
Private mbSaveSort As Boolean

Public Property Let MyGUI(poGUI As V2ECKeyBoard.clsCarGUI)
    Set moGUI = poGUI
End Property
Public Property Set MyGUI(poGUI As V2ECKeyBoard.clsCarGUI)
    Set moGUI = poGUI
End Property
Public Property Get MyGUI() As V2ECKeyBoard.clsCarGUI
    Set MyGUI = moGUI
End Property

Public Property Let AssignmentsID(psAssignmentsID As String)
    msAssignmentsID = psAssignmentsID
End Property
Public Property Get AssignmentsID() As String
    AssignmentsID = msAssignmentsID
End Property

Public Property Let UnloadMe(pbFlag As Boolean)
    mbUnloadMe = pbFlag
End Property
Public Property Get UnloadMe() As Boolean
    UnloadMe = mbUnloadMe
End Property

Public Property Let MyfrmClaim(pofrmClaim As Object)
    Set mfrmClaim = pofrmClaim
End Property
Public Property Set MyfrmClaim(pofrmClaim As Object)
    Set mfrmClaim = pofrmClaim
End Property
Public Property Get MyfrmClaim() As Object
    Set MyfrmClaim = mfrmClaim
End Property

Private Property Get msClassName() As String
    msClassName = Me.Name
End Property



'Private Sub cboAttachments_Click()
'    On Error GoTo EH
'    cboAttachments.ToolTipText = Trim(left(cboAttachments.Text, 100))
'    Exit Sub
'EH:
'    cboAttachments.ToolTipText = vbNullString
'End Sub


'Private Sub cboCarSpecReports_Click()
'    On Error GoTo EH
'    cboCarSpecReports.ToolTipText = Trim(left(cboCarSpecReports.Text, 100))
'    Exit Sub
'EH:
'    cboCarSpecReports.ToolTipText = vbNullString
'End Sub


'Private Sub cboIB_Click()
'    On Error GoTo EH
'    cboIB.ToolTipText = Trim(left(cboIB.Text, 100))
'    Exit Sub
'EH:
'    cboIB.ToolTipText = vbNullString
'End Sub


Private Sub cboMainReports_Click()
    On Error GoTo EH
    If Not mbLoading Then
        cboMainReports.ToolTipText = Trim(left(cboMainReports.Text, 100))
    End If
    Exit Sub
EH:
    cboMainReports.ToolTipText = Err.Description
End Sub


Private Sub cboPackage_Click()
    On Error GoTo EH
    Dim lCount As Long
    Dim sIBText As String
    Dim sIBListText As String
    Dim sMess As String
    
    If mbUnloadMe Then
        Exit Sub
    End If
    
    If mbEditMode Then
        If cmdSave.Enabled Then
            Screen.MousePointer = MousePointerConstants.vbHourglass
            SaveMe
            Screen.MousePointer = MousePointerConstants.vbDefault
            cmdSave.Enabled = False
        End If
    End If
    
    If cboPackage.ListIndex = -1 Or cboPackage.ListIndex = 0 Then
        msPackageID = "0"
        'Only show the Add button if there is nothing currently
        'open.  Adjuster must close the currnt package before adding another.
        'Only allow this button to be visible if there are no exisiting packages
        'In the future we may be creating multiple packages but now now.
        If cboPackage.ListCount = 1 Then
            cmdAddMultiReport.Visible = True
        End If
        
        mbEditMode = False
        lvwPackageItem.ListItems.Clear
        EnableEditFrames False
        cboMainReports.Clear
'        cboCarSpecReports.Clear
'        cboIB.Clear
'        cboPayments.Clear
'        cboAttachments.Clear
        txtAdminComments.Text = vbNullString
    Else
        cmdAddMultiReport.Visible = False
        msPackageID = cboPackage.ItemData(cboPackage.ListIndex)
        mbEditMode = True
        PopulatelvwPackageItem
        LoadReports
        LoadPackageItems msPackageID
        RemoveAddedReports
        EnableEditFrames True
    End If
    
    
    Exit Sub
EH:
    Screen.MousePointer = MousePointerConstants.vbDefault
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cboPackage_Click"
End Sub

Public Sub LoadPackageItems(psPackageID As String)
    On Error GoTo EH
    Dim sPackageID As String
    
    
    mfrmClaim.SetadoRSPackageList msAssignmentsID
    
    If mfrmClaim.adoRSPackageList.RecordCount > 0 Then
        mfrmClaim.adoRSPackageList.MoveFirst
        Do Until mfrmClaim.adoRSPackageList.EOF
            sPackageID = CStr(mfrmClaim.adoRSPackageList.Fields("PackageID").Value)
            If sPackageID = psPackageID Then
                txtAdminComments.Text = goUtil.IsNullIsVbNullString(mfrmClaim.adoRSPackageList.Fields("AdminComments"))
                Exit Do
            End If
            mfrmClaim.adoRSPackageList.MoveNext
        Loop
        mfrmClaim.adoRSPackageList.MoveFirst
    End If
     
    Exit Sub
EH:
   goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub LoadPackageItems"
End Sub

'Private Sub cboPayments_Click()
'    On Error GoTo EH
'    cboPayments.ToolTipText = Trim(left(cboPayments.Text, 100))
'    Exit Sub
'EH:
'    cboPayments.ToolTipText = vbNullString
'End Sub

Private Sub chkAddAllAvailRptDoc_GotFocus()
    chkAddAllAvailRptDoc.BackColor = &H80000018
End Sub

Private Sub chkAddAllAvailRptDoc_LostFocus()
    chkAddAllAvailRptDoc.BackColor = &H8000000F
End Sub

Private Sub chkMaxPackageMaint_Click()
    On Error GoTo EH
    
    If chkMaxPackageMaint.Value = vbChecked Then
        MaxMinPackageMaint False
    Else
        MaxMinPackageMaint True
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub chkMaxPackageMaint_Click"
End Sub

Public Sub MaxMinPackageMaint(pbMaximize As Boolean)
    On Error GoTo EH
    
    If pbMaximize Then
        cboPackage.Visible = False
        framEditPackage.Visible = False
        framAdjSavedPackages.Visible = False
        framAvailRptDoc.Visible = False
    Else
        cboPackage.Visible = True
        framEditPackage.Visible = True
        framAdjSavedPackages.Visible = True
        framAvailRptDoc.Visible = True
    End If
    
    If Not chkMaxPackageMaint.Enabled Then
        chkMaxPackageMaint.Enabled = True
        ReSizeMe
        chkMaxPackageMaint.Enabled = False
    Else
        ReSizeMe
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub MaxMinPackageMaint"
End Sub

Private Sub chkPrintPreview_Click()

    On Error GoTo EH
    If mbLoading Then
        Exit Sub
    End If
    
    If chkPrintPreview.Value = vbChecked Then
        mbPrintPreview = True
    Else
        mbPrintPreview = False
    End If
    
    SaveSetting App.EXEName, "GENERAL", "PRINT_PREVIEW", mbPrintPreview
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub chkPrintPreview_Click"
End Sub

Private Sub chkUsePass_Click()
    On Error GoTo EH
    
    If chkUsePass.Value = vbChecked Then
        lblPass(0).Visible = True
        txtPassWord(0).Visible = True
        lblPass(1).Visible = True
        txtPassWord(1).Visible = True
    Else
        lblPass(0).Visible = False
        txtPassWord(0).Visible = False
        lblPass(1).Visible = False
        txtPassWord(1).Visible = False
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub chkUsePass_Click"
End Sub

Private Sub chkXMLExport_Click()
    On Error GoTo EH
    If mbLoading Then
        Exit Sub
    End If
    
    If chkXMLExport.Value = vbChecked Then
        mbXMLExport = True
        If chkXMLExport.Enabled Then
            chkXMLOnly.Enabled = True
        End If
    Else
        mbXMLExport = False
        chkXMLOnly.Enabled = False
    End If
    
    SaveSetting App.EXEName, "GENERAL", "XML_EXPORT", mbXMLExport
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub chkXMLExport_Click"
End Sub

Private Sub chkXMLOnly_Click()
    On Error GoTo EH
    If mbLoading Then
        Exit Sub
    End If
    
    If chkXMLOnly.Value = vbChecked Then
        mbXMLOnly = True
    Else
        mbXMLOnly = False
    End If
    
    SaveSetting App.EXEName, "GENERAL", "XML_ONLY", mbXMLOnly
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub chkXMLOnly_Click"
End Sub

Private Sub cmdAddAvailRptDoc_Click(Index As Integer)
    On Error GoTo EH
    Dim itmX As ListItem
    Dim colRemoveItems As Collection
    Dim lCount As Long
    Dim vCount As Variant
    Dim bAddedReport As Boolean
    Select Case Index
        Case MAIN_REPORT
            If AddReportItem(cboMainReports) Then
                cboMainReports.RemoveItem cboMainReports.ListIndex
                bAddedReport = True
            End If
        Case CARSPEC_REPORT
'            If AddReportItem(cboCarSpecReports) Then
'                cboCarSpecReports.RemoveItem cboCarSpecReports.ListIndex
'                bAddedReport = True
'            End If
        Case IB_REPORT
'            If AddReportItem(cboIB) Then
'                cboIB.RemoveItem cboIB.ListIndex
'                bAddedReport = True
'            End If
        Case PAYMENTS_REPORT
'            If AddReportItem(cboPayments) Then
'                cboPayments.RemoveItem cboPayments.ListIndex
'                bAddedReport = True
'            End If
        Case ATTACHMENTS_REPORT
'            If AddReportItem(cboAttachments) Then
'                cboAttachments.RemoveItem cboAttachments.ListIndex
'                bAddedReport = True
'            End If
    End Select
    
    If chkAddAllAvailRptDoc.Value = vbChecked Then
        If bAddedReport Then
            cmdAddAvailRptDoc_Click Index
        End If
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdAddAvailRptDoc_Click"
End Sub

Private Function AddReportItem(pocboReport As Object) As Boolean
    On Error GoTo EH
    Dim MycboReport As ListBox
    Dim itmX As ListItem
    Dim EditItmx As ListItem
    Dim iMyIcon As Long
    Dim sFlagText As String
    Dim MyPackageListItem As GuiPackageItemListView
    Dim MypackageItem As GuiPackageItem
    Dim sReportFormat As String
    Dim sReportFormatTag As String
    Dim sAttachmentsID As String
    Dim sNumber As String  'certain reports can have multiple items per format
    Dim sData() As String
    Dim sAttachName As String  'The actual Attachment File name
    Dim bGetNumber As Boolean
    Dim sSortOrder As String
    Dim sName As String
    Dim sDescription As String
    Dim sTemp As String
    
    
    If Not TypeOf pocboReport Is ListBox Then
        Exit Function
    End If
    Set MycboReport = pocboReport
    
    'Check to see if the listindex is valid
    
    If chkAddAllAvailRptDoc.Value = vbChecked Then
        If MycboReport.ListIndex = -1 Then
            If MycboReport.ListCount > 0 Then
                MycboReport.ListIndex = 0
            End If
        End If
    End If
    
    If MycboReport.ListIndex = -1 Then
        Exit Function
    End If
    
    sReportFormat = MycboReport.Text
    
    'See if this is an atachment
    If InStr(1, Right(sReportFormat, 5), ".pdf|", vbTextCompare) > 0 Then
        sAttachmentsID = MycboReport.ItemData(MycboReport.ListIndex)
        sAttachName = Mid(sReportFormat, InStr(1, sReportFormat, String(200, " "), vbBinaryCompare))
        sAttachName = Trim(sAttachName)
        sData() = Split(sAttachName, "|")
        sAttachName = sData(0)
        'See if this Item already exists in the list
        'if it Does then need to set the EditItmX
        Set EditItmx = GetEditItmX(sReportFormat, , sAttachmentsID)
        sNumber = "Null"
    Else
        sAttachmentsID = "Null"
        sAttachName = vbNullString
        'These items can have multiple Items
        'Thus need to Populate the Number
        If InStr(1, sReportFormat, "ECrpt" & goUtil.gsCurCarDBName & "_arRptPhotos", vbTextCompare) > 0 Then
            bGetNumber = True
        ElseIf InStr(1, sReportFormat, "ECrpt" & goUtil.gsCurCarDBName & "_arWorkSheetDiag", vbTextCompare) > 0 Then
            bGetNumber = True
        ElseIf InStr(1, sReportFormat, "ECrpt" & goUtil.gsCurCarDBName & "_arRptAddlChk", vbTextCompare) > 0 Then
            bGetNumber = True
        ElseIf InStr(1, sReportFormat, "ECrpt" & goUtil.gsCurCarDBName & "_arRptIB" & goUtil.gsCurCarDBName, vbTextCompare) > 0 Then
            bGetNumber = True
        Else
            bGetNumber = False
        End If
        
        If bGetNumber Then
            sNumber = Mid(sReportFormat, InStr(1, sReportFormat, String(200, " "), vbBinaryCompare))
            sNumber = Trim(sNumber)
            sData() = Split(sNumber, "|")
            sNumber = sData(3)
            Set EditItmx = GetEditItmX(sReportFormat, sNumber)
        Else
            sNumber = "Null"
            Set EditItmx = GetEditItmX(sReportFormat)
        End If
        
    End If
    
    sName = sReportFormat
    
    sName = left(sName, InStr(1, sName, "_") - 1)
    sName = Trim(sName)
    
    sDescription = sReportFormat
    sDescription = Mid(sDescription, InStr(1, sDescription, "_") + 1, 200)
    sDescription = Trim(sDescription)
    
    sSortOrder = GetNextSortOrder()
    
    With MypackageItem
        If EditItmx Is Nothing Then
            .PackageItemID = "Null" 'Not Set Here
            .PackageID = msPackageID
            .AssignmentsID = msAssignmentsID
            .ID = "Null" 'not Set Here
            .IDPackage = msPackageID
            .IDAssignments = msAssignmentsID
            .ReportFormat = sReportFormat
            .RTAttachmentsID = sAttachmentsID
            .IDRTAttachments = sAttachmentsID
            .Number = sNumber
            .AttachmentName = sAttachName
            .SortOrder = sSortOrder
            .Name = sName
            .Description = sDescription
            .IsCoApprove = "False"
            .CoApproveDate = "Null"
            .CoApproveDesc = vbNullString
            .IsClientCoReject = "False"
            .ClientCoRejectDate = "Null"
            .ClientCoRejectDesc = vbNullString
            .IsClientCoDelete = "False"
            .ClientCoDeleteDate = "Null"
            .ClientCoDeleteDesc = vbNullString
            .IsClientCoApprove = "False"
            .ClientCoApproveDate = "Null"
            .ClientCoApproveDesc = vbNullString
            .PackageItemGUID = vbNullString
            .SendMe = "True"
            .SentDate = "Null"
            .IsDeleted = "False"
            .DownLoadMe = "False"
            .UpLoadMe = "True"
            .AdminComments = vbNullString
            .DateLastUpdated = Now()
            .UpdateByUserID = goUtil.gsCurUsersID
        Else
            'If the EditItemx is not nothing and the form is loading...
            'that means this process is attempting to add an item that
            'was previously deleted / removed from the package / file (what ever the heck its called at this time)
            'This action should not be allowed to occur. when mbLoading is true
            If mbLoading Then
                GoTo CLEAN_UP
            End If
            .PackageItemID = EditItmx.ListSubItems(GuiPackageItemListView.PackageItemID - 1)
            .PackageID = msPackageID
            .AssignmentsID = msAssignmentsID
            .ID = EditItmx.ListSubItems(GuiPackageItemListView.ID - 1)
            .IDPackage = msPackageID
            .IDAssignments = msAssignmentsID
            .ReportFormat = sReportFormat
            .RTAttachmentsID = sAttachmentsID
            .IDRTAttachments = sAttachmentsID
            .Number = sNumber
            .AttachmentName = sAttachName
            .SortOrder = sSortOrder
            .Name = sName
            .Description = sDescription
            sFlagText = EditItmx.ListSubItems(GuiPackageItemListView.IsCoApprove - 1)
            .IsCoApprove = goUtil.GetFlagFromText(sFlagText)
            .CoApproveDate = EditItmx.ListSubItems(GuiPackageItemListView.CoApproveDate - 1)
            .CoApproveDesc = EditItmx.ListSubItems(GuiPackageItemListView.CoApproveDesc - 1)
            sFlagText = EditItmx.ListSubItems(GuiPackageItemListView.IsClientCoReject - 1)
            .IsClientCoReject = goUtil.GetFlagFromText(sFlagText)
            .ClientCoRejectDate = EditItmx.ListSubItems(GuiPackageItemListView.ClientCoRejectDate - 1)
            .ClientCoRejectDesc = EditItmx.ListSubItems(GuiPackageItemListView.ClientCoRejectDesc - 1)
            sFlagText = EditItmx.ListSubItems(GuiPackageItemListView.IsClientCoDelete - 1)
            .IsClientCoDelete = goUtil.GetFlagFromText(sFlagText)
            .ClientCoDeleteDate = EditItmx.ListSubItems(GuiPackageItemListView.ClientCoDeleteDate - 1)
            .ClientCoDeleteDesc = EditItmx.ListSubItems(GuiPackageItemListView.ClientCoDeleteDesc - 1)
            sFlagText = EditItmx.ListSubItems(GuiPackageItemListView.IsClientCoApprove - 1)
            .IsClientCoApprove = goUtil.GetFlagFromText(sFlagText)
            .ClientCoApproveDate = EditItmx.ListSubItems(GuiPackageItemListView.ClientCoApproveDate - 1)
            .ClientCoApproveDesc = EditItmx.ListSubItems(GuiPackageItemListView.ClientCoApproveDesc - 1)
            .PackageItemGUID = EditItmx.ListSubItems(GuiPackageItemListView.PackageItemGUID - 1)
            .SendMe = "True"
            .SentDate = EditItmx.ListSubItems(GuiPackageItemListView.SentDate - 1)
            .IsDeleted = "False"
            sFlagText = EditItmx.ListSubItems(GuiPackageItemListView.DownLoadMe - 1)
            .DownLoadMe = goUtil.GetFlagFromText(sFlagText)
            .UpLoadMe = "True"
            .AdminComments = EditItmx.ListSubItems(GuiPackageItemListView.AdminComments - 1)
            .DateLastUpdated = Now()
            .UpdateByUserID = goUtil.gsCurUsersID
        End If
    End With
       
    'SortOrder
    If EditItmx Is Nothing Then
        Set itmX = lvwPackageItem.ListItems.Add(, , MypackageItem.SortOrder) '"""" & MyPackageItem.ID & MyPackageItem.SortOrder & """"
    Else
        Set itmX = EditItmx
        itmX.Text = MypackageItem.SortOrder
        'itmX.Key = """" & MyPackageItem.ID & MyPackageItem.SortOrder & """"
    End If
    'SortOrderSort
    itmX.SubItems(GuiPackageItemListView.SortOrderSort - 1) = goUtil.utNumInTextSortFormat(MypackageItem.SortOrder)
    'piName
    itmX.SubItems(GuiPackageItemListView.piName - 1) = MypackageItem.Name
    'Description
    itmX.SubItems(GuiPackageItemListView.Description - 1) = MypackageItem.Description
    'AttachmentName
    itmX.SubItems(GuiPackageItemListView.AttachmentName - 1) = MypackageItem.AttachmentName
    'IsCoApprove
    If CBool(MypackageItem.IsCoApprove) Then
        iMyIcon = GuiPackageItemStatusList.IsApproved
    Else
        iMyIcon = Empty
    End If
    sFlagText = goUtil.GetFlagText(CBool(MypackageItem.IsCoApprove))
    itmX.SubItems(GuiPackageItemListView.IsCoApprove - 1) = sFlagText
    itmX.ListSubItems(GuiPackageItemListView.IsCoApprove - 1).ReportIcon = iMyIcon
    'CoApproveDate
    If Not IsNull(MypackageItem.CoApproveDate) Then
        If IsDate(MypackageItem.CoApproveDate) Then
            itmX.SubItems(GuiPackageItemListView.CoApproveDate - 1) = Format(MypackageItem.CoApproveDate, "MM/DD/YYYY HH:MM:SS")
        Else
            itmX.SubItems(GuiPackageItemListView.CoApproveDate - 1) = vbNullString
        End If
    Else
        itmX.SubItems(GuiPackageItemListView.CoApproveDate - 1) = vbNullString
    End If
    'CoApproveDesc
    itmX.SubItems(GuiPackageItemListView.CoApproveDesc - 1) = MypackageItem.CoApproveDesc
    'IsClientCoReject
    If CBool(MypackageItem.IsClientCoReject) Then
        iMyIcon = GuiPackageItemStatusList.IsRejected
    Else
        iMyIcon = Empty
    End If
    sFlagText = goUtil.GetFlagText(CBool(MypackageItem.IsClientCoReject))
    itmX.SubItems(GuiPackageItemListView.IsClientCoReject - 1) = sFlagText
    itmX.ListSubItems(GuiPackageItemListView.IsClientCoReject - 1).ReportIcon = iMyIcon
    'ClientCoRejectDate
    If Not IsNull(MypackageItem.ClientCoRejectDate) Then
        If IsDate(MypackageItem.ClientCoRejectDate) Then
            itmX.SubItems(GuiPackageItemListView.ClientCoRejectDate - 1) = Format(MypackageItem.ClientCoRejectDate, "MM/DD/YYYY HH:MM:SS")
        Else
            itmX.SubItems(GuiPackageItemListView.ClientCoRejectDate - 1) = vbNullString
        End If
    Else
        itmX.SubItems(GuiPackageItemListView.ClientCoRejectDate - 1) = vbNullString
    End If
    'ClientCoRejectDesc
    itmX.SubItems(GuiPackageItemListView.ClientCoRejectDesc - 1) = MypackageItem.ClientCoRejectDesc
    'IsClientCoApprove
    If CBool(MypackageItem.IsClientCoApprove) Then
        iMyIcon = GuiPackageItemStatusList.IsApproved
    Else
        iMyIcon = Empty
    End If
    sFlagText = goUtil.GetFlagText(CBool(MypackageItem.IsClientCoApprove))
    itmX.SubItems(GuiPackageItemListView.IsClientCoApprove - 1) = sFlagText
    itmX.ListSubItems(GuiPackageItemListView.IsClientCoApprove - 1).ReportIcon = iMyIcon
    'ClientCoApproveDate
    If Not IsNull(MypackageItem.ClientCoApproveDate) Then
        If IsDate(MypackageItem.ClientCoApproveDate) Then
            itmX.SubItems(GuiPackageItemListView.ClientCoApproveDate - 1) = Format(MypackageItem.ClientCoApproveDate, "MM/DD/YYYY HH:MM:SS")
        Else
            itmX.SubItems(GuiPackageItemListView.ClientCoApproveDate - 1) = vbNullString
        End If
    Else
        itmX.SubItems(GuiPackageItemListView.ClientCoApproveDate - 1) = vbNullString
    End If
    'ClientCoApproveDesc
    itmX.SubItems(GuiPackageItemListView.ClientCoApproveDesc - 1) = MypackageItem.ClientCoApproveDesc
    'IsClientCoDelete
    If CBool(MypackageItem.IsClientCoDelete) Then
        iMyIcon = GuiPackageItemStatusList.IsDeleted
    Else
        iMyIcon = Empty
    End If
    sFlagText = goUtil.GetFlagText(CBool(MypackageItem.IsClientCoDelete))
    itmX.SubItems(GuiPackageItemListView.IsClientCoDelete - 1) = sFlagText
    itmX.ListSubItems(GuiPackageItemListView.IsClientCoDelete - 1).ReportIcon = iMyIcon
    'ClientCoDeleteDate
    If Not IsNull(MypackageItem.ClientCoDeleteDate) Then
        If IsDate(MypackageItem.ClientCoDeleteDate) Then
            itmX.SubItems(GuiPackageItemListView.ClientCoDeleteDate - 1) = Format(MypackageItem.ClientCoDeleteDate, "MM/DD/YYYY HH:MM:SS")
        Else
            itmX.SubItems(GuiPackageItemListView.ClientCoDeleteDate - 1) = vbNullString
        End If
    Else
        itmX.SubItems(GuiPackageItemListView.ClientCoDeleteDate - 1) = vbNullString
    End If
    'ClientCoDeleteDesc
    itmX.SubItems(GuiPackageItemListView.ClientCoDeleteDesc - 1) = MypackageItem.ClientCoDeleteDesc
    'SendMe
    If CBool(MypackageItem.SendMe) Then
        iMyIcon = GuiPackageItemStatusList.IsChecked
    Else
        iMyIcon = Empty
    End If
    sFlagText = goUtil.GetFlagText(CBool(MypackageItem.SendMe))
    itmX.SubItems(GuiPackageItemListView.SendMe - 1) = sFlagText
    itmX.ListSubItems(GuiPackageItemListView.SendMe - 1).ReportIcon = iMyIcon
    'SentDate
    If Not IsNull(MypackageItem.SentDate) Then
        If IsDate(MypackageItem.SentDate) Then
            itmX.SubItems(GuiPackageItemListView.SentDate - 1) = Format(MypackageItem.SentDate, "MM/DD/YYYY HH:MM:SS")
        Else
            itmX.SubItems(GuiPackageItemListView.SentDate - 1) = vbNullString
        End If
    Else
        itmX.SubItems(GuiPackageItemListView.SentDate - 1) = vbNullString
    End If
    'IsDeleted
    If CBool(MypackageItem.IsDeleted) Then
        iMyIcon = GuiPackageItemStatusList.IsDeleted
    Else
        iMyIcon = Empty
    End If
    sFlagText = goUtil.GetFlagText(CBool(MypackageItem.IsDeleted))
    itmX.SubItems(GuiPackageItemListView.IsDeleted - 1) = sFlagText
    itmX.ListSubItems(GuiPackageItemListView.IsDeleted - 1).ReportIcon = iMyIcon
    'UpLoadMe
    If CBool(MypackageItem.UpLoadMe) Then
        iMyIcon = GuiPackageItemStatusList.UpLoadMe
    Else
        iMyIcon = Empty
    End If
    sFlagText = goUtil.GetFlagText(CBool(MypackageItem.UpLoadMe))
    itmX.SubItems(GuiPackageItemListView.UpLoadMe - 1) = sFlagText
    itmX.ListSubItems(GuiPackageItemListView.UpLoadMe - 1).ReportIcon = iMyIcon
    'DateLastUpdated
    If Not IsNull(MypackageItem.DateLastUpdated) Then
        If IsDate(MypackageItem.DateLastUpdated) Then
            itmX.SubItems(GuiPackageItemListView.DateLastUpdated - 1) = Format(MypackageItem.DateLastUpdated, "MM/DD/YYYY HH:MM:SS")
        Else
            itmX.SubItems(GuiPackageItemListView.DateLastUpdated - 1) = vbNullString
        End If
    Else
        itmX.SubItems(GuiPackageItemListView.DateLastUpdated - 1) = vbNullString
    End If
    'AdminComments
    itmX.SubItems(GuiPackageItemListView.AdminComments - 1) = MypackageItem.AdminComments
    '----------------------------Hidden Items-----------------------
    'PackageItemID
    itmX.SubItems(GuiPackageItemListView.PackageItemID - 1) = MypackageItem.PackageItemID
    'PackageID
    itmX.SubItems(GuiPackageItemListView.PackageID - 1) = MypackageItem.PackageID
    'AssignmentsID
    itmX.SubItems(GuiPackageItemListView.AssignmentsID - 1) = MypackageItem.AssignmentsID
    'ID
    itmX.SubItems(GuiPackageItemListView.ID - 1) = MypackageItem.ID
    'IDPackage
    itmX.SubItems(GuiPackageItemListView.IDPackage - 1) = MypackageItem.IDPackage
    'IDAssignments
    itmX.SubItems(GuiPackageItemListView.IDAssignments - 1) = MypackageItem.IDAssignments
    'ReportFormat
    itmX.SubItems(GuiPackageItemListView.ReportFormat - 1) = MypackageItem.ReportFormat
    sReportFormatTag = MypackageItem.ReportFormat
    sReportFormatTag = Mid(sReportFormatTag, 200)
    sReportFormatTag = Trim(sReportFormatTag)
    itmX.Tag = sReportFormatTag
    'RTAttachmentsID
    itmX.SubItems(GuiPackageItemListView.RTAttachmentsID - 1) = MypackageItem.RTAttachmentsID
    'IDRTAttachments
    itmX.SubItems(GuiPackageItemListView.IDRTAttachments - 1) = MypackageItem.IDRTAttachments
    'Number
    itmX.SubItems(GuiPackageItemListView.Number - 1) = MypackageItem.Number
    'PackageItemGUID
    itmX.SubItems(GuiPackageItemListView.PackageItemGUID - 1) = MypackageItem.PackageItemGUID
    'DownLoadMe
    sFlagText = goUtil.GetFlagText(CBool(MypackageItem.DownLoadMe))
    itmX.SubItems(GuiPackageItemListView.DownLoadMe - 1) = sFlagText
    
    'UpdateByUserID
    itmX.SubItems(GuiPackageItemListView.UpdateByUserID - 1) = MypackageItem.UpdateByUserID
    itmX.Selected = False
    
    lvwPackageItem.SortKey = GuiPackageItemListView.SortOrder
    lvwPackageItem.Sorted = True
    
CLEAN_UP:
    AddReportItem = True
    
    'cleanup
    Set MycboReport = Nothing
    Set itmX = Nothing
    Set EditItmx = Nothing
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Function AddReportItem"
End Function

Public Function GetEditItmX(Optional psReportFormat As String, _
                            Optional psNumber As String, _
                            Optional psAttachmentsID As String) As ListItem
    On Error GoTo EH
    Dim itmX As ListItem
    Dim sReportFormat As String
    Dim sLookForReportFormat As String
    Dim sNumber As String
    
    'Need to see what type of report looking for
    If psAttachmentsID <> vbNullString Then
        'Looking for attachment item
        For Each itmX In lvwPackageItem.ListItems
            If itmX.SubItems(GuiPackageItemListView.RTAttachmentsID - 1) = psAttachmentsID Then
                Set GetEditItmX = itmX
                Exit Function
            End If
        Next
    ElseIf psNumber <> vbNullString And psReportFormat <> vbNullString Then
        sLookForReportFormat = psReportFormat
        sLookForReportFormat = Mid(sLookForReportFormat, 200)
        sLookForReportFormat = Trim(sLookForReportFormat)
        'Looking for Multi report ite m
        For Each itmX In lvwPackageItem.ListItems
            sReportFormat = itmX.ListSubItems(GuiPackageItemListView.ReportFormat - 1)
            sReportFormat = Mid(sReportFormat, InStr(1, sReportFormat, String(200, " "), vbBinaryCompare))
            sReportFormat = Trim(sReportFormat)
            If StrComp(sReportFormat, sLookForReportFormat, vbTextCompare) = 0 Then
                sNumber = itmX.SubItems(GuiPackageItemListView.Number - 1)
                If sNumber = psNumber Then
                    Set GetEditItmX = itmX
                    Exit Function
                End If
            End If
        Next
    ElseIf psReportFormat <> vbNullString Then
        sLookForReportFormat = psReportFormat
        sLookForReportFormat = Mid(sLookForReportFormat, 200)
        sLookForReportFormat = Trim(sLookForReportFormat)
        'Looking for Multi report ite m
        For Each itmX In lvwPackageItem.ListItems
            sReportFormat = itmX.ListSubItems(GuiPackageItemListView.ReportFormat - 1)
            sReportFormat = Mid(sReportFormat, InStr(1, sReportFormat, String(200, " "), vbBinaryCompare))
            sReportFormat = Trim(sReportFormat)
            If StrComp(sReportFormat, sLookForReportFormat, vbTextCompare) = 0 Then
                Set GetEditItmX = itmX
                Exit Function
            End If
        Next
    End If
    
    Set GetEditItmX = Nothing
    
    'cleanup
    Set itmX = Nothing
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function GetEditItmX"
End Function

Public Function GetNextSortOrder() As String
    On Error GoTo EH
    Dim itmX As ListItem
    Dim lNextOrder As Long
    Dim sFlagText As String
    
    
    If lvwPackageItem.ListItems.Count = 0 Then
        GetNextSortOrder = "1"
    Else
        lvwPackageItem.SortKey = GuiPackageItemListView.SortOrder
        lvwPackageItem.Sorted = True
        'need to loop through all of the Items and select the last one that is not
        'Deleted.  All the deleted should be sorted to the bottom of the list.
        For Each itmX In lvwPackageItem.ListItems
            sFlagText = itmX.SubItems(GuiPackageItemListView.IsDeleted - 1)
            If Not goUtil.GetFlagFromText(sFlagText) Then
                lNextOrder = CLng(itmX.Text)
            Else
                Exit For
            End If
        Next
        lNextOrder = lNextOrder + 1
        GetNextSortOrder = CStr(lNextOrder)
    End If
    
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function GetNextSortOrder"
End Function


Private Sub cmdAddMultiReport_Click()
    On Error GoTo EH
    Dim MyfrmAddMultiReportItem As AddMultiReportItem
    
    Set MyfrmAddMultiReportItem = New AddMultiReportItem
    
    With MyfrmAddMultiReportItem
        .MyfrmClaim = mfrmClaim
        .AssignmentsID = msAssignmentsID
        .TableName = "Package"
    End With
    
    
    Load MyfrmAddMultiReportItem
    
    MyfrmAddMultiReportItem.Timer_SaveMe.Enabled = True
    MyfrmAddMultiReportItem.Show vbModal
    
    MyfrmAddMultiReportItem.CLEANUP
    
    Unload MyfrmAddMultiReportItem
    
    Set MyfrmAddMultiReportItem = Nothing
    
    If cboPackage.ListCount > 1 Then
        cmdAddMultiReport.Visible = False
    End If
    
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdAddMultiReport_Click"
End Sub

Private Sub cmdDelItem_Click()
    On Error GoTo EH
    Dim sMess As String
    Dim sFilePath As String
    Dim sFileName As String
    
    If lvwSavedPackages.SelectedItem Is Nothing Then
        Exit Sub
    End If
    
    cmdDelItem.Enabled = False
    
    sFileName = lvwSavedPackages.SelectedItem.Text
    sFilePath = goUtil.AttachReposPath & sFileName
    
    sMess = "Are you sure you want to Delete the Selected File " & sFileName & " ? "
    
    If MsgBox(sMess, vbQuestion + vbYesNo, "Delete Selected File") = vbYes Then
        sMess = goUtil.utDeleteFile(sFilePath)
        If sMess <> vbNullString Then
            sMess = "Error " & sMess
            MsgBox sMess, vbCritical + vbOKOnly, "Error"
        End If
    End If
    
    PopulatelvwSavedPackages
    
    cmdDelItem.Enabled = True
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdDelItem_Click"
End Sub

Private Sub cmdDelPackageItem_Click()
    On Error GoTo EH
    Dim itmX As ListItem
    Dim iMyIcon As Long
    Dim sFlagText As String
    Dim bFlag As Boolean
    Dim sMess As String
    Dim sSentDate As String
    
    
    If lvwPackageItem.ListItems.Count = 0 Then
        Exit Sub
    End If
    
    sMess = "Are you sure you want to Remove the Selected item(s)?"
    
    If MsgBox(sMess, vbQuestion + vbYesNo, "Remove Item(s)") = vbNo Then
        Exit Sub
    End If
    
    sMess = vbNullString
    For Each itmX In lvwPackageItem.ListItems
        'Only Flag items that are selected
        If Not itmX.Selected Then
            GoTo NEXT_ITEM
        End If
        'First need to check some things
        'Check for Client co Approve
        sFlagText = itmX.ListSubItems(GuiPackageItemListView.IsClientCoApprove - 1).Text
        bFlag = goUtil.GetFlagFromText(sFlagText)
        If bFlag Then
            sMess = sMess & "Can't delete, client already approved Item: " & itmX.Text & " Name: " & itmX.ListSubItems(GuiPackageItemListView.piName - 1).Text
            sMess = sMess & " Desc: " & itmX.ListSubItems(GuiPackageItemListView.Description - 1).Text & vbCrLf
            GoTo NEXT_ITEM
        End If
        
        'Check for Client CO Delete
        sFlagText = itmX.ListSubItems(GuiPackageItemListView.IsClientCoDelete - 1).Text
        bFlag = goUtil.GetFlagFromText(sFlagText)
        If bFlag Then
            sMess = sMess & "Can't delete, client deleted Item: " & itmX.Text & " Name: " & itmX.ListSubItems(GuiPackageItemListView.piName - 1).Text
            sMess = sMess & " Desc: " & itmX.ListSubItems(GuiPackageItemListView.Description - 1).Text & vbCrLf
            GoTo NEXT_ITEM
        End If
        
        'Check for Deleted
        sFlagText = itmX.ListSubItems(GuiPackageItemListView.IsDeleted - 1).Text
        bFlag = goUtil.GetFlagFromText(sFlagText)
        If bFlag Then
            sMess = sMess & "Can't delete, deleted Item: " & itmX.Text & " Name: " & itmX.ListSubItems(GuiPackageItemListView.piName - 1).Text
            sMess = sMess & " Desc: " & itmX.ListSubItems(GuiPackageItemListView.Description - 1).Text & vbCrLf
            GoTo NEXT_ITEM
        End If
        
        'Check for Sent Date
        sSentDate = itmX.ListSubItems(GuiPackageItemListView.SentDate - 1).Text
        If IsDate(sSentDate) Then
            'Can't delete something that was already sent
            sMess = sMess & "Can't delete already sent item: " & itmX.Text & " Name: " & itmX.ListSubItems(GuiPackageItemListView.piName - 1).Text
            sMess = sMess & " Desc: " & itmX.ListSubItems(GuiPackageItemListView.Description - 1).Text & vbCrLf
            GoTo NEXT_ITEM
        End If
        
        iMyIcon = GuiPackageItemStatusList.IsDeleted
        sFlagText = goUtil.GetFlagText(True)
        itmX.SubItems(GuiPackageItemListView.IsDeleted - 1) = sFlagText
        itmX.ListSubItems(GuiPackageItemListView.IsDeleted - 1).ReportIcon = iMyIcon
        iMyIcon = Empty
        sFlagText = goUtil.GetFlagText(False)
        itmX.ListSubItems(GuiPackageItemListView.SendMe - 1).Text = sFlagText
        itmX.ListSubItems(GuiPackageItemListView.SendMe - 1).ReportIcon = iMyIcon
        iMyIcon = GuiPackageItemStatusList.UpLoadMe
        itmX.SubItems(GuiPackageItemListView.UpLoadMe - 1) = sFlagText
        itmX.ListSubItems(GuiPackageItemListView.UpLoadMe - 1).ReportIcon = iMyIcon
        itmX.SubItems(GuiPackageItemListView.DateLastUpdated - 1) = Format(Now(), "MM/DD/YYYY HH:MM:SS")
        itmX.SubItems(GuiPackageItemListView.UpdateByUserID - 1) = goUtil.gsCurUsersID
NEXT_ITEM:
    Next
    
    If sMess <> vbNullString Then
        MsgBox sMess, vbExclamation + vbOKOnly, "Could not mark some item(s) to be deleted."
    End If
    
    'Resave the sort order
    SaveSortOrder
    lvwPackageItem.SortKey = GuiPackageItemListView.SortOrder
    lvwPackageItem.Sorted = True
    LoadReports
    RemoveAddedReports
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdDelPackageItem_Click"
End Sub

Private Sub cmdDown_Click()
    On Error GoTo EH
    goUtil.utMoveListItem lvwPackageItem, MoveDown
    If Not lvwPackageItem.SelectedItem Is Nothing Then
        mbSaveSort = True
    End If
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdDown_Click"
End Sub


Private Sub cmdExit_Click()
    On Error GoTo EH
    Dim sMess As String
    
'    If cmdSave.Enabled Then
'        sMess = "Do you want to Save Changes?" & vbCrLf & vbCrLf & Me.Caption
'        If MsgBox(sMess, vbQuestion + vbYesNo, "Save Changes") = vbNo Then
'            cmdSave.Enabled = False
'        End If
'    End If
    
    mbUnloadMe = True
    Me.Visible = False
    mfrmClaim.Timer_UnloadForm.Enabled = True
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdExit_Click"
End Sub

Private Sub cmdFind_Click()
    On Error GoTo EH
    If lvwPackageItem.ListItems.Count > 0 Then
        mlLastFindIndex = 0
        goUtil.utFindListItem Me, lvwPackageItem, msFindText, mlLastFindIndex
    End If
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdFind_Click"
End Sub

Private Sub cmdFindNext_Click()
    On Error GoTo EH
    
    If msFindText <> vbNullString And lvwPackageItem.ListItems.Count > 0 Then
        goUtil.utFindListItem Me, lvwPackageItem, msFindText, mlLastFindIndex
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdFindNext_Click"
End Sub

Private Sub cmdMailTo_Click()
    On Error GoTo EH
    Dim sPath As String
    Dim sSelFile As String
    Dim lFileSize As Long
    Dim sMyFilter As String
    Dim sFile As String
    Dim itmX As ListItem
    Dim lRet As Long
    Dim sFileName As String
    Dim sMapiLaunchPath As String
    
    
    
    If lvwSavedPackages.SelectedItem Is Nothing Then
        Exit Sub
    End If

    Set itmX = lvwSavedPackages.SelectedItem
    
    sFileName = itmX.Text

    sPath = goUtil.AttachReposPath

    sPath = goUtil.AttachReposPath & sFileName
   
    sMapiLaunchPath = goUtil.gsInstallDir & "\" & "SendMail.exe"
    
    
    If goUtil.utFileExists(sPath) Then
        sMapiLaunchPath = """" & sMapiLaunchPath & """  """ & sPath & """"
        Shell sMapiLaunchPath, vbMaximizedFocus
    End If

    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdMailTo_Click"
End Sub


Private Sub cmdPrint_Click()
    On Error GoTo EH
    
    PrintPackageItems False
    
    Exit Sub
EH:
   goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdPrint_Click"
End Sub

Private Function PrintPackageItems(Optional pbPrintToFile As Boolean, Optional psBuilDir As String) As Boolean
    On Error GoTo EH
    Dim itmX As ListItem
    Dim sCopy As String
    Dim sReportFormat As String
    Dim sSaveToFileName As String
    Dim bPrintPackageitems As Boolean
    Dim lSelItemCount As Long
    Dim oProg As V2ECKeyBoard.clsProgForm
    Dim sData As String
    Dim saryData() As String
    Dim sReportTitle As String
    Dim sPDFFilePath As String
    Dim sLRFormat As String
    Dim bIsAttachment As Boolean
    Dim bIsLossReport As Boolean
    Dim sMess As String
    
    
    Screen.MousePointer = MousePointerConstants.vbHourglass
    
    For Each itmX In lvwPackageItem.ListItems
        If itmX.Selected Then
            'When Printing to File and Export to XML only is flagged,
            'DO NOT ALLOW process of Loss Report Or Attachments are included in
            'the Process!!
            If pbPrintToFile And mbXMLExport And mbXMLOnly Then
                sData = itmX.ListSubItems(GuiPackageItemListView.ReportFormat - 1)
                sReportTitle = Trim(left(sData, 200))
                goUtil.utCleanFileFolderName sReportTitle, False
                sData = Mid(sData, InStr(1, sData, String(200, " "), vbBinaryCompare))
                sData = Trim(sData)
                saryData() = Split(sData, "|", , vbBinaryCompare)
                If UBound(saryData, 1) <= 1 Then
                    'Check for Loss Report
                    sLRFormat = saryData(0)
                    If StrComp(sLRFormat, "LRFormat", vbTextCompare) = 0 Then
                        bIsLossReport = True
                    Else
                        sPDFFilePath = saryData(0)
                        If InStr(1, sPDFFilePath, ".pdf", vbTextCompare) > 0 Then
                            bIsAttachment = True
                        End If
                    End If
                End If
                If bIsAttachment Then
                    sMess = sMess & "Attachment: " & sReportTitle & vbCrLf
                    bIsAttachment = False
                    itmX.Selected = False
                    GoTo SKIP_ITEMX
                ElseIf bIsLossReport Then
                    sMess = sMess & "Loss Report: " & sReportTitle & vbCrLf
                    bIsLossReport = False
                    itmX.Selected = False
                    GoTo SKIP_ITEMX
                End If
            End If
            lSelItemCount = lSelItemCount + 1
        End If
SKIP_ITEMX:
    Next
    
    If sMess <> vbNullString Then
        sMess = "Loss Reports and Attachments can not be part of an XML ONLY Export!" & vbCrLf & vbCrLf & sMess
        MsgBox sMess, vbExclamation
    End If
    
    
    If lSelItemCount = 0 Then
        Screen.MousePointer = MousePointerConstants.vbDefault
        Exit Function
    End If
    
    Set oProg = New V2ECKeyBoard.clsProgForm
    oProg.LoadForm
    oProg.cmdCancelEnable = False
    If pbPrintToFile Then
        oProg.Caption = "Print To File Progress"
    Else
        oProg.Caption = "Print Progress"
    End If
    oProg.framFileText = "Printing File(s)"
    oProg.framRecordText = vbNullString
    oProg.framTableText = vbNullString
    oProg.PBarFile.Max = lSelItemCount
    oProg.ShowForm True
    
    lSelItemCount = 0
    For Each itmX In lvwPackageItem.ListItems
        
        If itmX.Selected Then
            lSelItemCount = lSelItemCount + 1
            oProg.lblFileText = oProg.lblFileText & " - " & itmX.ListSubItems(GuiPackageItemListView.Description - 1)
            oProg.lblFileText = itmX.ListSubItems(GuiPackageItemListView.piName - 1)
            oProg.PBarFile.Value = lSelItemCount
            oProg.SetFocus
            oProg.Refresh
            Sleep 100
            sReportFormat = itmX.ListSubItems(GuiPackageItemListView.ReportFormat - 1)
            'Only allow the IB at this time to pass in Copy parameter
            If InStr(1, sReportFormat, "ECrpt" & goUtil.gsCurCarDBName & "_arRptPhotos", vbTextCompare) > 0 Then
                sCopy = vbNullString
            ElseIf InStr(1, sReportFormat, "ECrpt" & goUtil.gsCurCarDBName & "_arWorkSheetDiag", vbTextCompare) > 0 Then
                sCopy = vbNullString
            ElseIf InStr(1, sReportFormat, "ECrpt" & goUtil.gsCurCarDBName & "_arRptAddlChk", vbTextCompare) > 0 Then
                sCopy = vbNullString
            ElseIf InStr(1, sReportFormat, "ECrpt" & goUtil.gsCurCarDBName & "_arRptIB" & goUtil.gsCurCarDBName, vbTextCompare) > 0 Then
                sCopy = cboCopy.Text
            Else
                sCopy = vbNullString
            End If
            
            cmdPrint.Enabled = False
            If pbPrintToFile Then
                sSaveToFileName = Format(itmX.Text, "000")
                sSaveToFileName = sSaveToFileName & "_" & itmX.ListSubItems(GuiPackageItemListView.piName - 1)
                sSaveToFileName = sSaveToFileName & "_" & itmX.ListSubItems(GuiPackageItemListView.Description - 1)
                'Max Len of File name is 40 Chars
                sSaveToFileName = left(sSaveToFileName, 50)
                goUtil.utCleanFileFolderName sSaveToFileName, False
                sSaveToFileName = sSaveToFileName & ".pdf"
                bPrintPackageitems = mfrmClaim.PrintActiveReport(itmX, , sCopy, mbPrintPreview, psBuilDir, sSaveToFileName, mbXMLExport, mbXMLOnly)
            Else
                bPrintPackageitems = mfrmClaim.PrintActiveReport(itmX, , sCopy, mbPrintPreview)
            End If
            If Not bPrintPackageitems Then
                PrintPackageItems = bPrintPackageitems
                
                GoTo CLEAN_UP
            End If
        End If
    Next
    
    If Not mbUnloadMe Then
        cmdPrint.Enabled = True
    End If
    
    Screen.MousePointer = MousePointerConstants.vbDefault
    
    PrintPackageItems = bPrintPackageitems
    
CLEAN_UP:
    If Not oProg Is Nothing Then
        oProg.CLEANUP
        Set oProg = Nothing
    End If
    
    Exit Function
EH:
    Screen.MousePointer = MousePointerConstants.vbDefault
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Function PrintPackageItems"
End Function




Private Sub cmdSave_Click()
    On Error GoTo EH
    Dim sMess As String
    
    
    cmdSave.Enabled = False
    Screen.MousePointer = MousePointerConstants.vbArrowHourglass
    
    If SaveMe() Then
        cboPackage.ListIndex = 0
    End If
    Screen.MousePointer = MousePointerConstants.vbDefault
    Exit Sub
EH:
    Screen.MousePointer = MousePointerConstants.vbDefault
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdSave_Click"
End Sub

Private Sub cmdSavePackage_Click()
    On Error GoTo EH
    Dim sMess As String
    Dim bUsePassword As Boolean
    Dim sBuildDir As String
    Dim itmX As ListItem
    Dim sFlagText As String
    
    cmdSavePackage.Enabled = False
    'If password selected, must validate it
    If chkUsePass.Value = vbChecked Then
        If StrComp(txtPassWord(0).Text, txtPassWord(1).Text, vbBinaryCompare) <> 0 Then
            sMess = "Password not confirmed!"
        ElseIf txtPassWord(0).Text = vbNullString Then
            sMess = "Password can not be blank!"
        End If
        If sMess <> vbNullString Then
            MsgBox sMess, vbExclamation + vbOKOnly, "Invalid Password Entry"
            cmdSavePackage.Enabled = True
            Exit Sub
        End If
        bUsePassword = True
    End If
    
    sMess = "Do you want save all items in this package?" & vbCrLf & vbCrLf
    sMess = sMess & "Click ""Yes"" to select all items." & vbCrLf
    sMess = sMess & "Click ""No"" save the current selection."
    
    If MsgBox(sMess, vbQuestion + vbYesNo, "Select all items for save") = vbYes Then
        For Each itmX In lvwPackageItem.ListItems
            'Do not select Deleted items
            sFlagText = itmX.ListSubItems(GuiPackageItemListView.IsDeleted - 1)
            If Not goUtil.GetFlagFromText(sFlagText) Then
                itmX.Selected = True
            Else
                itmX.Selected = False
            End If
        Next
    End If
    
    'Need to create Build Folder to store all the print to file .pdf docs
    'this folder will be used to create the Zip File and then move to
    'the Attach repos folder.
    sBuildDir = goUtil.gsInstallDir & "\BuildSave\"
    If Not goUtil.utFileExists(sBuildDir, True) Then
        goUtil.utMakeDir sBuildDir
    Else
        'Need to be sure nothing exisits in the build dir
        goUtil.utDeleteFile sBuildDir & "*.*"
        Sleep 100
    End If
    
    'Save this Item to File
    If SaveThisPackageToFile(bUsePassword, sBuildDir) Then
        PopulatelvwSavedPackages
        chkUsePass.Value = vbUnchecked
        MsgBox "Save Succeeded!", vbInformation + vbOKOnly, "Success!"
    End If
    
    cmdSavePackage.Enabled = True
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdSavePackage_Click"
End Sub

Public Function SaveThisPackageToFile(pbUsePassword As Boolean, psBuildDir As String) As Boolean
    On Error GoTo EH
    Dim RS As ADODB.Recordset
    Dim oXZip As V2ECKeyBoard.clsXZip
    Dim sSuffixName As String
    Dim sMess As String
    Dim sIbnumber As String
    Dim sZipName As String
    Dim sPassWord As String
    Dim sEncryptPassWord As String
    Dim sDestDir As String
    
    'Get Filename from user
    sMess = "Please enter a file name for this item."
    
    sSuffixName = InputBox(sMess, "Enter File Name")
    
    sMess = vbNullString
    If Len(sSuffixName) > 30 Then
        sMess = "Name too Big!"
        sSuffixName = vbNullString
    End If
    goUtil.utCleanFileFolderName sSuffixName, False
    
    If Trim(sSuffixName) = vbNullString Then
        MsgBox "File not Saved!" & vbCrLf & vbCrLf & sMess, vbExclamation + vbOKOnly, "Save Aborted"
        GoTo CLEAN_UP
    End If
    
    'Create Zipname
    mfrmClaim.SetadoRSAssignments msAssignmentsID
    Set RS = mfrmClaim.adoRSAssignments
    sIbnumber = goUtil.IsNullIsVbNullString(RS.Fields("IBNUM"))
    sZipName = sIbnumber & "_" & sSuffixName & ".zip"
    
    'The ultimate destination for this file will be
    'the attach repos
    sDestDir = goUtil.AttachReposPath
    
    'Check to see if it already Exists
    If goUtil.utFileExists(sDestDir & sZipName, False) Then
        sMess = "The file """ & sZipName & """ already exists!" & vbCrLf & vbCrLf
        sMess = sMess & "Click ""Yes"" to update this existing file." & vbCrLf
        sMess = sMess & "(Replace exisiting documents and Append new documents)" & vbCrLf & vbCrLf
        sMess = sMess & "Click ""No"" to Abort!"
        If MsgBox(sMess, vbExclamation + vbYesNo, "File Already Exists") = vbNo Then
            GoTo CLEAN_UP
        End If
    End If
    
    If Not PrintPackageItems(True, psBuildDir) Then
        sMess = "Package items NOT saved!" & vbCrLf & vbCrLf
        MsgBox sMess, vbExclamation + vbOKOnly, "Problems saving to file"
        GoTo CLEAN_UP
    End If
    'Need to save these created files into 1 zip file
    
    'Create the Zip utility
    Set oXZip = New V2ECKeyBoard.clsXZip
    oXZip.SetUtilObject goUtil
    
    If pbUsePassword Then
        sPassWord = txtPassWord(0).Text
        sEncryptPassWord = goUtil.Encode(sPassWord)
    End If
    If Not oXZip.SaveZIPFiles(psBuildDir, sZipName, "*.*", sEncryptPassWord, sDestDir) Then
        GoTo CLEAN_UP
    End If
    
    SaveThisPackageToFile = True
    
CLEAN_UP:
    Screen.MousePointer = vbDefault
    'cleanup
    Set RS = Nothing
    Set oXZip = Nothing
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function SaveThisPackageToFile"
End Function

Private Sub SaveSortOrder()
    On Error GoTo EH
    Dim itmX As ListItem
    Dim lSort As Long
    Dim sFlagText As String
    Dim iMyIcon As Long
    Dim sSort As String
    
    For Each itmX In lvwPackageItem.ListItems
        'Only renumber items that are not Deleted
        sFlagText = itmX.SubItems(GuiPackageItemListView.IsDeleted - 1)
        If Not goUtil.GetFlagFromText(sFlagText) Then
            lSort = lSort + 1
            'Only update this item if it is not already in the right sort order
            sSort = itmX.Text
            If StrComp(sSort, CStr(lSort), vbTextCompare) <> 0 Then
                itmX.Text = lSort
                itmX.SubItems(GuiPackageItemListView.SortOrderSort - 1) = goUtil.utNumInTextSortFormat(itmX.Text)
                'UpLoadMe
                iMyIcon = GuiPackageItemStatusList.UpLoadMe
                sFlagText = goUtil.GetFlagText(True)
                itmX.SubItems(GuiPackageItemListView.UpLoadMe - 1) = sFlagText
                itmX.ListSubItems(GuiPackageItemListView.UpLoadMe - 1).ReportIcon = iMyIcon
                itmX.SubItems(GuiPackageItemListView.DateLastUpdated - 1) = Format(Now(), "MM/DD/YYYY HH:MM:SS")
                itmX.SubItems(GuiPackageItemListView.UpdateByUserID - 1) = goUtil.gsCurUsersID
            End If
        Else
            sSort = itmX.Text
            'Only change a deleted Item if it wasn't already marked as Deleted
            If StrComp(sSort, "DEL", vbTextCompare) <> 0 Then
                itmX.Text = "DEL"
                itmX.SubItems(GuiPackageItemListView.SortOrderSort - 1) = itmX.Text
                'UpLoadMe
                iMyIcon = GuiPackageItemStatusList.UpLoadMe
                sFlagText = goUtil.GetFlagText(True)
                itmX.SubItems(GuiPackageItemListView.UpLoadMe - 1) = sFlagText
                itmX.ListSubItems(GuiPackageItemListView.UpLoadMe - 1).ReportIcon = iMyIcon
                itmX.SubItems(GuiPackageItemListView.DateLastUpdated - 1) = Format(Now(), "MM/DD/YYYY HH:MM:SS")
                itmX.SubItems(GuiPackageItemListView.UpdateByUserID - 1) = goUtil.gsCurUsersID
            End If
        End If
    Next
    
    mbSaveSort = False
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub SaveSortOrder"
End Sub

Private Sub cmdSaveSort_Click()
    SaveSortOrder
End Sub

Private Sub cmdSelAll_Click()
    On Error GoTo EH
    Dim itmX As ListItem
    For Each itmX In lvwPackageItem.ListItems
        itmX.Selected = True
    Next
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdSelAll_Click"
End Sub


Private Sub cmdPrintList_Click()
    On Error GoTo EH
    goUtil.utPrintListView App.EXEName, lvwPackageItem, "Package Items"
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdPrintList_Click"
End Sub

Private Sub cmdSendMe_Click()
    On Error GoTo EH
    
    cmdSendMe.Enabled = False
    FlagItemsSendMe True, lvwPackageItem
    cmdSendMe.Enabled = True
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdSendMe_Click"
End Sub

Public Function FlagItemsSendMe(pbFlagSendMe As Boolean, poLvw As ListView) As Boolean
    On Error GoTo EH
    Dim iMyIcon As Long
    Dim sFlagText As String
    Dim bFlag As Boolean
    Dim sTemp As String
    Dim sSentDate As String
    Dim itmX As ListItem
    Dim oLVW As ListView
    Dim sMess As String


    
    Set oLVW = poLvw
    
    If oLVW.SelectedItem Is Nothing Then
        sMess = "Nothing Selected!"
        GoTo SKIP_DOWN_HERE
    End If
    
    For Each itmX In oLVW.ListItems
        If Not itmX.Selected Then
            GoTo NEXT_ITEM
        End If
        'SendMe
        sFlagText = itmX.ListSubItems(GuiPackageItemListView.SendMe - 1).Text
        bFlag = goUtil.GetFlagFromText(sFlagText)
        If bFlag Then
            'Already Flagged to Send
            If pbFlagSendMe Then
                GoTo NEXT_ITEM
            Else
                iMyIcon = Empty
                sFlagText = goUtil.GetFlagText(False)
            End If
        Else
            If pbFlagSendMe Then
                'Need to be sure not sending something that has been Deleted
                'By Client Or Accepted By Client.  The Only way a user can
                'Flag to Send an Item is if the Item Has been Rejected by Client
                'If it has been Deleted by the Client, Or Accepted... CANT SEND IT AGAIN!!!
                
                'Check for Client co Approve
                sFlagText = itmX.ListSubItems(GuiPackageItemListView.IsClientCoApprove - 1).Text
                bFlag = goUtil.GetFlagFromText(sFlagText)
                If bFlag Then
                    sMess = sMess & "Can't send, client already approved Item: " & itmX.Text & " Name: " & itmX.ListSubItems(GuiPackageItemListView.piName - 1).Text
                    sMess = sMess & " Desc: " & itmX.ListSubItems(GuiPackageItemListView.Description - 1).Text & vbCrLf
                    GoTo NEXT_ITEM
                End If
                
                'Check for Client CO Delete
                sFlagText = itmX.ListSubItems(GuiPackageItemListView.IsClientCoDelete - 1).Text
                bFlag = goUtil.GetFlagFromText(sFlagText)
                If bFlag Then
                    sMess = sMess & "Can't send, client deleted Item: " & itmX.Text & " Name: " & itmX.ListSubItems(GuiPackageItemListView.piName - 1).Text
                    sMess = sMess & " Desc: " & itmX.ListSubItems(GuiPackageItemListView.Description - 1).Text & vbCrLf
                    GoTo NEXT_ITEM
                End If
                
                'Check for Deleted
                sFlagText = itmX.ListSubItems(GuiPackageItemListView.IsDeleted - 1).Text
                bFlag = goUtil.GetFlagFromText(sFlagText)
                If bFlag Then
                    sMess = sMess & "Can't send, deleted Item: " & itmX.Text & " Name: " & itmX.ListSubItems(GuiPackageItemListView.piName - 1).Text
                    sMess = sMess & " Desc: " & itmX.ListSubItems(GuiPackageItemListView.Description - 1).Text & vbCrLf
                    GoTo NEXT_ITEM
                End If
                
                'Check for Sent Date
                sSentDate = itmX.ListSubItems(GuiPackageItemListView.SentDate - 1).Text
                If IsDate(sSentDate) Then
                    'If the sent date is set the only way this item can be sent again is if it was
                    'rejected by the client to make corrections
                    sFlagText = itmX.ListSubItems(GuiPackageItemListView.IsClientCoReject - 1).Text
                    bFlag = goUtil.GetFlagFromText(sFlagText)
                    If Not bFlag Then
                        sMess = sMess & "Can't send already sent item, unless client rejects Item: " & itmX.Text & " Name: " & itmX.ListSubItems(GuiPackageItemListView.piName - 1).Text
                        sMess = sMess & " Desc: " & itmX.ListSubItems(GuiPackageItemListView.Description - 1).Text & vbCrLf
                        GoTo NEXT_ITEM
                    End If
                End If
                
                iMyIcon = GuiPackageItemStatusList.IsChecked
                sFlagText = goUtil.GetFlagText(True)
             Else
                GoTo NEXT_ITEM
             End If
        End If
        itmX.ListSubItems(GuiPackageItemListView.SendMe - 1).Text = sFlagText
        itmX.ListSubItems(GuiPackageItemListView.SendMe - 1).ReportIcon = iMyIcon
        itmX.ListSubItems(GuiPackageItemListView.DateLastUpdated - 1).Text = Format(Now(), "MM/DD/YYYY HH:MM:SS")
        sFlagText = goUtil.GetFlagText(True)
        itmX.ListSubItems(GuiPackageItemListView.UpLoadMe - 1).Text = sFlagText
        itmX.ListSubItems(GuiPackageItemListView.UpLoadMe - 1).ReportIcon = GuiPackageItemStatusList.UpLoadMe
        'Need to unflag the clientco Rejected after flagging to send a new
        'document.  Leave the Date And time there for history sake
        sFlagText = goUtil.GetFlagText(False)
        iMyIcon = Empty
        itmX.ListSubItems(GuiPackageItemListView.IsClientCoReject - 1).Text = sFlagText
        itmX.ListSubItems(GuiPackageItemListView.IsClientCoReject - 1).ReportIcon = iMyIcon
NEXT_ITEM:
    Next
    
SKIP_DOWN_HERE:
    If sMess <> vbNullString Then
        MsgBox sMess, vbExclamation + vbOKOnly, "Could not mark some item(s) to be sent."
    End If
    
    FlagItemsSendMe = True
    
CLEAN_UP:
    Set itmX = Nothing
    Set oLVW = Nothing
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function FlagItemsSendMe"
End Function

Private Sub cmdUp_Click()
    On Error GoTo EH
    goUtil.utMoveListItem lvwPackageItem, MoveUp
    If Not lvwPackageItem.SelectedItem Is Nothing Then
        mbSaveSort = True
    End If
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdUp_Click"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo EH
    
    Select Case KeyCode
        Case KeyCodeConstants.vbKeyF11 'maximize window / Normalize window
            If Me.WindowState = VBRUN.FormWindowStateConstants.vbMaximized Then
                Me.WindowState = VBRUN.FormWindowStateConstants.vbNormal
            Else
                Me.WindowState = VBRUN.FormWindowStateConstants.vbMaximized
            End If
    End Select
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_KeyDown"
End Sub

Private Sub Form_Load()
    On Error GoTo EH
    Dim sLossFormat As String
    
    sLossFormat = goUtil.IsNullIsVbNullString(mfrmClaim.adoRSAssignments.Fields("LRFormat"))
    chkMaxPackageMaint.Visible = True
    If StrComp(sLossFormat, "V2ECcarFarmers.clsLossXML01", vbTextCompare) = 0 Then
        cmdPrintList.Visible = True
        cmdSendMe.Visible = True
    End If
    
    mbLoading = True
    'Set the Form Icon
    Me.Icon = mfrmClaim.optClaim(GuiClaimOptions.opt09_Print).Picture
    
    mbPrintPreview = CBool(GetSetting(App.EXEName, "GENERAL", "PRINT_PREVIEW", True))
    If mbPrintPreview Then
        chkPrintPreview.Value = vbChecked
    Else
        chkPrintPreview.Value = vbUnchecked
    End If
    
    mbXMLExport = CBool(GetSetting(App.EXEName, "GENERAL", "XML_EXPORT", False))
    If mbXMLExport Then
        chkXMLExport.Value = vbChecked
    Else
        chkXMLExport.Value = vbUnchecked
    End If
    
    mbXMLOnly = CBool(GetSetting(App.EXEName, "GENERAL", "XML_ONLY", False))
    If mbXMLOnly Then
        chkXMLOnly.Value = vbChecked
    Else
        chkXMLOnly.Value = vbUnchecked
    End If
    
    LoadCopy
    LoadHeaderlvwSavedPackages
    LoadHeaderlvwPackageItem
    
    LoadMe
    
    ShowFrame
    
    cmdSave.Enabled = False
    
    'Add A Package if its visible
    If cboPackage.ListCount = 1 Then
        cmdAddMultiReport_Click
    End If
    
    'Select the package
    If cboPackage.ListCount > 1 Then
        cboPackage.ListIndex = 1
    End If
    
    'Add all available items
    'note if an item was previously deleted / Removed from the
    'package it will not be added to the package by the
    'AddReportItem function while mbLoading = true
    chkAddAllAvailRptDoc.Value = vbChecked
    cmdAddAvailRptDoc_Click 1
    chkAddAllAvailRptDoc.Value = vbUnchecked
    LoadMe
    'Select the package
    If cboPackage.ListCount > 1 Then
        cboPackage.ListIndex = 1
    End If
    Timer_MaximizeMe.Enabled = True
    mbLoading = False
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_Load"
End Sub

Public Function LoadMe() As Boolean
    On Error GoTo EH
    
    mbLoadingMe = True
    
    RefreshPackages
    PopulatelvwSavedPackages
    
    LoadMe = True
    mbLoadingMe = False
    Exit Function
EH:
    mbLoadingMe = False
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function LoadMe"
End Function

Public Function SaveMe() As Boolean
    On Error GoTo EH
    Dim itmX As ListItem
    Dim MypackageItem As GuiPackageItem
    Dim sFlagText As String
    Dim sMess As String
    Dim lSortOrder As Long
    Dim lThisItemSortOrder As Long
    Dim sSortOrder As String
    
    If lvwPackageItem.ListItems.Count = 0 Then
        SaveMe = False
        Exit Function
    End If
    
    'Save the sort order if it has changed
SAVE_SORT:
    If mbSaveSort Then
        SaveSortOrder
    Else
        lSortOrder = 0
        'Check to see if the current sort order is valid...
        For Each itmX In lvwPackageItem.ListItems
            sSortOrder = itmX.Text
            If StrComp(sSortOrder, "DEL", vbTextCompare) = 0 Then
                sSortOrder = DELETED_SORTORDER
            Else
                sSortOrder = itmX.Text
            End If
            If IsNumeric(sSortOrder) Then
                If sSortOrder > 0 Then
                    lSortOrder = lSortOrder + 1
                    lThisItemSortOrder = CLng(sSortOrder)
                    If lSortOrder <> lThisItemSortOrder Then
                        mbSaveSort = True
                        GoTo SAVE_SORT
                    End If
                End If
            End If
            
        Next
    End If
    
    For Each itmX In lvwPackageItem.ListItems
        With MypackageItem
            .PackageItemID = itmX.SubItems(GuiPackageItemListView.PackageItemID - 1)
            .PackageID = itmX.SubItems(GuiPackageItemListView.PackageID - 1)
            .AssignmentsID = itmX.SubItems(GuiPackageItemListView.AssignmentsID - 1)
            .ID = itmX.SubItems(GuiPackageItemListView.ID - 1)
            .IDPackage = itmX.SubItems(GuiPackageItemListView.IDPackage - 1)
            .IDAssignments = itmX.SubItems(GuiPackageItemListView.IDAssignments - 1)
            .ReportFormat = itmX.SubItems(GuiPackageItemListView.ReportFormat - 1)
            .RTAttachmentsID = itmX.SubItems(GuiPackageItemListView.RTAttachmentsID - 1)
            .IDRTAttachments = itmX.SubItems(GuiPackageItemListView.IDRTAttachments - 1)
            .Number = itmX.SubItems(GuiPackageItemListView.Number - 1)
            .AttachmentName = itmX.SubItems(GuiPackageItemListView.AttachmentName - 1)
            If StrComp(itmX.Text, "DEL", vbTextCompare) = 0 Then
                .SortOrder = DELETED_SORTORDER
            Else
                .SortOrder = itmX.Text
            End If
            .Name = itmX.SubItems(GuiPackageItemListView.piName - 1)
            .Description = itmX.SubItems(GuiPackageItemListView.Description - 1)
            sFlagText = itmX.SubItems(GuiPackageItemListView.IsCoApprove - 1)
            .IsCoApprove = goUtil.GetFlagFromText(sFlagText)
            .CoApproveDate = itmX.SubItems(GuiPackageItemListView.CoApproveDate - 1)
            .CoApproveDesc = itmX.SubItems(GuiPackageItemListView.CoApproveDesc - 1)
            sFlagText = itmX.SubItems(GuiPackageItemListView.IsClientCoReject - 1)
            .IsClientCoReject = goUtil.GetFlagFromText(sFlagText)
            .ClientCoRejectDate = itmX.SubItems(GuiPackageItemListView.ClientCoRejectDate - 1)
            .ClientCoRejectDesc = itmX.SubItems(GuiPackageItemListView.ClientCoRejectDesc - 1)
            sFlagText = itmX.SubItems(GuiPackageItemListView.IsClientCoDelete - 1)
            .IsClientCoDelete = goUtil.GetFlagFromText(sFlagText)
            .ClientCoDeleteDate = itmX.SubItems(GuiPackageItemListView.ClientCoDeleteDate - 1)
            .ClientCoDeleteDesc = itmX.SubItems(GuiPackageItemListView.ClientCoDeleteDesc - 1)
            sFlagText = itmX.SubItems(GuiPackageItemListView.IsClientCoApprove - 1)
            .IsClientCoApprove = goUtil.GetFlagFromText(sFlagText)
            .ClientCoApproveDate = itmX.SubItems(GuiPackageItemListView.ClientCoApproveDate - 1)
            .ClientCoApproveDesc = itmX.SubItems(GuiPackageItemListView.ClientCoApproveDesc - 1)
            .PackageItemGUID = itmX.SubItems(GuiPackageItemListView.PackageItemGUID - 1)
            sFlagText = itmX.SubItems(GuiPackageItemListView.SendMe - 1)
            .SendMe = goUtil.GetFlagFromText(sFlagText)
            .SentDate = itmX.SubItems(GuiPackageItemListView.SentDate - 1)
            sFlagText = itmX.SubItems(GuiPackageItemListView.IsDeleted - 1)
            .IsDeleted = goUtil.GetFlagFromText(sFlagText)
            sFlagText = itmX.SubItems(GuiPackageItemListView.DownLoadMe - 1)
            .DownLoadMe = goUtil.GetFlagFromText(sFlagText)
            sFlagText = itmX.SubItems(GuiPackageItemListView.UpLoadMe - 1)
            .UpLoadMe = goUtil.GetFlagFromText(sFlagText)
            .AdminComments = itmX.SubItems(GuiPackageItemListView.AdminComments - 1)
            .DateLastUpdated = itmX.SubItems(GuiPackageItemListView.DateLastUpdated - 1)
            .UpdateByUserID = itmX.SubItems(GuiPackageItemListView.UpdateByUserID - 1)
        End With
        
        'check wether to Add or edit this Item
        If StrComp(MypackageItem.PackageItemID, "Null", vbTextCompare) = 0 Then
            If Not AddPackageItem(MypackageItem) Then
                sMess = "Could not Add " & MypackageItem.Name & " " & MypackageItem.Description
                MsgBox sMess, vbCritical + vbOKOnly, "Error Saving"
                cmdSave.Enabled = False
                Exit Function
            End If
        Else
            If Not EditPackageItem(MypackageItem) Then
                sMess = "Could not Edit " & MypackageItem.Name & " " & MypackageItem.Description
                MsgBox sMess, vbCritical + vbOKOnly, "Error Saving"
                cmdSave.Enabled = False
                Exit Function
            End If
        End If
    Next
    
    cmdSave.Enabled = False
    SaveMe = True
    mbSaveSort = False
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function SaveMe"
End Function


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo EH
    Dim sMess As String
    
    Select Case UnloadMode
        Case vbFormControlMenu
            sMess = "Are you sure you want to Exit Package Screen?"
'            If MsgBox(sMess, vbQuestion + vbOKCancel, "Exit Package Screen") = vbCancel Then
'                Cancel = True
'                Exit Sub
'            End If
'            If cmdSave.Enabled Then
'                sMess = "Do you want to Save Changes?" & vbCrLf & vbCrLf & Me.Caption
'                If MsgBox(sMess, vbQuestion + vbYesNo, "Save Changes") = vbNo Then
'                    cmdSave.Enabled = False
'                End If
'            End If
            Cancel = True
            mbUnloadMe = True
            Me.Visible = False
            mfrmClaim.Timer_UnloadForm.Enabled = True
        Case Else
            CLEANUP
    End Select
    
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_QueryUnload"
End Sub

Private Sub Form_Resize()
    ReSizeMe
End Sub

Public Sub ReSizeMe()
    On Error Resume Next
    Dim sNavScreenPos As String
    
    If Me.WindowState <> vbMinimized Then
        Me.Visible = False
    End If
    
    If Me.WindowState <> vbMinimized And Me.WindowState <> vbMaximized Then
        Me.Width = Screen.Width - 10
        Me.Height = Screen.Height - (10 + mfrmClaim.Height + goUtil.utGetTaskbarHeight)
        Me.top = mfrmClaim.top + mfrmClaim.Height
        Me.left = 10
    End If
    
    'Widths and lefts
    framPackage.Width = Me.Width - 285
    framEditPackage.Width = Me.Width - 5700
    cboPackage.Width = Me.Width - 9690
'    txtAdminComments.Width = Me.Width - 8740
    framPackageItemMaint.Width = Me.Width - 5685
    chkPrintPreview.left = Me.Width - 7410
    lblCopy.left = Me.Width - 8340
    cboCopy.left = Me.Width - 8340
    cmdPrint.left = Me.Width - 6450
'    cmdDelPackageItem.left = Me.Width - 9300
    lvwPackageItem.Width = Me.Width - 5925
    framCommands.Width = Me.Width - 285
    txtAdminComments.Width = Me.Width - 7965
    cmdSave.left = Me.Width - 2460
    cmdExit.left = Me.Width - 1380
    
    
    'Heights and tops
    framPackage.Height = Me.Height - 1785
    framAdjSavedPackages.Height = Me.Height - 5505
    lvwSavedPackages.Height = Me.Height - 6270
    framAvailRptDoc.top = Me.Height - 5280
    framPackageItemMaint.Height = Me.Height - 3585
    lvwPackageItem.Height = Me.Height - 4905
    
    framCommands.top = Me.Height - 1680
    
    'See if Maximize or minimize PackageItems
    If chkMaxPackageMaint.Enabled Then
        If chkMaxPackageMaint.Value = vbUnchecked Then
            framPackageItemMaint.top = framAdjSavedPackages.top - 100
            framPackageItemMaint.left = framAdjSavedPackages.left - 100
            framPackageItemMaint.Width = framPackage.Width - 240
            framPackageItemMaint.Height = framPackage.Height - 360
            lvwPackageItem.Width = framPackageItemMaint.Width - 240
            lvwPackageItem.Height = framPackageItemMaint.Height - 1320
            chkPrintPreview.left = framPackageItemMaint.Width - 1725
            lblCopy.left = framPackageItemMaint.Width - 2655
            cboCopy.left = framPackageItemMaint.Width - 2655
            cmdPrint.left = framPackageItemMaint.Width - 765
        Else
            framPackageItemMaint.top = framEditPackage.top + framEditPackage.Height - 80
            framPackageItemMaint.left = framEditPackage.left - 80
            framPackageItemMaint.Width = framEditPackage.Width
            framPackageItemMaint.Height = Me.Height - 3585
            lvwPackageItem.Width = framPackageItemMaint.Width - 240
            lvwPackageItem.Height = framPackageItemMaint.Height - 1320
            chkPrintPreview.left = framPackageItemMaint.Width - 1725
            lblCopy.left = framPackageItemMaint.Width - 2655
            cboCopy.left = framPackageItemMaint.Width - 2655
            cmdPrint.left = framPackageItemMaint.Width - 765
        End If
    End If
    
    If Me.WindowState <> vbMinimized Then
        Me.Visible = True
    End If
End Sub

Public Function CLEANUP() As Boolean
    On Error GoTo EH
   If cmdSave.Enabled Then
        SaveMe
        If Not mfrmClaim Is Nothing And Not mbUnloadMe Then
            mfrmClaim.RefreshMe
        End If
    End If
    Set mfrmClaim = Nothing
    Set MyGUI = Nothing
    
    
    CLEANUP = True
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function CLEANUP"
End Function

Private Sub lvwPackageItem_DblClick()
    On Error GoTo EH
    
    chkPrintPreview.Value = vbChecked
    
    PrintPackageItems False
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lvwPackageItem_DblClick"
End Sub

Private Sub lvwPackageItem_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo EH
    Select Case KeyCode
        Case KeyCodeConstants.vbKeyReturn, KeyCodeConstants.vbKeySpace
            chkPrintPreview.Value = vbChecked
            PrintPackageItems False
        Case KeyCodeConstants.vbKeyDelete
            cmdDelPackageItem_Click
    End Select
Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lvwPackageItem_KeyDown"
End Sub

Private Sub lvwSavedPackages_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    On Error GoTo EH
    Dim sMess As String
    If lvwSavedPackages.SortOrder = lvwAscending Then
        lvwSavedPackages.SortOrder = lvwDescending
        sMess = "Descending"
    Else
        lvwSavedPackages.SortOrder = lvwAscending
        sMess = "Ascending"
    End If
    
    'Set the Tool Tip
    lvwSavedPackages.ToolTipText = "Sort By " & ColumnHeader.Text & " " & sMess
    
    'Need to see if the Column clicked was a NON text sort
    'Like Date or number.  If a non textual column was clicked
    'need to use the next column as the sort since this next column
    'will be hidden and contain a sort friendly format.
    'Sort Key is Base 0 where CoulmnHeader is not
    Select Case ColumnHeader.Index
        Case SavedPackagesListView.DateCreated, SavedPackagesListView.DateLastUpdated
            lvwSavedPackages.SortKey = ColumnHeader.Index
        Case Else
            lvwSavedPackages.SortKey = ColumnHeader.Index - 1
    End Select
    
    lvwSavedPackages.Sorted = True
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lvwAvail_ColumnClick"
End Sub

Private Sub lvwSavedPackages_DblClick()
    cmdViewFile_Click
End Sub


Private Sub lvwSavedPackages_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case KeyCodeConstants.vbKeyReturn, KeyCodeConstants.vbKeySpace
            cmdViewFile_Click
        Case KeyCodeConstants.vbKeyDelete
            cmdDelItem_Click
    End Select
End Sub


'Private Sub txtAdminComments_GotFocus()
'    goUtil.utSelText txtAdminComments
'End Sub

Private Sub EnableEditFrames(pbEnabled As Boolean)
    On Error GoTo EH
 
    framPackage.Enabled = pbEnabled
    framEditPackage.Enabled = pbEnabled
    ShowFrame
    cmdSave.Enabled = pbEnabled
    chkMaxPackageMaint.Enabled = pbEnabled
    cmdPrintList.Enabled = pbEnabled
    cmdSendMe.Enabled = pbEnabled
    If Not pbEnabled And chkMaxPackageMaint.Value = vbChecked Then
        chkMaxPackageMaint.Value = vbUnchecked
    End If
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub EnableEditFrames"
End Sub

Public Function ShowFrame() As Boolean
    On Error GoTo EH
    Dim sFrameName As String
    Dim oFrame As Control
    Dim MyFrame As Frame
    Dim oControl As Control
    
    For Each oFrame In Me.Controls
        If TypeOf oFrame Is Frame Then
            Set MyFrame = oFrame
            For Each oControl In Me.Controls
                If Not TypeOf oControl Is ImageList And Not TypeOf oControl Is Timer Then
                    If oControl.Container.Name = MyFrame.Name Then
                        'Only allow enabling of chkXMLOnly if
                        'the Export to Xml is True
                        If StrComp(oControl.Name, chkXMLOnly.Name, vbTextCompare) = 0 Then
                            If Not mbXMLExport And MyFrame.Enabled Then
                                oControl.Enabled = False
                            Else
                                oControl.Enabled = MyFrame.Enabled
                            End If
                        Else
                            If StrComp(oControl.Name, chkMaxPackageMaint.Name, vbTextCompare) = 0 Then
                            ElseIf StrComp(oControl.Name, cmdPrintList.Name, vbTextCompare) = 0 Then
                            ElseIf StrComp(oControl.Name, cmdSendMe.Name, vbTextCompare) = 0 Then
                            Else
                                oControl.Enabled = MyFrame.Enabled
                            End If
                        End If
                        
'                            Debug.Print oControl.Name
                    End If
                End If
            Next
        End If
    Next
    
    ShowFrame = True
    Set oFrame = Nothing
    Set MyFrame = Nothing
    Set oControl = Nothing
        
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function ShowFrame"
End Function


Private Sub cmdViewFile_Click()
    On Error GoTo EH
    Dim sPath As String
    Dim sSelFile As String
    Dim lFileSize As Long
    Dim sMyFilter As String
    Dim sFile As String
    Dim itmX As ListItem
    Dim lRet As Long
    Dim sFileName As String
    
    
    If lvwSavedPackages.SelectedItem Is Nothing Then
        Exit Sub
    End If

    Set itmX = lvwSavedPackages.SelectedItem
    
    sFileName = itmX.Text

    sPath = goUtil.AttachReposPath

    sMyFilter = sMyFilter & "ZIP File" & " (" & sFileName & ")" & SD & sFileName & SD

    sPath = goUtil.utGetPath(App.EXEName, "TempZipDir", "You Can Drag and Drop " & sFileName & " to your email program.", "You Can Drag and Drop " & sFileName & " to your email program.", sPath, Me.hWnd, sMyFilter, sSelFile)
    
    If sSelFile <> vbNullString Then
        sPath = goUtil.AttachReposPath & sFileName
    End If
    
    If goUtil.utFileExists(sPath) Then
        lRet = goUtil.utShellExecute(GetDesktopWindow, "OPEN", sPath, vbNullString, App.Path, vbNormalFocus, False, False, True)
    End If
    
    PopulatelvwSavedPackages

    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdViewFile_Click"
End Sub

Private Sub LoadHeaderlvwSavedPackages()
    On Error GoTo EH
    'set the columnheaders
    With lvwSavedPackages
        .Sorted = True
        .ColumnHeaders.Add , "Name", "Name"
        .ColumnHeaders.Add , "DateLastUpdated", "Date Last Updated"
        .ColumnHeaders.Add , "DateLastUpdatedSort", "Sort Date Last Updated"
        .ColumnHeaders.Add , "DateCreated", "Date Created"
        .ColumnHeaders.Add , "DateCreatedSort", "Sort Date Created"
        
        '"Avail WOrd XL Forms"
        .ColumnHeaders.Item(SavedPackagesListView.Name).Width = 3000
        .ColumnHeaders.Item(SavedPackagesListView.Name).Alignment = lvwColumnLeft
        .ColumnHeaders.Item(SavedPackagesListView.DateLastUpdated).Width = 1500
        .ColumnHeaders.Item(SavedPackagesListView.DateLastUpdated).Alignment = lvwColumnLeft
        .ColumnHeaders.Item(SavedPackagesListView.DateLastUpdatedSort).Width = 0  'Hidden
        .ColumnHeaders.Item(SavedPackagesListView.DateLastUpdatedSort).Alignment = lvwColumnLeft
        .ColumnHeaders.Item(SavedPackagesListView.DateCreated).Width = 1500
        .ColumnHeaders.Item(SavedPackagesListView.DateCreated).Alignment = lvwColumnLeft
        .ColumnHeaders.Item(SavedPackagesListView.DateCreatedSort).Width = 0   'Hidden
        .ColumnHeaders.Item(SavedPackagesListView.DateCreatedSort).Alignment = lvwColumnLeft
       
    End With
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub LoadHeaderlvwSavedPackages"
End Sub

Public Sub LoadHeaderlvwPackageItem()
    On Error GoTo EH
    Dim bGridOn As Boolean
    Dim bHideDeleted As Boolean
    Dim bHideUploadFlags As Boolean
    Dim sLossFormat As String
    
    sLossFormat = goUtil.IsNullIsVbNullString(mfrmClaim.adoRSAssignments.Fields("LRFormat"))
    
    
    bHideDeleted = CBool(GetSetting(App.EXEName, "GENERAL", "HIDE_DELETED", True))
    bHideUploadFlags = CBool(GetSetting(App.EXEName, "GENERAL", "HIDE_UPLOAD_FLAGS", True))
    
    'set the columnheaders
    With lvwPackageItem
        .ColumnHeaders.Add , "SortOrder", "Sort"
        .ColumnHeaders.Add , "SortOrderSort", "SortOrderSort"
        .ColumnHeaders.Add , "piName", "Name"
        .ColumnHeaders.Add , "Description", "Desc"
        .ColumnHeaders.Add , "AttachmentName", "Attachment"
        .ColumnHeaders.Add , "IsCoApprove", "CO APV"
        .ColumnHeaders.Add , "CoApproveDate", "Date"
        .ColumnHeaders.Add , "CoApproveDesc", "Desc"
        .ColumnHeaders.Add , "IsClientCoReject", "REJECT"
        .ColumnHeaders.Add , "ClientCoRejectDate", "Date"
        .ColumnHeaders.Add , "ClientCoRejectDesc", "Desc"
        .ColumnHeaders.Add , "IsClientCoApprove", "Approve"
        .ColumnHeaders.Add , "ClientCoApproveDate", "Date"
        .ColumnHeaders.Add , "ClientCoApproveDesc", "Desc"
        .ColumnHeaders.Add , "IsClientCoDelete", "Client DEL"
        .ColumnHeaders.Add , "ClientCoDeleteDate", "Date"
        .ColumnHeaders.Add , "ClientCoDeleteDesc", "Desc"
        .ColumnHeaders.Add , "SendMe", "SendME"
        .ColumnHeaders.Add , "SentDate", "Date"
        .ColumnHeaders.Add , "IsDeleted", "Is Deleted"
        .ColumnHeaders.Add , "UpLoadMe", "UploadMe"
        .ColumnHeaders.Add , "DateLastUpdated", "Date Last Updated"
        .ColumnHeaders.Add , "AdminComments", "Admin Comments"
        'Hidden
        .ColumnHeaders.Add , "PackageItemID", "PackageItemID"
        .ColumnHeaders.Add , "PackageID", "PackageID"
        .ColumnHeaders.Add , "AssignmentsID", "AssignmentsID"
        .ColumnHeaders.Add , "ID", "ID"
        .ColumnHeaders.Add , "IDPackage", "IDPackage"
        .ColumnHeaders.Add , "IDAssignments", "IDAssignments"
        .ColumnHeaders.Add , "ReportFormat", "IDAssignments"
        .ColumnHeaders.Add , "RTAttachmentsID", "RTAttachmentsID"
        .ColumnHeaders.Add , "IDRTAttachments", "IDRTAttachments"
        .ColumnHeaders.Add , "Number", "Number"
        .ColumnHeaders.Add , "PackageItemGUID", "PackageItemGUID"
        .ColumnHeaders.Add , "DownLoadMe", "DownLoadMe"
        .ColumnHeaders.Add , "UpdateByUserID", "UpdateByUserID"
        
        
        .Sorted = False
        .SortOrder = lvwAscending
        'SortOrder
        .ColumnHeaders.Item(GuiPackageItemListView.SortOrder).Width = 1000
        .ColumnHeaders.Item(GuiPackageItemListView.SortOrder).Alignment = lvwColumnLeft
        'SortOrderSort
        .ColumnHeaders.Item(GuiPackageItemListView.SortOrderSort).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiPackageItemListView.SortOrderSort).Alignment = lvwColumnLeft
        'piName
        .ColumnHeaders.Item(GuiPackageItemListView.piName).Width = 4000
        .ColumnHeaders.Item(GuiPackageItemListView.piName).Alignment = lvwColumnLeft
        'Description
        .ColumnHeaders.Item(GuiPackageItemListView.Description).Width = 4000
        .ColumnHeaders.Item(GuiPackageItemListView.Description).Alignment = lvwColumnLeft
        'AttachmentName
        .ColumnHeaders.Item(GuiPackageItemListView.AttachmentName).Width = 1500
        .ColumnHeaders.Item(GuiPackageItemListView.AttachmentName).Alignment = lvwColumnLeft
        'IsCoApprove
        .ColumnHeaders.Item(GuiPackageItemListView.IsCoApprove).Width = 400 '0
        .ColumnHeaders.Item(GuiPackageItemListView.IsCoApprove).Alignment = lvwColumnCenter
        .ColumnHeaders.Item(GuiPackageItemListView.IsCoApprove).Icon = GuiPackageItemStatusList.IsApproved
        'CoApproveDate
        .ColumnHeaders.Item(GuiPackageItemListView.CoApproveDate).Width = 1000 '0
        .ColumnHeaders.Item(GuiPackageItemListView.CoApproveDate).Alignment = lvwColumnLeft
        'CoApproveDesc
        .ColumnHeaders.Item(GuiPackageItemListView.CoApproveDesc).Width = 2000 '0
        .ColumnHeaders.Item(GuiPackageItemListView.CoApproveDesc).Alignment = lvwColumnLeft
        'IsClientCoReject
        If StrComp(sLossFormat, "V2ECcarFarmers.clsLossXML01", vbTextCompare) = 0 Then
            .ColumnHeaders.Item(GuiPackageItemListView.IsClientCoReject).Width = 400 '0
        Else
            .ColumnHeaders.Item(GuiPackageItemListView.IsClientCoReject).Width = 0
        End If
        .ColumnHeaders.Item(GuiPackageItemListView.IsClientCoReject).Alignment = lvwColumnCenter
        .ColumnHeaders.Item(GuiPackageItemListView.IsClientCoReject).Icon = GuiPackageItemStatusList.IsRejected
        'ClientCoRejectDate
        If StrComp(sLossFormat, "V2ECcarFarmers.clsLossXML01", vbTextCompare) = 0 Then
            .ColumnHeaders.Item(GuiPackageItemListView.ClientCoRejectDate).Width = 1000 '0
        Else
            .ColumnHeaders.Item(GuiPackageItemListView.ClientCoRejectDate).Width = 0
        End If
        .ColumnHeaders.Item(GuiPackageItemListView.ClientCoRejectDate).Alignment = lvwColumnLeft
        'ClientCoRejectDesc
        If StrComp(sLossFormat, "V2ECcarFarmers.clsLossXML01", vbTextCompare) = 0 Then
            .ColumnHeaders.Item(GuiPackageItemListView.ClientCoRejectDesc).Width = 2000 '0
        Else
            .ColumnHeaders.Item(GuiPackageItemListView.ClientCoRejectDesc).Width = 0
        End If
        .ColumnHeaders.Item(GuiPackageItemListView.ClientCoRejectDesc).Alignment = lvwColumnLeft
        'IsClientCoApprove
        If StrComp(sLossFormat, "V2ECcarFarmers.clsLossXML01", vbTextCompare) = 0 Then
            .ColumnHeaders.Item(GuiPackageItemListView.IsClientCoApprove).Width = 400 '0
        Else
            .ColumnHeaders.Item(GuiPackageItemListView.IsClientCoApprove).Width = 0
        End If
        .ColumnHeaders.Item(GuiPackageItemListView.IsClientCoApprove).Alignment = lvwColumnCenter
        .ColumnHeaders.Item(GuiPackageItemListView.IsClientCoApprove).Icon = GuiPackageItemStatusList.IsApproved
        'ClientCoApproveDate
        If StrComp(sLossFormat, "V2ECcarFarmers.clsLossXML01", vbTextCompare) = 0 Then
            .ColumnHeaders.Item(GuiPackageItemListView.ClientCoApproveDate).Width = 1000 '0
        Else
            .ColumnHeaders.Item(GuiPackageItemListView.ClientCoApproveDate).Width = 0
        End If
        .ColumnHeaders.Item(GuiPackageItemListView.ClientCoApproveDate).Alignment = lvwColumnLeft
        'ClientCoApproveDesc
        If StrComp(sLossFormat, "V2ECcarFarmers.clsLossXML01", vbTextCompare) = 0 Then
            .ColumnHeaders.Item(GuiPackageItemListView.ClientCoApproveDesc).Width = 2000 '0
        Else
            .ColumnHeaders.Item(GuiPackageItemListView.ClientCoApproveDesc).Width = 0
        End If
        .ColumnHeaders.Item(GuiPackageItemListView.ClientCoApproveDesc).Alignment = lvwColumnLeft
        'IsClientCoDelete
        If StrComp(sLossFormat, "V2ECcarFarmers.clsLossXML01", vbTextCompare) = 0 Then
            .ColumnHeaders.Item(GuiPackageItemListView.IsClientCoDelete).Width = 400 '0
        Else
            .ColumnHeaders.Item(GuiPackageItemListView.IsClientCoDelete).Width = 400 '0
        End If
        .ColumnHeaders.Item(GuiPackageItemListView.IsClientCoDelete).Alignment = lvwColumnCenter
        .ColumnHeaders.Item(GuiPackageItemListView.IsClientCoDelete).Icon = GuiPackageItemStatusList.IsDeleted
        'ClientCoDeleteDate
        If StrComp(sLossFormat, "V2ECcarFarmers.clsLossXML01", vbTextCompare) = 0 Then
            .ColumnHeaders.Item(GuiPackageItemListView.ClientCoDeleteDate).Width = 1000 '0
        Else
            .ColumnHeaders.Item(GuiPackageItemListView.ClientCoDeleteDate).Width = 0
        End If
        .ColumnHeaders.Item(GuiPackageItemListView.ClientCoDeleteDate).Alignment = lvwColumnLeft
        'ClientCoDeleteDesc
        If StrComp(sLossFormat, "V2ECcarFarmers.clsLossXML01", vbTextCompare) = 0 Then
            .ColumnHeaders.Item(GuiPackageItemListView.ClientCoDeleteDesc).Width = 2000 '0
        Else
            .ColumnHeaders.Item(GuiPackageItemListView.ClientCoDeleteDesc).Width = 0
        End If
        .ColumnHeaders.Item(GuiPackageItemListView.ClientCoDeleteDesc).Alignment = lvwColumnLeft
        'SendMe
        .ColumnHeaders.Item(GuiPackageItemListView.SendMe).Width = 400
        .ColumnHeaders.Item(GuiPackageItemListView.SendMe).Alignment = lvwColumnCenter
        .ColumnHeaders.Item(GuiPackageItemListView.SendMe).Icon = GuiPackageItemStatusList.ItemHasBeenSent
        'SentDate
        .ColumnHeaders.Item(GuiPackageItemListView.SentDate).Width = 1000
        .ColumnHeaders.Item(GuiPackageItemListView.SentDate).Alignment = lvwColumnLeft
        'Is Deleted
        .ColumnHeaders.Item(GuiPackageItemListView.IsDeleted).Width = 400
        .ColumnHeaders.Item(GuiPackageItemListView.IsDeleted).Alignment = lvwColumnCenter
        .ColumnHeaders.Item(GuiPackageItemListView.IsDeleted).Icon = GuiPackageItemStatusList.IsDeleted
        'UpLoad Me
        .ColumnHeaders.Item(GuiPackageItemListView.UpLoadMe).Width = 400
        .ColumnHeaders.Item(GuiPackageItemListView.UpLoadMe).Alignment = lvwColumnCenter
        .ColumnHeaders.Item(GuiPackageItemListView.UpLoadMe).Icon = GuiPackageItemStatusList.UpLoadMe
        'DateLastUpdated
        .ColumnHeaders.Item(GuiPackageItemListView.DateLastUpdated).Width = 2200
        .ColumnHeaders.Item(GuiPackageItemListView.DateLastUpdated).Alignment = lvwColumnLeft
        'AdminComments
        .ColumnHeaders.Item(GuiPackageItemListView.AdminComments).Width = 10000
        .ColumnHeaders.Item(GuiPackageItemListView.AdminComments).Alignment = lvwColumnLeft
        'hidden
        .ColumnHeaders.Item(GuiPackageItemListView.PackageItemID).Width = 0
        .ColumnHeaders.Item(GuiPackageItemListView.PackageID).Width = 0
        .ColumnHeaders.Item(GuiPackageItemListView.AssignmentsID).Width = 0
        .ColumnHeaders.Item(GuiPackageItemListView.ID).Width = 0
        .ColumnHeaders.Item(GuiPackageItemListView.IDPackage).Width = 0
        .ColumnHeaders.Item(GuiPackageItemListView.IDAssignments).Width = 0
        .ColumnHeaders.Item(GuiPackageItemListView.ReportFormat).Width = 0
        .ColumnHeaders.Item(GuiPackageItemListView.RTAttachmentsID).Width = 0
        .ColumnHeaders.Item(GuiPackageItemListView.IDRTAttachments).Width = 0
        .ColumnHeaders.Item(GuiPackageItemListView.Number).Width = 0
        .ColumnHeaders.Item(GuiPackageItemListView.PackageItemGUID).Width = 0
        .ColumnHeaders.Item(GuiPackageItemListView.DownLoadMe).Width = 0
        .ColumnHeaders.Item(GuiPackageItemListView.UpdateByUserID).Width = 0
    End With
    
    bGridOn = CBool(GetSetting(App.EXEName, "GENERAL", "GRID_ON", False))
    lvwPackageItem.GridLines = bGridOn
    
'    bHideDeleted = CBool(GetSetting(App.EXEName, "GENERAL", "HIDE_DELETED", True))
'    If bHideDeleted Then
'        chkHideDeleted.Value = vbChecked
'    Else
'        chkHideDeleted.Value = vbUnchecked
'    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub LoadHeaderlvwPackageItem"
End Sub



Public Sub PopulatelvwSavedPackages()
    On Error GoTo EH
    'Source Variables
    Dim varySavedPackages As Variant
    Dim sSavedPackage As String
    Dim sSavedPackagePath As String
    Dim iCount As Integer
    Dim itmX As ListItem
    Dim sDateLastUpdated As String
    Dim sDateCreated As String
    Dim oFI As V2ECKeyBoard.clsFileVersion
    Dim myFI As V2ECKeyBoard.FILE_INFORMATION

    lvwSavedPackages.ListItems.Clear
    'BGS 1.2.2002 load the Avail reports
    If Not GetSavedPackages(varySavedPackages) Then
        Exit Sub
    Else
        If IsArray(varySavedPackages) Then
            'Set the File info Object
            Set oFI = New V2ECKeyBoard.clsFileVersion
            
            For iCount = LBound(varySavedPackages) To UBound(varySavedPackages)
                sSavedPackage = varySavedPackages(iCount)
                
                sSavedPackagePath = goUtil.AttachReposPath & sSavedPackage
                
                myFI = oFI.GetFileInformation(sSavedPackagePath)
                sDateLastUpdated = myFI.dtLastModifyTime
                sDateCreated = myFI.dtCreationDate
                
                'If the DateLastUpdated is Earlier than the Date Created...
                'That means the User has not Done anything with it
                If IsDate(sDateLastUpdated) And IsDate(sDateCreated) Then
                    If CDate(sDateLastUpdated) < CDate(sDateCreated) Then
                        'So Make it Blank
                        sDateLastUpdated = vbNullString
                    End If
                End If
                
                Set itmX = lvwSavedPackages.ListItems.Add(, , sSavedPackage)
                
                itmX.SubItems(SavedPackagesListView.DateLastUpdated - 1) = Format(sDateLastUpdated, "MM/DD/YYYY HH:MM:SS")
                itmX.SubItems(SavedPackagesListView.DateLastUpdatedSort - 1) = Format(sDateLastUpdated, "YYYY/MM/DD HH:MM:SS")
                itmX.SubItems(SavedPackagesListView.DateCreated - 1) = Format(sDateCreated, "MM/DD/YYYY HH:MM:SS")
                itmX.SubItems(SavedPackagesListView.DateCreatedSort - 1) = Format(sDateCreated, "YYYY/MM/DD HH:MM:SS")
            Next
        End If
    End If

    'cleanup
    Set itmX = Nothing
    Set oFI = Nothing
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub PopulatelvwSavedPackages"
End Sub


Private Sub PopulatelvwPackageItem()
    On Error GoTo EH
    Dim iMyIcon As Long
    Dim sFlagText As String
    Dim sTemp As String
    Dim itmX As ListItem
    Dim oListView As MSComctlLib.ListView
    Dim RS As ADODB.Recordset
    Dim sReportFormatTag As String 'used to remove this item from the Add Item Boxes
    Dim lSortOrder As Long
    Dim sSortOrder As String
    Dim sRTAttachmentsID As String
    Dim sNumber As String
    'Clear the List view
    Set oListView = lvwPackageItem
    
    oListView.ListItems.Clear
    
    mfrmClaim.SetadoRSPackageItem msAssignmentsID, msPackageID
    Set RS = mfrmClaim.adoRSPackageItem
    
    If Not RS.EOF Then
        RS.MoveFirst
        Do Until RS.EOF
            'SortOrder
            lSortOrder = goUtil.IsNullIsVbNullString(RS.Fields("SortOrder"))
            If lSortOrder = DELETED_SORTORDER Then
                sSortOrder = "DEL"
            Else
                sSortOrder = lSortOrder
            End If
            Set itmX = oListView.ListItems.Add(, , sSortOrder) ' """" & goUtil.IsNullIsVbNullString(RS.Fields("ID")) & """"
            'SortOrderSort
            itmX.SubItems(GuiPackageItemListView.SortOrderSort - 1) = goUtil.utNumInTextSortFormat(sSortOrder)
            'piName
            itmX.SubItems(GuiPackageItemListView.piName - 1) = goUtil.IsNullIsVbNullString(RS.Fields("Name"))
            'Description
            itmX.SubItems(GuiPackageItemListView.Description - 1) = goUtil.IsNullIsVbNullString(RS.Fields("Description"))
            'AttachmentName
            itmX.SubItems(GuiPackageItemListView.AttachmentName - 1) = goUtil.IsNullIsVbNullString(RS.Fields("AttachmentName"))
            'IsCoApprove
            If CBool(RS.Fields("IsCoApprove")) Then
                iMyIcon = GuiPackageItemStatusList.IsApproved
            Else
                iMyIcon = Empty
            End If
            sFlagText = goUtil.GetFlagText(RS.Fields("IsCoApprove"))
            itmX.SubItems(GuiPackageItemListView.IsCoApprove - 1) = sFlagText
            itmX.ListSubItems(GuiPackageItemListView.IsCoApprove - 1).ReportIcon = iMyIcon
            'CoApproveDate
            If Not IsNull(RS.Fields("CoApproveDate").Value) Then
                If IsDate(RS.Fields("CoApproveDate").Value) Then
                    itmX.SubItems(GuiPackageItemListView.CoApproveDate - 1) = Format(RS.Fields("CoApproveDate").Value, "MM/DD/YYYY HH:MM:SS")
                Else
                    itmX.SubItems(GuiPackageItemListView.CoApproveDate - 1) = vbNullString
                End If
            Else
                itmX.SubItems(GuiPackageItemListView.CoApproveDate - 1) = vbNullString
            End If
            'CoApproveDesc
            itmX.SubItems(GuiPackageItemListView.CoApproveDesc - 1) = goUtil.IsNullIsVbNullString(RS.Fields("CoApproveDesc"))
            'IsClientCoReject
            If CBool(RS.Fields("IsClientCoReject")) Then
                iMyIcon = GuiPackageItemStatusList.IsRejected
            Else
                iMyIcon = Empty
            End If
            sFlagText = goUtil.GetFlagText(RS.Fields("IsClientCoReject"))
            itmX.SubItems(GuiPackageItemListView.IsClientCoReject - 1) = sFlagText
            itmX.ListSubItems(GuiPackageItemListView.IsClientCoReject - 1).ReportIcon = iMyIcon
            'ClientCoRejectDate
            If Not IsNull(RS.Fields("ClientCoRejectDate").Value) Then
                If IsDate(RS.Fields("ClientCoRejectDate").Value) Then
                    itmX.SubItems(GuiPackageItemListView.ClientCoRejectDate - 1) = Format(RS.Fields("ClientCoRejectDate").Value, "MM/DD/YYYY HH:MM:SS")
                Else
                    itmX.SubItems(GuiPackageItemListView.ClientCoRejectDate - 1) = vbNullString
                End If
            Else
                itmX.SubItems(GuiPackageItemListView.ClientCoRejectDate - 1) = vbNullString
            End If
            'ClientCoRejectDesc
            itmX.SubItems(GuiPackageItemListView.ClientCoRejectDesc - 1) = goUtil.IsNullIsVbNullString(RS.Fields("ClientCoRejectDesc"))
            'IsClientCoApprove
            If CBool(RS.Fields("IsClientCoApprove")) Then
                iMyIcon = GuiPackageItemStatusList.IsApproved
            Else
                iMyIcon = Empty
            End If
            sFlagText = goUtil.GetFlagText(RS.Fields("IsClientCoApprove"))
            itmX.SubItems(GuiPackageItemListView.IsClientCoApprove - 1) = sFlagText
            itmX.ListSubItems(GuiPackageItemListView.IsClientCoApprove - 1).ReportIcon = iMyIcon
            'ClientCoApproveDate
            If Not IsNull(RS.Fields("ClientCoApproveDate").Value) Then
                If IsDate(RS.Fields("ClientCoApproveDate").Value) Then
                    itmX.SubItems(GuiPackageItemListView.ClientCoApproveDate - 1) = Format(RS.Fields("ClientCoApproveDate").Value, "MM/DD/YYYY HH:MM:SS")
                Else
                    itmX.SubItems(GuiPackageItemListView.ClientCoApproveDate - 1) = vbNullString
                End If
            Else
                itmX.SubItems(GuiPackageItemListView.ClientCoApproveDate - 1) = vbNullString
            End If
            'ClientCoApproveDesc
            itmX.SubItems(GuiPackageItemListView.ClientCoApproveDesc - 1) = goUtil.IsNullIsVbNullString(RS.Fields("ClientCoApproveDesc"))
            'IsClientCoDelete
            If CBool(RS.Fields("IsClientCoDelete")) Then
                iMyIcon = GuiPackageItemStatusList.IsDeleted
            Else
                iMyIcon = Empty
            End If
            sFlagText = goUtil.GetFlagText(RS.Fields("IsClientCoDelete"))
            itmX.SubItems(GuiPackageItemListView.IsClientCoDelete - 1) = sFlagText
            itmX.ListSubItems(GuiPackageItemListView.IsClientCoDelete - 1).ReportIcon = iMyIcon
            'ClientCoDeleteDate
            If Not IsNull(RS.Fields("ClientCoDeleteDate").Value) Then
                If IsDate(RS.Fields("ClientCoDeleteDate").Value) Then
                    itmX.SubItems(GuiPackageItemListView.ClientCoDeleteDate - 1) = Format(RS.Fields("ClientCoDeleteDate").Value, "MM/DD/YYYY HH:MM:SS")
                Else
                    itmX.SubItems(GuiPackageItemListView.ClientCoDeleteDate - 1) = vbNullString
                End If
            Else
                itmX.SubItems(GuiPackageItemListView.ClientCoDeleteDate - 1) = vbNullString
            End If
            'ClientCoDeleteDesc
            itmX.SubItems(GuiPackageItemListView.ClientCoDeleteDesc - 1) = goUtil.IsNullIsVbNullString(RS.Fields("ClientCoDeleteDesc"))
            'SendMe
            If CBool(RS.Fields("SendMe")) Then
                iMyIcon = GuiPackageItemStatusList.IsChecked
            Else
                'use the Sent Icon if not flaged to send but was sent
                If Not IsNull(RS.Fields("SentDate").Value) Then
                    If IsDate(RS.Fields("SentDate").Value) Then
                        iMyIcon = GuiPackageItemStatusList.ItemHasBeenSent
                    Else
                        iMyIcon = Empty
                    End If
                Else
                    iMyIcon = Empty
                End If
            End If
            sFlagText = goUtil.GetFlagText(RS.Fields("SendMe"))
            itmX.SubItems(GuiPackageItemListView.SendMe - 1) = sFlagText
            itmX.ListSubItems(GuiPackageItemListView.SendMe - 1).ReportIcon = iMyIcon
            'SentDate
            If Not IsNull(RS.Fields("SentDate").Value) Then
                If IsDate(RS.Fields("SentDate").Value) Then
                    itmX.SubItems(GuiPackageItemListView.SentDate - 1) = Format(RS.Fields("SentDate").Value, "MM/DD/YYYY HH:MM:SS")
                Else
                    itmX.SubItems(GuiPackageItemListView.SentDate - 1) = vbNullString
                End If
            Else
                itmX.SubItems(GuiPackageItemListView.SentDate - 1) = vbNullString
            End If
            'IsDeleted
            If CBool(RS.Fields("IsDeleted")) Then
                iMyIcon = GuiPackageItemStatusList.IsDeleted
            Else
                iMyIcon = Empty
            End If
            sFlagText = goUtil.GetFlagText(RS.Fields("IsDeleted"))
            itmX.SubItems(GuiPackageItemListView.IsDeleted - 1) = sFlagText
            itmX.ListSubItems(GuiPackageItemListView.IsDeleted - 1).ReportIcon = iMyIcon
            'UpLoadMe
            If CBool(RS.Fields("UpLoadMe")) Then
                iMyIcon = GuiPackageItemStatusList.UpLoadMe
            Else
                iMyIcon = Empty
            End If
            sFlagText = goUtil.GetFlagText(RS.Fields("UpLoadMe"))
            itmX.SubItems(GuiPackageItemListView.UpLoadMe - 1) = sFlagText
            itmX.ListSubItems(GuiPackageItemListView.UpLoadMe - 1).ReportIcon = iMyIcon
            'DateLastUpdated
            If Not IsNull(RS.Fields("DateLastUpdated").Value) Then
                If IsDate(RS.Fields("DateLastUpdated").Value) Then
                    itmX.SubItems(GuiPackageItemListView.DateLastUpdated - 1) = Format(RS.Fields("DateLastUpdated").Value, "MM/DD/YYYY HH:MM:SS")
                Else
                    itmX.SubItems(GuiPackageItemListView.DateLastUpdated - 1) = vbNullString
                End If
            Else
                itmX.SubItems(GuiPackageItemListView.DateLastUpdated - 1) = vbNullString
            End If
            'AdminComments
            itmX.SubItems(GuiPackageItemListView.AdminComments - 1) = goUtil.IsNullIsVbNullString(RS.Fields("AdminComments"))
            '----------------------------Hidden Items-----------------------
            'PackageItemID
            itmX.SubItems(GuiPackageItemListView.PackageItemID - 1) = goUtil.IsNullIsVbNullString(RS.Fields("PackageItemID"))
            'PackageID
            itmX.SubItems(GuiPackageItemListView.PackageID - 1) = goUtil.IsNullIsVbNullString(RS.Fields("PackageID"))
            'AssignmentsID
            itmX.SubItems(GuiPackageItemListView.AssignmentsID - 1) = goUtil.IsNullIsVbNullString(RS.Fields("AssignmentsID"))
            'ID
            itmX.SubItems(GuiPackageItemListView.ID - 1) = goUtil.IsNullIsVbNullString(RS.Fields("ID"))
            'IDPackage
            itmX.SubItems(GuiPackageItemListView.IDPackage - 1) = goUtil.IsNullIsVbNullString(RS.Fields("IDPackage"))
            'IDAssignments
            itmX.SubItems(GuiPackageItemListView.IDAssignments - 1) = goUtil.IsNullIsVbNullString(RS.Fields("IDAssignments"))
            'ReportFormat
            itmX.SubItems(GuiPackageItemListView.ReportFormat - 1) = goUtil.IsNullIsVbNullString(RS.Fields("ReportFormat"))
            sReportFormatTag = goUtil.IsNullIsVbNullString(RS.Fields("ReportFormat"))
            sReportFormatTag = Mid(sReportFormatTag, 200)
            sReportFormatTag = Trim(sReportFormatTag)
            itmX.Tag = sReportFormatTag
            'RTAttachmentsID
            sRTAttachmentsID = goUtil.IsNullIsVbNullString(RS.Fields("RTAttachmentsID"))
            If sRTAttachmentsID = "0" Then
                sRTAttachmentsID = "Null"
            End If
            itmX.SubItems(GuiPackageItemListView.RTAttachmentsID - 1) = sRTAttachmentsID
            'IDRTAttachments
            sRTAttachmentsID = goUtil.IsNullIsVbNullString(RS.Fields("IDRTAttachments"))
            If sRTAttachmentsID = "0" Then
                sRTAttachmentsID = "Null"
            End If
            itmX.SubItems(GuiPackageItemListView.IDRTAttachments - 1) = sRTAttachmentsID
            'Number
            If IsNull(RS.Fields("Number")) Then
                sNumber = "Null"
            Else
                sNumber = RS.Fields("Number")
            End If
            itmX.SubItems(GuiPackageItemListView.Number - 1) = sNumber
            'PackageItemGUID
            itmX.SubItems(GuiPackageItemListView.PackageItemGUID - 1) = goUtil.IsNullIsVbNullString(RS.Fields("PackageItemGUID"))
            'DownLoadMe
            sFlagText = goUtil.GetFlagText(RS.Fields("DownLoadMe"))
            itmX.SubItems(GuiPackageItemListView.DownLoadMe - 1) = sFlagText
            'UpdateByUserID
            itmX.SubItems(GuiPackageItemListView.UpdateByUserID - 1) = goUtil.IsNullIsVbNullString(RS.Fields("UpdateByUserID"))
            itmX.Selected = False
            
            RS.MoveNext
        Loop
    End If
    oListView.SortKey = GuiPackageItemListView.SortOrder
    oListView.Sorted = True
    'Cleanup
    Set itmX = Nothing
    Set oListView = Nothing
    Set RS = Nothing
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub PopulatelvwPackageItem"
    oListView.Visible = True
End Sub

Public Function GetSavedPackages(pvarySavedPackages As Variant) As Boolean
    On Error GoTo EH
    Dim sReport As String
    Dim saryReports() As String
    Dim iReportCount As Integer
    Dim bFound As Boolean
    Dim sIbnumber As String
    
    sIbnumber = mfrmClaim.MyClaimsList.GetClaimItemAsString(GuiAssignments.IBNUM)
    sIbnumber = sIbnumber & "_"
    'BGS get the .zip reports
    sReport = Dir(goUtil.AttachReposPath & "\" & sIbnumber & "*.zip")
    Do Until sReport = vbNullString
        bFound = True
        iReportCount = iReportCount + 1
        ReDim Preserve saryReports(1 To iReportCount)
        saryReports(iReportCount) = sReport
        sReport = Dir
    Loop
    If bFound Then
        pvarySavedPackages = saryReports
        GetSavedPackages = True
    Else
        GetSavedPackages = False
    End If
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function GetSavedPackages"
End Function

Public Function LoadReports() As Boolean
    On Error GoTo EH
    Dim lNewIndex As Long
    Dim sData As String
    Dim RS As ADODB.Recordset
    Dim RSIB As ADODB.Recordset
    Dim RSPayment As ADODB.Recordset
    Dim RSAttach As ADODB.Recordset
    'Any Report types that have Multiple parts go here
    'Main Reports (Multi Report Type )
    Dim RSRTPhotoReportList As ADODB.Recordset  'Photos will have multiple Reports
    Dim sPhotoReportData As String
    Dim sPhotoReportName As String
    Dim sPhotoReportNumber As String
    Dim sPhotoReportDesc As String
    Dim RSRTWSDiagramList As ADODB.Recordset    'Worksheet Diagrams will have multiple Reports
    Dim sDiagramData As String
    Dim sDiagramName As String
    Dim sDiagramNumber As String
    Dim sDiagramDesc As String
    Dim sName As String
    'Loss Report
    Dim MyadoRSAssignments As ADODB.Recordset
    Dim sIBNUM As String
    Dim sCLIENTNUM As String
    
    'Load RecordSets
    mfrmClaim.SetadoRSMainReports msAssignmentsID
    mfrmClaim.SetadoRSMainReportsHistory msAssignmentsID
    
    'Load the Multi Report RS
    mfrmClaim.SetadoRSRTPhotoReportList msAssignmentsID     ' Photo Reports
    mfrmClaim.SetadoRSRTWSDiagramList msAssignmentsID       ' WorkSheet Diagrams
    'Set the Multi Report RS
    Set RSRTPhotoReportList = mfrmClaim.adoRSRTPhotoReportList    ' Photo Reports
    Set RSRTWSDiagramList = mfrmClaim.adoRSRTWSDiagramList      ' WorkSheet Diagrams
    
    'Load the Main Reports
    Set RS = mfrmClaim.adoRSMainReports
    cboMainReports.Clear
    
    'First Add the Loss Report to main Reports
    mfrmClaim.SetadoRSAssignments msAssignmentsID
    Set MyadoRSAssignments = mfrmClaim.adoRSAssignments
    sIBNUM = goUtil.IsNullIsVbNullString(MyadoRSAssignments.Fields("IBNUM"))
    sCLIENTNUM = goUtil.IsNullIsVbNullString(MyadoRSAssignments.Fields("CLIENTNUM"))
    
    sData = vbNullString
    sData = "Loss Report_" & sIBNUM & "_" & sCLIENTNUM
    sData = sData & String(200, " ")
    sData = sData & "LRFormat"
    cboMainReports.AddItem sData
    lNewIndex = cboMainReports.NewIndex
    cboMainReports.ItemData(lNewIndex) = msAssignmentsID
    
    If RS.RecordCount > 0 Then
        
        Do Until RS.EOF
            'Check For Multi Report item
            If StrComp(goUtil.IsNullIsVbNullString(RS.Fields("SectionLevel04")), "Multi", vbTextCompare) = 0 Then
                If StrComp(goUtil.IsNullIsVbNullString(RS.Fields("ProjectName")), "ECrpt" & goUtil.gsCurCarDBName & "_arRptPhotos", vbTextCompare) = 0 Then
                    ' Photo Reports
                    Do Until RSRTPhotoReportList.EOF
                        sData = vbNullString
                        sName = goUtil.IsNullIsVbNullString(RS.Fields("SectionLevel05"))
                        sPhotoReportName = RSRTPhotoReportList.Fields("Name").Value
                        sPhotoReportNumber = RSRTPhotoReportList.Fields("Number").Value
                        sPhotoReportDesc = RSRTPhotoReportList.Fields("Description").Value
                        If Len(sPhotoReportDesc) > 20 Then
                            sPhotoReportDesc = left(sPhotoReportDesc, 10) & "..."
                        End If
                        sData = sPhotoReportName & " _ " & sName & " - (" & sPhotoReportDesc & ")"
                        sData = sData & String(200, " ")
                        sData = sData & RS.Fields("ProjectName").Value & "|" & RS.Fields("ClassName").Value & "|" & RS.Fields("Version").Value
                        sPhotoReportData = sData & "|"
                        sPhotoReportData = sPhotoReportData & RSRTPhotoReportList.Fields("Number")
                        cboMainReports.AddItem sPhotoReportData
                        lNewIndex = cboMainReports.NewIndex
                        cboMainReports.ItemData(lNewIndex) = RS.Fields("ApplicationID").Value
                        RSRTPhotoReportList.MoveNext
                    Loop
                    
                ElseIf StrComp(goUtil.IsNullIsVbNullString(RS.Fields("ProjectName")), "ECrpt" & goUtil.gsCurCarDBName & "_arWorkSheetDiag", vbTextCompare) = 0 Then
                    ' WorkSheet Diagrams
                    Do Until RSRTWSDiagramList.EOF
                        sData = vbNullString
                        sName = goUtil.IsNullIsVbNullString(RS.Fields("SectionLevel05"))
                        sDiagramName = RSRTWSDiagramList.Fields("Name").Value
                        sDiagramNumber = RSRTWSDiagramList.Fields("Number").Value
                        sDiagramDesc = RSRTWSDiagramList.Fields("Description").Value
                        If Len(sDiagramDesc) > 20 Then
                            sDiagramDesc = left(sDiagramDesc, 10) & "..."
                        End If
                        sData = sDiagramName & " _ " & sName & " - (" & sDiagramDesc & ")"
                        sData = sData & String(200, " ")
                        sData = sData & RS.Fields("ProjectName").Value & "|" & RS.Fields("ClassName").Value & "|" & RS.Fields("Version").Value
                        sDiagramData = sData & "|"
                        sDiagramData = sDiagramData & RSRTWSDiagramList.Fields("Number")
                        sDiagramData = sDiagramData & "|"
                        sDiagramData = sDiagramData & RSRTWSDiagramList.Fields("DiagramPhotoName")
                        cboMainReports.AddItem sDiagramData
                        lNewIndex = cboMainReports.NewIndex
                        cboMainReports.ItemData(lNewIndex) = RS.Fields("ApplicationID").Value
                        RSRTWSDiagramList.MoveNext
                    Loop
                End If
            Else
                sData = vbNullString
                sName = goUtil.IsNullIsVbNullString(RS.Fields("SectionLevel05"))
                sData = RS.Fields("Description").Value & " _ " & sName
                sData = sData & String(200, " ")
                sData = sData & RS.Fields("ProjectName").Value & "|" & RS.Fields("ClassName").Value & "|" & RS.Fields("Version").Value
                cboMainReports.AddItem sData
                lNewIndex = cboMainReports.NewIndex
                cboMainReports.ItemData(lNewIndex) = RS.Fields("ApplicationID").Value
            End If
           
            RS.MoveNext
        Loop
    End If
    
    'Need to Add Items from History Table
    Set RS = mfrmClaim.adoRSMainReportsHistory
    
    If RS.RecordCount > 0 Then
        Do Until RS.EOF
            'Check For Multi Report item
            If StrComp(goUtil.IsNullIsVbNullString(RS.Fields("SectionLevel04")), "Multi", vbTextCompare) = 0 Then
                If StrComp(goUtil.IsNullIsVbNullString(RS.Fields("ProjectName")), "ECrpt" & goUtil.gsCurCarDBName & "_arRptPhotos", vbTextCompare) = 0 Then
                    ' Photo Reports
                    Do Until RSRTPhotoReportList.EOF
                        sData = vbNullString
                        sName = goUtil.IsNullIsVbNullString(RS.Fields("SectionLevel05"))
                        sPhotoReportName = RSRTPhotoReportList.Fields("Name").Value
                        sPhotoReportNumber = RSRTPhotoReportList.Fields("Number").Value
                        sPhotoReportDesc = RSRTPhotoReportList.Fields("Description").Value
                        If Len(sPhotoReportDesc) > 20 Then
                            sPhotoReportDesc = left(sPhotoReportDesc, 10) & "..."
                        End If
                        sData = sPhotoReportName & " _ " & sName & " - (" & sPhotoReportDesc & ")"
                        sData = sData & String(200, " ")
                        sData = sData & RS.Fields("ProjectName").Value & "|" & RS.Fields("ClassName").Value & "|" & RS.Fields("Version").Value
                        sPhotoReportData = sData & "|"
                        sPhotoReportData = sPhotoReportData & RSRTPhotoReportList.Fields("Number")
                        cboMainReports.AddItem sPhotoReportData
                        lNewIndex = cboMainReports.NewIndex
                        cboMainReports.ItemData(lNewIndex) = RS.Fields("ApplicationID").Value
                        RSRTPhotoReportList.MoveNext
                    Loop
                    
                ElseIf StrComp(goUtil.IsNullIsVbNullString(RS.Fields("ProjectName")), "ECrpt" & goUtil.gsCurCarDBName & "_arWorkSheetDiag", vbTextCompare) = 0 Then
                    ' WorkSheet Diagrams
                    Do Until RSRTWSDiagramList.EOF
                        sData = vbNullString
                        sName = goUtil.IsNullIsVbNullString(RS.Fields("SectionLevel05"))
                        sDiagramName = RSRTWSDiagramList.Fields("Name").Value
                        sDiagramNumber = RSRTWSDiagramList.Fields("Number").Value
                        sDiagramDesc = RSRTWSDiagramList.Fields("Description").Value
                        If Len(sDiagramDesc) > 20 Then
                            sDiagramDesc = left(sDiagramDesc, 10) & "..."
                        End If
                        sData = sDiagramName & " _ " & sName & " - (" & sDiagramDesc & ")"
                        sData = sData & String(200, " ")
                        sData = sData & RS.Fields("ProjectName").Value & "|" & RS.Fields("ClassName").Value & "|" & RS.Fields("Version").Value
                        sDiagramData = sData & "|"
                        sDiagramData = sDiagramData & RSRTWSDiagramList.Fields("Number")
                        sDiagramData = sDiagramData & "|"
                        sDiagramData = sDiagramData & RSRTWSDiagramList.Fields("DiagramPhotoName")
                        cboMainReports.AddItem sDiagramData
                        lNewIndex = cboMainReports.NewIndex
                        cboMainReports.ItemData(lNewIndex) = RS.Fields("ApplicationID").Value
                        RSRTWSDiagramList.MoveNext
                    Loop
                End If
            Else
                sData = vbNullString
                sName = goUtil.IsNullIsVbNullString(RS.Fields("SectionLevel05"))
                sData = RS.Fields("Description").Value & " _ " & sName & " - (Previous Version) "
                sData = sData & String(200, " ")
                sData = sData & RS.Fields("ProjectName").Value & "|" & RS.Fields("ClassName").Value & "|" & RS.Fields("Version").Value
                cboMainReports.AddItem sData
                lNewIndex = cboMainReports.NewIndex
                cboMainReports.ItemData(lNewIndex) = RS.Fields("ApplicationID").Value
            End If
            
            RS.MoveNext
        Loop
    End If
    
    'Load RecordSets
    mfrmClaim.SetadoRSCarSpecReports msAssignmentsID
    mfrmClaim.SetadoRSCarSpecReportsHistory msAssignmentsID
    
   'Load the Carrier Specific Reports
    Set RS = mfrmClaim.adoRSCarSpecReports
'    cboCarSpecReports.Clear
    If RS.RecordCount > 0 Then
        Do Until RS.EOF
            sData = vbNullString
            sName = goUtil.IsNullIsVbNullString(RS.Fields("SectionLevel05"))
            sData = RS.Fields("Description").Value & " _ " & sName
            sData = sData & String(200, " ")
            sData = sData & RS.Fields("ProjectName").Value & "|" & RS.Fields("ClassName").Value & "|" & RS.Fields("Version").Value
'            cboCarSpecReports.AddItem sData
            cboMainReports.AddItem sData
'            lNewIndex = cboCarSpecReports.NewIndex
            lNewIndex = cboMainReports.NewIndex
'            cboCarSpecReports.ItemData(lNewIndex) = RS.Fields("ApplicationID").Value
            cboMainReports.ItemData(lNewIndex) = RS.Fields("ApplicationID").Value
            RS.MoveNext
        Loop
    End If
    
    'Need to Add Items from History Table
    Set RS = mfrmClaim.adoRSCarSpecReportsHistory
    
    If RS.RecordCount > 0 Then
        Do Until RS.EOF
            sData = vbNullString
            sName = goUtil.IsNullIsVbNullString(RS.Fields("SectionLevel05"))
            sData = RS.Fields("Description").Value & " _ " & sName & " - (Previous Version) "
            sData = sData & String(200, " ")
            sData = sData & RS.Fields("ProjectName").Value & "|" & RS.Fields("ClassName").Value & "|" & RS.Fields("Version").Value
'            cboCarSpecReports.AddItem sData
            cboMainReports.AddItem sData
'            lNewIndex = cboCarSpecReports.NewIndex
            lNewIndex = cboMainReports.NewIndex
'            cboCarSpecReports.ItemData(lNewIndex) = RS.Fields("ApplicationID").Value
            cboMainReports.ItemData(lNewIndex) = RS.Fields("ApplicationID").Value
            RS.MoveNext
        Loop
    End If
    
    'Load Recordsets
    mfrmClaim.SetadoRSIB msAssignmentsID 'Actual Bills
    mfrmClaim.SetadoRSBillingReports ' software for Bills
    mfrmClaim.SetadoRSBillingReportsHistory 'software history for Bills
    
    'Load Billing Report
    Set RS = mfrmClaim.adoRSBillingReports ' Software
    Set RSIB = mfrmClaim.adoRSIB ' Actual IB Data
'    cboIB.Clear
    If RS.RecordCount = 1 And RSIB.RecordCount > 0 Then
        RS.MoveFirst
        Do Until RSIB.EOF
            sData = vbNullString
            sName = goUtil.IsNullIsVbNullString(RS.Fields("SectionLevel05"))
            sData = RSIB.Fields("sIBNumber").Value & " _ " & sName
            sData = sData & String(200, " ")
            sData = sData & RS.Fields("ProjectName").Value & "|" & RS.Fields("ClassName").Value & "|" & RS.Fields("Version").Value
            sData = sData & "|"
            sData = sData & RSIB.Fields("IB14a_sSupplement")
'            cboIB.AddItem sData
            cboMainReports.AddItem sData
'            lNewIndex = cboIB.NewIndex
            lNewIndex = cboMainReports.NewIndex
            'Use the Actual IB Data Unique ID
'            cboIB.ItemData(lNewIndex) = RSIB.Fields("IBID").Value
            cboMainReports.ItemData(lNewIndex) = RSIB.Fields("IBID").Value
            RSIB.MoveNext
        Loop
    End If
    
    'Need to Add Items from History Table
    Set RS = mfrmClaim.adoRSBillingReportsHistory
    
    If RS.RecordCount > 0 And RSIB.RecordCount > 0 Then
        RS.MoveFirst
        Do Until RSIB.EOF
            sData = vbNullString
            sName = goUtil.IsNullIsVbNullString(RS.Fields("SectionLevel05"))
            sData = RSIB.Fields("sIBNumber").Value & " _ " & sName & " - (Previous Version) "
            sData = sData & String(200, " ")
            sData = sData & RS.Fields("ProjectName").Value & "|" & RS.Fields("ClassName").Value & "|" & RS.Fields("Version").Value
            sData = sData & "|"
            sData = sData & RSIB.Fields("IB14a_sSupplement")
'            cboIB.AddItem sData
            cboMainReports.AddItem sData
'            lNewIndex = cboIB.NewIndex
            lNewIndex = cboMainReports.NewIndex
            'Use the Actual IB Data Unique ID
'            cboIB.ItemData(lNewIndex) = RSIB.Fields("IBID").Value
            cboMainReports.ItemData(lNewIndex) = RSIB.Fields("IBID").Value
            RSIB.MoveNext
        Loop
    End If
    
    'Load Recordsets
    mfrmClaim.SetadoRSPayment msAssignmentsID ' Actual payments
    mfrmClaim.SetadoRSPaymentReports 'Software for Payments
    mfrmClaim.SetadoRSPaymentReportsHistory 'Software History for Payments
    
    'Load Payment Report
    Set RS = mfrmClaim.adoRSPaymentReports ' Software
    Set RSPayment = mfrmClaim.adoRSPayment ' Actual Payment Data
'    cboPayments.Clear
    If RS.RecordCount = 1 And RSPayment.RecordCount > 0 Then
        RS.MoveFirst
        Do Until RSPayment.EOF
            sData = vbNullString
            sName = goUtil.IsNullIsVbNullString(RS.Fields("SectionLevel05"))
            If CBool(RSPayment.Fields("PrintOnIB").Value) Then
                sData = RS.Fields("Description").Value & " (" & RSPayment.Fields("CheckNum").Value & ") - Print On IB" & " _ " & sName
            Else
                sData = RS.Fields("Description").Value & " (" & RSPayment.Fields("CheckNum").Value & ")" & " _ " & sName
            End If
            sData = sData & String(200, " ")
            sData = sData & RS.Fields("ProjectName").Value & "|" & RS.Fields("ClassName").Value & "|" & RS.Fields("Version").Value
            sData = sData & "|"
            sData = sData & RSPayment.Fields("CheckNum")
'            cboPayments.AddItem sData
            cboMainReports.AddItem sData
'            lNewIndex = cboPayments.NewIndex
            lNewIndex = cboMainReports.NewIndex
            'Use the Actual Payment Data Unique ID
'            cboPayments.ItemData(lNewIndex) = RSPayment.Fields("RTChecksID").Value
            cboMainReports.ItemData(lNewIndex) = RSPayment.Fields("RTChecksID").Value
            RSPayment.MoveNext
        Loop
    End If
    
    'Need to Add Items from History Table
    Set RS = mfrmClaim.adoRSPaymentReportsHistory
    
    If RS.RecordCount > 0 And RSPayment.RecordCount > 0 Then
        RS.MoveFirst
        Do Until RSPayment.EOF
            sData = vbNullString
            sName = goUtil.IsNullIsVbNullString(RS.Fields("SectionLevel05"))
            sData = RS.Fields("Description").Value & " (" & RSPayment.Fields("CheckNum").Value & ")" & " _ " & sName
            sData = sData & String(200, " ")
            sData = sData & RS.Fields("ProjectName").Value & "|" & RS.Fields("ClassName").Value & "|" & RS.Fields("Version").Value
            sData = sData & "|"
            sData = sData & RSPayment.Fields("CheckNum")
'            cboPayments.AddItem sData
            cboMainReports.AddItem sData
'            lNewIndex = cboPayments.NewIndex
            lNewIndex = cboMainReports.NewIndex
            'Use the Actual IB Data Unique ID
'            cboPayments.ItemData(lNewIndex) = RSPayment.Fields("RTChecksID").Value
            cboMainReports.ItemData(lNewIndex) = RSPayment.Fields("RTChecksID").Value
            RSPayment.MoveNext
        Loop
    End If
    
    'Need to Add Attahcments
    mfrmClaim.SetadoRSRTAttachments msAssignmentsID, True
    Set RS = mfrmClaim.adoRSRTAttachments
'    cboAttachments.Clear
    
    If RS.RecordCount > 0 Then
        RS.MoveFirst
        Do Until RS.EOF
            sData = vbNullString
            sName = goUtil.IsNullIsVbNullString(RS.Fields("AttachName"))
            sData = sName & " _ " & RS.Fields("Description").Value
            sData = sData & String(200, " ")
            sData = sData & RS.Fields("Attachment") & "|"
'            cboAttachments.AddItem sData
            cboMainReports.AddItem sData
'            lNewIndex = cboAttachments.NewIndex
            lNewIndex = cboMainReports.NewIndex
            'Use the Actual IB Data Unique ID
'            cboAttachments.ItemData(lNewIndex) = RS.Fields("RTAttachmentsID").Value
            cboMainReports.ItemData(lNewIndex) = RS.Fields("RTAttachmentsID").Value
            RS.MoveNext
        Loop
    End If
    
    LoadReports = True
    
    'cleanup
    Set RS = Nothing
    Set RSIB = Nothing
    Set RSPayment = Nothing
    Set RSRTPhotoReportList = Nothing
    Set RSRTWSDiagramList = Nothing
    Set RSAttach = Nothing
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function LoadReports"
End Function

Public Function RemoveAddedReports() As Boolean
    On Error GoTo EH
    Dim itmX As ListItem
    Dim sReportFormatTag As String
    Dim sFlagText As String
    
    'Need to Remove any reports that have already been added to the PackageItem List view
    
    For Each itmX In lvwPackageItem.ListItems
        'Do not remove items that are deleted
        'So the user can add them back in if need be
        sFlagText = itmX.SubItems(GuiPackageItemListView.IsDeleted - 1)
        If Not goUtil.GetFlagFromText(sFlagText) Then
            sReportFormatTag = itmX.ListSubItems(GuiPackageItemListView.ReportFormat - 1)
            '1 Main Reports:
            If RemoveThisTagItem(sReportFormatTag, cboMainReports) Then
                GoTo NEXT_ITEM
            End If
            '2 Carrier Specific Reports:
'            If RemoveThisTagItem(sReportFormatTag, cboCarSpecReports) Then
'                GoTo NEXT_ITEM
'            End If
            '3 Bills Posted:
'            If RemoveThisTagItem(sReportFormatTag, cboIB) Then
'                GoTo NEXT_ITEM
'            End If
            '4 Payment Requests:
'            If RemoveThisTagItem(sReportFormatTag, cboPayments) Then
'                GoTo NEXT_ITEM
'            End If
            '5 Attachments:
'            If RemoveThisTagItem(sReportFormatTag, cboAttachments) Then
'                GoTo NEXT_ITEM
'            End If
        End If
NEXT_ITEM:
    Next
    
    RemoveAddedReports = True
    Exit Function
EH:
    
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function RemoveAddedReports"
End Function

Private Function RemoveThisTagItem(psTagItem As String, pocboItem As Object) As Boolean
    On Error GoTo EH
    Dim lPos As Long
    Dim MyCbo As ListBox
    Dim sTagItem As String
    Dim sThisItem As String
    
    If Not TypeOf pocboItem Is ListBox Then
        Exit Function
    End If
    
    Set MyCbo = pocboItem
    sTagItem = psTagItem
    
    sTagItem = Mid(sTagItem, InStr(1, sTagItem, String(200, " "), vbBinaryCompare))
    sTagItem = LTrim(sTagItem)
    
    For lPos = 0 To MyCbo.ListCount - 1
        sThisItem = MyCbo.List(lPos)
        If InStr(1, sThisItem, sTagItem, vbTextCompare) > 0 Then
            MyCbo.RemoveItem lPos
            RemoveThisTagItem = True
            Exit Function
        End If
    Next
    
    RemoveThisTagItem = False
    
    Set MyCbo = Nothing
        
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Function RemoveThisTagItem"
End Function

Public Function PrintReportsFromLVW(poListView As Object, Optional pbSelectedOnly As Boolean = True) As Boolean
    On Error GoTo EH
    Dim sCaption As String
    Dim sIBNUM As String
    Dim sCLIENTNUM As String
    Dim sPDFFileName As String
    Dim sPDFFilePath As String
    Dim itmX As ListItem
    Dim oListView As ListView
    Dim RS As ADODB.Recordset
    
    If TypeOf poListView Is ListView Then
        Set oListView = poListView
    Else
        Exit Function
    End If
    
    Screen.MousePointer = vbHourglass
    
    'GEt ref to Assignments Record
    
    Set RS = mfrmClaim.adoRSAssignments
    
    'Loop through the Selected Items and Send them to Adobe Viewer
    'Adjuster Must have Adobe Viewer Installed
    
    For Each itmX In oListView.ListItems
        If itmX.Selected Then
            sIBNUM = goUtil.IsNullIsVbNullString(RS.Fields("IBNUM"))
            sCLIENTNUM = goUtil.IsNullIsVbNullString(RS.Fields("CLIENTNUM"))
            sCaption = "Attachment - " & itmX.SubItems(GuiAttachListView.AttachName - 1) ' & Chr(160) & " "
            sCaption = sCaption & "(" & sIBNUM & "_" & sCLIENTNUM & ")"
            
            sPDFFileName = itmX.SubItems(GuiAttachListView.Attachment - 1)
            sPDFFilePath = goUtil.gsInstallDir & "\AttachRepos\" & sPDFFileName
            'Need to shell the PDF Loss Report to Adobe Reader
            goUtil.utShellExecute , "OPEN", sPDFFilePath, , , vbNormalFocus, True, True, True, sCaption
            
        End If
    Next

    Screen.MousePointer = vbDefault
    PrintReportsFromLVW = True
    
    'clean up
    Set RS = Nothing
    Set itmX = Nothing
    Set oListView = Nothing
    
    Exit Function
EH:
    Screen.MousePointer = vbNormal
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function PrintReportsFromLVW"
End Function

Public Sub RefreshPackage()
    On Error GoTo EH
    Dim sData As String
    Dim RS As ADODB.Recordset
    Dim lCount As Long
    Dim sIBID As String
    Dim sSupplement As String
    Dim sReportTitle As String
    
'    'Load Billing RS
'    mfrmClaim.SetadoRSBillingCount msAssignmentsID, , , True
'    cboBillingID.Clear
'    cboBillingID.AddItem "(--Select Billing--)"
'    '0 indicates Null ID since ID must be >= 1 or <= -1
'    '>=1    : WEB Server Synched
'    '<=-1   : Client has yet to synch Data ID with Web Server.
'    cboBillingID.ItemData(cboBillingID.NewIndex) = 0
'    mfrmClaim.PopulateLookUp mfrmClaim.adoRSBillingCount, _
'                        Nothing, _
'                        cboBillingID, _
'                        "ID", _
'                        vbNullString, _
'                        "IB", _
'                        "IBDescription", , , True, "IBDescription2"
'    'Need to Add some Data items to Text so that
'    'the IB can be printed... Wether it is Closed or Current
'    'this contains the Software info needed to Print IB
'    mfrmClaim.SetadoRSBillingReports ' software for Bills
'    mfrmClaim.SetadoRSBillingReportsHistory 'software history for Bills
'    Set RS = mfrmClaim.adoRSBillingReports
'
'    If RS.RecordCount = 1 Then
'        RS.MoveFirst
'        For lCount = 1 To cboBillingID.ListCount - 1
'            sData = cboBillingID.List(lCount)
'            sReportTitle = Trim(left(sData, InStr(1, sData, "-", vbTextCompare) - 1)) & ")"
'            If InStr(1, sData, "Current", vbTextCompare) > 0 Then
'                'Set the IBID to "" if this is Current Billing
'                'that way the IB Report will use the RTTable instead of the
'                'Closed IB Table
'                sIBID = vbNullString
'                sSupplement = mfrmClaim.GetSupplement(cboBillingID.ItemData(lCount))
'            Else
'                sIBID = mfrmClaim.GetIBID(cboBillingID.ItemData(lCount))
'                sSupplement = mfrmClaim.GetSupplement(cboBillingID.ItemData(lCount))
'            End If
'            sData = sData & String(200, " ")
'            sData = sData & RS.Fields("ProjectName").Value & "|" & RS.Fields("ClassName").Value & "|" & RS.Fields("Version").Value
'            sData = sData & "|" & "pIBID=" & sIBID
'            sData = sData & "|" & "pSupplement=" & sSupplement
'            sData = sData & "|" & "psReportTitle=" & sReportTitle
'            'Add the Software Data to this IB item
'            cboBillingID.List(lCount) = sData
'        Next
'    Else
'        Set RS = mfrmClaim.adoRSBillingReportsHistory
'        If RS.RecordCount > 0 Then
'            RS.MoveFirst
'            For lCount = 1 To cboBillingID.ListCount - 1
'                sData = cboBillingID.List(lCount)
'                sReportTitle = Trim(left(sData, InStr(1, sData, "-", vbTextCompare) - 1)) & ")"
'                If InStr(1, sData, "Current", vbTextCompare) > 0 Then
'                    'Set the IBID to "" if this is Current Billing
'                    'that way the IB Report will use the RTTable instead of the
'                    'Closed IB Table
'                    sIBID = vbNullString
'                    sSupplement = mfrmClaim.GetSupplement(cboBillingID.ItemData(lCount))
'                Else
'                    sIBID = mfrmClaim.GetIBID(cboBillingID.ItemData(lCount))
'                    sSupplement = mfrmClaim.GetSupplement(cboBillingID.ItemData(lCount))
'                End If
'                sData = sData & String(200, " ")
'                sData = sData & RS.Fields("ProjectName").Value & "|" & RS.Fields("ClassName").Value & "|" & RS.Fields("Version").Value
'                sData = sData & "|" & "pIBID=" & sIBID
'                sData = sData & "|" & "pSupplement=" & sSupplement
'                sData = sData & "|" & "psReportTitle=" & sReportTitle
'                'Add the Software Data to this IB item
'                cboBillingID.List(lCount) = sData
'            Next
'        End If
'    End If
'
'    'select the first Element
'    cboBillingID.ListIndex = 0
    
    'cleanup
    
    Set RS = Nothing
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub RefreshPackage"
End Sub

Private Sub LoadCopy()
    On Error GoTo EH
    
    cboCopy.Clear
    cboCopy.AddItem "(-ALL COPIES-)"
    cboCopy.AddItem goUtil.gsCurCarDBName & " Copy"
    cboCopy.AddItem GetSetting(goUtil.gsAppEXEName, "GENERAL", "CURRENT_COMPANY_NAME", "Company") & " Copy"
    cboCopy.AddItem "Remit Copy"
    cboCopy.AddItem "Adjuster Copy"
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function LoadCopy"
End Sub

Public Sub RefreshPackages()
    On Error GoTo EH
        
    'Load Fee Schedule RS
    mfrmClaim.SetadoRSPackageList msAssignmentsID
    cboPackage.Clear
    cboPackage.AddItem "(--Select File--)"
    '0 indicates Null ID since ID must be >= 1 or <= -1
    '>=1    : WEB Server Synched
    '<=-1   : Client has yet to synch Data ID with Web Server.
    cboPackage.ItemData(cboPackage.NewIndex) = 0
    mfrmClaim.PopulateLookUp mfrmClaim.adoRSPackageList, _
                        Nothing, _
                        cboPackage, _
                        "ID", _
                        vbNullString, _
                        "Name", _
                        "Description"
                        
    cboPackage.ListIndex = 0
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub RefreshPackages"
End Sub

Private Sub Timer_MaximizeMe_Timer()
    On Error GoTo EH
    Timer_MaximizeMe.Enabled = False
    MaxMinPackageMaint True
    Me.WindowState = VBRUN.FormWindowStateConstants.vbMaximized
    
    
    Exit Sub
EH:
    'do nothing
End Sub

Private Sub txtAdminComments_DblClick()
    On Error GoTo EH
    Dim lRet As Long
    Dim sData As String
    Dim sTickCount As String
    Dim sFileName As String
    
    
    sTickCount = goUtil.utGetTickCount
    sData = txtAdminComments.Text
    sFileName = goUtil.AttachReposPath & "\Comments_" & sTickCount & ".txt"
    
    goUtil.utSaveFileData sFileName, sData
    
    lRet = goUtil.utShellExecute(GetDesktopWindow, "OPEN", sFileName, vbNullString, App.Path, vbNormalFocus, False, False, True)
    
    Sleep 1000
    
    goUtil.utDeleteFile sFileName
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub RefreshPackages"
End Sub

Private Sub txtAdminComments_GotFocus()
    goUtil.utSelText txtAdminComments
End Sub

Private Sub txtPassWord_GotFocus(Index As Integer)
    goUtil.utSelText txtPassWord(Index)
End Sub

Public Function AddPackageItem(pudtPackageItem As GuiPackageItem) As Boolean
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    Dim sID As String
    Dim lRecordsAffected As Long
    
    Screen.MousePointer = MousePointerConstants.vbHourglass
    
    sID = goUtil.GetAccessDBUID("ID", "PackageItem")
    
    With pudtPackageItem
        .PackageItemID = sID
        .ID = sID
    End With
    
    sSQL = "INSERT INTO PackageItem "
    sSQL = sSQL & "( "
    sSQL = sSQL & "[PackageItemID], "
    sSQL = sSQL & "[PackageID], "
    sSQL = sSQL & "[AssignmentsID], "
    sSQL = sSQL & "[ID], "
    sSQL = sSQL & "[IDPackage], "
    sSQL = sSQL & "[IDAssignments], "
    sSQL = sSQL & "[ReportFormat], "
    sSQL = sSQL & "[RTAttachmentsID], "
    sSQL = sSQL & "[IDRTAttachments], "
    sSQL = sSQL & "[Number], "
    sSQL = sSQL & "[AttachmentName], "
    sSQL = sSQL & "[SortOrder], "
    sSQL = sSQL & "[Name], "
    sSQL = sSQL & "[Description], "
    sSQL = sSQL & "[IsCoApprove], "
    sSQL = sSQL & "[CoApproveDate], "
    sSQL = sSQL & "[CoApproveDesc], "
    sSQL = sSQL & "[IsClientCoReject], "
    sSQL = sSQL & "[ClientCoRejectDate], "
    sSQL = sSQL & "[ClientCoRejectDesc], "
    sSQL = sSQL & "[IsClientCoDelete], "
    sSQL = sSQL & "[ClientCoDeleteDate], "
    sSQL = sSQL & "[ClientCoDeleteDesc], "
    sSQL = sSQL & "[IsClientCoApprove], "
    sSQL = sSQL & "[ClientCoApproveDate], "
    sSQL = sSQL & "[ClientCoApproveDesc], "
    sSQL = sSQL & "[PackageItemGUID], "
    sSQL = sSQL & "[SendMe], "
    sSQL = sSQL & "[SentDate], "
    sSQL = sSQL & "[IsDeleted], "
    sSQL = sSQL & "[DownLoadMe], "
    sSQL = sSQL & "[UpLoadMe], "
    sSQL = sSQL & "[AdminComments], "
    sSQL = sSQL & "[DateLastUpdated], "
    sSQL = sSQL & "[UpdateByUserID] "
    sSQL = sSQL & ") "
    sSQL = sSQL & "SELECT "
    sSQL = sSQL & pudtPackageItem.PackageItemID & " As [PackageItemID], "
    sSQL = sSQL & pudtPackageItem.PackageID & " As [PackageID], "
    sSQL = sSQL & pudtPackageItem.AssignmentsID & " As [AssignmentsID], "
    sSQL = sSQL & pudtPackageItem.ID & " As [ID], "
    sSQL = sSQL & pudtPackageItem.IDPackage & " As [IDPackage], "
    sSQL = sSQL & pudtPackageItem.IDAssignments & " As [IDAssignments], "
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtPackageItem.ReportFormat) & "' As [ReportFormat], "
    sSQL = sSQL & pudtPackageItem.RTAttachmentsID & " As [RTAttachmentsID], "
    sSQL = sSQL & pudtPackageItem.IDRTAttachments & " As [IDRTAttachments], "
    sSQL = sSQL & pudtPackageItem.Number & " As [Number], "
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtPackageItem.AttachmentName) & "' As [AttachmentName], "
    sSQL = sSQL & pudtPackageItem.SortOrder & " As [SortOrder], "
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtPackageItem.Name) & "' As [Name], "
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtPackageItem.Description) & "' As [Description], "
    sSQL = sSQL & pudtPackageItem.IsCoApprove & " As [IsCoApprove], "
    If StrComp(pudtPackageItem.CoApproveDate, "Null", vbTextCompare) = 0 Or pudtPackageItem.CoApproveDate = vbNullString Then
        sSQL = sSQL & "Null As [CoApproveDate], "
    Else
        sSQL = sSQL & "#" & pudtPackageItem.CoApproveDate & "# As [CoApproveDate], "
    End If
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtPackageItem.CoApproveDesc) & "' As [CoApproveDesc], "
    sSQL = sSQL & pudtPackageItem.IsClientCoReject & " As [IsClientCoReject], "
    If StrComp(pudtPackageItem.ClientCoRejectDate, "Null", vbTextCompare) = 0 Or pudtPackageItem.ClientCoRejectDate = vbNullString Then
        sSQL = sSQL & "Null As [ClientCoRejectDate], "
    Else
        sSQL = sSQL & "#" & pudtPackageItem.ClientCoRejectDate & "# As [ClientCoRejectDate], "
    End If
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtPackageItem.ClientCoRejectDesc) & "' As [ClientCoRejectDesc], "
    sSQL = sSQL & pudtPackageItem.IsClientCoDelete & " As [IsClientCoDelete], "
    If StrComp(pudtPackageItem.ClientCoDeleteDate, "Null", vbTextCompare) = 0 Or pudtPackageItem.ClientCoDeleteDate = vbNullString Then
        sSQL = sSQL & "Null As [ClientCoDeleteDate], "
    Else
        sSQL = sSQL & "#" & pudtPackageItem.ClientCoDeleteDate & "# As [ClientCoDeleteDate], "
    End If
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtPackageItem.ClientCoDeleteDesc) & "' As [ClientCoDeleteDesc], "
    sSQL = sSQL & pudtPackageItem.IsClientCoApprove & " As [IsClientCoApprove], "
    If StrComp(pudtPackageItem.ClientCoApproveDate, "Null", vbTextCompare) = 0 Or pudtPackageItem.ClientCoApproveDate = vbNullString Then
        sSQL = sSQL & "Null As [ClientCoApproveDate], "
    Else
        sSQL = sSQL & "#" & pudtPackageItem.ClientCoApproveDate & "# As [ClientCoApproveDate], "
    End If
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtPackageItem.ClientCoApproveDesc) & "' As [ClientCoApproveDesc], "
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtPackageItem.PackageItemGUID) & "' As [PackageItemGUID], "
    sSQL = sSQL & pudtPackageItem.SendMe & " As [SendMe], "
    If StrComp(pudtPackageItem.SentDate, "Null", vbTextCompare) = 0 Or pudtPackageItem.SentDate = vbNullString Then
        sSQL = sSQL & "Null As [SentDate], "
    Else
        sSQL = sSQL & "#" & pudtPackageItem.SentDate & "# As [SentDate], "
    End If
    sSQL = sSQL & pudtPackageItem.IsDeleted & " As [IsDeleted], "
    sSQL = sSQL & pudtPackageItem.DownLoadMe & " As [DownLoadMe], "
    sSQL = sSQL & pudtPackageItem.UpLoadMe & " As [UpLoadMe], "
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtPackageItem.AdminComments) & "'" & " As [AdminComments], "
    sSQL = sSQL & "#" & pudtPackageItem.DateLastUpdated & "#" & " As [DateLastUpdated], "
    sSQL = sSQL & pudtPackageItem.UpdateByUserID & " As [UpdateByUserID] "
   
    Set oConn = New ADODB.Connection
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    oConn.Execute sSQL, lRecordsAffected
    
    Screen.MousePointer = MousePointerConstants.vbDefault
    AddPackageItem = CBool(lRecordsAffected)
    
    'Clean up
    Set oConn = Nothing
    Exit Function
EH:
    Screen.MousePointer = vbDefault
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function AddPackageItem"
End Function

Public Function EditPackageItem(pudtPackageItem As GuiPackageItem) As Boolean
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    
    sSQL = "UPDATE PackageItem Set "
    sSQL = sSQL & "[PackageItemID] = " & pudtPackageItem.PackageItemID & ", "
    sSQL = sSQL & "[PackageID] = " & pudtPackageItem.PackageID & ", "
    sSQL = sSQL & "[AssignmentsID] = " & pudtPackageItem.AssignmentsID & ", "
    sSQL = sSQL & "[ID] = " & pudtPackageItem.ID & ", "
    sSQL = sSQL & "[IDPackage] = " & pudtPackageItem.IDPackage & ", "
    sSQL = sSQL & "[IDAssignments] = " & pudtPackageItem.IDAssignments & ", "
    sSQL = sSQL & "[ReportFormat] = '" & goUtil.utCleanSQLString(pudtPackageItem.ReportFormat) & "', "
    sSQL = sSQL & "[RTAttachmentsID] = " & pudtPackageItem.RTAttachmentsID & ", "
    sSQL = sSQL & "[IDRTAttachments] = " & pudtPackageItem.IDRTAttachments & ", "
    sSQL = sSQL & "[Number] = " & pudtPackageItem.Number & ", "
    sSQL = sSQL & "[AttachmentName] = '" & goUtil.utCleanSQLString(pudtPackageItem.AttachmentName) & "', "
    sSQL = sSQL & "[SortOrder] = " & pudtPackageItem.SortOrder & ", "
    sSQL = sSQL & "[Name] = '" & goUtil.utCleanSQLString(pudtPackageItem.Name) & "', "
    sSQL = sSQL & "[Description] = '" & goUtil.utCleanSQLString(pudtPackageItem.Description) & "', "
    sSQL = sSQL & "[IsCoApprove] = " & pudtPackageItem.IsCoApprove & ", "
    If StrComp(pudtPackageItem.CoApproveDate, "Null", vbTextCompare) = 0 Or pudtPackageItem.CoApproveDate = vbNullString Then
        sSQL = sSQL & "[CoApproveDate] = Null, "
    Else
        sSQL = sSQL & "[CoApproveDate] = #" & pudtPackageItem.CoApproveDate & "#, "
    End If
    sSQL = sSQL & "[CoApproveDesc] = '" & goUtil.utCleanSQLString(pudtPackageItem.CoApproveDesc) & "', "
    sSQL = sSQL & "[IsClientCoReject] = " & pudtPackageItem.IsClientCoReject & ", "
    If StrComp(pudtPackageItem.ClientCoRejectDate, "Null", vbTextCompare) = 0 Or pudtPackageItem.ClientCoRejectDate = vbNullString Then
        sSQL = sSQL & "[ClientCoRejectDate] = Null, "
    Else
        sSQL = sSQL & "[ClientCoRejectDate] = #" & pudtPackageItem.ClientCoRejectDate & "#, "
    End If
    sSQL = sSQL & "[ClientCoRejectDesc] = '" & goUtil.utCleanSQLString(pudtPackageItem.ClientCoRejectDesc) & "', "
    sSQL = sSQL & "[IsClientCoDelete] = " & pudtPackageItem.IsClientCoDelete & ", "
    If StrComp(pudtPackageItem.ClientCoDeleteDate, "Null", vbTextCompare) = 0 Or pudtPackageItem.ClientCoDeleteDate = vbNullString Then
        sSQL = sSQL & "[ClientCoDeleteDate] = Null, "
    Else
        sSQL = sSQL & "[ClientCoDeleteDate] = #" & pudtPackageItem.ClientCoDeleteDate & "#, "
    End If
    sSQL = sSQL & "[ClientCoDeleteDesc] = '" & goUtil.utCleanSQLString(pudtPackageItem.ClientCoDeleteDesc) & "', "
    sSQL = sSQL & "[IsClientCoApprove] = " & pudtPackageItem.IsClientCoApprove & ", "
    If StrComp(pudtPackageItem.ClientCoApproveDate, "Null", vbTextCompare) = 0 Or pudtPackageItem.ClientCoApproveDate = vbNullString Then
        sSQL = sSQL & "[ClientCoApproveDate] = Null, "
    Else
        sSQL = sSQL & "[ClientCoApproveDate] = #" & pudtPackageItem.ClientCoApproveDate & "#, "
    End If
    sSQL = sSQL & "[ClientCoApproveDesc] = '" & goUtil.utCleanSQLString(pudtPackageItem.ClientCoApproveDesc) & "', "
    sSQL = sSQL & "[PackageItemGUID] = '" & goUtil.utCleanSQLString(pudtPackageItem.PackageItemGUID) & "', "
    sSQL = sSQL & "[SendMe] = " & pudtPackageItem.SendMe & ", "
    If StrComp(pudtPackageItem.SentDate, "Null", vbTextCompare) = 0 Or pudtPackageItem.SentDate = vbNullString Then
        sSQL = sSQL & "[SentDate] = Null, "
    Else
        sSQL = sSQL & "[SentDate] = #" & pudtPackageItem.SentDate & "#, "
    End If
    sSQL = sSQL & "[IsDeleted] = " & pudtPackageItem.IsDeleted & ", "
    sSQL = sSQL & "[DownLoadMe] = " & pudtPackageItem.DownLoadMe & ", "
    sSQL = sSQL & "[UpLoadMe] = " & pudtPackageItem.UpLoadMe & ", "
    sSQL = sSQL & "[AdminComments] = '" & goUtil.utCleanSQLString(pudtPackageItem.AdminComments) & "', "
    sSQL = sSQL & "[DateLastUpdated] = #" & pudtPackageItem.DateLastUpdated & "#, "
    sSQL = sSQL & "[UpdateByUserID]  = " & pudtPackageItem.UpdateByUserID & " "
    sSQL = sSQL & "WHERE [IDAssignments] = " & pudtPackageItem.IDAssignments & " "
    sSQL = sSQL & "AND [ID] = " & pudtPackageItem.ID & " "
    sSQL = sSQL & "AND [IDPackage] = " & pudtPackageItem.IDPackage & " "

    Set oConn = New ADODB.Connection
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    oConn.Execute sSQL
    
    EditPackageItem = True
    'Clean up
    Set oConn = Nothing
    Exit Function
    
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function EditPackageItem"
End Function

