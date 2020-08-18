VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPhotos 
   AutoRedraw      =   -1  'True
   Caption         =   "Photos"
   ClientHeight    =   6705
   ClientLeft      =   60
   ClientTop       =   345
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
   Icon            =   "frmPhotos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6705
   ScaleWidth      =   11820
   StartUpPosition =   3  'Windows Default
   Tag             =   "Photos"
   Begin VB.Frame framCommands 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7245
      TabIndex        =   21
      Top             =   5400
      Width           =   4455
      Begin VB.CommandButton cmdPrintList 
         Caption         =   "Sho&w Item List"
         Height          =   855
         Left            =   1200
         MaskColor       =   &H00000000&
         Picture         =   "frmPhotos.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Exit"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   855
         Left            =   3360
         MaskColor       =   &H00000000&
         Picture         =   "frmPhotos.frx":044E
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Exit"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Sa&ve"
         Enabled         =   0   'False
         Height          =   855
         Left            =   2280
         MaskColor       =   &H00000000&
         Picture         =   "frmPhotos.frx":0758
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Exit"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdSpelling 
         Caption         =   "Spellin&g"
         Height          =   855
         Left            =   120
         MaskColor       =   &H00000000&
         Picture         =   "frmPhotos.frx":08A2
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Exit"
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame framAssocBillingID 
      Caption         =   "Associate IB (Internal Billing) to selected Items:"
      Height          =   1215
      Left            =   120
      TabIndex        =   18
      Top             =   5400
      Width           =   6735
      Begin VB.ComboBox cboAssocBillingID 
         Height          =   360
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   720
         Width           =   5055
      End
      Begin VB.CommandButton cmdAssocBillingID 
         Caption         =   "Associate I&B"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5280
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   713
         Width           =   1335
      End
   End
   Begin VB.Frame framPhotos 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   11580
      Begin VB.CheckBox chkPhotoView 
         Caption         =   "Ico&n View"
         Height          =   975
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "View Thumbs"
         Top             =   2400
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CommandButton CmdReNumberSort 
         Caption         =   "&Save Sort Order"
         Enabled         =   0   'False
         Height          =   975
         Left            =   120
         TabIndex        =   11
         ToolTipText     =   "Renumber and Save"
         Top             =   1440
         Width           =   735
      End
      Begin VB.CommandButton cmdDown 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   120
         Picture         =   "frmPhotos.frx":0CEC
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Move Selected Photo DOWN"
         Top             =   840
         Width           =   720
      End
      Begin VB.CommandButton cmdUp 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   120
         Picture         =   "frmPhotos.frx":112E
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Move Selected Photo UP"
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton cmdFindNext 
         Caption         =   "Find &Next"
         Height          =   375
         Left            =   4920
         TabIndex        =   5
         Top             =   240
         Width           =   1100
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Find"
         Height          =   375
         Left            =   3840
         TabIndex        =   4
         Top             =   240
         Width           =   1100
      End
      Begin VB.CommandButton cmdSelAll 
         Caption         =   "&Select A&ll"
         Height          =   375
         Left            =   2760
         TabIndex        =   3
         Top             =   240
         Width           =   1100
      End
      Begin VB.CommandButton cmdEditPhotos 
         Caption         =   "&Edit"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdAddPhotos 
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdAddMultiReport 
         Caption         =   "&Add"
         Height          =   375
         Left            =   8280
         TabIndex        =   6
         Top             =   225
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox cboPhotoReports 
         Height          =   360
         Left            =   9000
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   240
         Width           =   2415
      End
      Begin MSComctlLib.ImageList imgListPhotos 
         Left            =   1080
         Top             =   1320
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         MaskColor       =   12632256
         _Version        =   393216
      End
      Begin MSComctlLib.ImageList imgPhotoStatus 
         Left            =   1680
         Top             =   1320
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483633
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   16777215
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   6
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPhotos.frx":1570
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPhotos.frx":1968
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPhotos.frx":1D32
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPhotos.frx":1E8C
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPhotos.frx":2278
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPhotos.frx":2587
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.TextBox txtSpellMe 
         Height          =   1335
         Left            =   1080
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   1320
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.Frame framPhotoMaint 
         Caption         =   "Photo Maintenance"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   13
         Top             =   4440
         Width           =   11355
         Begin VB.CommandButton cmdPrintPhotos 
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
            Left            =   1080
            TabIndex        =   15
            Top             =   240
            Width           =   975
         End
         Begin VB.CheckBox chkHideDeleted 
            Caption         =   "Sho&w Deleted"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   8520
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   240
            Visible         =   0   'False
            Width           =   1100
         End
         Begin VB.CommandButton cmdDelPhotos 
            Caption         =   "&Delete"
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
            Left            =   9720
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   240
            Width           =   1455
         End
         Begin VB.CommandButton cmdRefreshPhotos 
            Caption         =   "&Refresh"
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
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   975
         End
      End
      Begin MSComctlLib.ListView lstvPhotos 
         Height          =   3855
         Left            =   840
         TabIndex        =   8
         Tag             =   "Enable"
         ToolTipText     =   "Right Click for Menu"
         Top             =   600
         Width           =   10635
         _ExtentX        =   18759
         _ExtentY        =   6800
         Arrange         =   2
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         OLEDragMode     =   1
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ColHdrIcons     =   "imgPhotoStatus"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OLEDragMode     =   1
         NumItems        =   0
      End
   End
   Begin VB.Menu PopupMnuPhoto 
      Caption         =   "PopUpPhoto"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuEditPhoto 
         Caption         =   "&Edit"
      End
      Begin VB.Menu mnuDeletePhoto 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuSelectAllPhoto 
         Caption         =   "&Select All"
      End
   End
End
Attribute VB_Name = "frmPhotos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mfrmClaim As frmClaim
Private mbUnloadMe As Boolean
Private msAssignmentsID As String
Private msIDRTPhotoReport As String
Private msIBNUM As String
Private moGUI As V2ECKeyBoard.clsCarGUI
Private mitmXSelected As ListItem 'Currently selected Photo Item
Private mbLoading As Boolean 'Loading Form
Private mbLoadingMe As Boolean 'Just loading data
Private msFindText As String
Private mlLastFindIndex As Long
Private mArv As V2ARViewer.clsARViewer
Private moForm As Form
Private mbRenumberSort As Boolean
Private msRenumSortBillingCountID As String
Private msRenumSortIDBillingCount As String


Public Property Let itmXSelected(pitmX As ListItem)
    On Error GoTo EH
    Set mitmXSelected = pitmX
    If Not mitmXSelected Is Nothing Then
        cmdEditPhotos.Enabled = True
    Else
        cmdEditPhotos.Enabled = False
    End If
    Exit Property
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Property Let itmXSelected"
End Property
Public Property Set itmXSelected(pitmX As ListItem)
    On Error GoTo EH
    Set mitmXSelected = pitmX
    If Not mitmXSelected Is Nothing Then
        cmdEditPhotos.Enabled = True
    Else
        cmdEditPhotos.Enabled = False
    End If
    Exit Property
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Property Set itmXSelected"
End Property
Public Property Get itmXSelected() As ListItem
    Set itmXSelected = mitmXSelected
End Property

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

Public Property Let IBNUM(psIBNUM As String)
    msIBNUM = psIBNUM
End Property
Public Property Get IBNUM() As String
    IBNUM = msIBNUM
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

Public Property Get MAX_PHOTOS_ALLOWED() As Long
    MAX_PHOTOS_ALLOWED = 20
End Property

Private Property Get msClassName() As String
    msClassName = Me.Name
End Property

Private Sub cboAssocBillingID_Click()
    On Error GoTo EH
    
    If cboAssocBillingID.ListIndex > -1 Then
        cmdAssocBillingID.Enabled = True
    Else
        cmdAssocBillingID.Enabled = False
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cboAssocBillingID_Click"
End Sub




Private Sub cboPhotoReports_Click()
    On Error GoTo EH
    
    If cboPhotoReports.ListIndex = -1 Or mbLoadingMe Then
        Exit Sub
    End If
    
    msIDRTPhotoReport = cboPhotoReports.ItemData(cboPhotoReports.ListIndex)
    
    LoadMe
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub chkHideDeleted_Click"
End Sub

Private Sub chkHideDeleted_Click()
    On Error GoTo EH
    Dim bHideDeleted As Boolean
    
    If chkHideDeleted.Value = vbChecked Then
        chkHideDeleted.Caption = "&Hide Deleted"
        bHideDeleted = True
    Else
        chkHideDeleted.Caption = "Sho&w Deleted"
        bHideDeleted = False
    End If
    
    SaveSetting App.EXEName, "GENERAL", "HIDE_DELETED", bHideDeleted
    If Not mbLoading Then
        LoadMe
    End If
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub chkHideDeleted_Click"
End Sub


Private Sub chkPhotoView_Click()
    On Error GoTo EH
    lstvPhotos.Visible = False
    SaveSetting App.EXEName, "GENERAL", "PHOTO_PhotoView", chkPhotoView.Value
    If chkPhotoView.Value = vbChecked Then
        chkPhotoView.Caption = "Ico&n View"
        lstvPhotos.View = lvwIcon
        framAssocBillingID.Enabled = False
        cboAssocBillingID.Enabled = False
        cmdAssocBillingID.Enabled = False
    Else
        chkPhotoView.Caption = "Repor&t View"
        lstvPhotos.View = lvwReport
        framAssocBillingID.Enabled = True
        cboAssocBillingID.Enabled = True
        cmdAssocBillingID.Enabled = True
    End If
    lstvPhotos.Visible = True
    Exit Sub
EH:
    lstvPhotos.Visible = True
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub chkPhotoView_Click"
End Sub

Private Sub cmdAddMultiReport_Click()
    On Error GoTo EH
    Dim MyfrmAddMultiReportItem As AddMultiReportItem
    
    Set MyfrmAddMultiReportItem = New AddMultiReportItem
    
    With MyfrmAddMultiReportItem
        .MyfrmClaim = mfrmClaim
        .AssignmentsID = msAssignmentsID
        .TableName = "RTPhotoReport"
    End With
    
    
    Load MyfrmAddMultiReportItem
    
    MyfrmAddMultiReportItem.Show vbModal
    
    MyfrmAddMultiReportItem.CLEANUP
    
    Unload MyfrmAddMultiReportItem
    
    Set MyfrmAddMultiReportItem = Nothing
    
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdAddMultiReport_Click"
End Sub

Private Sub cmdAddPhotos_Click()
    On Error GoTo EH
'    If cboPhotoReports.ListIndex = -1 Then
'        MsgBox "You must add a Photo Report in Reports section first!", vbOKOnly + vbExclamation, "Create a Photo Report!"
'        Exit Sub
'    End If
    
    'BGS Making GUI Changes to Photos 1.10.2005 Per Rob Petrovics and Elizabeth Warner-Simpson Request
    'MAde the Add button invisible, if for some reason need to go back to Add button just make the
    'Add Button Visible again.
    If cmdAddMultiReport.Visible = False Then
        'Need to select the Last Photo Report created.
        'Start adding photos to that Report.
        If cboPhotoReports.ListCount = 1 Then
            'If there are no photo reports yet... need to add the first one.
            AddNextPhotoReport
        ElseIf cboPhotoReports.ListIndex = 0 Then
            AddNextPhotoReport
        End If
        
    End If
    
    cmdAddPhotos.Enabled = False
    With AddPhoto
        .MyPhotos = Me
        .MyfrmClaim = Me.MyfrmClaim
        Load AddPhoto
        .AssignmentsID = msAssignmentsID
        .IDRTPhotoReport = msIDRTPhotoReport
        .IBNUM = msIBNUM
        .Adding = True
        .Caption = "Photo Add"
        .MaxSort = GetMaxSort(msAssignmentsID, msIDRTPhotoReport)
        .PhotoCount = GetPhotoCount(msAssignmentsID, msIDRTPhotoReport)
        .cmdLoadAll.Enabled = True
        .cmdMenu(2).Enabled = False
        .txtPhotoDate.Text = Format(Now, "MM/DD/YYYY")
        .Show vbModal
    End With
    
    'Clean Up
    Unload AddPhoto
    Set AddPhoto = Nothing
    cmdAddPhotos.Enabled = True
    
    If lstvPhotos.Visible Then
        lstvPhotos.SetFocus
    End If
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdAddPhotos_Click"
End Sub

Private Function CheckNextReportID() As Long
    On Error GoTo EH
    Dim sReportID As String
    Dim lPhotoCount As Long
    Dim lCount As Long
    
    CheckNextReportID = 0
    lCount = cboPhotoReports.ListIndex + 1
    Do Until lCount > cboPhotoReports.ListCount
        sReportID = cboPhotoReports.ItemData(lCount - 1)
        lPhotoCount = GetPhotoCount(msAssignmentsID, sReportID)
        If lPhotoCount < MAX_PHOTOS_ALLOWED And sReportID <> "0" Then
            CheckNextReportID = cboPhotoReports.ItemData(lCount - 1)
            cboPhotoReports.ListIndex = lCount - 1
            Exit Do
        End If
        lCount = lCount + 1
    Loop
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Function CheckNextReportID"
End Function

Public Function AddNextPhotoReport(Optional plIDRTPhotoReport As Long) As Boolean
    On Error GoTo EH
    Dim RS As ADODB.Recordset
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    Dim sID As String
    Dim sTableName As String
    Dim sName As String
    Dim sDescription As String
    Dim sNumber As String
    Dim sBeginNum As String
    Dim sEndNum As String
    Dim lNextReportID As Long
    
    
    'First Check next report.
    'If there is a next report and that report has Room to add more photos...
    'Then select that report
    lNextReportID = CheckNextReportID()
    If lNextReportID <> 0 Then
        GoTo SET_IDRTPHOTOREPORT
    End If
    
    'Set the Table name for Photo Reports
    sTableName = "RTPhotoReport"
    
    Set oConn = New ADODB.Connection
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    'If adding new then
    sID = goUtil.GetAccessDBUID("ID", sTableName)
    
    'Need to get the Max number
    sSQL = "SELECT   MAX([Number]) + 1 As [Number] "
    sSQL = sSQL & "FROM     " & sTableName & " "
    sSQL = sSQL & "WHERE    [IDAssignments] = " & msAssignmentsID & " "
    
    Set RS = New ADODB.Recordset
    'Use Disconnected Record Set on asUseClient Cusor ONLY !
    RS.CursorLocation = adUseClient
    RS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set RS.ActiveConnection = Nothing
    
    If RS.RecordCount = 1 Then
        sNumber = goUtil.IsNullIsVbNullString(RS.Fields("Number"))
        If sNumber = vbNullString Or sNumber = 0 Then
            sNumber = "1"
        End If
    Else
        sNumber = "1"
    End If
    
    sName = "PhotoReport" & Format(sNumber, "000")
    sEndNum = CLng(sNumber) * MAX_PHOTOS_ALLOWED
    sBeginNum = (CLng(sEndNum) - MAX_PHOTOS_ALLOWED) + 1
    sDescription = Format(sBeginNum, "000") & " - " & Format(sEndNum, "000")
    
    mfrmClaim.AddMultiReport "RTPhotoReport", sName, sDescription, CLng(sNumber)
    
    cboPhotoReports.ListIndex = cboPhotoReports.ListCount - 1
    
SET_IDRTPHOTOREPORT:
    plIDRTPhotoReport = cboPhotoReports.ItemData(cboPhotoReports.ListIndex)
    
    AddNextPhotoReport = True
    
    'cleanup
    Set RS = Nothing
    Set oConn = Nothing
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function AddNextPhotoReport"
End Function

Private Sub cmdAssocBillingID_Click()
    On Error GoTo EH
    Dim sMess As String
    Dim sIBDesc As String
    Dim sTitle As String
    Dim bItemSelected As Boolean
    Dim itmX As MSComctlLib.ListItem
    Dim sPhotoID As String
    Dim vPhotoID As Variant
    Dim colPhotoID As Collection
    
    
    If lstvPhotos.ListItems.Count > 0 Then
        
        If cboAssocBillingID.ListIndex = -1 Then
            sMess = "You must select an IB from the Drop down List!"
            MsgBox sMess, vbExclamation + vbOKOnly, "Select an IB."
            cmdAssocBillingID.Enabled = False
            Exit Sub
        End If
        
        For Each itmX In lstvPhotos.ListItems
            If itmX.Selected Then
                bItemSelected = True
                Exit For
            End If
        Next
        
        'See if there is a selected item
        If Not bItemSelected Then
            sMess = "You must select at least one item from the View!"
            MsgBox sMess, vbExclamation + vbOKOnly, "Select at least one item."
            Exit Sub
        End If
    
        
        sIBDesc = cboAssocBillingID.Text
        'If the selected Billing is Closed thenGive message
        If InStr(1, sIBDesc, "Closed", vbTextCompare) > 0 Then
            sMess = "The selected IB is CLOSED." & vbCrLf & vbCrLf
            sMess = sMess & "Are you sure you really want to associated the selected item(s)" & vbCrLf
            sMess = sMess & "to this CLOSED IB?  If you do, you will have to Rebill the IB" & vbCrLf
            sMess = sMess & "and any IB(s) that are associated with the selected item(s) and calculate again." & vbCrLf & vbCrLf
            sMess = sMess & "Click ""OK"" to associcate these items to the CLOSED IB." & vbCrLf
            sMess = sMess & "Click ""CANCEL"" to abort this action."
            sTitle = "Associate Items to CLOSED IB"
        ElseIf InStr(1, sIBDesc, "(--Disassociate Billing--)", vbTextCompare) > 0 Then
            sMess = "Are you sure you really want to disassociate the selected item(s)" & vbCrLf
            sMess = sMess & "If you do, you will have to Rebill" & vbCrLf
            sMess = sMess & "any IB(s) that are associated with the selected item(s) and calculate again." & vbCrLf & vbCrLf
            sMess = sMess & "Click ""OK"" to disassociate these items " & vbCrLf
            sMess = sMess & "Click ""CANCEL"" to abort this action."
            sTitle = "Disassociate Items"
        ElseIf InStr(1, sIBDesc, "Curent", vbTextCompare) > 0 Then
            sMess = "Are you sure you want to associate the selected item(s)" & vbCrLf
            sMess = sMess & "to the CURRENT IB?" & vbCrLf & vbCrLf
            sMess = sMess & "Click ""OK"" to associate these items " & vbCrLf
            sMess = sMess & "Click ""CANCEL"" to abort this action."
            sTitle = "Associate Items to CURRENT IB"
        End If
        
        If sMess <> vbNullString Then
            If MsgBox(sMess, vbInformation + vbOKCancel, sTitle) = vbCancel Then
                Exit Sub
            End If
        End If
        
        lstvPhotos.Visible = False
        Set colPhotoID = New Collection
        For Each itmX In lstvPhotos.ListItems
            If itmX.Selected Then
                colPhotoID.Add itmX.SubItems(GuiPhotoListView.ID - 1), itmX.SubItems(GuiPhotoListView.ID - 1)
            End If
        Next
        For Each vPhotoID In colPhotoID
            sPhotoID = vPhotoID
            If Not AssocPhotoItemToBillingID(sPhotoID) Then
                Exit Sub
            End If
        Next
    End If
    
    RefreshPhotos
    
    lstvPhotos.Visible = True
    cmdAssocBillingID.Enabled = True
    
    Set itmX = Nothing
    Set colPhotoID = Nothing
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdAssocBillingID_Click"
End Sub

Private Sub cmdDelPhotos_Click()
    On Error GoTo EH
    Dim itmX As MSComctlLib.ListItem
    Dim sPhotoID As String
    Dim vPhotoID As Variant
    Dim colPhotoID As Collection
    
    
    If lstvPhotos.ListItems.Count > 0 Then
        If MsgBox("Are you sure ?", vbYesNo, "DELETE SELECTED PHOTOS") = vbYes Then
            lstvPhotos.Visible = False
            Set colPhotoID = New Collection
            For Each itmX In lstvPhotos.ListItems
                If itmX.Selected Then
                    colPhotoID.Add itmX.SubItems(GuiPhotoListView.ID - 1), itmX.SubItems(GuiPhotoListView.ID - 1)
                End If
            Next
            For Each vPhotoID In colPhotoID
                sPhotoID = vPhotoID
                If DeletePhotoItem(sPhotoID) Then
                    lstvPhotos.ListItems.Remove ("""" & sPhotoID & """")
                End If
            Next
        End If
    End If
    lstvPhotos.Visible = True
    Set itmX = Nothing
    Set colPhotoID = Nothing
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdDelPhotos_Click"
End Sub

Private Sub cmdDown_Click()
    goUtil.utMoveListItem lstvPhotos, MoveDown
End Sub

Private Sub cmdEditPhotos_Click()
    On Error GoTo EH
    cmdEditPhotos.Enabled = False
    EditPhoto
    cmdEditPhotos.Enabled = True
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdEditPhotos_Click"
End Sub

Private Sub cmdExit_Click()
    On Error GoTo EH
    
    mbUnloadMe = True
    Me.Visible = False
    mfrmClaim.Timer_UnloadForm.Enabled = True
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdExit_Click"
End Sub

Private Sub cmdFind_Click()
    On Error GoTo EH
    If lstvPhotos.ListItems.Count > 0 Then
        mlLastFindIndex = 0
        goUtil.utFindListItem Me, lstvPhotos, msFindText, mlLastFindIndex
    End If
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdFind_Click"
End Sub

Private Sub cmdFindNext_Click()
    On Error GoTo EH
    
    If msFindText <> vbNullString And lstvPhotos.ListItems.Count > 0 Then
        goUtil.utFindListItem Me, lstvPhotos, msFindText, mlLastFindIndex
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdFindNext_Click"
End Sub

Private Sub cmdPrintList_Click()
    On Error GoTo EH
    goUtil.utPrintListView App.EXEName, lstvPhotos, "Photos"
     Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdPrintList_Click"
End Sub

Private Sub cmdPrintPhotos_Click()
    On Error GoTo EH
    Screen.MousePointer = vbHourglass
    cmdPrintPhotos.Enabled = False
    If PrintPhotos(msAssignmentsID, msIDRTPhotoReport) Then
        If Not mbUnloadMe Then
            cmdPrintPhotos.Enabled = True
        End If
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
EH:
    Screen.MousePointer = vbDefault
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdPrintPhotos_Click"
End Sub

Private Sub cmdRefreshPhotos_Click()
    On Error GoTo EH
    cmdRefreshPhotos.Enabled = False
    Screen.MousePointer = vbHourglass
    RefreshPhotos
    Screen.MousePointer = vbDefault
    cmdRefreshPhotos.Enabled = True
Exit Sub
EH:
    Screen.MousePointer = vbDefault
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdRefreshPhotos_Click"
End Sub

Public Sub RefreshPhotos()
    LoadMe
End Sub

Private Sub CmdReNumberSort_Click()
    On Error GoTo EH
    Dim bHideDeleted As Boolean
    Dim sMess As String
    
    bHideDeleted = CBool(GetSetting(App.EXEName, "GENERAL", "HIDE_DELETED", True))
    
    'Can't save sort while able to view deleted records...
    'so give message box indicating this
    sMess = "Can't Save Sort while ""hide deleted records on all screens"" is unchecked!" & vbCrLf & vbCrLf
    sMess = sMess & "You can check this item under the Fee Schedule."
    If Not bHideDeleted And Not mbUnloadMe Then
        MsgBox sMess, vbExclamation + vbOKOnly, "Can't Save Sort!"
        Exit Sub
    End If
    
    
    Screen.MousePointer = vbHourglass
    If CmdReNumberSort.Enabled Then
        CmdReNumberSort.Enabled = False
        ReNumberPhotoSort
        CmdReNumberSort.Enabled = True
        Screen.MousePointer = vbDefault
    End If
    Exit Sub
EH:
    Screen.MousePointer = vbDefault
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub CmdReNumberSort_Click"
End Sub

Private Sub cmdSave_Click()
    On Error GoTo EH
    If SaveMe Then
        mfrmClaim.RefreshMe
        cmdSave.Enabled = False
    End If
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdSave_Click"
End Sub

Private Sub cmdSelAll_Click()
    On Error GoTo EH
    Dim itmX As ListItem
    For Each itmX In lstvPhotos.ListItems
        itmX.Selected = True
    Next
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdSelAll_Click"
End Sub

Private Sub cmdSpelling_Click()
    On Error GoTo EH
    Dim itmX As ListItem
    Dim sText As String
    Dim saryText() As String
    Dim lPos As Long
    Dim udtPhoto As GuiPhotoItem
    
    txtSpellMe.Text = vbNullString
    
    For Each itmX In lstvPhotos.ListItems
        sText = sText & itmX.SubItems(GuiPhotoListView.Description - 1) & vbCrLf
    Next
    'take off the last VBCRLF
    If sText <> vbNullString Then
        sText = left(sText, InStrRev(sText, vbCrLf, , vbBinaryCompare) - 1)
    Else
        Exit Sub
    End If
    
    'Set the Spelling text box
    txtSpellMe.Text = sText
    
    cmdSpelling.Enabled = False
    goUtil.utLoadSP
    goUtil.goSP.CheckSP txtSpellMe
    
    'Now Get the Corrected Text into Array
    sText = txtSpellMe.Text
    saryText() = Split(sText, vbCrLf, , vbBinaryCompare)
    
    'check the spelling against the List view...
    'if any changes then need to save those changes to the db
    For lPos = LBound(saryText, 1) To UBound(saryText, 1)
        sText = saryText(lPos)
        Set itmX = lstvPhotos.ListItems(lPos + 1)
        
        If StrComp(sText, itmX.SubItems(GuiPhotoListView.Description - 1), vbTextCompare) <> 0 Then
            With udtPhoto
                .RTPhotoLogID = itmX.SubItems(GuiPhotoListView.RTPhotoLogID - 1)
                .AssignmentsID = itmX.SubItems(GuiPhotoListView.AssignmentsID - 1)
                .BillingCountID = itmX.SubItems(GuiPhotoListView.BillingCountID - 1)
                .ID = itmX.SubItems(GuiPhotoListView.ID - 1)
                .IDAssignments = itmX.SubItems(GuiPhotoListView.IDAssignments - 1)
                .IDBillingCount = itmX.SubItems(GuiPhotoListView.IDBillingCount - 1)
                .PhotoDate = itmX.SubItems(GuiPhotoListView.PhotoDate - 1)
                .SortOrder = itmX.SubItems(GuiPhotoListView.SortOrder - 1)
                'Set the Correct spelling Description.
                itmX.SubItems(GuiPhotoListView.Description - 1) = sText
                .Description = sText
                .PhotoName = itmX.SubItems(GuiPhotoListView.PhotoName - 1)
                .Photo = itmX.SubItems(GuiPhotoListView.Photo - 1)
                .DownloadPhoto = goUtil.GetFlagFromText(itmX.SubItems(GuiPhotoListView.DownloadPhoto - 1))
                .UpLoadPhoto = goUtil.GetFlagFromText(itmX.SubItems(GuiPhotoListView.UpLoadPhoto - 1))
                .PhotoThumb = itmX.SubItems(GuiPhotoListView.PhotoThumb - 1)
                .DownloadPhotoThumb = goUtil.GetFlagFromText(itmX.SubItems(GuiPhotoListView.DownloadPhotoThumb - 1))
                .UpLoadPhotoThumb = goUtil.GetFlagFromText(itmX.SubItems(GuiPhotoListView.UpLoadPhotoThumb - 1))
                ' This option is not yet supported 7.22.2004
                .PhotoHighRes = itmX.SubItems(GuiPhotoListView.PhotoHighRes - 1)
                .DownloadPhotoHighRes = goUtil.GetFlagFromText(itmX.SubItems(GuiPhotoListView.DownloadPhotoHighRes - 1))
                .UploadPhotoHighRes = goUtil.GetFlagFromText(itmX.SubItems(GuiPhotoListView.UploadPhotoHighRes - 1))
                .IsDeleted = goUtil.GetFlagFromText(itmX.SubItems(GuiPhotoListView.IsDeleted - 1))
                .DownLoadMe = goUtil.GetFlagFromText(itmX.SubItems(GuiPhotoListView.DownLoadMe - 1))
                .UpLoadMe = "True"
                .AdminComments = itmX.SubItems(GuiPhotoListView.AdminComments - 1)
                .DateLastUpdated = Format(Now(), "MM/DD/YYYY HH:MM:SS")
                .UpdateByUserID = goUtil.gsCurUsersID
            End With
            EditPhotoItem udtPhoto
        End If
    Next
    
    cmdSpelling.Enabled = True
    
    'cleanup
    Set itmX = Nothing
    txtSpellMe.Text = vbNullString
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdSpelling_Click"
End Sub

Private Sub cmdUp_Click()
    goUtil.utMoveListItem lstvPhotos, MoveUp
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
            
        Case Else
            If Not mfrmClaim Is Nothing Then
                mfrmClaim.Form_KeyDown KeyCode, Shift
            End If
    End Select
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_KeyDown"
End Sub
Private Sub Form_Load()
    On Error GoTo EH
    mbLoading = True
    'Set the Form Icon
    Me.Icon = mfrmClaim.optClaim(GuiClaimOptions.opt05_Photos).Picture
    
    LoadHeaderlstvPhotos
    Screen.MousePointer = vbHourglass
    LoadMe
    Screen.MousePointer = vbDefault
    
    CheckStatus
    
    cmdSave.Enabled = False
    mbLoading = False
    Exit Sub
EH:
    Screen.MousePointer = vbDefault
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_Load"
End Sub

Public Function LoadMe() As Boolean
    On Error GoTo EH
    Dim lPos As Long
    
    mbLoadingMe = True
    
    'Check for Selected Photo Report
    If msIDRTPhotoReport = vbNullString Then
        If cboPhotoReports.ListIndex <> -1 Then
            msIDRTPhotoReport = cboPhotoReports.ItemData(cboPhotoReports.ListIndex)
        Else
            msIDRTPhotoReport = "0"
        End If
    End If
    
    'Load Photo Reports RS
    mfrmClaim.SetadoRSRTPhotoReportList msAssignmentsID
    cboPhotoReports.Clear
    cboPhotoReports.AddItem "(--All Photos--)"
    cboPhotoReports.ItemData(cboPhotoReports.NewIndex) = 0
    mfrmClaim.PopulateLookUp mfrmClaim.adoRSRTPhotoReportList, _
                        Nothing, _
                        cboPhotoReports, _
                        "ID", _
                        vbNullString, _
                        "Name", _
                        "Description"
                        
    'Make sure the correct Photo Report is Selected
    If msIDRTPhotoReport <> "0" Then
        For lPos = 0 To cboPhotoReports.ListCount - 1
            If cboPhotoReports.ItemData(lPos) = msIDRTPhotoReport Then
                cboPhotoReports.ListIndex = lPos
                CmdReNumberSort.Enabled = False
                cmdUp.Enabled = False
                cmdDown.Enabled = False
                cmdPrintPhotos.Enabled = True
                cmdAddPhotos.Enabled = True
                cmdEditPhotos.Enabled = True
                Exit For
            End If
        Next
    Else
        cboPhotoReports.ListIndex = 0
        CmdReNumberSort.Enabled = True
        cmdUp.Enabled = True
        cmdDown.Enabled = True
'        cmdPrintPhotos.Enabled = False
'        cmdAddPhotos.Enabled = False
'        cmdEditPhotos.Enabled = False
    End If
    
    
    If Not mfrmClaim.SetadoRSRTPhotoLog(msAssignmentsID, msIDRTPhotoReport) Then
        Exit Function
    End If
    
    PopulatelstvPhotos
    
    'Load Billing RS
    mfrmClaim.SetadoRSBillingCount msAssignmentsID, , True
    cboAssocBillingID.Clear
    cboAssocBillingID.AddItem "(--Disassociate Billing--)"
    '0 indicates Null ID since ID must be >= 1 or <= -1
    '>=1    : WEB Server Synched
    '<=-1   : Client has yet to synch Data ID with Web Server.
    cboAssocBillingID.ItemData(cboAssocBillingID.NewIndex) = 0
    mfrmClaim.PopulateLookUp mfrmClaim.adoRSBillingCount, _
                        Nothing, _
                        cboAssocBillingID, _
                        "ID", _
                        vbNullString, _
                        "IB", _
                        "IBDescription", , , True, "IBDescription2"
                        
    
    LoadMe = True
    mbLoadingMe = False
    Exit Function
EH:
    mbLoadingMe = False
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function LoadMe"
End Function



Public Function SaveMe() As Boolean
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim MyadoRSAssignments As ADODB.Recordset
    Dim iCurrentStatus As V2ECKeyBoard.AssgnStatus
    Dim sSQL As String
    'If Close date is Set then be sure all the other dates are set tooooo
    Dim bCloseDateIsSet As Boolean
    ' Vars
    
    'validate all the fields on this form
    goUtil.utValidate Me
    
'    'Check for Drop Down Items Not Selected that should be
'    If cboAssignmentType.ListIndex = -1 Then
'        sMess = sMess & "Assignment Type not selected !" & vbCrLf
'    End If
'    If cboCatCode.ListIndex = -1 Then
'        sMess = sMess & "Cat Code not selected !" & vbCrLf
'    End If
'    If cboACID.ListIndex = -1 Then
'        sMess = sMess & "ACID not selected !" & vbCrLf
'    End If
'    If cboACIDDisplay.ListIndex = -1 Then
'        sMess = sMess & "ACID Display not selected !" & vbCrLf
'    End If
'    If cboMAState.ListIndex = -1 Then
'        sMess = sMess & "Mailing State not selected !" & vbCrLf
'    End If
'    If cboPAState.ListIndex = -1 Then
'        sMess = sMess & "Property State not selected !" & vbCrLf
'    End If
'    If cboTypeOfLoss.ListIndex = -1 Then
'        sMess = sMess & "Type Of Loss not selected !" & vbCrLf
'    End If
'
'    'DATES !!!!
'    'Close Date
'    If IsDate(txtCloseDate.Text) Then
'        'Check for Close date but no other dates filled out
'        bCloseDateIsSet = True
'        sCloseDate = "#" & Format(txtCloseDate.Text, "MM/DD/YYYY") & "#"
'    Else
'        sCloseDate = "null"
'    End If
'
'    'Loss Date
'    If IsDate(txtLossDate.Text) Then
'        sLossDate = "#" & Format(txtLossDate.Text, "MM/DD/YYYY") & "#"
'    Else
'        'Check for Close date but no other dates filled out
'        If bCloseDateIsSet Then
'            sMess = sMess & "Loss Date is not set!" & vbCrLf
'        End If
'        sLossDate = "null"
'    End If
'    'Assigned Date
'    If IsDate(txtAssignedDate.Text) Then
'        sAssignedDate = "#" & Format(txtAssignedDate.Text, "MM/DD/YYYY") & "#"
'    Else
'        If bCloseDateIsSet Then
'            sMess = sMess & "Assigned Date is not set!" & vbCrLf
'        End If
'        sAssignedDate = "null"
'    End If
'    'Received Date
'    If IsDate(txtReceivedDate.Text) Then
'        sReceivedDate = "#" & Format(txtReceivedDate.Text, "MM/DD/YYYY") & "#"
'    Else
'        If bCloseDateIsSet Then
'            sMess = sMess & "Received Date is not set!" & vbCrLf
'        End If
'        sReceivedDate = "null"
'    End If
'    'Contact Date
'    If IsDate(txtContactDate.Text) Then
'        sContactDate = "#" & Format(txtContactDate.Text, "MM/DD/YYYY") & "#"
'    Else
'        If bCloseDateIsSet Then
'            sMess = sMess & "Contact Date is not set!" & vbCrLf
'        End If
'        sContactDate = "null"
'    End If
'    'Inspected Date
'    If IsDate(txtInspectedDate.Text) Then
'        sInspectedDate = "#" & Format(txtInspectedDate.Text, "MM/DD/YYYY") & "#"
'    Else
'        If bCloseDateIsSet Then
'            sMess = sMess & "Inspected Date is not set!" & vbCrLf
'        End If
'        sInspectedDate = "null"
'    End If
'
'    If sMess <> vbNullString Then
'        sMess = "Could not save " & Me.Caption & vbCrLf & vbCrLf & sMess
'        MsgBox sMess, vbExclamation + vbOKOnly, "Could Not Save Claim Information."
'        Exit Function
'    End If
'
'    'Use this to check new values to be inserted
'    'against the current values in this recordset
'    Set MyadoRSAssignments = mfrmClaim.adoRSAssignments
'
'    'set the Assignemtn vars
'    sAssignmentsID = msAssignmentsID
'
'    sID = msAssignmentsID
'
'    sAssignmentTypeID = cboAssignmentType.ItemData(cboAssignmentType.ListIndex)
'
'    sClientCompanyCatSpecID = cboCatCode.ItemData(cboCatCode.ListIndex)
'
'    sAdjusterSpecID = cboACID.ItemData(cboACID.ListIndex)
'
'    sAdjusterSpecIDDisplay = cboACIDDisplay.ItemData(cboACIDDisplay.ListIndex)
'
'    sSPVersion = "[SPVersion]"
'    'IBNUM
'    sIBNUM = "'" & goUtil.utCleanSQLString(UCase(txtIBNUM.Text)) & "'"
'    'CLIENTNUM
'    sCLIENTNUM = "'" & goUtil.utCleanSQLString(UCase(txtCLIENTNUM.Text)) & "'"
'    'Policy Number
'    sPolicyNo = "'" & goUtil.utCleanSQLString(UCase(txtPolicyNo.Text)) & "'"
'    'Policty Description
'    sPolicyDescription = "'" & goUtil.utCleanSQLString(UCase(txtPolicyDescription.Text)) & "'"
'    'Insured
'    sInsured = "'" & goUtil.utCleanSQLString(UCase(txtInsured.Text)) & "'"
'
'    'Mailing Address
'    'Street
'    sMAStreet = UCase(txtMAStreet.Text)
'    'City
'    sMACity = UCase(txtMACity.Text)
'    'State
'    sMAState = left(UCase(cboMAState.Text), 2)
'    'Zip
'    sMAZIP = txtMAZIP.Text
'    'Zip4
'    sMAZIP4 = txtMAZIP4.Text
'    'Other Post Code
'    sMAOtherPostCode = UCase(txtMAOtherPostCode.Text)
'    'Build entire Address
'    If sMAZIP = "00000" & sMAZIP4 = "0000" Then
'        sMailingAddress = "'" & goUtil.utCleanSQLString(goUtil.utJoinToAddress(sMAStreet, sMACity, sMAState, sMAOtherPostCode)) & "'"
'    Else
'        sMailingAddress = "'" & goUtil.utCleanSQLString(goUtil.utJoinToAddress(sMAStreet, sMACity, sMAState, Format(sMAZIP, "00000") & "-" & Format(sMAZIP4, "0000"))) & "'"
'    End If
'    'Street
'    sMAStreet = "'" & goUtil.utCleanSQLString(UCase(txtMAStreet.Text)) & "'"
'    'City
'    sMACity = "'" & goUtil.utCleanSQLString(UCase(txtMACity.Text)) & "'"
'    'State
'    sMAState = "'" & goUtil.utCleanSQLString(left(UCase(cboMAState.Text), 2)) & "'"
'    'Zip
'    sMAZIP = txtMAZIP.Text
'    'Zip4
'    sMAZIP4 = txtMAZIP4.Text
'    'Other Post Code
'    sMAOtherPostCode = "'" & goUtil.utCleanSQLString(UCase(txtMAOtherPostCode.Text)) & "'"
'    'End Mailing Address
'
'    'Property Address
'    'Street
'    sPAStreet = UCase(txtPAStreet.Text)
'    'City
'    sPACity = UCase(txtPACity.Text)
'    'State
'    sPAState = left(UCase(cboPAState.Text), 2)
'    'Zip
'    sPAZIP = txtPAZIP.Text
'    'Zip4
'    sPAZIP4 = txtPAZIP4.Text
'    'other PostCode
'    sPAOtherPostCode = UCase(txtPAOtherPostCode.Text)
'    'Build entire Address
'    If sPAZIP = "00000" & sPAZIP4 = "0000" Then
'        sPropertyAddress = "'" & goUtil.utCleanSQLString(goUtil.utJoinToAddress(sPAStreet, sPACity, sPAState, sPAOtherPostCode)) & "'"
'    Else
'        sPropertyAddress = "'" & goUtil.utCleanSQLString(goUtil.utJoinToAddress(sPAStreet, sPACity, sPAState, Format(sPAZIP, "00000") & "-" & Format(sPAZIP4, "0000"))) & "'"
'    End If
'    'Street
'    sPAStreet = "'" & goUtil.utCleanSQLString(UCase(txtPAStreet.Text)) & "'"
'    'City
'    sPACity = "'" & goUtil.utCleanSQLString(UCase(txtPACity.Text)) & "'"
'    'State
'    sPAState = "'" & goUtil.utCleanSQLString(left(UCase(cboPAState.Text), 2)) & "'"
'    'Zip
'    sPAZIP = txtPAZIP.Text
'    'Zip4
'    sPAZIP4 = txtPAZIP4.Text
'    'Other Post Code
'    sPAOtherPostCode = "'" & goUtil.utCleanSQLString(UCase(txtPAOtherPostCode.Text)) & "'"
'    'End Property Address
'
'    'Home Phone
'    sHomePhone = "'" & goUtil.utCleanSQLString(UCase(txtHomePhone.Text)) & "'"
'    'Business Phone
'    sBusinessPhone = "'" & goUtil.utCleanSQLString(UCase(txtBusinessPhone.Text)) & "'"
'    'Mortgage name
'    sMortgageeName = "'" & goUtil.utCleanSQLString(UCase(txtMortgageeName.Text)) & "'"
'    'Agent No
'    sAgentNo = "'" & goUtil.utCleanSQLString(UCase(txtAgentNo.Text)) & "'"
'    'Reported By
'    sReportedBy = "'" & goUtil.utCleanSQLString(UCase(txtReportedBy.Text)) & "'"
'    'Reported by Phone
'    sReportedByPhone = "'" & goUtil.utCleanSQLString(UCase(txtReportedByPhone.Text)) & "'"
'    'Deductible
'    sDeductible = txtDeductible.Text
'
'    sAppDedClassTypeIDOrder = "[AppDedClassTypeIDOrder]"
'
'    'if the Loss report was changed to TEXT then need to update these vars
'    'otherwise they remain the same!
'    '(Attaching a PDF Loss Report already updates Assignments table See --> Private Sub cmdAttachPDFLossReport_Click)
'    If StrComp(cboAssignmentLossReportFormat.Text, "TEXT", vbTextCompare) = 0 Then
'        sLRFormat = "'TEXT'"
'        sLossReport = "'" & goUtil.utCleanSQLString(txtLossReport.Text) & "'"
'
'        sDownLoadLossReport = "[DownLoadLossReport]"
'
'        sUpLoadLossReport = "True"
'    Else
'        sLRFormat = "[LRFormat]"
'
'        sLossReport = "[LossReport]"
'
'        sDownLoadLossReport = "[DownLoadLossReport]"
'
'        sUpLoadLossReport = "[UpLoadLossReport]"
'    End If
'    'Type Of Loss
'    sTypeOfLossID = cboTypeOfLoss.ItemData(cboTypeOfLoss.ListIndex)
'
'    sXactTypeOfLoss = "[XactTypeOfLoss]"
'
'    sSentToXact = "[SentToXact]"
'
'    sReassigned = "[Reassigned]"
'
'    sDateReassigned = "[DateReassigned]"
'
'
'    'STATUS ID !
'    'Check for Closed Date
'    'Change Status ID
'    If IsDate(txtCloseDate.Text) Then
'        sStatusID = CStr(V2ECKeyBoard.AssgnStatus.iAssignmentsStatus_CLOSED)
'    Else
'        'Check to see if the Current Status is Closed if it Is Need to
'        'Change the Status to NEW
'        iCurrentStatus = MyadoRSAssignments.Fields("StatusID").Value
'        Select Case iCurrentStatus
'            Case V2ECKeyBoard.AssgnStatus.iAssignmentsStatus_CLOSED
'                sStatusID = V2ECKeyBoard.AssgnStatus.iAssignmentsStatus_NEW
'            Case Else
'                sStatusID = "[StatusID]"
'        End Select
'
'    End If
'
'    sRAAdjusterSpecID = "[RAAdjusterSpecID]"
'
'    sIsLocked = "[IsLocked]"
'
'    sIsDeleted = "[IsDeleted]"
'
'    sDownLoadMe = "[DownLoadMe]"
'
'    sUpLoadMe = "True"
'
'    sDownLoadAll = "[DownLoadAll]"
'
'    sUpLoadAll = "[UpLoadAll]"
'    'Admin Comments
'    sAdminComments = "'" & goUtil.utCleanSQLString(txtAdminComments.Text) & "'"
'
'    sMiscDelimSettings = "[MiscDelimSettings]"
'
'    sDateLastUpdated = "#" & Format(Now(), "MM/DD/YYYY") & "#"
'
'    sUID = goUtil.gsCurUsersID
'
'
'    sSQL = "Update Assignments Set "
'    sSQL = sSQL & "[AssignmentsID] = " & sAssignmentsID & ", "
'    sSQL = sSQL & "[ID] = " & sID & ", "
'    sSQL = sSQL & "[AssignmentTypeID] = " & sAssignmentTypeID & ", "
'    sSQL = sSQL & "[ClientCompanyCatSpecID] = " & sClientCompanyCatSpecID & ", "
'    sSQL = sSQL & "[AdjusterSpecID] = " & sAdjusterSpecID & ", "
'    sSQL = sSQL & "[AdjusterSpecIDDisplay] =" & sAdjusterSpecIDDisplay & ", "
'    sSQL = sSQL & "[SPVersion] = " & sSPVersion & ", "
'    sSQL = sSQL & "[IBNUM] = " & sIBNUM & ", "
'    sSQL = sSQL & "[CLIENTNUM] = " & sCLIENTNUM & ", "
'    sSQL = sSQL & "[PolicyNo] = " & sPolicyNo & ", "
'    sSQL = sSQL & "[PolicyDescription] = " & sPolicyDescription & ", "
'    sSQL = sSQL & "[Insured] = " & sInsured & ", "
'    sSQL = sSQL & "[MailingAddress] = " & sMailingAddress & ", "
'    sSQL = sSQL & "[MAStreet] = " & sMAStreet & ", "
'    sSQL = sSQL & "[MACity] = " & sMACity & ", "
'    sSQL = sSQL & "[MAState] = " & sMAState & ", "
'    sSQL = sSQL & "[MAZIP] = " & sMAZIP & ", "
'    sSQL = sSQL & "[MAZIP4] = " & sMAZIP4 & ", "
'    sSQL = sSQL & "[MAOtherPostCode] = " & sMAOtherPostCode & ", "
'    sSQL = sSQL & "[HomePhone]  = " & sHomePhone & ", "
'    sSQL = sSQL & "[BusinessPhone] = " & sBusinessPhone & ", "
'    sSQL = sSQL & "[PropertyAddress] = " & sPropertyAddress & ", "
'    sSQL = sSQL & "[PAStreet]  = " & sPAStreet & ", "
'    sSQL = sSQL & "[PACity]  = " & sPACity & ", "
'    sSQL = sSQL & "[PAState] = " & sPAState & ", "
'    sSQL = sSQL & "[PAZIP]  = " & sPAZIP & ", "
'    sSQL = sSQL & "[PAZIP4] = " & sPAZIP4 & ", "
'    sSQL = sSQL & "[PAOtherPostCode]  = " & sPAOtherPostCode & ", "
'    sSQL = sSQL & "[MortgageeName]  = " & sMortgageeName & ", "
'    sSQL = sSQL & "[AgentNo]  = " & sAgentNo & ", "
'    sSQL = sSQL & "[ReportedBy] = " & sReportedBy & ", "
'    sSQL = sSQL & "[ReportedByPhone] = " & sReportedByPhone & ", "
'    sSQL = sSQL & "[Deductible]  = " & sDeductible & ", "
'    sSQL = sSQL & "[AppDedClassTypeIDOrder] = " & sAppDedClassTypeIDOrder & ", "
'    sSQL = sSQL & "[LRFormat]  = " & sLRFormat & ", "
'    sSQL = sSQL & "[LossReport] = " & sLossReport & ", "
'    sSQL = sSQL & "[DownLoadLossReport] = " & sDownLoadLossReport & ", "
'    sSQL = sSQL & "[UpLoadLossReport] = " & sUpLoadLossReport & ", "
'    sSQL = sSQL & "[StatusID]  = " & sStatusID & ", "
'    sSQL = sSQL & "[TypeOfLossID] = " & sTypeOfLossID & ", "
'    sSQL = sSQL & "[XactTypeOfLoss] = " & sXactTypeOfLoss & ", "
'    sSQL = sSQL & "[SentToXact] = " & sSentToXact & ", "
'    sSQL = sSQL & "[LossDate] = " & sLossDate & ", "
'    sSQL = sSQL & "[AssignedDate] = " & sAssignedDate & ", "
'    sSQL = sSQL & "[ReceivedDate] = " & sReceivedDate & ", "
'    sSQL = sSQL & "[ContactDate] = " & sContactDate & ", "
'    sSQL = sSQL & "[InspectedDate] = " & sInspectedDate & ", "
'    sSQL = sSQL & "[CloseDate]  = " & sCloseDate & ", "
'    sSQL = sSQL & "[Reassigned]  = " & sReassigned & ", "
'    sSQL = sSQL & "[DateReassigned] = " & sDateReassigned & ", "
'    sSQL = sSQL & "[RAAdjusterSpecID] = " & sRAAdjusterSpecID & ", "
'    sSQL = sSQL & "[IsLocked] = " & sIsLocked & ", "
'    sSQL = sSQL & "[IsDeleted] = " & sIsDeleted & ", "
'    sSQL = sSQL & "[DownLoadMe] = " & sDownLoadMe & ", "
'    sSQL = sSQL & "[UpLoadMe] = " & sUpLoadMe & ", "
'    sSQL = sSQL & "[DownLoadAll] = " & sDownLoadAll & ", "
'    sSQL = sSQL & "[UpLoadAll] = " & sUpLoadAll & ", "
'    sSQL = sSQL & "[AdminComments] = " & sAdminComments & ", "
'    sSQL = sSQL & "[MiscDelimSettings] = " & sMiscDelimSettings & ", "
'    sSQL = sSQL & "[DateLastUpdated] = " & sDateLastUpdated & ", "
'    sSQL = sSQL & "[UpdateByUserID] = " & sUID & " "
'    sSQL = sSQL & "WHERE AssignmentsID = " & sAssignmentsID & " "
'
'
'    Set oConn = New ADODB.Connection
'    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
'
'    oConn.Execute sSQL
    
    cmdSave.Enabled = False
    SaveMe = True
    
    'cleanup
    Set oConn = Nothing
'    Set MyadoRSAssignments = Nothing
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function SaveMe"
End Function

Public Function CheckStatus() As Boolean
    On Error GoTo EH
'    Dim lTabsPos As Long
'    Dim oFrame As Control
'    Dim MyFrame As Frame
'    Dim oControl As Control
'    Dim MyTextBox As TextBox
'    Dim MycmdButton As CommandButton
'    Dim sFrameName As String
'
'     'If this claim is closed only certain things can be edited
'    If mfrmClaim.MyStatus = iAssignmentsStatus_CLOSED Then
'        For lTabsPos = 1 To TSClaimInfo.Tabs.Count
'            Select Case UCase(TSClaimInfo.Tabs(lTabsPos).Tag)
'                Case UCase(framSpecifics.Name), _
'                        UCase(framInsuredInfo.Name), _
'                        UCase(framPolicyLimits.Name)
'                    sFrameName = TSClaimInfo.Tabs(lTabsPos).Tag
'                    For Each oFrame In Me.Controls
'                        If TypeOf oFrame Is Frame Then
'                            Set MyFrame = oFrame
'                            If StrComp(MyFrame.Name, sFrameName, vbTextCompare) = 0 Then
'                                MyFrame.Enabled = False
'                                Exit For
'                            End If
'                        End If
'                    Next
'                Case UCase(framDates.Name)
'                    'Need to disable all dates except the closedate
'                    For Each oControl In Me.Controls
'                        If TypeOf oControl Is TextBox Then
'                            Set MyTextBox = oControl
'                            If StrComp(MyTextBox.Tag, "Date", vbTextCompare) = 0 Then
'                                If StrComp(MyTextBox.Name, txtCloseDate.Name, vbTextCompare) <> 0 Then
'                                    MyTextBox.Enabled = False
'                                End If
'                            End If
'                        ElseIf TypeOf oControl Is CommandButton Then
'                            Set MycmdButton = oControl
'                            If StrComp(MycmdButton.Tag, "Date", vbTextCompare) = 0 Then
'                                If StrComp(MycmdButton.Name, cmdCloseDate.Name, vbTextCompare) <> 0 Then
'                                    MycmdButton.Enabled = False
'                                End If
'                            End If
'                        End If
'                    Next
'                Case UCase(framLossReport.Name)
'                    'Need to disable all control except the closedate
'                    For Each oControl In Me.Controls
'                        If TypeOf oControl Is CommandButton Then
'                            Set MycmdButton = oControl
'                            If StrComp(MycmdButton.Name, cmdViewPDFLossReport.Name, vbTextCompare) = 0 Then
'                                MycmdButton.Enabled = True
'                            End If
'                        Else
'                            If (Not TypeOf oControl Is ImageList) And (Not TypeOf oControl Is TabStrip) Then
'                                If oControl.Container.Name = framLossReport.Name Then
'                                    oControl.Enabled = False
'                                End If
'                            End If
'                        End If
'                    Next
'            End Select
'        Next
'    End If
    
    CheckStatus = True
    
    'cleanup
'    Set oFrame = Nothing
'    Set MyFrame = Nothing
'    Set oControl = Nothing
'    Set MyTextBox = Nothing
'    Set MycmdButton = Nothing
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function CheckStatus"
End Function


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo EH
    
    Select Case UnloadMode
        Case vbFormControlMenu
            Cancel = True
            mbUnloadMe = True
            Me.Visible = False
            mfrmClaim.Timer_UnloadForm.Enabled = True
        Case Else
            CmdReNumberSort_Click
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
    
    'RePos Controls
    'Width and Lefts
    cboPhotoReports.Width = Me.Width - 9525
    framPhotos.Width = Me.Width - 360
    lstvPhotos.Width = framPhotos.Width - 945
    framPhotoMaint.Width = framPhotos.Width - 225
    chkHideDeleted.left = framPhotoMaint.Width - 2775
    cmdDelPhotos.left = framPhotoMaint.Width - 1575
    'framCommands
    framCommands.left = Me.Width - 4695
    
    'Heights and Tops
    framPhotos.Height = Me.Height - 1815
    lstvPhotos.Height = framPhotos.Height - 1440
    framPhotoMaint.top = framPhotos.Height - 855
    
    'framAssocBillingID
    framAssocBillingID.top = Me.Height - 1710
    
    'framCommands
    framCommands.top = Me.Height - 1710
    
    If Me.WindowState <> vbMinimized Then
        Me.Visible = True
    End If
End Sub

Public Function CLEANUP() As Boolean
    On Error GoTo EH
    If cmdSave.Enabled Then
        SaveMe
        If Not mfrmClaim Is Nothing Then
            mfrmClaim.RefreshMe
        End If
    End If
    Set mfrmClaim = Nothing
    Set MyGUI = Nothing
    Set mitmXSelected = Nothing
    
    If Not moForm Is Nothing Then
        Unload moForm
        Set moForm = Nothing
    End If
    If Not mArv Is Nothing Then
        mArv.CLEANUP
        Set mArv = Nothing
    End If
    
    CLEANUP = True
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function CLEANUP"
End Function

Public Function DeletePhotoItem(psID As String) As Boolean
    On Error GoTo EH
    Dim sSQL As String
    Dim sPath As String
    Dim bUpdateAsDeletedOnly As Boolean
    Dim oConn As ADODB.Connection
    Dim RS As ADODB.Recordset
    
    'Need to remove the actual jpeg files as well because they are
    'not needed anymore. only if this Record has never been uploaded
    
    Set oConn = New ADODB.Connection
    Set RS = New ADODB.Recordset
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name


    sSQL = "SELECT P.[PhotoName] "
    sSQL = sSQL & "FROM RTPhotolog P "
    sSQL = sSQL & "WHERE P.[ID] = " & psID & " "
    'Only allow actual deletion of Photos that have never been uploaded
    'The Main Table Indentity will be negative number if this is true.
    sSQL = sSQL & "AND (P.[RTPhotologID] Is Null Or P.[RTPhotologID] < 0)  "
    
    'Use Disconnected Record Set on asUseClient Cusor ONLY !
    RS.CursorLocation = adUseClient
    RS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set RS.ActiveConnection = Nothing
    
    If RS.RecordCount = 1 Then
        RS.MoveFirst
        If Not IsNull(RS!PhotoName) Then
            'Remove Photo
            sPath = goUtil.PhotoReposPath & RS!PhotoName
            goUtil.utDeleteFile sPath
            'Remove Thumbnail
            sPath = Replace(sPath, "_1.jpg", "_2.jpg")
            goUtil.utDeleteFile sPath
            'Remove HighRes
            sPath = Replace(sPath, "_2.jpg", "_0.jpg")
            goUtil.utDeleteFile sPath
        End If
    Else
        bUpdateAsDeletedOnly = True
    End If

    
    If bUpdateAsDeletedOnly Then
        sSQL = "UPDATE RTPhotolog SET "
        sSQL = sSQL & "[IsDeleted] = IIF([IsDeleted], False, True), "
        sSQL = sSQL & "[UpLoadMe] = True, "
        sSQL = sSQL & "[DateLastUpdated] = #" & Format(Now(), "MM/DD/YYYY HH:MM:SS") & "#, "
        sSQL = sSQL & "[UpdateByUserID] = " & goUtil.gsCurUsersID & " "
        sSQL = sSQL & "WHERE [ID] = " & psID & " "
        sSQL = sSQL & "AND [IDAssignments] = " & msAssignmentsID & " "
    Else
        sSQL = "DELETE * FROM RTPhotolog P "
        sSQL = sSQL & "WHERE P.[ID] = " & psID & " "
        sSQL = sSQL & "AND P.[IDAssignments] = " & msAssignmentsID & " "
    End If

    oConn.Execute sSQL
    
    DeletePhotoItem = True
    'clean up
    
    Set RS = Nothing
    Set oConn = Nothing
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function DeletePhotoItem"
End Function

Public Function AssocPhotoItemToBillingID(psID As String) As Boolean
    On Error GoTo EH
    Dim sSQL As String
    Dim sPath As String
    Dim sIDBillingCount As String
    Dim oConn As ADODB.Connection
    
    
    'Set the IDBillingCOunt to Drop down item data
    
    sIDBillingCount = cboAssocBillingID.ItemData(cboAssocBillingID.ListIndex)
    
    'Check to See if it is Null value
    
    If sIDBillingCount = 0 Then
        sIDBillingCount = "null"
    End If
    
    Set oConn = New ADODB.Connection
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    sSQL = "UPDATE RTPhotolog SET "
    sSQL = sSQL & "[BillingCountID] = " & sIDBillingCount & ", "
    sSQL = sSQL & "[IDBillingCount] = " & sIDBillingCount & ", "
    sSQL = sSQL & "[UpLoadMe] = True, "
    sSQL = sSQL & "[DateLastUpdated] = #" & Format(Now(), "MM/DD/YYYY HH:MM:SS") & "#, "
    sSQL = sSQL & "[UpdateByUserID] = " & goUtil.gsCurUsersID & " "
    sSQL = sSQL & "WHERE [ID] = " & psID & " "
    sSQL = sSQL & "AND IDAssignments = " & msAssignmentsID & " "

    oConn.Execute sSQL
    
    AssocPhotoItemToBillingID = True
    
    'clean up
    Set oConn = Nothing
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function AssocPhotoItemToBillingID"
End Function

Private Sub lstvPhotos_Click()
    On Error GoTo EH
    'Set the selected Photo
    itmXSelected = lstvPhotos.SelectedItem
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lstvPhotos_Click"
End Sub

Private Sub lstvPhotos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    On Error GoTo EH
    Dim sMess As String
    If lstvPhotos.SortOrder = lvwAscending Then
        lstvPhotos.SortOrder = lvwDescending
        sMess = "Descending"
    Else
        lstvPhotos.SortOrder = lvwAscending
        sMess = "Ascending"
    End If
    
    'Set the Tool Tip
    lstvPhotos.ToolTipText = "Sort By " & ColumnHeader.Text & " " & sMess
    
    'Need to see if the Column clicked was a NON text sort
    'Like Date or number.  If a non textual column was clicked
    'need to use the next column as the sort since this next column
    'will be hidden and contain a sort friendly format.
    'Sort Key is Base 0 where CoulmnHeader is not
    Select Case ColumnHeader.Index
        Case GuiPhotoListView.PhotoDate
            lstvPhotos.SortKey = ColumnHeader.Index
        Case GuiPhotoListView.DateLastUpdated
            lstvPhotos.SortKey = ColumnHeader.Index
        Case GuiPhotoListView.SortOrder
            lstvPhotos.SortKey = ColumnHeader.Index
        Case Else
            lstvPhotos.SortKey = ColumnHeader.Index - 1
    End Select
    
    lstvPhotos.Sorted = True
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lstvPhotos_ColumnClick"
End Sub
Private Sub lstvPhotos_DblClick()
    On Error GoTo EH
    'Set the selected claim
    
    itmXSelected = lstvPhotos.SelectedItem
    If Not lstvPhotos.SelectedItem Is Nothing Then
        EditPhoto
    End If
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lstvPhotos_DblClick"
End Sub

Private Sub lstvPhotos_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error GoTo EH
    'Set the selected Photo
    itmXSelected = lstvPhotos.SelectedItem
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lstvPhotos_ItemClick"
End Sub

Private Sub lstvPhotos_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo EH
    Select Case KeyCode
        Case KeyCodeConstants.vbKeyReturn
            EditPhoto
        Case vbKeyDelete
            cmdDelPhotos_Click
    End Select
Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lstvPhotos_KeyDown"
End Sub

Public Sub LoadHeaderlstvPhotos()
    On Error GoTo EH
    Dim bGridOn As Boolean
    Dim bHideDeleted As Boolean
    Dim bHideUploadFlags As Boolean
    
    bHideDeleted = CBool(GetSetting(App.EXEName, "GENERAL", "HIDE_DELETED", True))
    bHideUploadFlags = CBool(GetSetting(App.EXEName, "GENERAL", "HIDE_UPLOAD_FLAGS", True))
    
    'set the columnheaders
    With lstvPhotos
        .ColumnHeaders.Add , "Thumb", "Thumb"
        .ColumnHeaders.Add , "SortOrder", "Sort Order"
        .ColumnHeaders.Add , "SortOrderSort", "Sort Sort Order" 'Hidden
        .ColumnHeaders.Add , "Status", "Status"
        .ColumnHeaders.Add , "IB", "IB" ' Shows the IB this Photo Items is Asscoiated with
        .ColumnHeaders.Add , "PhotoDate", "Date"
        .ColumnHeaders.Add , "PhotoDateSort", "Sort Date"
        .ColumnHeaders.Add , "Description", "Description"
        .ColumnHeaders.Add , "PhotoName", "Photo Name" 'hidden
        .ColumnHeaders.Add , "UpLoadPhoto", "UpLoad Photo" 'hidden
        .ColumnHeaders.Add , "UpLoadPhotoThumb", "UpLoad Thumb" 'hidden
        .ColumnHeaders.Add , "UploadPhotoHighRes", "Upload HighRes" ' hidden  this option is not yet supported 7.22.2004
        .ColumnHeaders.Add , "IsDeleted", "Is Deleted"
        .ColumnHeaders.Add , "UpLoadMe", "UpLoad Me"
        .ColumnHeaders.Add , "DateLastUpdated", "Date Last Updated"
        .ColumnHeaders.Add , "DateLastUpdatedSort", "Sort Date Last Updated" ' hidden
        .ColumnHeaders.Add , "AdminComments", "Admin Comments" 'hidden
        .ColumnHeaders.Add , "ID", "ID" 'Hidden
        .ColumnHeaders.Add , "IDRTPhotoReport", "IDRTPhotoReport" 'Hidden
        .ColumnHeaders.Add , "IDAssignments", "IDAssignments" 'Hidden
        .ColumnHeaders.Add , "RTPhotoLogID", "RTPhotoLogID" ' Hidden
        .ColumnHeaders.Add , "RTPhotoReportID", "RTPhotoReportID" 'Hidden
        .ColumnHeaders.Add , "AssignmentsID", "AssignmentsID"  ' Hidden
        .ColumnHeaders.Add , "BillingCountID", "BillingCountID"  ' Hidden
        .ColumnHeaders.Add , "IDBillingCount", "IDBillingCount" ' Hidden
        .ColumnHeaders.Add , "Photo", "Photo"  'Hidden
        .ColumnHeaders.Add , "DownloadPhoto", "DownloadPhoto"  'hidden
        .ColumnHeaders.Add , "PhotoThumb", "PhotoThumb"  'hidden
        .ColumnHeaders.Add , "DownloadPhotoThumb", "DownloadPhotoThumb"  'hidden
        .ColumnHeaders.Add , "PhotoHighRes", "PhotoHighRes"  'hidden
        .ColumnHeaders.Add , "DownloadPhotoHighRes", "DownloadPhotoHighRes"  'hidden
        .ColumnHeaders.Add , "DownLoadMe", "DownLoadMe"  ' hidden
        .ColumnHeaders.Add , "UpdateByUserID", "UpdateByUserID"  ' Hidden
        
        
        .Sorted = False
        .SortOrder = lvwAscending
        'Thumb
        .ColumnHeaders.Item(GuiPhotoListView.Thumb).Width = 1470
        .ColumnHeaders.Item(GuiPhotoListView.Thumb).Alignment = lvwColumnLeft
        'SortOrder
        .ColumnHeaders.Item(GuiPhotoListView.SortOrder).Width = 1230
        .ColumnHeaders.Item(GuiPhotoListView.SortOrder).Alignment = lvwColumnLeft
        'SortOrderSort
        .ColumnHeaders.Item(GuiPhotoListView.SortOrderSort).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiPhotoListView.SortOrderSort).Alignment = lvwColumnLeft
        'Status
        .ColumnHeaders.Item(GuiPhotoListView.Status).Width = 750
        .ColumnHeaders.Item(GuiPhotoListView.Status).Alignment = lvwColumnLeft
        'IB
        .ColumnHeaders.Item(GuiPhotoListView.IB).Width = 750
        .ColumnHeaders.Item(GuiPhotoListView.IB).Alignment = lvwColumnLeft
        'PhotoDate
        .ColumnHeaders.Item(GuiPhotoListView.PhotoDate).Width = 1335
        .ColumnHeaders.Item(GuiPhotoListView.PhotoDate).Alignment = lvwColumnLeft
        'PhotoDateSort
        .ColumnHeaders.Item(GuiPhotoListView.PhotoDateSort).Width = 0
        .ColumnHeaders.Item(GuiPhotoListView.PhotoDateSort).Alignment = lvwColumnLeft
        'Description
        .ColumnHeaders.Item(GuiPhotoListView.Description).Width = 9000
        .ColumnHeaders.Item(GuiPhotoListView.Description).Alignment = lvwColumnLeft
        'PhotoName
        .ColumnHeaders.Item(GuiPhotoListView.PhotoName).Width = 0 ' hidden 5000
        .ColumnHeaders.Item(GuiPhotoListView.PhotoName).Alignment = lvwColumnLeft
        'UpLoadPhoto
        .ColumnHeaders.Item(GuiPhotoListView.UpLoadPhoto).Width = 0 ' hidden 400
        .ColumnHeaders.Item(GuiPhotoListView.UpLoadPhoto).Alignment = lvwColumnCenter
        .ColumnHeaders.Item(GuiPhotoListView.UpLoadPhoto).Icon = GuiPhotoStatusList.UpLoadMeColHeader
        'UpLoadPhotoThumb
        .ColumnHeaders.Item(GuiPhotoListView.UpLoadPhotoThumb).Width = 0 ' hidden 400
        .ColumnHeaders.Item(GuiPhotoListView.UpLoadPhotoThumb).Alignment = lvwColumnCenter
        .ColumnHeaders.Item(GuiPhotoListView.UpLoadPhotoThumb).Icon = GuiPhotoStatusList.UpLoadMeColHeader
        'UploadPhotoHighRes
        .ColumnHeaders.Item(GuiPhotoListView.UploadPhotoHighRes).Width = 0 ' hidden  this option is not yet supported 7.22.2004
        .ColumnHeaders.Item(GuiPhotoListView.UploadPhotoHighRes).Alignment = lvwColumnCenter
        .ColumnHeaders.Item(GuiPhotoListView.UploadPhotoHighRes).Icon = GuiPhotoStatusList.UpLoadMeColHeader
        'Is Deleted
        If bHideDeleted Then
            .ColumnHeaders.Item(GuiPhotoListView.IsDeleted).Width = 0 ' hidden 400
        Else
            .ColumnHeaders.Item(GuiPhotoListView.IsDeleted).Width = 400
        End If
        .ColumnHeaders.Item(GuiPhotoListView.IsDeleted).Alignment = lvwColumnCenter
        .ColumnHeaders.Item(GuiPhotoListView.IsDeleted).Icon = GuiPhotoStatusList.IsDeletedColHeader
        'UpLoad Me
        If bHideUploadFlags Then
            .ColumnHeaders.Item(GuiPhotoListView.UpLoadMe).Width = 0 ' hidden 400
        Else
            .ColumnHeaders.Item(GuiPhotoListView.UpLoadMe).Width = 400
        End If
        .ColumnHeaders.Item(GuiPhotoListView.UpLoadMe).Alignment = lvwColumnCenter
        .ColumnHeaders.Item(GuiPhotoListView.UpLoadMe).Icon = GuiPhotoStatusList.UpLoadMeColHeader
        'DateLastUpdated
        .ColumnHeaders.Item(GuiPhotoListView.DateLastUpdated).Width = 2200
        .ColumnHeaders.Item(GuiPhotoListView.DateLastUpdated).Alignment = lvwColumnLeft
        'DateLastUpdatedSort
        .ColumnHeaders.Item(GuiPhotoListView.DateLastUpdatedSort).Width = 0  'Hidden Sort Column
        .ColumnHeaders.Item(GuiPhotoListView.DateLastUpdatedSort).Alignment = lvwColumnLeft
        'AdminComments
        .ColumnHeaders.Item(GuiPhotoListView.AdminComments).Width = 0  'Hidden 10000
        .ColumnHeaders.Item(GuiPhotoListView.AdminComments).Alignment = lvwColumnLeft
        'ID
        .ColumnHeaders.Item(GuiPhotoListView.ID).Width = 0  'hidden
        .ColumnHeaders.Item(GuiPhotoListView.ID).Alignment = lvwColumnLeft
        'ID
        .ColumnHeaders.Item(GuiPhotoListView.IDRTPhotoReport).Width = 0   'hidden
        .ColumnHeaders.Item(GuiPhotoListView.IDRTPhotoReport).Alignment = lvwColumnLeft
        'IDAssignments
        .ColumnHeaders.Item(GuiPhotoListView.IDAssignments).Width = 0  'Hidden
        .ColumnHeaders.Item(GuiPhotoListView.IDAssignments).Alignment = lvwColumnLeft
        'RTPhotoLogID
        .ColumnHeaders.Item(GuiPhotoListView.RTPhotoLogID).Width = 0  'Hidden
        .ColumnHeaders.Item(GuiPhotoListView.RTPhotoLogID).Alignment = lvwColumnLeft
        'RTPhotoLogID
        .ColumnHeaders.Item(GuiPhotoListView.RTPhotoReportID).Width = 0  'Hidden
        .ColumnHeaders.Item(GuiPhotoListView.RTPhotoReportID).Alignment = lvwColumnLeft
        'AssignmentsID
        .ColumnHeaders.Item(GuiPhotoListView.AssignmentsID).Width = 0  'Hidden
        .ColumnHeaders.Item(GuiPhotoListView.AssignmentsID).Alignment = lvwColumnLeft
        'BillingCountID
        .ColumnHeaders.Item(GuiPhotoListView.BillingCountID).Width = 0  'Hidden
        .ColumnHeaders.Item(GuiPhotoListView.BillingCountID).Alignment = lvwColumnLeft
        'IDBillingCount
        .ColumnHeaders.Item(GuiPhotoListView.IDBillingCount).Width = 0  'Hidden
        .ColumnHeaders.Item(GuiPhotoListView.IDBillingCount).Alignment = lvwColumnLeft
        'Photo
        .ColumnHeaders.Item(GuiPhotoListView.Photo).Width = 0  'Hidden
        .ColumnHeaders.Item(GuiPhotoListView.Photo).Alignment = lvwColumnLeft
        'DownloadPhoto
        .ColumnHeaders.Item(GuiPhotoListView.DownloadPhoto).Width = 0  'Hidden
        .ColumnHeaders.Item(GuiPhotoListView.DownloadPhoto).Alignment = lvwColumnLeft
        'PhotoThumb
        .ColumnHeaders.Item(GuiPhotoListView.PhotoThumb).Width = 0  'Hidden
        .ColumnHeaders.Item(GuiPhotoListView.PhotoThumb).Alignment = lvwColumnLeft
        'DownloadPhotoThumb
        .ColumnHeaders.Item(GuiPhotoListView.DownloadPhotoThumb).Width = 0  'Hidden
        .ColumnHeaders.Item(GuiPhotoListView.DownloadPhotoThumb).Alignment = lvwColumnLeft
        'PhotoHighRes
        .ColumnHeaders.Item(GuiPhotoListView.PhotoHighRes).Width = 0  'Hidden
        .ColumnHeaders.Item(GuiPhotoListView.PhotoHighRes).Alignment = lvwColumnLeft
        'DownloadPhotoHighRes
        .ColumnHeaders.Item(GuiPhotoListView.DownloadPhotoHighRes).Width = 0  'Hidden
        .ColumnHeaders.Item(GuiPhotoListView.DownloadPhotoHighRes).Alignment = lvwColumnLeft
        'DownLoadMe
        .ColumnHeaders.Item(GuiPhotoListView.DownLoadMe).Width = 0  'Hidden
        .ColumnHeaders.Item(GuiPhotoListView.DownLoadMe).Alignment = lvwColumnLeft
        'UpdateByUserID
        .ColumnHeaders.Item(GuiPhotoListView.UpdateByUserID).Width = 0  'Hidden
        .ColumnHeaders.Item(GuiPhotoListView.UpdateByUserID).Alignment = lvwColumnLeft
    End With
    
    bGridOn = CBool(GetSetting(App.EXEName, "GENERAL", "GRID_ON", False))
    lstvPhotos.GridLines = bGridOn
    
    If bHideDeleted Then
        chkHideDeleted.Value = vbChecked
    Else
        chkHideDeleted.Value = vbUnchecked
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub LoadHeaderlstvPhotos"
End Sub

Private Sub PopulatelstvPhotos()
    On Error GoTo EH
    Dim iMyIcon As Long
    Dim sFlagText As String
    Dim sTemp As String
    Dim sPhotoName As String
    Dim sUploadPhotoPath As String
    Dim bMissingPhoto As Boolean
    Dim bUploadedAllPhotos As Boolean
    Dim sMessMissingPhoto As String
    Dim lPhotoCount As Long
    Dim lPhotoIndex As Long
    Dim lPhotoStatusIndex As Long
    Dim itmX As ListItem
    Dim oListView As MSComctlLib.ListView
    Dim oImageList As MSComctlLib.ImageList
    Dim oPhotoStatusList As MSComctlLib.ImageList
    Dim MyadoRSRTPhotoLog As ADODB.Recordset
    
    'Clear the List view
    Set oListView = lstvPhotos
    
    oListView.Visible = False
    oListView.ListItems.Clear
    
    'Reset the Bound Icons
    Set oListView.Icons = Nothing
    Set oListView.SmallIcons = Nothing
    
    'Set the Image lists
    Set oImageList = imgListPhotos
    Set oPhotoStatusList = imgPhotoStatus
    
    'Clear the thumbs
    oImageList.ListImages.Clear
    
    Set MyadoRSRTPhotoLog = mfrmClaim.adoRSRTPhotoLog
    
    If Not MyadoRSRTPhotoLog.EOF Then
        MyadoRSRTPhotoLog.MoveFirst
        'First populate the image list  with thumbnail images
        'If the thumnail is missing for some reason then add Deleted Image in its place
        Do Until MyadoRSRTPhotoLog.EOF
            sPhotoName = goUtil.IsNullIsVbNullString(MyadoRSRTPhotoLog.Fields("PhotoName"))
            sPhotoName = Replace(sPhotoName, "_1.jpg", "_2.jpg", , 1, vbTextCompare)
            If goUtil.utFileExists(goUtil.PhotoReposPath & sPhotoName) Then
                'If there is an error loading the picture then use error picture
                On Error Resume Next
                oImageList.ListImages.Add , , LoadPicture(goUtil.PhotoReposPath & sPhotoName)
                If Err.Number > 0 Then
                    Err.Clear
                    oImageList.ListImages.Add , , oPhotoStatusList.ListImages(GuiPhotoStatusList.IsDeleted).Picture
                End If
                On Error GoTo EH
            Else
                oImageList.ListImages.Add , , oPhotoStatusList.ListImages(GuiPhotoStatusList.IsDeleted).Picture
            End If
            MyadoRSRTPhotoLog.MoveNext
        Loop
        
        lPhotoCount = oImageList.ListImages.Count
        
        'Now se the thumnail images to the list view so it
        'can point to them
        If MyadoRSRTPhotoLog.RecordCount > 0 Then
            Set oListView.Icons = oImageList
            Set oListView.SmallIcons = oImageList
            MyadoRSRTPhotoLog.MoveFirst
        End If
        
        lPhotoIndex = 0
        'Start the status index at the end of the thumbnail count
        'The same image list for thumbnails will also hold Status Images
        lPhotoStatusIndex = lPhotoCount
        Do Until MyadoRSRTPhotoLog.EOF
            lPhotoIndex = lPhotoIndex + 1
            bMissingPhoto = False
            'Start this off as true and set to false if
            'any of the three photos has not yet been uploaded...
            '(Highres=(_0.jpg) ,Main Photo=(_1.jpg), Thumb Nail=(_2.jpg)
            bUploadedAllPhotos = True
            sMessMissingPhoto = vbNullString
            
            '1. ThumbNail
            Set itmX = oListView.ListItems.Add(, """" & goUtil.IsNullIsVbNullString(MyadoRSRTPhotoLog.Fields("ID")) & """", goUtil.IsNullIsVbNullString(MyadoRSRTPhotoLog.Fields("SortOrder")), lPhotoIndex, lPhotoIndex)
            
            '2. Sort Order
            itmX.SubItems(GuiPhotoListView.SortOrder - 1) = goUtil.IsNullIsVbNullString(MyadoRSRTPhotoLog.Fields("SortOrder"))
            'Sort Order Sort
            itmX.SubItems(GuiPhotoListView.SortOrderSort - 1) = goUtil.utNumInTextSortFormat(goUtil.IsNullIsVbNullString(MyadoRSRTPhotoLog.Fields("SortOrder")))
            
            '3. Status
            'If the Upoad photo is missing then Add the Report Icon to Show it was uploaded.
            'If the Source Photo is Missing then display error icon
            'Check for the Highres, MainPhoto, and Thumbnail
            'Highres
            'Check to see if it is flagged for upload if it is need to reset bUploadedAllPhotos
            If CBool(MyadoRSRTPhotoLog.Fields("UploadPhotoHighRes")) Then
                bUploadedAllPhotos = False
            End If
            sPhotoName = goUtil.IsNullIsVbNullString(MyadoRSRTPhotoLog.Fields("PhotoName"))
            sUploadPhotoPath = goUtil.PhotoReposPath & Replace(sPhotoName, "_1.jpg", "_0.jpg", , 1, vbTextCompare)
            If Not goUtil.utFileExists(sUploadPhotoPath) Then
                bMissingPhoto = True
                sMessMissingPhoto = sMessMissingPhoto & "Missing Highres Photo. "
            End If
            'Main Photo
            'Check to see if it is flagged for upload if it is need to reset bUploadedAllPhotos
            If CBool(MyadoRSRTPhotoLog.Fields("UpLoadPhoto")) Then
                bUploadedAllPhotos = False
            End If
            sUploadPhotoPath = goUtil.PhotoReposPath & sPhotoName
            If Not goUtil.utFileExists(sUploadPhotoPath) Then
                bMissingPhoto = True
                sMessMissingPhoto = sMessMissingPhoto & "Missing Main Photo. "
            End If
            'Thumb Nail
            'Check to see if it is flagged for upload if it is need to reset bUploadedAllPhotos
            If CBool(MyadoRSRTPhotoLog.Fields("UpLoadPhotoThumb")) Then
                bUploadedAllPhotos = False
            End If
            sUploadPhotoPath = goUtil.PhotoReposPath & Replace(sPhotoName, "_1.jpg", "_2.jpg", , 1, vbTextCompare)
            If Not goUtil.utFileExists(sUploadPhotoPath) Then
                bMissingPhoto = True
                sMessMissingPhoto = sMessMissingPhoto & "Missing Thumbnail Photo. "
            End If
            
            'Set the Status Column  (This is actually not in the DB as a Field)
            If Not bMissingPhoto And bUploadedAllPhotos Then
                oImageList.ListImages.Add , , oPhotoStatusList.ListImages(GuiPhotoStatusList.PhotoUploaded).Picture
                lPhotoStatusIndex = lPhotoStatusIndex + 1
                itmX.SubItems(GuiPhotoListView.Status - 1) = "SENT"
                itmX.ListSubItems(GuiPhotoListView.Status - 1).ReportIcon = lPhotoStatusIndex
            ElseIf bMissingPhoto Then
                oImageList.ListImages.Add , , oPhotoStatusList.ListImages(GuiPhotoStatusList.IsDeleted).Picture
                lPhotoStatusIndex = lPhotoStatusIndex + 1
                itmX.SubItems(GuiPhotoListView.Status - 1) = sMessMissingPhoto
                itmX.ListSubItems(GuiPhotoListView.Status - 1).ReportIcon = lPhotoStatusIndex
            Else
                itmX.SubItems(GuiPhotoListView.Status - 1) = sMessMissingPhoto
            End If
            
            '4. IB
            itmX.SubItems(GuiPhotoListView.IB - 1) = goUtil.IsNullIsVbNullString(MyadoRSRTPhotoLog.Fields("IB"))
            
            '5. Photo Date
            If Not IsNull(MyadoRSRTPhotoLog.Fields("PhotoDate").Value) Then
                If IsDate(MyadoRSRTPhotoLog.Fields("PhotoDate").Value) Then
                    itmX.SubItems(GuiPhotoListView.PhotoDate - 1) = Format(MyadoRSRTPhotoLog.Fields("PhotoDate").Value, "MM/DD/YYYY")
                    itmX.SubItems(GuiPhotoListView.PhotoDateSort - 1) = Format(MyadoRSRTPhotoLog.Fields("PhotoDate").Value, "YYYY/MM/DD")
                Else
                    itmX.SubItems(GuiPhotoListView.PhotoDate - 1) = vbNullString
                    itmX.SubItems(GuiPhotoListView.PhotoDateSort - 1) = vbNullString
                End If
            Else
                itmX.SubItems(GuiPhotoListView.PhotoDate - 1) = vbNullString
                itmX.SubItems(GuiPhotoListView.PhotoDateSort - 1) = vbNullString
            End If
            
            '6.Description
            itmX.SubItems(GuiPhotoListView.Description - 1) = goUtil.IsNullIsVbNullString(MyadoRSRTPhotoLog.Fields("Description"))
            '7.PhotoName
            itmX.SubItems(GuiPhotoListView.PhotoName - 1) = goUtil.IsNullIsVbNullString(MyadoRSRTPhotoLog.Fields("PhotoName"))
            '8. UpLoadPhoto
            sFlagText = goUtil.GetFlagText(MyadoRSRTPhotoLog.Fields("UpLoadPhoto").Value)
            itmX.SubItems(GuiPhotoListView.UpLoadPhoto - 1) = sFlagText
            If CBool(MyadoRSRTPhotoLog.Fields("UpLoadPhoto").Value) Then
                oImageList.ListImages.Add , , oPhotoStatusList.ListImages(GuiPhotoStatusList.UpLoadMe).Picture
                lPhotoStatusIndex = lPhotoStatusIndex + 1
                itmX.ListSubItems(GuiPhotoListView.UpLoadPhoto - 1).ReportIcon = lPhotoStatusIndex
            Else
                itmX.ListSubItems(GuiPhotoListView.UpLoadPhoto - 1).ReportIcon = Empty
            End If
            
            '9.UpLoadPhotoThumb
            sFlagText = goUtil.GetFlagText(MyadoRSRTPhotoLog.Fields("UpLoadPhotoThumb").Value)
            itmX.SubItems(GuiPhotoListView.UpLoadPhotoThumb - 1) = sFlagText
            If CBool(MyadoRSRTPhotoLog.Fields("UpLoadPhotoThumb").Value) Then
                oImageList.ListImages.Add , , oPhotoStatusList.ListImages(GuiPhotoStatusList.UpLoadMe).Picture
                lPhotoStatusIndex = lPhotoStatusIndex + 1
                itmX.ListSubItems(GuiPhotoListView.UpLoadPhotoThumb - 1).ReportIcon = lPhotoStatusIndex
            Else
                itmX.ListSubItems(GuiPhotoListView.UpLoadPhotoThumb - 1).ReportIcon = Empty
            End If
            
            '10.UploadPhotoHighRes
            sFlagText = goUtil.GetFlagText(MyadoRSRTPhotoLog.Fields("UploadPhotoHighRes").Value)
            itmX.SubItems(GuiPhotoListView.UploadPhotoHighRes - 1) = sFlagText
            If CBool(MyadoRSRTPhotoLog.Fields("UploadPhotoHighRes").Value) Then
                oImageList.ListImages.Add , , oPhotoStatusList.ListImages(GuiPhotoStatusList.UpLoadMe).Picture
                lPhotoStatusIndex = lPhotoStatusIndex + 1
                itmX.ListSubItems(GuiPhotoListView.UploadPhotoHighRes - 1).ReportIcon = lPhotoStatusIndex
            Else
                itmX.ListSubItems(GuiPhotoListView.UploadPhotoHighRes - 1).ReportIcon = Empty
            End If
            
            '11.IsDeleted
            sFlagText = goUtil.GetFlagText(MyadoRSRTPhotoLog.Fields("IsDeleted").Value)
            itmX.SubItems(GuiPhotoListView.IsDeleted - 1) = sFlagText
            If CBool(MyadoRSRTPhotoLog.Fields("IsDeleted").Value) Then
                oImageList.ListImages.Add , , oPhotoStatusList.ListImages(GuiPhotoStatusList.IsDeleted).Picture
                lPhotoStatusIndex = lPhotoStatusIndex + 1
                itmX.ListSubItems(GuiPhotoListView.IsDeleted - 1).ReportIcon = lPhotoStatusIndex
            Else
                itmX.ListSubItems(GuiPhotoListView.IsDeleted - 1).ReportIcon = Empty
            End If
            
            '12.UpLoadMe
            sFlagText = goUtil.GetFlagText(MyadoRSRTPhotoLog.Fields("UpLoadMe").Value)
            itmX.SubItems(GuiPhotoListView.UpLoadMe - 1) = sFlagText
            If CBool(MyadoRSRTPhotoLog.Fields("UpLoadMe").Value) Then
                oImageList.ListImages.Add , , oPhotoStatusList.ListImages(GuiPhotoStatusList.UpLoadMe).Picture
                lPhotoStatusIndex = lPhotoStatusIndex + 1
                itmX.ListSubItems(GuiPhotoListView.UpLoadMe - 1).ReportIcon = lPhotoStatusIndex
            Else
                itmX.ListSubItems(GuiPhotoListView.UpLoadMe - 1).ReportIcon = Empty
            End If
            
            '13. Date Last Updated
            If Not IsNull(MyadoRSRTPhotoLog.Fields("DateLastUpdated").Value) Then
                If IsDate(MyadoRSRTPhotoLog.Fields("DateLastUpdated").Value) Then
                    itmX.SubItems(GuiPhotoListView.DateLastUpdated - 1) = Format(MyadoRSRTPhotoLog.Fields("DateLastUpdated").Value, "MM/DD/YYYY HH:MM:SS")
                    itmX.SubItems(GuiPhotoListView.DateLastUpdatedSort - 1) = Format(MyadoRSRTPhotoLog.Fields("DateLastUpdated").Value, "YYYY/MM/DD HH:MM:SS")
                Else
                    itmX.SubItems(GuiPhotoListView.DateLastUpdated - 1) = vbNullString
                    itmX.SubItems(GuiPhotoListView.DateLastUpdatedSort - 1) = vbNullString
                End If
            Else
                itmX.SubItems(GuiPhotoListView.DateLastUpdated - 1) = vbNullString
                itmX.SubItems(GuiPhotoListView.DateLastUpdatedSort - 1) = vbNullString
            End If
            
            '14. AdminComments
            itmX.SubItems(GuiPhotoListView.AdminComments - 1) = goUtil.IsNullIsVbNullString(MyadoRSRTPhotoLog.Fields("AdminComments"))
            
            '15. ID
            itmX.SubItems(GuiPhotoListView.ID - 1) = goUtil.IsNullIsVbNullString(MyadoRSRTPhotoLog.Fields("ID")) 'Hidden
            
            '15a. IDRTPhotoReport
            itmX.SubItems(GuiPhotoListView.IDRTPhotoReport - 1) = goUtil.IsNullIsVbNullString(MyadoRSRTPhotoLog.Fields("IDRTPhotoReport"))  'Hidden
            
            '16. IDAssignments
            itmX.SubItems(GuiPhotoListView.IDAssignments - 1) = goUtil.IsNullIsVbNullString(MyadoRSRTPhotoLog.Fields("IDAssignments")) 'Hidden
            
            '17. RTPhotoLogID
            itmX.SubItems(GuiPhotoListView.RTPhotoLogID - 1) = goUtil.IsNullIsVbNullString(MyadoRSRTPhotoLog.Fields("RTPhotoLogID")) 'Hidden
            
            '17a. RTPhotoReportID
            itmX.SubItems(GuiPhotoListView.RTPhotoReportID - 1) = goUtil.IsNullIsVbNullString(MyadoRSRTPhotoLog.Fields("RTPhotoReportID"))  'Hidden
            
            '18. AssignmentsID
            itmX.SubItems(GuiPhotoListView.AssignmentsID - 1) = goUtil.IsNullIsVbNullString(MyadoRSRTPhotoLog.Fields("AssignmentsID")) 'Hidden
            
            '19. BillingCountID
            itmX.SubItems(GuiPhotoListView.BillingCountID - 1) = goUtil.IsNullIsVbNullString(MyadoRSRTPhotoLog.Fields("BillingCountID")) 'Hidden
            
            '20. IDBillingCount
            itmX.SubItems(GuiPhotoListView.IDBillingCount - 1) = goUtil.IsNullIsVbNullString(MyadoRSRTPhotoLog.Fields("IDBillingCount")) 'Hidden
            
            '21. Photo
            itmX.SubItems(GuiPhotoListView.Photo - 1) = goUtil.IsNullIsVbNullString(MyadoRSRTPhotoLog.Fields("Photo")) 'Hidden
            
            '22. DownloadPhoto
            sFlagText = goUtil.GetFlagText(MyadoRSRTPhotoLog.Fields("DownloadPhoto").Value)
            itmX.SubItems(GuiPhotoListView.DownloadPhoto - 1) = sFlagText 'Hidden
            If CBool(MyadoRSRTPhotoLog.Fields("DownloadPhoto").Value) Then
                oImageList.ListImages.Add , , oPhotoStatusList.ListImages(GuiPhotoStatusList.UpLoadMe).Picture
                lPhotoStatusIndex = lPhotoStatusIndex + 1
                itmX.ListSubItems(GuiPhotoListView.DownloadPhoto - 1).ReportIcon = lPhotoStatusIndex
            Else
                itmX.ListSubItems(GuiPhotoListView.DownloadPhoto - 1).ReportIcon = Empty
            End If
            
            '23. PhotoThumb
            itmX.SubItems(GuiPhotoListView.PhotoThumb - 1) = goUtil.IsNullIsVbNullString(MyadoRSRTPhotoLog.Fields("PhotoThumb")) 'Hidden
            
            '24. DownloadPhotoThumb
            sFlagText = goUtil.GetFlagText(MyadoRSRTPhotoLog.Fields("DownloadPhotoThumb").Value)
            itmX.SubItems(GuiPhotoListView.DownloadPhotoThumb - 1) = sFlagText 'Hidden
            If CBool(MyadoRSRTPhotoLog.Fields("DownloadPhotoThumb").Value) Then
                oImageList.ListImages.Add , , oPhotoStatusList.ListImages(GuiPhotoStatusList.UpLoadMe).Picture
                lPhotoStatusIndex = lPhotoStatusIndex + 1
                itmX.ListSubItems(GuiPhotoListView.DownloadPhotoThumb - 1).ReportIcon = lPhotoStatusIndex
            Else
                itmX.ListSubItems(GuiPhotoListView.DownloadPhotoThumb - 1).ReportIcon = Empty
            End If
            
            '25. PhotoHighRes
            itmX.SubItems(GuiPhotoListView.PhotoHighRes - 1) = goUtil.IsNullIsVbNullString(MyadoRSRTPhotoLog.Fields("PhotoHighRes")) 'Hidden
            
            '26. DownloadPhotoHighRes
            sFlagText = goUtil.GetFlagText(MyadoRSRTPhotoLog.Fields("DownloadPhotoHighRes").Value)
            itmX.SubItems(GuiPhotoListView.DownloadPhotoHighRes - 1) = sFlagText 'Hidden
            If CBool(MyadoRSRTPhotoLog.Fields("DownloadPhotoHighRes").Value) Then
                oImageList.ListImages.Add , , oPhotoStatusList.ListImages(GuiPhotoStatusList.UpLoadMe).Picture
                lPhotoStatusIndex = lPhotoStatusIndex + 1
                itmX.ListSubItems(GuiPhotoListView.DownloadPhotoHighRes - 1).ReportIcon = lPhotoStatusIndex
            Else
                itmX.ListSubItems(GuiPhotoListView.DownloadPhotoHighRes - 1).ReportIcon = Empty
            End If
            
            '27. DownLoadMe
            sFlagText = goUtil.GetFlagText(MyadoRSRTPhotoLog.Fields("DownLoadMe").Value)
            itmX.SubItems(GuiPhotoListView.DownLoadMe - 1) = sFlagText 'Hidden
            If CBool(MyadoRSRTPhotoLog.Fields("DownLoadMe").Value) Then
                oImageList.ListImages.Add , , oPhotoStatusList.ListImages(GuiPhotoStatusList.UpLoadMe).Picture
                lPhotoStatusIndex = lPhotoStatusIndex + 1
                itmX.ListSubItems(GuiPhotoListView.DownLoadMe - 1).ReportIcon = lPhotoStatusIndex
            Else
                itmX.ListSubItems(GuiPhotoListView.DownLoadMe - 1).ReportIcon = Empty
            End If
            
            '28. UpdateByUserID
            itmX.SubItems(GuiPhotoListView.UpdateByUserID - 1) = goUtil.IsNullIsVbNullString(MyadoRSRTPhotoLog.Fields("UpdateByUserID")) 'Hidden
            
            itmX.Selected = False
            
            MyadoRSRTPhotoLog.MoveNext
        Loop
    End If
    oListView.Visible = True
    'Cleanup
    Set itmX = Nothing
    Set oListView = Nothing
    Set oImageList = Nothing
    Set oPhotoStatusList = Nothing
    Set MyadoRSRTPhotoLog = Nothing
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub PopulatelstvPhotos"
    oListView.Visible = True
End Sub


Private Sub lstvPhotos_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo EH
    If Button = vbRightButton Then
        PopupMenu PopupMnuPhoto
    End If
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lstvPhotos_MouseUp"
End Sub


Public Sub mnuDeletePhoto_Click()
    On Error GoTo EH
    Dim itmX As MSComctlLib.ListItem
    Dim sPhotoID As String
    Dim sSortOrder As String
    
    Set itmX = lstvPhotos.SelectedItem
    
    If Not itmX Is Nothing Then
        sPhotoID = itmX.SubItems(GuiPhotoListView.ID - 1)
        sSortOrder = itmX.SubItems(GuiPhotoListView.SortOrder - 1)
        If MsgBox("Are you sure you want to delete photo " & sSortOrder & " ?", vbYesNo, "DELETE SELECTED PHOTO") = vbYes Then
            If DeletePhotoItem(sPhotoID) Then
                lstvPhotos.ListItems.Remove ("""" & sPhotoID & """")
            End If
        End If
        lstvPhotos.SetFocus
    End If
    
    Set itmX = Nothing
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub mnuDeletePhoto_Click"
End Sub

Public Sub mnuEditPhoto_Click()
    On Error GoTo EH
    
    EditPhoto
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub mnuEditPhoto_Click"
End Sub


Private Sub mnuSelectAllPhoto_Click()
    On Error GoTo EH
    Dim itmX As ListItem
    
    For Each itmX In lstvPhotos.ListItems
        itmX.Selected = True
    Next
    
    Set itmX = Nothing
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub mnuSelectAllPhoto_Click"
End Sub

Public Function EditPhoto() As Boolean
    On Error GoTo EH
    Dim oListView As MSComctlLib.ListView
    Dim itmX As MSComctlLib.ListItem
    Dim sDeresPhoto As String

    If lstvPhotos.ListItems.Count = 0 Then
        Exit Function
    Else
        Set oListView = lstvPhotos
    End If

    Set itmX = oListView.SelectedItem
    
    With AddPhoto
        .MyPhotos = Me
        .MyfrmClaim = Me.MyfrmClaim
         Load AddPhoto
        .Adding = False
        sDeresPhoto = itmX.SubItems(GuiPhotoListView.PhotoName - 1)
        sDeresPhoto = Replace(sDeresPhoto, "_1.jpg", "_0.jpg", , , vbTextCompare)
        sDeresPhoto = goUtil.PhotoReposPath & sDeresPhoto
        .EditDeResPath = sDeresPhoto
        .Caption = "Photo Edit"
        .AssignmentsID = itmX.SubItems(GuiPhotoListView.IDAssignments - 1)
        .IDRTPhotoReport = itmX.SubItems(GuiPhotoListView.IDRTPhotoReport - 1)
        .IBNUM = Me.IBNUM
        .PhotoID = itmX.SubItems(GuiPhotoListView.ID - 1)
        .cmdLoadAll.Enabled = False
        .cmdMenu(2).Enabled = True
        .lblFileName.Caption = sDeresPhoto
        .txtSortOrder.Text = itmX.SubItems(GuiPhotoListView.SortOrder - 1)
        .txtDescription.Text = itmX.SubItems(GuiPhotoListView.Description - 1)
        .txtPhotoDate = itmX.SubItems(GuiPhotoListView.PhotoDate - 1)
        If goUtil.utFileExists(sDeresPhoto) Then
            On Error Resume Next
            .Image1 = LoadPicture(sDeresPhoto)
            If Err.Number > 0 Then
                Err.Clear
                .Image1 = imgPhotoStatus.ListImages(GuiPhotoStatusList.IsDeleted).Picture
            End If
            On Error GoTo EH
            .SetOriginalPhoto
        End If
        .Show vbModal
    End With
   

    Unload AddPhoto
    Set AddPhoto = Nothing
    If lstvPhotos.Visible Then
        lstvPhotos.SetFocus
    End If
    
    EditPhoto = True
    
    Set oListView = Nothing
    Set itmX = Nothing
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function EditPhoto"
    Unload AddPhoto
    Set AddPhoto = Nothing
End Function

Public Function AddPhotoItem(pudtPhoto As GuiPhotoItem) As Boolean
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    Dim sID As String
    
    Screen.MousePointer = MousePointerConstants.vbHourglass
    
    sID = goUtil.GetAccessDBUID("ID", "RTPhotoLog")
    
    With pudtPhoto
        .RTPhotoLogID = sID
        .RTPhotoReportID = msIDRTPhotoReport
        .AssignmentsID = msAssignmentsID
        'Use Current Billing Item if Available
        .BillingCountID = mfrmClaim.GetCurrentBillingCountID(False)
        .ID = sID
        .IDRTPhotoReport = msIDRTPhotoReport
        .IDAssignments = msAssignmentsID
        .IDBillingCount = mfrmClaim.GetCurrentBillingCountID(True)
    End With
    
    sSQL = "INSERT INTO RTPhotoLog "
    sSQL = sSQL & "( "
    sSQL = sSQL & "[RTPhotoLogID], "
    sSQL = sSQL & "[RTPhotoReportID], "
    sSQL = sSQL & "[AssignmentsID], "
    sSQL = sSQL & "[BillingCountID], "
    sSQL = sSQL & "[ID], "
    sSQL = sSQL & "[IDRTPhotoReport], "
    sSQL = sSQL & "[IDAssignments], "
    sSQL = sSQL & "[IDBillingCount], "
    sSQL = sSQL & "[PhotoDate], "
    sSQL = sSQL & "[SortOrder], "
    sSQL = sSQL & "[Description], "
    sSQL = sSQL & "[PhotoName], "
    sSQL = sSQL & "[Photo], "
    sSQL = sSQL & "[DownloadPhoto], "
    sSQL = sSQL & "[UpLoadPhoto], "
    sSQL = sSQL & "[PhotoThumb], "
    sSQL = sSQL & "[DownloadPhotoThumb], "
    sSQL = sSQL & "[UpLoadPhotoThumb], "
    sSQL = sSQL & "[PhotoHighRes], "
    sSQL = sSQL & "[DownloadPhotoHighRes], "
    sSQL = sSQL & "[UploadPhotoHighRes], "
    sSQL = sSQL & "[IsDeleted], "
    sSQL = sSQL & "[DownLoadMe], "
    sSQL = sSQL & "[UpLoadMe], "
    sSQL = sSQL & "[AdminComments], "
    sSQL = sSQL & "[DateLastUpdated], "
    sSQL = sSQL & "[UpdateByUserID] "
    sSQL = sSQL & ") "
    sSQL = sSQL & "SELECT "
    sSQL = sSQL & pudtPhoto.RTPhotoLogID & " As [RTPhotoLogID], "
    sSQL = sSQL & pudtPhoto.RTPhotoReportID & " As [RTPhotoReportID], "
    sSQL = sSQL & pudtPhoto.AssignmentsID & " As [AssignmentsID], "
    sSQL = sSQL & pudtPhoto.BillingCountID & " As [BillingCountID] , "
    sSQL = sSQL & pudtPhoto.ID & " As [ID], "
    sSQL = sSQL & pudtPhoto.IDRTPhotoReport & " As [IDRTPhotoReport], "
    sSQL = sSQL & pudtPhoto.IDAssignments & " As [IDAssignments], "
    sSQL = sSQL & pudtPhoto.IDBillingCount & " As [IDBillingCount], "
    sSQL = sSQL & "#" & pudtPhoto.PhotoDate & "#" & " As [PhotoDate], "
    sSQL = sSQL & pudtPhoto.SortOrder & " As [SortOrder], "
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtPhoto.Description) & "'" & " As [Description], "
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtPhoto.PhotoName) & "'" & " As [PhotoName], "
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtPhoto.Photo) & "'" & " As [Photo], "
    sSQL = sSQL & pudtPhoto.DownloadPhoto & " As [DownloadPhoto], "
    sSQL = sSQL & pudtPhoto.UpLoadPhoto & " As [UpLoadPhoto], "
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtPhoto.PhotoThumb) & "'" & " As [PhotoThumb], "
    sSQL = sSQL & pudtPhoto.DownloadPhotoThumb & " As [DownloadPhotoThumb], "
    sSQL = sSQL & pudtPhoto.UpLoadPhotoThumb & " As [UpLoadPhotoThumb], "
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtPhoto.PhotoHighRes) & "'" & " As [PhotoHighRes], "
    sSQL = sSQL & pudtPhoto.DownloadPhotoHighRes & " As [DownloadPhotoHighRes], "
    sSQL = sSQL & pudtPhoto.UploadPhotoHighRes & " As [UploadPhotoHighRes], "
    sSQL = sSQL & pudtPhoto.IsDeleted & " As [IsDeleted], "
    sSQL = sSQL & pudtPhoto.DownLoadMe & " As [DownLoadMe], "
    sSQL = sSQL & pudtPhoto.UpLoadMe & " As [UpLoadMe], "
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtPhoto.AdminComments) & "'" & " As [AdminComments], "
    sSQL = sSQL & "#" & pudtPhoto.DateLastUpdated & "#" & " As [DateLastUpdated], "
    sSQL = sSQL & pudtPhoto.UpdateByUserID & " As [UpdateByUserID] "
   
    Set oConn = New ADODB.Connection
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    oConn.Execute sSQL
    
    Screen.MousePointer = MousePointerConstants.vbDefault
    AddPhotoItem = True
    
    'Clean up
    Set oConn = Nothing
    Exit Function
EH:
    Screen.MousePointer = vbDefault
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function AddPhotoItem"
End Function

Public Function EditPhotoItem(pudtPhoto As GuiPhotoItem) As Boolean
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim sSQL As String

    With pudtPhoto
        If .RTPhotoLogID = vbNullString Or .RTPhotoLogID = "0" Then
            .RTPhotoLogID = "Null"
        End If
        If .RTPhotoReportID = vbNullString Or .RTPhotoReportID = "0" Then
            .RTPhotoReportID = msIDRTPhotoReport
        End If
        .AssignmentsID = msAssignmentsID
        'Use Current Billing Item if Available
        If mbRenumberSort Then
            If msRenumSortBillingCountID = vbNullString Then
                .BillingCountID = mfrmClaim.GetCurrentBillingCountID(False)
                msRenumSortBillingCountID = .BillingCountID
            Else
                .BillingCountID = msRenumSortBillingCountID
            End If
        Else
            .BillingCountID = mfrmClaim.GetCurrentBillingCountID(False)
        End If
        
        If .ID = vbNullString Or .ID = "0" Then
            .ID = "Null"
        End If
        If .IDRTPhotoReport = vbNullString Or .IDRTPhotoReport = "0" Then
            .IDRTPhotoReport = msIDRTPhotoReport
        End If
        .IDAssignments = msAssignmentsID
        If mbRenumberSort Then
            If msRenumSortIDBillingCount = vbNullString Then
                .IDBillingCount = mfrmClaim.GetCurrentBillingCountID(True)
                msRenumSortIDBillingCount = .IDBillingCount
            Else
                .IDBillingCount = msRenumSortIDBillingCount
            End If
        Else
            .IDBillingCount = mfrmClaim.GetCurrentBillingCountID(True)
        End If
    End With
    
    sSQL = "UPDATE RTPhotoLog Set "
    If Not mbRenumberSort Then
        sSQL = sSQL & "[RTPhotoLogID] = " & pudtPhoto.RTPhotoLogID & ", "
    End If
        If pudtPhoto.RTPhotoReportID <> "0" Then
            sSQL = sSQL & "[RTPhotoReportID] = " & pudtPhoto.RTPhotoReportID & ", "
        End If
    If Not mbRenumberSort Then
        sSQL = sSQL & "[AssignmentsID] = " & pudtPhoto.AssignmentsID & ", "
        sSQL = sSQL & "[BillingCountID] = " & pudtPhoto.BillingCountID & ", "
        sSQL = sSQL & "[ID] = " & pudtPhoto.ID & ", "
    End If
    If pudtPhoto.IDRTPhotoReport <> "0" Then
        sSQL = sSQL & "[IDRTPhotoReport] = " & pudtPhoto.IDRTPhotoReport & ", "
    End If
    If Not mbRenumberSort Then
        sSQL = sSQL & "[IDAssignments] = " & pudtPhoto.IDAssignments & ", "
        sSQL = sSQL & "[IDBillingCount] = " & pudtPhoto.IDBillingCount & ", "
        sSQL = sSQL & "[PhotoDate] = #" & pudtPhoto.PhotoDate & "#, "
    End If
    sSQL = sSQL & "[SortOrder] = " & pudtPhoto.SortOrder & ", "
    If Not mbRenumberSort Then
        sSQL = sSQL & "[Description] = '" & goUtil.utCleanSQLString(pudtPhoto.Description) & "', "
        sSQL = sSQL & "[PhotoName] = '" & goUtil.utCleanSQLString(pudtPhoto.PhotoName) & "', "
        sSQL = sSQL & "[Photo] = '" & goUtil.utCleanSQLString(pudtPhoto.Photo) & "', "
        sSQL = sSQL & "[DownloadPhoto] = " & pudtPhoto.DownloadPhoto & ", "
        sSQL = sSQL & "[UpLoadPhoto] = " & pudtPhoto.UpLoadPhoto & ", "
        sSQL = sSQL & "[PhotoThumb] = '" & goUtil.utCleanSQLString(pudtPhoto.PhotoThumb) & "', "
        sSQL = sSQL & "[DownloadPhotoThumb] = " & pudtPhoto.DownloadPhotoThumb & ", "
        sSQL = sSQL & "[UpLoadPhotoThumb] = " & pudtPhoto.UpLoadPhotoThumb & ", "
        sSQL = sSQL & "[PhotoHighRes] = '" & goUtil.utCleanSQLString(pudtPhoto.PhotoHighRes) & "', "
        sSQL = sSQL & "[DownloadPhotoHighRes] = " & pudtPhoto.DownloadPhotoHighRes & ", "
        sSQL = sSQL & "[UploadPhotoHighRes] = " & pudtPhoto.UploadPhotoHighRes & ", "
        sSQL = sSQL & "[IsDeleted] = " & pudtPhoto.IsDeleted & ", "
        sSQL = sSQL & "[DownLoadMe] = " & pudtPhoto.DownLoadMe & ", "
    End If
    sSQL = sSQL & "[UpLoadMe] = " & pudtPhoto.UpLoadMe & ", "
    If Not mbRenumberSort Then
        sSQL = sSQL & "[AdminComments] = '" & goUtil.utCleanSQLString(pudtPhoto.AdminComments) & "', "
    End If
    sSQL = sSQL & "[DateLastUpdated] = #" & pudtPhoto.DateLastUpdated & "#, "
    sSQL = sSQL & "[UpdateByUserID]  = " & pudtPhoto.UpdateByUserID & " "
    sSQL = sSQL & "WHERE IDAssignments = " & pudtPhoto.IDAssignments & " "
    sSQL = sSQL & "AND [ID] = " & pudtPhoto.ID & " "
    If mbRenumberSort Then
        sSQL = sSQL & "AND [SortOrder] <> " & pudtPhoto.SortOrder & " "
    End If

    Set oConn = New ADODB.Connection
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    oConn.Execute sSQL
    
    EditPhotoItem = True
    'Clean up
    Set oConn = Nothing
    Exit Function
    
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function EditPhotoItem"
End Function

Public Function GetMaxSort(psIDAssignments As String, psIDRTPhotoReport As String) As Long
    On Error GoTo EH
    Dim RS As ADODB.Recordset
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    
    Set oConn = New ADODB.Connection
    Set RS = New ADODB.Recordset
    
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    sSQL = "SELECT MAX(PL.SortOrder) As MaxSort "
    sSQL = sSQL & "FROM RTPhotoLog PL "
    sSQL = sSQL & "WHERE PL.IDAssignments = " & psIDAssignments & " "
    If cmdAddMultiReport.Visible Then
        sSQL = sSQL & "AND PL.IDRTPhotoReport = " & psIDRTPhotoReport & " "
    End If
    
    'Use Disconnected Record Set on asUseClient Cusor ONLY !
    RS.CursorLocation = adUseClient
    RS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set RS.ActiveConnection = Nothing
    
    If Not RS.EOF Then
        RS.MoveFirst
        GetMaxSort = IIf(IsNull(RS!MaxSort), 0, RS!MaxSort)
    End If
    
    Set oConn = Nothing
    Set RS = Nothing
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function GetMaxSort"
End Function

Public Function GetPhotoCount(psIDAssignments As String, psIDRTPhotoReport As String) As Long
    On Error GoTo EH
    Dim RS As ADODB.Recordset
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    
    Set oConn = New ADODB.Connection
    Set RS = New ADODB.Recordset
    
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    sSQL = "SELECT Count(PL.RTPhotoLogID) As PhotoCount "
    sSQL = sSQL & "FROM RTPhotoLog PL "
    sSQL = sSQL & "WHERE PL.IDAssignments = " & psIDAssignments & " "
    sSQL = sSQL & "AND PL.IDRTPhotoReport = " & psIDRTPhotoReport & " "
    
    'Use Disconnected Record Set on asUseClient Cusor ONLY !
    RS.CursorLocation = adUseClient
    RS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set RS.ActiveConnection = Nothing
    
    If Not RS.EOF Then
        RS.MoveFirst
        GetPhotoCount = IIf(IsNull(RS!PhotoCount), 0, RS!PhotoCount)
    End If
    
    Set oConn = Nothing
    Set RS = Nothing
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function GetPhotoCount"
End Function

Public Function PrintPhotos(psIDAssignments As String, psIDRTPhotoReport As String) As Boolean
    On Error GoTo EH
    Dim MyPhoto As Object
    Dim lrptVersion As Long
    Dim sParams As String
    Dim lMainSPVersion As Long
    Dim oConn As ADODB.Connection
    Dim oCarList As V2ECKeyBoard.clsCarLists
    'If using Adobe PDF Viewer
    Dim sPDFFilePath As String
    Dim sPrintPreview As String
    Dim bUseAdobeReader As Boolean
    Dim sReportName As String
    Dim sProjectName As String
    Dim sClassName As String
    Dim adoRSApplication As ADODB.Recordset
    
    'Photo Reports (Multi Report)
    Dim sPhotoReportNumber As String
    Dim RS As ADODB.Recordset
    Dim sSQL As String
    'Section Levels (Application Software Table)
    Dim sSL01 As String
    Dim sSL02 As String
    Dim sSL03 As String
    Dim sSL04 As String
    Dim sSL05 As String
    Dim sSL06 As String
    Dim sSL07 As String
    Dim sSL08 As String
    Dim sSL09 As String
    Dim sSL10 As String
    
    'Need to populate the Section Levels via Project name Lookup
    mfrmClaim.PopulateSectionLevels msAssignmentsID, _
                                    "_arRptPhotos", _
                                    sSL01, _
                                    sSL02, _
                                    sSL03, _
                                    sSL04, _
                                    sSL05, _
                                    sSL06, _
                                    sSL07, _
                                    sSL08, _
                                    sSL09, _
                                    sSL10
    
    Set adoRSApplication = mfrmClaim.GetadoRSApplication(msAssignmentsID, sSL01, sSL02, sSL03, sSL04, sSL05)
    
    sProjectName = goUtil.IsNullIsVbNullString(adoRSApplication.Fields("ProjectName"))
    sClassName = goUtil.IsNullIsVbNullString(adoRSApplication.Fields("ClassName"))
    
    sReportName = sProjectName & "." & sClassName
    
    Set oConn = New ADODB.Connection
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    'get the Photo Report Num for the passsed in ID
    sSQL = "SELECT  [Number] "
    sSQL = sSQL & "FROM     RTPhotoReport "
    sSQL = sSQL & "WHERE    ID = " & psIDRTPhotoReport & " "
    
    'Use Disconnected Record Set on asUseClient Cusor ONLY !
    Set RS = New ADODB.Recordset
    RS.CursorLocation = adUseClient
    RS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set RS.ActiveConnection = Nothing
    
    If RS.RecordCount = 1 Then
        sPhotoReportNumber = goUtil.IsNullIsVbNullString(RS.Fields("Number"))
    Else
        If cmdAddMultiReport.Visible Then
            GoTo CLEAN_UP
        End If
    End If
    
    sPrintPreview = GetSetting(goUtil.gsMainAppEXEName, "GENERAL", "PRINT_PREVIEW", "USE_ADOBE")
    
    Select Case UCase(sPrintPreview)
        Case "USE_ADOBE"
            bUseAdobeReader = True
    End Select
    
    lMainSPVersion = mfrmClaim.adoRSAssignments.Fields("SPVersion").Value
    
    
    lrptVersion = goUtil.GetApplicationVersionNumber(lMainSPVersion, sProjectName, oConn)
    
    sParams = sParams & "psAssignmentsID=" & psIDAssignments & "|"
    sParams = sParams & "pPhotoReposPath=" & goUtil.PhotoReposPath & "|"
    sParams = sParams & "pbPreview=" & "True" & "|"
    If bUseAdobeReader Then
        sPDFFilePath = goUtil.gsInstallDir & "\TempActiveReport" & goUtil.utGetTickCount & ".pdf"
        sParams = sParams & "psXportPath=" & sPDFFilePath & "|"
        sParams = sParams & "pPDFJPEGQuality=" & "50" & "|"
        sParams = sParams & "pXportType=" & ExportType.ARPdf & "|"
    Else
        sParams = sParams & "pbGetObjectOnly=" & "True" & "|"
    End If
    'Photo Reports (Multi Report)
    If sPhotoReportNumber <> vbNullString Then
        sParams = sParams & "pNumber=" & sPhotoReportNumber & "|"
    Else
        sParams = sParams & "pNumber=0" & "|"
    End If

    Set oCarList = CreateObject(goUtil.goCurCarList.ClassName)
    
    If bUseAdobeReader Then
        oCarList.GetARReport sReportName, lrptVersion, sParams
        If goUtil.utFileExists(sPDFFilePath) Then
            goUtil.utShellExecute , "OPEN", sPDFFilePath, , , vbNormalFocus, True, True, True, "Activity Log"
            DoEvents
            Sleep 1000
            goUtil.utDeleteFile sPDFFilePath
            oCarList.CLEANUP
            Set oCarList = Nothing
        End If
    Else
    
        Set MyPhoto = oCarList.GetARReport(sReportName, lrptVersion, sParams)
        
        
        If mArv Is Nothing Then
            Set mArv = New V2ARViewer.clsARViewer
            mArv.SetUtilObject goUtil
        End If
        
        If Not moForm Is Nothing Then
            Unload moForm
            Set moForm = Nothing
        End If
        
        With mArv
            'Pass in true to have Active reports process on separate thread.
            'This will allow the viewer to load while the report is processing
            'false will force the report to run on single thread
            MyPhoto.Run False 'True
            .objARvReport = MyPhoto
            .sRptTitle = "Photos"
            .HidePrintButton = False
            .ShowReportOnForm moForm, vbModeless
            Unload .objARvReport
            Set .objARvReport = Nothing
            Unload MyPhoto
            Set MyPhoto = Nothing
            oCarList.CLEANUP
            Set oCarList = Nothing
        End With
    End If
    
    'Clear the Local ref to this report object only
    'The actual cleanup of this active report object will occur within gARV
    PrintPhotos = True
    
CLEAN_UP:
    Set MyPhoto = Nothing
    Set oConn = Nothing
    Set oCarList = Nothing
    Set RS = Nothing
    Set adoRSApplication = Nothing
    
    Exit Function
EH:
    Screen.MousePointer = vbDefault
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function PrintPhotos"
End Function

Public Sub PopReportIDBySort(lSortOrder As Long, psRTPhotoReportID As String, psIDRTPhotoReport As String)
    On Error GoTo EH
    Dim lNumber As Long
    Dim sSQL As String
    Dim oConn As ADODB.Connection
    Dim RS As ADODB.Recordset
    
    'Need to Get the Correct Number to query the Correct Report ID
    For lNumber = 1 To 51
        If lSortOrder <= (MAX_PHOTOS_ALLOWED * lNumber) Then
            Exit For
        End If
    Next
    If lNumber = 51 Then
        GoTo CLEAN_UP
    End If
    
    Set oConn = New ADODB.Connection
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    'get the Photo Report Num for the passsed in ID
    sSQL = "SELECT  [RTPhotoReportID], "
    sSQL = sSQL & "[ID] "
    sSQL = sSQL & "FROM     RTPhotoReport "
    sSQL = sSQL & "WHERE    [AssignmentsID] = " & msAssignmentsID & " "
    sSQL = sSQL & "AND      [Number] = " & lNumber & " "
    
    'Use Disconnected Record Set on asUseClient Cusor ONLY !
    Set RS = New ADODB.Recordset
    RS.CursorLocation = adUseClient
    RS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set RS.ActiveConnection = Nothing
    
    If RS.RecordCount = 1 Then
        RS.MoveFirst
        psRTPhotoReportID = RS.Fields("RTPhotoReportID")
        psIDRTPhotoReport = RS.Fields("ID")
    End If
    
CLEAN_UP:
    'cleanup
    Set oConn = Nothing
    Set RS = Nothing
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub PopReportIDBySort"
End Sub

Public Sub ReNumberPhotoSort()
    On Error GoTo EH
    Dim oListView As MSComctlLib.ListView
    Dim itmX  As MSComctlLib.ListItem
    Dim lCount As Long
    Dim udtPhoto As GuiPhotoItem
    Dim sRTPhotoReportID As String
    Dim sIDRTPhotoReport As String
    
    Set oListView = lstvPhotos
    lCount = 0
    mbRenumberSort = True
    msRenumSortBillingCountID = vbNullString
    msRenumSortIDBillingCount = vbNullString
    For Each itmX In oListView.ListItems
        lCount = lCount + 1
        
        'Populate the Photo Report ID
        sRTPhotoReportID = itmX.SubItems(GuiPhotoListView.RTPhotoReportID - 1) 'Set Default value
        sIDRTPhotoReport = itmX.SubItems(GuiPhotoListView.IDRTPhotoReport - 1) 'Set Default value
        PopReportIDBySort lCount, sRTPhotoReportID, sIDRTPhotoReport
        
        With udtPhoto
            .RTPhotoReportID = sRTPhotoReportID
            .ID = itmX.SubItems(GuiPhotoListView.ID - 1)
            .IDRTPhotoReport = sIDRTPhotoReport
            .IDAssignments = itmX.SubItems(GuiPhotoListView.IDAssignments - 1)
            .SortOrder = lCount
            .UpLoadMe = "True"
            .DateLastUpdated = Format(Now(), "MM/DD/YYYY HH:MM:SS")
            .UpdateByUserID = goUtil.gsCurUsersID
        End With
        If EditPhotoItem(udtPhoto) Then
            itmX.Text = lCount
            itmX.SubItems(GuiPhotoListView.SortOrder - 1) = lCount
        Else
            Exit For
        End If
    Next
    
    If Not mbUnloadMe Then
        LoadMe
    End If
    
    Set itmX = Nothing
    Set oListView = Nothing
    mbRenumberSort = False
    Exit Sub
EH:
    mbRenumberSort = False
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub ReNumberPhotoSort"
End Sub
