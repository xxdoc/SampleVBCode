VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form EditPayReq 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payment Request"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5910
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "EditPayReq.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   5910
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstvbUserDefinedType 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5340
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   5655
   End
   Begin VB.Timer Timer_GetAmount 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   240
      Top             =   5760
   End
   Begin VB.CheckBox chkAddress 
      Caption         =   "..."
      Height          =   340
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Reset Address"
      Top             =   5100
      Width           =   375
   End
   Begin VB.TextBox txtAddress 
      Height          =   360
      Left            =   240
      MaxLength       =   255
      TabIndex        =   25
      Tag             =   "UCASE"
      Top             =   5107
      Width           =   5535
   End
   Begin VB.Frame framCommands 
      Height          =   1215
      Left            =   3480
      TabIndex        =   45
      Top             =   5520
      Width           =   2295
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   855
         Left            =   1200
         MaskColor       =   &H00000000&
         Picture         =   "EditPayReq.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Exit"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Sa&ve"
         Default         =   -1  'True
         Enabled         =   0   'False
         Height          =   855
         Left            =   120
         MaskColor       =   &H00000000&
         Picture         =   "EditPayReq.frx":0316
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Exit"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdSpelling 
         Caption         =   "Spellin&g"
         Height          =   855
         Left            =   120
         MaskColor       =   &H00000000&
         Picture         =   "EditPayReq.frx":0460
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Exit"
         Top             =   240
         Visible         =   0   'False
         Width           =   975
      End
   End
   Begin VB.Frame framEdit 
      Height          =   4695
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   5655
      Begin VB.TextBox txtAppliedDeductible 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
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
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   19
         ToolTipText     =   "Applied Deductible Totals"
         Top             =   3960
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txtIndemTotals 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
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
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   17
         ToolTipText     =   "Associated Indemnity Totals"
         Top             =   3960
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CheckBox chkViewCalc 
         Height          =   340
         Left            =   480
         Picture         =   "EditPayReq.frx":08AA
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "View Calculator"
         Top             =   4200
         Width           =   375
      End
      Begin VB.CheckBox chkAmountOfCheck 
         Caption         =   "..."
         Height          =   340
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Calculate Amount Of Check"
         Top             =   4200
         Width           =   375
      End
      Begin VB.CheckBox chkInsured 
         Caption         =   "..."
         Height          =   340
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Reset Insured Name"
         Top             =   3000
         Width           =   375
      End
      Begin VB.CheckBox chkIncludeMortgagee 
         Caption         =   "..."
         Height          =   340
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Include Mortgagee on Draft"
         Top             =   3600
         Width           =   375
      End
      Begin VB.CheckBox chkPrintOnIB 
         Caption         =   "Print On IB (Internal Billing)"
         Enabled         =   0   'False
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
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Visible         =   0   'False
         Width           =   3975
      End
      Begin VB.TextBox txtAmountOfCheck 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   120
         MaxLength       =   14
         TabIndex        =   23
         Tag             =   "Currency"
         Top             =   4200
         Width           =   5415
      End
      Begin VB.ComboBox cboCatCode 
         Height          =   360
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1200
         Width           =   5415
      End
      Begin VB.TextBox txtMortgageeName 
         Height          =   360
         Left            =   120
         MaxLength       =   100
         TabIndex        =   14
         Tag             =   "UCASE"
         Top             =   3600
         Width           =   5415
      End
      Begin VB.TextBox txtInsured 
         Height          =   360
         Left            =   120
         MaxLength       =   100
         TabIndex        =   11
         Tag             =   "UCASE"
         Top             =   3000
         Width           =   5415
      End
      Begin VB.ComboBox cboClassOfLoss 
         CausesValidation=   0   'False
         Height          =   360
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1800
         Width           =   5415
      End
      Begin VB.ComboBox cboTypeOfLoss 
         CausesValidation=   0   'False
         Height          =   360
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2400
         Width           =   5415
      End
      Begin VB.CommandButton cmdFindNext 
         Caption         =   "Find &Next"
         Height          =   375
         Left            =   8400
         TabIndex        =   29
         Top             =   240
         Visible         =   0   'False
         Width           =   1100
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Find"
         Height          =   375
         Left            =   7200
         TabIndex        =   28
         Top             =   240
         Visible         =   0   'False
         Width           =   1100
      End
      Begin VB.ComboBox cboPayments 
         Enabled         =   0   'False
         Height          =   360
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   600
         Width           =   5415
      End
      Begin VB.CommandButton cmdSelAll 
         Caption         =   "&Select A&ll"
         Height          =   375
         Left            =   6000
         TabIndex        =   27
         Top             =   240
         Visible         =   0   'False
         Width           =   1100
      End
      Begin VB.TextBox txtSpellMe 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   6840
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   960
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.CommandButton cmdPrintReport 
         Caption         =   "&Print"
         Enabled         =   0   'False
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
         Left            =   10200
         TabIndex        =   30
         Top             =   240
         Visible         =   0   'False
         Width           =   975
      End
      Begin MSComctlLib.ImageList imgRptParamsStatus 
         Left            =   6120
         Top             =   960
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
               Picture         =   "EditPayReq.frx":0E34
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "EditPayReq.frx":0F8E
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "EditPayReq.frx":137A
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "EditPayReq.frx":14EB
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "EditPayReq.frx":1919
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "EditPayReq.frx":1D6B
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Frame framMultiUpdate 
         Caption         =   "Multi Update"
         Height          =   735
         Left            =   5880
         TabIndex        =   39
         Top             =   3840
         Visible         =   0   'False
         Width           =   5535
         Begin VB.CommandButton cmdUpdateMulti 
            Caption         =   "&Update"
            Height          =   375
            Index           =   0
            Left            =   1800
            TabIndex        =   41
            Top             =   240
            Width           =   855
         End
         Begin VB.CheckBox chkMultiUpdate 
            Caption         =   "Uncheck Selected Checkable Items"
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
            TabIndex        =   40
            Top             =   240
            Width           =   1695
         End
         Begin VB.CommandButton cmdDateMultiUpdate 
            Height          =   375
            Left            =   4080
            Picture         =   "EditPayReq.frx":2131
            Style           =   1  'Graphical
            TabIndex        =   43
            Tag             =   "Date"
            Top             =   240
            Width           =   375
         End
         Begin VB.CommandButton cmdUpdateMulti 
            Caption         =   "&Update"
            Height          =   375
            Index           =   1
            Left            =   4560
            TabIndex        =   44
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox txtDateMultiUpdate 
            Height          =   375
            Left            =   2880
            TabIndex        =   42
            Tag             =   "Date"
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame framEditParam 
         Height          =   3255
         Left            =   5880
         TabIndex        =   32
         Top             =   600
         Visible         =   0   'False
         Width           =   5535
         Begin VB.CommandButton cmdUpdateEdit 
            Caption         =   "&Update"
            Height          =   375
            Left            =   4560
            TabIndex        =   38
            Top             =   2760
            Width           =   855
         End
         Begin VB.CheckBox chkParamBoolean 
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   2400
            Visible         =   0   'False
            Width           =   4095
         End
         Begin VB.CommandButton cmdParamDate 
            Height          =   375
            Left            =   4080
            Picture         =   "EditPayReq.frx":2573
            Style           =   1  'Graphical
            TabIndex        =   36
            Tag             =   "Date"
            Top             =   2760
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.CommandButton cmdCancelEdit 
            Caption         =   "&Cancel"
            Height          =   375
            Left            =   4560
            TabIndex        =   37
            Top             =   2340
            Width           =   855
         End
         Begin VB.TextBox txtParamCaption 
            Appearance      =   0  'Flat
            Height          =   1935
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   240
            Visible         =   0   'False
            Width           =   5295
         End
         Begin VB.TextBox txtParamValue 
            Height          =   375
            Left            =   120
            MaxLength       =   300
            TabIndex        =   35
            Top             =   2760
            Visible         =   0   'False
            Width           =   4335
         End
      End
      Begin MSComctlLib.ListView lvwRptParams 
         Height          =   3255
         Left            =   5880
         TabIndex        =   31
         Top             =   600
         Visible         =   0   'False
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   5741
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "imgRptParamsStatus"
         ColHdrIcons     =   "imgRptParamsStatus"
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
         NumItems        =   0
      End
      Begin VB.Label lblEquals 
         Alignment       =   2  'Center
         Caption         =   "="
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5280
         TabIndex        =   20
         Top             =   3960
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lbMinus 
         Alignment       =   2  'Center
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   18
         Top             =   3960
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblAmountOfCheck 
         Caption         =   "Amount Of Check:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   3960
         Width           =   1695
      End
      Begin VB.Label lblCatCode 
         Caption         =   "Cat Code:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   3255
      End
      Begin VB.Label lblMortgageeName 
         Caption         =   "Other Payee Name(s):"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   3360
         Width           =   5415
      End
      Begin VB.Label lblInsured 
         Caption         =   "Insured Payee Name(s):"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   2760
         Width           =   2535
      End
      Begin VB.Label lblClassOfLoss 
         Caption         =   "Line Of Coverage (Class):"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1560
         Width           =   4455
      End
      Begin VB.Label lblTypeOfLoss 
         Caption         =   "Type Of Loss:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   2160
         Width           =   3255
      End
   End
   Begin VB.Label lblAddress 
      Caption         =   "Address"
      Height          =   255
      Left            =   240
      TabIndex        =   24
      Top             =   4800
      Width           =   5415
   End
End
Attribute VB_Name = "EditPayReq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmClaim As frmClaim
Private mfrmIndemnity As frmIndemnity
Private moGUI As V2ECKeyBoard.clsCarGUI
Private msAssignmentsID As String
Private mbUnloadMe As Boolean
Private moActivecboReport As ComboBox
Private msFindText As String
Private mlLastFindIndex As Long
Private mbShowingEditRptParam As Boolean
Private msPayReqID As String
Private mbLoading As Boolean
Private msInsured As String
Private msMortgagee As String
Private msAddress As String
Private mlCOLOrigListIndex As Long
Private mbAssociateFromIndem As Boolean
Private mbEditGetAmount As Boolean
Private mbSaveEnabled As Boolean

Public Property Let EditGetAmount(pbFlag As Boolean)
    mbEditGetAmount = pbFlag
End Property
Public Property Get EditGetAmount() As Boolean
    EditGetAmount = mbEditGetAmount
End Property

Public Property Let AssociateFromIndem(pbFlag As Boolean)
    mbAssociateFromIndem = pbFlag
End Property
Public Property Get AssociateFromIndem() As Boolean
    AssociateFromIndem = mbAssociateFromIndem
End Property

Public Property Let COLOrigListIndex(pIndex As Long)
    mlCOLOrigListIndex = pIndex
End Property
Public Property Get COLOrigListIndex() As Long
    COLOrigListIndex = mlCOLOrigListIndex
End Property

Public Property Let Address(psAddress As String)
    msAddress = psAddress
End Property
Public Property Get Address() As String
    Address = msAddress
End Property

Public Property Let Mortgagee(psMortgagee As String)
    msMortgagee = psMortgagee
End Property
Public Property Get Mortgagee() As String
    Mortgagee = msMortgagee
End Property

Public Property Let Insured(psInsured As String)
    msInsured = psInsured
End Property
Public Property Get Insured() As String
    Insured = msInsured
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

Public Property Let MyfrmClaim(poForm As Object)
    Set mfrmClaim = poForm
End Property
Public Property Set MyfrmClaim(poForm As Object)
    Set mfrmClaim = poForm
End Property
Public Property Get MyfrmClaim() As Object
    Set MyfrmClaim = mfrmClaim
End Property

Public Property Let MyIndemnity(poForm As Object)
    Set mfrmIndemnity = poForm
End Property
Public Property Set MyIndemnity(poForm As Object)
    Set mfrmIndemnity = poForm
End Property
Public Property Get MyIndemnity() As Object
    Set MyIndemnity = mfrmIndemnity
End Property

Public Property Let PayReqID(psID As String)
    msPayReqID = psID
End Property

Public Property Let Loading(pbFlag As Boolean)
    mbLoading = pbFlag
End Property
Public Property Get Loading() As Boolean
    Loading = mbLoading
End Property

Public Property Let UnloadMe(pbFlag As Boolean)
    mbUnloadMe = pbFlag
End Property
Public Property Get UnloadMe() As Boolean
    UnloadMe = mbUnloadMe
End Property

Private Property Get msClassName() As String
    msClassName = Me.Name
End Property



Private Sub cboCatCode_Click()
    On Error GoTo EH
    
    If mbLoading Then
        Exit Sub
    End If
    
    cmdSave.Enabled = True
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cboCatCode_Click"
End Sub

Private Sub cboClassOfLoss_Click()
    On Error GoTo EH
    
    If mbLoading Then
        Exit Sub
    End If
    
    cmdSave.Enabled = True
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cboClassOfLoss_Click"
End Sub

Private Sub cboTypeOfLoss_Click()
    On Error GoTo EH
    
    If mbLoading Then
        Exit Sub
    End If
    
    cmdSave.Enabled = True
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cboTypeOfLoss_Click"
End Sub

Private Sub chkAddress_Click()
    On Error GoTo EH
    
    If chkAddress.Value = vbChecked Then
        txtAddress.Text = msAddress
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub chkAddress_Click"
End Sub

Private Sub chkAmountOfCheck_Click()
    On Error GoTo EH
    Dim sClassOfLossID As String
    Dim RS As ADODB.Recordset
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    Dim sClassTypeID As String
    Dim cAppliedDeductible As Currency
    Dim cAmountOfCheck As Currency
    Dim itmX As ListItem
    Dim oListView As ListView
    
    If chkAmountOfCheck.Value = vbUnchecked Then
        Exit Sub
    End If
    
    If cboClassOfLoss.ListIndex = -1 Then
        MsgBox "Must Select " & lblClassOfLoss.Caption & "! ", vbExclamation + vbOKOnly, "Amount Of Check"
        Exit Sub
    End If
    
    'Class of Loss
    sClassTypeID = cboClassOfLoss.ItemData(cboClassOfLoss.ListIndex)
    sSQL = "SELECT COL.[ClassOfLossID], "
    sSQL = sSQL & "(CT.Description + ' (' + COL.Description + ')') As ClassOfLoss, "
    sSQL = sSQL & "COL.Code As ClassOfLossCode "
    sSQL = sSQL & "FROM ClassOfLoss COL "
    sSQL = sSQL & "INNER JOIN ClassType CT ON CT.ClassTypeID = COL.ClassTypeID "
    sSQL = sSQL & "WHERE COL.[ClientCompanyID] = " & goUtil.gsCurCar & " "
    sSQL = sSQL & "AND COL.[ClassTypeID] = " & sClassTypeID & " "
    
    Set oConn = New ADODB.Connection
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    Set RS = New ADODB.Recordset
    RS.CursorLocation = adUseClient
    RS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set RS.ActiveConnection = Nothing
    
    If RS.RecordCount = 1 Then
        RS.MoveFirst
        sClassOfLossID = goUtil.IsNullIsVbNullString(RS.Fields("ClassOfLossID"))
    End If
    
    With AssociatePayReqToIndem
        .MyIndemnity = mfrmIndemnity
        .MyfrmClaim = mfrmClaim
        .MyGUI = moGUI
        .AssignmentsID = msAssignmentsID
        .PayReqID = msPayReqID
        .SelRTChecksID = CLng(msPayReqID)
        .SelClassOfLossID = sClassOfLossID
        .AssociateFromIndem = mbAssociateFromIndem
        .EditGetAmount = mbEditGetAmount
        Load AssociatePayReqToIndem
        .Caption = "Associate Payment Request to selected Indemnity Items:"
        .WindowState = vbNormal
        .Show vbModal
    End With
    
    'Total up the Amount of Check and Applied deductibles
    Set oListView = AssociatePayReqToIndem.lvwIndemnity
    For Each itmX In oListView.ListItems
        If itmX.SubItems(GuiIndemListView.IDRTChecks - 1) = msPayReqID Then
            cAmountOfCheck = cAmountOfCheck + CCur(itmX.SubItems(GuiIndemListView.ACVLessExcessLimits - 1))
            cAppliedDeductible = cAppliedDeductible + CCur(itmX.SubItems(GuiIndemListView.AppliedDeductible - 1))
        End If
    Next
    
    AssociatePayReqToIndem.CLEANUP
    Unload AssociatePayReqToIndem
    Set AssociatePayReqToIndem = Nothing
    
    txtIndemTotals.Visible = True
    lbMinus.Visible = True
    txtAppliedDeductible.Visible = True
    lblEquals.Visible = True
    txtIndemTotals.Text = Format(cAmountOfCheck, "#,###,###,##0.00")
    txtAppliedDeductible.Text = Format(cAppliedDeductible, "#,###,###,##0.00")
    cAmountOfCheck = cAmountOfCheck - cAppliedDeductible
    If cAmountOfCheck < 0 Then
        cAmountOfCheck = 0
    End If
    
    txtAmountOfCheck.Text = Format(cAmountOfCheck, "#,###,###,##0.00")
    
    'cleanup
    Set RS = Nothing
    Set oConn = Nothing
    Set itmX = Nothing
    Set oListView = Nothing
    
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub chkAmountOfCheck_Click"
End Sub

Private Sub chkIncludeMortgagee_Click()
    On Error GoTo EH
    
    If chkIncludeMortgagee.Value = vbChecked Then
        txtMortgageeName.Text = msMortgagee
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub chkIncludeMortgagee_Click"
End Sub

Private Sub chkInsured_Click()
    On Error GoTo EH
    
    If chkInsured.Value = vbChecked Then
        txtInsured.Text = msInsured
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub chkInsured_Click"
End Sub

Private Sub chkMultiUpdate_Click()
    On Error GoTo EH
    
    If chkMultiUpdate.Value = vbChecked Then
        chkMultiUpdate.Caption = "Check Selected Checkable Items"
    Else
        chkMultiUpdate.Caption = "Uncheck Selected Checkable Items"
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub chkMultiUpdate_Click"
End Sub

Private Sub chkParamBoolean_Click()
    On Error GoTo EH
    
    If mbShowingEditRptParam Then
        Exit Sub
    End If
    
    UpdateEdit
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub chkParamBoolean_Click"
End Sub

Private Sub chkPrintOnIB_Click()
    cmdSave.Enabled = True
End Sub

Private Sub chkViewCalc_Click()
    On Error GoTo EH
    
    If chkViewCalc.Value = vbChecked Then
        Shell "calc.exe ", vbNormalFocus
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub chkViewCalc_Click"
End Sub

Private Sub cmdCancelEdit_Click()
    On Error GoTo EH
    'Set Ecit button back to the defualt Cancel
    cmdExit.Cancel = True
    cmdUpdateEdit.Default = False
    framEditParam.Visible = False
    lvwRptParams.Visible = True
    cmdSelAll.Enabled = True
    cmdFind.Enabled = True
    cmdFindNext.Enabled = True
    cmdPrintReport.Enabled = True
'    framMultiUpdate.Visible = True
    
    cmdSpelling.Enabled = True
    lvwRptParams.SetFocus
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdCancelEdit_Click"
End Sub

Private Sub cmdDateMultiUpdate_Click()
    On Error GoTo EH
    
    MyGUI.ShowCalendar txtDateMultiUpdate
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdDateMultiUpdate_Click"
End Sub

Private Sub cmdExit_Click()
    On Error GoTo EH
    If cmdSave.Enabled Then
'        If MsgBox("Do you want to save changes?", vbYesNo + vbQuestion, "Save Changes") = vbYes Then
            If Not SaveMe() Then
                Exit Sub
            End If
'        End If
    ElseIf mlCOLOrigListIndex = -1 Then
        cmdSave.Enabled = True
    End If
    mbUnloadMe = True
    goUtil.utFormWinRegPos goUtil.gsMainAppEXEName, Me, True, , , , True
    Me.Visible = False
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdExit_Click"
End Sub

Private Sub cmdFind_Click()
    On Error GoTo EH
    If lvwRptParams.ListItems.Count > 0 Then
        mlLastFindIndex = 0
        goUtil.utFindListItem Me, lvwRptParams, msFindText, mlLastFindIndex
    End If
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdFind_Click"
End Sub

Private Sub cmdFindNext_Click()
    On Error GoTo EH
    
    If msFindText <> vbNullString And lvwRptParams.ListItems.Count > 0 Then
        goUtil.utFindListItem Me, lvwRptParams, msFindText, mlLastFindIndex
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdFindNext_Click"
End Sub

Private Sub cmdParamDate_Click()
    On Error GoTo EH
    
    MyGUI.ShowCalendar txtParamValue
    UpdateEdit
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdParamDate_Click"
End Sub


Private Sub cmdPrintReport_Click()
    On Error GoTo EH
    
    If cmdSave.Enabled Then
'        If MsgBox("Do you want to save changes?", vbYesNo + vbQuestion, "Save Changes") = vbYes Then
            If Not SaveMe() Then
                Exit Sub
            Else
                cmdSave.Enabled = False
            End If
'        End If
    End If
    'First be sure control are valid
    cmdPrintReport.Enabled = False
    If mfrmClaim.PrintActiveReport(cboPayments, vbModal) Then
        If Not mbUnloadMe Then
            cmdPrintReport.Enabled = True
            cmdPrintReport.Enabled = True
        End If
    End If


    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdPrintReport_Click"
End Sub

Private Sub cmdSave_Click()
    On Error GoTo EH
    
    If SaveMe() Then
        goUtil.utFormWinRegPos goUtil.gsMainAppEXEName, Me, True, , , , True
        Me.Visible = False
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdSave_Click"
End Sub

Private Sub cmdSelAll_Click()
    On Error GoTo EH
    Dim itmX As ListItem
    For Each itmX In lvwRptParams.ListItems
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
    Dim sFlagText As String
    Dim saryText() As String
    Dim lPos As Long
    Dim iDataType As VBA.VbVarType
    Dim MyParam As V2ECKeyBoard.MiscReportParam
    
    txtSpellMe.Text = vbNullString
    
    For Each itmX In lvwRptParams.ListItems
        sText = sText & itmX.SubItems(GuiRptParamsListView.ParamValue - 1) & vbCrLf
    Next
    'take off the last VBCRLF
    If sText <> vbNullString Then
        sText = left(sText, InStrRev(sText, vbCrLf, , vbBinaryCompare) - 1)
    Else
        Exit Sub
    End If
    Screen.MousePointer = MousePointerConstants.vbHourglass
    
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
        Set itmX = lvwRptParams.ListItems(lPos + 1)
        'Only update string values
        iDataType = itmX.SubItems(GuiRptParamsListView.ParamDataType - 1)
        If iDataType <> vbString Then
            GoTo NEXT_ITEM
        End If
        If StrComp(sText, itmX.SubItems(GuiRptParamsListView.ParamValue - 1), vbTextCompare) <> 0 Then
            With MyParam
                .MiscReportParamID = itmX.SubItems(GuiRptParamsListView.MiscReportParamID - 1)
                .AssignmentsID = itmX.SubItems(GuiRptParamsListView.AssignmentsID - 1)
                .ID = itmX.SubItems(GuiRptParamsListView.ID - 1)
                .IDAssignments = itmX.SubItems(GuiRptParamsListView.IDAssignments - 1)
                .Number = itmX.SubItems(GuiRptParamsListView.Number - 1)
                .ProjectName = itmX.SubItems(GuiRptParamsListView.ProjectName - 1)
                .ClassName = itmX.SubItems(GuiRptParamsListView.ClassName - 1)
                .ParamName = itmX.SubItems(GuiRptParamsListView.ParamName - 1)
                .ParamCaption = itmX.Text
                .ParamValue = sText
                itmX.SubItems(GuiRptParamsListView.ParamValue - 1) = sText
                .ParamDataType = itmX.SubItems(GuiRptParamsListView.ParamDataType - 1)
                .SortMe = itmX.SubItems(GuiRptParamsListView.SortMe - 1)
                .IsDeleted = goUtil.GetFlagFromText(itmX.SubItems(GuiRptParamsListView.IsDeleted - 1))
                .DownLoadMe = goUtil.GetFlagFromText(itmX.SubItems(GuiRptParamsListView.DownLoadMe - 1))
                .UpLoadMe = "True"
                sFlagText = goUtil.GetFlagText(True)
                itmX.SubItems(GuiRptParamsListView.UpLoadMe - 1) = sFlagText
                itmX.ListSubItems(GuiRptParamsListView.UpLoadMe - 1).ReportIcon = GuiRptParamsStatusList.UpLoadMe
                .AdminComments = itmX.SubItems(GuiRptParamsListView.AdminComments - 1)
                .DateLastUpdated = Format(Now(), "MM/DD/YYYY HH:MM:SS")
                itmX.SubItems(GuiRptParamsListView.DateLastUpdated - 1) = .DateLastUpdated
                itmX.SubItems(GuiRptParamsListView.DateLastUpdatedSort - 1) = Format(.DateLastUpdated, "YYYY/MM/DD HH:MM:SS")
                .UpdateByUserID = goUtil.gsCurUsersID
                itmX.SubItems(GuiRptParamsListView.UpdateByUserID - 1) = .UpdateByUserID
            End With
            mfrmClaim.EditRptParamItem MyParam
        End If
NEXT_ITEM:
    Next
    
    cmdSpelling.Enabled = True
    Screen.MousePointer = MousePointerConstants.vbDefault
    
    'cleanup
    Set itmX = Nothing
    txtSpellMe.Text = vbNullString
    Exit Sub
EH:
    Screen.MousePointer = MousePointerConstants.vbDefault
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdSpelling_Click"
End Sub

Private Sub cmdUpdateEdit_Click()
    On Error GoTo EH
    
    cmdUpdateEdit.Enabled = False
    
    If UpdateEdit Then
        cmdUpdateEdit.Enabled = True
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdUpdateEdit_Click"
End Sub

Private Sub cmdUpdateMulti_Click(Index As Integer)
    On Error GoTo EH
    Dim itmX As ListItem
    Dim bMultiUpdateChecks As Boolean
    Dim bMultiUpdateDates As Boolean
    Dim iDataType As VBA.VbVarType
    Dim MyParam As V2ECKeyBoard.MiscReportParam
    'Checks Update
    Dim bFlag As Boolean
    Dim bUpdateFlag As Boolean
    Dim sFlagText As String
    Dim iMyIcon As Long
    'Dates Update
    Dim sDate As String
    Dim sUpdateDate As String
    
    Screen.MousePointer = MousePointerConstants.vbHourglass
    cmdUpdateMulti(Index).Enabled = False
    
    Select Case Index
        Case 0
            bMultiUpdateChecks = True
            If chkMultiUpdate.Value = vbChecked Then
                bUpdateFlag = True
            Else
                bUpdateFlag = False
            End If
        Case 1
            bMultiUpdateDates = True
            If IsDate(txtDateMultiUpdate.Text) Then
                If CDate(txtDateMultiUpdate.Text) = NULL_DATE Then
                    sUpdateDate = vbNullString
                Else
                    sUpdateDate = Format(txtDateMultiUpdate.Text, "MM/DD/YYYY")
                End If
            Else
                sUpdateDate = vbNullString
            End If
    End Select

    For Each itmX In lvwRptParams.ListItems
        If Not itmX.Selected Then
            GoTo NEXT_ITEM
        End If
        'Set the Param Value Here according to the Data Type
        iDataType = itmX.SubItems(GuiRptParamsListView.ParamDataType - 1)
        If bMultiUpdateChecks Then
            If iDataType <> vbBoolean Then
                GoTo NEXT_ITEM
            Else
                bFlag = goUtil.GetFlagFromText(itmX.SubItems(GuiRptParamsListView.ParamValue - 1))
                'Check to see if the current value already matches the updatevalue.
                'if it does then can skip this item
                If bFlag = bUpdateFlag Then
                    GoTo NEXT_ITEM
                End If
            End If
        End If
        If bMultiUpdateDates Then
            If iDataType <> vbDate Then
                GoTo NEXT_ITEM
            Else
                sDate = goUtil.GetFlagFromText(itmX.SubItems(GuiRptParamsListView.ParamValue - 1))
                'Check to see if the current value already matches the updatevalue.
                'if it does then can skip this item
                If IsDate(sDate) And IsDate(sUpdateDate) Then
                    If CDate(sDate) = CDate(sUpdateDate) Then
                        GoTo NEXT_ITEM
                    End If
                End If
                'Check for nulstring matched
                If Trim(sDate) = vbNullString And Trim(sUpdateDate) = vbNullString Then
                     GoTo NEXT_ITEM
                End If
            End If
        End If
        With MyParam
            .MiscReportParamID = itmX.SubItems(GuiRptParamsListView.MiscReportParamID - 1)
            .AssignmentsID = itmX.SubItems(GuiRptParamsListView.AssignmentsID - 1)
            .ID = itmX.SubItems(GuiRptParamsListView.ID - 1)
            .IDAssignments = itmX.SubItems(GuiRptParamsListView.IDAssignments - 1)
            .Number = itmX.SubItems(GuiRptParamsListView.Number - 1)
            .ProjectName = itmX.SubItems(GuiRptParamsListView.ProjectName - 1)
            .ClassName = itmX.SubItems(GuiRptParamsListView.ClassName - 1)
            .ParamName = itmX.SubItems(GuiRptParamsListView.ParamName - 1)
            .ParamCaption = itmX.Text
            'Update .ParamValue Below
            .ParamDataType = itmX.SubItems(GuiRptParamsListView.ParamDataType - 1)
            .SortMe = itmX.SubItems(GuiRptParamsListView.SortMe - 1)
            .IsDeleted = goUtil.GetFlagFromText(itmX.SubItems(GuiRptParamsListView.IsDeleted - 1))
            .DownLoadMe = goUtil.GetFlagFromText(itmX.SubItems(GuiRptParamsListView.DownLoadMe - 1))
            .UpLoadMe = "True"
            sFlagText = goUtil.GetFlagText(True)
            itmX.SubItems(GuiRptParamsListView.UpLoadMe - 1) = sFlagText
            itmX.ListSubItems(GuiRptParamsListView.UpLoadMe - 1).ReportIcon = GuiRptParamsStatusList.UpLoadMe
            .AdminComments = itmX.SubItems(GuiRptParamsListView.AdminComments - 1)
            .DateLastUpdated = Format(Now(), "MM/DD/YYYY HH:MM:SS")
            itmX.SubItems(GuiRptParamsListView.DateLastUpdated - 1) = .DateLastUpdated
            itmX.SubItems(GuiRptParamsListView.DateLastUpdatedSort - 1) = Format(.DateLastUpdated, "YYYY/MM/DD HH:MM:SS")
            .UpdateByUserID = goUtil.gsCurUsersID
            itmX.SubItems(GuiRptParamsListView.UpdateByUserID - 1) = .UpdateByUserID
        End With
        
        If bMultiUpdateChecks Then
            MyParam.ParamValue = bUpdateFlag
            sFlagText = goUtil.GetFlagText(bUpdateFlag)
            If bUpdateFlag Then
                iMyIcon = GuiRptParamsStatusList.ValueIsChecked
                MyParam.ParamValue = "-1"
            Else
                iMyIcon = GuiRptParamsStatusList.ValueIsUnchecked
                MyParam.ParamValue = "0"
            End If
            itmX.SubItems(GuiRptParamsListView.ParamValue - 1) = sFlagText
            itmX.ListSubItems(GuiRptParamsListView.ParamValue - 1).ReportIcon = iMyIcon
        ElseIf bMultiUpdateDates Then
            itmX.SubItems(GuiRptParamsListView.ParamValue - 1) = sUpdateDate
        End If
        'now Edit this param
        mfrmClaim.EditRptParamItem MyParam
NEXT_ITEM:
    Next
    
    cmdUpdateMulti(Index).Enabled = True
    Screen.MousePointer = MousePointerConstants.vbDefault
    Exit Sub
EH:
    Screen.MousePointer = MousePointerConstants.vbDefault
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdUpdateMulti_Click"
End Sub

Private Sub Form_Load()
    On Error GoTo EH
    Dim adoRS As ADODB.Recordset
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    Dim lPos As Long
    Dim lId As Long
    Dim lMyPayReqID As Long
    Dim adoClassTypeIDInRS As ADODB.Recordset
    Dim sLossFormat As String
    Dim bIncludeDeletedItems As Boolean
    
    'Check the Loss Format
    sLossFormat = goUtil.IsNullIsVbNullString(mfrmClaim.adoRSAssignments.Fields("LRFormat"))
    If StrComp(sLossFormat, "V2ECcarFarmers.clsLossXML01", vbTextCompare) = 0 Then
        bIncludeDeletedItems = True
    End If
    
    mbLoading = True
    
    'Set the Form Icon
    Me.Icon = mfrmClaim.optClaim(GuiClaimOptions.opt03_Indemnity).Picture
'    goUtil.utFormWinRegPos goUtil.gsMainAppEXEName, Me, , , , , True
    
    LoadHeaderlvwRptParams
    LoadReports
    
    
    'Select the Correct Payment Request
    lMyPayReqID = msPayReqID
    For lPos = 0 To cboPayments.ListCount - 1
       lId = cboPayments.ItemData(lPos)
       If lId = lMyPayReqID Then
           cboPayments.ListIndex = lPos
           Exit For
       End If
    Next
    
    sSQL = "SELECT "
    sSQL = sSQL & "[RTChecksID], "
    sSQL = sSQL & "[AssignmentsID], "
    sSQL = sSQL & "[BillingCountID], "
    sSQL = sSQL & "[ID], "
    sSQL = sSQL & "[IDAssignments], "
    sSQL = sSQL & "[IDBillingCount], "
    sSQL = sSQL & "[CheckNum], "
    sSQL = sSQL & "[RT42_ClassOfLossID], "
    sSQL = sSQL & "[RT43_TypeOfLossID], "
    sSQL = sSQL & "[RT50_sInsuredPayeeName], "
    sSQL = sSQL & "[RT51_sPayeeNames], "
    sSQL = sSQL & "[RT52_sAddress], "
    sSQL = sSQL & "[RT53_cAmountOfCheck], "
    sSQL = sSQL & "[AppliedDeductible], "
    sSQL = sSQL & "[RT54_CompanyCatSpecID], "
    sSQL = sSQL & "[tempCHeckName], "
    sSQL = sSQL & "[PrintOnIB], "
    sSQL = sSQL & "[IsDeleted], "
    sSQL = sSQL & "[DownLoadMe], "
    sSQL = sSQL & "[UpLoadMe], "
    sSQL = sSQL & "[AdminComments], "
    sSQL = sSQL & "[DateLastUpdated], "
    sSQL = sSQL & "[UpdateByUserID]"
    sSQL = sSQL & "FROM RTChecks "
    sSQL = sSQL & "WHERE [RTChecksID] = " & msPayReqID & " "
    
    Set oConn = New ADODB.Connection
    Set adoRS = New ADODB.Recordset
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    adoRS.CursorLocation = adUseClient
    adoRS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set adoRS.ActiveConnection = Nothing
    
    If adoRS.Fields("PrintOnIB") Then
        chkPrintOnIB.Enabled = True
    ElseIf goUtil.IsNullIsVbNullString(adoRS.Fields("CheckNum")) = 1 Then
        chkPrintOnIB.Enabled = True
    End If
    
    'Populate the Available Cat Code
    If Not MyGUI.adoRSCatCode Is Nothing Then
        cboCatCode.Clear
        mfrmClaim.PopulateLookUp MyGUI.adoRSCatCode, _
                        adoRS, _
                        cboCatCode, _
                        "ClientCompanyCatSpecID", _
                        "RT54_CompanyCatSpecID", _
                        "CatCode", _
                        "Comments"
    End If
    
    'Only Include items where the ClassTypeID are included under
    'Policy Limits Table
    sSQL = "SELECT  [ClassTypeID] "
    sSQL = sSQL & "FROM PolicyLimits "
    sSQL = sSQL & "WHERE [AssignmentsID] = " & msAssignmentsID & " "
    If Not bIncludeDeletedItems Then
        sSQL = sSQL & "AND [IsDeleted] = 0 "
    End If
    
    Set adoClassTypeIDInRS = New ADODB.Recordset
    adoClassTypeIDInRS.CursorLocation = adUseClient
    adoClassTypeIDInRS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set adoRS.ActiveConnection = Nothing
    
    
    'Populate the Available Class Of Loss
    If Not MyGUI.adoClassType Is Nothing Then
        cboClassOfLoss.Clear
        mfrmClaim.PopulateLookUp MyGUI.adoClassType, _
                        adoRS, _
                        cboClassOfLoss, _
                        "ClassTypeID", _
                        "RT42_ClassOfLossID", _
                        "Class", _
                        "Description", _
                        , _
                        , _
                        , _
                        , _
                        adoClassTypeIDInRS
    End If
    
    'record the Orginal Index
    mlCOLOrigListIndex = cboClassOfLoss.ListIndex
    
    'Populate the Available Type Of Loss Info
    If Not MyGUI.adoRSTypeOfLoss Is Nothing Then
        cboTypeOfLoss.Clear
        mfrmClaim.PopulateLookUp MyGUI.adoRSTypeOfLoss, _
                        adoRS, _
                        cboTypeOfLoss, _
                        "TypeOfLossID", _
                        "RT43_TypeOfLossID", _
                        "TypeOfLoss", _
                        "Code"
    End If
    
    txtInsured.Text = goUtil.IsNullIsVbNullString(adoRS.Fields("RT50_sInsuredPayeeName"))
    txtMortgageeName.Text = goUtil.IsNullIsVbNullString(adoRS.Fields("RT51_sPayeeNames"))
    txtAppliedDeductible.Text = Format(goUtil.IsNullIsVbNullString(adoRS.Fields("AppliedDeductible")), "#,###,###,##0.00")
    txtAmountOfCheck.Text = Format(goUtil.IsNullIsVbNullString(adoRS.Fields("RT53_cAmountOfCheck")), "#,###,###,##0.00")
    txtAddress.Text = Replace(goUtil.IsNullIsVbNullString(adoRS.Fields("RT52_sAddress")), F_VBCRLF, "    ")
    
    If CBool(adoRS.Fields("PrintOnIB").Value) Then
        chkPrintOnIB.Value = vbChecked
    Else
        chkPrintOnIB.Value = vbUnchecked
    End If
    
    If mfrmClaim.GetRptParamColAndLoadLvw(cboPayments, lvwRptParams, framMultiUpdate) Then
        'Edit Payreq will make sure there is nothing selected
        Set lvwRptParams.SelectedItem = Nothing
        'Enable the Print button
        cmdPrintReport.Enabled = True
    Else
        cmdPrintReport.Enabled = False
    End If
    
    If mbAssociateFromIndem Or mbEditGetAmount Then
        Timer_GetAmount.Enabled = True
    End If
    
    'cleanup
    
    Set adoRS = Nothing
    Set adoClassTypeIDInRS = Nothing
    Set oConn = Nothing
    
    mbLoading = False
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_Load"
End Sub

Public Function LoadReports() As Boolean
    On Error GoTo EH
    Dim lNewIndex As Long
    Dim sData As String
    Dim RS As ADODB.Recordset
    Dim RSPayment As ADODB.Recordset
    
    'Load Recordsets
    mfrmClaim.SetadoRSPayment msAssignmentsID ' Actual payments
    mfrmClaim.SetadoRSPaymentReports 'Software for Payments
    mfrmClaim.SetadoRSPaymentReportsHistory 'Software History for Payments
    
    'Load Payment Report
    Set RS = mfrmClaim.adoRSPaymentReports ' Software
    Set RSPayment = mfrmClaim.adoRSPayment ' Actual Payment Data
    cboPayments.Clear
    If RS.RecordCount = 1 And RSPayment.RecordCount > 0 Then
        RS.MoveFirst
        Do Until RSPayment.EOF
            sData = vbNullString
            If CBool(RSPayment.Fields("PrintOnIB").Value) Then
                sData = RS.Fields("Description").Value & " (" & RSPayment.Fields("CheckNum").Value & ") - Print On IB"
            Else
                sData = RS.Fields("Description").Value & " (" & RSPayment.Fields("CheckNum").Value & ")"
            End If
            sData = sData & String(200, " ")
            sData = sData & RS.Fields("ProjectName").Value & "|" & RS.Fields("ClassName").Value & "|" & RS.Fields("Version").Value
            cboPayments.AddItem sData
            lNewIndex = cboPayments.NewIndex
            'Use the Actual Payment Data Unique ID
            cboPayments.ItemData(lNewIndex) = RSPayment.Fields("RTChecksID").Value
            RSPayment.MoveNext
        Loop
    End If
    
    'Need to Add Items from History Table
    Set RS = mfrmClaim.adoRSPaymentReportsHistory
    
    If RS.RecordCount > 0 And RSPayment.RecordCount > 0 Then
        RS.MoveFirst
        Do Until RSPayment.EOF
            sData = vbNullString
            sData = RS.Fields("Description").Value & " (" & RSPayment.Fields("CheckNum").Value & ")"
            sData = sData & String(200, " ")
            sData = sData & RS.Fields("ProjectName").Value & "|" & RS.Fields("ClassName").Value & "|" & RS.Fields("Version").Value
            cboPayments.AddItem sData
            lNewIndex = cboPayments.NewIndex
            'Use the Actual IB Data Unique ID
            cboPayments.ItemData(lNewIndex) = RSPayment.Fields("RTChecksID").Value
            RSPayment.MoveNext
        Loop
    End If
    
    'cleanup
    Set RS = Nothing
    Set RSPayment = Nothing
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function LoadReports"
End Function

Public Sub LoadHeaderlvwRptParams()
    On Error GoTo EH
    Dim bGridOn As Boolean
    
    'set the columnheaders
    With lvwRptParams
        .ColumnHeaders.Add , "ParamCaption", "Caption"
        .ColumnHeaders.Add , "ParamValue", "Value"
        .ColumnHeaders.Add , "IsDeleted", "Is Deleted"
        .ColumnHeaders.Add , "UpLoadMe", "UpLoad Me"
        .ColumnHeaders.Add , "DateLastUpdated", "Date Last Updated"
        .ColumnHeaders.Add , "DateLastUpdatedSort", "Sort Date Last Updated" ' hidden
        .ColumnHeaders.Add , "AdminComments", "Admin Comments"
        .ColumnHeaders.Add , "Number", "Number" ' hidden
        .ColumnHeaders.Add , "ProjectName", "ProjectName" ' hidden
        .ColumnHeaders.Add , "ClassName", "ClassName" ' hidden
        .ColumnHeaders.Add , "ParamName", "ParamName" ' hidden
        .ColumnHeaders.Add , "ParamDataType", "ParamDataType" ' hidden
        .ColumnHeaders.Add , "SortMe", "SortMe" ' hidden
        .ColumnHeaders.Add , "SortMeSort", "Sort SortMe" ' hidden
        .ColumnHeaders.Add , "ID", "ID" 'Hidden
        .ColumnHeaders.Add , "IDAssignments", "IDAssignments" 'Hidden
        .ColumnHeaders.Add , "MiscReportParamID", "MiscReportParamID" ' Hidden
        .ColumnHeaders.Add , "AssignmentsID", "AssignmentsID"  ' Hidden
        .ColumnHeaders.Add , "DownLoadMe", "DownLoadMe"  ' hidden
        .ColumnHeaders.Add , "UpdateByUserID", "UpdateByUserID"  ' Hidden
        
        .Sorted = False
        .SortOrder = lvwAscending
        
        'ParamCaption
        .ColumnHeaders.Item(GuiRptParamsListView.ParamCaption).Width = 5000
        .ColumnHeaders.Item(GuiRptParamsListView.ParamCaption).Alignment = lvwColumnLeft
        'ParamValue
        .ColumnHeaders.Item(GuiRptParamsListView.ParamValue).Width = 2500
        .ColumnHeaders.Item(GuiRptParamsListView.ParamValue).Alignment = lvwColumnLeft
        'Is Deleted
        .ColumnHeaders.Item(GuiRptParamsListView.IsDeleted).Width = 400
        .ColumnHeaders.Item(GuiRptParamsListView.IsDeleted).Alignment = lvwColumnCenter
        .ColumnHeaders.Item(GuiRptParamsListView.IsDeleted).Icon = GuiRptParamsStatusList.IsDeleted
        'UpLoad Me
        .ColumnHeaders.Item(GuiRptParamsListView.UpLoadMe).Width = 400
        .ColumnHeaders.Item(GuiRptParamsListView.UpLoadMe).Alignment = lvwColumnCenter
        .ColumnHeaders.Item(GuiRptParamsListView.UpLoadMe).Icon = GuiRptParamsStatusList.UpLoadMe
        'DateLastUpdated
        .ColumnHeaders.Item(GuiRptParamsListView.DateLastUpdated).Width = 2200
        .ColumnHeaders.Item(GuiRptParamsListView.DateLastUpdated).Alignment = lvwColumnLeft
        'DateLastUpdatedSort
        .ColumnHeaders.Item(GuiRptParamsListView.DateLastUpdatedSort).Width = 0  'Hidden Sort Column
        .ColumnHeaders.Item(GuiRptParamsListView.DateLastUpdatedSort).Alignment = lvwColumnLeft
        'AdminComments
        .ColumnHeaders.Item(GuiRptParamsListView.AdminComments).Width = 10000
        .ColumnHeaders.Item(GuiRptParamsListView.AdminComments).Alignment = lvwColumnLeft
        'Number
        .ColumnHeaders.Item(GuiRptParamsListView.Number).Width = 0  'hidden
        .ColumnHeaders.Item(GuiRptParamsListView.Number).Alignment = lvwColumnLeft
        'ProjectName
        .ColumnHeaders.Item(GuiRptParamsListView.ProjectName).Width = 0  'hidden
        .ColumnHeaders.Item(GuiRptParamsListView.ProjectName).Alignment = lvwColumnLeft
        'ClassName
        .ColumnHeaders.Item(GuiRptParamsListView.ClassName).Width = 0  'hidden
        .ColumnHeaders.Item(GuiRptParamsListView.ClassName).Alignment = lvwColumnLeft
        'ParamName
        .ColumnHeaders.Item(GuiRptParamsListView.ParamName).Width = 0  'hidden
        .ColumnHeaders.Item(GuiRptParamsListView.ParamName).Alignment = lvwColumnLeft
        'ParamDataType
        .ColumnHeaders.Item(GuiRptParamsListView.ParamDataType).Width = 0  'hidden
        .ColumnHeaders.Item(GuiRptParamsListView.ParamDataType).Alignment = lvwColumnLeft
        'SortMe
        .ColumnHeaders.Item(GuiRptParamsListView.SortMe).Width = 0  'hidden
        .ColumnHeaders.Item(GuiRptParamsListView.SortMe).Alignment = lvwColumnLeft
        'SortMeSort
        .ColumnHeaders.Item(GuiRptParamsListView.SortMeSort).Width = 0  'hidden
        .ColumnHeaders.Item(GuiRptParamsListView.SortMeSort).Alignment = lvwColumnLeft
        'ID
        .ColumnHeaders.Item(GuiRptParamsListView.ID).Width = 0  'hidden
        .ColumnHeaders.Item(GuiRptParamsListView.ID).Alignment = lvwColumnLeft
        'IDAssignments
        .ColumnHeaders.Item(GuiRptParamsListView.IDAssignments).Width = 0  'Hidden
        .ColumnHeaders.Item(GuiRptParamsListView.IDAssignments).Alignment = lvwColumnLeft
        'MiscReportParamID
        .ColumnHeaders.Item(GuiRptParamsListView.MiscReportParamID).Width = 0   'Hidden
        .ColumnHeaders.Item(GuiRptParamsListView.MiscReportParamID).Alignment = lvwColumnLeft
        'AssignmentsID
        .ColumnHeaders.Item(GuiRptParamsListView.AssignmentsID).Width = 0  'Hidden
        .ColumnHeaders.Item(GuiRptParamsListView.AssignmentsID).Alignment = lvwColumnLeft
        'DownLoadMe
        .ColumnHeaders.Item(GuiRptParamsListView.DownLoadMe).Width = 0  'Hidden
        .ColumnHeaders.Item(GuiRptParamsListView.DownLoadMe).Alignment = lvwColumnLeft
        'UpdateByUserID
        .ColumnHeaders.Item(GuiRptParamsListView.UpdateByUserID).Width = 0  'Hidden
        .ColumnHeaders.Item(GuiRptParamsListView.UpdateByUserID).Alignment = lvwColumnLeft
    End With
    
    bGridOn = CBool(GetSetting(App.EXEName, "GENERAL", "GRID_ON", False))
    lvwRptParams.GridLines = bGridOn
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub LoadHeaderlvwRptParams"
End Sub

Public Function CLEANUP() As Boolean
    On Error GoTo EH

    Set mfrmClaim = Nothing
    Set mfrmIndemnity = Nothing
    Set moGUI = Nothing
    Set moActivecboReport = Nothing
    
    CLEANUP = True
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function CLEANUP"
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo EH
    
     Select Case UnloadMode
        Case vbFormControlMenu
            If cmdSave.Enabled Then
'                If MsgBox("Do you want to save changes?", vbYesNo + vbQuestion, "Save Changes") = vbYes Then
                    If Not SaveMe() Then
                        Cancel = True
                        Exit Sub
                    End If
'                End If
            ElseIf mlCOLOrigListIndex = -1 Then
                cmdSave.Enabled = True
            End If
'            goUtil.utFormWinRegPos goUtil.gsMainAppEXEName, Me, True, , , , True
            mbUnloadMe = True
            Me.Visible = False
            Cancel = True
    End Select
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_QueryUnload"
End Sub


Private Sub Form_Resize()
    If Not mbLoading Then
'        ReSizeMe
    End If
End Sub

Public Sub ReSizeMe()
    On Error Resume Next
    
    'RePos Controls
    'Width and Lefts
    framEdit.Width = Me.Width - 375
    cmdPrintReport.left = Me.Width - 1470
    lvwRptParams.Width = Me.Width - 6135
    framEditParam.Width = Me.Width - 6135
    txtParamCaption.Width = Me.Width - 6375
    chkParamBoolean.Width = Me.Width - 7575
    txtParamValue.Width = Me.Width - 7335
    cmdParamDate.left = Me.Width - 7590
    cmdCancelEdit.left = Me.Width - 7110
    cmdUpdateEdit.left = Me.Width - 7110
    framMultiUpdate.Width = Me.Width - 6135
    txtAddress.Width = Me.Width - 3975
    chkAddress.left = Me.Width - 4110
    
    'framCommands
    framCommands.left = Me.Width - 3630
    
    
    'Heights and Tops
    framEdit.Height = Me.Height - 1740
    framEditParam.Height = Me.Height - 3180
    lvwRptParams.Height = Me.Height - 3180
    txtParamCaption.Height = Me.Height - 4500
    chkParamBoolean.top = Me.Height - 4035
    txtParamValue.top = Me.Height - 3675
    cmdParamDate.top = Me.Height - 3675
    cmdCancelEdit.top = Me.Height - 4095
    cmdUpdateEdit.top = Me.Height - 3675
    framMultiUpdate.top = Me.Height - 2595
    lblAddress.top = Me.Height - 1635
    txtAddress.top = Me.Height - 1328
    chkAddress.top = Me.Height - 1335
    'framCommands
    framCommands.top = Me.Height - 1755
    
End Sub

Public Function SaveMe() As Boolean
    On Error GoTo EH
    Dim udtPayReq As GuiPayReqsItem
    Dim oListView As ListView
    Dim itmX As ListItem
    Dim sClassOfLossID As String
    Dim sClassOfLoss As String
    Dim sClassOfLossCode  As String
    Dim sTypeOfLossID As String
    Dim sTypeOfLoss As String
    Dim sTypeOfLossCode As String
    Dim sCompanyCatSpecID As String
    Dim sCatCode As String
    Dim RS As ADODB.Recordset
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    Dim bCancel As Boolean
    Dim sMess As String
    Dim sFlagText As String
    
    ' Validate some stuff first
    goUtil.utValidate Me
    'Make sure the required stuff was filled out
    If cboCatCode.ListIndex = -1 Then
        bCancel = True
        sMess = sMess & "Must Select CatCode!" & vbCrLf
    ElseIf cboClassOfLoss.ListIndex = -1 Then
        bCancel = True
        sMess = sMess & "Must Select Class!" & vbCrLf
    ElseIf cboTypeOfLoss.ListIndex = -1 Then
        bCancel = True
        sMess = sMess & "Must Select Type Of Loss!" & vbCrLf
    End If
    
    If bCancel Then
        MsgBox sMess, vbOKOnly + vbExclamation, "Could Not Save!"
        Exit Function
    End If
    
    cmdSave.Enabled = False
    
    Set oListView = mfrmIndemnity.lvwPayReqs
    
    'Need to Set some Vars first
    
    Set oConn = New ADODB.Connection
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    'Class of Loss
    sSQL = "SELECT COL.[ClassOfLossID], "
    sSQL = sSQL & "(CT.Description + ' (' + COL.Description + ')') As ClassOfLoss, "
    sSQL = sSQL & "COL.Code As ClassOfLossCode "
    sSQL = sSQL & "FROM ClassOfLoss COL "
    sSQL = sSQL & "INNER JOIN ClassType CT ON CT.ClassTypeID = COL.ClassTypeID "
    sSQL = sSQL & "WHERE COL.[ClientCompanyID] = " & goUtil.gsCurCar & " "
    sSQL = sSQL & "AND COL.[ClassTypeID] = " & cboClassOfLoss.ItemData(cboClassOfLoss.ListIndex) & " "
    
    Set RS = New ADODB.Recordset
    RS.CursorLocation = adUseClient
    RS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set RS.ActiveConnection = Nothing
    
    If RS.RecordCount = 1 Then
        RS.MoveFirst
        sClassOfLossID = goUtil.IsNullIsVbNullString(RS.Fields("ClassOfLossID"))
        sClassOfLoss = goUtil.IsNullIsVbNullString(RS.Fields("ClassOfLoss"))
        sClassOfLossCode = goUtil.IsNullIsVbNullString(RS.Fields("ClassOfLossCode"))
    End If
    
    'type Of Loss
    Set RS = Nothing
    sSQL = "SELECT TOL.[TypeOfLossID], "
    sSQL = sSQL & "TOL.TypeOfLoss + ' (' + TOL.Description + ')' As TypeOfLoss, "
    sSQL = sSQL & "TOL.Code As TypeOfLossCode "
    sSQL = sSQL & "FROM TypeOfLoss TOL "
    sSQL = sSQL & "WHERE TOL.[ClientCompanyID] = " & goUtil.gsCurCar & " "
    sSQL = sSQL & "AND TOL.[TypeOfLossID] = " & cboTypeOfLoss.ItemData(cboTypeOfLoss.ListIndex) & " "
    
    Set RS = New ADODB.Recordset
    RS.CursorLocation = adUseClient
    RS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set RS.ActiveConnection = Nothing
    
    If RS.RecordCount = 1 Then
        RS.MoveFirst
        sTypeOfLossID = goUtil.IsNullIsVbNullString(RS.Fields("TypeOfLossID"))
        sTypeOfLoss = goUtil.IsNullIsVbNullString(RS.Fields("TypeOfLoss"))
        sTypeOfLossCode = goUtil.IsNullIsVbNullString(RS.Fields("TypeOfLossCode"))
    End If
    
    'Cat Code
    Set RS = Nothing
    sSQL = "SELECT CCCS.[ClientCompanyCatSpecID], "
    sSQL = sSQL & "CCCS.[CatCode] "
    sSQL = sSQL & "FROM ClientCompanyCatSpec CCCS "
    sSQL = sSQL & "WHERE CCCS.ClientCompanyCatSpecID = " & cboCatCode.ItemData(cboCatCode.ListIndex) & " "
    
    Set RS = New ADODB.Recordset
    RS.CursorLocation = adUseClient
    RS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set RS.ActiveConnection = Nothing
    
    If RS.RecordCount = 1 Then
        RS.MoveFirst
        sCompanyCatSpecID = goUtil.IsNullIsVbNullString(RS.Fields("ClientCompanyCatSpecID"))
        sCatCode = goUtil.IsNullIsVbNullString(RS.Fields("CatCode"))
    End If
    
    Set oConn = Nothing
    Set RS = Nothing

    With udtPayReq
        Set itmX = oListView.ListItems("""" & msPayReqID & """")
        .RTChecksID = itmX.SubItems(GuiPayReqsListView.RTChecksID - 1)
        .AssignmentsID = itmX.SubItems(GuiPayReqsListView.AssignmentsID - 1)
        .BillingCountID = itmX.SubItems(GuiPayReqsListView.BillingCountID - 1)
        .ID = itmX.SubItems(GuiPayReqsListView.ID - 1)
        .IDAssignments = itmX.SubItems(GuiPayReqsListView.IDAssignments - 1)
        .IDBillingCount = itmX.SubItems(GuiPayReqsListView.IDBillingCount - 1)
        .CheckNum = itmX.Text
        .RT42_ClassOfLossID = sClassOfLossID
        itmX.SubItems(GuiPayReqsListView.RT42_ClassOfLossID - 1) = sClassOfLossID
        itmX.SubItems(GuiPayReqsListView.ClassOfLoss - 1) = sClassOfLoss
        itmX.SubItems(GuiPayReqsListView.ClassOfLossCode - 1) = sClassOfLossCode
        .RT43_TypeOfLossID = sTypeOfLossID
        itmX.SubItems(GuiPayReqsListView.RT43_TypeOfLossID - 1) = sTypeOfLossID
        itmX.SubItems(GuiPayReqsListView.TypeOfLoss - 1) = sTypeOfLoss
        itmX.SubItems(GuiPayReqsListView.TypeOfLossCode - 1) = sTypeOfLossCode
        .RT50_sInsuredPayeeName = txtInsured.Text
        itmX.SubItems(GuiPayReqsListView.RT50_sInsuredPayeeName - 1) = txtInsured.Text
        .RT51_sPayeeNames = txtMortgageeName.Text
        itmX.SubItems(GuiPayReqsListView.RT51_sPayeeNames - 1) = txtMortgageeName.Text
        .RT52_sAddress = txtAddress.Text
        itmX.SubItems(GuiPayReqsListView.RT52_sAddress - 1) = txtAddress.Text
        itmX.SubItems(GuiPayReqsListView.RT52_sAddressSort - 1) = goUtil.utNumInTextSortFormat(txtAddress.Text)
        .RT53_cAmountOfCheck = txtAmountOfCheck.Text
        itmX.SubItems(GuiPayReqsListView.RT53_cAmountOfCheck - 1) = txtAmountOfCheck.Text
        itmX.SubItems(GuiPayReqsListView.RT53_cAmountOfCheckSort - 1) = goUtil.utNumInTextSortFormat(txtAmountOfCheck.Text)
        .AppliedDeductible = txtAppliedDeductible.Text
        itmX.SubItems(GuiPayReqsListView.AppliedDeductible - 1) = txtAppliedDeductible.Text
        itmX.SubItems(GuiPayReqsListView.AppliedDeductibleSort - 1) = goUtil.utNumInTextSortFormat(txtAppliedDeductible.Text)
        .RT54_CompanyCatSpecID = sCompanyCatSpecID
        itmX.SubItems(GuiPayReqsListView.RT54_CompanyCatSpecID - 1) = sCompanyCatSpecID
        itmX.SubItems(GuiPayReqsListView.CatCode - 1) = sCatCode
        .tempCHeckName = itmX.SubItems(GuiPayReqsListView.tempCHeckName - 1)
        
        'PrintOnIB
        If chkPrintOnIB.Value = vbChecked Then
            itmX.ListSubItems(GuiPayReqsListView.PrintOnIB - 1).ReportIcon = GuiPayReqsStatusList.IsChecked
            .PrintOnIB = True
        Else
            itmX.ListSubItems(GuiPayReqsListView.PrintOnIB - 1).ReportIcon = Empty
            .PrintOnIB = False
        End If
        sFlagText = goUtil.GetFlagText(CBool(.PrintOnIB))
        itmX.SubItems(GuiPayReqsListView.PrintOnIB - 1) = sFlagText
        
        .IsDeleted = goUtil.GetFlagFromText(itmX.SubItems(GuiPayReqsListView.IsDeleted - 1))
        .DownLoadMe = goUtil.GetFlagFromText(itmX.SubItems(GuiPayReqsListView.DownLoadMe - 1))
        itmX.ListSubItems(GuiPayReqsListView.UpLoadMe - 1).ReportIcon = GuiPayReqsStatusList.UpLoadMe
        .UpLoadMe = True
        .AdminComments = itmX.SubItems(GuiPayReqsListView.AdminComments - 1)
        .DateLastUpdated = Format(Now(), "MM/DD/YYYY HH:MM:SS")
        itmX.SubItems(GuiPayReqsListView.DateLastUpdated - 1) = .DateLastUpdated
        itmX.SubItems(GuiPayReqsListView.DateLastUpdatedSort - 1) = Format(.DateLastUpdated, "YYYY/MM/DD HH:MM:SS")
        itmX.SubItems(GuiPayReqsListView.UpdateByUserID - 1) = goUtil.gsCurUsersID
        .UpdateByUserID = itmX.SubItems(GuiPayReqsListView.UpdateByUserID - 1)
    End With

    'Edit this entry
    mfrmIndemnity.EditPayReqItem udtPayReq
    oListView.SortKey = GuiPayReqsListView.CheckNum
    oListView.Sorted = True

    'now be sure its visible
    Set itmX = oListView.SelectedItem
    itmX.EnsureVisible

    SaveMe = True
    
    'cleanup
    Set oListView = Nothing
    Set itmX = Nothing

    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function SaveMe"
End Function

Private Sub lstvbUserDefinedType_DblClick()
    On Error GoTo EH
    
    If lstvbUserDefinedType.ListIndex > -1 Then
        txtParamValue.Text = lstvbUserDefinedType.Text
        UpdateEdit
        lstvbUserDefinedType.Visible = False
        cmdSave.Enabled = mbSaveEnabled
        cmdSave.Default = True
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lstvbUserDefinedType_DblClick"
End Sub

Private Sub lstvbUserDefinedType_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case KeyCodeConstants.vbKeyReturn
            If lstvbUserDefinedType.Enabled And lstvbUserDefinedType.Visible Then
                lstvbUserDefinedType_DblClick
            End If
    End Select
End Sub

Private Sub lvwRptParams_DblClick()
    On Error GoTo EH
    
    If Not lvwRptParams.SelectedItem Is Nothing Then
        lvwRptParams.Visible = False
        ShowEditRptParam
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lvwRptParams_DblClick"
End Sub

Private Sub lvwRptParams_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo EH
    Select Case KeyCode
        Case KeyCodeConstants.vbKeyReturn, KeyCodeConstants.vbKeySpace
            If Not lvwRptParams.SelectedItem Is Nothing Then
                lvwRptParams.Visible = False
                ShowEditRptParam
            End If
        Case vbKeyDelete
    End Select
Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lvwRptParams_KeyDown"
End Sub

Private Sub lvwRptParams_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo EH
    Dim itmX As ListItem
    Dim iDataType As VBA.VbVarType
    If Button = vbLeftButton Then
        Set itmX = lvwRptParams.SelectedItem
        If Not itmX Is Nothing Then
            iDataType = itmX.SubItems(GuiRptParamsListView.ParamDataType - 1)
            If iDataType = vbBoolean Then
                ShowEditRptParam True
            ElseIf iDataType = vbDate Then
                ShowEditRptParam True
            End If
        End If
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lvwRptParams_MouseUp"
End Sub

Private Sub Timer_GetAmount_Timer()
    On Error GoTo EH
    Dim itmX As ListItem
    Dim sClassOfLossID As String
    Dim sClassTypeID As String
    Dim sSQL As String
    Dim sThisClassTypeID As String
    Dim lCount As Long
    Dim RS As ADODB.Recordset
    Dim oConn As ADODB.Connection
    
    Timer_GetAmount.Enabled = False
    
    'First need to Select Line of Coverage (Class)
    'Use the selected item from Indemnity to get the class
    Set itmX = mfrmIndemnity.lvwIndemnity.SelectedItem
    sClassOfLossID = itmX.SubItems(GuiIndemListView.ClassOfLossID - 1)
    
    'From the ClassOfLossID need to get the ClassType ID
    '(Class Of Loss is the Carrier Definition of Class type)
    Set oConn = New ADODB.Connection
    Set RS = New ADODB.Recordset
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    sSQL = "SELECT  ClassTypeID "
    sSQL = sSQL & "FROM ClassOfLoss "
    sSQL = sSQL & "WHERE ClassOfLossID = " & sClassOfLossID & " "
    Set RS = New ADODB.Recordset
    RS.CursorLocation = adUseClient
    RS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    
    Set RS.ActiveConnection = Nothing
    If RS.RecordCount = 1 Then
        sClassTypeID = goUtil.IsNullIsVbNullString(RS.Fields("ClassTypeID"))
    End If
    
    'Select the correct one
    For lCount = 0 To cboClassOfLoss.ListCount - 1
        sThisClassTypeID = cboClassOfLoss.ItemData(lCount)
        If StrComp(sClassTypeID, sThisClassTypeID, vbTextCompare) = 0 Then
            cboClassOfLoss.ListIndex = lCount
            Exit For
        End If
    Next
    
    chkAmountOfCheck.Value = vbChecked
    
    If lstvbUserDefinedType.Visible Then
        mbSaveEnabled = cmdSave.Enabled
        cmdSave.Enabled = False
        cmdSave.Default = False
        lstvbUserDefinedType.SetFocus
    End If
    
    'Cleanup
    Set itmX = Nothing
    Set RS = Nothing
    Set oConn = Nothing
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Timer_GetAmount_Timer"
End Sub

Private Sub txtAddress_Change()
    cmdSave.Enabled = True
End Sub

Private Sub txtAddress_GotFocus()
    goUtil.utSelText txtAddress
End Sub

Private Sub txtAddress_LostFocus()
    goUtil.utValidate , txtAddress
End Sub

Private Sub txtAmountOfCheck_Change()
    cmdSave.Enabled = True
End Sub

Private Sub txtAmountOfCheck_GotFocus()
    goUtil.utSelText txtAmountOfCheck
End Sub

Public Function UpdateEdit(Optional pbMouseClickBoolean As Boolean = False) As Boolean
    On Error GoTo EH
    Dim itmX As ListItem
    Dim MyParam As V2ECKeyBoard.MiscReportParam
    Dim iDataType As VBA.VbVarType
    Dim iMyIcon As Long
    Dim sFlagText As String
    
    Screen.MousePointer = MousePointerConstants.vbHourglass
    'First validate the ParamValue
    If txtParamValue.Visible Then
        goUtil.utValidate , txtParamValue
    End If
    
    Set itmX = lvwRptParams.SelectedItem
    
    If itmX Is Nothing Then
        Screen.MousePointer = MousePointerConstants.vbDefault
        Exit Function
    End If
    
    With MyParam
        .MiscReportParamID = itmX.SubItems(GuiRptParamsListView.MiscReportParamID - 1)
        .AssignmentsID = itmX.SubItems(GuiRptParamsListView.AssignmentsID - 1)
        .ID = itmX.SubItems(GuiRptParamsListView.ID - 1)
        .IDAssignments = itmX.SubItems(GuiRptParamsListView.IDAssignments - 1)
        .Number = itmX.SubItems(GuiRptParamsListView.Number - 1)
        .ProjectName = itmX.SubItems(GuiRptParamsListView.ProjectName - 1)
        .ClassName = itmX.SubItems(GuiRptParamsListView.ClassName - 1)
        .ParamName = itmX.SubItems(GuiRptParamsListView.ParamName - 1)
        .ParamCaption = itmX.Text
        'Update .ParamValue Below
        .ParamDataType = itmX.SubItems(GuiRptParamsListView.ParamDataType - 1)
        .SortMe = itmX.SubItems(GuiRptParamsListView.SortMe - 1)
        .IsDeleted = goUtil.GetFlagFromText(itmX.SubItems(GuiRptParamsListView.IsDeleted - 1))
        .DownLoadMe = goUtil.GetFlagFromText(itmX.SubItems(GuiRptParamsListView.DownLoadMe - 1))
        .UpLoadMe = "True"
        sFlagText = goUtil.GetFlagText(True)
        itmX.SubItems(GuiRptParamsListView.UpLoadMe - 1) = sFlagText
        itmX.ListSubItems(GuiRptParamsListView.UpLoadMe - 1).ReportIcon = GuiRptParamsStatusList.UpLoadMe
        .AdminComments = itmX.SubItems(GuiRptParamsListView.AdminComments - 1)
        .DateLastUpdated = Format(Now(), "MM/DD/YYYY HH:MM:SS")
        itmX.SubItems(GuiRptParamsListView.DateLastUpdated - 1) = .DateLastUpdated
        itmX.SubItems(GuiRptParamsListView.DateLastUpdatedSort - 1) = Format(.DateLastUpdated, "YYYY/MM/DD HH:MM:SS")
        .UpdateByUserID = goUtil.gsCurUsersID
        itmX.SubItems(GuiRptParamsListView.UpdateByUserID - 1) = .UpdateByUserID
    End With
    
    'Set the Param Value Here according to the Data Type
    iDataType = itmX.SubItems(GuiRptParamsListView.ParamDataType - 1)
    Select Case iDataType
        Case VBA.VbVarType.vbBoolean
            If chkParamBoolean.Value = vbChecked Then
                sFlagText = goUtil.GetFlagText(True)
                iMyIcon = GuiRptParamsStatusList.ValueIsChecked
                MyParam.ParamValue = "-1"
            Else
                sFlagText = goUtil.GetFlagText(False)
                iMyIcon = GuiRptParamsStatusList.ValueIsUnchecked
                MyParam.ParamValue = "0"
            End If
            itmX.SubItems(GuiRptParamsListView.ParamValue - 1) = sFlagText
            itmX.ListSubItems(GuiRptParamsListView.ParamValue - 1).ReportIcon = iMyIcon
            
        Case VBA.VbVarType.vbCurrency, VBA.VbVarType.vbDecimal, VBA.VbVarType.vbDouble, VBA.VbVarType.vbLong, VBA.VbVarType.vbInteger
            MyParam.ParamValue = txtParamValue.Text
            itmX.SubItems(GuiRptParamsListView.ParamValue - 1) = MyParam.ParamValue
        Case VBA.VbVarType.vbDate
            MyParam.ParamValue = txtParamValue.Text
            itmX.SubItems(GuiRptParamsListView.ParamValue - 1) = MyParam.ParamValue
        Case VBA.VbVarType.vbUserDefinedType
            MyParam.ParamValue = txtParamValue.Text
            itmX.SubItems(GuiRptParamsListView.ParamValue - 1) = MyParam.ParamValue
        Case VBA.VbVarType.vbString
            MyParam.ParamValue = txtParamValue.Text
            itmX.SubItems(GuiRptParamsListView.ParamValue - 1) = MyParam.ParamValue
    End Select
    
    'now Edit this param
    
    mfrmClaim.EditRptParamItem MyParam
    
    If Not pbMouseClickBoolean Then
        'Set Ecit button back to the defualt Cancel
        cmdExit.Cancel = True
        cmdUpdateEdit.Default = False
        framEditParam.Visible = False
        lvwRptParams.Visible = True
        cmdSelAll.Enabled = True
        cmdFind.Enabled = True
        cmdFindNext.Enabled = True
        cmdPrintReport.Enabled = True
'        framMultiUpdate.Visible = True
        cmdSpelling.Enabled = True
        lvwRptParams.SetFocus
    End If
    
   
    Screen.MousePointer = MousePointerConstants.vbDefault
    Exit Function
EH:
    Screen.MousePointer = MousePointerConstants.vbDefault
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdUpdateEdit_Click"
End Function

Public Function ShowEditRptParam(Optional pbMouseClickBoolean As Boolean = False) As Boolean
    On Error GoTo EH
    Dim iDataType As VBA.VbVarType
    Dim itmX As ListItem
    Dim sFlagText As String
    Dim sLossFormat As String
    
    If lvwRptParams.SelectedItem Is Nothing Then
        sLossFormat = goUtil.IsNullIsVbNullString(mfrmClaim.adoRSAssignments.Fields("LRFormat"))
        If StrComp(sLossFormat, "V2ECcarFarmers.clsLossXML01", vbTextCompare) = 0 Then
            For Each itmX In lvwRptParams.ListItems
                If StrComp(itmX.Text, "Contact Payee Id", vbTextCompare) = 0 Then
                    itmX.Selected = True
                    Exit For
                End If
            Next
        Else
            Exit Function
        End If
    End If
    If lvwRptParams.SelectedItem Is Nothing Then
        Exit Function
    End If
    
    mbShowingEditRptParam = True
    Set itmX = lvwRptParams.SelectedItem
    txtParamCaption.Text = itmX.Text
    
    If Not pbMouseClickBoolean Then
        framEditParam.Visible = True
        cmdCancelEdit.Cancel = True
        framEditParam.Visible = True
        cmdFind.Enabled = False
        cmdSelAll.Enabled = False
        cmdFindNext.Enabled = False
        cmdPrintReport.Enabled = False
        cmdUpdateEdit.Enabled = True
'        cmdUpdateEdit.Default = True
        framMultiUpdate.Visible = False
        cmdSpelling.Enabled = False
    End If
    
    iDataType = itmX.SubItems(GuiRptParamsListView.ParamDataType - 1)
    Select Case iDataType
        Case VBA.VbVarType.vbBoolean
            chkParamBoolean.Visible = True
            cmdParamDate.Visible = False
            txtParamValue.Visible = False
            sFlagText = itmX.SubItems(GuiRptParamsListView.ParamValue - 1)
            chkParamBoolean.Caption = "(Check or Uncheck this Item)"
            If goUtil.GetFlagFromText(sFlagText) Then
                If pbMouseClickBoolean Then
                    'Switch the Value if this is A boolean and then update
                    chkParamBoolean.Value = vbUnchecked
                    UpdateEdit pbMouseClickBoolean
                Else
                    chkParamBoolean.Value = vbChecked
                    chkParamBoolean.SetFocus
                End If
            Else
                If pbMouseClickBoolean Then
                    'Switch the Value if this is A boolean and then update
                    chkParamBoolean.Value = vbChecked
                    UpdateEdit pbMouseClickBoolean
                Else
                    chkParamBoolean.Value = vbUnchecked
                    chkParamBoolean.SetFocus
                End If
            End If
            
        Case VBA.VbVarType.vbCurrency, VBA.VbVarType.vbDecimal, VBA.VbVarType.vbDouble, VBA.VbVarType.vbLong, VBA.VbVarType.vbInteger
            chkParamBoolean.Visible = False
            cmdParamDate.Visible = False
            txtParamValue.Visible = True
            txtParamValue.Locked = False
            txtParamValue.Tag = "Numeric"
            'Numbers have diff max len
            Select Case iDataType
                Case VBA.VbVarType.vbLong
                    txtParamValue.MaxLength = 9
                Case VBA.VbVarType.vbInteger
                    txtParamValue.MaxLength = 5
                Case Else
                    txtParamValue.MaxLength = 10
            End Select
            txtParamValue.Text = itmX.SubItems(GuiRptParamsListView.ParamValue - 1)
            txtParamValue.SetFocus
        Case VBA.VbVarType.vbDate
            If pbMouseClickBoolean Then
                txtParamValue.Tag = "Date"
                txtParamValue.MaxLength = 20
                txtParamValue.Text = itmX.SubItems(GuiRptParamsListView.ParamValue - 1)
                MyGUI.ShowCalendar txtParamValue
                UpdateEdit pbMouseClickBoolean
            Else
                chkParamBoolean.Visible = False
                cmdParamDate.Visible = True
                txtParamValue.Visible = True
                txtParamValue.Locked = True
                txtParamValue.Tag = "Date"
                txtParamValue.MaxLength = 20
                txtParamValue.Text = itmX.SubItems(GuiRptParamsListView.ParamValue - 1)
                cmdParamDate.SetFocus
                cmdParamDate_Click
            End If
        Case VBA.VbVarType.vbUserDefinedType
            'Check for LRFOrmat
            If StrComp(sLossFormat, "V2ECcarFarmers.clsLossXML01", vbTextCompare) = 0 Then
                'Famers XML01 6.1.2005 Need to Show Selection box for
                'All available Farmers Contacts to associate with this Payment Request
                txtParamValue.Tag = vbNullString
                txtParamValue.MaxLength = 500
                txtParamValue.Text = itmX.SubItems(GuiRptParamsListView.ParamValue - 1)
                LoadvbUserDefinedTypeList txtParamValue
            End If
        Case VBA.VbVarType.vbString
            chkParamBoolean.Visible = False
            cmdParamDate.Visible = False
            txtParamValue.Visible = True
            txtParamValue.Locked = False
            txtParamValue.Tag = ""
            txtParamValue.MaxLength = 300
            txtParamValue.Text = itmX.SubItems(GuiRptParamsListView.ParamValue - 1)
            txtParamValue.SetFocus
    End Select
    
    ShowEditRptParam = True
    mbShowingEditRptParam = False
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function ShowEditRptParam"
End Function

Public Sub LoadvbUserDefinedTypeList(poTextBox As Object)
    On Error GoTo EH
    Dim sDescList As String
    Dim saryList() As String
    Dim lCount As Long
    Dim sSelItem As String
    Dim lSelIndex As Long
    Dim sLossFormat As String

    Dim oTextBox As TextBox
    'Wddx Objects
    Dim sLossReportData As String
    Dim sContactRowID As String
    Dim sFirstName As String
    Dim sLastName As String
    Dim sContactRole As String
    Dim oDeser As WDDXDeserializer
    Dim oMyStruct As WDDXStruct
    Dim oContactsRS As WDDXRecordset
    
    sLossFormat = goUtil.IsNullIsVbNullString(mfrmClaim.adoRSAssignments.Fields("LRFormat"))
    If StrComp(sLossFormat, "V2ECcarFarmers.clsLossXML01", vbTextCompare) = 0 Then
        lstvbUserDefinedType.Visible = True
        sLossReportData = goUtil.IsNullIsVbNullString(mfrmClaim.adoRSAssignments.Fields("LossReport"))
        Set oDeser = New WDDXDeserializer
        Set oMyStruct = oDeser.deserialize(sLossReportData)
        Set oContactsRS = oMyStruct.getProp("ContactRS")
        For lCount = 1 To oContactsRS.getRowCount
            sContactRowID = oContactsRS.getField(lCount, "ContactRowID")
            sFirstName = oContactsRS.getField(lCount, "FirstName")
            sLastName = oContactsRS.getField(lCount, "LastName")
            sContactRole = oContactsRS.getField(lCount, "ContactRole")
            If sDescList = vbNullString Then
                sDescList = sLastName & ", " & sFirstName & "  Contact Role: " & sContactRole & String(200, Chr(32)) & "_" & sContactRowID
            Else
                sDescList = sDescList & "|"
                sDescList = sDescList & sLastName & ", " & sFirstName & "  Contact Role: " & sContactRole & String(200, Chr(32)) & "_" & sContactRowID
            End If
        Next
    Else
        Exit Sub
    End If
    
    Set oTextBox = poTextBox
    
    sSelItem = Trim(oTextBox.Text)
    lSelIndex = -1
    
    If sDescList <> vbNullString Then
        saryList() = Split(sDescList, "|")
        'BUbble sort this
        goUtil.utBubbleSort saryList
        lstvbUserDefinedType.Clear
        For lCount = LBound(saryList, 1) To UBound(saryList, 1)
            lstvbUserDefinedType.AddItem saryList(lCount)
            If sSelItem <> vbNullString Then
                If StrComp(lstvbUserDefinedType.List(lstvbUserDefinedType.NewIndex), sSelItem, vbTextCompare) = 0 Then
                    lSelIndex = lstvbUserDefinedType.NewIndex
                End If
            End If
        Next
    Else
        lstvbUserDefinedType.Visible = False
    End If
    
    lstvbUserDefinedType.ListIndex = lSelIndex
    
    
    If lstvbUserDefinedType.ListIndex = -1 Then
        oTextBox = vbNullString
    End If
    
    Set oDeser = Nothing
    Set oMyStruct = Nothing
    Set oContactsRS = Nothing
    Set oTextBox = Nothing
    Exit Sub
EH:

    Set oDeser = Nothing
    Set oMyStruct = Nothing
    Set oContactsRS = Nothing
    Set oTextBox = Nothing
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub LoadvbUserDefinedTypeList"
End Sub


Private Sub txtAmountOfCheck_LostFocus()
    goUtil.utValidate , txtAmountOfCheck
End Sub

Private Sub txtAppliedDeductible_Change()
    cmdSave.Enabled = True
End Sub

Private Sub txtAppliedDeductible_GotFocus()
    goUtil.utSelText txtAppliedDeductible
End Sub

Private Sub txtDateMultiUpdate_GotFocus()
    goUtil.utSelText txtDateMultiUpdate
End Sub

Private Sub txtDateMultiUpdate_LostFocus()
    goUtil.utValidate , txtDateMultiUpdate
End Sub

Private Sub txtIndemTotals_GotFocus()
    goUtil.utSelText txtIndemTotals
End Sub

Private Sub txtInsured_Change()
    cmdSave.Enabled = True
End Sub

Private Sub txtInsured_GotFocus()
    goUtil.utSelText txtInsured
End Sub

Private Sub txtInsured_LostFocus()
    goUtil.utValidate , txtInsured
End Sub

Private Sub txtMortgageeName_Change()
    cmdSave.Enabled = True
End Sub

Private Sub txtMortgageeName_GotFocus()
    goUtil.utSelText txtMortgageeName
End Sub

Private Sub txtMortgageeName_LostFocus()
    goUtil.utValidate , txtMortgageeName
End Sub

Private Sub txtParamValue_GotFocus()
    goUtil.utSelText txtParamValue
End Sub

Private Sub txtParamValue_LostFocus()
    goUtil.utValidate , txtParamValue
End Sub
