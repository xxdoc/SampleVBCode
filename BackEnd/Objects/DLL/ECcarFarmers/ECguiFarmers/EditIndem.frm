VERSION 5.00
Begin VB.Form EditIndem 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Indemnity"
   ClientHeight    =   7845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11550
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "EditIndem.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   11550
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame framEdit 
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   11295
      Begin VB.Frame framIndemAmounts 
         Caption         =   "Indemnity Amounts"
         Height          =   4815
         Left            =   5760
         TabIndex        =   25
         Top             =   1560
         Width           =   5415
         Begin VB.Frame framHelper 
            Height          =   735
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Width           =   5175
            Begin VB.CheckBox chkViewCalc 
               Height          =   360
               Left            =   4680
               Picture         =   "EditIndem.frx":000C
               Style           =   1  'Graphical
               TabIndex        =   27
               ToolTipText     =   "View Calculator"
               Top             =   240
               Width           =   375
            End
         End
         Begin VB.TextBox txtACVLessExcessLimits 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E9E9E9&
            Height          =   375
            Left            =   3120
            Locked          =   -1  'True
            MaxLength       =   14
            TabIndex        =   43
            Tag             =   "Currency"
            Top             =   4320
            Width           =   2175
         End
         Begin VB.TextBox txtMiscDescription 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   720
            Left            =   3120
            MaxLength       =   50
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   41
            Tag             =   "UCASE"
            Top             =   3360
            Width           =   2175
         End
         Begin VB.TextBox txtMiscellaneous 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   3120
            MaxLength       =   14
            TabIndex        =   39
            Tag             =   "Currency"
            Top             =   3000
            Width           =   2175
         End
         Begin VB.TextBox txtExcessLimits 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E9E9E9&
            Height          =   375
            Left            =   3120
            Locked          =   -1  'True
            MaxLength       =   14
            TabIndex        =   37
            Tag             =   "Currency"
            Top             =   2640
            Width           =   2175
         End
         Begin VB.TextBox txtACVClaim 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E9E9E9&
            Height          =   375
            Left            =   3120
            Locked          =   -1  'True
            MaxLength       =   14
            TabIndex        =   35
            Tag             =   "Currency"
            Top             =   2280
            Width           =   2175
         End
         Begin VB.TextBox txtNonRecoverableDep 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   3120
            MaxLength       =   14
            TabIndex        =   33
            Tag             =   "Currency"
            Top             =   1920
            Width           =   2175
         End
         Begin VB.TextBox txtRecoverableDep 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   3120
            MaxLength       =   14
            TabIndex        =   31
            Tag             =   "Currency"
            Top             =   1560
            Width           =   2175
         End
         Begin VB.TextBox txtReplacementCost 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   3120
            MaxLength       =   14
            TabIndex        =   29
            Tag             =   "Currency"
            Top             =   1200
            Width           =   2175
         End
         Begin VB.Line Line2 
            X1              =   105
            X2              =   5305
            Y1              =   4200
            Y2              =   4200
         End
         Begin VB.Line Line1 
            X1              =   105
            X2              =   5305
            Y1              =   4260
            Y2              =   4260
         End
         Begin VB.Label lblMiscDesc 
            Alignment       =   1  'Right Justify
            Caption         =   "Misc. Description:"
            Enabled         =   0   'False
            Height          =   255
            Left            =   240
            TabIndex        =   40
            Top             =   3600
            Width           =   2775
         End
         Begin VB.Label lblFullCostOfRepair 
            Alignment       =   1  'Right Justify
            Caption         =   "Full Cost of Repair/Replacement:"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   1260
            Width           =   2895
         End
         Begin VB.Label lblRecoverableDepreciation 
            Alignment       =   1  'Right Justify
            Caption         =   "Recoverable Depreciation:"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   1620
            Width           =   2895
         End
         Begin VB.Label lblACVLoss 
            Alignment       =   1  'Right Justify
            Caption         =   "Actual Cash Value:"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   2340
            Width           =   2895
         End
         Begin VB.Label lblNonRecovDepr 
            Alignment       =   1  'Right Justify
            Caption         =   "Non-Recoverable Depreciation:"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   1980
            Width           =   2895
         End
         Begin VB.Label lblLessExcessLimits 
            Alignment       =   1  'Right Justify
            Caption         =   "Less Excess Limits:"
            Height          =   255
            Left            =   240
            TabIndex        =   36
            Top             =   2700
            Width           =   2775
         End
         Begin VB.Label lblLessMisc 
            Alignment       =   1  'Right Justify
            Caption         =   "Less Miscellaneous:"
            Height          =   255
            Left            =   240
            TabIndex        =   38
            Top             =   3060
            Width           =   2775
         End
         Begin VB.Label lblNetActualCashValueClaim 
            Alignment       =   1  'Right Justify
            Caption         =   "ACVC:"
            Height          =   255
            Left            =   240
            TabIndex        =   42
            Top             =   4380
            Width           =   2775
         End
      End
      Begin VB.Frame framPreviousPayment 
         Caption         =   "Previous Payment"
         Enabled         =   0   'False
         Height          =   1455
         Left            =   120
         TabIndex        =   15
         Top             =   4920
         Width           =   5535
         Begin VB.TextBox txtPPayCheckNumber 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   375
            Left            =   3240
            MaxLength       =   20
            TabIndex        =   22
            Tag             =   "UCASE"
            Top             =   960
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.CommandButton cmdPPayDatePaid 
            Enabled         =   0   'False
            Height          =   375
            Left            =   5025
            Picture         =   "EditIndem.frx":0596
            Style           =   1  'Graphical
            TabIndex        =   18
            Tag             =   "Date"
            Top             =   240
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txtPPayDatePaid 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   360
            Left            =   3240
            TabIndex        =   17
            Tag             =   "Date"
            Top             =   240
            Visible         =   0   'False
            Width           =   1680
         End
         Begin VB.TextBox txtPPayAmountPaid 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   375
            Left            =   3240
            MaxLength       =   14
            TabIndex        =   20
            Tag             =   "Currency"
            Top             =   600
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.Label lblPPayCheckNumber 
            Alignment       =   1  'Right Justify
            Caption         =   "Check Number:"
            Enabled         =   0   'False
            Height          =   255
            Left            =   240
            TabIndex        =   21
            Top             =   1020
            Visible         =   0   'False
            Width           =   2895
         End
         Begin VB.Label lblPPayDatePaid 
            Alignment       =   1  'Right Justify
            Caption         =   "Date Paid:"
            Enabled         =   0   'False
            Height          =   255
            Left            =   1560
            TabIndex        =   16
            Top             =   293
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Label lbPPayAmountPaid 
            Alignment       =   1  'Right Justify
            Caption         =   "Amount Paid:"
            Enabled         =   0   'False
            Height          =   255
            Left            =   240
            TabIndex        =   19
            Top             =   660
            Visible         =   0   'False
            Width           =   2895
         End
      End
      Begin VB.CheckBox chkEnableSpecialLimits 
         Caption         =   "Is there a special limit involved for this item?"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1560
         Width           =   5535
      End
      Begin VB.TextBox txtDescription 
         Height          =   960
         Left            =   5760
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   24
         Tag             =   "UCASE"
         Top             =   480
         Width           =   5415
      End
      Begin VB.Frame framSpecialLimits 
         Caption         =   "Special Limits"
         Enabled         =   0   'False
         Height          =   2655
         Left            =   120
         TabIndex        =   6
         Top             =   1800
         Width           =   5535
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
            Left            =   3960
            Locked          =   -1  'True
            TabIndex        =   11
            Tag             =   "Currency"
            ToolTipText     =   "Applied Deductible For this Item"
            Top             =   1680
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.CheckBox chkExcessAbsorbsDeductible 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            Caption         =   "ACV in excess of limitation will absorb deductible."
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   9
            Top             =   1440
            Visible         =   0   'False
            Width           =   5055
         End
         Begin VB.TextBox txtSpecialLimits 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   375
            Left            =   3240
            MaxLength       =   14
            TabIndex        =   13
            Tag             =   "Currency"
            Top             =   2160
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.CheckBox chkIsAddAmountOfInsurance 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            Caption         =   "Is additional amount of insurance."
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   8
            Top             =   1200
            Visible         =   0   'False
            Width           =   5055
         End
         Begin VB.Label lblAppliedDed 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            Caption         =   "Applied Deductible:"
            Height          =   255
            Left            =   240
            TabIndex        =   10
            Top             =   1680
            Visible         =   0   'False
            Width           =   3735
         End
         Begin VB.Label lblMess 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   $"EditIndem.frx":09D8
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   1815
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Visible         =   0   'False
            Width           =   5295
         End
         Begin VB.Label lblSpecialLimits 
            Alignment       =   1  'Right Justify
            Caption         =   "Amount of Limitation:"
            Enabled         =   0   'False
            Height          =   255
            Left            =   240
            TabIndex        =   12
            Top             =   2220
            Visible         =   0   'False
            Width           =   2895
         End
      End
      Begin VB.ComboBox cboTypeOfLoss 
         CausesValidation=   0   'False
         Height          =   360
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1080
         Width           =   5415
      End
      Begin VB.ComboBox cboClassOfLoss 
         CausesValidation=   0   'False
         Height          =   360
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   480
         Width           =   5415
      End
      Begin VB.CheckBox chkIsPreviousPayment 
         Caption         =   "Is this item a previous payment?"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   4680
         Width           =   5415
      End
      Begin VB.Label lblDescription 
         Caption         =   "Description:"
         Height          =   255
         Left            =   5760
         TabIndex        =   23
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label lblTypeOfLoss 
         Caption         =   "Type Of Loss:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   3255
      End
      Begin VB.Label lblClassOfLoss 
         Caption         =   "Line Of Coverage (Class):"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.Frame framCommands 
      Height          =   1215
      Left            =   8040
      TabIndex        =   44
      Top             =   6480
      Width           =   3375
      Begin VB.CommandButton cmdSpelling 
         Caption         =   "Spellin&g"
         Height          =   855
         Left            =   120
         MaskColor       =   &H00000000&
         Picture         =   "EditIndem.frx":0A80
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Exit"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Sa&ve"
         Default         =   -1  'True
         Enabled         =   0   'False
         Height          =   855
         Left            =   1200
         MaskColor       =   &H00000000&
         Picture         =   "EditIndem.frx":0ECA
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Exit"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   855
         Left            =   2280
         MaskColor       =   &H00000000&
         Picture         =   "EditIndem.frx":1014
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Exit"
         Top             =   240
         Width           =   975
      End
   End
End
Attribute VB_Name = "EditIndem"
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
Private mbLoading As Boolean
Private mlCOLOrigListIndex As Long
Private moCurrentTextBox As TextBox


Private mbAdding As Boolean
Private msIndemID As String
Private Const BG_COLOR_GRAY = &HE0E0E0
Private Const BG_COLOR_DRKGRAY = &H8000000F
Private Const BG_COLOR_WHITE = &H80000005
Private Const BG_COLOR_YELLOW = &H80000018

Public Property Let CurrentTextBox(poTextBox As Object)
    On Error GoTo EH
    Set moCurrentTextBox = poTextBox
    If moCurrentTextBox Is Nothing Then
        cmdSpelling.Enabled = False
    Else
        cmdSpelling.Enabled = True
    End If
    Exit Property
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Property Let CurrentTextBox"
End Property
Public Property Set CurrentTextBox(poTextBox As Object)
    On Error GoTo EH
    Set moCurrentTextBox = poTextBox
    If moCurrentTextBox Is Nothing Then
        cmdSpelling.Enabled = False
    Else
        cmdSpelling.Enabled = True
    End If
    Exit Property
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Property Set CurrentTextBox"
End Property
Public Property Get CurrentTextBox() As Object
    Set CurrentTextBox = moCurrentTextBox
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

Public Property Let COLOrigListIndex(pIndex As Long)
    mlCOLOrigListIndex = pIndex
End Property
Public Property Get COLOrigListIndex() As Long
    COLOrigListIndex = mlCOLOrigListIndex
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

Public Property Let IndemID(psID As String)
    msIndemID = psID
End Property

Public Property Let Adding(pbFlag As Boolean)
    mbAdding = pbFlag
End Property

Public Property Let Loading(pbFlag As Boolean)
    mbLoading = pbFlag
End Property
Public Property Get Loading() As Boolean
    Loading = mbLoading
End Property

Private Property Get msClassName() As String
    msClassName = Me.Name
End Property


Private Sub cboClassOfLoss_Click()
    cmdSave.Enabled = True
End Sub

Private Sub cboTypeOfLoss_Click()
    cmdSave.Enabled = True
End Sub

Private Sub chkEnableSpecialLimits_Click()
    On Error GoTo EH
    
    If mbLoading Then
        Exit Sub
    End If
    
    If chkEnableSpecialLimits.Value = vbChecked Then
        framSpecialLimits.Enabled = True
        lblMess.Enabled = True
        lblMess.Visible = True
        chkIsAddAmountOfInsurance.Enabled = True
        chkIsAddAmountOfInsurance.Visible = True
        chkExcessAbsorbsDeductible.Enabled = True
        chkExcessAbsorbsDeductible.Visible = True
        lblSpecialLimits.Enabled = True
        lblSpecialLimits.Visible = True
        txtSpecialLimits.Enabled = True
        txtSpecialLimits.Visible = True
        txtSpecialLimits.BackColor = BG_COLOR_WHITE
        lblAppliedDed.Enabled = True
        lblAppliedDed.Visible = True
        txtAppliedDeductible.Enabled = True
        txtAppliedDeductible.Visible = True
        txtAppliedDeductible.BackColor = BG_COLOR_YELLOW
    Else
        framSpecialLimits.Enabled = False
        lblMess.Enabled = False
        lblMess.Visible = False
        chkIsAddAmountOfInsurance.Enabled = False
        chkIsAddAmountOfInsurance.Visible = False
        chkExcessAbsorbsDeductible.Enabled = False
        chkExcessAbsorbsDeductible.Visible = False
        chkIsAddAmountOfInsurance.Value = vbUnchecked
        chkExcessAbsorbsDeductible.Value = vbChecked
        lblSpecialLimits.Enabled = False
        lblSpecialLimits.Visible = False
        txtSpecialLimits.Enabled = False
        txtSpecialLimits.Visible = False
        txtSpecialLimits.BackColor = BG_COLOR_DRKGRAY
        txtSpecialLimits.Text = vbNullString
        lblAppliedDed.Enabled = False
        lblAppliedDed.Visible = False
        txtAppliedDeductible.Enabled = False
        txtAppliedDeductible.Visible = False
        txtAppliedDeductible.BackColor = BG_COLOR_DRKGRAY
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub chkEnableSpecialLimits_Click"
End Sub

Private Sub chkExcessAbsorbsDeductible_Click()
    On Error GoTo EH
    
    If mbLoading Then
        Exit Sub
    End If
    cmdSave.Enabled = True
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub chkExcessAbsorbsDeductible_Click"
End Sub

Private Sub chkIsAddAmountOfInsurance_Click()
    On Error GoTo EH
    
    If mbLoading Then
        Exit Sub
    End If
    cmdSave.Enabled = True
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub chkIsAddAmountOfInsurance_Click"
End Sub

Private Sub chkIsPreviousPayment_Click()
    On Error GoTo EH
    
    If mbLoading Then
        Exit Sub
    End If
    
    If chkIsPreviousPayment.Value = vbChecked Then
        framPreviousPayment.Enabled = True
        lblPPayDatePaid.Enabled = True
        lblPPayDatePaid.Visible = True
        lbPPayAmountPaid.Enabled = True
        lbPPayAmountPaid.Visible = True
        lblPPayCheckNumber.Enabled = True
        lblPPayCheckNumber.Visible = True
        txtPPayDatePaid.Enabled = True
        txtPPayDatePaid.Visible = True
        txtPPayDatePaid.BackColor = BG_COLOR_WHITE
        cmdPPayDatePaid.Enabled = True
        cmdPPayDatePaid.Visible = True
        txtPPayAmountPaid.Enabled = True
        txtPPayAmountPaid.Visible = True
        txtPPayAmountPaid.BackColor = BG_COLOR_WHITE
        txtPPayCheckNumber.Enabled = True
        txtPPayCheckNumber.Visible = True
        txtPPayCheckNumber.BackColor = BG_COLOR_WHITE
    Else
        framPreviousPayment.Enabled = False
        lblPPayDatePaid.Enabled = False
        lblPPayDatePaid.Visible = False
        lbPPayAmountPaid.Enabled = False
        lbPPayAmountPaid.Visible = False
        lblPPayCheckNumber.Enabled = False
        lblPPayCheckNumber.Visible = False
        txtPPayDatePaid.Enabled = False
        txtPPayDatePaid.Visible = False
        txtPPayDatePaid.BackColor = BG_COLOR_DRKGRAY
        txtPPayDatePaid.Text = vbNullString
        cmdPPayDatePaid.Enabled = False
        cmdPPayDatePaid.Visible = False
        txtPPayAmountPaid.Enabled = False
        txtPPayAmountPaid.Visible = False
        txtPPayAmountPaid.BackColor = BG_COLOR_DRKGRAY
        txtPPayAmountPaid.Text = vbNullString
        txtPPayCheckNumber.Enabled = False
        txtPPayCheckNumber.Visible = False
        txtPPayCheckNumber.BackColor = BG_COLOR_DRKGRAY
        txtPPayCheckNumber.Text = vbNullString
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub chkIsPreviousPayment_Click"
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

Private Sub cmdExit_Click()
    On Error GoTo EH

    If cmdSave.Enabled Then
        If MsgBox("Do you want to save changes?", vbYesNo + vbQuestion, "Save Changes") = vbYes Then
            If Not SaveMe() Then
                Exit Sub
            End If
        End If
    ElseIf mlCOLOrigListIndex = -1 Then
        cmdSave.Enabled = True
    End If
    mbUnloadMe = True
    Me.Visible = False

    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdExit_Click"
End Sub

Private Sub cmdPPayDatePaid_Click()
    On Error GoTo EH
    
    mfrmClaim.ShowCalendar txtPPayDatePaid
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdPPayDatePaid_Click"
End Sub

Private Sub cmdSave_Click()
    On Error GoTo EH

    If SaveMe() Then
        Me.Visible = False
    End If

    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdSave_Click"
End Sub

Private Sub cmdSpelling_Click()
    On Error GoTo EH
    cmdSpelling.Enabled = False
    If Not CurrentTextBox Is Nothing Then
        goUtil.utLoadSP
        goUtil.goSP.CheckSP CurrentTextBox
        If CurrentTextBox.Visible Then
            CurrentTextBox.SetFocus
            Sleep 100
            cmdSpelling.Enabled = False
        End If
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdSpelling_Click"
End Sub

Private Sub Form_Load()
    On Error GoTo EH
    mbLoading = True
    Dim sSQL As String
    Dim oConn As ADODB.Connection
    Dim adoRS As ADODB.Recordset
    Dim adoClassTypeIDInRS As ADODB.Recordset
    Dim sLossFormat As String
    Dim bIncludeDeletedItems As Boolean
    
    'Check the Loss Format
    sLossFormat = goUtil.IsNullIsVbNullString(mfrmClaim.adoRSAssignments.Fields("LRFormat"))
    If StrComp(sLossFormat, "V2ECcarFarmers.clsLossXML01", vbTextCompare) = 0 Then
        bIncludeDeletedItems = True
    End If
    
    'Set the Form Icon
    Me.Icon = mfrmClaim.optClaim(GuiClaimOptions.opt03_Indemnity).Picture
    
'    LoadHeaderlvwRptParams
'    LoadReports
    
    
    sSQL = "SELECT "
    sSQL = sSQL & "[RTIndemnityID], "
    sSQL = sSQL & "[AssignmentsID], "
    sSQL = sSQL & "[RTChecksID], "
    sSQL = sSQL & "[ID], "
    sSQL = sSQL & "[IDAssignments], "
    sSQL = sSQL & "[IDRTChecks], "
    sSQL = sSQL & "[ACVClaim], "
    sSQL = sSQL & "[ACVLessExcessLimits], "
    sSQL = sSQL & "[SpecialLimits], "
    sSQL = sSQL & "[ExcessLimits], "
    sSQL = sSQL & "[Miscellaneous], "
    sSQL = sSQL & "[MiscDescription], "
    sSQL = sSQL & "[IsAddAmountOfInsurance], "
    sSQL = sSQL & "[ExcessAbsorbsDeductible], "
    sSQL = sSQL & "[AppliedDeductible], "
    sSQL = sSQL & "[NonRecoverableDep], "
    sSQL = sSQL & "[RecoverableDep], "
    sSQL = sSQL & "[ReplacementCost], "
    sSQL = sSQL & "[TypeOfLossID], "
    sSQL = sSQL & "[ClassOfLossID], "
    sSQL = sSQL & "[Description], "
    sSQL = sSQL & "[IsPreviousPayment], "
    sSQL = sSQL & "[PPayDatePaid], "
    sSQL = sSQL & "[PPayAmountPaid], "
    sSQL = sSQL & "[PPayCheckNumber], "
    sSQL = sSQL & "[IsDeleted], "
    sSQL = sSQL & "[DownLoadMe], "
    sSQL = sSQL & "[UpLoadMe], "
    sSQL = sSQL & "[AdminComments], "
    sSQL = sSQL & "[DateLastUpdated], "
    sSQL = sSQL & "[UpdateByUserID] "
    sSQL = sSQL & "FROM RTIndemnity "
    sSQL = sSQL & "WHERE [RTIndemnityID] = " & msIndemID & " "
    
    Set oConn = New ADODB.Connection
    Set adoRS = New ADODB.Recordset
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    adoRS.CursorLocation = adUseClient
    adoRS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set adoRS.ActiveConnection = Nothing
    
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
                        "ClassOfLossID", _
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
                        "TypeOfLossID", _
                        "TypeOfLoss", _
                        "Code"
    End If
    
    txtDescription.Text = goUtil.IsNullIsVbNullString(adoRS.Fields("Description"))
    
    If goUtil.IsNullIsVbNullString(adoRS.Fields("SpecialLimits")) > 0 Then
        chkEnableSpecialLimits.Value = vbChecked
        If adoRS.Fields("IsAddAmountOfInsurance") Then
            chkIsAddAmountOfInsurance.Value = vbChecked
        Else
            chkIsAddAmountOfInsurance.Value = vbUnchecked
        End If
        If adoRS.Fields("ExcessAbsorbsDeductible") Then
            chkExcessAbsorbsDeductible.Value = vbChecked
        Else
            chkExcessAbsorbsDeductible.Value = vbUnchecked
        End If
        framSpecialLimits.Enabled = True
        lblMess.Enabled = True
        lblMess.Visible = True
        chkIsAddAmountOfInsurance.Enabled = True
        chkIsAddAmountOfInsurance.Visible = True
        chkExcessAbsorbsDeductible.Enabled = True
        chkExcessAbsorbsDeductible.Visible = True
        lblSpecialLimits.Enabled = True
        lblSpecialLimits.Visible = True
        txtSpecialLimits.Enabled = True
        txtSpecialLimits.Visible = True
        txtSpecialLimits.BackColor = BG_COLOR_WHITE
        txtSpecialLimits.Text = Format(goUtil.IsNullIsVbNullString(adoRS.Fields("SpecialLimits")), "#,###,###,##0.00")
        lblAppliedDed.Enabled = True
        lblAppliedDed.Visible = True
        txtAppliedDeductible.Enabled = True
        txtAppliedDeductible.Visible = True
        txtAppliedDeductible.BackColor = BG_COLOR_YELLOW
        txtAppliedDeductible.Text = Format(goUtil.IsNullIsVbNullString(adoRS.Fields("AppliedDeductible")), "#,###,###,##0.00")
    Else
        framSpecialLimits.Enabled = False
        lblMess.Enabled = False
        lblMess.Visible = False
        chkIsAddAmountOfInsurance.Enabled = False
        chkIsAddAmountOfInsurance.Visible = False
        chkExcessAbsorbsDeductible.Enabled = False
        chkExcessAbsorbsDeductible.Visible = False
        chkIsAddAmountOfInsurance.Value = vbUnchecked
        chkExcessAbsorbsDeductible.Value = vbChecked
        lblSpecialLimits.Enabled = False
        lblSpecialLimits.Visible = False
        txtSpecialLimits.Enabled = False
        txtSpecialLimits.Visible = False
        txtSpecialLimits.BackColor = BG_COLOR_DRKGRAY
        txtSpecialLimits.Text = vbNullString
        lblAppliedDed.Enabled = False
        lblAppliedDed.Visible = False
        txtAppliedDeductible.Enabled = False
        txtAppliedDeductible.Visible = False
        txtAppliedDeductible.BackColor = BG_COLOR_DRKGRAY
        txtAppliedDeductible.Text = Format(goUtil.IsNullIsVbNullString(adoRS.Fields("AppliedDeductible")), "#,###,###,##0.00")
    End If
    
    If adoRS.Fields("IsPreviousPayment") Then
        chkIsPreviousPayment.Value = vbChecked
        framPreviousPayment.Enabled = True
        lblPPayDatePaid.Enabled = True
        lblPPayDatePaid.Visible = True
        lbPPayAmountPaid.Enabled = True
        lbPPayAmountPaid.Visible = True
        lblPPayCheckNumber.Enabled = True
        lblPPayCheckNumber.Visible = True
        txtPPayDatePaid.Enabled = True
        txtPPayDatePaid.Visible = True
        txtPPayDatePaid.BackColor = BG_COLOR_WHITE
        cmdPPayDatePaid.Enabled = True
        cmdPPayDatePaid.Visible = True
        txtPPayAmountPaid.Enabled = True
        txtPPayAmountPaid.Visible = True
        txtPPayAmountPaid.BackColor = BG_COLOR_WHITE
        txtPPayCheckNumber.Enabled = True
        txtPPayCheckNumber.Visible = True
        txtPPayCheckNumber.BackColor = BG_COLOR_WHITE
        txtPPayDatePaid.Text = Format(goUtil.IsNullIsVbNullString(adoRS.Fields("PPayDatePaid")), "MM/DD/YYYY")
        txtPPayAmountPaid.Text = Format(goUtil.IsNullIsVbNullString(adoRS.Fields("PPayAmountPaid")), "#,###,###,##0.00")
        txtPPayCheckNumber.Text = goUtil.IsNullIsVbNullString(adoRS.Fields("PPayCheckNumber"))
    Else
        chkIsPreviousPayment.Value = vbUnchecked
        framPreviousPayment.Enabled = False
        lblPPayDatePaid.Enabled = False
        lblPPayDatePaid.Visible = False
        lbPPayAmountPaid.Enabled = False
        lbPPayAmountPaid.Visible = False
        lblPPayCheckNumber.Enabled = False
        lblPPayCheckNumber.Visible = False
        txtPPayDatePaid.Enabled = False
        txtPPayDatePaid.Visible = False
        txtPPayDatePaid.BackColor = BG_COLOR_DRKGRAY
        txtPPayDatePaid.Text = vbNullString
        cmdPPayDatePaid.Enabled = False
        cmdPPayDatePaid.Visible = False
        txtPPayAmountPaid.Enabled = False
        txtPPayAmountPaid.Visible = False
        txtPPayAmountPaid.BackColor = BG_COLOR_DRKGRAY
        txtPPayAmountPaid.Text = vbNullString
        txtPPayCheckNumber.Enabled = False
        txtPPayCheckNumber.Visible = False
        txtPPayCheckNumber.BackColor = BG_COLOR_DRKGRAY
        txtPPayCheckNumber.Text = vbNullString
    End If
    
    txtReplacementCost.Text = Format(goUtil.IsNullIsVbNullString(adoRS.Fields("ReplacementCost")), "#,###,###,##0.00")
    txtRecoverableDep.Text = Format(goUtil.IsNullIsVbNullString(adoRS.Fields("RecoverableDep")), "#,###,###,##0.00")
    txtNonRecoverableDep.Text = Format(goUtil.IsNullIsVbNullString(adoRS.Fields("NonRecoverableDep")), "#,###,###,##0.00")
    txtACVClaim.Text = Format(goUtil.IsNullIsVbNullString(adoRS.Fields("ACVClaim")), "#,###,###,##0.00")
    txtExcessLimits.Text = Format(goUtil.IsNullIsVbNullString(adoRS.Fields("ExcessLimits")), "#,###,###,##0.00")
    txtMiscellaneous.Text = Format(goUtil.IsNullIsVbNullString(adoRS.Fields("Miscellaneous")), "#,###,###,##0.00")
    txtMiscDescription.Text = goUtil.IsNullIsVbNullString(adoRS.Fields("MiscDescription"))
    If Val(txtMiscellaneous) <> 0 Then
        txtMiscDescription.Enabled = True
        txtMiscDescription.BackColor = BG_COLOR_WHITE
    Else
        txtMiscDescription.Enabled = False
        txtMiscDescription.BackColor = BG_COLOR_DRKGRAY
    End If
    txtACVLessExcessLimits.Text = Format(goUtil.IsNullIsVbNullString(adoRS.Fields("ACVLessExcessLimits")), "#,###,###,##0.00")
    
    'cleanup
    
    Set adoRS = Nothing
    Set adoClassTypeIDInRS = Nothing
    Set oConn = Nothing
    
    mbLoading = False

    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_Load"
End Sub

Public Function CLEANUP() As Boolean
    On Error GoTo EH

    Set mfrmClaim = Nothing
    Set mfrmIndemnity = Nothing
    Set moGUI = Nothing
    Set moCurrentTextBox = Nothing

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
                If MsgBox("Do you want to save changes?", vbYesNo + vbQuestion, "Save Changes") = vbYes Then
                    If Not SaveMe() Then
                        Cancel = True
                        Exit Sub
                    End If
                End If
            ElseIf mlCOLOrigListIndex = -1 Then
                cmdSave.Enabled = True
            End If
            mbUnloadMe = True
            Me.Visible = False
            Cancel = True
    End Select

    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_QueryUnload"
End Sub

Public Function SaveMe() As Boolean
    On Error GoTo EH
    Dim udtIndem As GuiIndemItem
    Dim oListView As ListView
    Dim itmX As ListItem
    Dim sClassOfLossID As String
    Dim sClassOfLoss As String
    Dim sClassOfLossCode  As String
    Dim sTypeOfLossID As String
    Dim sTypeOfLoss As String
    Dim sTypeOfLossCode As String
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
    If cboClassOfLoss.ListIndex = -1 Then
        bCancel = True
        sMess = sMess & "Must Select Class!" & vbCrLf
        cboClassOfLoss.SetFocus
    ElseIf cboTypeOfLoss.ListIndex = -1 Then
        bCancel = True
        sMess = sMess & "Must Select Type Of Loss!" & vbCrLf
        cboTypeOfLoss.SetFocus
    ElseIf txtDescription.Text = vbNullString Then
        bCancel = True
        sMess = sMess & "Must Enter a Description!" & vbCrLf
        txtDescription.SetFocus
        goUtil.utSelText txtDescription
    ElseIf chkEnableSpecialLimits.Value = vbChecked And Val(txtSpecialLimits.Text) = 0 Then
        bCancel = True
        sMess = sMess & "(Special Limits) - Must Enter Amount of Limitation!" & vbCrLf
        txtSpecialLimits.SetFocus
        goUtil.utSelText txtSpecialLimits
    ElseIf chkIsPreviousPayment.Value = vbChecked And Not IsDate(txtPPayDatePaid.Text) Then
        bCancel = True
        sMess = sMess & "(Previous Payment) - Must Enter Date Paid!" & vbCrLf
        txtPPayDatePaid.SetFocus
        goUtil.utSelText txtPPayDatePaid
    ElseIf chkIsPreviousPayment.Value = vbChecked And Val(txtPPayAmountPaid.Text) = 0 Then
        bCancel = True
        sMess = sMess & "(Previous Payment) - Must Enter Amount Paid!" & vbCrLf
        txtPPayAmountPaid.SetFocus
        goUtil.utSelText txtPPayAmountPaid
    ElseIf chkIsPreviousPayment.Value = vbChecked And txtPPayCheckNumber.Text = vbNullString Then
        bCancel = True
        sMess = sMess & "(Previous Payment) - Must Enter Check Number!" & vbCrLf
        txtPPayCheckNumber.SetFocus
        goUtil.utSelText txtPPayCheckNumber
    ElseIf Val(txtReplacementCost.Text) = 0 Then
        bCancel = True
        sMess = sMess & "Must Enter Full Cost of Repair/Replacement!" & vbCrLf
        txtReplacementCost.SetFocus
        goUtil.utSelText txtReplacementCost
    ElseIf Val(txtMiscellaneous.Text) <> 0 And txtMiscDescription.Text = vbNullString Then
        bCancel = True
        sMess = sMess & "Must Enter Misc. Description!" & vbCrLf
        txtMiscDescription.SetFocus
        goUtil.utSelText txtMiscDescription
    End If
    
    If bCancel Then
        MsgBox sMess, vbOKOnly + vbExclamation, "Could Not Save!"
        Exit Function
    End If
    
    cmdSave.Enabled = False
    
    Set oListView = mfrmIndemnity.lvwIndemnity
    
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
    
    Set oConn = Nothing
    Set RS = Nothing

    With udtIndem
        Set itmX = oListView.ListItems("""" & msIndemID & """")
        .RTIndemnityID = itmX.SubItems(GuiIndemListView.RTIndemnityID - 1)
        .AssignmentsID = itmX.SubItems(GuiIndemListView.AssignmentsID - 1)
        .RTChecksID = itmX.SubItems(GuiIndemListView.RTChecksID - 1)
        .ID = itmX.SubItems(GuiIndemListView.ID - 1)
        .IDAssignments = itmX.SubItems(GuiIndemListView.IDAssignments - 1)
        .IDRTChecks = itmX.SubItems(GuiIndemListView.IDRTChecks - 1)
        .ACVClaim = txtACVClaim.Text
        itmX.SubItems(GuiIndemListView.ACVClaim - 1) = txtACVClaim.Text
        itmX.SubItems(GuiIndemListView.ACVClaimSort - 1) = goUtil.utNumInTextSortFormat(txtACVClaim.Text)
        .ACVLessExcessLimits = txtACVLessExcessLimits.Text
        itmX.SubItems(GuiIndemListView.ACVLessExcessLimits - 1) = txtACVLessExcessLimits.Text
        itmX.SubItems(GuiIndemListView.ACVLessExcessLimitsSort - 1) = goUtil.utNumInTextSortFormat(txtACVLessExcessLimits.Text)
        .SpecialLimits = txtSpecialLimits.Text
        itmX.SubItems(GuiIndemListView.SpecialLimits - 1) = txtSpecialLimits.Text
        itmX.SubItems(GuiIndemListView.SpecialLimitsSort - 1) = goUtil.utNumInTextSortFormat(txtSpecialLimits.Text)
        .ExcessLimits = txtExcessLimits.Text
        itmX.SubItems(GuiIndemListView.ExcessLimits - 1) = txtExcessLimits.Text
        itmX.SubItems(GuiIndemListView.ExcessLimitsSort - 1) = goUtil.utNumInTextSortFormat(txtExcessLimits.Text)
        .Miscellaneous = txtMiscellaneous.Text
        itmX.SubItems(GuiIndemListView.Miscellaneous - 1) = txtMiscellaneous.Text
        itmX.SubItems(GuiIndemListView.MiscellaneousSort - 1) = goUtil.utNumInTextSortFormat(txtMiscellaneous.Text)
        .MiscDescription = txtMiscDescription.Text
        itmX.SubItems(GuiIndemListView.MiscellaneousDesc - 1) = txtMiscDescription.Text
        'IsAddAmountOfInsurance
        If chkIsAddAmountOfInsurance.Value = vbChecked Then
            itmX.ListSubItems(GuiIndemListView.IsAddAmountOfInsurance - 1).ReportIcon = GuiIndemStatusList.IsChecked
            .IsAddAmountOfInsurance = True
        Else
            itmX.ListSubItems(GuiIndemListView.IsAddAmountOfInsurance - 1).ReportIcon = Empty
            .IsAddAmountOfInsurance = False
        End If
        sFlagText = goUtil.GetFlagText(CBool(.IsAddAmountOfInsurance))
        itmX.SubItems(GuiIndemListView.IsAddAmountOfInsurance - 1) = sFlagText
        
        'ExcessAbsorbsDeductible
        If chkExcessAbsorbsDeductible.Value = vbChecked Then
            itmX.ListSubItems(GuiIndemListView.ExcessAbsorbsDeductible - 1).ReportIcon = GuiIndemStatusList.IsChecked
            .ExcessAbsorbsDeductible = True
        Else
            itmX.ListSubItems(GuiIndemListView.ExcessAbsorbsDeductible - 1).ReportIcon = Empty
            .ExcessAbsorbsDeductible = False
        End If
        sFlagText = goUtil.GetFlagText(CBool(.ExcessAbsorbsDeductible))
        itmX.SubItems(GuiIndemListView.ExcessAbsorbsDeductible - 1) = sFlagText
        
        .AppliedDeductible = txtAppliedDeductible.Text
        itmX.SubItems(GuiIndemListView.AppliedDeductible - 1) = txtAppliedDeductible.Text
        itmX.SubItems(GuiIndemListView.AppliedDeductibleSort - 1) = goUtil.utNumInTextSortFormat(txtAppliedDeductible.Text)
        .NonRecoverableDep = txtNonRecoverableDep.Text
        itmX.SubItems(GuiIndemListView.NonRecoverableDep - 1) = txtNonRecoverableDep.Text
        itmX.SubItems(GuiIndemListView.NonRecoverableDepSort - 1) = goUtil.utNumInTextSortFormat(txtNonRecoverableDep.Text)
        .RecoverableDep = txtRecoverableDep.Text
        itmX.SubItems(GuiIndemListView.RecoverableDep - 1) = txtRecoverableDep.Text
        itmX.SubItems(GuiIndemListView.RecoverableDepSort - 1) = goUtil.utNumInTextSortFormat(txtRecoverableDep.Text)
        .ReplacementCost = txtReplacementCost.Text
        itmX.SubItems(GuiIndemListView.ReplacementCost - 1) = txtReplacementCost.Text
        itmX.SubItems(GuiIndemListView.ReplacementCostSort - 1) = goUtil.utNumInTextSortFormat(txtReplacementCost.Text)
        .TypeOfLossID = sTypeOfLossID
        itmX.SubItems(GuiIndemListView.TypeOfLossID - 1) = sTypeOfLossID
        itmX.SubItems(GuiIndemListView.TypeOfLoss - 1) = sTypeOfLoss
        itmX.SubItems(GuiIndemListView.TypeOfLossCode - 1) = sTypeOfLossCode
        .ClassOfLossID = sClassOfLossID
        itmX.SubItems(GuiIndemListView.ClassOfLossID - 1) = sClassOfLossID
        itmX.SubItems(GuiIndemListView.ClassOfLoss - 1) = sClassOfLoss
        itmX.SubItems(GuiIndemListView.ClassOfLossCode - 1) = sClassOfLossCode
        .Description = txtDescription.Text
        itmX.SubItems(GuiIndemListView.Description - 1) = txtDescription.Text
        
        'IsPreviousPayment
        If chkIsPreviousPayment.Value = vbChecked Then
            itmX.ListSubItems(GuiIndemListView.IsPreviousPayment - 1).ReportIcon = GuiIndemStatusList.IsChecked
            .IsPreviousPayment = True
            .PPayDatePaid = Format(txtPPayDatePaid.Text, "MM/DD/YYYY")
            itmX.SubItems(GuiIndemListView.PPayDatePaid - 1) = .PPayDatePaid
            itmX.SubItems(GuiIndemListView.PPayDatePaidSort - 1) = Format(.DateLastUpdated, "YYYY/MM/DD")
            .PPayAmountPaid = txtPPayAmountPaid.Text
            itmX.SubItems(GuiIndemListView.PPayAmountPaid - 1) = txtPPayAmountPaid.Text
            itmX.SubItems(GuiIndemListView.PPayAmountPaidSort - 1) = goUtil.utNumInTextSortFormat(txtPPayAmountPaid.Text)
            .PPayCheckNumber = txtPPayCheckNumber.Text
            itmX.SubItems(GuiIndemListView.PPayCheckNumber - 1) = txtPPayCheckNumber.Text
            itmX.SubItems(GuiIndemListView.PPayCheckNumberSort - 1) = goUtil.utNumInTextSortFormat(txtPPayCheckNumber.Text)
        Else
            itmX.ListSubItems(GuiIndemListView.IsPreviousPayment - 1).ReportIcon = Empty
            .IsPreviousPayment = False
            .PPayDatePaid = "Null"
            itmX.SubItems(GuiIndemListView.PPayDatePaid - 1) = vbNullString
            itmX.SubItems(GuiIndemListView.PPayDatePaidSort - 1) = vbNullString
            .PPayAmountPaid = "0.00"
            itmX.SubItems(GuiIndemListView.PPayAmountPaid - 1) = "0.00"
            itmX.SubItems(GuiIndemListView.PPayAmountPaidSort - 1) = goUtil.utNumInTextSortFormat("0.00")
            .PPayCheckNumber = "Null"
            itmX.SubItems(GuiIndemListView.PPayCheckNumber - 1) = vbNullString
            itmX.SubItems(GuiIndemListView.PPayCheckNumberSort - 1) = vbNullString
        End If
        sFlagText = goUtil.GetFlagText(CBool(.IsPreviousPayment))
        itmX.SubItems(GuiIndemListView.IsPreviousPayment - 1) = sFlagText
        
        .IsDeleted = goUtil.GetFlagFromText(itmX.SubItems(GuiIndemListView.IsDeleted - 1))
        .DownLoadMe = goUtil.GetFlagFromText(itmX.SubItems(GuiIndemListView.DownLoadMe - 1))
        itmX.ListSubItems(GuiIndemListView.UpLoadMe - 1).ReportIcon = GuiIndemStatusList.UpLoadMe
        .UpLoadMe = True
        .AdminComments = itmX.SubItems(GuiIndemListView.AdminComments - 1)
        .DateLastUpdated = Format(Now(), "MM/DD/YYYY HH:MM:SS")
        itmX.SubItems(GuiIndemListView.DateLastUpdated - 1) = .DateLastUpdated
        itmX.SubItems(GuiIndemListView.DateLastUpdatedSort - 1) = Format(.DateLastUpdated, "YYYY/MM/DD HH:MM:SS")
        itmX.SubItems(GuiIndemListView.UpdateByUserID - 1) = goUtil.gsCurUsersID
        .UpdateByUserID = itmX.SubItems(GuiIndemListView.UpdateByUserID - 1)
    End With

    'Edit this entry
    mfrmIndemnity.EditIdemnityItem udtIndem
    Sleep 500
    oListView.SortKey = GuiIndemListView.PaymentRequest
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
Public Sub SumACVClaim()
    On Error GoTo EH
    Dim cReplacementCost As Currency
    Dim cRecoverableDep As Currency
    Dim cNonRecoverableDep As Currency
    Dim cACVClaim As Currency
    
    If IsNumeric(txtReplacementCost.Text) Then
        cReplacementCost = CCur(txtReplacementCost.Text)
    Else
        cReplacementCost = 0
    End If
    
    If IsNumeric(txtRecoverableDep.Text) Then
        cRecoverableDep = CCur(txtRecoverableDep.Text)
    Else
        cRecoverableDep = 0
    End If
    
    If IsNumeric(txtNonRecoverableDep.Text) Then
        cNonRecoverableDep = CCur(txtNonRecoverableDep.Text)
    Else
        cNonRecoverableDep = 0
    End If
    
    cACVClaim = cReplacementCost - (cRecoverableDep + cNonRecoverableDep)
    
    txtACVClaim.Text = Format(cACVClaim, "#,###,###,##0.00")
    
    SumExcessLimits
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub SumACVClaim"
End Sub

Public Sub SumExcessLimits()
    On Error GoTo EH
    Dim cACVClaim As Currency
    Dim cSpecialLimits As Currency
    Dim cExcessLimits As Currency
    
    If chkEnableSpecialLimits.Value = vbChecked Then
        If IsNumeric(txtACVClaim.Text) Then
            cACVClaim = CCur(txtACVClaim.Text)
        Else
            cACVClaim = 0
        End If
        If IsNumeric(txtSpecialLimits.Text) Then
            cSpecialLimits = CCur(txtSpecialLimits.Text)
        Else
            cSpecialLimits = 0
        End If
        cExcessLimits = cACVClaim - cSpecialLimits
        If cExcessLimits > 0 Then
            txtExcessLimits.Text = Format(cExcessLimits, "#,###,###,##0.00")
        Else
            txtExcessLimits.Text = "0.00"
        End If
    Else
        txtExcessLimits.Text = "0.00"
    End If
    
    SumACVLessExcessLimits
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub SumExcessLimits"
End Sub

Public Sub SumACVLessExcessLimits()
    On Error GoTo EH
    Dim cACVClaim As Currency
    Dim cExcessLimits As Currency
    Dim cMiscellaneous As Currency
    Dim cACVLessExcessLimits As Currency
    
    
    If IsNumeric(txtACVClaim.Text) Then
        cACVClaim = CCur(txtACVClaim.Text)
    Else
        cACVClaim = 0
    End If
    If IsNumeric(txtExcessLimits.Text) Then
        cExcessLimits = CCur(txtExcessLimits.Text)
    Else
        cExcessLimits = 0
    End If
     If IsNumeric(txtMiscellaneous.Text) Then
        cMiscellaneous = CCur(txtMiscellaneous.Text)
    Else
        cMiscellaneous = 0
    End If
    
    cACVLessExcessLimits = cACVClaim - (cExcessLimits + cMiscellaneous)
    
    txtACVLessExcessLimits.Text = Format(cACVLessExcessLimits, "#,###,###,##0.00")
    
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub SumACVLessExcessLimits"
End Sub

Private Sub txtACVClaim_GotFocus()
    goUtil.utSelText txtACVClaim
End Sub

Private Sub txtACVClaim_LostFocus()
    goUtil.utValidate , txtACVClaim
End Sub

Private Sub txtACVLessExcessLimits_GotFocus()
    goUtil.utSelText txtACVLessExcessLimits
End Sub

Private Sub txtACVLessExcessLimits_LostFocus()
    goUtil.utValidate , txtACVLessExcessLimits
End Sub

Private Sub txtAppliedDeductible_GotFocus()
    goUtil.utSelText txtAppliedDeductible
End Sub

Private Sub txtAppliedDeductible_LostFocus()
    goUtil.utValidate , txtAppliedDeductible
End Sub

Private Sub txtDescription_Change()
    On Error GoTo EH
    Dim lPos As Long
    
    cmdSave.Enabled = True
    
    If InStr(1, txtDescription.Text, vbCrLf, vbBinaryCompare) > 0 Then
        lPos = txtDescription.SelStart
        txtDescription.Text = Replace(txtDescription.Text, vbCrLf, vbNullString)
        txtDescription.SelStart = lPos
    End If
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub txtDescription_Change"
End Sub

Private Sub txtDescription_GotFocus()
    goUtil.utSelText txtDescription
    Set CurrentTextBox = txtDescription
End Sub

Private Sub txtDescription_LostFocus()
    goUtil.utValidate , txtDescription
End Sub

Private Sub txtExcessLimits_Change()
    SumExcessLimits
End Sub

Private Sub txtExcessLimits_GotFocus()
    goUtil.utSelText txtExcessLimits
End Sub

Private Sub txtExcessLimits_LostFocus()
    goUtil.utValidate , txtExcessLimits
End Sub

Private Sub txtMiscDescription_Change()
    On Error GoTo EH
    Dim lPos As Long
    
    cmdSave.Enabled = True
    
    If InStr(1, txtMiscDescription.Text, vbCrLf, vbBinaryCompare) > 0 Then
        lPos = txtMiscDescription.SelStart
        txtMiscDescription.Text = Replace(txtMiscDescription.Text, vbCrLf, vbNullString)
        txtMiscDescription.SelStart = lPos
    End If
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub txtMiscDescription_Change"
End Sub

Private Sub txtMiscDescription_GotFocus()
    goUtil.utSelText txtMiscDescription
    If Not txtMiscDescription.Locked Then
        Set CurrentTextBox = txtMiscDescription
    End If
End Sub

Private Sub txtMiscDescription_LostFocus()
    goUtil.utValidate , txtMiscDescription
End Sub

Private Sub txtMiscellaneous_Change()
    On Error GoTo EH
    Dim cMiscellaneous As Currency
    SumACVLessExcessLimits
    cmdSave.Enabled = True
    
    If IsNumeric(txtMiscellaneous.Text) Then
        cMiscellaneous = txtMiscellaneous.Text
        If cMiscellaneous > 0 Then
            lblMiscDesc.Enabled = True
            txtMiscDescription.Enabled = True
            txtMiscDescription.BackColor = BG_COLOR_WHITE
        Else
            lblMiscDesc.Enabled = False
            txtMiscDescription.Enabled = False
            txtMiscDescription.BackColor = BG_COLOR_DRKGRAY
            txtMiscDescription.Text = vbNullString
        End If
    End If
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub txtMiscellaneous_Change"
End Sub

Private Sub txtMiscellaneous_GotFocus()
    goUtil.utSelText txtMiscellaneous
End Sub

Private Sub txtMiscellaneous_LostFocus()
    goUtil.utValidate , txtMiscellaneous
End Sub

Private Sub txtNonRecoverableDep_Change()
    SumACVClaim
    cmdSave.Enabled = True
End Sub

Private Sub txtNonRecoverableDep_GotFocus()
    goUtil.utSelText txtNonRecoverableDep
End Sub

Private Sub txtNonRecoverableDep_LostFocus()
    goUtil.utValidate , txtNonRecoverableDep
End Sub

Private Sub txtPPayAmountPaid_Change()
    cmdSave.Enabled = True
End Sub

Private Sub txtPPayAmountPaid_GotFocus()
    goUtil.utSelText txtPPayAmountPaid
End Sub

Private Sub txtPPayAmountPaid_LostFocus()
    goUtil.utValidate , txtPPayAmountPaid
End Sub

Private Sub txtPPayCheckNumber_Change()
    cmdSave.Enabled = True
End Sub

Private Sub txtPPayCheckNumber_GotFocus()
    goUtil.utSelText txtPPayCheckNumber
End Sub

Private Sub txtPPayCheckNumber_LostFocus()
    goUtil.utValidate , txtPPayCheckNumber
End Sub

Private Sub txtPPayDatePaid_Change()
    cmdSave.Enabled = True
End Sub

Private Sub txtPPayDatePaid_GotFocus()
    goUtil.utSelText txtPPayDatePaid
End Sub

Private Sub txtPPayDatePaid_LostFocus()
    goUtil.utValidate , txtPPayDatePaid
End Sub

Private Sub txtRecoverableDep_Change()
    SumACVClaim
    cmdSave.Enabled = True
End Sub

Private Sub txtRecoverableDep_GotFocus()
    goUtil.utSelText txtRecoverableDep
End Sub

Private Sub txtRecoverableDep_LostFocus()
    goUtil.utValidate , txtRecoverableDep
End Sub

Private Sub txtReplacementCost_Change()
    SumACVClaim
    cmdSave.Enabled = True
End Sub

Private Sub txtReplacementCost_GotFocus()
     goUtil.utSelText txtReplacementCost
End Sub

Private Sub txtReplacementCost_LostFocus()
    goUtil.utValidate , txtReplacementCost
End Sub

Private Sub txtSpecialLimits_Change()
    SumExcessLimits
    cmdSave.Enabled = True
End Sub

Private Sub txtSpecialLimits_GotFocus()
    goUtil.utSelText txtSpecialLimits
End Sub

Private Sub txtSpecialLimits_LostFocus()
    goUtil.utValidate , txtSpecialLimits
End Sub
