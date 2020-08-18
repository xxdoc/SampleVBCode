VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFeeSchedule 
   Caption         =   "Fee Schedule"
   ClientHeight    =   9045
   ClientLeft      =   60
   ClientTop       =   11130
   ClientWidth     =   10695
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFeeSchedule.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9045
   ScaleWidth      =   10695
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPrintList 
      Caption         =   "Sho&w Item List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8520
      MaskColor       =   &H00000000&
      Picture         =   "frmFeeSchedule.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   43
      ToolTipText     =   "Exit"
      Top             =   8040
      Width           =   975
   End
   Begin VB.Frame framFeeSchedLevels 
      Caption         =   "Fee Schedule Levels"
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      Begin MSComctlLib.ListView lstvFeeSchedLevels 
         Height          =   3255
         Left            =   120
         TabIndex        =   1
         Tag             =   "Enable"
         Top             =   240
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   5741
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         OLEDragMode     =   1
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "imgActLogStatus"
         SmallIcons      =   "imgActLogStatus"
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
   Begin VB.Frame framFeeTypes 
      Caption         =   "Fee Types"
      Height          =   3615
      Left            =   5400
      TabIndex        =   2
      Top             =   120
      Width           =   5175
      Begin MSComctlLib.ImageList imgFeeTypesStatus 
         Left            =   3840
         Top             =   480
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483633
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   16777215
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFeeSchedule.frx":0884
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFeeSchedule.frx":09DE
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFeeSchedule.frx":0DCA
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lstvFeeTypes 
         Height          =   3255
         Left            =   120
         TabIndex        =   3
         Tag             =   "Enable"
         Top             =   240
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   5741
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         OLEDragMode     =   1
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "imgFeeTypesStatus"
         SmallIcons      =   "imgFeeTypesStatus"
         ColHdrIcons     =   "imgFeeTypesStatus"
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
   Begin VB.Frame framOther 
      Caption         =   "Other"
      Height          =   1095
      Left            =   120
      TabIndex        =   6
      Top             =   3720
      Width           =   10455
      Begin VB.CheckBox chkHideDeleted 
         Alignment       =   1  'Right Justify
         Caption         =   "Check to hide deleted records on all screens"
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
         TabIndex        =   4
         Top             =   240
         Value           =   1  'Checked
         Width           =   4455
      End
      Begin VB.TextBox txtTaxPercent 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E9E9E9&
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
         Left            =   7440
         Locked          =   -1  'True
         MaxLength       =   14
         TabIndex        =   10
         Top             =   480
         Width           =   2895
      End
      Begin VB.TextBox txtFeeServiceHourlyRate 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E9E9E9&
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
         Left            =   7440
         Locked          =   -1  'True
         MaxLength       =   14
         TabIndex        =   8
         Top             =   240
         Width           =   2895
      End
      Begin VB.CheckBox chkHideUploadFlags 
         Alignment       =   1  'Right Justify
         Caption         =   "Check to hide upload flags on all screens"
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
         TabIndex        =   5
         Top             =   460
         Value           =   1  'Checked
         Width           =   4455
      End
      Begin VB.Label lblTaxPercent 
         Alignment       =   1  'Right Justify
         Caption         =   "TAX Percent:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4800
         TabIndex        =   9
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label lblFeeServiceHourlyRate 
         Alignment       =   1  'Right Justify
         Caption         =   "Service Hourly Rate FEE:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4800
         TabIndex        =   7
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   855
      Left            =   9600
      MaskColor       =   &H00000000&
      Picture         =   "frmFeeSchedule.frx":121C
      Style           =   1  'Graphical
      TabIndex        =   44
      ToolTipText     =   "Exit"
      Top             =   8040
      Width           =   975
   End
   Begin VB.Frame framInitialOptions 
      Caption         =   "Initial Options"
      Height          =   1575
      Left            =   120
      TabIndex        =   11
      Top             =   4800
      Width           =   10455
      Begin VB.TextBox txtInitialOptions 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
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
         Index           =   5
         Left            =   7440
         MaxLength       =   100
         TabIndex        =   26
         Top             =   1200
         Width           =   2895
      End
      Begin VB.TextBox txtInitialOptions 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
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
         Index           =   4
         Left            =   7440
         MaxLength       =   100
         TabIndex        =   24
         Top             =   960
         Width           =   2895
      End
      Begin VB.TextBox txtInitialOptions 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
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
         Index           =   3
         Left            =   7440
         MaxLength       =   100
         TabIndex        =   22
         Top             =   720
         Width           =   2895
      End
      Begin VB.TextBox txtInitialOptions 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
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
         Index           =   2
         Left            =   7440
         MaxLength       =   100
         TabIndex        =   20
         Top             =   480
         Width           =   2895
      End
      Begin VB.TextBox txtInitialOptions 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
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
         Index           =   1
         Left            =   7440
         MaxLength       =   100
         TabIndex        =   18
         Top             =   240
         Width           =   2895
      End
      Begin VB.CheckBox chkInitialOptions 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   16
         Top             =   1200
         Width           =   4335
      End
      Begin VB.CheckBox chkInitialOptions 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   15
         Top             =   960
         Width           =   4335
      End
      Begin VB.CheckBox chkInitialOptions 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   14
         Top             =   720
         Width           =   4335
      End
      Begin VB.CheckBox chkInitialOptions 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   13
         Top             =   480
         Width           =   4335
      End
      Begin VB.CheckBox chkInitialOptions 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   4335
      End
      Begin VB.Label lblInitialOptions 
         Alignment       =   1  'Right Justify
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
         Height          =   240
         Index           =   5
         Left            =   4800
         TabIndex        =   25
         Top             =   1200
         Width           =   2535
      End
      Begin VB.Label lblInitialOptions 
         Alignment       =   1  'Right Justify
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
         Height          =   240
         Index           =   4
         Left            =   4800
         TabIndex        =   23
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label lblInitialOptions 
         Alignment       =   1  'Right Justify
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
         Height          =   240
         Index           =   3
         Left            =   4800
         TabIndex        =   21
         Top             =   720
         Width           =   2535
      End
      Begin VB.Label lblInitialOptions 
         Alignment       =   1  'Right Justify
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
         Height          =   240
         Index           =   2
         Left            =   4800
         TabIndex        =   19
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label lblInitialOptions 
         Alignment       =   1  'Right Justify
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
         Height          =   240
         Index           =   1
         Left            =   4800
         TabIndex        =   17
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Frame framOptions 
      Caption         =   "Options"
      Height          =   1575
      Left            =   120
      TabIndex        =   27
      Top             =   6360
      Width           =   10455
      Begin VB.TextBox txtOptions 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
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
         Index           =   5
         Left            =   7440
         MaxLength       =   100
         TabIndex        =   42
         Top             =   1200
         Width           =   2895
      End
      Begin VB.TextBox txtOptions 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
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
         Index           =   4
         Left            =   7440
         MaxLength       =   100
         TabIndex        =   40
         Top             =   960
         Width           =   2895
      End
      Begin VB.TextBox txtOptions 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
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
         Index           =   3
         Left            =   7440
         MaxLength       =   100
         TabIndex        =   38
         Top             =   720
         Width           =   2895
      End
      Begin VB.TextBox txtOptions 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
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
         Index           =   2
         Left            =   7440
         MaxLength       =   100
         TabIndex        =   36
         Top             =   480
         Width           =   2895
      End
      Begin VB.TextBox txtOptions 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
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
         Index           =   1
         Left            =   7440
         MaxLength       =   100
         TabIndex        =   34
         Top             =   240
         Width           =   2895
      End
      Begin VB.CheckBox chkOptions 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   32
         Top             =   1200
         Width           =   4335
      End
      Begin VB.CheckBox chkOptions 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   31
         Top             =   960
         Width           =   4335
      End
      Begin VB.CheckBox chkOptions 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   30
         Top             =   720
         Width           =   4335
      End
      Begin VB.CheckBox chkOptions 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   29
         Top             =   480
         Width           =   4335
      End
      Begin VB.CheckBox chkOptions 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   28
         Top             =   240
         Width           =   4335
      End
      Begin VB.Label lblOptions 
         Alignment       =   1  'Right Justify
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
         Height          =   240
         Index           =   5
         Left            =   4800
         TabIndex        =   41
         Top             =   1200
         Width           =   2535
      End
      Begin VB.Label lblOptions 
         Alignment       =   1  'Right Justify
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
         Height          =   240
         Index           =   4
         Left            =   4800
         TabIndex        =   39
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label lblOptions 
         Alignment       =   1  'Right Justify
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
         Height          =   240
         Index           =   3
         Left            =   4800
         TabIndex        =   37
         Top             =   720
         Width           =   2535
      End
      Begin VB.Label lblOptions 
         Alignment       =   1  'Right Justify
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
         Height          =   240
         Index           =   2
         Left            =   4800
         TabIndex        =   35
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label lblOptions 
         Alignment       =   1  'Right Justify
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
         Height          =   240
         Index           =   1
         Left            =   4800
         TabIndex        =   33
         Top             =   240
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frmFeeSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbFormLoading As Boolean
Private mbLoadingMe As Boolean
Private mbUnloadMe As Boolean
Private moGUI As V2ECKeyBoard.clsCarGUI
Private mcolInitialOptions As Collection
Private mcolOptions As Collection
Private msFeeScheduleID As String
Private msClientCompanyID As String

Private Const BG_COLOR_GRAY = &HE0E0E0
Private Const BG_COLOR_DRKGRAY = &H8000000F
Private Const BG_COLOR_WHITE = &H80000005
Private Const BG_COLOR_YELLOW = &H80000018


Private Property Get msClassName() As String
    msClassName = Me.Name
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

Private Sub chkHideDeleted_Click()
    On Error GoTo EH
    Dim bHideDeleted As Boolean
    
    If chkHideDeleted.Value = vbChecked Then
        bHideDeleted = True
    Else
        bHideDeleted = False
    End If
    
    SaveSetting App.EXEName, "GENERAL", "HIDE_DELETED", bHideDeleted
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub chkHideDeleted_Click"
End Sub

Private Sub chkHideUploadFlags_Click()
    On Error GoTo EH
    Dim bHideUploadFlags As Boolean
    
    If chkHideUploadFlags.Value = vbChecked Then
        bHideUploadFlags = True
    Else
        bHideUploadFlags = False
    End If
    
    SaveSetting App.EXEName, "GENERAL", "HIDE_UPLOAD_FLAGS", bHideUploadFlags
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub chkHideUploadFlags_Click"
End Sub

Private Sub chkInitialOptions_Click(Index As Integer)
    On Error GoTo EH
    Dim sctlName As String
    Dim sOptKey As String
    Dim sValue As String
    If mbFormLoading Then
        Exit Sub
    End If
    
    sctlName = chkInitialOptions(Index).Name
    sOptKey = """" & sctlName & "_" & CStr(Index) & """"
    
    If chkInitialOptions(Index).Value = vbChecked Then
        sValue = "1"
    Else
        sValue = "0"
    End If
    
    SaveOptSettings mcolInitialOptions, sOptKey, sValue
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub chkInitialOptions_Click"
End Sub

Private Sub chkOptions_Click(Index As Integer)
    On Error GoTo EH
    Dim sctlName As String
    Dim sOptKey As String
    Dim sValue As String
    If mbFormLoading Then
        Exit Sub
    End If
    
    sctlName = chkOptions(Index).Name
    sOptKey = """" & sctlName & "_" & CStr(Index) & """"
    
    If chkOptions(Index).Value = vbChecked Then
        sValue = "1"
    Else
        sValue = "0"
    End If
    
    SaveOptSettings mcolOptions, sOptKey, sValue
    
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub chkOptions_Click"
End Sub

Private Sub cmdExit_Click()
    On Error GoTo EH
    Dim sMess As String
    mbUnloadMe = True
    cmdExit.Enabled = False
    CLEANUP
    goUtil.utFormWinRegPos goUtil.gsMainAppEXEName, Me, True, , , , True
    Unload Me
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdExit_Click"
End Sub

Private Sub cmdPrintList_Click()
    On Error GoTo EH
    goUtil.utPrintListView App.EXEName, lstvFeeSchedLevels, "Fee Schedule Levels"
    goUtil.utPrintListView App.EXEName, lstvFeeTypes, "Fee Types"
     Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdPrintList_Click"
End Sub

Private Sub Form_Load()
    On Error GoTo EH
    mbFormLoading = True
    
    LoadHeaderlstvFeeSchedLevels
    LoadHeaderlstvFeeTypes
    
    Screen.MousePointer = vbHourglass
    LoadMe
    Screen.MousePointer = vbDefault
    
    mbFormLoading = False
    mbUnloadMe = False
    Exit Sub
EH:
    mbFormLoading = False
   goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_Load"
End Sub

Public Function LoadMe() As Boolean
    On Error GoTo EH
    Dim RS As ADODB.Recordset
    Dim lMaxCheckBox As Long
    Dim lMaxTextBox As Long
    Dim sData As String
    Dim sCaption As String
    Dim bHideDeleted As Boolean
    Dim bHideUploadFlags As Boolean
    
    mbLoadingMe = True

    PopulatelvwFeeSchedLevels
    
    PopulatelvwFeeTypes
    
    'populate the Rest of the Form
    Set RS = moGUI.adoFeeSchedule

    If RS.RecordCount = 0 Then
        Exit Function
    End If
    
    RS.MoveFirst
    
    
    'Set the Form Caption using Sched name and Desc
    sCaption = "Fee Schedule (" & goUtil.IsNullIsVbNullString(RS.Fields("ScheduleName")) & " - "
    sCaption = sCaption & goUtil.IsNullIsVbNullString(RS.Fields("Description")) & ") "
    Me.Caption = sCaption
    
    'set the FeeSchedId and ClientComp ID
    msFeeScheduleID = goUtil.IsNullIsVbNullString(RS.Fields("FeeScheduleID"))
    msClientCompanyID = goUtil.IsNullIsVbNullString(RS.Fields("ClientCompanyID"))
    
    
    txtFeeServiceHourlyRate.Text = Format(goUtil.IsNullIsVbNullString(RS.Fields("FeeServiceHourlyRate")), "#,###,###,##0.00")
    txtTaxPercent.Text = Format(goUtil.IsNullIsVbNullString(RS.Fields("TaxPercent")), "0.0000")
    
    'Populate Initial Options Collection
    lMaxCheckBox = 5
    lMaxTextBox = 5
    sData = goUtil.IsNullIsVbNullString(RS.Fields("InitialOptions"))
    PopulateOpt mcolInitialOptions, sData, lMaxCheckBox, lMaxTextBox, chkInitialOptions, txtInitialOptions, lblInitialOptions
    
    'Populate Options Collection
    lMaxCheckBox = 5
    lMaxTextBox = 5
    sData = goUtil.IsNullIsVbNullString(RS.Fields("Options"))
    PopulateOpt mcolOptions, sData, lMaxCheckBox, lMaxTextBox, chkOptions, txtOptions, lblOptions
    
    'Set the Hide Deleted value
    bHideDeleted = CBool(GetSetting(App.EXEName, "GENERAL", "HIDE_DELETED", True))
    If bHideDeleted Then
        chkHideDeleted.Value = vbChecked
    Else
        chkHideDeleted.Value = vbUnchecked
    End If
    
    'Set the hide upload flags value
    bHideUploadFlags = CBool(GetSetting(App.EXEName, "GENERAL", "HIDE_UPLOAD_FLAGS", True))
    If bHideUploadFlags Then
        chkHideUploadFlags.Value = vbChecked
    Else
        chkHideUploadFlags.Value = vbUnchecked
    End If
    
    LoadMe = True
    mbLoadingMe = False
    'Clean Up
    Set RS = Nothing
    Exit Function
EH:
    mbLoadingMe = False
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function LoadMe"
End Function

Private Function PopulateOpt(pcolOpt As Collection, _
                            psData As String, _
                            plMaxCheckBox As Long, _
                            plMaxTextBox As Long, _
                            pochkOpt As Object, _
                            potxtOpt As Object, _
                            polblOpt As Object) As Boolean
    On Error GoTo EH
    Dim saryOpt() As String
    Dim saryOptItem() As String
    Dim saryValue() As String
    Dim lCount As Long
    Dim lMaxCheckBox As Long
    Dim lRemainCheckBox As Long
    Dim lMaxTextBox As Long
    Dim lRemainTextBox As Long
    Dim lIndex As Long
    Dim sctlName As String
    Dim MyFSOptions As FSOptions
    Dim sData As String
    Dim bSkipThisItem As Boolean
    Dim ochkOpt As Object
    Dim otxtOpt As Object
    Dim olblOpt As Object
    Dim sDefaultValue As String
    Dim sAppName As String
    Dim sSection As String
    Dim sKey As String
    
    
    'Init local from param
    sData = psData
    lMaxCheckBox = plMaxCheckBox
    lRemainCheckBox = lMaxCheckBox
    lMaxTextBox = plMaxTextBox
    lRemainTextBox = lMaxTextBox

    Set ochkOpt = pochkOpt
    Set otxtOpt = potxtOpt
    Set olblOpt = polblOpt
    
    If Trim(sData) <> vbNullString Then
        saryOpt() = Split(sData, "@", , vbBinaryCompare)
        For lCount = LBound(saryOpt, 1) To UBound(saryOpt, 1)
            sData = saryOpt(lCount)
            If sData <> vbNullString Then
                bSkipThisItem = False
                saryOptItem() = Split(sData, "^", , vbBinaryCompare)
                'Populate the Udt
                With MyFSOptions
                    saryValue() = Split(saryOptItem(0), "=", , vbBinaryCompare)
                    .optCAPTION = saryValue(1)
                    saryValue() = Split(saryOptItem(2), "=", , vbBinaryCompare)
                    .optDATATYPE = saryValue(1)
                    saryValue() = Split(saryOptItem(4), "=", , vbBinaryCompare)
                    .optDEFAULTVALUE = saryValue(1)
                    saryValue() = Split(saryOptItem(1), "=", , vbBinaryCompare)
                    .optTYPE = saryValue(1)
                    saryValue() = Split(saryOptItem(3), "=", , vbBinaryCompare)
                    .optVBREGSTR = saryValue(1)
                End With
                sData = MyFSOptions.optTYPE
                'Figure out what type this control is
                If StrComp(sData, "checkbox", vbTextCompare) = 0 Then
                    lRemainCheckBox = lRemainCheckBox - 1
                    lIndex = lMaxCheckBox - lRemainCheckBox
                    'if index is =0 then went past the max
                    If lIndex >= 1 And lIndex <= lMaxCheckBox Then
                        sAppName = App.EXEName
                        sSection = msFeeScheduleID & "_" & msClientCompanyID
                        sKey = MyFSOptions.optVBREGSTR
                        sDefaultValue = MyFSOptions.optDEFAULTVALUE
                        sDefaultValue = GetSetting(sAppName, "FeeSchedule\" & sSection, sKey, sDefaultValue)
                        
                        ochkOpt(lIndex).Caption = MyFSOptions.optCAPTION
                        ochkOpt(lIndex).Enabled = True
                        If CBool(sDefaultValue) Then
                            ochkOpt(lIndex).Value = vbChecked
                        Else
                            ochkOpt(lIndex).Value = vbUnchecked
                        End If
                        sctlName = ochkOpt(lIndex).Name
                        
                        'Now Save the Setting
                        'This will SetUp the Default Value as well as allow for User Changes
                        'to the default value
                        SaveSetting sAppName, "FeeSchedule\" & sSection, sKey, sDefaultValue
                        
                    Else
                        bSkipThisItem = True
                    End If
                ElseIf StrComp(sData, "textbox", vbTextCompare) = 0 Then
                    lRemainTextBox = lRemainTextBox - 1
                    lIndex = lMaxTextBox - lRemainTextBox
                    'if index is =0 then went past the max
                    If lIndex >= 1 And lIndex <= lMaxTextBox Then
                        sAppName = App.EXEName
                        sSection = msFeeScheduleID & "_" & msClientCompanyID
                        sKey = MyFSOptions.optVBREGSTR
                        sDefaultValue = MyFSOptions.optDEFAULTVALUE
                        sDefaultValue = GetSetting(sAppName, "FeeSchedule\" & sSection, sKey, sDefaultValue)
                        
                        olblOpt(lIndex).Caption = MyFSOptions.optCAPTION
                        olblOpt(lIndex).Enabled = True
                        otxtOpt(lIndex).Enabled = True
                        otxtOpt(lIndex).BackColor = BG_COLOR_WHITE
                        otxtOpt(lIndex).Text = sDefaultValue
                        sctlName = otxtOpt(lIndex).Name
                        
                        'Now Save the Setting
                        'This will SetUp the Default Value as well as allow for User Changes
                        'to the default value
                        SaveSetting sAppName, "FeeSchedule\" & sSection, sKey, sDefaultValue
                    Else
                        bSkipThisItem = True
                    End If
                Else
                    bSkipThisItem = True
                End If
                
                If pcolOpt Is Nothing Then
                    Set pcolOpt = New Collection
                End If
                If Not bSkipThisItem Then
                    pcolOpt.Add MyFSOptions, """" & sctlName & "_" & CStr(lIndex) & """"
                End If
            End If
        Next
    End If
    
    'cleanup
    Set ochkOpt = Nothing
    Set otxtOpt = Nothing
    Set olblOpt = Nothing
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Function PopulateOpt"
End Function

Private Sub PopulatelvwFeeSchedLevels()
    On Error GoTo EH
    Dim iMyIcon As Long
    Dim sFlagText As String
    Dim sTemp As String
    Dim itmX As ListItem
    Dim oListView As MSComctlLib.ListView
    Dim RS As ADODB.Recordset
    
    'Clear the List view
    Set oListView = lstvFeeSchedLevels

    oListView.Visible = False
    oListView.ListItems.Clear

    Set RS = moGUI.adoFeeScheduleLevels

    If RS.RecordCount > 0 Then
        RS.MoveFirst
        Do Until RS.EOF
            'LevelNum
             Set itmX = oListView.ListItems.Add(, """" & goUtil.IsNullIsVbNullString(RS.Fields("FeeScheduleLevelsID")) & """", goUtil.IsNullIsVbNullString(RS.Fields("LevelNum")))
            'LevelNumSort 'hidden
            itmX.SubItems(GuiFeeSchedLevelsListView.LevelNumSort - 1) = goUtil.utNumInTextSortFormat(goUtil.IsNullIsVbNullString(RS.Fields("LevelNum")))
            'LevelMax
            itmX.SubItems(GuiFeeSchedLevelsListView.LevelMax - 1) = Format(goUtil.IsNullIsVbNullString(RS.Fields("LevelMax")), "#,###,###,##0.00")
            'LevelMaxSort
            itmX.SubItems(GuiFeeSchedLevelsListView.LevelMaxSort - 1) = goUtil.utNumInTextSortFormat(Format(goUtil.IsNullIsVbNullString(RS.Fields("LevelMax")), "#,###,###,##0.00"))
            'LevelPctApp
            itmX.SubItems(GuiFeeSchedLevelsListView.LevelPctApp - 1) = Format(goUtil.IsNullIsVbNullString(RS.Fields("LevelPctApp")), "#,###,###,##0.00")
            'LevelPctAppSort
            itmX.SubItems(GuiFeeSchedLevelsListView.LevelPctAppSort - 1) = goUtil.utNumInTextSortFormat(Format(goUtil.IsNullIsVbNullString(RS.Fields("LevelPctApp")), "0.0000"))
            'LevelMin
            itmX.SubItems(GuiFeeSchedLevelsListView.LevelMin - 1) = Format(goUtil.IsNullIsVbNullString(RS.Fields("LevelMin")), "#,###,###,##0.00")
            'LevelMinSort
            itmX.SubItems(GuiFeeSchedLevelsListView.LevelMinSort - 1) = goUtil.utNumInTextSortFormat(Format(goUtil.IsNullIsVbNullString(RS.Fields("LevelMin")), "#,###,###,##0.00"))
            'DateLastUpdated
            If Not IsNull(RS.Fields("DateLastUpdated").Value) Then
                If IsDate(RS.Fields("DateLastUpdated").Value) Then
                    itmX.SubItems(GuiFeeSchedLevelsListView.DateLastUpdated - 1) = Format(RS.Fields("DateLastUpdated").Value, "MM/DD/YYYY HH:MM:SS")
                    itmX.SubItems(GuiFeeSchedLevelsListView.DateLastUpdatedSort - 1) = Format(RS.Fields("DateLastUpdated").Value, "YYYY/MM/DD HH:MM:SS")
                Else
                    itmX.SubItems(GuiFeeSchedLevelsListView.DateLastUpdated - 1) = vbNullString
                    itmX.SubItems(GuiFeeSchedLevelsListView.DateLastUpdatedSort - 1) = vbNullString
                End If
            Else
                itmX.SubItems(GuiFeeSchedLevelsListView.DateLastUpdated - 1) = vbNullString
                itmX.SubItems(GuiFeeSchedLevelsListView.DateLastUpdatedSort - 1) = vbNullString
            End If
            'UpdateByUserID ' Hidden
            itmX.SubItems(GuiFeeSchedLevelsListView.UpdateByUserID - 1) = goUtil.IsNullIsVbNullString(RS.Fields("UpdateByUserID"))
            'FeeScheduleLevelsID ' Hidden
            itmX.SubItems(GuiFeeSchedLevelsListView.FeeScheduleLevelsID - 1) = goUtil.IsNullIsVbNullString(RS.Fields("FeeScheduleLevelsID"))
            'FeeScheduleID ' Hidden
            itmX.SubItems(GuiFeeSchedLevelsListView.FeeScheduleID - 1) = goUtil.IsNullIsVbNullString(RS.Fields("FeeScheduleID"))
        
            itmX.Selected = False

            RS.MoveNext
        Loop
    End If
    oListView.Visible = True
    'Cleanup
    Set itmX = Nothing
    Set RS = Nothing
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub PopulatelvwFeeSchedLevels"
    oListView.Visible = True
End Sub

Private Sub PopulatelvwFeeTypes()
    On Error GoTo EH
    Dim iMyIcon As Long
    Dim sFlagText As String
    Dim sTemp As String
    Dim itmX As ListItem
    Dim oListView As MSComctlLib.ListView
    Dim RS As ADODB.Recordset
    
    'Clear the List view
    Set oListView = lstvFeeTypes

    oListView.Visible = False
    oListView.ListItems.Clear

    Set RS = moGUI.adoFeeScheduleFeeTypes

    If RS.RecordCount > 0 Then
        RS.MoveFirst
        Do Until RS.EOF
            'TypeNum = 1
             Set itmX = oListView.ListItems.Add(, """" & goUtil.IsNullIsVbNullString(RS.Fields("FeeScheduleFeeTypesID")) & """", goUtil.IsNullIsVbNullString(RS.Fields("TypeNum")))
            'TypeNumSort
            itmX.SubItems(GuiFeeTypesListView.TypeNumSort - 1) = goUtil.utNumInTextSortFormat(goUtil.IsNullIsVbNullString(RS.Fields("TypeNum")))
            'ftName
            itmX.SubItems(GuiFeeTypesListView.ftName - 1) = goUtil.IsNullIsVbNullString(RS.Fields("Name"))
            'Description
            itmX.SubItems(GuiFeeTypesListView.Description - 1) = goUtil.IsNullIsVbNullString(RS.Fields("Description"))
            'FeeAmount
            itmX.SubItems(GuiFeeTypesListView.FeeAmount - 1) = Format(goUtil.IsNullIsVbNullString(RS.Fields("FeeAmount")), "#,###,###,##0.00")
            'FeeAmountSort
            itmX.SubItems(GuiFeeTypesListView.FeeAmountSort - 1) = goUtil.utNumInTextSortFormat(Format(goUtil.IsNullIsVbNullString(RS.Fields("FeeAmount")), "#,###,###,##0.00"))
            'IsExpense
            If CBool(RS.Fields("IsExpense")) Then
                iMyIcon = GuiFeeTypesStatusList.IsChecked
            Else
                iMyIcon = Empty
            End If
            sFlagText = goUtil.GetFlagText(RS.Fields("IsExpense"))
            itmX.SubItems(GuiFeeTypesListView.IsExpense - 1) = sFlagText
            itmX.ListSubItems(GuiFeeTypesListView.IsExpense - 1).ReportIcon = iMyIcon
            'MaxNumberOfItems
            itmX.SubItems(GuiFeeTypesListView.MaxNumberOfItems - 1) = goUtil.IsNullIsVbNullString(RS.Fields("MaxNumberOfItems"))
            'MaxNumberOfItemsSort
            itmX.SubItems(GuiFeeTypesListView.MaxNumberOfItemsSort - 1) = goUtil.utNumInTextSortFormat(goUtil.IsNullIsVbNullString(RS.Fields("MaxNumberOfItems")))
            'MaxFeeAmount
            itmX.SubItems(GuiFeeTypesListView.MaxFeeAmount - 1) = Format(goUtil.IsNullIsVbNullString(RS.Fields("MaxFeeAmount")), "#,###,###,##0.00")
            'MaxFeeAmountSort
            itmX.SubItems(GuiFeeTypesListView.MaxFeeAmountSort - 1) = goUtil.utNumInTextSortFormat(Format(goUtil.IsNullIsVbNullString(RS.Fields("MaxFeeAmount")), "#,###,###,##0.00"))
            'UseFormula
            If CBool(RS.Fields("UseFormula")) Then
                iMyIcon = GuiFeeTypesStatusList.IsChecked
            Else
                iMyIcon = Empty
            End If
            sFlagText = goUtil.GetFlagText(RS.Fields("UseFormula"))
            itmX.SubItems(GuiFeeTypesListView.UseFormula - 1) = sFlagText
            itmX.ListSubItems(GuiFeeTypesListView.UseFormula - 1).ReportIcon = iMyIcon
            'VBFormula
            itmX.SubItems(GuiFeeTypesListView.VBFormula - 1) = goUtil.IsNullIsVbNullString(RS.Fields("VBFormula"))
            'DateLastUpdated
            If Not IsNull(RS.Fields("DateLastUpdated").Value) Then
                If IsDate(RS.Fields("DateLastUpdated").Value) Then
                    itmX.SubItems(GuiFeeTypesListView.DateLastUpdated - 1) = Format(RS.Fields("DateLastUpdated").Value, "MM/DD/YYYY HH:MM:SS")
                    itmX.SubItems(GuiFeeTypesListView.DateLastUpdatedSort - 1) = Format(RS.Fields("DateLastUpdated").Value, "YYYY/MM/DD HH:MM:SS")
                Else
                    itmX.SubItems(GuiFeeTypesListView.DateLastUpdated - 1) = vbNullString
                    itmX.SubItems(GuiFeeTypesListView.DateLastUpdatedSort - 1) = vbNullString
                End If
            Else
                itmX.SubItems(GuiFeeTypesListView.DateLastUpdated - 1) = vbNullString
                itmX.SubItems(GuiFeeTypesListView.DateLastUpdatedSort - 1) = vbNullString
            End If
            'UpdateByUserID ' Hidden
            itmX.SubItems(GuiFeeTypesListView.UpdateByUserID - 1) = goUtil.IsNullIsVbNullString(RS.Fields("UpdateByUserID"))
            'FeeScheduleFeeTypesID ' Hidden
            itmX.SubItems(GuiFeeTypesListView.FeeScheduleFeeTypesID - 1) = goUtil.IsNullIsVbNullString(RS.Fields("FeeScheduleFeeTypesID"))
            'FeeScheduleID ' Hidden
            itmX.SubItems(GuiFeeTypesListView.FeeScheduleID - 1) = goUtil.IsNullIsVbNullString(RS.Fields("FeeScheduleID"))
             
            itmX.Selected = False

            RS.MoveNext
        Loop
    End If
    oListView.Visible = True
    'Cleanup
    Set itmX = Nothing
    Set RS = Nothing
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub PopulatelvwFeeTypes"
    oListView.Visible = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo EH
    
    Select Case UnloadMode
        Case vbFormControlMenu
            CLEANUP
            goUtil.utFormWinRegPos goUtil.gsMainAppEXEName, Me, True, , , , True
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
    Dim lHeightDiff As Long
    Dim lWidthDiff As Long
    
    If Me.WindowState <> vbMinimized Then
        Me.Visible = False
    End If
    
    'widths
    framFeeTypes.Width = Me.Width - 5610
    lstvFeeTypes.Width = Me.Width - 5850
    framOther.Width = Me.Width - 330
    framInitialOptions.Width = Me.Width - 330
    framOptions.Width = Me.Width - 330
    'lefts
    lblFeeServiceHourlyRate.left = Me.Width - 5985
    lblTaxPercent.left = Me.Width - 5985
    txtFeeServiceHourlyRate.left = Me.Width - 3345
    txtTaxPercent.left = Me.Width - 3345
    lblInitialOptions(1).left = Me.Width - 5985
    lblInitialOptions(2).left = Me.Width - 5985
    lblInitialOptions(3).left = Me.Width - 5985
    lblInitialOptions(4).left = Me.Width - 5985
    lblInitialOptions(5).left = Me.Width - 5985
    txtInitialOptions(1).left = Me.Width - 3345
    txtInitialOptions(2).left = Me.Width - 3345
    txtInitialOptions(3).left = Me.Width - 3345
    txtInitialOptions(4).left = Me.Width - 3345
    txtInitialOptions(5).left = Me.Width - 3345
    lblOptions(1).left = Me.Width - 5985
    lblOptions(2).left = Me.Width - 5985
    lblOptions(3).left = Me.Width - 5985
    lblOptions(4).left = Me.Width - 5985
    lblOptions(5).left = Me.Width - 5985
    txtOptions(1).left = Me.Width - 3345
    txtOptions(2).left = Me.Width - 3345
    txtOptions(3).left = Me.Width - 3345
    txtOptions(4).left = Me.Width - 3345
    txtOptions(5).left = Me.Width - 3345
    cmdExit.left = Me.Width - 1215
    cmdPrintList.left = Me.Width - 2295
    
    'heights
    framFeeSchedLevels.Height = Me.Height - 5745
    lstvFeeSchedLevels.Height = Me.Height - 6105
    framFeeTypes.Height = Me.Height - 5745
    lstvFeeTypes.Height = Me.Height - 6105
    
    'Tops
    framOther.top = Me.Height - 5640
    framInitialOptions.top = Me.Height - 4560
    framOptions.top = Me.Height - 3000
    cmdExit.top = Me.Height - 1320
    cmdPrintList.top = Me.Height - 1320
    
    If Me.WindowState <> vbMinimized Then
        Me.Visible = True
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo EH
    If Not goUtil.gfrmECTray Is Nothing Then
        goUtil.gfrmECTray.ShowMe False
    End If
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_Unload"
End Sub

Public Function CLEANUP() As Boolean
    On Error GoTo EH
    
    Set moGUI = Nothing
    Set mcolInitialOptions = Nothing
    Set mcolOptions = Nothing
    
    CLEANUP = True
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function CLEANUP"
End Function




Private Sub lstvFeeSchedLevels_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    On Error GoTo EH
    Dim sMess As String
    If lstvFeeSchedLevels.SortOrder = lvwAscending Then
        lstvFeeSchedLevels.SortOrder = lvwDescending
        sMess = "Descending"
    Else
        lstvFeeSchedLevels.SortOrder = lvwAscending
        sMess = "Ascending"
    End If
    
    'Set the Tool Tip
    lstvFeeSchedLevels.ToolTipText = "Sort By " & ColumnHeader.Text & " " & sMess
    
    'Need to see if the Column clicked was a NON text sort
    'Like Date or number.  If a non textual column was clicked
    'need to use the next column as the sort since this next column
    'will be hidden and contain a sort friendly format.
    'Sort Key is Base 0 where CoulmnHeader is not
    
    Select Case ColumnHeader.Index
        Case GuiFeeSchedLevelsListView.LevelNum
            lstvFeeSchedLevels.SortKey = ColumnHeader.Index
        Case GuiFeeSchedLevelsListView.LevelMax
            lstvFeeSchedLevels.SortKey = ColumnHeader.Index
        Case GuiFeeSchedLevelsListView.LevelPctApp
            lstvFeeSchedLevels.SortKey = ColumnHeader.Index
        Case GuiFeeSchedLevelsListView.LevelMin
            lstvFeeSchedLevels.SortKey = ColumnHeader.Index
        Case GuiFeeSchedLevelsListView.DateLastUpdated
            lstvFeeSchedLevels.SortKey = ColumnHeader.Index
        Case Else
            lstvFeeSchedLevels.SortKey = ColumnHeader.Index - 1
    End Select
    
    lstvFeeSchedLevels.Sorted = True
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lstvFeeSchedLevels_ColumnClick"
End Sub


Private Sub lstvFeeTypes_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    On Error GoTo EH
    Dim sMess As String
    If lstvFeeTypes.SortOrder = lvwAscending Then
        lstvFeeTypes.SortOrder = lvwDescending
        sMess = "Descending"
    Else
        lstvFeeTypes.SortOrder = lvwAscending
        sMess = "Ascending"
    End If
    
    'Set the Tool Tip
    lstvFeeTypes.ToolTipText = "Sort By " & ColumnHeader.Text & " " & sMess
    
    'Need to see if the Column clicked was a NON text sort
    'Like Date or number.  If a non textual column was clicked
    'need to use the next column as the sort since this next column
    'will be hidden and contain a sort friendly format.
    'Sort Key is Base 0 where CoulmnHeader is not
    
    'Debug Column Width
'    Debug.Print lstvFeeTypes.ColumnHeaders(ColumnHeader.Index).Text & " = " & lstvFeeTypes.ColumnHeaders(ColumnHeader.Index).Width
    'End
    
    Select Case ColumnHeader.Index
        Case GuiFeeTypesListView.TypeNum
            lstvFeeTypes.SortKey = ColumnHeader.Index
        Case GuiFeeTypesListView.FeeAmount
            lstvFeeTypes.SortKey = ColumnHeader.Index
        Case GuiFeeTypesListView.DateLastUpdated
            lstvFeeTypes.SortKey = ColumnHeader.Index
        Case Else
            lstvFeeTypes.SortKey = ColumnHeader.Index - 1
    End Select
    
    lstvFeeTypes.Sorted = True
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lstvFeeTypes_ColumnClick"
End Sub

Private Sub txtInitialOptions_Change(Index As Integer)
    On Error GoTo EH
    Dim sctlName As String
    Dim sOptKey As String
    Dim sValue As String
    
    If mbFormLoading Then
        Exit Sub
    End If
    
    sctlName = txtInitialOptions(Index).Name
    sOptKey = """" & sctlName & "_" & CStr(Index) & """"
    
    sValue = txtInitialOptions(Index).Text
    
    SaveOptSettings mcolInitialOptions, sOptKey, sValue
    
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub txtInitialOptions_Change"
End Sub

Private Sub txtInitialOptions_GotFocus(Index As Integer)
    On Error GoTo EH
    
    goUtil.utSelText txtInitialOptions(Index)
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub txtInitialOptions_GotFocus"
End Sub

Public Sub LoadHeaderlstvFeeSchedLevels()
    On Error GoTo EH
    Dim bGridOn As Boolean
    
    'set the columnheaders
    With lstvFeeSchedLevels
        .ColumnHeaders.Add , "LevelNum", "No."
        .ColumnHeaders.Add , "LevelNumSort", "Level No. Sort" 'Hidden
        .ColumnHeaders.Add , "LevelMax", "Up To Max"
        .ColumnHeaders.Add , "LevelMaxSort", "Up To Max Sort" 'Hidden
        .ColumnHeaders.Add , "LevelPctApp", "Pct. If Applicable"
        .ColumnHeaders.Add , "LevelPctAppSort", "Pct. If Applicable Sort" 'Hidden
        .ColumnHeaders.Add , "LevelMin", "Fee Amount"
        .ColumnHeaders.Add , "LevelMinSort", "Fee Amount Sort" 'Hidden
        .ColumnHeaders.Add , "DateLastUpdated", "Date Last Updated"
        .ColumnHeaders.Add , "DateLastUpdatedSort", "DateLastUpdatedSort" ' Hidden
        .ColumnHeaders.Add , "UpdateByUserID", "UpdateByUserID" ' Hidden
        .ColumnHeaders.Add , "FeeScheduleLevelsID", "FeeScheduleLevelsID" ' Hidden
        .ColumnHeaders.Add , "FeeScheduleID", "FeeScheduleID" ' Hidden
        
        .Sorted = False
        .SortOrder = lvwAscending
        
        'LevelNum
        .ColumnHeaders.Item(GuiFeeSchedLevelsListView.LevelNum).Width = 700
        .ColumnHeaders.Item(GuiFeeSchedLevelsListView.LevelNum).Alignment = lvwColumnLeft
        'LevelNumSort
        .ColumnHeaders.Item(GuiFeeSchedLevelsListView.LevelNumSort).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiFeeSchedLevelsListView.LevelNumSort).Alignment = lvwColumnLeft
        'LevelMax
        .ColumnHeaders.Item(GuiFeeSchedLevelsListView.LevelMax).Width = 1400
        .ColumnHeaders.Item(GuiFeeSchedLevelsListView.LevelMax).Alignment = lvwColumnRight
        'LevelMaxSort
        .ColumnHeaders.Item(GuiFeeSchedLevelsListView.LevelMaxSort).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiFeeSchedLevelsListView.LevelMaxSort).Alignment = lvwColumnRight
        'LevelPctApp
        .ColumnHeaders.Item(GuiFeeSchedLevelsListView.LevelPctApp).Width = 1000
        .ColumnHeaders.Item(GuiFeeSchedLevelsListView.LevelPctApp).Alignment = lvwColumnRight
        'LevelPctAppSort
        .ColumnHeaders.Item(GuiFeeSchedLevelsListView.LevelPctAppSort).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiFeeSchedLevelsListView.LevelPctAppSort).Alignment = lvwColumnRight
        'LevelMin
        .ColumnHeaders.Item(GuiFeeSchedLevelsListView.LevelMin).Width = 1400
        .ColumnHeaders.Item(GuiFeeSchedLevelsListView.LevelMin).Alignment = lvwColumnRight
        'LevelMinSort
        .ColumnHeaders.Item(GuiFeeSchedLevelsListView.LevelMinSort).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiFeeSchedLevelsListView.LevelMinSort).Alignment = lvwColumnRight
        'DateLastUpdated
        .ColumnHeaders.Item(GuiFeeSchedLevelsListView.DateLastUpdated).Width = 2200
        .ColumnHeaders.Item(GuiFeeSchedLevelsListView.DateLastUpdated).Alignment = lvwColumnLeft
        'DateLastUpdatedSort
        .ColumnHeaders.Item(GuiFeeSchedLevelsListView.DateLastUpdatedSort).Width = 0  'Hidden Sort Column
        .ColumnHeaders.Item(GuiFeeSchedLevelsListView.DateLastUpdatedSort).Alignment = lvwColumnLeft
        'UpdateByUserID
        .ColumnHeaders.Item(GuiFeeSchedLevelsListView.UpdateByUserID).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiFeeSchedLevelsListView.UpdateByUserID).Alignment = lvwColumnLeft
        'FeeScheduleLevelsID
        .ColumnHeaders.Item(GuiFeeSchedLevelsListView.FeeScheduleLevelsID).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiFeeSchedLevelsListView.FeeScheduleLevelsID).Alignment = lvwColumnLeft
        'FeeScheduleID
        .ColumnHeaders.Item(GuiFeeSchedLevelsListView.FeeScheduleID).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiFeeSchedLevelsListView.FeeScheduleID).Alignment = lvwColumnLeft
    End With
    
    bGridOn = CBool(GetSetting(App.EXEName, "GENERAL", "GRID_ON", False))
    lstvFeeSchedLevels.GridLines = bGridOn
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub LoadHeaderlstvFeeSchedLevels"
End Sub

Public Sub LoadHeaderlstvFeeTypes()
    On Error GoTo EH
    Dim bGridOn As Boolean
    
    'set the columnheaders
    With lstvFeeTypes
        .ColumnHeaders.Add , "TypeNum", "No."
        .ColumnHeaders.Add , "TypeNumSort", "TypeNumSort" 'Hidden
        .ColumnHeaders.Add , "ftName", "Name" 'Hidden
        .ColumnHeaders.Add , "Description", "Description"
        .ColumnHeaders.Add , "FeeAmount", "Fee Amount"
        .ColumnHeaders.Add , "FeeAmountSort", "FeeAmountSort" 'Hidden
        .ColumnHeaders.Add , "IsExpense", "Is Expense"
        .ColumnHeaders.Add , "MaxNumberOfItems", "Max Number Of Items"
        .ColumnHeaders.Add , "MaxNumberOfItemsSort", "MaxNumberOfItemsSort" 'hidden
        .ColumnHeaders.Add , "MaxFeeAmount", "Max Fee Amount"
        .ColumnHeaders.Add , "MaxFeeAmountSort", "MaxFeeAmountSort" 'Hidden
        .ColumnHeaders.Add , "UseFormula", "Use Formula"
        .ColumnHeaders.Add , "VBFormula", "VBFormula"
        .ColumnHeaders.Add , "DateLastUpdated", "Date Last Updated"
        .ColumnHeaders.Add , "DateLastUpdatedSort", "DateLastUpdatedSort" ' Hidden
        .ColumnHeaders.Add , "UpdateByUserID", "UpdateByUserID" ' Hidden
        .ColumnHeaders.Add , "FeeScheduleFeeTypesID", "FeeScheduleFeeTypesID" ' Hidden
        .ColumnHeaders.Add , "FeeScheduleID", "FeeScheduleID" ' Hidden
    
        .Sorted = False
        .SortOrder = lvwAscending

        'TypeNum
        .ColumnHeaders.Item(GuiFeeTypesListView.TypeNum).Width = 700
        .ColumnHeaders.Item(GuiFeeTypesListView.TypeNum).Alignment = lvwColumnLeft
        'TypeNumSort
        .ColumnHeaders.Item(GuiFeeTypesListView.TypeNumSort).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiFeeTypesListView.TypeNumSort).Alignment = lvwColumnLeft
        'ftName
        .ColumnHeaders.Item(GuiFeeTypesListView.ftName).Width = 0
        .ColumnHeaders.Item(GuiFeeTypesListView.ftName).Alignment = lvwColumnLeft
        'Description
        .ColumnHeaders.Item(GuiFeeTypesListView.Description).Width = 2460
        .ColumnHeaders.Item(GuiFeeTypesListView.Description).Alignment = lvwColumnLeft
        'FeeAmount
        .ColumnHeaders.Item(GuiFeeTypesListView.FeeAmount).Width = 1400
        .ColumnHeaders.Item(GuiFeeTypesListView.FeeAmount).Alignment = lvwColumnRight
        'FeeAmountSort
        .ColumnHeaders.Item(GuiFeeTypesListView.FeeAmountSort).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiFeeTypesListView.FeeAmountSort).Alignment = lvwColumnRight
        'IsExpense
        .ColumnHeaders.Item(GuiFeeTypesListView.IsExpense).Width = 1215
        .ColumnHeaders.Item(GuiFeeTypesListView.IsExpense).Alignment = lvwColumnCenter
        'MaxNumberOfItems
        .ColumnHeaders.Item(GuiFeeTypesListView.MaxNumberOfItems).Width = 2160
        .ColumnHeaders.Item(GuiFeeTypesListView.MaxNumberOfItems).Alignment = lvwColumnLeft
        'MaxNumberOfItemsSort
        .ColumnHeaders.Item(GuiFeeTypesListView.MaxNumberOfItemsSort).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiFeeTypesListView.MaxNumberOfItemsSort).Alignment = lvwColumnLeft
        'MaxFeeAmount
        .ColumnHeaders.Item(GuiFeeTypesListView.MaxFeeAmount).Width = 1965
        .ColumnHeaders.Item(GuiFeeTypesListView.MaxFeeAmount).Alignment = lvwColumnRight
        'MaxFeeAmountSort
        .ColumnHeaders.Item(GuiFeeTypesListView.MaxFeeAmountSort).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiFeeTypesListView.MaxFeeAmountSort).Alignment = lvwColumnRight
        'UseFormula
        .ColumnHeaders.Item(GuiFeeTypesListView.UseFormula).Width = 1410
        .ColumnHeaders.Item(GuiFeeTypesListView.UseFormula).Alignment = lvwColumnCenter
        'VBFormula
        .ColumnHeaders.Item(GuiFeeTypesListView.VBFormula).Width = 3480
        .ColumnHeaders.Item(GuiFeeTypesListView.VBFormula).Alignment = lvwColumnLeft
        'DateLastUpdated
        .ColumnHeaders.Item(GuiFeeTypesListView.DateLastUpdated).Width = 2200
        .ColumnHeaders.Item(GuiFeeTypesListView.DateLastUpdated).Alignment = lvwColumnLeft
        'DateLastUpdatedSort
        .ColumnHeaders.Item(GuiFeeTypesListView.DateLastUpdatedSort).Width = 0  'Hidden Sort Column
        .ColumnHeaders.Item(GuiFeeTypesListView.DateLastUpdatedSort).Alignment = lvwColumnLeft
        'UpdateByUserID
        .ColumnHeaders.Item(GuiFeeTypesListView.UpdateByUserID).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiFeeTypesListView.UpdateByUserID).Alignment = lvwColumnLeft
        'FeeScheduleFeeTypesID
        .ColumnHeaders.Item(GuiFeeTypesListView.FeeScheduleFeeTypesID).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiFeeTypesListView.FeeScheduleFeeTypesID).Alignment = lvwColumnLeft
        'FeeScheduleID
        .ColumnHeaders.Item(GuiFeeTypesListView.FeeScheduleID).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiFeeTypesListView.FeeScheduleID).Alignment = lvwColumnLeft
        
    End With
    
    bGridOn = CBool(GetSetting(App.EXEName, "GENERAL", "GRID_ON", False))
    lstvFeeTypes.GridLines = bGridOn
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub LoadHeaderlstvFeeTypes"
End Sub

Private Sub txtOptions_Change(Index As Integer)
    On Error GoTo EH
    Dim sctlName As String
    Dim sOptKey As String
    Dim sValue As String
    
    If mbFormLoading Then
        Exit Sub
    End If
    
    sctlName = txtOptions(Index).Name
    sOptKey = """" & sctlName & "_" & CStr(Index) & """"
    
    sValue = txtOptions(Index).Text
    
    SaveOptSettings mcolOptions, sOptKey, sValue
    
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub txtOptions_Change"
End Sub

Private Function SaveOptSettings(pcolOpt As Collection, psOptKey As String, psValue As String) As Boolean
    On Error GoTo EH
    Dim colOpt As Collection
    Dim sOptKey As String
    Dim sAppName As String
    Dim sSection As String
    Dim sKey As String
    Dim sValue As String
    Dim MyFSOptions As FSOptions
    
    
    'Init the local vars
    Set colOpt = pcolOpt
    sOptKey = psOptKey
    
    
    MyFSOptions = goUtil.GetItemFromCollection(colOpt, sOptKey)
    
    'set vars for Registry save
    sAppName = App.EXEName
    sSection = msFeeScheduleID & "_" & msClientCompanyID
    sKey = MyFSOptions.optVBREGSTR
    sValue = psValue
    'Save the setting to Registry
    SaveSetting sAppName, "FeeSchedule\" & sSection, sKey, sValue
    
    'cleanup
    Set colOpt = Nothing
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Function SaveOptSettings"
End Function

Private Sub txtOptions_GotFocus(Index As Integer)
    On Error GoTo EH
    
    goUtil.utSelText txtOptions(Index)
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub txtOptions_GotFocus"
End Sub
