VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReports 
   AutoRedraw      =   -1  'True
   Caption         =   "File"
   ClientHeight    =   6555
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6555
   ScaleWidth      =   11820
   StartUpPosition =   3  'Windows Default
   Tag             =   "Reports"
   Begin VB.Frame framReports 
      Height          =   4695
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   11295
      Begin MSComctlLib.ListView lvwRptParams 
         Height          =   3975
         Left            =   4320
         TabIndex        =   14
         Top             =   600
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   7011
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
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.ListBox cboMainReports 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2610
         ItemData        =   "frmReports.frx":0000
         Left            =   240
         List            =   "frmReports.frx":0002
         Sorted          =   -1  'True
         TabIndex        =   44
         Top             =   1995
         Width           =   3975
      End
      Begin VB.ListBox lstvbUserDefinedType 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3885
         Left            =   4320
         TabIndex        =   0
         Top             =   720
         Visible         =   0   'False
         Width           =   6855
      End
      Begin VB.CheckBox chkPrintPreview 
         Alignment       =   1  'Right Justify
         Caption         =   "Print Preview"
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
         Left            =   9120
         TabIndex        =   12
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdEditDiagram 
         Caption         =   "Diagra&m"
         Enabled         =   0   'False
         Height          =   375
         Left            =   7920
         TabIndex        =   11
         Top             =   240
         Visible         =   0   'False
         Width           =   1100
      End
      Begin VB.Frame framMultiReport 
         Caption         =   "Multi Report Items"
         Height          =   1455
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   3975
         Begin VB.Frame framMultiReportCommand 
            Height          =   1455
            Left            =   3000
            TabIndex        =   5
            Top             =   0
            Width           =   975
            Begin VB.CommandButton cmdAddMultiReport 
               Caption         =   "&Add"
               Height          =   375
               Left            =   120
               TabIndex        =   6
               Top             =   960
               Width           =   735
            End
            Begin VB.Image imgDiagram 
               Height          =   615
               Left            =   120
               Stretch         =   -1  'True
               Top             =   240
               Width           =   735
            End
         End
         Begin VB.OptionButton optMultiReport 
            Caption         =   "WorkSheet Diagram"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   4
            Tag             =   "RTWSDiagram"
            Top             =   360
            Width           =   2775
         End
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
         TabIndex        =   13
         Top             =   240
         Width           =   975
      End
      Begin MSComctlLib.ImageList imgRptParamsStatus 
         Left            =   8760
         Top             =   1200
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
               Picture         =   "frmReports.frx":0004
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmReports.frx":015E
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmReports.frx":054A
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmReports.frx":06BB
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmReports.frx":0AE9
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmReports.frx":0F3B
               Key             =   ""
            EndProperty
         EndProperty
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
         Left            =   5880
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   960
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.CommandButton cmdSelAll 
         Caption         =   "&Select A&ll"
         Height          =   375
         Left            =   4320
         TabIndex        =   8
         Top             =   240
         Width           =   1100
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Find"
         Height          =   375
         Left            =   5520
         TabIndex        =   9
         Top             =   240
         Width           =   1100
      End
      Begin VB.CommandButton cmdFindNext 
         Caption         =   "Find &Next"
         Height          =   375
         Left            =   6720
         TabIndex        =   10
         Top             =   240
         Width           =   1100
      End
      Begin VB.Frame framMultiUpdate 
         Caption         =   "Multi Update"
         Height          =   735
         Left            =   4320
         TabIndex        =   22
         Top             =   3840
         Visible         =   0   'False
         Width           =   6855
         Begin VB.CommandButton cmdUpdateMulti 
            Caption         =   "&Update"
            Height          =   375
            Index           =   0
            Left            =   1800
            TabIndex        =   24
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton cmdUpdateMulti 
            Caption         =   "Up&date"
            Height          =   375
            Index           =   1
            Left            =   4560
            TabIndex        =   27
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton cmdDateMultiUpdate 
            Height          =   375
            Left            =   4080
            Picture         =   "frmReports.frx":1301
            Style           =   1  'Graphical
            TabIndex        =   26
            Tag             =   "Date"
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox txtDateMultiUpdate 
            Height          =   375
            Left            =   2880
            TabIndex        =   25
            Tag             =   "Date"
            Top             =   240
            Width           =   1575
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
            TabIndex        =   23
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Frame framEditParam 
         Height          =   3255
         Left            =   4320
         TabIndex        =   15
         Top             =   600
         Visible         =   0   'False
         Width           =   6855
         Begin VB.CommandButton cmdCancelEdit 
            Caption         =   "&Cancel"
            Height          =   375
            Left            =   5880
            TabIndex        =   20
            Top             =   2340
            Width           =   855
         End
         Begin VB.CommandButton cmdParamDate 
            Height          =   375
            Left            =   5400
            Picture         =   "frmReports.frx":1743
            Style           =   1  'Graphical
            TabIndex        =   19
            Tag             =   "Date"
            Top             =   2760
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txtParamCaption 
            Appearance      =   0  'Flat
            Height          =   1935
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   240
            Width           =   6615
         End
         Begin VB.TextBox txtParamValue 
            Height          =   375
            Left            =   120
            MaxLength       =   300
            TabIndex        =   18
            Top             =   2760
            Visible         =   0   'False
            Width           =   5655
         End
         Begin VB.CheckBox chkParamBoolean 
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   2400
            Visible         =   0   'False
            Width           =   5655
         End
         Begin VB.CommandButton cmdUpdateEdit 
            Caption         =   "&Update"
            Height          =   375
            Left            =   5880
            TabIndex        =   21
            Top             =   2760
            Width           =   855
         End
      End
      Begin VB.Label lblMainReports 
         Caption         =   "Reports:"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1760
         Width           =   2655
      End
   End
   Begin VB.Frame framWordExcel 
      Height          =   4695
      Left            =   240
      TabIndex        =   28
      Top             =   480
      Visible         =   0   'False
      Width           =   11295
      Begin VB.CommandButton cmdReloadWordXL 
         Caption         =   "&Reload"
         Enabled         =   0   'False
         Height          =   615
         Left            =   10200
         Picture         =   "frmReports.frx":1B85
         Style           =   1  'Graphical
         TabIndex        =   31
         TabStop         =   0   'False
         Tag             =   "Enable"
         ToolTipText     =   "Reload Word and Excel Applications"
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdPrintDoc 
         Caption         =   "View &Document"
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
         Height          =   615
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   3900
         Width           =   975
      End
      Begin MSComctlLib.ImageList imgVarDoc 
         Left            =   10440
         Top             =   2400
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmReports.frx":1F47
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmReports.frx":2283
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmReports.frx":25AF
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Frame framName 
         Caption         =   "Selected Report"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   680
         Left            =   1200
         TabIndex        =   35
         Top             =   3840
         Width           =   9975
         Begin VB.Image imgSelected 
            Height          =   360
            Left            =   120
            Stretch         =   -1  'True
            Top             =   247
            Width           =   360
         End
         Begin VB.Label lblName 
            Height          =   375
            Left            =   555
            TabIndex        =   36
            Top             =   240
            Width           =   6495
         End
         Begin VB.Label lblDate 
            Height          =   375
            Left            =   7680
            TabIndex        =   37
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.ComboBox cboWordXLDocs 
         Height          =   360
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   480
         Width           =   9975
      End
      Begin MSComctlLib.ListView lvwAvail 
         Height          =   2535
         Left            =   120
         TabIndex        =   33
         Tag             =   "Enable"
         Top             =   1200
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   4471
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
      Begin VB.Label lblAvail 
         Caption         =   "Available Reports"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label lblDocPackages 
         Caption         =   "Document Packages:"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   2655
      End
   End
   Begin MSComctlLib.TabStrip TSReports 
      Height          =   5175
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   9128
      ShowTips        =   0   'False
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Reports"
            Object.Tag             =   "framReports"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Word && Excel Documents"
            Object.Tag             =   "framWordExcel"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
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
      Left            =   7200
      TabIndex        =   38
      Top             =   5280
      Width           =   4455
      Begin VB.CommandButton cmdPrintList 
         Caption         =   "Sho&w Item List"
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
         MaskColor       =   &H00000000&
         Picture         =   "frmReports.frx":28D3
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Exit"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdSpelling 
         Caption         =   "Spellin&g"
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
         Left            =   1200
         MaskColor       =   &H00000000&
         Picture         =   "frmReports.frx":2D15
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Exit"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Sa&ve"
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
         Height          =   855
         Left            =   2280
         MaskColor       =   &H00000000&
         Picture         =   "frmReports.frx":315F
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Exit"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
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
         Left            =   3360
         MaskColor       =   &H00000000&
         Picture         =   "frmReports.frx":32A9
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Exit"
         Top             =   240
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mfrmClaim As frmClaim
Private mbUnloadMe As Boolean
Private msAssignmentsID As String
Private moGUI As V2ECKeyBoard.clsCarGUI
Private moCurrentTextBox As TextBox
Private moActivecboReport As ListBox
Private mbLoading As Boolean 'Loading Form
Private mbLoadingMe As Boolean 'Just loading data
Private msFindText As String
Private mlLastFindIndex As Long
Private mbShowingEditRptParam As Boolean
Private mbPrintPreview As Boolean
Private mbLoadingDesc As Boolean

'WordExcel Vars
Private moWordXL As V2ECKeyBoard.clsWordXL
Private mbLoadingWordXL As Boolean
'Multi Report Control
Private moOptMultiReport As OptionButton
Private mlEditDiagramNumber As Long
Private msIBNUM As String

Public Property Let IBNUM(psIBNUM As String)
    msIBNUM = psIBNUM
End Property
Public Property Get IBNUM() As String
    IBNUM = msIBNUM
End Property


Public Property Let EditDiagramNumber(plNum As Long)
    mlEditDiagramNumber = plNum
End Property
Public Property Get EditDiagramNumber() As Long
    EditDiagramNumber = mlEditDiagramNumber
End Property


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

'Private Sub cboCarSpecReports_Click()
'    On Error GoTo EH
'
'    If Not cboCarSpecReports.ListIndex = -1 Then
'        cboMainReports.ListIndex = -1
'        cboIB.ListIndex = -1
'        cboPayments.ListIndex = -1
'        If mfrmClaim.GetRptParamColAndLoadLvw(cboCarSpecReports, lvwRptParams, framMultiUpdate) Then
'            'Enable the Print button
'            cmdPrintReport.Enabled = True
'        Else
'            cmdPrintReport.Enabled = False
'        End If
'        Set moActivecboReport = cboCarSpecReports
'    End If
'
'    Exit Sub
'EH:
'    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cboCarSpecReports_Click"
'End Sub


'Private Sub cboIB_Click()
'    On Error GoTo EH
'
'    If Not cboIB.ListIndex = -1 Then
'        cboMainReports.ListIndex = -1
'        cboCarSpecReports.ListIndex = -1
'        cboPayments.ListIndex = -1
'        If mfrmClaim.GetRptParamColAndLoadLvw(cboIB, lvwRptParams, framMultiUpdate) Then
'            'Enable the Print button
'            cmdPrintReport.Enabled = True
'        Else
'            cmdPrintReport.Enabled = False
'        End If
'        Set moActivecboReport = cboIB
'    End If
'
'    Exit Sub
'EH:
'    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cboIB_Click"
'End Sub

Private Sub cboMainReports_Click()
    On Error GoTo EH
    
    If Not cboMainReports.ListIndex = -1 Then
        Screen.MousePointer = VBRUN.MousePointerConstants.vbHourglass
        If mfrmClaim.GetRptParamColAndLoadLvw(cboMainReports, lvwRptParams, framMultiUpdate) Then
            'Enable the Print button
            cmdPrintReport.Enabled = True
        Else
            cmdPrintReport.Enabled = False
        End If
        Set moActivecboReport = cboMainReports
    End If
    Screen.MousePointer = VBRUN.MousePointerConstants.vbDefault
    Exit Sub
EH:
    Screen.MousePointer = VBRUN.MousePointerConstants.vbDefault
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cboMainReports_Click"
End Sub

'Private Sub cboPayments_Click()
'    On Error GoTo EH
'
'    If Not cboPayments.ListIndex = -1 Then
'        cboMainReports.ListIndex = -1
'        cboCarSpecReports.ListIndex = -1
'        cboIB.ListIndex = -1
'        If mfrmClaim.GetRptParamColAndLoadLvw(cboPayments, lvwRptParams, framMultiUpdate) Then
'            'Enable the Print button
'            cmdPrintReport.Enabled = True
'        Else
'            cmdPrintReport.Enabled = False
'        End If
'        Set moActivecboReport = cboPayments
'    End If
'
'    Exit Sub
'EH:
'    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cboPayments_Click"
'End Sub

Private Sub cboWordXLDocs_Click()
    On Error GoTo EH
    Dim sData As String
    Dim sWordXLDocPath As String
    
    If cboWordXLDocs.ListIndex = -1 Then
        Exit Sub
    End If
    
    If LoadWordXL() Then
        sData = cboWordXLDocs.Text
        If sData <> vbNullString Then
            sWordXLDocPath = Trim(left(sData, 200))
            sData = Mid(sData, 200)
            sData = Trim(sData)
            sWordXLDocPath = sData
            moWordXL.WordXLDocPath = sWordXLDocPath
            PopulatelvwAvail
            cmdPrintDoc.Enabled = False
            cmdReloadWordXL.Enabled = True
        End If
    End If
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cboWordXLDocs_Click"
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

Private Sub cmdAddMultiReport_Click()
    On Error GoTo EH
    Dim MyfrmAddMultiReportItem As AddMultiReportItem
    
    'Check to see if a multi report is selected
    If moOptMultiReport Is Nothing Then
        cmdAddMultiReport.Enabled = False
        Exit Sub
    End If
    
    Set MyfrmAddMultiReportItem = New AddMultiReportItem
    
    With MyfrmAddMultiReportItem
        .MyfrmClaim = mfrmClaim
        .AssignmentsID = msAssignmentsID
        .TableName = moOptMultiReport.Tag
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
    cboMainReports.Enabled = True
'    cboCarSpecReports.Enabled = True
'    cboIB.Enabled = True
'    cboPayments.Enabled = True
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

Private Sub cmdEditDiagram_Click()
    On Error GoTo EH
    
    ShowDiagramEdit
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdEditDiagram_Click"
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

Private Sub cmdPrintDoc_Click()
    PrintPreviewDoc
End Sub

Public Function PrintPreviewDoc() As Boolean
    On Error GoTo EH
    Dim iType As Pic
    
    cmdPrintDoc.Enabled = False
    If lblName.Caption <> vbNullString Then
        If InStr(1, lblName.Caption, ".xls", vbTextCompare) > 0 Then
            iType = Pic.XL
        ElseIf InStr(1, lblName.Caption, ".doc", vbTextCompare) > 0 Then
            iType = Pic.Word
        End If
        Screen.MousePointer = vbHourglass
        If moWordXL.PrintIt(Me, iType, lblName.Caption, lblDate.Caption, Nothing, Nothing, Nothing) Then
            PrintPreviewDoc = True
        End If
        lvwAvail.ListItems.Clear
        imgSelected.Picture = LoadPicture(vbNullString)
        lblName.Caption = vbNullString
        lblDate.Caption = vbNullString
        Screen.MousePointer = vbDefault
    End If
    
    Exit Function
EH:
    Screen.MousePointer = vbDefault
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function PrintPreviewDoc"
End Function

Private Sub cmdPrintList_Click()
    On Error GoTo EH
    
    If lvwAvail.Visible Then
        goUtil.utPrintListView App.EXEName, lvwAvail, "Word And Excel Documents"
    End If
    
    If lvwRptParams.Visible Then
        goUtil.utPrintListView App.EXEName, lvwRptParams, "Report Parameters"
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdPrintList_Click"
End Sub


Private Sub cmdPrintReport_Click()
    On Error GoTo EH
    Dim sCopy As String
    
    If moActivecboReport Is Nothing Then
        Exit Sub
    End If
    
    'Only allow the IB at this time to pass in Copy parameter
    
'    If StrComp(moActivecboReport.Name, cboIB.Name, vbTextCompare) = 0 Then
'        sCopy = cboCopy.Text
'    Else
'        sCopy = vbNullString
'    End If
        
    Screen.MousePointer = vbHourglass
    'First be sure control are valid
    
    cmdPrintReport.Enabled = False
    cboMainReports.Enabled = False
'    cboCarSpecReports.Enabled = False
'    cboIB.Enabled = False
'    cboPayments.Enabled = False
    If mfrmClaim.PrintActiveReport(moActivecboReport, , sCopy, mbPrintPreview) Then
        If Not mbUnloadMe Then
            DoEvents
            Sleep 1000
            cmdPrintReport.Enabled = True
            cmdPrintReport.Enabled = True
            cboMainReports.Enabled = True
'            cboCarSpecReports.Enabled = True
'            cboIB.Enabled = True
'            cboPayments.Enabled = True
        End If
    End If
    Screen.MousePointer = vbDefault
    
    Exit Sub
EH:
    Screen.MousePointer = vbDefault
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdPrintReport_Click"
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
        cboMainReports.Enabled = True
'        cboCarSpecReports.Enabled = True
'        cboIB.Enabled = True
'        cboPayments.Enabled = True
        cmdSpelling.Enabled = True
        lvwRptParams.SetFocus
    End If
    
   
    Screen.MousePointer = MousePointerConstants.vbDefault
    Exit Function
EH:
    Screen.MousePointer = MousePointerConstants.vbDefault
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdUpdateEdit_Click"
End Function

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

Public Function LoadWordXL() As Boolean
    On Error GoTo EH

    If mbLoadingWordXL Then
        Exit Function
    End If
    mbLoadingWordXL = True
    Screen.MousePointer = VBRUN.MousePointerConstants.vbHourglass
    If Not moWordXL Is Nothing Then
        moWordXL.CLEANUP
    End If
    Set moWordXL = New V2ECKeyBoard.clsWordXL
    moWordXL.SetUtilObject goUtil
    With moWordXL
        .aryQVariables = GetVariables
        .DocVarID = msAssignmentsID
    End With
    If moWordXL.LoadWordXLAPP(Me) Then
        LoadWordXL = True
    End If
    mbLoadingWordXL = False
    Screen.MousePointer = VBRUN.MousePointerConstants.vbDefault
    Exit Function
EH:
    Screen.MousePointer = VBRUN.MousePointerConstants.vbDefault
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdWordXL_Click"
End Function

Private Sub cmdReloadWordXL_Click()
    On Error GoTo EH
    
    If cboWordXLDocs.ListIndex <> -1 Then
        cboWordXLDocs_Click
    Else
        cmdReloadWordXL.Enabled = False
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdReloadWordXL_Click"
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
    Me.Icon = mfrmClaim.optClaim(GuiClaimOptions.opt06_Reports).Picture
    
    mbPrintPreview = CBool(GetSetting(App.EXEName, "GENERAL", "PRINT_PREVIEW", True))
    If mbPrintPreview Then
        chkPrintPreview.Value = vbChecked
    Else
        chkPrintPreview.Value = vbUnchecked
    End If
    
    LoadHeaderlvwRptParams
    LoadHeaderWordXL
    If LoadWordXLReports() Then
        cmdPrintDoc.Enabled = False
    End If
    LoadMe
    
    ShowFrame
    cmdSave.Enabled = False
    mbLoading = False
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_Load"
End Sub

Public Function LoadMe() As Boolean
    On Error GoTo EH
   
    mbLoadingMe = True
    
    LoadReports
'    LoadCopy
    lvwRptParams.ListItems.Clear
    cmdPrintReport.Enabled = False
    Set moActivecboReport = Nothing
    
    LoadMe = True
    mbLoadingMe = False
    Exit Function
EH:
    mbLoadingMe = False
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function LoadMe"
End Function

'Private Sub LoadCopy()
'    On Error GoTo EH
'
'    cboCopy.Clear
'    cboCopy.AddItem "(-ALL COPIES-)"
'    cboCopy.AddItem goUtil.gsCurCarDBName & " Copy"
'    cboCopy.AddItem GetSetting(goUtil.gsAppEXEName, "GENERAL", "CURRENT_COMPANY_NAME", "Company") & " Copy"
'    cboCopy.AddItem "Remit Copy"
'    cboCopy.AddItem "Adjuster Copy"
'
'    Exit Sub
'EH:
'    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function LoadCopy"
'End Sub

Public Function SaveMe() As Boolean
    
    cmdSave.Enabled = False
    SaveMe = True
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function SaveMe"
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
    TSReports.Width = Me.Width - 405
    framReports.Width = Me.Width - 645
    cmdPrintReport.left = Me.Width - 1740
    chkPrintPreview.left = Me.Width - 2820
    lvwRptParams.Width = Me.Width - 5085
    framEditParam.Width = Me.Width - 5085
    lstvbUserDefinedType.Width = Me.Width - 5085
    framMultiUpdate.Width = Me.Width - 5085
    txtParamCaption.Width = Me.Width - 5300
    txtParamValue.Width = Me.Width - 6285
    cmdParamDate.left = Me.Width - 6540
    cmdCancelEdit.left = Me.Width - 6060
    cmdUpdateEdit.left = Me.Width - 6060
    
    'WordXL
    framWordExcel.Width = Me.Width - 645
    cboWordXLDocs.Width = Me.Width - 1965
    cmdReloadWordXL.left = Me.Width - 1740
    lvwAvail.Width = Me.Width - 885
    framName.Width = Me.Width - 1965
    lblName.Width = Me.Width - 5445
    lblDate.left = Me.Width - 4260
    
    'framCommands
    framCommands.left = Me.Width - 4740
    
    
    'Heights and Tops
    TSReports.Height = Me.Height - 1785
    framReports.Height = Me.Height - 2265
    lvwRptParams.Height = Me.Height - 3075
    framEditParam.Height = Me.Height - 3705
    lstvbUserDefinedType.Height = Me.Height - 3075
    framMultiUpdate.top = Me.Height - 3120
    cboMainReports.Height = Me.Height - 4350
    
    'wordXl
    framWordExcel.Height = Me.Height - 2265
    lvwAvail.Height = Me.Height - 4425
    cmdPrintDoc.top = Me.Height - 3060
    framName.top = Me.Height - 3120
    
    'framCommands
    framCommands.top = Me.Height - 1680
    
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

    Set moCurrentTextBox = Nothing
    Set moActivecboReport = Nothing
    
    If Not moWordXL Is Nothing Then
        moWordXL.CLEANUP
        Set moWordXL = Nothing
    End If
    
    Set moOptMultiReport = Nothing
    
    CLEANUP = True
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function CLEANUP"
End Function

Private Sub imgDiagram_Click()
    On Error GoTo EH
    
    ShowDiagramEdit
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub imgDiagram_Click"
End Sub

Private Sub lstvbUserDefinedType_DblClick()
    On Error GoTo EH
    
    If mbLoadingDesc Then
        Exit Sub
    End If
    
    If lstvbUserDefinedType.ListIndex > -1 Then
        txtParamValue.Text = lstvbUserDefinedType.Text
        UpdateEdit
        lstvbUserDefinedType.Visible = False
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

Private Sub lvwAvail_Click()
    LoadAvail
End Sub

Private Sub lvwAvail_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    On Error GoTo EH
    Dim sMess As String
    If lvwAvail.SortOrder = lvwAscending Then
        lvwAvail.SortOrder = lvwDescending
        sMess = "Descending"
    Else
        lvwAvail.SortOrder = lvwAscending
        sMess = "Ascending"
    End If
    
    'Set the Tool Tip
    lvwAvail.ToolTipText = "Sort By " & ColumnHeader.Text & " " & sMess
    
    'Need to see if the Column clicked was a NON text sort
    'Like Date or number.  If a non textual column was clicked
    'need to use the next column as the sort since this next column
    'will be hidden and contain a sort friendly format.
    'Sort Key is Base 0 where CoulmnHeader is not
    Select Case ColumnHeader.Index
        Case AvailDocs.DateCreated, AvailDocs.DateLastUpdated
            lvwAvail.SortKey = ColumnHeader.Index
        Case Else
            lvwAvail.SortKey = ColumnHeader.Index - 1
    End Select
    
    lvwAvail.Sorted = True
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lvwAvail_ColumnClick"
End Sub

Private Sub lvwAvail_DblClick()
    PrintPreviewDoc
End Sub

Private Sub lvwAvail_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo EH
    
    Select Case KeyCode
        Case vbKeyReturn
            LoadAvail
            PrintPreviewDoc
    End Select
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lvwAvail_KeyDown"
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
            ElseIf iDataType = vbUserDefinedType Then
                ShowEditRptParam True
            End If
        End If
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lvwRptParams_MouseUp"
End Sub

Private Sub optMultiReport_Click(Index As Integer)
    On Error GoTo EH
    
    Set moOptMultiReport = optMultiReport(Index)
    cmdAddMultiReport.Enabled = True
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub optMultiReport_Click"
End Sub

Private Sub TSReports_Click()
    ShowFrame
End Sub

Public Function ShowFrame() As Boolean
    On Error GoTo EH
    Dim sFrameName As String
    Dim oFrame As Control
    Dim MyFrame As Frame
    Dim oControl As Control
    
    sFrameName = TSReports.SelectedItem.Tag
    
    For Each oFrame In Me.Controls
        If TypeOf oFrame Is Frame Then
            Set MyFrame = oFrame
            If StrComp(MyFrame.Name, sFrameName, vbTextCompare) = 0 Then
                MyFrame.Visible = True
                'If the frame is disabled then disable all the controls
                If Not MyFrame.Enabled Then
                    For Each oControl In Me.Controls
                        If Not TypeOf oControl Is ImageList Then
                            If oControl.Container.Name = MyFrame.Name Then
                                oControl.Enabled = False
                            End If
                        End If
                    Next
                End If
            Else
                If StrComp(MyFrame.Name, framCommands.Name, vbTextCompare) = 0 Then
                    MyFrame.Visible = True
                ElseIf StrComp(MyFrame.Name, framName.Name, vbTextCompare) = 0 Then
                    MyFrame.Visible = True
                ElseIf StrComp(MyFrame.Name, framMultiReport.Name, vbTextCompare) = 0 Then
                    MyFrame.Visible = True
                ElseIf StrComp(MyFrame.Name, framMultiReportCommand.Name, vbTextCompare) = 0 Then
                    MyFrame.Visible = True
                Else
                    MyFrame.Visible = False
                End If
            End If
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

Private Sub txtDateMultiUpdate_GotFocus()
    goUtil.utSelText txtDateMultiUpdate
End Sub

Private Sub txtDateMultiUpdate_LostFocus()
    goUtil.utValidate , txtDateMultiUpdate
End Sub

Private Sub txtParamValue_GotFocus()
    goUtil.utSelText txtParamValue
End Sub

Public Function LoadReports() As Boolean
    On Error GoTo EH
    Dim lNewIndex As Long
    Dim sData As String
    Dim RS As ADODB.Recordset
    Dim RSIB As ADODB.Recordset
    Dim RSPayment As ADODB.Recordset
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
    sData = "Loss Report _" & sIBNUM & "_" & sCLIENTNUM
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
                        sPhotoReportName = RSRTPhotoReportList.Fields("Name").Value
                        sPhotoReportNumber = RSRTPhotoReportList.Fields("Number").Value
                        sPhotoReportDesc = RSRTPhotoReportList.Fields("Description").Value
                        If Len(sPhotoReportDesc) > 20 Then
                            sPhotoReportDesc = left(sPhotoReportDesc, 10) & "..."
                        End If
                        sData = sPhotoReportName & " - (" & sPhotoReportDesc & ")"
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
                        sDiagramName = RSRTWSDiagramList.Fields("Name").Value
                        sDiagramNumber = RSRTWSDiagramList.Fields("Number").Value
                        sDiagramDesc = RSRTWSDiagramList.Fields("Description").Value
                        If Len(sDiagramDesc) > 20 Then
                            sDiagramDesc = left(sDiagramDesc, 10) & "..."
                        End If
                        sData = sDiagramName & " - (" & sDiagramDesc & ")"
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
                sData = RS.Fields("Description").Value
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
                        sPhotoReportName = RSRTPhotoReportList.Fields("Name").Value
                        sPhotoReportNumber = RSRTPhotoReportList.Fields("Number").Value
                        sPhotoReportDesc = RSRTPhotoReportList.Fields("Description").Value
                        If Len(sPhotoReportDesc) > 20 Then
                            sPhotoReportDesc = left(sPhotoReportDesc, 10) & "..."
                        End If
                        sData = sPhotoReportName & " - (" & sPhotoReportDesc & ")"
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
                        sDiagramName = RSRTWSDiagramList.Fields("Name").Value
                        sDiagramNumber = RSRTWSDiagramList.Fields("Number").Value
                        sDiagramDesc = RSRTWSDiagramList.Fields("Description").Value
                        If Len(sDiagramDesc) > 20 Then
                            sDiagramDesc = left(sDiagramDesc, 10) & "..."
                        End If
                        sData = sDiagramName & " - (" & sDiagramDesc & ")"
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
                sData = RS.Fields("Description").Value & " - (Previous Version) "
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
            sData = RS.Fields("Description").Value
            sData = sData & String(200, " ")
            sData = sData & RS.Fields("ProjectName").Value & "|" & RS.Fields("ClassName").Value & "|" & RS.Fields("Version").Value
'            cboCarSpecReports.AddItem sData
            cboMainReports.AddItem sData
'            lNewIndex = cboCarSpecReports.NewIndex
            lNewIndex = cboMainReports.NewIndex
            'cboCarSpecReports.ItemData(lNewIndex) = RS.Fields("ApplicationID").Value
            cboMainReports.ItemData(lNewIndex) = RS.Fields("ApplicationID").Value
            RS.MoveNext
        Loop
    End If
    
    'Need to Add Items from History Table
    Set RS = mfrmClaim.adoRSCarSpecReportsHistory
    
    If RS.RecordCount > 0 Then
        Do Until RS.EOF
            sData = vbNullString
            sData = RS.Fields("Description").Value & " - (Previous Version) "
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
            sData = RSIB.Fields("sIBNumber").Value
            sData = sData & String(200, " ")
            sData = sData & RS.Fields("ProjectName").Value & "|" & RS.Fields("ClassName").Value & "|" & RS.Fields("Version").Value
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
            sData = RSIB.Fields("sIBNumber").Value & " - (Previous Version) "
            sData = sData & String(200, " ")
            sData = sData & RS.Fields("ProjectName").Value & "|" & RS.Fields("ClassName").Value & "|" & RS.Fields("Version").Value
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
            If CBool(RSPayment.Fields("PrintOnIB").Value) Then
                sData = RS.Fields("Description").Value & " (" & RSPayment.Fields("CheckNum").Value & ") - Print On IB"
            Else
                sData = RS.Fields("Description").Value & " (" & RSPayment.Fields("CheckNum").Value & ")"
            End If
            sData = sData & String(200, " ")
            sData = sData & RS.Fields("ProjectName").Value & "|" & RS.Fields("ClassName").Value & "|" & RS.Fields("Version").Value
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
            sData = RS.Fields("Description").Value & " (" & RSPayment.Fields("CheckNum").Value & ")"
            sData = sData & String(200, " ")
            sData = sData & RS.Fields("ProjectName").Value & "|" & RS.Fields("ClassName").Value & "|" & RS.Fields("Version").Value
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
    
    'cleanup
    Set RS = Nothing
    Set RSIB = Nothing
    Set RSPayment = Nothing
    Set RSRTPhotoReportList = Nothing
    Set RSRTWSDiagramList = Nothing
    Set MyadoRSAssignments = Nothing
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function LoadReports"
End Function

Public Function LoadWordXLReports() As Boolean
    Dim lNewIndex As Long
    Dim sData As String
    Dim sWordXLInstallFileLocation As String
    Dim RS As ADODB.Recordset
    
    'Load RS for Word XL Documents
    mfrmClaim.SetadoRSWordXLDocs msAssignmentsID
    
    'Load the Word XL Documents
    Set RS = mfrmClaim.adoRSWordXLDocs
    cboWordXLDocs.Clear
    lvwAvail.ListItems.Clear
    If RS.RecordCount > 0 Then
        Do Until RS.EOF
            sData = vbNullString
            sData = RS.Fields("Description").Value
            sData = sData & String(200, " ")
            sWordXLInstallFileLocation = RS.Fields("InstallFileLocation").Value
            sWordXLInstallFileLocation = Replace(sWordXLInstallFileLocation, "{InstallDir}", goUtil.gsInstallDir, , , vbTextCompare)
            sData = sData & sWordXLInstallFileLocation
            cboWordXLDocs.AddItem sData
            lNewIndex = cboWordXLDocs.NewIndex
            cboWordXLDocs.ItemData(lNewIndex) = RS.Fields("DocumentID").Value
            RS.MoveNext
        Loop
    End If
    LoadWordXLReports = True
    
    'cleanup
    Set RS = Nothing
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function LoadWordXLReports"
End Function

Public Sub LoadHeaderlvwRptParams()
    On Error GoTo EH
    Dim bGridOn As Boolean
    Dim bHideDeleted As Boolean
    Dim bHideUploadFlags As Boolean
    
    bHideDeleted = CBool(GetSetting(App.EXEName, "GENERAL", "HIDE_DELETED", True))
    bHideUploadFlags = CBool(GetSetting(App.EXEName, "GENERAL", "HIDE_UPLOAD_FLAGS", True))
    
    'set the columnheaders
    With lvwRptParams
        .ColumnHeaders.Add , "ParamCaption", "Caption"
        .ColumnHeaders.Add , "ParamValue", "Value"
        .ColumnHeaders.Add , "IsDeleted", "Is Deleted" 'hidden
        .ColumnHeaders.Add , "UpLoadMe", "UpLoad Me" 'hidden
        .ColumnHeaders.Add , "DateLastUpdated", "Date Last Updated"
        .ColumnHeaders.Add , "DateLastUpdatedSort", "Sort Date Last Updated" ' hidden
        .ColumnHeaders.Add , "AdminComments", "Admin Comments" ' hidden
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
        .ColumnHeaders.Item(GuiRptParamsListView.ParamCaption).Width = 8000
        .ColumnHeaders.Item(GuiRptParamsListView.ParamCaption).Alignment = lvwColumnLeft
        'ParamValue
        .ColumnHeaders.Item(GuiRptParamsListView.ParamValue).Width = 2500
        .ColumnHeaders.Item(GuiRptParamsListView.ParamValue).Alignment = lvwColumnLeft
        'Is Deleted
        If bHideDeleted Then
            .ColumnHeaders.Item(GuiRptParamsListView.IsDeleted).Width = 0  'Hidden 400
        Else
            .ColumnHeaders.Item(GuiRptParamsListView.IsDeleted).Width = 400
        End If
        .ColumnHeaders.Item(GuiRptParamsListView.IsDeleted).Alignment = lvwColumnCenter
        .ColumnHeaders.Item(GuiRptParamsListView.IsDeleted).Icon = GuiRptParamsStatusList.IsDeleted
        'UpLoad Me
        If bHideUploadFlags Then
            .ColumnHeaders.Item(GuiRptParamsListView.UpLoadMe).Width = 0  'Hidden 400
        Else
            .ColumnHeaders.Item(GuiRptParamsListView.UpLoadMe).Width = 400
        End If
        .ColumnHeaders.Item(GuiRptParamsListView.UpLoadMe).Alignment = lvwColumnCenter
        .ColumnHeaders.Item(GuiRptParamsListView.UpLoadMe).Icon = GuiRptParamsStatusList.UpLoadMe
        'DateLastUpdated
        .ColumnHeaders.Item(GuiRptParamsListView.DateLastUpdated).Width = 2200
        .ColumnHeaders.Item(GuiRptParamsListView.DateLastUpdated).Alignment = lvwColumnLeft
        'DateLastUpdatedSort
        .ColumnHeaders.Item(GuiRptParamsListView.DateLastUpdatedSort).Width = 0  'Hidden Sort Column
        .ColumnHeaders.Item(GuiRptParamsListView.DateLastUpdatedSort).Alignment = lvwColumnLeft
        'AdminComments
        .ColumnHeaders.Item(GuiRptParamsListView.AdminComments).Width = 0  'Hidden 10000
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

Private Function ShowEditRptParam(Optional pbMouseClickBoolean As Boolean = False) As Boolean
    On Error GoTo EH
    Dim iDataType As VBA.VbVarType
    Dim itmX As ListItem
    Dim sFlagText As String
    Dim sLossFormat As String
    
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
        cmdUpdateEdit.Default = True
        framMultiUpdate.Visible = False
        cboMainReports.Enabled = False
'        cboCarSpecReports.Enabled = False
'        cboIB.Enabled = False
'        cboPayments.Enabled = False
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
                txtParamValue.Tag = "Date"
                txtParamValue.MaxLength = 20
                txtParamValue.Text = itmX.SubItems(GuiRptParamsListView.ParamValue - 1)
                cmdParamDate.SetFocus
            End If
        Case VBA.VbVarType.vbUserDefinedType
            'Check for LRFOrmat
            sLossFormat = goUtil.IsNullIsVbNullString(mfrmClaim.adoRSAssignments.Fields("LRFormat"))
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
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Function ShowEditRptParam"
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
    Dim itmX As ListItem
    Dim sUDTName As String
    'Wddx Objects For ContactsID
    Dim sLossReportData As String
    Dim sContactRowID As String
    Dim sFirstName As String
    Dim sLastName As String
    Dim sContactRole As String
    Dim oDeser As WDDXDeserializer
    Dim oMyStruct As WDDXStruct
    Dim oContactsRS As WDDXRecordset
    'Vars for Texas sub coverage Code Lookups
    
    Set itmX = lvwRptParams.SelectedItem
    sUDTName = itmX.ListSubItems((GuiRptParamsListView.ParamName - 1))
    
    sLossFormat = goUtil.IsNullIsVbNullString(mfrmClaim.adoRSAssignments.Fields("LRFormat"))
    If StrComp(sLossFormat, "V2ECcarFarmers.clsLossXML01", vbTextCompare) = 0 Then
        lstvbUserDefinedType.Visible = True
        'Check for Different User Defined Types Name to populate
        If StrComp(sUDTName, "f_p057_CRNVar_ContactPayeeId", vbTextCompare) = 0 Then
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
        ElseIf StrComp(sUDTName, "f_p0000_sTexasSubCovCode", vbTextCompare) = 0 Then
            LoadDescriptionList poTextBox, "XML01_TXCOVERAGECODE"
            GoTo CLEAN_UP
        End If
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
    lstvbUserDefinedType.SetFocus
    
    If lstvbUserDefinedType.ListIndex = -1 Then
        oTextBox = vbNullString
    End If
    
CLEAN_UP:
       
    Set oDeser = Nothing
    Set oMyStruct = Nothing
    Set oContactsRS = Nothing
    Set oTextBox = Nothing
    Set itmX = Nothing
    Exit Sub
EH:

    Set oDeser = Nothing
    Set oMyStruct = Nothing
    Set oContactsRS = Nothing
    Set oTextBox = Nothing
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub LoadvbUserDefinedTypeList"
End Sub

Public Sub LoadDescriptionList(poTextBox As Object, psUDTNameKey As String)
    On Error GoTo EH
    Dim sDescList As String
    Dim saryList() As String
    Dim lCount As Long
    Dim sSelItem As String
    Dim lSelIndex As Long
    Dim sLossFormat As String
    
    sLossFormat = goUtil.IsNullIsVbNullString(mfrmClaim.adoRSAssignments.Fields("LRFormat"))
    If StrComp(sLossFormat, "V2ECcarFarmers.clsLossXML01", vbTextCompare) = 0 Then
        sDescList = GetSetting(goUtil.gsMainAppEXEName, "V2ECcarFarmers.clsLossXML01", psUDTNameKey, vbNullString)
    Else
        Exit Sub
    End If
    
    mbLoadingDesc = True
    
    sSelItem = Trim(poTextBox.Text)
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
        poTextBox.Visible = True
    End If
    
    lstvbUserDefinedType.ListIndex = lSelIndex
    
    If lstvbUserDefinedType.ListIndex = -1 Then
        poTextBox = vbNullString
    End If
    
    mbLoadingDesc = False
   
    Exit Sub
EH:
    mbLoadingDesc = False
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub LoadDescriptionList"
End Sub



Private Sub txtParamValue_LostFocus()
    goUtil.utValidate , txtParamValue
End Sub

Private Sub LoadHeaderWordXL()
    On Error GoTo EH
    'set the columnheaders
    With lvwAvail
        .Sorted = True
        .ColumnHeaders.Add , "Name", "Name"
        .ColumnHeaders.Add , "DateLastUpdated", "Date Last Updated"
        .ColumnHeaders.Add , "DateLastUpdatedSort", "Sort Date Last Updated"
        .ColumnHeaders.Add , "DateCreated", "Date Created"
        .ColumnHeaders.Add , "DateCreatedSort", "Sort Date Created"
        
        '"Avail WOrd XL Forms"
        .ColumnHeaders.Item(AvailDocs.Name).Width = 7000
        .ColumnHeaders.Item(AvailDocs.Name).Alignment = lvwColumnLeft
        .ColumnHeaders.Item(AvailDocs.DateLastUpdated).Width = 2500
        .ColumnHeaders.Item(AvailDocs.DateLastUpdated).Alignment = lvwColumnLeft
        .ColumnHeaders.Item(AvailDocs.DateLastUpdatedSort).Width = 0  'Hidden
        .ColumnHeaders.Item(AvailDocs.DateLastUpdatedSort).Alignment = lvwColumnLeft
        .ColumnHeaders.Item(AvailDocs.DateCreated).Width = 2500
        .ColumnHeaders.Item(AvailDocs.DateCreated).Alignment = lvwColumnLeft
        .ColumnHeaders.Item(AvailDocs.DateCreatedSort).Width = 0   'Hidden
        .ColumnHeaders.Item(AvailDocs.DateCreatedSort).Alignment = lvwColumnLeft
       
    End With
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub LoadHeaderWordXL"
End Sub

Public Sub PopulatelvwAvail()
    On Error GoTo EH
    'Source Variables
    Dim varyAvail As Variant
    Dim sAvail As String
    Dim iCount As Integer
    Dim iPic As Integer
    Dim RS As ADODB.Recordset
    Dim itmX As ListItem
    Dim sIBNUM As String
    Dim sTemplatePath As String
    Dim sDestPath As String
    Dim sErrorMess As String
    Dim sMess As String
    Dim sNewAvail As String
    Dim sDateLastUpdated As String
    Dim sDateCreated As String
    Dim oFI As V2ECKeyBoard.clsFileVersion
    Dim myFI As V2ECKeyBoard.FILE_INFORMATION

    lvwAvail.ListItems.Clear
    'BGS 1.2.2002 load the Avail reports
    If Not moWordXL.GetAvail(varyAvail) Then
        sMess = "Template documents are missing for the selected package!"
    Else
        If IsArray(varyAvail) Then
            'Set the RS to Current assignment RS
            Set RS = mfrmClaim.adoRSAssignments
            sIBNUM = RS.Fields("IBNUM").Value
            
            'Set the File info Object
            Set oFI = New V2ECKeyBoard.clsFileVersion
            
            For iCount = LBound(varyAvail) To UBound(varyAvail)
                sAvail = varyAvail(iCount)
                If InStr(1, sAvail, ".xls", vbTextCompare) Then
                    iPic = Pic.XL
                ElseIf InStr(1, sAvail, ".doc", vbTextCompare) Then
                    iPic = Pic.Word
                End If
                'Copy over the Avail template to attach repos for this
                'Assignments using the IB number as prefix.
                sTemplatePath = moWordXL.WordXLDocPath & "\" & sAvail
                sNewAvail = sIBNUM & "_" & sAvail
                sDestPath = goUtil.AttachReposPath & sNewAvail
                'Only copy it over if it  doesn't already exist
                If Not goUtil.utFileExists(sDestPath) Then
                    sErrorMess = goUtil.utCopyFile(sTemplatePath, sDestPath)
                    If sErrorMess <> vbNullString Then
                        sErrorMess = "Error copying document " & vbCrLf & vbCrLf & sErrorMess & vbCrLf
                        sMess = sMess & sErrorMess
                    End If
                Else
                    'If the File Was already there...
                    'Need to get file info on this Doc to get the Date last Modified
                    myFI = oFI.GetFileInformation(sDestPath)
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
                End If
                
                Set itmX = lvwAvail.ListItems.Add(, , sNewAvail, , iPic)
                
                itmX.SubItems(AvailDocs.DateLastUpdated - 1) = Format(sDateLastUpdated, "MM/DD/YYYY HH:MM:SS")
                itmX.SubItems(AvailDocs.DateLastUpdatedSort - 1) = Format(sDateLastUpdated, "YYYY/MM/DD HH:MM:SS")
                itmX.SubItems(AvailDocs.DateCreated - 1) = Format(sDateCreated, "MM/DD/YYYY HH:MM:SS")
                itmX.SubItems(AvailDocs.DateCreatedSort - 1) = Format(sDateCreated, "YYYY/MM/DD HH:MM:SS")
                
            Next
        End If
    End If
    
    If sMess <> vbNullString Then
        MsgBox sMess, vbOKOnly + vbCritical, "Error"
    End If

    'cleanup
    Set itmX = Nothing
    Set RS = Nothing
    Set oFI = Nothing
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub PopulatelvwAvail"
End Sub

Public Function GetVariables() As Variant
    On Error GoTo EH
    Dim MyVariable As V2ECKeyBoard.QVariable
    Dim aryVariables() As V2ECKeyBoard.QVariable
    Dim RS As ADODB.Recordset
    Dim oField As ADODB.Field
    Dim lFCount As Long
    Dim lEndCount As Long
    Dim sADJFName As String
    Dim sADJLName As String
    Dim sADJSSN As String
    Dim sADJEmail As String
    Dim sADJContactPhone As String
    Dim sADJEmergencyPhone As String
    Dim sADJAddress As String
    Dim sADJCity As String
    Dim sADJState As String
    Dim sADJZip As String
    Dim sADJZip4 As String
    Dim sADJOtherPostCode As String
    Dim sADJTeamLeaderSup As String
    'Used these to look up Info from ID keys
    Dim oConn As ADODB.Connection
    Dim MyRS As ADODB.Recordset
    Dim sSQL As String
    
    
    'Load Assignments Table Info
    Set RS = mfrmClaim.adoRSAssignments
    
    'Establish connection for Lookup info
    Set oConn = New ADODB.Connection
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    
    If RS.RecordCount = 1 Then
        RS.MoveFirst
        Erase aryVariables
        lFCount = 0
        ReDim aryVariables(1 To RS.Fields.Count)
        For Each oField In RS.Fields
            lFCount = lFCount + 1
            MyVariable.Name = oField.Name
            'BGS 1.15.2002 If Date is in the Field name then parse it out
            If InStr(1, MyVariable.Name, "Date", vbTextCompare) > 0 Then
                If Not IsNull(goUtil.IsNullIsVbNullString(oField)) Then
                    If Not IsDate(Trim(goUtil.IsNullIsVbNullString(oField))) Then
                        MyVariable.Value = vbNullString
                    Else
                        MyVariable.Value = goUtil.IsNullIsVbNullString(oField)
                    End If
                Else
                    MyVariable.Value = vbNullString
                End If
            ElseIf StrComp(MyVariable.Name, "AssignmentTypeID", vbTextCompare) = 0 Then
                MyVariable.Value = goUtil.IsNullIsVbNullString(oField)
                MyVariable.Name = "AssignmentTypeType"
                If MyVariable.Value <> vbNullString Then
                    'Need to get the Assignment type Name
                    sSQL = "SELECT          [Type] As [AssignmentTypeType] "
                    sSQL = sSQL & "FROM     AssignmentType "
                    sSQL = sSQL & "WHERE    AssignmentType.[AssignmentTypeID] = " & MyVariable.Value & " "
                    Set MyRS = New ADODB.Recordset
                    MyRS.CursorLocation = adUseClient
                    MyRS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
                    Set MyRS.ActiveConnection = Nothing
                    If Not MyRS.EOF Then
                        MyVariable.Value = MyRS.Fields("AssignmentTypeType").Value
                    End If
                End If
            ElseIf StrComp(MyVariable.Name, "ClientCompanyCatSpecID", vbTextCompare) = 0 Then
                MyVariable.Value = goUtil.IsNullIsVbNullString(oField)
                MyVariable.Name = "CatCode"
                If MyVariable.Value <> vbNullString Then
                    'Need to get the CatCode
                    sSQL = "SELECT          [CatCode] "
                    sSQL = sSQL & "FROM     ClientCompanyCatSpec CCCS "
                    sSQL = sSQL & "WHERE    CCCS.[ClientCompanyCatSpecID] = " & MyVariable.Value & " "
                    Set MyRS = New ADODB.Recordset
                    MyRS.CursorLocation = adUseClient
                    MyRS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
                    Set MyRS.ActiveConnection = Nothing
                    If Not MyRS.EOF Then
                        MyVariable.Value = MyRS.Fields("CatCode").Value
                    End If
                End If
            ElseIf StrComp(MyVariable.Name, "AdjusterSpecID", vbTextCompare) = 0 Then
                MyVariable.Value = goUtil.IsNullIsVbNullString(oField)
                MyVariable.Name = "ACID"
                If MyVariable.Value <> vbNullString Then
                    'Need to get the ACID
                    sSQL = "SELECT          [ACID] "
                    sSQL = sSQL & "FROM     ClientCoAdjusterSpec CCAS "
                    sSQL = sSQL & "WHERE    CCAS.[ClientCoAdjusterSpecID] = " & MyVariable.Value & " "
                    Set MyRS = New ADODB.Recordset
                    MyRS.CursorLocation = adUseClient
                    MyRS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
                    Set MyRS.ActiveConnection = Nothing
                    If Not MyRS.EOF Then
                        MyVariable.Value = MyRS.Fields("ACID").Value
                    End If
                End If
            ElseIf StrComp(MyVariable.Name, "AdjusterSpecIDDisplay", vbTextCompare) = 0 Then
                MyVariable.Value = goUtil.IsNullIsVbNullString(oField)
                MyVariable.Name = "ACIDDisplay"
                If MyVariable.Value <> vbNullString Then
                    'Need to get the ACID Display
                    sSQL = "SELECT          [ACID] As [ACIDDisplay] "
                    sSQL = sSQL & "FROM     ClientCoAdjusterSpec CCAS "
                    sSQL = sSQL & "WHERE    CCAS.[ClientCoAdjusterSpecID] = " & MyVariable.Value & " "
                    Set MyRS = New ADODB.Recordset
                    MyRS.CursorLocation = adUseClient
                    MyRS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
                    Set MyRS.ActiveConnection = Nothing
                    If Not MyRS.EOF Then
                        MyVariable.Value = MyRS.Fields("ACIDDisplay").Value
                    End If
                End If
            ElseIf StrComp(MyVariable.Name, "StatusID", vbTextCompare) = 0 Then
                MyVariable.Value = goUtil.IsNullIsVbNullString(oField)
                MyVariable.Name = "StatusStatus"
                If MyVariable.Value <> vbNullString Then
                    'Need to get the Assignment type Name
                    sSQL = "SELECT          [Status] As [StatusStatus]"
                    sSQL = sSQL & "FROM     Status S "
                    sSQL = sSQL & "WHERE    S.[StatusID] = " & MyVariable.Value & " "
                    Set MyRS = New ADODB.Recordset
                    MyRS.CursorLocation = adUseClient
                    MyRS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
                    Set MyRS.ActiveConnection = Nothing
                    If Not MyRS.EOF Then
                        MyVariable.Value = MyRS.Fields("StatusStatus").Value
                    End If
                End If
            ElseIf StrComp(MyVariable.Name, "TypeOfLossID", vbTextCompare) = 0 Then
                MyVariable.Value = goUtil.IsNullIsVbNullString(oField)
                MyVariable.Name = "TypeOfLoss"
                If MyVariable.Value <> vbNullString Then
                    'Need to get the Type Of Loss
                    sSQL = "SELECT          [TypeOfLoss] "
                    sSQL = sSQL & "FROM     TypeOfLoss TOL "
                    sSQL = sSQL & "WHERE    TOL.[TypeOfLossID] = " & MyVariable.Value & " "
                    Set MyRS = New ADODB.Recordset
                    MyRS.CursorLocation = adUseClient
                    MyRS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
                    Set MyRS.ActiveConnection = Nothing
                    If Not MyRS.EOF Then
                        MyVariable.Value = MyRS.Fields("TypeOfLoss").Value
                    End If
                End If
            Else
                MyVariable.Value = Trim(goUtil.IsNullIsVbNullString(oField))
            End If
            
            aryVariables(lFCount) = MyVariable
        Next
    End If
    
    'Load Policy Limits table
    Set RS = mfrmClaim.adoRSPolicyLimits
    If RS Is Nothing Then
        GoTo Load_Vars
    End If
    If RS.RecordCount > 0 Then
        RS.MoveFirst
        Do Until RS.EOF
            If Not CBool(RS.Fields("IsDeleted").Value) Then
                ReDim Preserve aryVariables(1 To UBound(aryVariables, 1) + 1)
            End If
            RS.MoveNext
        Loop
        
        RS.MoveFirst
        
        Do Until RS.EOF
            If Not CBool(RS.Fields("IsDeleted").Value) Then
                lFCount = lFCount + 1
                MyVariable.Name = "COVERAGE_" & RS.Fields("ClassTypeClass").Value
                MyVariable.Value = RS.Fields("LimitAmount").Value
                aryVariables(lFCount) = MyVariable
            End If
            RS.MoveNext
        Loop
        RS.MoveFirst
    End If
Load_Vars:
    
    'Load some Vars from Preferences Screen
    ReDim Preserve aryVariables(1 To lFCount + 13)
    'Adjuster Information
    lFCount = lFCount + 1
    sADJFName = GetSetting(goUtil.gsMainAppEXEName, "GENERAL", "ADJUSTOR_FIRST_NAME", vbNullString)
    MyVariable.Name = "sADJFName"
    MyVariable.Value = sADJFName
    aryVariables(lFCount) = MyVariable
    
    lFCount = lFCount + 1
    sADJLName = GetSetting(goUtil.gsMainAppEXEName, "GENERAL", "ADJUSTOR_LAST_NAME", vbNullString)
    MyVariable.Name = "sADJLName"
    MyVariable.Value = sADJLName
    aryVariables(lFCount) = MyVariable
    
    lFCount = lFCount + 1
    sADJSSN = goUtil.utGetECSCryptSetting("ECS", "WEB_SECURITY", "SSN")
    MyVariable.Name = "sADJSSN"
    MyVariable.Value = sADJSSN
    aryVariables(lFCount) = MyVariable
    
    lFCount = lFCount + 1
    sADJEmail = GetSetting(goUtil.gsMainAppEXEName, "GENERAL", "ADJ_EMAIL", vbNullString)
    MyVariable.Name = "sADJEmail"
    MyVariable.Value = sADJEmail
    aryVariables(lFCount) = MyVariable
    
    lFCount = lFCount + 1
    sADJContactPhone = GetSetting(goUtil.gsMainAppEXEName, "GENERAL", "ADJ_CONTACT_PHONE", vbNullString)
    MyVariable.Name = "sADJContactPhone"
    MyVariable.Value = sADJContactPhone
    aryVariables(lFCount) = MyVariable
    
    lFCount = lFCount + 1
    sADJEmergencyPhone = GetSetting(goUtil.gsMainAppEXEName, "GENERAL", "ADJ_EMERGENCY_PHONE", vbNullString)
    MyVariable.Name = "sADJEmergencyPhone"
    MyVariable.Value = sADJEmergencyPhone
    aryVariables(lFCount) = MyVariable
    
    lFCount = lFCount + 1
    sADJAddress = GetSetting(goUtil.gsMainAppEXEName, "GENERAL", "ADJ_ADDRESS", vbNullString)
    MyVariable.Name = "sADJAddress"
    MyVariable.Value = sADJAddress
    aryVariables(lFCount) = MyVariable
    
    lFCount = lFCount + 1
    sADJCity = GetSetting(goUtil.gsMainAppEXEName, "GENERAL", "ADJ_CITY", vbNullString)
    MyVariable.Name = "sADJCity"
    MyVariable.Value = sADJCity
    aryVariables(lFCount) = MyVariable
    
    lFCount = lFCount + 1
    sADJState = GetSetting(goUtil.gsMainAppEXEName, "GENERAL", "ADJ_STATE", vbNullString)
    MyVariable.Name = "sADJState"
    MyVariable.Value = sADJState
    aryVariables(lFCount) = MyVariable
    
    lFCount = lFCount + 1
    sADJZip = GetSetting(goUtil.gsMainAppEXEName, "GENERAL", "ADJ_ZIP", vbNullString)
    MyVariable.Name = "sADJZip"
    MyVariable.Value = sADJZip
    aryVariables(lFCount) = MyVariable
    
    lFCount = lFCount + 1
    sADJZip4 = GetSetting(goUtil.gsMainAppEXEName, "GENERAL", "ADJ_ZIP4", vbNullString)
    MyVariable.Name = "sADJZip4"
    MyVariable.Value = sADJZip4
    aryVariables(lFCount) = MyVariable
    
    lFCount = lFCount + 1
    sADJOtherPostCode = GetSetting(goUtil.gsMainAppEXEName, "GENERAL", "ADJ_OTHER_POSTCODE", vbNullString)
    MyVariable.Name = "sADJOtherPostCode"
    MyVariable.Value = sADJOtherPostCode
    aryVariables(lFCount) = MyVariable
    
    lFCount = lFCount + 1
    sADJTeamLeaderSup = GetSetting(goUtil.gsMainAppEXEName, "GENERAL", "TEAM_LEADER", vbNullString)
    MyVariable.Name = "sADJTeamLeaderSup"
    MyVariable.Value = sADJTeamLeaderSup
    aryVariables(lFCount) = MyVariable
    
    'Set the function return
    GetVariables = aryVariables
    
'    'Debgug only
'    For lFCount = LBound(aryVariables, 1) To UBound(aryVariables, 1)
'        MyVariable = aryVariables(lFCount)
'        Debug.Print MyVariable.Name & " = " & MyVariable.Value
'    Next
    'End Debug
    
    Set MyRS = Nothing
    Set oConn = Nothing
    Set RS = Nothing
    Set oField = Nothing
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function GetVariables"
End Function

Private Function LoadAvail() As Boolean
    On Error GoTo EH
    Dim itmX As ListItem
    For Each itmX In lvwAvail.ListItems
        If itmX.Selected Then
            imgSelected.Picture = imgVarDoc.ListImages.Item(Pic.Hourglass).Picture
            imgSelected.Refresh
            lblName.Caption = itmX.Text
            lblName.Refresh
            lblDate.Caption = "Loading, please wait..."
            lblDate.Refresh
            
            If moWordXL.LoadaryDocVariables(itmX.SmallIcon, itmX.Text) Then
                lblDate.Caption = "Print ?"
                lblDate.Refresh
                imgSelected.Picture = imgVarDoc.ListImages.Item(itmX.SmallIcon).Picture
                imgSelected.Refresh
            Else
                imgSelected.Picture = imgVarDoc.ListImages.Item(itmX.SmallIcon).Picture
                imgSelected.Refresh
                lblDate.Caption = "No variables found."
            End If
            LoadAvail = True
            cmdPrintDoc.Enabled = True
            Exit For
        End If
    Next
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Function LoadAvail"
End Function

Public Function EditWSDiagram(plNumber As Long, psDiagramXML As String, psDiagramPhotoName As String) As Boolean
    On Error GoTo EH
    Dim sMess As String
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    Dim sID As String
    Dim sDiagramXML As String
    Dim sDiagramPhotoName As String
    Dim sNumber As String
    
    'Set from Passed in params
    sNumber = CStr(plNumber)
    sDiagramXML = psDiagramXML
    sDiagramPhotoName = psDiagramPhotoName
    
    
    Set oConn = New ADODB.Connection
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    Screen.MousePointer = vbHourglass
    
    sSQL = "UPDATE  RTWSDiagram SET  "
    sSQL = sSQL & "[DiagramPhotoName] = '" & goUtil.utCleanSQLString(sDiagramPhotoName) & "', "
    sSQL = sSQL & "[UploadDiagramPhoto] = True, "
    sSQL = sSQL & "[DiagramXML] = '" & goUtil.utCleanSQLString(sDiagramXML) & "', "
    sSQL = sSQL & "[UpLoadMe] = True, "
    sSQL = sSQL & "[DateLastUpdated] = #" & Now() & "# , "
    sSQL = sSQL & "[UpdateByUserID] = " & goUtil.gsCurUsersID & " "
    sSQL = sSQL & "WHERE IDAssignments = " & msAssignmentsID & " "
    sSQL = sSQL & "AND [Number] = " & sNumber & " "

    oConn.Execute sSQL
    
    Sleep 500

    Screen.MousePointer = vbNormal
    
    EditWSDiagram = True
CLEAN_UP:
    'cleanup
    Set oConn = Nothing
    Exit Function
EH:
    Screen.MousePointer = vbNormal
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function EditWSDiagram"
End Function

Public Function ShowDiagramEdit() As Boolean
    On Error GoTo EH
    Dim RS As ADODB.Recordset
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    Dim oECSketch As ECSKETCH.clsECSKETCH
    Dim sWddxXMLIN As String
    Dim sWddxXMLOut As String
    Dim sTempJPGPhotoPathOUT As String
    Dim sDiagramPhotoName As String
    Dim sTimeStamp As String
    Dim sBuildPath As String
    
    
    Set oConn = New ADODB.Connection
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    sSQL = "SELECT "
    sSQL = sSQL & "DiagramXML "
    sSQL = sSQL & "FROM     RTWSDiagram "
    sSQL = sSQL & "WHERE    IDAssignments = " & msAssignmentsID & " "
    sSQL = sSQL & "AND      [Number] = " & CStr(mlEditDiagramNumber) & " "
    
    Set RS = New ADODB.Recordset
    'Use Disconnected Record Set on asUseClient Cusor ONLY !
    RS.CursorLocation = adUseClient
    RS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set RS.ActiveConnection = Nothing
    
    If RS.EOF Then
        ShowDiagramEdit = False
        GoTo CLEAN_UP
    ElseIf RS.RecordCount > 1 Then
        ShowDiagramEdit = False
        GoTo CLEAN_UP
    End If
    
    RS.MoveFirst
    
    'get the Wddx Packet to send to Sketch
    sWddxXMLIN = goUtil.IsNullIsVbNullString(RS.Fields("DiagramXML"))
    
    'Done with these
    Set RS = Nothing
    Set oConn = Nothing
    
    Set oECSketch = New ECSKETCH.clsECSKETCH
    
    oECSketch.WddxXml = sWddxXMLIN
    
    'This will show modal
    oECSketch.ShowSketch
    
    'After Modal Show Will Get JPG Path and Updated Wddx
    sWddxXMLOut = oECSketch.WddxXml
    sTempJPGPhotoPathOUT = oECSketch.myJPGPath
    
    If goUtil.utFileExists(sTempJPGPhotoPathOUT) Then
        DoEvents
        Sleep 200
        sTimeStamp = Format(Now, "YYMMDDHHMMSS")
        sDiagramPhotoName = msIBNUM & "_" & sTimeStamp & ".jpg"
        sBuildPath = goUtil.PhotoReposPath & sDiagramPhotoName
    Else
        sDiagramPhotoName = vbNullString
    End If
    
    If oECSketch.Save Then
        If EditWSDiagram(mlEditDiagramNumber, sWddxXMLOut, sDiagramPhotoName) Then
            If sDiagramPhotoName <> vbNullString Then
                If goUtil.utCopyFile(sTempJPGPhotoPathOUT, sBuildPath) = vbNullString Then
                    imgDiagram.Picture = LoadPicture(sBuildPath)
                End If
                goUtil.utDeleteFile sTempJPGPhotoPathOUT
            Else
                imgDiagram.Picture = LoadPicture()
            End If
        End If
    End If
    
        
    ShowDiagramEdit = True
CLEAN_UP:
    
    Set RS = Nothing
    Set oConn = Nothing
    Set oECSketch = Nothing
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function ShowDiagramEdit"
End Function
