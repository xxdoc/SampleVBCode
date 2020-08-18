VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFixXactProjects 
   AutoRedraw      =   -1  'True
   Caption         =   "Send To Xactimate"
   ClientHeight    =   6585
   ClientLeft      =   60
   ClientTop       =   345
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
   Icon            =   "frmFixXactProjects.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   9480
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame framMain 
      Appearance      =   0  'Flat
      Caption         =   "Assignments (&Xactimate Projects)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9255
      Begin VB.CommandButton cmdGetFromExportPath 
         Height          =   330
         Left            =   8760
         Picture         =   "frmFixXactProjects.frx":1B42
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Browse"
         Top             =   255
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtGetFromExportPath 
         Height          =   360
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   240
         Visible         =   0   'False
         Width           =   2880
      End
      Begin VB.CheckBox chkGetFromExport 
         Caption         =   "Get From Export"
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
         Left            =   5280
         TabIndex        =   5
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdSendToExportPath 
         Height          =   330
         Left            =   4800
         Picture         =   "frmFixXactProjects.frx":1FBC
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Browse"
         Top             =   255
         Width           =   375
      End
      Begin VB.TextBox txtSendToExportPath 
         Height          =   360
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   2880
      End
      Begin VB.CheckBox chkSendToExport 
         Caption         =   "Send To Export"
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
         Left            =   1320
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdSelAll 
         Caption         =   "&Select A&ll"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1100
      End
      Begin VB.Timer Timer_Resize 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   8520
         Top             =   960
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Find"
         Height          =   375
         Left            =   1200
         TabIndex        =   11
         Top             =   2680
         Width           =   1100
      End
      Begin VB.CommandButton cmdFindNext 
         Caption         =   "Find &Next"
         Height          =   375
         Left            =   2400
         TabIndex        =   12
         Top             =   2680
         Width           =   1100
      End
      Begin VB.CommandButton cmdViewFailed 
         Caption         =   "Vie&w Failed"
         Height          =   375
         Left            =   3600
         TabIndex        =   13
         Top             =   2680
         Width           =   1455
      End
      Begin VB.CheckBox chkShowGrid 
         Caption         =   "Show Grid"
         Height          =   240
         Left            =   7920
         TabIndex        =   15
         Top             =   2680
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CommandButton cmdDown 
         Height          =   375
         Left            =   660
         Picture         =   "frmFixXactProjects.frx":2436
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Next Record"
         Top             =   2680
         Width           =   400
      End
      Begin VB.CommandButton cmdUp 
         Height          =   375
         Left            =   120
         Picture         =   "frmFixXactProjects.frx":2878
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Previous Record"
         Top             =   2680
         Width           =   400
      End
      Begin VB.CheckBox chkBrowseBadData 
         Caption         =   "Find Failed Validation Only"
         Height          =   240
         Left            =   5160
         TabIndex        =   14
         Top             =   2680
         Value           =   1  'Checked
         Width           =   2655
      End
      Begin MSComctlLib.ImageList imgFixXact 
         Left            =   8400
         Top             =   1560
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFixXactProjects.frx":2CBA
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFixXactProjects.frx":310E
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFixXactProjects.frx":3562
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFixXactProjects.frx":39B6
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFixXactProjects.frx":3E0A
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lvwXactProjects 
         Height          =   2055
         Left            =   120
         TabIndex        =   8
         Tag             =   "Enable"
         Top             =   600
         Width           =   8985
         _ExtentX        =   15849
         _ExtentY        =   3625
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "imgFixXact"
         SmallIcons      =   "imgFixXact"
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
   End
   Begin VB.Frame framCommands 
      Appearance      =   0  'Flat
      Caption         =   "Commands"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   3360
      TabIndex        =   20
      Top             =   3240
      Width           =   6015
      Begin VB.CommandButton cmdPrintList 
         Caption         =   "Sho&w Item List"
         Height          =   855
         Left            =   3840
         MaskColor       =   &H00000000&
         Picture         =   "frmFixXactProjects.frx":425E
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Exit"
         Top             =   360
         Width           =   975
      End
      Begin VB.ComboBox cboXactVS 
         Height          =   360
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   420
         Width           =   2535
      End
      Begin VB.ComboBox cboXactSpeed 
         Height          =   360
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   790
         Width           =   2535
      End
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "&Exit"
         Height          =   855
         Left            =   4920
         MaskColor       =   &H00000000&
         Picture         =   "frmFixXactProjects.frx":46A0
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Exit"
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdSendAll 
         Caption         =   "Send &All"
         Default         =   -1  'True
         Enabled         =   0   'False
         Height          =   855
         Left            =   120
         Picture         =   "frmFixXactProjects.frx":4922
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Send All"
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblXactVS 
         Caption         =   "Xactimate Settings"
         Height          =   255
         Left            =   1200
         TabIndex        =   22
         Top             =   165
         Width           =   2175
      End
   End
   Begin VB.Frame framSendStatus 
      Appearance      =   0  'Flat
      Caption         =   "Send to Xact Status"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   120
      TabIndex        =   16
      Top             =   3240
      Width           =   3135
      Begin VB.OptionButton optSent 
         Caption         =   "Sent"
         Enabled         =   0   'False
         Height          =   855
         Left            =   2160
         Picture         =   "frmFixXactProjects.frx":5564
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   360
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.OptionButton optSkip 
         Caption         =   "Ski&p"
         Height          =   855
         Left            =   1140
         Picture         =   "frmFixXactProjects.frx":59A6
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Mark to Skip"
         Top             =   360
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.OptionButton optSend 
         Caption         =   "&Send"
         Height          =   855
         Left            =   120
         Picture         =   "frmFixXactProjects.frx":5DE8
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Mark to Send"
         Top             =   360
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin VB.Frame framTol 
      Appearance      =   0  'Flat
      Caption         =   "Type of Loss"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   6120
      TabIndex        =   39
      Top             =   4680
      Width           =   3255
      Begin VB.TextBox txtCatCode 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   360
         Left            =   2670
         MaxLength       =   2
         TabIndex        =   44
         Top             =   1320
         Width           =   465
      End
      Begin VB.ComboBox cboTOL 
         Height          =   360
         Left            =   1560
         TabIndex        =   43
         Top             =   1320
         Width           =   1155
      End
      Begin VB.TextBox txtDeductible 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   42
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtPolicyNumber 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1560
         MaxLength       =   40
         TabIndex        =   41
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtClaimNumber 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1560
         MaxLength       =   40
         TabIndex        =   40
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblFixXact 
         Caption         =   "TOL/Cat Code"
         Height          =   255
         Index           =   11
         Left            =   60
         TabIndex        =   56
         Top             =   1425
         Width           =   3015
      End
      Begin VB.Label lblFixXact 
         Caption         =   "Policy Number"
         Height          =   255
         Index           =   9
         Left            =   60
         TabIndex        =   55
         Top             =   720
         Width           =   3015
      End
      Begin VB.Label lblFixXact 
         Caption         =   "Deductible"
         Height          =   255
         Index           =   10
         Left            =   60
         TabIndex        =   54
         Top             =   1080
         Width           =   3015
      End
      Begin VB.Label lblFixXact 
         Caption         =   "Claim Number"
         Height          =   255
         Index           =   8
         Left            =   60
         TabIndex        =   53
         Top             =   360
         Width           =   3015
      End
   End
   Begin VB.Frame framStatus 
      Appearance      =   0  'Flat
      Caption         =   "Status (Dates)"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   3360
      TabIndex        =   34
      Top             =   4680
      Width           =   2655
      Begin VB.TextBox txtDateEntered 
         Alignment       =   2  'Center
         Height          =   360
         Left            =   1320
         MaxLength       =   40
         TabIndex        =   38
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txtDateInspected 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   1320
         MaxLength       =   40
         TabIndex        =   37
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtDateReceived 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   1320
         MaxLength       =   40
         TabIndex        =   36
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtDateofLoss 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   1320
         MaxLength       =   40
         TabIndex        =   35
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblFixXact 
         Caption         =   "Loss"
         Height          =   255
         Index           =   4
         Left            =   60
         TabIndex        =   52
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label lblFixXact 
         Caption         =   "Entered"
         Height          =   255
         Index           =   7
         Left            =   60
         TabIndex        =   51
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label lblFixXact 
         Caption         =   "Inspected"
         Height          =   255
         Index           =   6
         Left            =   60
         TabIndex        =   50
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label lblFixXact 
         Caption         =   "Received"
         Height          =   255
         Index           =   5
         Left            =   60
         TabIndex        =   49
         Top             =   720
         Width           =   2415
      End
   End
   Begin VB.Frame framInsured 
      Appearance      =   0  'Flat
      Caption         =   "Insured Information"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   120
      TabIndex        =   27
      Top             =   4680
      Width           =   3135
      Begin VB.TextBox txtZip5 
         Alignment       =   2  'Center
         Height          =   360
         Left            =   1575
         MaxLength       =   5
         TabIndex        =   32
         Top             =   1320
         Width           =   675
      End
      Begin VB.ComboBox cboState 
         Height          =   360
         Left            =   840
         TabIndex        =   31
         Top             =   1320
         Width           =   795
      End
      Begin VB.TextBox txtZip4 
         Alignment       =   2  'Center
         Height          =   360
         Left            =   2475
         MaxLength       =   4
         TabIndex        =   33
         Top             =   1320
         Width           =   540
      End
      Begin VB.TextBox txtCity 
         Height          =   375
         Left            =   840
         MaxLength       =   30
         TabIndex        =   30
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox txtStreet 
         Height          =   375
         Left            =   840
         MaxLength       =   40
         TabIndex        =   29
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox txtName 
         Height          =   375
         Left            =   840
         MaxLength       =   40
         TabIndex        =   28
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Caption         =   "-"
         Height          =   255
         Left            =   2300
         TabIndex        =   57
         Top             =   1350
         Width           =   135
      End
      Begin VB.Label lblFixXact 
         Caption         =   "Street"
         Height          =   255
         Index           =   1
         Left            =   60
         TabIndex        =   48
         Top             =   720
         Width           =   2895
      End
      Begin VB.Label lblFixXact 
         Caption         =   "St/Zip"
         Height          =   255
         Index           =   3
         Left            =   60
         TabIndex        =   47
         Top             =   1440
         Width           =   2895
      End
      Begin VB.Label lblFixXact 
         Caption         =   "City"
         Height          =   255
         Index           =   2
         Left            =   60
         TabIndex        =   46
         Top             =   1080
         Width           =   2895
      End
      Begin VB.Label lblFixXact 
         Caption         =   "Name"
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   45
         Top             =   360
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmFixXactProjects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const RED_BG As Long = &HC0C0FF
Private Const WHITE_BG As Long = &H80000005

'Control Size in Ref to Form Diff Constants
Private Const FORM_W As Long = 9600
Private Const FORM_H As Long = 6990
Private Const framMain_H As Long = 3855
Private Const framMain_W As Long = 345
'Use Actual Base Left for Export controls
Private Const txtSendToExportPath_W As Long = 2880
Private Const cmdSendToExportPath_L As Long = 4800
Private Const chkGetFromExport_L As Long = 5280
Private Const txtGetFromExportPath_L As Long = 6240
Private Const txtGetFromExportPath_W As Long = 2880
Private Const cmdGetFromExportPath_L As Long = 8760
Private Const lvwXactProjects_H As Long = 4935
Private Const lvwXactProjects_W As Long = 615
Private Const cmdUp_T As Long = 4310
Private Const cmdDown_T As Long = 4310
Private Const cmdFind_T As Long = 4310
Private Const cmdFindNext_T As Long = 4310
Private Const cmdViewFailed_T As Long = 4310
Private Const chkBrowseBadData_T As Long = 4310
Private Const chkShowGrid_T As Long = 4310
'Fram Send Status
Private Const framSendStatus_T As Long = 3750
'framCommands
Private Const framCommands_T As Long = 3750
Private Const framCommands_W As Long = 3585
Private Const cmdPrintList_L As Long = 5760
Private Const cmdExit_L As Long = 4680
'framInsured
Private Const framInsured_T As Long = 2310
Private Const framInsured_W As Long = 6465
Private Const txtName_W As Long = 7425
Private Const txtStreet_W As Long = 7425
Private Const txtCity_W As Long = 7425
'framStatus
Private Const framStatus_T As Long = 2310
Private Const framStatus_L As Long = 6240
'framTol
Private Const framTol_T As Long = 2310
Private Const framTol_L As Long = 3480


Private mbResize As Boolean
Private mitmX As listItem
Private mXProj As udtXactProject
Private mcolXactProjects As Collection
Private mcolValidXProjects As Collection
Private mbProjectsValidated As Boolean
Private moXact As clsXact
Private mbAlreadySent As Boolean
Private mbLoadingEdit As Boolean
Private mbSkipAll As Boolean
Private mbModalFlag As Boolean
Private msFindText As String
Private mlLastFindIndex As Long

Private Property Get msClassName() As String
    msClassName = Me.Name
End Property

Public Property Let ModalFlag(pbFlag As Boolean)
    mbModalFlag = pbFlag
End Property
Public Property Get ModalFlag() As Boolean
    ModalFlag = mbModalFlag
End Property

Public Property Get SkipAll() As Boolean
    SkipAll = mbSkipAll
End Property
Public Property Get ProjectsValidated() As Boolean
    ProjectsValidated = mbProjectsValidated
End Property

Public Property Let XactProjects(pcolXactProjects As Collection)
    Set mcolXactProjects = pcolXactProjects
End Property
Public Property Set XactProjects(pcolXactProjects As Collection)
    Set mcolXactProjects = pcolXactProjects
End Property
Public Property Get XactProjects() As Collection
    Set XactProjects = mcolXactProjects
End Property
Public Property Let Xact(poXACT As clsXact)
    Set moXact = poXACT
End Property
Public Property Set Xact(poXACT As clsXact)
    Set moXact = poXACT
End Property

Private Sub LoadHeaderXactProjects()
    On Error GoTo EH
    Dim bGridOn As Boolean
    
    'set the columnheaders
    With lvwXactProjects
      
        .ColumnHeaders.Add , "Status", "Status"
        .ColumnHeaders.Add , "Data", "Data"
        .ColumnHeaders.Add , "Name", "Name"
        .ColumnHeaders.Add , "Street", "Street"
        .ColumnHeaders.Add , "StreetSort", "StreetSort"
        .ColumnHeaders.Add , "City", "City"
        .ColumnHeaders.Add , "State", "State"
        .ColumnHeaders.Add , "Zip", "Zip"
        .ColumnHeaders.Add , "Zip4", "Zip4"
        .ColumnHeaders.Add , "DateOfLoss", "Date Of Loss"
        .ColumnHeaders.Add , "DateOfLossSort", "DateOfLossSort"
        .ColumnHeaders.Add , "DateReceived", "Date Received"
        .ColumnHeaders.Add , "DateReceivedSort", "DateReceivedSort"
        .ColumnHeaders.Add , "DateInspected", "Date Inspected"
        .ColumnHeaders.Add , "DateInspectedSort", "DateInspectedSort"
        .ColumnHeaders.Add , "DateEntered", "Date Entered"
        .ColumnHeaders.Add , "DateEnteredSort", "DateEnteredSort"
        .ColumnHeaders.Add , "ClaimNumber", "Claim Number"
        .ColumnHeaders.Add , "ClaimNumberSort", "ClaimNumberSort"
        .ColumnHeaders.Add , "PolicyNumber", "Policy Number"
        .ColumnHeaders.Add , "PolicyNumberSort", "PolicyNumberSort"
        .ColumnHeaders.Add , "TypeOfLoss", "Type Of Loss"
        .ColumnHeaders.Add , "Deductible", "Deductible"
        .ColumnHeaders.Add , "DeductibleSort", "DeductibleSort"
        .ColumnHeaders.Add , "CatCode", "Cat Code"
        .ColumnHeaders.Add , "CatCodeSort", "CatCodeSort"
        .Sorted = False
        
        .ColumnHeaders.Item(FixXactProj.Status).Width = 800
        .ColumnHeaders.Item(FixXactProj.Status).Alignment = lvwColumnLeft
        .ColumnHeaders.Item(FixXactProj.Data).Width = 600
        .ColumnHeaders.Item(FixXactProj.Data).Alignment = lvwColumnLeft
        .ColumnHeaders.Item(FixXactProj.Name).Width = 1200
        .ColumnHeaders.Item(FixXactProj.Name).Alignment = lvwColumnLeft
        .ColumnHeaders.Item(FixXactProj.Street).Width = 1200
        .ColumnHeaders.Item(FixXactProj.Street).Alignment = lvwColumnLeft
        .ColumnHeaders.Item(FixXactProj.StreetSort).Width = 0
        .ColumnHeaders.Item(FixXactProj.StreetSort).Alignment = lvwColumnLeft
        .ColumnHeaders.Item(FixXactProj.City).Width = 1200
        .ColumnHeaders.Item(FixXactProj.City).Alignment = lvwColumnLeft
        .ColumnHeaders.Item(FixXactProj.State).Width = 700
        .ColumnHeaders.Item(FixXactProj.State).Alignment = lvwColumnLeft
        .ColumnHeaders.Item(FixXactProj.Zip).Width = 700
        .ColumnHeaders.Item(FixXactProj.Zip).Alignment = lvwColumnLeft
        .ColumnHeaders.Item(FixXactProj.Zip4).Width = 700
        .ColumnHeaders.Item(FixXactProj.Zip4).Alignment = lvwColumnLeft
        .ColumnHeaders.Item(FixXactProj.DateOfLoss).Width = 1200
        .ColumnHeaders.Item(FixXactProj.DateOfLoss).Alignment = lvwColumnLeft
        .ColumnHeaders.Item(FixXactProj.DateOfLossSort).Width = 0
        .ColumnHeaders.Item(FixXactProj.DateOfLossSort).Alignment = lvwColumnLeft
        .ColumnHeaders.Item(FixXactProj.DateReceived).Width = 1200
        .ColumnHeaders.Item(FixXactProj.DateReceived).Alignment = lvwColumnLeft
        .ColumnHeaders.Item(FixXactProj.DateReceivedSort).Width = 0
        .ColumnHeaders.Item(FixXactProj.DateReceivedSort).Alignment = lvwColumnLeft
        .ColumnHeaders.Item(FixXactProj.DateInspected).Width = 1200
        .ColumnHeaders.Item(FixXactProj.DateInspected).Alignment = lvwColumnLeft
        .ColumnHeaders.Item(FixXactProj.DateInspectedSort).Width = 0
        .ColumnHeaders.Item(FixXactProj.DateInspectedSort).Alignment = lvwColumnLeft
        .ColumnHeaders.Item(FixXactProj.DateEntered).Width = 1200
        .ColumnHeaders.Item(FixXactProj.DateEntered).Alignment = lvwColumnLeft
        .ColumnHeaders.Item(FixXactProj.DateEnteredSort).Width = 0
        .ColumnHeaders.Item(FixXactProj.DateEnteredSort).Alignment = lvwColumnLeft
        .ColumnHeaders.Item(FixXactProj.ClaimNumber).Width = 1200
        .ColumnHeaders.Item(FixXactProj.ClaimNumber).Alignment = lvwColumnLeft
        .ColumnHeaders.Item(FixXactProj.ClaimNumberSort).Width = 0
        .ColumnHeaders.Item(FixXactProj.ClaimNumberSort).Alignment = lvwColumnLeft
        .ColumnHeaders.Item(FixXactProj.PolicyNumber).Width = 1200
        .ColumnHeaders.Item(FixXactProj.PolicyNumber).Alignment = lvwColumnLeft
        .ColumnHeaders.Item(FixXactProj.PolicyNumberSort).Width = 0
        .ColumnHeaders.Item(FixXactProj.PolicyNumberSort).Alignment = lvwColumnLeft
        .ColumnHeaders.Item(FixXactProj.TypeOfLoss).Width = 1200
        .ColumnHeaders.Item(FixXactProj.TypeOfLoss).Alignment = lvwColumnLeft
        .ColumnHeaders.Item(FixXactProj.Deductible).Width = 1200
        .ColumnHeaders.Item(FixXactProj.Deductible).Alignment = lvwColumnLeft
        .ColumnHeaders.Item(FixXactProj.DeductibleSort).Width = 0
        .ColumnHeaders.Item(FixXactProj.DeductibleSort).Alignment = lvwColumnLeft
        .ColumnHeaders.Item(FixXactProj.CatCode).Width = 800
        .ColumnHeaders.Item(FixXactProj.CatCode).Alignment = lvwColumnLeft
        .ColumnHeaders.Item(FixXactProj.CatCodeSort).Width = 0
        .ColumnHeaders.Item(FixXactProj.CatCodeSort).Alignment = lvwColumnLeft
    End With
    
    bGridOn = CBool(GetSetting(App.EXEName, "GENERAL", "GRID_ON", False))
    If bGridOn Then
        chkShowGrid.Value = vbChecked
    Else
        chkShowGrid.Value = vbUnchecked
    End If
    
    lvwXactProjects.Gridlines = bGridOn
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub LoadHeaderXactProjects"
End Sub

Private Sub PopulatelvwXactProjects()
    On Error GoTo EH
    'Source Variables
    Dim vXProj As Variant   'udtXactProject '
    Dim XProj As V2ECKeyBoard.udtXactProject
    Dim itmX As listItem
    Dim iStatus As Integer
    Dim sStatus As String
    Dim resetcolXactProjects As Collection
    lvwXactProjects.ListItems.Clear

    If Not mcolXactProjects Is Nothing Then
        Set resetcolXactProjects = New Collection
        For Each vXProj In mcolXactProjects
            XProj = vXProj
            If XProj.SentToXact Then
                iStatus = FixXactPic.AlreadySentToXact
                sStatus = "Sent"
            Else
                iStatus = FixXactPic.NeedsSentToXact
                sStatus = "Send"
            End If
                
            Set itmX = lvwXactProjects.ListItems.Add(, """" & CStr(XProj.Loss.ClaimNumber) & """", sStatus, , iStatus)
            moXact.ValidateXProject XProj, itmX
            itmX.Selected = False
            resetcolXactProjects.Add XProj, XProj.Loss.ClaimNumber
        Next
        Set mcolXactProjects = Nothing
        Set mcolXactProjects = resetcolXactProjects
        EnableSend
    End If
    
CLEANUP:
    'Cleanup

    Set itmX = Nothing
    Set resetcolXactProjects = Nothing
    Exit Sub
EH:
    Set itmX = Nothing
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub PopulatelvwXactProjects"
End Sub

Private Sub cboState_Change()
    On Error GoTo EH
    Dim lPos As Long
    lPos = cboState.SelStart
    cboState.Text = UCase(cboState.Text)
    cboState.SelStart = lPos
    ValidateControl cboState, State, True
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cboState_Change"
End Sub

Private Sub cboState_Click()
    On Error GoTo EH
    Dim lPos As Long
    lPos = cboState.SelStart
    cboState.Text = UCase(cboState.Text)
    cboState.SelStart = lPos
    ValidateControl cboState, State, True
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cboState_Click"
End Sub

Private Sub cboState_LostFocus()
    LoadEdit
End Sub

Private Sub cboTOL_Change()
    ValidateControl cboTOL, TypeOfLoss, False
End Sub

Private Sub cboTOL_Click()
    ValidateControl cboTOL, TypeOfLoss, False
End Sub

Private Sub cboTOL_LostFocus()
    LoadEdit
End Sub

Private Sub cboXactSpeed_Click()
    On Error GoTo EH
    
    SaveSetting "ECS", "KEYBOARD", "SPEED", cboXactSpeed.Text
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cboXactSpeed_Click"
End Sub

Private Sub cboXactVS_Click()
    On Error GoTo EH
    
    SaveSetting "ECS", "XACTIMATE", "Version", cboXactVS.Text
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cboXactVS_Click"
End Sub



Private Sub chkSendToExport_Click()
    On Error GoTo EH
    
    If chkSendToExport.Value = vbChecked Then
        moXact.SendToExport = True
        moXact.ExportFilePath = txtSendToExportPath.Text
    Else
        moXact.SendToExport = False
        moXact.ExportFilePath = vbNullString
    End If
    
    EnableSend
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub chkSendToExport_Click"
End Sub

Private Sub chkShowGrid_Click()
    On Error GoTo EH
    Dim bGridOn As Boolean
    
    If chkShowGrid.Value = vbChecked Then
        chkShowGrid.Caption = "&Grid ON"
        bGridOn = True
    Else
        chkShowGrid.Caption = "&Grid OFF"
        bGridOn = False
    End If
    
    SaveSetting App.EXEName, "GENERAL", "GRID_ON", bGridOn
    lvwXactProjects.Gridlines = bGridOn
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub chkShowGrid_Click"
End Sub

Private Sub cmdDown_Click()
    On Error Resume Next
    Dim lCount As Long
    If mitmX Is Nothing Then
        Exit Sub
    End If
    lCount = mitmX.Index + 1
    Do Until lCount > lvwXactProjects.ListItems.Count
        Set mitmX = lvwXactProjects.ListItems(lCount)
        If chkBrowseBadData.Value = vbChecked Then
            If mitmX.ListSubItems(FixXactProj.Data - 1).ReportIcon = FixXactPic.BadData Then
                If mitmX.SmallIcon <> FixXactPic.Skip Then
                    mitmX.Selected = True
                    mitmX.EnsureVisible
                    LoadEdit
                    Exit Do
                End If
            End If
        Else
            mitmX.Selected = True
            mitmX.EnsureVisible
            LoadEdit
            Exit Do
        End If
        lCount = lCount + 1
    Loop
End Sub

Private Sub cmdExit_Click()
    On Error GoTo EH
    Dim sMess As String
    Dim itmX As listItem
    
    'See if All projects were corrected
    'If not then Give Chance not to Exit.
    'If not All Projects were fixed non of the Projects the
    'user made changes to will be saved...
    For Each itmX In lvwXactProjects.ListItems
        If itmX.ListSubItems(FixXactProj.Data - 1).ReportIcon = FixXactPic.BadData Then
            mbProjectsValidated = False
            Exit For
        End If
    Next
    If Not mbProjectsValidated Then
        sMess = "Are you sure you want to Exit?" & vbCrLf & vbCrLf
        sMess = sMess & "All projects must validate before changes made to " & vbCrLf
        sMess = sMess & "any one project will be saved." & vbCrLf & vbCrLf
        sMess = sMess & "Press ""Yes"" to Exit without saving changes." & vbCrLf
        sMess = sMess & "Press ""No"" to NOT EXIT!"
        If MsgBox(sMess, vbYesNo + vbExclamation, "Are You Sure?") = vbNo Then
            Exit Sub
        End If
    End If
    moXact.XactProjects = mcolXactProjects
    mbSkipAll = True
    
    LoadValidXProjects
    If Not mcolXactProjects Is Nothing Then
        moXact.CheckIsDirtyAndAdd mcolValidXProjects
    End If
   
    Me.Hide
    Set itmX = Nothing
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdExit_Click"
    Me.Hide
End Sub

Private Sub cmdFind_Click()
    On Error GoTo EH
    If lvwXactProjects.ListItems.Count > 0 Then
        mlLastFindIndex = 0
        goUtil.utFindListItem Me, lvwXactProjects, msFindText, mlLastFindIndex
    End If
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdFind_Click"
End Sub

Private Sub cmdFindNext_Click()
    On Error GoTo EH
    
    If msFindText <> vbNullString And lvwXactProjects.ListItems.Count > 0 Then
        goUtil.utFindListItem Me, lvwXactProjects, msFindText, mlLastFindIndex
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdFindNext_Click"
End Sub

Private Sub cmdGetFromExportPath_Click()
    On Error GoTo EH
    Dim sPath As String
    Dim sSelFile As String
    Dim lFileSize As Long
    Dim sMyFilter As String
    
    sMyFilter = sMyFilter & "Xactimate Export File" & " (*." & "xef" & ")" & SD & "*." & "xef" & SD
   
    
    sPath = goUtil.utGetPath(App.EXEName, "XactExportFile", "Browse to the Xactimate Export File you want to use", "", sPath, Me.hWnd, sMyFilter, sSelFile)
    
    If FileExists(sPath & sSelFile) Then
        txtGetFromExportPath.Text = sPath & sSelFile
        chkGetFromExport.Value = vbChecked
        moXact.GetFromExport = True
        moXact.ExportFilePath = txtGetFromExportPath.Text
        moXact.PopulateFromXactExport
        PopulatelvwXactProjects
    Else
        txtGetFromExportPath.Text = vbNullString
        chkGetFromExport.Enabled = False
        chkGetFromExport.Value = vbUnchecked
        Set mcolXactProjects = Nothing
        Set moXact.XactProjects = Nothing
        lvwXactProjects.ListItems.Clear
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdGetFromExportPath_Click"
End Sub


Private Sub cmdPrintList_Click()
    On Error GoTo EH

    goUtil.utPrintListView goUtil.gsAppEXEName, lvwXactProjects, "Xactimate Claims List", ddOPortrait, vbModal, 0, True
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdPrintList_Click"
End Sub

Private Sub cmdSelAll_Click()
    On Error GoTo EH
    Dim itmX As listItem
    For Each itmX In lvwXactProjects.ListItems
        itmX.Selected = True
    Next
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdSelAll_Click"
End Sub

Private Sub cmdSendAll_Click()
    On Error GoTo EH
    moXact.XactProjects = mcolXactProjects
    mbSkipAll = False
    LoadValidXProjects
    moXact.CheckIsDirtyAndAdd mcolValidXProjects
    Me.Hide
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdSendAll_Click"
End Sub

Private Sub cmdSendToExportPath_Click()
    On Error GoTo EH
    Dim sPath As String
    Dim sSelFile As String
    Dim lFileSize As Long
    Dim sMyFilter As String

    sMyFilter = sMyFilter & "Xactimate Export File" & " (*." & "xef" & ")" & SD & "*." & "xef" & SD
   
    
    sPath = goUtil.utGetSavePath(App.EXEName, "XactExportFile", "Browse to a folder and enter a Name for this Xactimate Export File", "", sPath, Me.hWnd, sMyFilter, sSelFile)
    
    If FileExists(sPath, True) Then
        txtSendToExportPath.Text = sPath & sSelFile & ".xef"
        chkSendToExport.Enabled = True
        chkSendToExport.Value = vbChecked
    Else
        txtSendToExportPath.Text = vbNullString
        chkSendToExport.Enabled = False
        chkSendToExport.Value = vbUnchecked
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdSendToExportPath_Click"
End Sub

Private Sub cmdUp_Click()
    On Error Resume Next
    Dim lCount As Long
    If mitmX Is Nothing Then
        Exit Sub
    End If
    lCount = mitmX.Index - 1
    Do Until lCount <= 0
        Set mitmX = lvwXactProjects.ListItems(lCount)
        If chkBrowseBadData.Value = vbChecked Then
            If mitmX.ListSubItems(FixXactProj.Data - 1).ReportIcon = FixXactPic.BadData Then
                If mitmX.SmallIcon <> FixXactPic.Skip Then
                    mitmX.Selected = True
                    mitmX.EnsureVisible
                    LoadEdit
                    Exit Do
                End If
            End If
        Else
            mitmX.Selected = True
            mitmX.EnsureVisible
            LoadEdit
            Exit Do
        End If
        lCount = lCount - 1
    Loop
End Sub

Private Sub cmdViewFailed_Click()
    On Error GoTo EH
    Dim itmX As listItem
    Dim lCount As Long
    
    
    'Select Alll the Failed Items in the List View
    If lvwXactProjects.ListItems.Count > 0 Then
        For lCount = 1 To lvwXactProjects.ListItems.Count
            Set itmX = lvwXactProjects.ListItems(lCount)
            If itmX.ListSubItems(FixXactProj.Data - 1).ReportIcon = FixXactPic.BadData Then
                itmX.Selected = True
            End If
        Next
    End If
    
    Set itmX = Nothing
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdViewFailed_Click"
End Sub


'Private Sub cmdUp_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyLeft Then
'        cmdUp_Click
'    End If
'End Sub

Private Sub Form_Load()
    On Error GoTo EH
    goUtil.utFormWinRegPos goUtil.gsMainAppEXEName, Me
    
    LoadHeaderXactProjects
    PopulatelvwXactProjects
    PopulateLookupCbo
    'Issue 224  9.23.2002 Send to Xactimate no work with Xactimate 2002
    PopulateXactSettings
    
    goUtil.utSuffixLabels lblFixXact
    
    'Check for Get From Export
    If moXact.GetFromExport Then
        chkGetFromExport.Visible = True
        txtGetFromExportPath.Visible = True
        cmdGetFromExportPath.Visible = True
        txtGetFromExportPath.Text = moXact.ExportFilePath
'        chkSendToExport.Visible = False
'        txtSendToExportPath.Visible = False
'        cmdSendToExportPath.Visible = False
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_Load"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo EH
    Dim sMess As String
    Dim itmX As listItem
    For Each itmX In lvwXactProjects.ListItems
        If itmX.ListSubItems(FixXactProj.Data - 1).ReportIcon = FixXactPic.BadData Then
            mbProjectsValidated = False
            Exit For
        End If
    Next
    'Hide instead of allowing unload if user closes window
    If UnloadMode = vbFormControlMenu Then
        If Not mbProjectsValidated Then
            sMess = "Are you sure you want to Exit?" & vbCrLf & vbCrLf
            sMess = sMess & "All projects must validate before changes made to " & vbCrLf
            sMess = sMess & "any one project will be saved." & vbCrLf & vbCrLf
            sMess = sMess & "Press ""Yes"" to Exit without saving changes." & vbCrLf
            sMess = sMess & "Press ""No"" to NOT EXIT!"
            If MsgBox(sMess, vbYesNo + vbExclamation, "Are You Sure?") = vbNo Then
                Cancel = True
                Exit Sub
            End If
        End If
        Cancel = True
        moXact.XactProjects = mcolXactProjects
        mbSkipAll = True
        LoadValidXProjects
        moXact.CheckIsDirtyAndAdd mcolValidXProjects
        Me.Hide
    End If
    Set itmX = Nothing
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_QueryUnload"
    Me.Hide
End Sub

Private Sub Form_Resize()
    On Error GoTo EH
    If Not mbResize Then
        VisibleFrames False
        DoEvents
        Sleep 100
        Timer_Resize.Enabled = True
    End If
    
    Exit Sub
EH:
    Err.Clear
End Sub

Public Sub VisibleFrames(pbVisible As Boolean)
    On Error GoTo EH
    framMain.Visible = pbVisible
    framSendStatus.Visible = pbVisible
    framCommands.Visible = pbVisible
    framInsured.Visible = pbVisible
    framStatus.Visible = pbVisible
    framTol.Visible = pbVisible
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub VisibleFrames"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo EH
    goUtil.utFormWinRegPos goUtil.gsMainAppEXEName, Me, True
    
    If Not goUtil.gfrmECTray Is Nothing Then
        goUtil.gfrmECTray.ShowMe False
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_Unload"
End Sub

Public Function CLEANUP() As Boolean
    On Error GoTo EH
    mbModalFlag = False
    Set mcolXactProjects = Nothing
    Set mcolValidXProjects = Nothing
    Set moXact = Nothing
    Set mitmX = Nothing
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function CLEANUP"
End Function

Private Sub PopulateLookupCbo()
    On Error GoTo EH
    Dim vStates As Variant
    Dim vTOL As Variant
    Dim lCount As Long
    
    vStates = moXact.States
    vTOL = moXact.TOL
    
    If IsArray(vStates) Then
        cboState.Clear
        For lCount = LBound(vStates, 1) To UBound(vStates, 1)
            cboState.AddItem vStates(lCount)
        Next
    End If
    
    If IsArray(vTOL) Then
        cboTOL.Clear
        For lCount = LBound(vTOL, 1) To UBound(vTOL, 1)
            cboTOL.AddItem vTOL(lCount)
        Next
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub PopulateLookupCbo"
End Sub

Private Sub lvwXactProjects_Click()
    LoadEdit True
End Sub

Private Sub LoadEdit(Optional pbUseSelected As Boolean)
    On Error GoTo EH
    Dim itmX As listItem
    Dim bFound As Boolean
    Dim sZip5 As String
    Dim sZip4 As String
    
    mbLoadingEdit = True
    For Each itmX In lvwXactProjects.ListItems
        If pbUseSelected Then
            If itmX.Selected Then
                bFound = True
                Exit For
            End If
        Else
            If mitmX Is Nothing Then
                Exit Sub
            End If
                
            If itmX.Index = mitmX.Index Then
                bFound = True
                Exit For
            End If
        End If
    Next
    
    If bFound Then
        framInsured.Enabled = True
        framStatus.Enabled = True
        framTol.Enabled = True
        framSendStatus.Enabled = True
        optSend.Visible = True
        optSkip.Visible = True
        optSent.Visible = True
        
        With itmX
            Select Case .SmallIcon
                Case FixXactPic.AlreadySentToXact
                    optSent.Value = True
                    mbAlreadySent = True
                    mXProj.SentToXact = True
                Case FixXactPic.NeedsSentToXact
                    optSend.Value = True
                    mbAlreadySent = False
                    mXProj.SentToXact = False
                Case FixXactPic.Skip
                    optSkip.Value = True
                    mbAlreadySent = False
                    mXProj.SkipThisProject = True
            End Select
            txtName.Text = .SubItems(FixXactProj.Name - 1)
            mXProj.Main.Name = txtName.Text
            txtName.BackColor = IIf(.ListSubItems(FixXactProj.Name - 1).ReportIcon = FixXactPic.BadData, RED_BG, WHITE_BG)
            
            txtStreet.Text = .SubItems(FixXactProj.Street - 1)
            mXProj.Main.Street = txtStreet.Text
            txtStreet.BackColor = IIf(.ListSubItems(FixXactProj.Street - 1).ReportIcon = FixXactPic.BadData, RED_BG, WHITE_BG)
            
            txtCity.Text = .SubItems(FixXactProj.City - 1)
            mXProj.Main.City = txtCity.Text
            txtCity.BackColor = IIf(.ListSubItems(FixXactProj.City - 1).ReportIcon = FixXactPic.BadData, RED_BG, WHITE_BG)
            
            cboState.Text = .SubItems(FixXactProj.State - 1)
            mXProj.Main.State = cboState.Text
            cboState.BackColor = IIf(.ListSubItems(FixXactProj.State - 1).ReportIcon = FixXactPic.BadData, RED_BG, WHITE_BG)
            
'            moXact.ZipValid .SubItems(FixXactProj.Zip - 1), sZip5, sZip4
            txtZip5.Text = Format(.SubItems(FixXactProj.Zip - 1), "00000")
            txtZip5.BackColor = IIf(.ListSubItems(FixXactProj.Zip - 1).ReportIcon = FixXactPic.BadData, RED_BG, WHITE_BG)
            txtZip4.Text = Format(.SubItems(FixXactProj.Zip4 - 1), "0000")
            txtZip4.BackColor = IIf(.ListSubItems(FixXactProj.Zip4 - 1).ReportIcon = FixXactPic.BadData, RED_BG, WHITE_BG)
            mXProj.Main.Zip = txtZip5.Text
            mXProj.Main.Zip4 = txtZip4.Text
            
            txtDateofLoss.Text = .SubItems(FixXactProj.DateOfLoss - 1)
            mXProj.Main.DateOfLoss = txtDateofLoss.Text
            txtDateofLoss.BackColor = IIf(.ListSubItems(FixXactProj.DateOfLoss - 1).ReportIcon = FixXactPic.BadData, RED_BG, WHITE_BG)
            
            txtDateReceived.Text = .SubItems(FixXactProj.DateReceived - 1)
            mXProj.Main.DateReceived = txtDateReceived.Text
            txtDateReceived.BackColor = IIf(.ListSubItems(FixXactProj.DateReceived - 1).ReportIcon = FixXactPic.BadData, RED_BG, WHITE_BG)
            
            txtDateInspected.Text = .SubItems(FixXactProj.DateInspected - 1)
            mXProj.Main.DateInspected = txtDateInspected.Text
            txtDateInspected.BackColor = IIf(.ListSubItems(FixXactProj.DateInspected - 1).ReportIcon = FixXactPic.BadData, RED_BG, WHITE_BG)
            
            txtDateEntered.Text = .SubItems(FixXactProj.DateEntered - 1)
            mXProj.Main.DateEntered = txtDateEntered.Text
            txtDateEntered.BackColor = IIf(.ListSubItems(FixXactProj.DateEntered - 1).ReportIcon = FixXactPic.BadData, RED_BG, WHITE_BG)
            
            txtClaimNumber.Text = .SubItems(FixXactProj.ClaimNumber - 1)
            mXProj.Loss.ClaimNumber = txtClaimNumber.Text
            txtClaimNumber.BackColor = IIf(.ListSubItems(FixXactProj.ClaimNumber - 1).ReportIcon = FixXactPic.BadData, RED_BG, WHITE_BG)
            
            txtPolicyNumber.Text = .SubItems(FixXactProj.PolicyNumber - 1)
            mXProj.Loss.PolicyNumber = txtPolicyNumber.Text
            txtPolicyNumber.BackColor = IIf(.ListSubItems(FixXactProj.PolicyNumber - 1).ReportIcon = FixXactPic.BadData, RED_BG, WHITE_BG)
            
            txtDeductible.Text = .SubItems(FixXactProj.Deductible - 1)
            mXProj.Loss.Deductible = txtDeductible.Text
            txtDeductible.BackColor = IIf(.ListSubItems(FixXactProj.Deductible - 1).ReportIcon = FixXactPic.BadData, RED_BG, WHITE_BG)
            
            cboTOL.Text = .SubItems(FixXactProj.TypeOfLoss - 1)
            mXProj.Loss.TypeOfLoss = cboTOL.Text
            cboTOL.BackColor = IIf(.ListSubItems(FixXactProj.TypeOfLoss - 1).ReportIcon = FixXactPic.BadData, RED_BG, WHITE_BG)
            
            txtCatCode.Text = .SubItems(FixXactProj.CatCode - 1)
            mXProj.Loss.CatCode = txtCatCode.Text
            txtCatCode.BackColor = IIf(.ListSubItems(FixXactProj.CatCode - 1).ReportIcon = FixXactPic.BadData, RED_BG, WHITE_BG)
        End With
        Set mitmX = itmX
    End If
    mbLoadingEdit = False
    EnableSend
    Exit Sub
EH:
    mbLoadingEdit = False
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub LoadEdit"
End Sub

Private Sub lvwXactProjects_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    On Error GoTo EH
    If lvwXactProjects.SortOrder = lvwAscending Then
        lvwXactProjects.SortOrder = lvwDescending
    Else
        lvwXactProjects.SortOrder = lvwAscending
    End If
    
    
    'Need to see if the Column clicked was a NON text sort
    'Like Date or number.  If a non textual column was clicked
    'need to use the next column as the sort since this next column
    'will be hidden and contain a sort friendly format.
    Select Case ColumnHeader.Index
        Case FixXactProj.Street
            lvwXactProjects.SortKey = ColumnHeader.Index
        Case FixXactProj.DateEntered
            lvwXactProjects.SortKey = ColumnHeader.Index
        Case FixXactProj.DateInspected
            lvwXactProjects.SortKey = ColumnHeader.Index
        Case FixXactProj.DateOfLoss
            lvwXactProjects.SortKey = ColumnHeader.Index
        Case FixXactProj.DateReceived
            lvwXactProjects.SortKey = ColumnHeader.Index
        Case FixXactProj.Deductible
            lvwXactProjects.SortKey = ColumnHeader.Index
        Case Else
            lvwXactProjects.SortKey = ColumnHeader.Index - 1
    End Select
    
    lvwXactProjects.Sorted = True
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lvwXactProjects_ColumnClick"
End Sub

Private Sub lvwXactProjects_KeyUp(KeyCode As Integer, Shift As Integer)
    LoadEdit True
End Sub

Private Sub optSend_Click()
    On Error GoTo EH
    Dim itmX As listItem
    
    If Not mitmX Is Nothing Then
        For Each itmX In lvwXactProjects.ListItems
            If itmX.Selected Then
                itmX.SmallIcon = FixXactPic.NeedsSentToXact
            End If
        Next
        EnableSend
    End If
    
    Set itmX = Nothing
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub optSend_Click"
End Sub

Private Sub optSend_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo EH
    
    If KeyCode = vbKeyS Then
        lvwXactProjects.SetFocus
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub optSend_KeyUp"
End Sub


Private Sub optSkip_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo EH
    
    If KeyCode = vbKeyP Then
        lvwXactProjects.SetFocus
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub optSkip_KeyUp"
End Sub

Private Sub optSkip_Click()
Dim itmX As listItem
    On Error GoTo EH
    If Not mitmX Is Nothing Then
        For Each itmX In lvwXactProjects.ListItems
            If itmX.Selected Then
                itmX.SmallIcon = FixXactPic.Skip
            End If
        Next
        EnableSend
    End If
    
    Set itmX = Nothing
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub optSkip_Click"
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
    
    'Main Frame
    framMain.Height = lH - framMain_H
    framMain.Width = lW - framMain_W
    txtSendToExportPath.Width = txtSendToExportPath_W + (lW - FORM_W) / 2
    cmdSendToExportPath.left = cmdSendToExportPath_L + (lW - FORM_W) / 2
    chkGetFromExport.left = chkGetFromExport_L + (lW - FORM_W) / 2
    txtGetFromExportPath.left = txtGetFromExportPath_L + (lW - FORM_W) / 2
    txtGetFromExportPath.Width = txtGetFromExportPath_W + (lW - FORM_W) / 2
    cmdGetFromExportPath.left = cmdGetFromExportPath_L + (lW - FORM_W)
    lvwXactProjects.Width = lW - lvwXactProjects_W
    lvwXactProjects.Height = lH - lvwXactProjects_H
    cmdUp.top = lH - cmdUp_T
    cmdDown.top = lH - cmdDown_T
    cmdFind.top = lH - cmdFind_T
    cmdFindNext.top = lH - cmdFindNext_T
    cmdViewFailed.top = lH - cmdViewFailed_T
    chkBrowseBadData.top = lH - chkBrowseBadData_T
    chkShowGrid.top = lH - chkShowGrid_T
    
    'Fram Status
    framSendStatus.top = lH - framSendStatus_T
    
    'Fram Commands
    framCommands.top = lH - framCommands_T
    framCommands.Width = lW - framCommands_W
    cmdPrintList.left = lW - cmdPrintList_L
    cmdExit.left = lW - cmdExit_L
    
    'framInsured
    framInsured.top = lH - framInsured_T
    framInsured.Width = lW - framInsured_W
    txtName.Width = lW - txtName_W
    txtStreet.Width = lW - txtStreet_W
    txtCity.Width = lW - txtCity_W
    
    'framStatus
    framStatus.top = lH - framStatus_T
    framStatus.left = lW - framStatus_L
    
    'framTol
    framTol.top = lH - framTol_T
    framTol.left = lW - framTol_L
    
    
    VisibleFrames True
    
    mbResize = False
    Exit Sub
EH:
    Err.Clear
    Resume Next
End Sub


Private Sub txtCatCode_Change()
    ValidateControl txtCatCode, FixXactProj.CatCode, False
End Sub

Private Sub txtCatCode_GotFocus()
    SelText txtCatCode
End Sub

Private Sub txtCatCode_LostFocus()
    LoadEdit
End Sub

Private Sub txtCity_Change()
    ValidateControl txtCity, City, True
End Sub

Private Sub txtCity_GotFocus()
    SelText txtCity
End Sub

Private Sub txtCity_LostFocus()
    LoadEdit
End Sub

Private Sub txtClaimNumber_Change()
    ValidateControl txtClaimNumber, ClaimNumber, False
End Sub

Private Sub txtClaimNumber_GotFocus()
    SelText txtClaimNumber
End Sub

Private Sub txtClaimNumber_LostFocus()
    LoadEdit
End Sub

Private Sub txtDateEntered_Change()
    ValidateControl txtDateEntered, DateEntered, True
End Sub

Private Sub txtDateEntered_GotFocus()
    SelText txtDateEntered
End Sub

Private Sub txtDateEntered_LostFocus()
    LoadEdit
End Sub

Private Sub txtDateInspected_Change()
    ValidateControl txtDateInspected, DateInspected, True
End Sub

Private Sub txtDateInspected_GotFocus()
    SelText txtDateInspected
End Sub

Private Sub txtDateInspected_LostFocus()
    LoadEdit
End Sub

Private Sub txtDateofLoss_Change()
    ValidateControl txtDateofLoss, DateOfLoss, True
End Sub

Private Sub txtDateofLoss_GotFocus()
    SelText txtDateofLoss
End Sub

Private Sub txtDateofLoss_LostFocus()
    LoadEdit
End Sub

Private Sub txtDateReceived_Change()
    ValidateControl txtDateReceived, DateReceived, True
End Sub

Private Sub txtDateReceived_GotFocus()
    SelText txtDateReceived
End Sub

Private Sub txtDateReceived_LostFocus()
    LoadEdit
End Sub

Private Sub txtDeductible_Change()
    ValidateControl txtDeductible, Deductible, False
End Sub

Private Sub txtDeductible_GotFocus()
    SelText txtDeductible
End Sub

Private Sub txtDeductible_LostFocus()
    LoadEdit
End Sub


Private Sub txtName_Change()
    ValidateControl txtName, FixXactProj.Name, True
End Sub

Private Sub txtName_GotFocus()
    SelText txtName
End Sub

Private Sub txtName_LostFocus()
    LoadEdit
End Sub

Private Sub txtPolicyNumber_Change()
    ValidateControl txtPolicyNumber, PolicyNumber, False
End Sub

Private Sub txtPolicyNumber_GotFocus()
    SelText txtPolicyNumber
End Sub

Private Sub txtPolicyNumber_LostFocus()
    LoadEdit
End Sub

Private Sub txtStreet_Change()
    ValidateControl txtStreet, Street, True
End Sub

Private Sub txtStreet_GotFocus()
    SelText txtStreet
End Sub

Private Sub txtStreet_LostFocus()
    LoadEdit
End Sub

Private Sub txtZip4_Change()
    On Error Resume Next
    Dim lPos As Long
    lPos = txtZip4.SelStart
    ValidateControl txtZip4, Zip4, True
    txtZip4.SelStart = lPos
End Sub

Private Sub txtZip4_GotFocus()
    SelText txtZip4
End Sub

Private Sub txtZip4_LostFocus()
    On Error Resume Next
    Dim lPos As Long
    LoadEdit
    lPos = txtZip4.SelStart
    ValidateControl txtZip4, Zip4, True
    txtZip4.SelStart = lPos
End Sub

Private Sub txtZip5_Change()
    On Error Resume Next
    Dim lPos As Long
    lPos = txtZip5.SelStart
    ValidateControl txtZip5, Zip, True
    txtZip5.SelStart = lPos
End Sub

Private Sub txtZip5_GotFocus()
    SelText txtZip5
End Sub

Private Sub txtZip5_LostFocus()
    On Error Resume Next
    Dim lPos As Long
    LoadEdit
    lPos = txtZip5.SelStart
    ValidateControl txtZip5, Zip, True
    txtZip5.SelStart = lPos
End Sub

Private Sub ValidateControl(pControl As Control, piXProjItem As FixXactProj, pbMain As Boolean)
    On Error GoTo EH
    Dim bZip As Boolean
    Dim bZip4 As Boolean
    Dim vXProj As Variant
    Dim XProj As udtXactProject
    Dim bValidControl As Boolean
    Dim itmX As listItem
    
    
    If mbLoadingEdit Then
        Exit Sub
    End If
    
    With mitmX
        If pbMain Then
            Select Case piXProjItem
                Case FixXactProj.Name
                    mXProj.Main.Name = pControl.Text
                    
                Case FixXactProj.Street
                    mXProj.Main.Street = pControl.Text
                    
                Case FixXactProj.City
                    mXProj.Main.City = pControl.Text
                    
                Case FixXactProj.State
                    mXProj.Main.State = pControl.Text
                Case FixXactProj.Zip
                    bZip = True
                    mXProj.Main.Zip = txtZip5.Text
                Case FixXactProj.Zip4
                    bZip4 = True
                     mXProj.Main.Zip4 = txtZip4.Text
                    
                Case FixXactProj.DateEntered
                    mXProj.Main.DateEntered = pControl.Text
                    
                Case FixXactProj.DateInspected
                    mXProj.Main.DateInspected = pControl.Text
                    
                Case FixXactProj.DateOfLoss
                    mXProj.Main.DateOfLoss = pControl.Text
                    
                Case FixXactProj.DateReceived
                    mXProj.Main.DateReceived = pControl.Text
            End Select
        Else
            Select Case piXProjItem
                Case FixXactProj.ClaimNumber
                    mXProj.Loss.ClaimNumber = pControl.Text
                    
                Case FixXactProj.PolicyNumber
                    mXProj.Loss.PolicyNumber = pControl.Text
                    
                Case FixXactProj.TypeOfLoss
                    mXProj.Loss.TypeOfLoss = pControl.Text
                    
                Case FixXactProj.Deductible
                    mXProj.Loss.Deductible = pControl.Text
                    
                Case FixXactProj.CatCode
                    mXProj.Loss.CatCode = pControl.Text
            End Select
        End If
        
        moXact.ValidateXProject mXProj, mitmX
        
        If Not bZip Then
            If .ListSubItems(piXProjItem - 1).ReportIcon = FixXactPic.BadData Then
                pControl.BackColor = RED_BG
            Else
                pControl.BackColor = WHITE_BG
                bValidControl = True
            End If
        Else
            If .ListSubItems(piXProjItem - 1).ReportIcon = FixXactPic.BadData Then
                txtZip5.BackColor = RED_BG
            Else
                txtZip5.BackColor = WHITE_BG
            End If
        End If
        If Not bZip4 Then
            If .ListSubItems(piXProjItem - 1).ReportIcon = FixXactPic.BadData Then
                pControl.BackColor = RED_BG
            Else
                pControl.BackColor = WHITE_BG
                bValidControl = True
            End If
        Else
            If .ListSubItems(piXProjItem - 1).ReportIcon = FixXactPic.BadData Then
                txtZip4.BackColor = RED_BG
            Else
                txtZip4.BackColor = WHITE_BG
            End If
        End If
        'CHeck for Multi Updates
        If bValidControl = True Then
            Select Case piXProjItem
                Case FixXactProj.State, FixXactProj.TypeOfLoss
                    For Each vXProj In mcolXactProjects
                        XProj = vXProj
                        Select Case piXProjItem
                            Case FixXactProj.State
                                XProj.Main.State = pControl.Text
                            Case FixXactProj.TypeOfLoss
                                XProj.Loss.TypeOfLoss = pControl.Text
                        End Select
                        If XProj.Loss.ClaimNumber <> mXProj.Loss.ClaimNumber Then
                            Set itmX = lvwXactProjects.ListItems("""" & CStr(XProj.Loss.ClaimNumber) & """")
                            If itmX.Selected Then
                                If piXProjItem <> FixXactProj.State Then
                                    XProj.Main.State = itmX.SubItems(FixXactProj.State - 1)
                                End If
                                If piXProjItem <> FixXactProj.TypeOfLoss Then
                                    XProj.Loss.TypeOfLoss = itmX.SubItems(FixXactProj.TypeOfLoss - 1)
                                End If
                                moXact.ValidateXProject XProj, itmX
                            End If
                        End If
                    Next
            End Select
        End If
        EnableSend
    End With
    
    'cleanup
    Set itmX = Nothing
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub ValidateControl"
End Sub



Private Sub EnableSend()
    On Error GoTo EH
    Dim itmX As listItem
    Dim bFoundSend As Boolean
    
    mbProjectsValidated = True
    For Each itmX In lvwXactProjects.ListItems
        If itmX.SmallIcon = FixXactPic.NeedsSentToXact Then
            bFoundSend = True
            If itmX.ListSubItems(FixXactProj.Data - 1).ReportIcon = FixXactPic.BadData Then
                'Allow failed Validation When Exporting to File
                'Only when Actually sending to Xactimate does this apply.
                If Not moXact.SendToExport Then
                    mbProjectsValidated = False
                    Exit For
                End If
            End If
        End If
    Next
    If mbProjectsValidated And bFoundSend Then
        cmdSendAll.Enabled = True
    Else
        cmdSendAll.Enabled = False
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub EnableSend"
End Sub

Private Sub LoadValidXProjects()
    On Error GoTo EH
    Dim itmX As listItem
    Dim XProj As udtXactProject
    
    Set mcolValidXProjects = New Collection
    
    For Each itmX In lvwXactProjects.ListItems
        
        If itmX.SmallIcon = FixXactPic.NeedsSentToXact Then
            XProj.SentToXact = False
        End If
        If itmX.SmallIcon = FixXactPic.AlreadySentToXact Then
            XProj.SentToXact = True
        End If
        If itmX.SmallIcon = FixXactPic.Skip Or mbSkipAll Then
            'Issue 248 9.23.2002 Send To Xact Status marked as Sent when it is not.
            'need to be sure set this flag is False
            '3.1.2004 Don't reset this if Loading from Export file
            If Not moXact.GetFromExport Then
                XProj.SentToXact = False
            End If
            XProj.SkipThisProject = True
        Else
            XProj.SkipThisProject = False
        End If
        If itmX.ListSubItems(FixXactProj.Data - 1).ReportIcon = FixXactPic.BadData Then
            XProj.ValidData = False
        Else
            XProj.ValidData = True
        End If
        
        With XProj.Main
            .City = itmX.SubItems(FixXactProj.City - 1)
            .DateEntered = itmX.SubItems(FixXactProj.DateEntered - 1)
            .DateInspected = itmX.SubItems(FixXactProj.DateInspected - 1)
            .DateOfLoss = itmX.SubItems(FixXactProj.DateOfLoss - 1)
            .DateReceived = itmX.SubItems(FixXactProj.DateReceived - 1)
            .Name = itmX.SubItems(FixXactProj.Name - 1)
            .State = itmX.SubItems(FixXactProj.State - 1)
            .Street = itmX.SubItems(FixXactProj.Street - 1)
            .Zip = itmX.SubItems(FixXactProj.Zip - 1)
            .Zip4 = itmX.SubItems(FixXactProj.Zip4 - 1)
        End With
        With XProj.Loss
            .CatCode = itmX.SubItems(FixXactProj.CatCode - 1)
            .ClaimNumber = itmX.SubItems(FixXactProj.ClaimNumber - 1)
            .Deductible = itmX.SubItems(FixXactProj.Deductible - 1)
            .PolicyNumber = itmX.SubItems(FixXactProj.PolicyNumber - 1)
            .TypeOfLoss = itmX.SubItems(FixXactProj.TypeOfLoss - 1)
        End With
        mcolValidXProjects.Add XProj, XProj.Loss.ClaimNumber
    Next
    
    'clean up
    Set itmX = Nothing
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub LoadValidXProjects"
End Sub

Private Sub PopulateXactSettings()
    'Issue 224  9.23.2002 Send to Xactimate no work with Xactimate 2002
    On Error GoTo EH
    
    'Version Info
    cboXactVS.AddItem "2001"
    cboXactVS.AddItem "2002"
    cboXactVS.AddItem "2002(05-29-2003 272)"
    cboXactVS.AddItem "2002(05-25-2004 276)"
    cboXactVS.Text = GetSetting("ECS", "XACTIMATE", "VERSION", "2001")
        
    'Speed Settings
    cboXactSpeed.AddItem "Entry Speed - FAST"
    cboXactSpeed.AddItem "Entry Speed - MEDIUM"
    cboXactSpeed.AddItem "Entry Speed - SLOW"
    
    cboXactSpeed.Text = GetSetting("ECS", "KEYBOARD", "SPEED", "Entry Speed - FAST")
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub PopulateXactSettings"
End Sub
