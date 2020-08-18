VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmClaimInfo 
   AutoRedraw      =   -1  'True
   Caption         =   "Claim Info"
   ClientHeight    =   6555
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
   ScaleHeight     =   6555
   ScaleWidth      =   11820
   Tag             =   "Claim Info"
   Begin VB.Frame framSpecifics 
      Height          =   4695
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   11295
      Begin VB.TextBox txtLossDate 
         Height          =   360
         Left            =   120
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "Date"
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox txtAssignedDate 
         Height          =   360
         Left            =   120
         MaxLength       =   10
         TabIndex        =   6
         Tag             =   "Date"
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox txtReceivedDate 
         Height          =   360
         Left            =   120
         MaxLength       =   10
         TabIndex        =   9
         Tag             =   "Date"
         Top             =   1680
         Width           =   1455
      End
      Begin VB.CommandButton cmdLossDate 
         Height          =   375
         Left            =   1560
         Picture         =   "frmClaimInfo.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Tag             =   "Date"
         Top             =   465
         Width           =   375
      End
      Begin VB.CommandButton cmdAssignedDate 
         Height          =   375
         Left            =   1560
         Picture         =   "frmClaimInfo.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   7
         Tag             =   "Date"
         Top             =   1065
         Width           =   375
      End
      Begin VB.CommandButton cmdReceivedDate 
         Height          =   375
         Left            =   1560
         Picture         =   "frmClaimInfo.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   10
         Tag             =   "Date"
         Top             =   1665
         Width           =   375
      End
      Begin VB.TextBox txtContactDate 
         Height          =   360
         Left            =   120
         MaxLength       =   10
         TabIndex        =   12
         Tag             =   "Date"
         Top             =   2280
         Width           =   1455
      End
      Begin VB.TextBox txtInspectedDate 
         Height          =   360
         Left            =   120
         MaxLength       =   10
         TabIndex        =   15
         Tag             =   "Date"
         Top             =   2880
         Width           =   1455
      End
      Begin VB.TextBox txtCloseDate 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   360
         Left            =   120
         MaxLength       =   10
         TabIndex        =   18
         Tag             =   "Date"
         ToolTipText     =   "Set the Close Date ONLY when you have completed this Claim!"
         Top             =   3480
         Width           =   1455
      End
      Begin VB.CommandButton cmdContactDate 
         Height          =   375
         Left            =   1560
         Picture         =   "frmClaimInfo.frx":0CC6
         Style           =   1  'Graphical
         TabIndex        =   13
         Tag             =   "Date"
         Top             =   2265
         Width           =   375
      End
      Begin VB.CommandButton cmdInspectedDate 
         Height          =   375
         Left            =   1560
         Picture         =   "frmClaimInfo.frx":1108
         Style           =   1  'Graphical
         TabIndex        =   16
         Tag             =   "Date"
         Top             =   2865
         Width           =   375
      End
      Begin VB.CommandButton cmdCloseDate 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1560
         Picture         =   "frmClaimInfo.frx":154A
         Style           =   1  'Graphical
         TabIndex        =   19
         Tag             =   "Date"
         Top             =   3465
         Width           =   375
      End
      Begin VB.TextBox txtReportedByPhone 
         Height          =   360
         Left            =   7440
         MaxLength       =   50
         TabIndex        =   33
         Tag             =   "UCASE"
         Top             =   1680
         Width           =   3735
      End
      Begin VB.TextBox txtReportedBy 
         Height          =   360
         Left            =   7440
         MaxLength       =   50
         TabIndex        =   31
         Tag             =   "UCASE"
         Top             =   1080
         Width           =   3735
      End
      Begin VB.TextBox txtAgentNo 
         Height          =   360
         Left            =   7440
         MaxLength       =   50
         TabIndex        =   29
         Tag             =   "UCASE"
         Top             =   480
         Width           =   3735
      End
      Begin VB.TextBox txtCLIENTNUM 
         Height          =   360
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   21
         Tag             =   "UCASE_ALPHANUM"
         Top             =   480
         Width           =   5295
      End
      Begin VB.ComboBox cboCatCode 
         Height          =   360
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   2280
         Width           =   5295
      End
      Begin VB.ComboBox cboTypeOfLoss 
         CausesValidation=   0   'False
         Height          =   360
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   1680
         Width           =   5295
      End
      Begin VB.ComboBox cboAssignmentType 
         Enabled         =   0   'False
         Height          =   360
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   1080
         Width           =   5295
      End
      Begin VB.Label lblLossDate 
         Caption         =   "Loss Date:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblAssignedDate 
         Caption         =   "Assigned Date:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label lblReceivedDate 
         Caption         =   "Received Date: "
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label lblContactDate 
         Caption         =   "Contact Date:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label lblInspectedDate 
         Caption         =   "Inspected Date:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label lblCloseDate 
         Caption         =   "Close Date: "
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   3240
         Width           =   1455
      End
      Begin VB.Label lblReportedByPhone 
         Caption         =   "Reported By Phone:"
         Height          =   255
         Left            =   7440
         TabIndex        =   32
         Top             =   1440
         Width           =   3615
      End
      Begin VB.Label lblReportedBy 
         Caption         =   "Reported By: "
         Height          =   255
         Left            =   7440
         TabIndex        =   30
         Top             =   840
         Width           =   3735
      End
      Begin VB.Label lblAgentNo 
         Caption         =   "Agent No:"
         Height          =   255
         Left            =   7440
         TabIndex        =   28
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label lblCLIENTNUM 
         Caption         =   "Claim Number:"
         Height          =   255
         Left            =   2040
         TabIndex        =   20
         Top             =   240
         Width           =   5175
      End
      Begin VB.Label lblCatCode 
         Caption         =   "Cat Code:"
         Height          =   255
         Left            =   2040
         TabIndex        =   26
         Top             =   2040
         Width           =   5175
      End
      Begin VB.Label lblTypeOfLoss 
         Caption         =   "Type of Loss:"
         Height          =   255
         Left            =   2040
         TabIndex        =   24
         Top             =   1440
         Width           =   5175
      End
      Begin VB.Label lblAssignmentType 
         Caption         =   "Assignment Type:"
         Height          =   255
         Left            =   2040
         TabIndex        =   22
         Top             =   840
         Width           =   5175
      End
   End
   Begin VB.Frame framPolicyLimits 
      Height          =   4695
      Left            =   240
      TabIndex        =   71
      Top             =   480
      Visible         =   0   'False
      Width           =   11295
      Begin VB.TextBox txtPolicyNo 
         Height          =   360
         Left            =   120
         MaxLength       =   20
         TabIndex        =   79
         Tag             =   "UCASE"
         Top             =   1080
         Width           =   2535
      End
      Begin VB.TextBox txtPolicyDescription 
         Height          =   360
         Left            =   2760
         MaxLength       =   100
         TabIndex        =   82
         Tag             =   "UCASE"
         Top             =   1080
         Width           =   3495
      End
      Begin VB.TextBox txtMortgageeName 
         Height          =   360
         Left            =   120
         MaxLength       =   100
         TabIndex        =   75
         Tag             =   "UCASE"
         Top             =   480
         Width           =   6135
      End
      Begin VB.TextBox txtPLReserves 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F2FFFF&
         Height          =   360
         Left            =   6600
         MaxLength       =   10
         TabIndex        =   95
         Tag             =   "Currency"
         Top             =   2280
         Width           =   1575
      End
      Begin VB.CommandButton cmdPrintList 
         Caption         =   "Sho&w Item List"
         Height          =   735
         Left            =   10200
         MaskColor       =   &H00000000&
         Picture         =   "frmClaimInfo.frx":198C
         Style           =   1  'Graphical
         TabIndex        =   83
         ToolTipText     =   "Exit"
         Top             =   240
         Visible         =   0   'False
         Width           =   975
      End
      Begin MSComctlLib.ImageList imgPolicyLimits 
         Left            =   5880
         Top             =   1680
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
               Picture         =   "frmClaimInfo.frx":1DCE
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClaimInfo.frx":2220
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClaimInfo.frx":260C
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdUpdatePLClassTypeID 
         Caption         =   "&UPDATE"
         Enabled         =   0   'False
         Height          =   735
         Left            =   9840
         Style           =   1  'Graphical
         TabIndex        =   100
         Top             =   3840
         Width           =   1335
      End
      Begin VB.CheckBox chkPLIsDeleted 
         Caption         =   "Is Deleted:"
         DownPicture     =   "frmClaimInfo.frx":291B
         Height          =   735
         Left            =   6600
         Style           =   1  'Graphical
         TabIndex        =   99
         Tag             =   "CCDisable"
         Top             =   3840
         Width           =   1335
      End
      Begin VB.TextBox txtPLAdminComments 
         BackColor       =   &H00F2FFFF&
         Height          =   960
         Left            =   6600
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   98
         Tag             =   "CCDisable"
         Top             =   2880
         Width           =   4575
      End
      Begin VB.TextBox txtPLRCSaidProp 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F2FFFF&
         Height          =   360
         Left            =   6600
         MaxLength       =   10
         TabIndex        =   93
         Tag             =   "Currency"
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox txtPLLimitAmount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F2FFFF&
         Height          =   360
         Left            =   6600
         MaxLength       =   10
         TabIndex        =   91
         Tag             =   "Currency"
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CommandButton cmdAddPLAppClassTypeID 
         Caption         =   "&ADD"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   85
         Tag             =   "CCDisable"
         Top             =   1680
         Width           =   975
      End
      Begin VB.ComboBox cboAddPLAppClassTypeID 
         BackColor       =   &H00F2FFFF&
         Height          =   360
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   86
         Tag             =   "CCDisable"
         Top             =   1680
         Width           =   5295
      End
      Begin VB.TextBox txtDeductible 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   6600
         MaxLength       =   10
         TabIndex        =   89
         Tag             =   "Numeric"
         Top             =   480
         Width           =   1575
      End
      Begin MSComctlLib.ListView lvwPLClassTypeID 
         Height          =   2535
         Left            =   120
         TabIndex        =   87
         Top             =   2040
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   4471
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "imgPolicyLimits"
         SmallIcons      =   "imgPolicyLimits"
         ColHdrIcons     =   "imgPolicyLimits"
         ForeColor       =   -2147483640
         BackColor       =   15925247
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
      Begin VB.Label lblPolicyNo 
         Caption         =   "Policy Number:"
         Height          =   255
         Left            =   120
         TabIndex        =   77
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label lblPolicyDescription 
         Caption         =   "Policy Description:"
         Height          =   255
         Left            =   2760
         TabIndex        =   81
         Top             =   840
         Width           =   3375
      End
      Begin VB.Label lblMortgageeName 
         Caption         =   "Mortgagee Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   73
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label lblPLReserves 
         Caption         =   "Reserves:"
         Height          =   255
         Left            =   6600
         TabIndex        =   94
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label lblPLCaption 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F2FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1575
         Left            =   8280
         TabIndex        =   96
         Top             =   1080
         Width           =   2895
      End
      Begin VB.Label lblPLAdminComments 
         Caption         =   "Admin Comments:"
         Height          =   255
         Left            =   6600
         TabIndex        =   97
         Top             =   2640
         Width           =   3135
      End
      Begin VB.Label lblPLRCSaidProp 
         Caption         =   "RC Said:"
         Height          =   255
         Left            =   6600
         TabIndex        =   92
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label lblPLLimitAmount 
         Caption         =   "Limit Amount:"
         Height          =   255
         Left            =   6600
         TabIndex        =   90
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lblPolicyLimit 
         Caption         =   "Policy Limits for Line of Coverage:"
         Height          =   255
         Left            =   120
         TabIndex        =   84
         Top             =   1440
         Width           =   4575
      End
      Begin VB.Label lblDeductible 
         Caption         =   "Deductible:"
         Height          =   255
         Left            =   6600
         TabIndex        =   88
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame framLossReport 
      Height          =   4695
      Left            =   240
      TabIndex        =   101
      Top             =   480
      Visible         =   0   'False
      Width           =   11295
      Begin VB.CommandButton cmdViewPDFLossReport 
         Height          =   375
         Left            =   8280
         Picture         =   "frmClaimInfo.frx":2A65
         Style           =   1  'Graphical
         TabIndex        =   105
         Top             =   480
         Width           =   375
      End
      Begin VB.CheckBox chkDelOrigPDF 
         Alignment       =   1  'Right Justify
         Caption         =   "Delete Original PDF After Attach:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8760
         TabIndex        =   110
         Top             =   940
         Width           =   2415
      End
      Begin VB.CommandButton cmdAttachPDFLossReport 
         Caption         =   "&Attach"
         Enabled         =   0   'False
         Height          =   375
         Left            =   8760
         TabIndex        =   111
         Top             =   1200
         Width           =   2415
      End
      Begin VB.TextBox txtLossReport 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   32000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   113
         Top             =   1920
         Width           =   11055
      End
      Begin VB.ComboBox cboAssignmentLossReportFormat 
         Height          =   360
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   103
         Tag             =   "CCDisable"
         Top             =   480
         Width           =   4455
      End
      Begin VB.CommandButton cmdBuildTextLossReport 
         Caption         =   "&Build Text Loss Report"
         Enabled         =   0   'False
         Height          =   375
         Left            =   8760
         TabIndex        =   106
         Top             =   465
         Width           =   2415
      End
      Begin VB.CommandButton cmdlAttachPDFLossReportPath 
         Enabled         =   0   'False
         Height          =   330
         Left            =   8280
         Picture         =   "frmClaimInfo.frx":2BAF
         Style           =   1  'Graphical
         TabIndex        =   109
         ToolTipText     =   "Browse"
         Top             =   1210
         Width           =   330
      End
      Begin VB.TextBox txtAttachPDFLossReportPath 
         Height          =   360
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   108
         Top             =   1200
         Width           =   8520
      End
      Begin VB.Label lblSelFormat 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   4680
         TabIndex        =   104
         Top             =   480
         Width           =   3495
      End
      Begin VB.Label lblAttachPDFLossReportPath 
         Caption         =   "PDF Loss Report Attachment Windows Path"
         Height          =   255
         Left            =   120
         TabIndex        =   107
         Top             =   960
         Width           =   4455
      End
      Begin VB.Label lblLossReport 
         Caption         =   "Text Formated Loss Report ( - 32000 character limit -): "
         Height          =   255
         Left            =   120
         TabIndex        =   112
         Top             =   1680
         Width           =   9855
      End
      Begin VB.Label lblAssignmentLossReportFormat 
         Caption         =   "Loss Report Format:"
         Height          =   255
         Left            =   120
         TabIndex        =   102
         Top             =   240
         Width           =   4455
      End
   End
   Begin VB.Frame framAdjusterInfo 
      Height          =   4695
      Left            =   240
      TabIndex        =   34
      Top             =   480
      Visible         =   0   'False
      Width           =   11295
      Begin VB.TextBox txtIBNUM 
         Height          =   360
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   40
         Tag             =   "UCASE"
         Top             =   1920
         Width           =   4215
      End
      Begin VB.ComboBox cboACID 
         Height          =   360
         ItemData        =   "frmClaimInfo.frx":3029
         Left            =   120
         List            =   "frmClaimInfo.frx":302B
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   480
         Width           =   4215
      End
      Begin VB.ComboBox cboACIDDisplay 
         Height          =   360
         ItemData        =   "frmClaimInfo.frx":302D
         Left            =   120
         List            =   "frmClaimInfo.frx":302F
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   1200
         Width           =   4215
      End
      Begin VB.Label lblIBNUM 
         Caption         =   "IB Number (Internal Billing):"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   1680
         Width           =   2775
      End
      Begin VB.Label lblACID 
         Caption         =   "ACID (Adjuster Client Identification): "
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Width           =   4215
      End
      Begin VB.Label lblACIDDisplay 
         Caption         =   "Display ACID (Will Print on Billing Info): "
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   960
         Width           =   4215
      End
   End
   Begin VB.Frame framInsuredInfo 
      Height          =   4695
      Left            =   240
      TabIndex        =   41
      Top             =   480
      Visible         =   0   'False
      Width           =   11295
      Begin VB.CheckBox chkUseProperty 
         Caption         =   "(Use Property Address)"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   3465
         Width           =   2535
      End
      Begin VB.CheckBox chkUseMailing 
         Caption         =   "(Use Mailing Address)"
         Height          =   360
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   1965
         Width           =   2535
      End
      Begin VB.TextBox txtPAOtherPostCode 
         Height          =   360
         Left            =   9480
         MaxLength       =   20
         TabIndex        =   61
         Tag             =   "UCASE"
         Top             =   2685
         Width           =   1695
      End
      Begin VB.TextBox txtPAZIP4 
         Height          =   360
         Left            =   8280
         MaxLength       =   4
         TabIndex        =   59
         Tag             =   "ZipCode4"
         Top             =   2685
         Width           =   1095
      End
      Begin VB.TextBox txtPAZIP 
         Height          =   360
         Left            =   7080
         MaxLength       =   5
         TabIndex        =   57
         Tag             =   "ZipCode5"
         Top             =   2685
         Width           =   1095
      End
      Begin VB.TextBox txtPACity 
         Height          =   360
         Left            =   120
         MaxLength       =   50
         TabIndex        =   53
         Tag             =   "UCASE"
         Top             =   2685
         Width           =   3255
      End
      Begin VB.TextBox txtPAStreet 
         Height          =   360
         Left            =   2640
         MaxLength       =   150
         TabIndex        =   51
         Tag             =   "UCASE"
         Top             =   1965
         Width           =   8535
      End
      Begin VB.ComboBox cboPAState 
         Height          =   360
         Left            =   3480
         Style           =   2  'Dropdown List
         TabIndex        =   55
         Top             =   2685
         Width           =   3495
      End
      Begin VB.TextBox txtMAOtherPostCode 
         Height          =   360
         Left            =   9480
         MaxLength       =   20
         TabIndex        =   80
         Tag             =   "UCASE"
         Top             =   4200
         Width           =   1695
      End
      Begin VB.TextBox txtMAZIP4 
         Height          =   360
         Left            =   8280
         MaxLength       =   4
         TabIndex        =   76
         Tag             =   "ZipCode4"
         Top             =   4200
         Width           =   1095
      End
      Begin VB.TextBox txtMAZIP 
         Height          =   360
         Left            =   7080
         MaxLength       =   5
         TabIndex        =   72
         Tag             =   "ZipCode5"
         Top             =   4200
         Width           =   1095
      End
      Begin VB.TextBox txtMACity 
         Height          =   360
         Left            =   120
         MaxLength       =   50
         TabIndex        =   67
         Tag             =   "UCASE"
         Top             =   4200
         Width           =   3255
      End
      Begin VB.TextBox txtMAStreet 
         Height          =   360
         Left            =   2640
         MaxLength       =   150
         TabIndex        =   65
         Tag             =   "UCASE"
         Top             =   3480
         Width           =   8535
      End
      Begin VB.ComboBox cboMAState 
         Height          =   360
         Left            =   3480
         Style           =   2  'Dropdown List
         TabIndex        =   69
         Top             =   4200
         Width           =   3495
      End
      Begin VB.TextBox txtHomePhone 
         Height          =   360
         Left            =   120
         MaxLength       =   50
         TabIndex        =   45
         Tag             =   "UCASE"
         Top             =   1200
         Width           =   4935
      End
      Begin VB.TextBox txtBusinessPhone 
         Height          =   360
         Left            =   5160
         MaxLength       =   50
         TabIndex        =   47
         Tag             =   "UCASE"
         Top             =   1200
         Width           =   6015
      End
      Begin VB.TextBox txtInsured 
         Height          =   360
         Left            =   120
         MaxLength       =   100
         TabIndex        =   43
         Tag             =   "UCASE"
         Top             =   480
         Width           =   11055
      End
      Begin VB.Label lblMAStreet 
         Caption         =   "Mailing Street:"
         Height          =   255
         Left            =   2640
         TabIndex        =   63
         Top             =   3240
         Width           =   2535
      End
      Begin VB.Label lblPAStreet 
         Caption         =   "Property Street:"
         Height          =   255
         Left            =   2640
         TabIndex        =   49
         Top             =   1725
         Width           =   2535
      End
      Begin VB.Label lblPAOtherPostCode 
         Caption         =   "Other Code:"
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
         Left            =   9480
         TabIndex        =   60
         Top             =   2445
         Width           =   1695
      End
      Begin VB.Label lblPAZIP4 
         Caption         =   "ZIP- 4:"
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
         Left            =   8280
         TabIndex        =   58
         Top             =   2445
         Width           =   1095
      End
      Begin VB.Label lblPAZIP 
         Caption         =   "Property ZIP:"
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
         Left            =   7080
         TabIndex        =   56
         Top             =   2445
         Width           =   1095
      End
      Begin VB.Label lblPACity 
         Caption         =   "Property City:"
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   2445
         Width           =   3255
      End
      Begin VB.Label lblPA 
         Caption         =   "Property Address:"
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   1725
         Width           =   2535
      End
      Begin VB.Label lblPAState 
         Caption         =   "Property State:"
         Height          =   255
         Left            =   3480
         TabIndex        =   54
         Top             =   2445
         Width           =   3495
      End
      Begin VB.Label lblMAOtherPostCode 
         Caption         =   "Other Code:"
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
         Left            =   9480
         TabIndex        =   78
         Top             =   3960
         Width           =   1695
      End
      Begin VB.Label lblMAZIP4 
         Caption         =   "ZIP- 4:"
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
         Left            =   8280
         TabIndex        =   74
         Top             =   3960
         Width           =   1095
      End
      Begin VB.Label lblMAZIP 
         Caption         =   "Mailing ZIP:"
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
         Left            =   7080
         TabIndex        =   70
         Top             =   3960
         Width           =   1095
      End
      Begin VB.Label lblMACity 
         Caption         =   "Mailing City:"
         Height          =   255
         Left            =   120
         TabIndex        =   66
         Top             =   3960
         Width           =   3255
      End
      Begin VB.Label lblMA 
         Caption         =   "Mailing Address:"
         Height          =   255
         Left            =   120
         TabIndex        =   62
         Top             =   3240
         Width           =   2535
      End
      Begin VB.Label lblMAState 
         Caption         =   "Mailing State:"
         Height          =   255
         Left            =   3480
         TabIndex        =   68
         Top             =   3960
         Width           =   3495
      End
      Begin VB.Label lblHomePhone 
         Caption         =   "Home Phone: "
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   960
         Width           =   4815
      End
      Begin VB.Label lblBusinessPhone 
         Caption         =   "Business Phone: "
         Height          =   255
         Left            =   5160
         TabIndex        =   46
         Top             =   960
         Width           =   5895
      End
      Begin VB.Label lblInsured 
         Caption         =   "Insured:"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   240
         Width           =   9735
      End
   End
   Begin VB.Frame framAdminComments 
      Height          =   4695
      Left            =   240
      TabIndex        =   114
      Top             =   480
      Visible         =   0   'False
      Width           =   11295
      Begin VB.TextBox txtAdminComments 
         Height          =   4095
         Left            =   120
         MaxLength       =   1000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   115
         Top             =   480
         Width           =   11055
      End
      Begin VB.Label lblAdminComments 
         Caption         =   "Admin Comments ( - 1000 character limit -): "
         Height          =   255
         Left            =   120
         TabIndex        =   120
         Top             =   240
         Width           =   9855
      End
   End
   Begin MSComctlLib.TabStrip TSClaimInfo 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   9128
      ShowTips        =   0   'False
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Claim Detail"
            Object.Tag             =   "framSpecifics"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Insured Info"
            Object.Tag             =   "framInsuredInfo"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Policy Information"
            Object.Tag             =   "framPolicyLimits"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Admin &Comments"
            Object.Tag             =   "framAdminComments"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Adjuster Info"
            Object.Tag             =   "framAdjusterInfo"
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
      Height          =   1215
      Left            =   8280
      TabIndex        =   116
      Top             =   5280
      Width           =   3375
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
         Left            =   2280
         MaskColor       =   &H00000000&
         Picture         =   "frmClaimInfo.frx":3031
         Style           =   1  'Graphical
         TabIndex        =   119
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
         Left            =   1200
         MaskColor       =   &H00000000&
         Picture         =   "frmClaimInfo.frx":333B
         Style           =   1  'Graphical
         TabIndex        =   118
         ToolTipText     =   "Exit"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdSpelling 
         Caption         =   "Spellin&g"
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
         Left            =   120
         MaskColor       =   &H00000000&
         Picture         =   "frmClaimInfo.frx":3485
         Style           =   1  'Graphical
         TabIndex        =   117
         ToolTipText     =   "Exit"
         Top             =   240
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmClaimInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mfrmClaim As frmClaim
Private mbUnloadMe As Boolean
Private msAssignmentsID As String
Private moGUI As V2ECKeyBoard.clsCarGUI
Private mitmXPLSelected As ListItem
Private moCurrentTextBox As TextBox
Private mbLoading As Boolean 'Loading Form
Private mbLoadingMe As Boolean 'Just loading data

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

Public Property Let itmXPLSelected(pitmX As ListItem)
    On Error GoTo EH
    Set mitmXPLSelected = pitmX
    PopulatePolicyLimitControls
    If Not mitmXPLSelected Is Nothing Then
        cmdUpdatePLClassTypeID.Enabled = True
    Else
        cmdUpdatePLClassTypeID.Enabled = False
    End If
    Exit Property
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Property Let itmXPLSelected"
End Property
Public Property Set itmXPLSelected(pitmX As ListItem)
    On Error GoTo EH
    Set mitmXPLSelected = pitmX
    PopulatePolicyLimitControls
    If Not mitmXPLSelected Is Nothing Then
        cmdUpdatePLClassTypeID.Enabled = True
    Else
        cmdUpdatePLClassTypeID.Enabled = False
    End If
    Exit Property
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Property Set itmXPLSelected"
End Property
Public Property Get itmXPLSelected() As ListItem
    Set itmXPLSelected = mitmXPLSelected
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

Private Sub cboACID_Change()
    cmdSave.Enabled = True
End Sub

Private Sub cboACID_Click()
    cmdSave.Enabled = True
End Sub

Private Sub cboACID_GotFocus()
    Set CurrentTextBox = Nothing
End Sub

Private Sub cboACIDDisplay_Change()
    cmdSave.Enabled = True
End Sub

Private Sub cboACIDDisplay_Click()
    cmdSave.Enabled = True
End Sub

Private Sub cboACIDDisplay_GotFocus()
    Set CurrentTextBox = Nothing
End Sub

Private Sub cboAddPLAppClassTypeID_GotFocus()
    Set CurrentTextBox = Nothing
End Sub

Private Sub cboAssignmentLossReportFormat_Click()
    On Error GoTo EH
    Dim sLRFormat As String
    Dim MyadoRSAssignments As ADODB.Recordset
    Dim AssLRFormat As String
    
    If mbLoading Then
        Exit Sub
    End If
    
    
    Set MyadoRSAssignments = mfrmClaim.adoRSAssignments
    
    sLRFormat = cboAssignmentLossReportFormat.Text
    AssLRFormat = goUtil.IsNullIsVbNullString(MyadoRSAssignments.Fields("LRFormat"))
    
    If StrComp(sLRFormat, "(--CHANGE FORMAT--)", vbTextCompare) = 0 Then
        If InStr(1, AssLRFormat, "TEXT", vbTextCompare) > 0 Then
            lblSelFormat.Caption = "Current Format - TEXT"
        ElseIf InStr(1, AssLRFormat, "OLEType_pdf", vbTextCompare) > 0 Then
            lblSelFormat.Caption = "Current Format - PDF (Adobe)"
        ElseIf Trim(AssLRFormat) <> vbNullString Then
            lblSelFormat.Caption = "Current Format - " & Trim(AssLRFormat)
        End If
        cmdlAttachPDFLossReportPath.Enabled = False
        cmdBuildTextLossReport.Enabled = False
        cmdlAttachPDFLossReportPath.Enabled = False
        txtLossReport.Locked = True
        Exit Sub
    End If
    
    If InStr(1, sLRFormat, "TEXT", vbTextCompare) > 0 Then
        txtLossReport.Locked = False
        cmdBuildTextLossReport.Enabled = True
        cmdAttachPDFLossReport.Enabled = False
        cmdlAttachPDFLossReportPath.Enabled = False
    ElseIf InStr(1, sLRFormat, "PDF (Adobe)", vbTextCompare) > 0 Then
        cmdlAttachPDFLossReportPath.Enabled = True
        If goUtil.utFileExists(txtAttachPDFLossReportPath.Text) Then
            cmdAttachPDFLossReport.Enabled = True
        End If
        cmdBuildTextLossReport.Enabled = False
        txtLossReport.Locked = True
    End If

    
    Set MyadoRSAssignments = Nothing
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cboAssignmentLossReportFormat_Click"
End Sub


Private Sub cboAssignmentType_Change()
    cmdSave.Enabled = True
End Sub

Private Sub cboAssignmentType_Click()
    cmdSave.Enabled = True
End Sub

Private Sub cboAssignmentType_GotFocus()
    Set CurrentTextBox = Nothing
End Sub

Private Sub cboCatCode_Change()
    cmdSave.Enabled = True
End Sub

Private Sub cboCatCode_Click()
    cmdSave.Enabled = True
End Sub

Private Sub cboCatCode_GotFocus()
    Set CurrentTextBox = Nothing
End Sub

Private Sub cboMAState_Change()
    cmdSave.Enabled = True
End Sub

Private Sub cboMAState_Click()
    cmdSave.Enabled = True
End Sub

Private Sub cboMAState_GotFocus()
    Set CurrentTextBox = Nothing
End Sub

Private Sub cboPAState_Change()
    cmdSave.Enabled = True
End Sub

Private Sub cboPAState_Click()
    cmdSave.Enabled = True
End Sub

Private Sub cboPAState_GotFocus()
    Set CurrentTextBox = Nothing
End Sub

Private Sub cboTypeOfLoss_Change()
    cmdSave.Enabled = True
End Sub

Private Sub cboTypeOfLoss_Click()
    cmdSave.Enabled = True
End Sub

Private Sub cboTypeOfLoss_GotFocus()
    Set CurrentTextBox = Nothing
End Sub

Private Sub chkDelOrigPDF_Click()
    On Error GoTo EH
    If mbLoading Then
        Exit Sub
    End If
    Dim bDelOrigPDF As Boolean
    
    If chkDelOrigPDF.Value = vbChecked Then
        bDelOrigPDF = True
    Else
        bDelOrigPDF = False
    End If
    
    SaveSetting App.EXEName, "GENERAL", "DELETE_ORIG_PDF", bDelOrigPDF
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub chkDelOrigPDF_Click"
End Sub

Private Sub chkPLIsDeleted_Click()
    On Error GoTo EH
    
    If chkPLIsDeleted.Value = vbChecked Then
        chkPLIsDeleted.Picture = imgPolicyLimits.ListImages(GuiPolicyLimitsPic.IsDeleted).Picture
    Else
        chkPLIsDeleted.Picture = LoadPicture()
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub chkPLIsDeleted_Click"
End Sub

Private Sub chkUseMailing_Click()
    On Error GoTo EH
    Dim adoRSAssignments As ADODB.Recordset
    Dim lCount As Long
    Dim sState As String
    Dim sMess As String
    Static bClicked As Boolean
    
    If bClicked Then
        Exit Sub
    End If
    
    If chkUseMailing.Value = vbChecked Then
        sMess = "Are you sure you want to use Mailing Address for Property Address? "
        If MsgBox(sMess, vbQuestion + vbYesNo, "Use Mailing Address") = vbNo Then
            bClicked = True
            chkUseMailing.Value = vbUnchecked
            bClicked = False
            Exit Sub
        End If
        txtPAStreet.Text = txtMAStreet.Text
        txtPACity.Text = txtMACity.Text
        cboPAState.Text = cboMAState.Text
        txtPAZIP.Text = Format(txtMAZIP.Text, "00000")
        txtPAZIP4.Text = Format(txtMAZIP4.Text, "0000")
        txtPAOtherPostCode.Text = txtMAOtherPostCode.Text
    Else
        Set adoRSAssignments = mfrmClaim.adoRSAssignments
        txtPAStreet.Text = goUtil.IsNullIsVbNullString(adoRSAssignments.Fields("PAStreet"))
        txtPACity.Text = goUtil.IsNullIsVbNullString(adoRSAssignments.Fields("PACity"))
        sState = goUtil.IsNullIsVbNullString(adoRSAssignments.Fields("PAState"))
        For lCount = 0 To cboPAState.ListCount - 1
            If StrComp(sState, left(cboPAState.List(lCount), 2), vbTextCompare) = 0 Then
                cboPAState.ListIndex = lCount
                Exit For
            End If
        Next
        txtPAZIP.Text = Format(goUtil.IsNullIsVbNullString(adoRSAssignments.Fields("PAZIP")), "00000")
        txtPAZIP4.Text = Format(goUtil.IsNullIsVbNullString(adoRSAssignments.Fields("PAZIP4")), "0000")
        txtPAOtherPostCode.Text = goUtil.IsNullIsVbNullString(adoRSAssignments.Fields("PAOtherPostCode"))
    End If
    
    'cleanup
    Set adoRSAssignments = Nothing
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub chkUseMailing_Click"
End Sub

Private Sub chkUseProperty_Click()
    On Error GoTo EH
    Dim adoRSAssignments As ADODB.Recordset
    Dim lCount As Long
    Dim sState As String
    Dim sMess As String
    Static bClicked As Boolean
    
    If bClicked Then
        Exit Sub
    End If
    
    If chkUseProperty.Value = vbChecked Then
        sMess = "Are you sure you want to use Property Address for Mailing Address? "
        If MsgBox(sMess, vbQuestion + vbYesNo, "Use Property Address") = vbNo Then
            bClicked = True
            chkUseProperty.Value = vbUnchecked
            bClicked = False
            Exit Sub
        End If
        txtMAStreet.Text = txtPAStreet.Text
        txtMACity.Text = txtPACity.Text
        cboMAState.Text = cboPAState.Text
        txtMAZIP.Text = Format(txtPAZIP.Text, "00000")
        txtMAZIP4.Text = Format(txtPAZIP4.Text, "0000")
        txtMAOtherPostCode.Text = txtPAOtherPostCode.Text
    Else
        Set adoRSAssignments = mfrmClaim.adoRSAssignments
        txtMAStreet.Text = goUtil.IsNullIsVbNullString(adoRSAssignments.Fields("MAStreet"))
        txtMACity.Text = goUtil.IsNullIsVbNullString(adoRSAssignments.Fields("MACity"))
        sState = goUtil.IsNullIsVbNullString(adoRSAssignments.Fields("MAState"))
        For lCount = 0 To cboMAState.ListCount - 1
            If StrComp(sState, left(cboMAState.List(lCount), 2), vbTextCompare) = 0 Then
                cboMAState.ListIndex = lCount
                Exit For
            End If
        Next
        txtMAZIP.Text = Format(goUtil.IsNullIsVbNullString(adoRSAssignments.Fields("MAZIP")), "00000")
        txtMAZIP4.Text = Format(goUtil.IsNullIsVbNullString(adoRSAssignments.Fields("MAZIP4")), "0000")
        txtMAOtherPostCode.Text = goUtil.IsNullIsVbNullString(adoRSAssignments.Fields("MAOtherPostCode"))
    End If
    
    'cleanup
    Set adoRSAssignments = Nothing
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub chkUseProperty_Click"
End Sub

Private Sub cmdAddPLAppClassTypeID_Click()
    On Error GoTo EH
    Dim sMess As String
    Dim sSelectID As String
    
    'make sure there is an item avail to add
    If cboAddPLAppClassTypeID.ListCount > 0 Then
        If cboAddPLAppClassTypeID.ListIndex = -1 Then
            sMess = "Choose an item from the list first!"
        End If
    Else
        sMess = "There are no more items!"
    End If
    
    If sMess <> vbNullString Then
        MsgBox sMess, vbExclamation + vbOKOnly, "Add"
        Exit Sub
    End If
    
    If AddPLAppClassTypeID(sSelectID) Then
        LoadPolicyLimitsStuff
        lvwPLClassTypeID.ListItems("""" & sSelectID & """").Selected = True
        If Not lvwPLClassTypeID.SelectedItem Is Nothing Then
            Set itmXPLSelected = lvwPLClassTypeID.SelectedItem
            txtPLLimitAmount.SetFocus
        End If
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdAddPLAppClassTypeID_Click"
End Sub

Public Function AddPLAppClassTypeID(psSelectID As String) As Boolean
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim adoRS As ADODB.Recordset
    Dim sSQL As String
    Dim sAssignmentsID As String
    Dim sID As String
    Dim sIDAssignments As String
    Dim sClassTypeID As String
    Dim sLimitAmount As String
    Dim sRCSaidProp As String
    Dim sReserves As String
    Dim sIsDeleted As String
    Dim sDownLoadMe As String
    Dim sUpLoadMe As String
    Dim sAdminComments As String
    Dim sDateLastUpdated As String
    Dim sUID As String
    
    Set oConn = New ADODB.Connection
    Set adoRS = New ADODB.Recordset
    
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    'Get field data from Query
    sSQL = "SELECT A.AssignmentsID, "
    'When INSERTING Records On Client SIDE ALWAYS USE
    'NEGATIVE NUMBERS, This Will indicate it needs to be Synched wth Server
    'If the CLIENT ID Is not Negative that means it was Synched with Server
    sSQL = sSQL & goUtil.GetAccessDBUID("ID", "PolicyLimits") & " AS ID, "
    sSQL = sSQL & "A.ID As IDAssignments, "
    sSQL = sSQL & cboAddPLAppClassTypeID.ItemData(cboAddPLAppClassTypeID.ListIndex) & " As ClassTypeID, "
    sSQL = sSQL & " 0.00 As LimitAmount, "
    sSQL = sSQL & " 0.00 As RCSaidProp, "
    sSQL = sSQL & " 0.00 As Reserves, "
    sSQL = sSQL & " False As IsDeleted, "
    sSQL = sSQL & " False As DownLoadMe, "
    sSQL = sSQL & " True As UpLoadMe, "
    sSQL = sSQL & "'' As AdminComments, "
    sSQL = sSQL & "#" & Now() & "# As DateLastUpdated, "
    sSQL = sSQL & goUtil.gsCurUsersID & " As UID "
    sSQL = sSQL & "FROM Assignments A "
    sSQL = sSQL & "WHERE A.AssignmentsID = " & msAssignmentsID & " "
    
    'Use Disconnected Record Set on asUseClient Cusor ONLY !
    adoRS.CursorLocation = adUseClient
    adoRS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set adoRS.ActiveConnection = Nothing
    'Populate the Fields to be inserted into Policy Limits
    
    If Not adoRS.EOF Then
        adoRS.MoveFirst
        sAssignmentsID = goUtil.IsNullIsVbNullString(adoRS.Fields("AssignmentsID"))
        sID = goUtil.IsNullIsVbNullString(adoRS.Fields("ID"))
        psSelectID = sID ' Select this ID in list view
        sIDAssignments = goUtil.IsNullIsVbNullString(adoRS.Fields("IDAssignments"))
        sClassTypeID = goUtil.IsNullIsVbNullString(adoRS.Fields("ClassTypeID"))
        sLimitAmount = Format(goUtil.IsNullIsVbNullString(adoRS.Fields("LimitAmount")), "#,###,###,##0.00")
        sRCSaidProp = Format(goUtil.IsNullIsVbNullString(adoRS.Fields("RCSaidProp")), "#,###,###,##0.00")
        sReserves = Format(goUtil.IsNullIsVbNullString(adoRS.Fields("Reserves")), "#,###,###,##0.00")
        sIsDeleted = goUtil.IsNullIsVbNullString(adoRS.Fields("IsDeleted"))
        sDownLoadMe = goUtil.IsNullIsVbNullString(adoRS.Fields("DownLoadMe"))
        sUpLoadMe = goUtil.IsNullIsVbNullString(adoRS.Fields("UpLoadMe"))
        sAdminComments = goUtil.IsNullIsVbNullString(adoRS.Fields("AdminComments"))
        sDateLastUpdated = goUtil.IsNullIsVbNullString(adoRS.Fields("DateLastUpdated"))
        sUID = goUtil.IsNullIsVbNullString(adoRS.Fields("UID"))
    End If
    
    'When adding an item the DB will be updated real time, NOT WHEN THE FORM EXITS
    'This way other forms can access this particular data while the Claim Info Form
    'is still open.
    sSQL = "INSERT INTO PolicyLimits "
    sSQL = sSQL & "( "
    sSQL = sSQL & "[PolicyLimitsID], "
    sSQL = sSQL & "[AssignmentsID] , "
    sSQL = sSQL & "[ID],"
    sSQL = sSQL & "[IDAssignments], "
    sSQL = sSQL & "[ClassTypeID] , "
    sSQL = sSQL & "[LimitAmount] , "
    sSQL = sSQL & "[RCSaidProp] , "
    sSQL = sSQL & "[Reserves] , "
    sSQL = sSQL & "[IsDeleted] , "
    sSQL = sSQL & "[DownLoadMe], "
    sSQL = sSQL & "[UpLoadMe], "
    sSQL = sSQL & "[AdminComments], "
    sSQL = sSQL & "[DateLastUpdated], "
    sSQL = sSQL & "[UpdateByUserID] "
    sSQL = sSQL & ") "
    sSQL = sSQL & "SELECT "
    sSQL = sSQL & sID & " As [PolicyLimitsID], "
    sSQL = sSQL & sAssignmentsID & " AS [AssignmentsID] , "
    sSQL = sSQL & sID & " As [ID], "
    sSQL = sSQL & sIDAssignments & " As [IDAssignments], "
    sSQL = sSQL & sClassTypeID & " As [ClassTypeID] , "
    sSQL = sSQL & CCur(sLimitAmount) & " As [LimitAmount] , "
    sSQL = sSQL & CCur(sRCSaidProp) & " As [RCSaidProp] , "
    sSQL = sSQL & CCur(sReserves) & " As [Reserves] , "
    sSQL = sSQL & sIsDeleted & " As [IsDeleted] , "
    sSQL = sSQL & sDownLoadMe & " As [DownLoadMe], "
    sSQL = sSQL & sUpLoadMe & " As [UpLoadMe], "
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(sAdminComments) & "' As [AdminComments], "
    sSQL = sSQL & "#" & sDateLastUpdated & "# As [DateLastUpdated], "
    sSQL = sSQL & sUID & " As [UpdateByUserID] "
    
    oConn.Execute sSQL
    
    
    AddPLAppClassTypeID = True
    
    'cleanup
    Set oConn = Nothing
    Set adoRS = Nothing
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function AddPLAppClassTypeID"
End Function

Public Function EditPLAppClassTypeID() As Boolean

    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    Dim sAssignmentsID As String
    Dim sID As String
    Dim sIDAssignments As String
    Dim sClassTypeID As String
    Dim sLimitAmount As String
    Dim sRCSaidProp As String
    Dim sReserves As String
    Dim sIsDeleted As String
    Dim sDownLoadMe As String
    Dim sUpLoadMe As String
    Dim sAdminComments As String
    Dim sDateLastUpdated As String
    Dim sUID As String
    Dim bDeleteRecord As Boolean
    
    
    If mitmXPLSelected Is Nothing Then
        Exit Function
    End If
    
    sAssignmentsID = msAssignmentsID
    sID = mitmXPLSelected.SubItems(GuiPolicyLimits.ID - 1)
    sIDAssignments = msAssignmentsID
    sClassTypeID = mitmXPLSelected.SubItems(GuiPolicyLimits.ClassTypeID - 1)
    sLimitAmount = txtPLLimitAmount.Text
    sRCSaidProp = txtPLRCSaidProp.Text
    sReserves = txtPLReserves.Text
    If chkPLIsDeleted.Value = vbChecked Then
        sIsDeleted = "-1"
        'if this record has never been uploaded and it
        'is marked for Delettion then really get rid of it
        If CDbl(sID) < 0 Then
            bDeleteRecord = True
        End If
    Else
        sIsDeleted = "0"
    End If
    sDownLoadMe = "0"
    sUpLoadMe = "-1"
    sAdminComments = txtPLAdminComments.Text
    sDateLastUpdated = Now()
    sUID = goUtil.gsCurUsersID
    
    Set oConn = New ADODB.Connection
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    If bDeleteRecord Then
        sSQL = " DELETE FROM PolicyLimits "
        sSQL = sSQL & "WHERE   PolicyLimitsID = " & sID & " "
    Else
        sSQL = "Update PolicyLimits Set "
        sSQL = sSQL & "[AssignmentsID] = " & sAssignmentsID & " , "
        sSQL = sSQL & "[ID] = " & sID & ", "
        sSQL = sSQL & "[IDAssignments] = " & sIDAssignments & ", "
        sSQL = sSQL & "[ClassTypeID] = " & sClassTypeID & ", "
        sSQL = sSQL & "[LimitAmount] = " & CCur(sLimitAmount) & ", "
        sSQL = sSQL & "[RCSaidProp] = " & CCur(sRCSaidProp) & ", "
        sSQL = sSQL & "[Reserves] = " & CCur(sReserves) & ", "
        sSQL = sSQL & "[IsDeleted] =  " & sIsDeleted & ", "
        sSQL = sSQL & "[DownLoadMe] = " & sDownLoadMe & ", "
        sSQL = sSQL & "[UpLoadMe] = " & sUpLoadMe & ", "
        sSQL = sSQL & "[AdminComments] = '" & goUtil.utCleanSQLString(sAdminComments) & "', "
        sSQL = sSQL & "[DateLastUpdated] = #" & sDateLastUpdated & "#, "
        sSQL = sSQL & "[UpdateByUserID] = " & sUID & " "
        sSQL = sSQL & "WHERE   PolicyLimitsID = " & sID & " "
    End If
    oConn.Execute sSQL
                    
    EditPLAppClassTypeID = True
    
    'cleanup
    Set oConn = Nothing
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function EditPLAppClassTypeID"
End Function

Private Sub cmdAssignedDate_Click()
    mfrmClaim.ShowCalendar txtAssignedDate
    CheckMyDateWhenCloseDateSet txtAssignedDate
End Sub

Private Sub cmdAttachPDFLossReport_Click()
    On Error GoTo EH
    Dim sPDFAttachPath As String
    Dim sIBNUM As String
    Dim sYYMMDDHHMMSS As String
    Dim sUsersID As String
    Dim sLRFormat As String
    Dim sPDFFileName As String
    Dim sNewAttachPath As String
    Dim sMess As String
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    
    
    Screen.MousePointer = vbHourglass
    
    'Build the PdfAttachment File name
    sIBNUM = txtIBNUM.Text
    sYYMMDDHHMMSS = Format(Now(), "YYMMDDHHMMSS")
    sUsersID = goUtil.gsCurUsersID
    
    'create the PDF FIle name (Example '"FRE21220_040709090032_1.pdf"
    'Note those pdf attachments uploaded via website will look like this ...
    'FRE21220_040709090032_1@3545969.pdf
    'the @#########  is Token id created by Cold Fusion...
    'Flags that it was created On the Web Site vs Easy Claim client
    'Easy Claim DOES NOT USE THE TOKEN ID !!!
    sPDFFileName = sIBNUM & "_" & sYYMMDDHHMMSS & "_" & sUsersID & ".pdf"
    
    'Set the Format
    'Note those pdf attachments  Format Name uploaded via website will look like this ...
    'OLEType_pdf_3545969_160870156
    'The first set of number is Token ID created by Cold Fusion
    'The Second set of numbers is the tickcount created by the Server (number of seconds since Serverf system reboot)
    'Flags that it was created On the Web Site vs Easy Claim client
    sLRFormat = "OLEType_pdf" ' Easy Claim client DOES NOT USE THE SPECIAL FORMAT NUmbers
    
    'Get the File path to the raw file that needs to be attached
    sPDFAttachPath = txtAttachPDFLossReportPath.Text
    
    If Not goUtil.utFileExists(sPDFAttachPath) Then
        Screen.MousePointer = vbNormal
        MsgBox "Can't find " & sPDFAttachPath, vbCritical + vbOKOnly, "Error reading file"
        Exit Sub
    End If
    
    'The Attachment path will be in AttachRepos under install dir
    'see if it exists, build it if it does not
    sNewAttachPath = goUtil.gsInstallDir & "\AttachRepos\"
    If Not goUtil.utFileExists(sNewAttachPath, True) Then
        goUtil.utMakeDir sNewAttachPath
    End If
    
    'Add the file to sNewAttachPath
    sNewAttachPath = sNewAttachPath & sPDFFileName
    
    'Copy the file over and then update the DB
    sMess = goUtil.utCopyFile(sPDFAttachPath, sNewAttachPath)
    
    If Not sMess = vbNullString Then
        Screen.MousePointer = vbNormal
        MsgBox "Error Attaching file!" & vbCrLf & vbCrLf & sMess, vbCritical + vbOKOnly, "Error"
        Exit Sub
    End If
    
    Set oConn = New ADODB.Connection
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    sSQL = "UPDATE Assignments SET "
    sSQL = sSQL & "LRFormat = '" & sLRFormat & "', "
    sSQL = sSQL & "LossReport = '" & sPDFFileName & "', "
    sSQL = sSQL & "UploadLossReport = True, "
    sSQL = sSQL & "UpLoadMe = True, "
    sSQL = sSQL & "DateLastUpdated = #" & Now() & "#, "
    sSQL = sSQL & "UpdateByUserID = " & goUtil.gsCurUsersID & " "
    sSQL = sSQL & "WHERE ID = " & msAssignmentsID
    
    oConn.Execute sSQL
    
    DoEvents
    Sleep 2000 'Need to let Db catch up with the Update
    
    If SaveMe Then
        'IF the Delete Original PDF After Attach check box
        'is checked then remove the original PDF
        If chkDelOrigPDF.Value = vbChecked Then
            goUtil.utDeleteFile sPDFAttachPath
        End If
        txtAttachPDFLossReportPath.Text = vbNullString
        cmdAttachPDFLossReport.Enabled = False
        LoadMe
        cmdSave.Enabled = False
        mfrmClaim.RefreshMe
    End If
    
    Screen.MousePointer = vbNormal
    
    Set oConn = Nothing
    Exit Sub
EH:
    Screen.MousePointer = vbNormal
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdAttachPDFLossReport_Click"
End Sub

Private Sub cmdBuildTextLossReport_Click()
    On Error GoTo EH
    Dim sAdditionalInfo As String
    Dim sLossReportText As String
    Dim lPos As Long
    'Need to get anything that is below ----ADDITIONAL INFO---- '
    'and concat it on to the bottom of the Build text
    sLossReportText = txtLossReport.Text
    lPos = InStr(1, sLossReportText, "----ADDITIONAL INFO----", vbTextCompare)
    If lPos > 0 Then
        lPos = InStr(lPos, sLossReportText, vbCrLf, vbBinaryCompare)
        If lPos > 0 Then
            sAdditionalInfo = Mid(sLossReportText, lPos + 4)
        End If
    End If

    txtLossReport.Text = BuildTextLossReport & sAdditionalInfo
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdBuildTextLossReport_Click"
End Sub

Private Sub cmdCloseDate_Click()
    mfrmClaim.ShowCalendar txtCloseDate
    CheckMyDateWhenCloseDateSet txtCloseDate
End Sub

Private Sub cmdContactDate_Click()
    mfrmClaim.ShowCalendar txtContactDate
    CheckMyDateWhenCloseDateSet txtContactDate
End Sub

Private Sub cmdInspectedDate_Click()
    mfrmClaim.ShowCalendar txtInspectedDate
    CheckMyDateWhenCloseDateSet txtInspectedDate
End Sub

Private Sub cmdlAttachPDFLossReportPath_Click()
    On Error GoTo EH
    Dim sPath As String
    Dim sSelFile As String
    Dim lFileSize As Long
    Dim sMyFilter As String
    
    sMyFilter = sMyFilter & "PDF Document File" & " (*." & "pdf" & ")" & SD & "*." & "pdf" & SD
   
    
    sPath = goUtil.utGetPath(App.EXEName, "PDFDocumentFile", "Browse to the PDF Document File you want to attach", "", sPath, Me.hWnd, sMyFilter, sSelFile)
    
    If goUtil.utFileExists(sPath & sSelFile) Then
        If StrComp(sPath, goUtil.AttachReposPath, vbTextCompare) = 0 Then
            MsgBox "Can't use this directory for attaching files!", vbExclamation + vbOKOnly, "INVALID DIRECTORY!"
            cmdAttachPDFLossReport.Enabled = False
            txtAttachPDFLossReportPath.Text = vbNullString
            Exit Sub
        End If
        cmdAttachPDFLossReport.Enabled = True
        txtAttachPDFLossReportPath.Text = sPath & sSelFile
    Else
        cmdAttachPDFLossReport.Enabled = False
        txtAttachPDFLossReportPath.Text = vbNullString
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdlAttachPDFLossReportPath_Click"
End Sub

Private Sub cmdLossDate_Click()
    mfrmClaim.ShowCalendar txtLossDate
    CheckMyDateWhenCloseDateSet txtLossDate
End Sub

Private Sub cmdPrintList_Click()
    On Error GoTo EH
    goUtil.utPrintListView App.EXEName, lvwPLClassTypeID, "Policy Limits"
     Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdPrintList_Click"
End Sub

Private Sub cmdReceivedDate_Click()
    mfrmClaim.ShowCalendar txtReceivedDate
    CheckMyDateWhenCloseDateSet txtReceivedDate
End Sub

Private Sub cmdSave_Click()
    On Error GoTo EH
    If cmdUpdatePLClassTypeID.Enabled Then
        cmdUpdatePLClassTypeID_Click
    End If
    If SaveMe Then
        mfrmClaim.RefreshMe
        cmdSave.Enabled = False
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
            DoEvents
            Sleep 100
            cmdSpelling.Enabled = False
        End If
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdSpelling_Click"
End Sub

Private Sub cmdUpdatePLClassTypeID_Click()
    On Error GoTo EH
    
    goUtil.utValidate , txtPLLimitAmount
    goUtil.utValidate , txtPLRCSaidProp
    goUtil.utValidate , txtPLReserves
    If EditPLAppClassTypeID Then
        LoadPolicyLimitsStuff
        cmdUpdatePLClassTypeID.Enabled = False
        If cmdAddPLAppClassTypeID.Visible And cmdAddPLAppClassTypeID.Enabled Then
            cmdAddPLAppClassTypeID.SetFocus
        End If
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdUpdatePLClassTypeID_Click"
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

Private Sub lvwPLClassTypeID_Click()
    On Error GoTo EH
    
    'Set the selected claim
    itmXPLSelected = lvwPLClassTypeID.SelectedItem
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lvwPLClassTypeID_Click"
End Sub

Private Sub lvwPLClassTypeID_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    On Error GoTo EH
    Dim sMess As String
    If lvwPLClassTypeID.SortOrder = lvwAscending Then
        lvwPLClassTypeID.SortOrder = lvwDescending
        sMess = "Descending"
    Else
        lvwPLClassTypeID.SortOrder = lvwAscending
        sMess = "Ascending"
    End If
    
    'Set the Tool Tip
    lvwPLClassTypeID.ToolTipText = "Sort By " & ColumnHeader.Text & " " & sMess
    
    'Need to see if the Column clicked was a NON text sort
    'Like Date or number.  If a non textual column was clicked
    'need to use the next column as the sort since this next column
    'will be hidden and contain a sort friendly format.
    'Sort Key is Base 0 where CoulmnHeader is not
    Select Case ColumnHeader.Index
        Case GuiPolicyLimits.LimitAmount, GuiPolicyLimits.RCSaidProp, GuiPolicyLimits.Reserves
            lvwPLClassTypeID.SortKey = ColumnHeader.Index
        Case GuiPolicyLimits.DateLastUpdated
            lvwPLClassTypeID.SortKey = ColumnHeader.Index
        Case Else
            lvwPLClassTypeID.SortKey = ColumnHeader.Index - 1
    End Select
    
    lvwPLClassTypeID.Sorted = True
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lvwPLClassTypeID_ColumnClick"
End Sub

Private Sub lvwPLClassTypeID_GotFocus()
    Set CurrentTextBox = Nothing
End Sub

Private Sub lvwPLClassTypeID_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo EH
    
    'Set the selected claim
    If Not lvwPLClassTypeID.SelectedItem Is Nothing Then
        itmXPLSelected = lvwPLClassTypeID.SelectedItem
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lvwPLClassTypeID_KeyUp"
End Sub


Private Sub txtAdminComments_Change()
    cmdSave.Enabled = True
End Sub

Private Sub txtAdminComments_GotFocus()
    goUtil.utSelText txtAdminComments
    Set CurrentTextBox = txtAdminComments
End Sub

Private Sub txtAdminComments_LostFocus()
    goUtil.utValidate , txtAdminComments
End Sub

Private Sub txtAgentNo_Change()
    cmdSave.Enabled = True
End Sub

Private Sub txtAgentNo_GotFocus()
    goUtil.utSelText txtAgentNo
        Set CurrentTextBox = txtAgentNo
End Sub

Private Sub txtAgentNo_LostFocus()
    goUtil.utValidate , txtAgentNo
End Sub

Private Sub txtAssignedDate_Change()
    cmdSave.Enabled = True
End Sub

Private Sub txtAssignedDate_GotFocus()
    goUtil.utSelText txtAssignedDate
    Set CurrentTextBox = Nothing
End Sub

Private Sub txtAssignedDate_LostFocus()
    goUtil.utValidate , txtAssignedDate
    CheckMyDateWhenCloseDateSet txtAssignedDate
End Sub

Private Sub txtAttachPDFLossReportPath_GotFocus()
    goUtil.utSelText txtAttachPDFLossReportPath
End Sub

Private Sub txtBusinessPhone_Change()
    cmdSave.Enabled = True
End Sub

Private Sub txtBusinessPhone_GotFocus()
    goUtil.utSelText txtBusinessPhone
        Set CurrentTextBox = txtBusinessPhone
End Sub

Private Sub txtBusinessPhone_LostFocus()
    goUtil.utValidate , txtBusinessPhone
End Sub

Private Sub txtCLIENTNUM_Change()
    cmdSave.Enabled = True
End Sub

Private Sub txtCLIENTNUM_GotFocus()
    goUtil.utSelText txtCLIENTNUM
    Set CurrentTextBox = Nothing
End Sub

Private Sub txtCLIENTNUM_LostFocus()
    goUtil.utValidate , txtCLIENTNUM
End Sub

Private Sub txtCloseDate_Change()
    On Error GoTo EH
    Dim sMess As String
    Dim lRet As VBA.VbMsgBoxResult
    Dim bShowCloseDateMessage As Boolean
    Dim bOtherDatesNotFilledOut As Boolean
    If Not mbLoadingMe Then
        cmdSave.Enabled = True
        
        'first Check to see if all the other dates have been filled out
        If Not IsDate(txtLossDate.Text) Then
             bOtherDatesNotFilledOut = True
        ElseIf Not IsDate(txtAssignedDate.Text) Then
             bOtherDatesNotFilledOut = True
        ElseIf Not IsDate(txtReceivedDate.Text) Then
             bOtherDatesNotFilledOut = True
        ElseIf Not IsDate(txtContactDate.Text) Then
             bOtherDatesNotFilledOut = True
        ElseIf Not IsDate(txtInspectedDate.Text) Then
             bOtherDatesNotFilledOut = True
        End If
        
        If bOtherDatesNotFilledOut And IsDate(txtCloseDate.Text) Then
            sMess = "You must fill out all other dates before the Close Date."
            MsgBox sMess, vbExclamation + vbOKOnly, "MISSING DATES"
            txtCloseDate.Text = vbNullString
            Exit Sub
        End If
        
        bShowCloseDateMessage = CBool(GetSetting(App.EXEName, "MESSAGES", "ShowCloseDateMessage", True))
        If IsDate(txtCloseDate.Text) And Not mfrmClaim.MyStatus = iAssignmentsStatus_CLOSED Then
            If bShowCloseDateMessage Then
                sMess = "Are you sure you are ready to close this Claim ?" & vbCrLf
                sMess = sMess & "You will not be able to change anything" & vbCrLf
                sMess = sMess & "after the CLOSE DATE is set and you exit this claim!" & vbCrLf & vbCrLf
                sMess = sMess & "Click ""YES"" to Close this claim." & vbCrLf
                sMess = sMess & "Click ""NO"" to not close this claim at this time." & vbCrLf
                sMess = sMess & "Click ""Cancel"" to Close this claim And not show this message again." & vbCrLf
                lRet = MsgBox(sMess, vbQuestion + vbYesNoCancel, "Close this Claim ?")
                
                If lRet = vbNo Then
                    txtCloseDate.Text = vbNullString
                ElseIf lRet = vbCancel Then
                    SaveSetting App.EXEName, "MESSAGES", "ShowCloseDateMessage", "False"
                End If
            End If
        End If
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub txtCloseDate_Change"
End Sub

Public Sub CheckMyDateWhenCloseDateSet(poDateBox As Object)
    On Error GoTo EH
    Dim oDateBox As TextBox
    Dim oOtherDate As TextBox
    Dim oControl As Control
    Dim sDateName As String
    Dim sMess As String
    
    If Not TypeOf poDateBox Is TextBox Then
        Exit Sub
    Else
        Set oDateBox = poDateBox
    End If
    
    If IsDate(txtCloseDate.Text) Then
        If Not IsDate(oDateBox.Text) Then
            sMess = "Dates can not be blank if the Close Date is Set!" & vbCrLf & vbCrLf
        End If
    End If
    
    'As well check this Date against Other Dates
    For Each oControl In Me.Controls
        If TypeOf oControl Is TextBox Then
            Set oOtherDate = oControl
            If StrComp(oOtherDate.Tag, "Date", vbTextCompare) = 0 And oOtherDate.Name <> oDateBox.Name Then
                Select Case UCase(oDateBox.Name)
                    Case UCase(txtLossDate.Name)
                        sDateName = lblLossDate.Caption
                        If Not IsDate(oDateBox.Text) Then
                            Exit For
                        End If
                        'Loss date Can't be > than any other Dates!
                        If IsDate(oOtherDate.Text) Then
                            If CDate(oDateBox.Text) > CDate(oOtherDate.Text) Then
                                sMess = sMess & sDateName & " can not be later than " & oOtherDate.Text & vbCrLf
                            End If
                        End If
                    Case UCase(txtAssignedDate.Name)
                        sDateName = lblAssignedDate.Caption
                        If Not IsDate(oDateBox.Text) Then
                            Exit For
                        End If
                        'Assigned Date Can't be > than any other Dates! excpet for Loss Date.
                        If IsDate(oOtherDate.Text) And oOtherDate.Name <> txtLossDate.Name Then
                            If (CDate(oDateBox.Text) > CDate(oOtherDate.Text)) Then
                                sMess = sMess & sDateName & " can not be later than " & oOtherDate.Text & vbCrLf
                            End If
                        End If
                        'Assigned Date Can't be < than Loss Date.
                        If IsDate(oOtherDate.Text) And oOtherDate.Name = txtLossDate.Name Then
                            If (CDate(oDateBox.Text) < CDate(oOtherDate.Text)) Then
                                sMess = sMess & sDateName & " can not be earlier than " & oOtherDate.Text & vbCrLf
                            End If
                        End If
                    Case UCase(txtReceivedDate.Name)
                        sDateName = lblReceivedDate.Caption
                        If Not IsDate(oDateBox.Text) Then
                            Exit For
                        End If
                         'Received Date Can't be > than any other Dates! excpet for Loss Date And Assigned Date.
                        If IsDate(oOtherDate.Text) And oOtherDate.Name <> txtLossDate.Name And oOtherDate.Name <> txtAssignedDate.Name Then
                            If (CDate(oDateBox.Text) > CDate(oOtherDate.Text)) Then
                                sMess = sMess & sDateName & " can not be later than " & oOtherDate.Text & vbCrLf
                            End If
                        End If
                        'Received Date Can't be < than Loss Date or Assigned Date
                        If IsDate(oOtherDate.Text) And ((oOtherDate.Name = txtLossDate.Name) Or (oOtherDate.Name = txtAssignedDate.Name)) Then
                            If (CDate(oDateBox.Text) < CDate(oOtherDate.Text)) Then
                                sMess = sMess & sDateName & " can not be earlier than " & oOtherDate.Text & vbCrLf
                            End If
                        End If
                    Case UCase(txtContactDate.Name)
                        sDateName = lblContactDate.Caption
                        If Not IsDate(oDateBox.Text) Then
                            Exit For
                        End If
                        'Contact Date  Can't be < than any other Dates! excpet for txtInspected Date And txtClose Date.
                        If IsDate(oOtherDate.Text) And oOtherDate.Name <> txtInspectedDate.Name And oOtherDate.Name <> txtCloseDate.Name Then
                            If (CDate(oDateBox.Text) < CDate(oOtherDate.Text)) Then
                                sMess = sMess & sDateName & " can not be earlier than " & oOtherDate.Text & vbCrLf
                            End If
                        End If
                        'Contact Date  Can't be > than txtInspected Date Or txtClose Date.
                        If IsDate(oOtherDate.Text) And ((oOtherDate.Name = txtInspectedDate.Name) Or (oOtherDate.Name = txtCloseDate.Name)) Then
                            If (CDate(oDateBox.Text) > CDate(oOtherDate.Text)) Then
                                sMess = sMess & sDateName & " can not be later than " & oOtherDate.Text & vbCrLf
                            End If
                        End If
                    Case UCase(txtInspectedDate.Name)
                        sDateName = lblInspectedDate.Caption
                        If Not IsDate(oDateBox.Text) Then
                            Exit For
                        End If
                        'txtInspected Date  Can't be < than any other Dates! excpet for txtClose Date.
                        If IsDate(oOtherDate.Text) And oOtherDate.Name <> txtCloseDate.Name Then
                            If (CDate(oDateBox.Text) < CDate(oOtherDate.Text)) Then
                                sMess = sMess & sDateName & " can not be earlier than " & oOtherDate.Text & vbCrLf
                            End If
                        End If
                        'Inspected Date  Can't be > than txtClose Date.
                        If IsDate(oOtherDate.Text) And oOtherDate.Name = txtCloseDate.Name Then
                            If (CDate(oDateBox.Text) > CDate(oOtherDate.Text)) Then
                                sMess = sMess & sDateName & " can not be later than " & oOtherDate.Text & vbCrLf
                            End If
                        End If
                    Case UCase(txtCloseDate.Name)
                        sDateName = lblCloseDate.Caption
                        If Not IsDate(oDateBox.Text) Then
                            Exit For
                        End If
                        'Close Date date Can't be < than any other Dates!
                        If IsDate(oOtherDate.Text) Then
                            If CDate(oDateBox.Text) < CDate(oOtherDate.Text) Then
                                sMess = sMess & sDateName & " can not be earlier than " & oOtherDate.Text & vbCrLf
                            End If
                        End If
                End Select
            End If
        End If
    Next

    If sMess <> vbNullString Then
        MsgBox sMess, vbExclamation + vbOKOnly, "Invalid Date Entry"
        oDateBox.Text = vbNullString
    End If
    'cleanup
    Set oDateBox = Nothing
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub CheckMyDateWhenCloseDateSet"
End Sub

Private Sub txtCloseDate_GotFocus()
    goUtil.utSelText txtCloseDate
    Set CurrentTextBox = Nothing
End Sub

Private Sub txtCloseDate_LostFocus()
    goUtil.utValidate , txtCloseDate
    CheckMyDateWhenCloseDateSet txtCloseDate
End Sub

Private Sub txtContactDate_Change()
    cmdSave.Enabled = True
End Sub

Private Sub txtContactDate_GotFocus()
    goUtil.utSelText txtContactDate
    Set CurrentTextBox = Nothing
End Sub

Private Sub txtContactDate_LostFocus()
    goUtil.utValidate , txtContactDate
    CheckMyDateWhenCloseDateSet txtContactDate
End Sub

Private Sub txtDeductible_Change()
    cmdSave.Enabled = True
End Sub

Private Sub txtDeductible_GotFocus()
    goUtil.utSelText txtDeductible
    Set CurrentTextBox = Nothing
End Sub

Private Sub txtDeductible_LostFocus()
    goUtil.utValidate , txtDeductible
End Sub

Private Sub txtHomePhone_Change()
    cmdSave.Enabled = True
End Sub

Private Sub txtHomePhone_GotFocus()
    goUtil.utSelText txtHomePhone
        Set CurrentTextBox = txtHomePhone
End Sub

Private Sub txtHomePhone_LostFocus()
    goUtil.utValidate , txtHomePhone
End Sub

Private Sub txtIBNUM_Change()
    cmdSave.Enabled = True
End Sub

Private Sub txtIBNUM_GotFocus()
    goUtil.utSelText txtIBNUM
    Set CurrentTextBox = Nothing
End Sub

Private Sub txtIBNUM_LostFocus()
    goUtil.utValidate , txtIBNUM
End Sub

Private Sub txtInspectedDate_Change()
    cmdSave.Enabled = True
End Sub

Private Sub txtInspectedDate_GotFocus()
    goUtil.utSelText txtInspectedDate
    Set CurrentTextBox = Nothing
End Sub

Private Sub txtInspectedDate_LostFocus()
    goUtil.utValidate , txtInspectedDate
    CheckMyDateWhenCloseDateSet txtInspectedDate
End Sub

Private Sub txtInsured_Change()
    cmdSave.Enabled = True
End Sub

Private Sub txtInsured_GotFocus()
    goUtil.utSelText txtInsured
    Set CurrentTextBox = txtInsured
End Sub

Private Sub txtInsured_LostFocus()
    goUtil.utValidate , txtInsured
End Sub

Private Sub txtLossDate_Change()
    cmdSave.Enabled = True
End Sub

Private Sub txtLossDate_GotFocus()
    goUtil.utSelText txtLossDate
    Set CurrentTextBox = Nothing
End Sub

Private Sub txtLossDate_LostFocus()
    goUtil.utValidate , txtLossDate
    CheckMyDateWhenCloseDateSet txtLossDate
End Sub

Private Sub txtLossReport_Change()
    cmdSave.Enabled = True
End Sub

Private Sub txtLossReport_GotFocus()
    Set CurrentTextBox = txtLossReport
End Sub

Private Sub txtLossReport_LostFocus()
    goUtil.utValidate , txtLossReport
End Sub

Private Sub txtMACity_Change()
    cmdSave.Enabled = True
End Sub

Private Sub txtMACity_GotFocus()
    goUtil.utSelText txtMACity
    Set CurrentTextBox = txtMACity
End Sub

Private Sub txtMACity_LostFocus()
    goUtil.utValidate , txtMACity
End Sub

Private Sub txtMAOtherPostCode_Change()
    cmdSave.Enabled = True
End Sub

Private Sub txtMAOtherPostCode_GotFocus()
    goUtil.utSelText txtMAOtherPostCode
    Set CurrentTextBox = Nothing
End Sub

Private Sub txtMAOtherPostCode_LostFocus()
    goUtil.utValidate , txtMAOtherPostCode
End Sub

Private Sub txtMAStreet_Change()
    cmdSave.Enabled = True
End Sub

Private Sub txtMAStreet_GotFocus()
    goUtil.utSelText txtMAStreet
    Set CurrentTextBox = txtMAStreet
End Sub

Private Sub txtMAStreet_LostFocus()
    goUtil.utValidate , txtMAStreet
End Sub

Private Sub txtMAZIP_Change()
    cmdSave.Enabled = True
End Sub

Private Sub txtMAZIP_GotFocus()
    goUtil.utSelText txtMAZIP
    Set CurrentTextBox = Nothing
End Sub

Private Sub txtMAZIP_LostFocus()
    goUtil.utValidate , txtMAZIP
End Sub

Private Sub txtMAZIP4_Change()
    cmdSave.Enabled = True
End Sub

Private Sub txtMAZIP4_GotFocus()
    goUtil.utSelText txtMAZIP4
    Set CurrentTextBox = Nothing
End Sub

Private Sub txtMAZIP4_LostFocus()
    goUtil.utValidate , txtMAZIP4
End Sub

Private Sub txtMortgageeName_Change()
    cmdSave.Enabled = True
End Sub

Private Sub txtMortgageeName_GotFocus()
    goUtil.utSelText txtMortgageeName
        Set CurrentTextBox = txtMortgageeName
End Sub

Private Sub txtMortgageeName_LostFocus()
    goUtil.utValidate , txtMortgageeName
End Sub

Private Sub txtPACity_Change()
    cmdSave.Enabled = True
End Sub

Private Sub txtPACity_GotFocus()
    goUtil.utSelText txtPACity
        Set CurrentTextBox = txtPACity
End Sub

Private Sub txtPACity_LostFocus()
    goUtil.utValidate , txtPACity
End Sub

Private Sub txtPAOtherPostCode_Change()
    cmdSave.Enabled = True
End Sub

Private Sub txtPAOtherPostCode_GotFocus()
    goUtil.utSelText txtPAOtherPostCode
    Set CurrentTextBox = Nothing
End Sub

Private Sub txtPAOtherPostCode_LostFocus()
    goUtil.utValidate , txtPAOtherPostCode
End Sub

Private Sub txtPAStreet_Change()
    cmdSave.Enabled = True
End Sub

Private Sub txtPAStreet_GotFocus()
    goUtil.utSelText txtPAStreet
        Set CurrentTextBox = txtPAStreet
End Sub

Private Sub txtPAStreet_LostFocus()
    goUtil.utValidate , txtPAStreet
End Sub

Private Sub txtPAZIP_Change()
    cmdSave.Enabled = True
End Sub

Private Sub txtPAZIP_GotFocus()
    goUtil.utSelText txtPAZIP
    Set CurrentTextBox = Nothing
End Sub

Private Sub txtPAZIP_LostFocus()
    goUtil.utValidate , txtPAZIP
End Sub

Private Sub txtPAZIP4_Change()
    cmdSave.Enabled = True
End Sub

Private Sub txtPAZIP4_GotFocus()
    goUtil.utSelText txtPAZIP4
    Set CurrentTextBox = Nothing
End Sub

Private Sub txtPAZIP4_LostFocus()
    goUtil.utValidate , txtPAZIP4
End Sub

Private Sub txtPLAdminComments_GotFocus()
    goUtil.utSelText txtPLAdminComments
    Set CurrentTextBox = txtPLAdminComments
End Sub

Private Sub txtPLAdminComments_LostFocus()
   goUtil.utValidate , txtPLAdminComments
End Sub


Private Sub txtPLLimitAmount_GotFocus()
    goUtil.utSelText txtPLLimitAmount
    Set CurrentTextBox = Nothing
End Sub

Private Sub txtPLLimitAmount_LostFocus()
    goUtil.utValidate , txtPLLimitAmount
End Sub

Private Sub txtPLRCSaidProp_GotFocus()
    goUtil.utSelText txtPLRCSaidProp
    Set CurrentTextBox = Nothing
End Sub

Private Sub txtPLRCSaidProp_LostFocus()
    goUtil.utValidate , txtPLRCSaidProp
End Sub

Private Sub txtPLReserves_GotFocus()
    goUtil.utSelText txtPLReserves
    Set CurrentTextBox = Nothing
End Sub

Private Sub txtPLReserves_LostFocus()
    goUtil.utValidate , txtPLReserves
End Sub

Private Sub txtPolicyDescription_Change()
    cmdSave.Enabled = True
End Sub

Private Sub txtPolicyDescription_GotFocus()
    goUtil.utSelText txtPolicyDescription
    Set CurrentTextBox = txtPolicyDescription
End Sub

Private Sub txtPolicyDescription_LostFocus()
    goUtil.utValidate , txtPolicyDescription
End Sub

Private Sub txtPolicyNo_Change()
    cmdSave.Enabled = True
End Sub

Private Sub txtPolicyNo_GotFocus()
    goUtil.utSelText txtPolicyNo
    Set CurrentTextBox = Nothing
End Sub

Private Sub txtPolicyNo_LostFocus()
    goUtil.utValidate , txtPolicyNo
End Sub

Private Sub txtReceivedDate_Change()
    cmdSave.Enabled = True
End Sub

Private Sub txtReceivedDate_GotFocus()
    goUtil.utSelText txtReceivedDate
    Set CurrentTextBox = Nothing
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

Private Sub Form_Load()
    On Error GoTo EH
    Dim bDelOrigPDF As Boolean
    
    mbLoading = True
    
    'Set Check box for Deleting original PDF File after Attach
    bDelOrigPDF = CBool(GetSetting(App.EXEName, "GENERAL", "DELETE_ORIG_PDF", False))
    If bDelOrigPDF Then
        chkDelOrigPDF.Value = vbChecked
    Else
        chkDelOrigPDF.Value = vbUnchecked
    End If
    
    'Set the Form Icon
    Me.Icon = mfrmClaim.optClaim(GuiClaimOptions.opt01_ClaimInfo).Picture
    
    LoadMe
    
    CheckStatus
    
    ShowFrame
    EnableCC False
    cmdSave.Enabled = False
    mbLoading = False
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_Load"
End Sub

Private Sub EnableCC(pbEnable As Boolean)
    On Error GoTo EH
    'Check for Client Co Specific Enableing
    Dim oControl As Control
    Dim sTag As String
    Dim sCC As String
    Dim bDoEnableCC As Boolean
    Dim sLossFormat As String
    
    sCC = goUtil.goCurCarList.ClassName
    'Set the CC Flag depending on what Client Co is Disabling controls
    If StrComp(sCC, "V2ECcarFarmers.clsLists", vbTextCompare) = 0 Then
        'For Famers Only Disable if Loss Format is XML
        sLossFormat = goUtil.IsNullIsVbNullString(mfrmClaim.adoRSAssignments.Fields("LRFormat"))
        If StrComp(sLossFormat, "V2ECcarFarmers.clsLossXML01", vbTextCompare) = 0 Then
            bDoEnableCC = True
        End If
    End If
    
    If bDoEnableCC Then
        For Each oControl In Me.Controls
            If Not TypeOf oControl Is ImageList Then
                sTag = oControl.Tag
                If pbEnable Then
                    If InStr(1, sTag, "CCEnable", vbTextCompare) > 0 Then
                        oControl.Enabled = True
                    End If
                Else
                    If InStr(1, sTag, "CCDisable", vbTextCompare) > 0 Then
                        oControl.Enabled = False
                    End If
                End If
            End If
        Next
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub EnableCC"
End Sub

Public Function CheckStatus() As Boolean
    On Error GoTo EH
    Dim lTabsPos As Long
    Dim oFrame As Control
    Dim MyFrame As Frame
    Dim oControl As Control
    Dim MyTextBox As TextBox
    Dim MycmdButton As CommandButton
    Dim sFrameName As String
    
     'If this claim is closed only certain things can be edited
    If mfrmClaim.MyStatus = iAssignmentsStatus_CLOSED Then
        For lTabsPos = 1 To TSClaimInfo.Tabs.Count
            Select Case UCase(TSClaimInfo.Tabs(lTabsPos).Tag)
                Case UCase(framSpecifics.Name), _
                        UCase(framInsuredInfo.Name), _
                        UCase(framPolicyLimits.Name)
                    sFrameName = TSClaimInfo.Tabs(lTabsPos).Tag
                    For Each oFrame In Me.Controls
                        If TypeOf oFrame Is Frame Then
                            Set MyFrame = oFrame
                            If StrComp(MyFrame.Name, sFrameName, vbTextCompare) = 0 Then
                                MyFrame.Enabled = False
                                Exit For
                            End If
                        End If
                    Next
                Case UCase(framSpecifics.Name)
                    'Need to disable all dates except the closedate
                    For Each oControl In Me.Controls
                        If TypeOf oControl Is TextBox Then
                            Set MyTextBox = oControl
                            If StrComp(MyTextBox.Tag, "Date", vbTextCompare) = 0 Then
                                If StrComp(MyTextBox.Name, txtCloseDate.Name, vbTextCompare) <> 0 Then
                                    MyTextBox.Enabled = False
                                End If
                            End If
                        ElseIf TypeOf oControl Is CommandButton Then
                            Set MycmdButton = oControl
                            If StrComp(MycmdButton.Tag, "Date", vbTextCompare) = 0 Then
                                If StrComp(MycmdButton.Name, cmdCloseDate.Name, vbTextCompare) <> 0 Then
                                    MycmdButton.Enabled = False
                                End If
                            End If
                        End If
                    Next
                Case UCase(framLossReport.Name)
                    'Need to disable all control except the closedate
                    For Each oControl In Me.Controls
                        If TypeOf oControl Is CommandButton Then
                            Set MycmdButton = oControl
                            If StrComp(MycmdButton.Name, cmdViewPDFLossReport.Name, vbTextCompare) = 0 Then
                                MycmdButton.Enabled = True
                            End If
                        Else
                            If (Not TypeOf oControl Is ImageList) And (Not TypeOf oControl Is TabStrip) Then
                                If oControl.Container.Name = framLossReport.Name Then
                                    oControl.Enabled = False
                                End If
                            End If
                        End If
                    Next
            End Select
        Next
    End If
    
    
    CheckStatus = True
    
    'cleanup
    Set oFrame = Nothing
    Set MyFrame = Nothing
    Set oControl = Nothing
    Set MyTextBox = Nothing
    Set MycmdButton = Nothing
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function CheckStatus"
End Function

Public Function LoadLossReportStuff() As Boolean
    On Error GoTo EH
    
    PopulateAssignmentLossReportFormatLookUp

    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function LoadLossReportStuff"
End Function

Public Function LoadPolicyLimitsStuff() As Boolean
    On Error GoTo EH
    
    'init these values
    lblPLCaption.Caption = vbNullString
    txtPLLimitAmount.Text = vbNullString
    txtPLRCSaidProp.Text = vbNullString
    txtPLReserves.Text = vbNullString
    chkPLIsDeleted.Value = vbUnchecked
    txtPLAdminComments.Text = vbNullString
    Set mitmXPLSelected = Nothing
    
    mfrmClaim.SetadoRSPolicyLimits msAssignmentsID
    PopulatelvwPLClassTypeID
    PopulateAddPLAppClassTypeIDLookUp
    
    LoadPolicyLimitsStuff = True
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function LoadPolicyLimitsStuff"
End Function

Public Function LoadMe() As Boolean
    On Error GoTo EH
    Dim oConn As New ADODB.Connection
    Dim adoRS As ADODB.Recordset
    Dim RS As ADODB.Recordset
    Dim sSQL As String
    Dim sMess As String
    Dim sTemp As String
    
    mbLoadingMe = True
    
    Set oConn = New ADODB.Connection
    Set adoRS = New ADODB.Recordset
    
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    sSQL = "SELECT A.*, "
    sSQL = sSQL & "( "
    sSQL = sSQL & "SELECT UserName "
    sSQL = sSQL & "FROM USERS "
    sSQL = sSQL & "WHERE UsersID =  " & goUtil.gsCurUsersID & " "
    sSQL = sSQL & ") As AdjusterSpecUserName, "
    sSQL = sSQL & "( "
    sSQL = sSQL & "SELECT  ACID "
    sSQL = sSQL & "FROM ClientCoAdjusterSpec "
    sSQL = sSQL & "Where ClientCoAdjusterSpecID = A.[AdjusterSPecID] "
    sSQL = sSQL & ") As AdjusterSpecACID, "
    sSQL = sSQL & "( "
    sSQL = sSQL & "SELECT  ACIDDescription "
    sSQL = sSQL & "FROM ClientCoAdjusterSpec "
    sSQL = sSQL & "Where ClientCoAdjusterSpecID = A.[AdjusterSPecID] "
    sSQL = sSQL & ") As AdjusterSpecAcidDescription, "
    sSQL = sSQL & "( "
    sSQL = sSQL & "SELECT  Type "
    sSQL = sSQL & "FROM    AssignmentType "
    sSQL = sSQL & "WHERE   AssignmentTypeID = A.[AssignmentTypeID] "
    sSQL = sSQL & ") As AssignmentTypeType, "
    sSQL = sSQL & "( "
    sSQL = sSQL & "SELECT  CatCode "
    sSQL = sSQL & "FROM    ClientCompanyCatSpec "
    sSQL = sSQL & "WHERE   ClientCompanyCatSpecID = A.[ClientCompanyCatSpecID] "
    sSQL = sSQL & ") As ClientCompanyCatSpecCatCode, "
    sSQL = sSQL & "( "
    sSQL = sSQL & "SELECT Name "
    sSQL = sSQL & "FROM CAT "
    sSQL = sSQL & "WHERE CATID = " & goUtil.gsCurCat & " "
    sSQL = sSQL & ") As CatName, "
    sSQL = sSQL & "S.Status As Status, "
    sSQL = sSQL & "S.StatusAlias As StatusAlias, "
    sSQL = sSQL & "CCCS.CatCode "
    sSQL = sSQL & "FROM (Assignments A "
    sSQL = sSQL & "INNER JOIN STATUS S ON A.StatusID = S.StatusID) "
    sSQL = sSQL & "INNER JOIN CLIENTCOMPANYCATSPEC CCCS ON (A.ClientCompanyCatSpecID = CCCS.ClientCompanyCatSpecID) "
    sSQL = sSQL & "WHERE A.ClientCompanyCatSpecID IN "
    sSQL = sSQL & "( "
    sSQL = sSQL & "SELECT  ClientCompanyCatSpecID "
    sSQL = sSQL & "FROM ClientCompanyCatSpec "
    sSQL = sSQL & "WHERE ClientCompanyID = " & goUtil.gsCurCar & " "
    sSQL = sSQL & "AND     CATID = " & goUtil.gsCurCat & " "
    sSQL = sSQL & ") "
    sSQL = sSQL & "AND A.AdjusterSpecID IN "
    sSQL = sSQL & "( "
    sSQL = sSQL & "SELECT  ClientCoAdjusterSpecID "
    sSQL = sSQL & "FROM ClientCoAdjusterSpec "
    sSQL = sSQL & "Where ClientCompanyID = " & goUtil.gsCurCar & " "
    sSQL = sSQL & "AND UsersID = " & goUtil.gsCurUsersID & " "
    sSQL = sSQL & ") "
    sSQL = sSQL & "AND A.ID = " & msAssignmentsID & " "
    
    'Use Disconnected Record Set on asUseClient Cusor ONLY !
    adoRS.CursorLocation = adUseClient
    adoRS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set adoRS.ActiveConnection = Nothing
    
    If adoRS.RecordCount > 1 Then
        adoRS.MoveFirst
        sMess = "Database Error.  Duplicate Record ID found!" & vbCrLf & vbCrLf & "ID = " & adoRS!ID & vbCrLf & vbCrLf & "AssignmentsID = " & adoRS!AssignmentsID
        Err.Raise -999, , sMess
    ElseIf adoRS.RecordCount = 0 Then
        sMess = "Database Error.  Record ID Not found!" & vbCrLf & vbCrLf & "AssignmentsID = " & msAssignmentsID
        Err.Raise -999, , sMess
    End If
    
    adoRS.MoveFirst
    
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
    
    'Populate the Available Assignment type
    If Not MyGUI.adoRSAssignmentType Is Nothing Then
        cboAssignmentType.Clear
        mfrmClaim.PopulateLookUp MyGUI.adoRSAssignmentType, _
                        adoRS, _
                        cboAssignmentType, _
                        "AssignmentTypeID", _
                        "AssignmentTypeID", _
                        "Type", _
                        "Description"
    End If
    
    'Populate the Available ACID (Adjuster Client Identification)
    If Not MyGUI.adoRSACID Is Nothing Then
        cboACID.Clear
        mfrmClaim.PopulateLookUp MyGUI.adoRSACID, _
                        adoRS, _
                        cboACID, _
                        "ClientCoAdjusterSpecID", _
                        "AdjusterSpecID", _
                        "ACID", _
                        "ACIDDescription"
    End If
    
    'Populate the Available ACID Display (Adjuster Client Identification)
    'This will show on billing information
    If Not MyGUI.adoRSACID Is Nothing Then
        cboACIDDisplay.Clear
        mfrmClaim.PopulateLookUp MyGUI.adoRSACID, _
                        adoRS, _
                        cboACIDDisplay, _
                        "ClientCoAdjusterSpecID", _
                        "AdjusterSpecIDDisplay", _
                        "ACID", _
                        "ACIDDescription"
    End If
    
    'Populate the Available Cat Code
    If Not MyGUI.adoRSCatCode Is Nothing Then
        cboCatCode.Clear
        mfrmClaim.PopulateLookUp MyGUI.adoRSCatCode, _
                        adoRS, _
                        cboCatCode, _
                        "ClientCompanyCatSpecID", _
                        "ClientCompanyCatSpecID", _
                        "CatCode", _
                        "Comments"
    End If
    
    'Specifics
    txtIBNUM.Text = goUtil.IsNullIsVbNullString(adoRS.Fields("IBNUM"))
    txtCLIENTNUM.Text = goUtil.IsNullIsVbNullString(adoRS.Fields("CLIENTNUM"))
    txtPolicyNo.Text = goUtil.IsNullIsVbNullString(adoRS.Fields("PolicyNo"))
    txtPolicyDescription.Text = goUtil.IsNullIsVbNullString(adoRS.Fields("PolicyDescription"))
    txtMortgageeName.Text = goUtil.IsNullIsVbNullString(adoRS.Fields("MortgageeName"))
    txtAgentNo.Text = goUtil.IsNullIsVbNullString(adoRS.Fields("AgentNo"))
    txtReportedBy.Text = goUtil.IsNullIsVbNullString(adoRS.Fields("ReportedBy"))
    txtReportedByPhone.Text = goUtil.IsNullIsVbNullString(adoRS.Fields("ReportedByPhone"))
    
    'Dates
    txtLossDate.Text = goUtil.IsNullIsVbNullString(adoRS.Fields("LossDate"))
    txtLossDate.Text = Format(txtLossDate.Text, "MM/DD/YYYY")
    txtAssignedDate.Text = goUtil.IsNullIsVbNullString(adoRS.Fields("AssignedDate"))
    txtAssignedDate.Text = Format(txtAssignedDate.Text, "MM/DD/YYYY")
    txtReceivedDate.Text = goUtil.IsNullIsVbNullString(adoRS.Fields("ReceivedDate"))
    txtReceivedDate.Text = Format(txtReceivedDate.Text, "MM/DD/YYYY")
    txtContactDate.Text = goUtil.IsNullIsVbNullString(adoRS.Fields("ContactDate"))
    txtContactDate.Text = Format(txtContactDate.Text, "MM/DD/YYYY")
    txtInspectedDate.Text = goUtil.IsNullIsVbNullString(adoRS.Fields("InspectedDate"))
    txtInspectedDate.Text = Format(txtInspectedDate.Text, "MM/DD/YYYY")
    txtCloseDate.Text = goUtil.IsNullIsVbNullString(adoRS.Fields("CloseDate"))
    txtCloseDate.Text = Format(txtCloseDate.Text, "MM/DD/YYYY")
    
    'Insured Info
    txtInsured.Text = goUtil.IsNullIsVbNullString(adoRS.Fields("Insured"))
    txtHomePhone.Text = goUtil.IsNullIsVbNullString(adoRS.Fields("HomePhone"))
    txtBusinessPhone.Text = goUtil.IsNullIsVbNullString(adoRS.Fields("BusinessPhone"))
    'Property Address
    txtPAStreet.Text = goUtil.IsNullIsVbNullString(adoRS.Fields("PAStreet"))
    txtPACity.Text = goUtil.IsNullIsVbNullString(adoRS.Fields("PACity"))
    'Populate the Available Sates for Property address
    If Not MyGUI.adoRSState Is Nothing Then
        cboPAState.Clear
        mfrmClaim.PopulateLookUp MyGUI.adoRSState, _
                        adoRS, _
                        cboPAState, _
                        "StateID", _
                        "", _
                        "Code", _
                        "Name", _
                        True, _
                        "PAState"
    End If
    txtPAZIP.Text = Format(goUtil.IsNullIsVbNullString(adoRS.Fields("PAZIP")), "00000")
    txtPAZIP4.Text = Format(goUtil.IsNullIsVbNullString(adoRS.Fields("PAZIP4")), "0000")
    txtPAOtherPostCode.Text = goUtil.IsNullIsVbNullString(adoRS.Fields("PAOtherPostCode"))
    
    'Mailing Address
    txtMAStreet.Text = goUtil.IsNullIsVbNullString(adoRS.Fields("MAStreet"))
    txtMACity.Text = goUtil.IsNullIsVbNullString(adoRS.Fields("MACity"))
    'Populate the Available Sates for Mailling address
    If Not MyGUI.adoRSState Is Nothing Then
        cboMAState.Clear
        mfrmClaim.PopulateLookUp MyGUI.adoRSState, _
                        adoRS, _
                        cboMAState, _
                        "StateID", _
                        "", _
                        "Code", _
                        "Name", _
                        True, _
                        "MAState"
    End If
    txtMAZIP.Text = Format(goUtil.IsNullIsVbNullString(adoRS.Fields("MAZIP")), "00000")
    txtMAZIP4.Text = Format(goUtil.IsNullIsVbNullString(adoRS.Fields("MAZIP4")), "0000")
    txtMAOtherPostCode.Text = goUtil.IsNullIsVbNullString(adoRS.Fields("MAOtherPostCode"))
    
    'Policy limits
    txtDeductible.Text = goUtil.IsNullIsVbNullString(adoRS.Fields("Deductible"))
    If mbLoading Then
        LoadHeaderPLClassTypeID
    End If
    LoadPolicyLimitsStuff
    
    'Loss Report
    LoadLossReportStuff
    sTemp = goUtil.IsNullIsVbNullString(adoRS.Fields("LRFormat"))
    If sTemp <> vbNullString Then
        If StrComp(sTemp, "TEXT", vbTextCompare) = 0 Then
            txtLossReport.Text = goUtil.IsNullIsVbNullString(adoRS.Fields("LossReport"))
        Else
            txtLossReport.Text = vbNullString
        End If
    Else
        txtLossReport.Text = vbNullString
    End If
    
    'Admin Comments
    txtAdminComments.Text = goUtil.IsNullIsVbNullString(adoRS.Fields("AdminComments"))
    
    'cleanup
    Set RS = Nothing
    Set adoRS = Nothing
    Set oConn = Nothing
    
    LoadMe = True
    mbLoadingMe = False
    Exit Function
EH:
    mbLoadingMe = False
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function LoadMe"
    Set adoRS = Nothing
    Set oConn = Nothing
End Function



Private Function PopulateAssignmentLossReportFormatLookUp() As Boolean
    On Error GoTo EH
    Dim MyadoRSAssignments As ADODB.Recordset
    Dim cboBox As ComboBox
    Dim sLRFormat As String

    Dim lSelIndex As Long

    Set MyadoRSAssignments = mfrmClaim.adoRSAssignments
    Set cboBox = cboAssignmentLossReportFormat

    cboBox.Clear
    cboBox.AddItem "(--CHANGE FORMAT--)"
    'Set the current Format for the currently selected assignment (claim)
    
    sLRFormat = goUtil.IsNullIsVbNullString(MyadoRSAssignments.Fields("LRFormat"))
    
    If InStr(1, sLRFormat, "TEXT", vbTextCompare) > 0 Then
        lblSelFormat.Caption = "Current Format - TEXT"
    ElseIf InStr(1, sLRFormat, "OLEType_pdf", vbTextCompare) > 0 Then
        lblSelFormat.Caption = "Current Format - PDF (Adobe)"
    ElseIf Trim(sLRFormat) <> vbNullString Then
        lblSelFormat.Caption = "Current Format - " & Trim(sLRFormat)
    End If

    cboBox.AddItem "TEXT"
    cboBox.AddItem "PDF (Adobe)"
    cboBox.ListIndex = 0

    PopulateAssignmentLossReportFormatLookUp = True

    Set cboBox = Nothing
    Set MyadoRSAssignments = Nothing
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Function PopulateAssignmentLossReportFormatLookUp"
End Function


Private Function PopulateAddPLAppClassTypeIDLookUp() As Boolean
    On Error GoTo EH
    Dim MyadoRSPolicyLimits As ADODB.Recordset
    Dim MyadoRSClassType As ADODB.Recordset
    Dim cboBox As ComboBox
    Dim sTemp As String
    Dim lId As Long
    Dim sID As String
    Dim sSeekID As String
    
    Set MyadoRSPolicyLimits = mfrmClaim.adoRSPolicyLimits
    Set MyadoRSClassType = MyGUI.adoClassType
    Set cboBox = cboAddPLAppClassTypeID
    
    cboBox.Clear
    If MyadoRSClassType.RecordCount > 0 Then
        MyadoRSClassType.MoveFirst
        Do Until MyadoRSClassType.EOF
            'If the class Type ID is Already part of adoRSPolicyLimits RecordSet
            'then do not add it to the list of available Class Types to Add.
            lId = MyadoRSClassType.Fields("ClassTypeID").Value
            If Not MyadoRSPolicyLimits.RecordCount = 0 Then
                MyadoRSPolicyLimits.MoveFirst
            End If
            Do Until MyadoRSPolicyLimits.EOF
                sID = lId
                sSeekID = goUtil.IsNullIsVbNullString(MyadoRSPolicyLimits.Fields("ClassTypeID"))
                If StrComp(sID, sSeekID, vbTextCompare) = 0 Then
                    Exit Do
                End If
                MyadoRSPolicyLimits.MoveNext
            Loop
            'the adoRSPolicyLimits should be at EOF if the Seek did not find it
            If MyadoRSPolicyLimits.EOF Then
                sTemp = MyadoRSClassType.Fields("Class").Value
                sTemp = sTemp & " ("
                sTemp = sTemp & MyadoRSClassType.Fields("Description").Value
                sTemp = sTemp & ")"
                cboBox.AddItem sTemp
                'Set the Record Id to the Itemdata of the element just added
                cboBox.ItemData(cboBox.NewIndex) = lId
            Else
                If Not MyadoRSPolicyLimits.RecordCount = 0 Then
                    MyadoRSPolicyLimits.MoveFirst
                End If
            End If
            MyadoRSClassType.MoveNext
        Loop
        
        'Select the first item on the list
        If cboBox.ListCount > 0 Then
            cboBox.ListIndex = 0
        End If
        
    End If
    
    PopulateAddPLAppClassTypeIDLookUp = True
    
    Set MyadoRSPolicyLimits = Nothing
    Set MyadoRSClassType = Nothing
    Set cboBox = Nothing
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Function PopulateAddPLAppClassTypeIDLookUp"
End Function


Public Function SaveMe() As Boolean
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim MyadoRSAssignments As ADODB.Recordset
    Dim iCurrentStatus As V2ECKeyBoard.AssgnStatus
    Dim sSQL As String
    'If Close date is Set then be sure all the other dates are set tooooo
    Dim bCloseDateIsSet As Boolean
    'Assignments Vars
    Dim sAssignmentsID As String
    Dim sID As String
    Dim sAssignmentTypeID As String
    Dim sClientCompanyCatSpecID As String
    Dim sAdjusterSpecID  As String
    Dim sAdjusterSpecIDDisplay As String
    Dim sSPVersion As String
    Dim sIBNUM As String
    Dim sCLIENTNUM As String
    Dim sPolicyNo As String
    Dim sPolicyDescription As String
    Dim sInsured As String
    Dim sMailingAddress As String
    Dim sMAStreet As String
    Dim sMACity As String
    Dim sMAState As String
    Dim sMAZIP As String
    Dim sMAZIP4 As String
    Dim sMAOtherPostCode As String
    Dim sHomePhone As String
    Dim sBusinessPhone As String
    Dim sPropertyAddress As String
    Dim sPAStreet As String
    Dim sPACity As String
    Dim sPAState As String
    Dim sPAZIP As String
    Dim sPAZIP4 As String
    Dim sPAOtherPostCode As String
    Dim sMortgageeName As String
    Dim sAgentNo As String
    Dim sReportedBy As String
    Dim sReportedByPhone As String
    Dim sDeductible As String
    Dim sAppDedClassTypeIDOrder As String
    Dim sLRFormat As String
    Dim sLossReport As String
    Dim sDownLoadLossReport As String
    Dim sUpLoadLossReport As String
    Dim sStatusID As String
    Dim sTypeOfLossID As String
    Dim sXactTypeOfLoss As String
    Dim sSentToXact  As String
    Dim sLossDate  As String
    Dim sAssignedDate As String
    Dim sReceivedDate As String
    Dim sContactDate As String
    Dim sInspectedDate As String
    Dim sCloseDate As String
    Dim sReassigned As String
    Dim sDateReassigned As String
    Dim sRAAdjusterSpecID As String
    Dim sIsLocked As String
    Dim sIsDeleted As String
    Dim sDownLoadMe As String
    Dim sUpLoadMe  As String
    Dim sDownLoadAll As String
    Dim sUpLoadAll  As String
    Dim sAdminComments As String
    Dim sMiscDelimSettings As String
    Dim sDateLastUpdated  As String
    Dim sUID As String
    Dim sMess As String
    
    'validate all the fields on this form
    goUtil.utValidate Me
    
    'Check for Drop Down Items Not Selected that should be
    If cboAssignmentType.ListIndex = -1 Then
        sMess = sMess & "Assignment Type not selected !" & vbCrLf
    End If
    If cboCatCode.ListIndex = -1 Then
        sMess = sMess & "Cat Code not selected !" & vbCrLf
    End If
    If cboACID.ListIndex = -1 Then
        sMess = sMess & "ACID not selected !" & vbCrLf
    End If
    If cboACIDDisplay.ListIndex = -1 Then
        sMess = sMess & "ACID Display not selected !" & vbCrLf
    End If
    If cboMAState.ListIndex = -1 Then
        sMess = sMess & "Mailing State not selected !" & vbCrLf
    End If
    If cboPAState.ListIndex = -1 Then
        sMess = sMess & "Property State not selected !" & vbCrLf
    End If
    If cboTypeOfLoss.ListIndex = -1 Then
        sMess = sMess & "Type Of Loss not selected !" & vbCrLf
    End If
    
    'DATES !!!!
    'Close Date
    If IsDate(txtCloseDate.Text) Then
        'Check for Close date but no other dates filled out
        bCloseDateIsSet = True
        sCloseDate = "#" & Format(txtCloseDate.Text, "MM/DD/YYYY") & "#"
    Else
        sCloseDate = "null"
    End If
    
    'Loss Date
    If IsDate(txtLossDate.Text) Then
        sLossDate = "#" & Format(txtLossDate.Text, "MM/DD/YYYY") & "#"
    Else
        'Check for Close date but no other dates filled out
        If bCloseDateIsSet Then
            sMess = sMess & "Loss Date is not set!" & vbCrLf
        End If
        sLossDate = "null"
    End If
    'Assigned Date
    If IsDate(txtAssignedDate.Text) Then
        sAssignedDate = "#" & Format(txtAssignedDate.Text, "MM/DD/YYYY") & "#"
    Else
        If bCloseDateIsSet Then
            sMess = sMess & "Assigned Date is not set!" & vbCrLf
        End If
        sAssignedDate = "null"
    End If
    'Received Date
    If IsDate(txtReceivedDate.Text) Then
        sReceivedDate = "#" & Format(txtReceivedDate.Text, "MM/DD/YYYY") & "#"
    Else
        If bCloseDateIsSet Then
            sMess = sMess & "Received Date is not set!" & vbCrLf
        End If
        sReceivedDate = "null"
    End If
    'Contact Date
    If IsDate(txtContactDate.Text) Then
        sContactDate = "#" & Format(txtContactDate.Text, "MM/DD/YYYY") & "#"
    Else
        If bCloseDateIsSet Then
            sMess = sMess & "Contact Date is not set!" & vbCrLf
        End If
        sContactDate = "null"
    End If
    'Inspected Date
    If IsDate(txtInspectedDate.Text) Then
        sInspectedDate = "#" & Format(txtInspectedDate.Text, "MM/DD/YYYY") & "#"
    Else
        If bCloseDateIsSet Then
            sMess = sMess & "Inspected Date is not set!" & vbCrLf
        End If
        sInspectedDate = "null"
    End If
    
    If sMess <> vbNullString Then
        sMess = "Could not save " & Me.Caption & vbCrLf & vbCrLf & sMess
        MsgBox sMess, vbExclamation + vbOKOnly, "Could Not Save Claim Information."
        Exit Function
    End If
    
    'Use this to check new values to be inserted
    'against the current values in this recordset
    Set MyadoRSAssignments = mfrmClaim.adoRSAssignments
    
    'set the Assignemtn vars
    sAssignmentsID = msAssignmentsID
    
    sID = msAssignmentsID
    
    sAssignmentTypeID = cboAssignmentType.ItemData(cboAssignmentType.ListIndex)
    
    sClientCompanyCatSpecID = cboCatCode.ItemData(cboCatCode.ListIndex)
    
    sAdjusterSpecID = cboACID.ItemData(cboACID.ListIndex)
    
    sAdjusterSpecIDDisplay = cboACIDDisplay.ItemData(cboACIDDisplay.ListIndex)
    
    sSPVersion = "[SPVersion]"
    'IBNUM
    sIBNUM = "'" & goUtil.utCleanSQLString(UCase(txtIBNUM.Text)) & "'"
    'CLIENTNUM
    sCLIENTNUM = "'" & goUtil.utCleanSQLString(UCase(txtCLIENTNUM.Text)) & "'"
    'Policy Number
    sPolicyNo = "'" & goUtil.utCleanSQLString(UCase(txtPolicyNo.Text)) & "'"
    'Policty Description
    sPolicyDescription = "'" & goUtil.utCleanSQLString(UCase(txtPolicyDescription.Text)) & "'"
    'Insured
    sInsured = "'" & goUtil.utCleanSQLString(UCase(txtInsured.Text)) & "'"
    
    'Mailing Address
    'Street
    sMAStreet = UCase(txtMAStreet.Text)
    'City
    sMACity = UCase(txtMACity.Text)
    'State
    sMAState = left(UCase(cboMAState.Text), 2)
    'Zip
    sMAZIP = txtMAZIP.Text
    'Zip4
    sMAZIP4 = txtMAZIP4.Text
    'Other Post Code
    sMAOtherPostCode = UCase(txtMAOtherPostCode.Text)
    'Build entire Address
    If sMAZIP & "-" & sMAZIP4 = "00000-0000" Then
        sMailingAddress = "'" & goUtil.utCleanSQLString(goUtil.utJoinToAddress(sMAStreet, sMACity, sMAState, sMAOtherPostCode)) & "'"
    Else
        sMailingAddress = "'" & goUtil.utCleanSQLString(goUtil.utJoinToAddress(sMAStreet, sMACity, sMAState, Format(sMAZIP, "00000") & "-" & Format(sMAZIP4, "0000"))) & "'"
    End If
    'Street
    sMAStreet = "'" & goUtil.utCleanSQLString(UCase(txtMAStreet.Text)) & "'"
    'City
    sMACity = "'" & goUtil.utCleanSQLString(UCase(txtMACity.Text)) & "'"
    'State
    sMAState = "'" & goUtil.utCleanSQLString(left(UCase(cboMAState.Text), 2)) & "'"
    'Zip
    sMAZIP = txtMAZIP.Text
    'Zip4
    sMAZIP4 = txtMAZIP4.Text
    'Other Post Code
    sMAOtherPostCode = "'" & goUtil.utCleanSQLString(UCase(txtMAOtherPostCode.Text)) & "'"
    'End Mailing Address
     
    'Property Address
    'Street
    sPAStreet = UCase(txtPAStreet.Text)
    'City
    sPACity = UCase(txtPACity.Text)
    'State
    sPAState = left(UCase(cboPAState.Text), 2)
    'Zip
    sPAZIP = txtPAZIP.Text
    'Zip4
    sPAZIP4 = txtPAZIP4.Text
    'other PostCode
    sPAOtherPostCode = UCase(txtPAOtherPostCode.Text)
    'Build entire Address
    If sPAZIP & "-" & sPAZIP4 = "00000-0000" Then
        sPropertyAddress = "'" & goUtil.utCleanSQLString(goUtil.utJoinToAddress(sPAStreet, sPACity, sPAState, sPAOtherPostCode)) & "'"
    Else
        sPropertyAddress = "'" & goUtil.utCleanSQLString(goUtil.utJoinToAddress(sPAStreet, sPACity, sPAState, Format(sPAZIP, "00000") & "-" & Format(sPAZIP4, "0000"))) & "'"
    End If
    'Street
    sPAStreet = "'" & goUtil.utCleanSQLString(UCase(txtPAStreet.Text)) & "'"
    'City
    sPACity = "'" & goUtil.utCleanSQLString(UCase(txtPACity.Text)) & "'"
    'State
    sPAState = "'" & goUtil.utCleanSQLString(left(UCase(cboPAState.Text), 2)) & "'"
    'Zip
    sPAZIP = txtPAZIP.Text
    'Zip4
    sPAZIP4 = txtPAZIP4.Text
    'Other Post Code
    sPAOtherPostCode = "'" & goUtil.utCleanSQLString(UCase(txtPAOtherPostCode.Text)) & "'"
    'End Property Address
    
    'Home Phone
    sHomePhone = "'" & goUtil.utCleanSQLString(UCase(txtHomePhone.Text)) & "'"
    'Business Phone
    sBusinessPhone = "'" & goUtil.utCleanSQLString(UCase(txtBusinessPhone.Text)) & "'"
    'Mortgage name
    sMortgageeName = "'" & goUtil.utCleanSQLString(UCase(txtMortgageeName.Text)) & "'"
    'Agent No
    sAgentNo = "'" & goUtil.utCleanSQLString(UCase(txtAgentNo.Text)) & "'"
    'Reported By
    sReportedBy = "'" & goUtil.utCleanSQLString(UCase(txtReportedBy.Text)) & "'"
    'Reported by Phone
    sReportedByPhone = "'" & goUtil.utCleanSQLString(UCase(txtReportedByPhone.Text)) & "'"
    'Deductible
    sDeductible = txtDeductible.Text
    
    sAppDedClassTypeIDOrder = "[AppDedClassTypeIDOrder]"
    
    'if the Loss report was changed to TEXT then need to update these vars
    'otherwise they remain the same!
    '(Attaching a PDF Loss Report already updates Assignments table See --> Private Sub cmdAttachPDFLossReport_Click)
    If StrComp(cboAssignmentLossReportFormat.Text, "TEXT", vbTextCompare) = 0 Then
        sLRFormat = "'TEXT'"
        sLossReport = "'" & goUtil.utCleanSQLString(txtLossReport.Text) & "'"
        
        sDownLoadLossReport = "[DownLoadLossReport]"
        
        sUpLoadLossReport = "True"
    Else
        sLRFormat = "[LRFormat]"
        
        sLossReport = "[LossReport]"
        
        sDownLoadLossReport = "[DownLoadLossReport]"
        
        sUpLoadLossReport = "[UpLoadLossReport]"
    End If
    'Type Of Loss
    sTypeOfLossID = cboTypeOfLoss.ItemData(cboTypeOfLoss.ListIndex)
    
    sXactTypeOfLoss = "[XactTypeOfLoss]"
    
    sSentToXact = "[SentToXact]"
    
    sReassigned = "[Reassigned]"
    
    sDateReassigned = "[DateReassigned]"
    
    
    'STATUS ID !
    'Check for Closed Date
    'Change Status ID
    If IsDate(txtCloseDate.Text) Then
        sStatusID = CStr(V2ECKeyBoard.AssgnStatus.iAssignmentsStatus_CLOSED)
    Else
        'Check to see if the Current Status is Closed if it Is Need to
        'Change the Status to NEW
        iCurrentStatus = MyadoRSAssignments.Fields("StatusID").Value
        Select Case iCurrentStatus
            Case V2ECKeyBoard.AssgnStatus.iAssignmentsStatus_CLOSED
                sStatusID = V2ECKeyBoard.AssgnStatus.iAssignmentsStatus_NEW
            Case Else
                sStatusID = "[StatusID]"
        End Select
        
    End If
    
    sRAAdjusterSpecID = "[RAAdjusterSpecID]"
    
    sIsLocked = "[IsLocked]"
    
    sIsDeleted = "[IsDeleted]"
    
    sDownLoadMe = "[DownLoadMe]"
    
    sUpLoadMe = "True"
    
    sDownLoadAll = "[DownLoadAll]"
    
    sUpLoadAll = "[UpLoadAll]"
    'Admin Comments
    sAdminComments = "'" & goUtil.utCleanSQLString(txtAdminComments.Text) & "'"
    
    sMiscDelimSettings = "[MiscDelimSettings]"
    
    sDateLastUpdated = "#" & Now() & "#"
    
    sUID = goUtil.gsCurUsersID
    
    
    sSQL = "Update Assignments Set "
    sSQL = sSQL & "[AssignmentsID] = " & sAssignmentsID & ", "
    sSQL = sSQL & "[ID] = " & sID & ", "
    sSQL = sSQL & "[AssignmentTypeID] = " & sAssignmentTypeID & ", "
    sSQL = sSQL & "[ClientCompanyCatSpecID] = " & sClientCompanyCatSpecID & ", "
    sSQL = sSQL & "[AdjusterSpecID] = " & sAdjusterSpecID & ", "
    sSQL = sSQL & "[AdjusterSpecIDDisplay] =" & sAdjusterSpecIDDisplay & ", "
    sSQL = sSQL & "[SPVersion] = " & sSPVersion & ", "
    sSQL = sSQL & "[IBNUM] = " & sIBNUM & ", "
    sSQL = sSQL & "[CLIENTNUM] = " & sCLIENTNUM & ", "
    sSQL = sSQL & "[PolicyNo] = " & sPolicyNo & ", "
    sSQL = sSQL & "[PolicyDescription] = " & sPolicyDescription & ", "
    sSQL = sSQL & "[Insured] = " & sInsured & ", "
    sSQL = sSQL & "[MailingAddress] = " & sMailingAddress & ", "
    sSQL = sSQL & "[MAStreet] = " & sMAStreet & ", "
    sSQL = sSQL & "[MACity] = " & sMACity & ", "
    sSQL = sSQL & "[MAState] = " & sMAState & ", "
    sSQL = sSQL & "[MAZIP] = " & sMAZIP & ", "
    sSQL = sSQL & "[MAZIP4] = " & sMAZIP4 & ", "
    sSQL = sSQL & "[MAOtherPostCode] = " & sMAOtherPostCode & ", "
    sSQL = sSQL & "[HomePhone]  = " & sHomePhone & ", "
    sSQL = sSQL & "[BusinessPhone] = " & sBusinessPhone & ", "
    sSQL = sSQL & "[PropertyAddress] = " & sPropertyAddress & ", "
    sSQL = sSQL & "[PAStreet]  = " & sPAStreet & ", "
    sSQL = sSQL & "[PACity]  = " & sPACity & ", "
    sSQL = sSQL & "[PAState] = " & sPAState & ", "
    sSQL = sSQL & "[PAZIP]  = " & sPAZIP & ", "
    sSQL = sSQL & "[PAZIP4] = " & sPAZIP4 & ", "
    sSQL = sSQL & "[PAOtherPostCode]  = " & sPAOtherPostCode & ", "
    sSQL = sSQL & "[MortgageeName]  = " & sMortgageeName & ", "
    sSQL = sSQL & "[AgentNo]  = " & sAgentNo & ", "
    sSQL = sSQL & "[ReportedBy] = " & sReportedBy & ", "
    sSQL = sSQL & "[ReportedByPhone] = " & sReportedByPhone & ", "
    sSQL = sSQL & "[Deductible]  = " & sDeductible & ", "
    sSQL = sSQL & "[AppDedClassTypeIDOrder] = " & sAppDedClassTypeIDOrder & ", "
    sSQL = sSQL & "[LRFormat]  = " & sLRFormat & ", "
    sSQL = sSQL & "[LossReport] = " & sLossReport & ", "
    sSQL = sSQL & "[DownLoadLossReport] = " & sDownLoadLossReport & ", "
    sSQL = sSQL & "[UpLoadLossReport] = " & sUpLoadLossReport & ", "
    sSQL = sSQL & "[StatusID]  = " & sStatusID & ", "
    sSQL = sSQL & "[TypeOfLossID] = " & sTypeOfLossID & ", "
    sSQL = sSQL & "[XactTypeOfLoss] = " & sXactTypeOfLoss & ", "
    sSQL = sSQL & "[SentToXact] = " & sSentToXact & ", "
    sSQL = sSQL & "[LossDate] = " & sLossDate & ", "
    sSQL = sSQL & "[AssignedDate] = " & sAssignedDate & ", "
    sSQL = sSQL & "[ReceivedDate] = " & sReceivedDate & ", "
    sSQL = sSQL & "[ContactDate] = " & sContactDate & ", "
    sSQL = sSQL & "[InspectedDate] = " & sInspectedDate & ", "
    sSQL = sSQL & "[CloseDate]  = " & sCloseDate & ", "
    sSQL = sSQL & "[Reassigned]  = " & sReassigned & ", "
    sSQL = sSQL & "[DateReassigned] = " & sDateReassigned & ", "
    sSQL = sSQL & "[RAAdjusterSpecID] = " & sRAAdjusterSpecID & ", "
    sSQL = sSQL & "[IsLocked] = " & sIsLocked & ", "
    sSQL = sSQL & "[IsDeleted] = " & sIsDeleted & ", "
    sSQL = sSQL & "[DownLoadMe] = " & sDownLoadMe & ", "
    sSQL = sSQL & "[UpLoadMe] = " & sUpLoadMe & ", "
    sSQL = sSQL & "[DownLoadAll] = " & sDownLoadAll & ", "
    sSQL = sSQL & "[UpLoadAll] = " & sUpLoadAll & ", "
    sSQL = sSQL & "[AdminComments] = " & sAdminComments & ", "
    sSQL = sSQL & "[MiscDelimSettings] = " & sMiscDelimSettings & ", "
    sSQL = sSQL & "[DateLastUpdated] = " & sDateLastUpdated & ", "
    sSQL = sSQL & "[UpdateByUserID] = " & sUID & " "
    sSQL = sSQL & "WHERE AssignmentsID = " & sAssignmentsID & " "
    
    
    Set oConn = New ADODB.Connection
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    oConn.Execute sSQL
    
    cmdSave.Enabled = False
    SaveMe = True
    
    'cleanup
    Set oConn = Nothing
    Set MyadoRSAssignments = Nothing
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function SaveMe"
End Function

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
    Set mitmXPLSelected = Nothing
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
    
    If Me.WindowState <> vbMinimized Then
        Me.Visible = True
    End If
End Sub

Private Sub TSClaimInfo_Click()
    ShowFrame
End Sub

Public Function ShowFrame() As Boolean
    On Error GoTo EH
    Dim sFrameName As String
    Dim oFrame As Control
    Dim MyFrame As Frame
    Dim oControl As Control
    Dim sTag As String
    
    sFrameName = TSClaimInfo.SelectedItem.Tag
    
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

Private Sub txtReceivedDate_LostFocus()
    goUtil.utValidate , txtReceivedDate
    CheckMyDateWhenCloseDateSet txtReceivedDate
End Sub

Private Sub txtReportedBy_Change()
    cmdSave.Enabled = True
End Sub

Private Sub txtReportedBy_GotFocus()
    goUtil.utSelText txtReportedBy
        Set CurrentTextBox = txtReportedBy
End Sub

Private Sub txtReportedBy_LostFocus()
    goUtil.utValidate , txtReportedBy
End Sub

Private Sub txtReportedByPhone_Change()
    cmdSave.Enabled = True
End Sub

Private Sub txtReportedByPhone_GotFocus()
    goUtil.utSelText txtReportedByPhone
        Set CurrentTextBox = txtReportedByPhone
End Sub

Private Sub PopulatelvwPLClassTypeID()
    On Error GoTo EH
    Dim itmX As ListItem
    Dim iMyIcon As Long
    Dim sFlagText As String
    Dim sTemp As String
    Dim MyadoRSPolicyLimits As ADODB.Recordset
   
    If Not mfrmClaim.adoRSPolicyLimits Is Nothing Then
        Set MyadoRSPolicyLimits = mfrmClaim.adoRSPolicyLimits
    Else
        Exit Sub
    End If
    
    'Clear Any Existing Items
    lvwPLClassTypeID.ListItems.Clear
    
    If Not MyadoRSPolicyLimits.EOF Then
        MyadoRSPolicyLimits.MoveFirst
        Do Until MyadoRSPolicyLimits.EOF
            'Class
            sTemp = goUtil.IsNullIsVbNullString(MyadoRSPolicyLimits.Fields("ID"))
            Set itmX = lvwPLClassTypeID.ListItems.Add(, """" & sTemp & """", goUtil.IsNullIsVbNullString(MyadoRSPolicyLimits.Fields("ClassTypeClass")))
            'ClassTypeID
            itmX.SubItems(GuiPolicyLimits.ClassTypeID - 1) = goUtil.IsNullIsVbNullString(MyadoRSPolicyLimits.Fields("ClassTypeID"))
            'Description
            itmX.SubItems(GuiPolicyLimits.ClassTypeDescription - 1) = goUtil.IsNullIsVbNullString(MyadoRSPolicyLimits.Fields("ClassTypeDescription"))
            'Limit Amount
            itmX.SubItems(GuiPolicyLimits.LimitAmount - 1) = Format(goUtil.IsNullIsVbNullString(MyadoRSPolicyLimits.Fields("LimitAmount")), "#,###,###,##0.00")
            'Sort Limit Amount
            itmX.SubItems(GuiPolicyLimits.LimitAmountSort - 1) = goUtil.utNumInTextSortFormat(goUtil.IsNullIsVbNullString(MyadoRSPolicyLimits.Fields("LimitAmount")))
            'RC Said
            itmX.SubItems(GuiPolicyLimits.RCSaidProp - 1) = Format(goUtil.IsNullIsVbNullString(MyadoRSPolicyLimits.Fields("RCSaidProp")), "#,###,###,##0.00")
            'Sort RC Said
            itmX.SubItems(GuiPolicyLimits.RCSaidPropSort - 1) = goUtil.utNumInTextSortFormat(goUtil.IsNullIsVbNullString(MyadoRSPolicyLimits.Fields("RCSaidProp")))
            'Reserves
            itmX.SubItems(GuiPolicyLimits.Reserves - 1) = Format(goUtil.IsNullIsVbNullString(MyadoRSPolicyLimits.Fields("Reserves")), "#,###,###,##0.00")
            'Sort Reserves
            itmX.SubItems(GuiPolicyLimits.ReservesSort - 1) = goUtil.utNumInTextSortFormat(goUtil.IsNullIsVbNullString(MyadoRSPolicyLimits.Fields("Reserves")))
            'Is Deleted
            If CBool(MyadoRSPolicyLimits.Fields("IsDeleted")) Then
                iMyIcon = GuiPolicyLimitsPic.IsDeleted
            Else
                iMyIcon = Empty
            End If
            sFlagText = goUtil.GetFlagText(MyadoRSPolicyLimits!IsDeleted)
            itmX.SubItems(GuiPolicyLimits.IsDeleted - 1) = sFlagText
            itmX.ListSubItems(GuiPolicyLimits.IsDeleted - 1).ReportIcon = iMyIcon
            'UpLoad Me
            If CBool(MyadoRSPolicyLimits.Fields("UpLoadMe")) Then
                iMyIcon = GuiPolicyLimitsPic.UpLoadMe
            Else
                iMyIcon = Empty
            End If
            sFlagText = goUtil.GetFlagText(MyadoRSPolicyLimits!UpLoadMe)
            itmX.SubItems(GuiPolicyLimits.UpLoadMe - 1) = sFlagText
            itmX.ListSubItems(GuiPolicyLimits.UpLoadMe - 1).ReportIcon = iMyIcon
            'DateLastUpdated
            If Not IsNull(MyadoRSPolicyLimits!DateLastUpdated) Then
                If IsDate(MyadoRSPolicyLimits!DateLastUpdated) Then
                    itmX.SubItems(GuiPolicyLimits.DateLastUpdated - 1) = Format(MyadoRSPolicyLimits!DateLastUpdated, "MM/DD/YYYY HH:MM:SS")
                    itmX.SubItems(GuiPolicyLimits.DateLastUpdatedSort - 1) = Format(MyadoRSPolicyLimits!DateLastUpdated, "YYYY/MM/DD HH:MM:SS")
                Else
                    itmX.SubItems(GuiPolicyLimits.DateLastUpdated - 1) = vbNullString
                    itmX.SubItems(GuiPolicyLimits.DateLastUpdatedSort - 1) = vbNullString
                End If
            Else
                itmX.SubItems(GuiPolicyLimits.DateLastUpdated - 1) = vbNullString
                itmX.SubItems(GuiPolicyLimits.DateLastUpdatedSort - 1) = vbNullString
            End If
            
            'AdminComments
            itmX.SubItems(GuiPolicyLimits.AdminComments - 1) = goUtil.IsNullIsVbNullString(MyadoRSPolicyLimits.Fields("AdminComments"))
            'ID
            itmX.SubItems(GuiPolicyLimits.ID - 1) = goUtil.IsNullIsVbNullString(MyadoRSPolicyLimits.Fields("ID"))
            'IDAssignments
            itmX.Selected = False
            MyadoRSPolicyLimits.MoveNext
        Loop
    End If
    
    'Cleanup
    Set itmX = Nothing
    Set MyadoRSPolicyLimits = Nothing
    
    Exit Sub
EH:
    Set itmX = Nothing
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub PopulatelvwPLClassTypeID"
End Sub

Public Sub LoadHeaderPLClassTypeID()
    On Error GoTo EH
    Dim bGridOn As Boolean
    Dim bHideDeleted As Boolean
    Dim bHideUploadFlags As Boolean
    
    bHideDeleted = CBool(GetSetting(App.EXEName, "GENERAL", "HIDE_DELETED", True))
    bHideUploadFlags = CBool(GetSetting(App.EXEName, "GENERAL", "HIDE_UPLOAD_FLAGS", True))
    
    'set the columnheaders
    With lvwPLClassTypeID
        .ColumnHeaders.Add , "ClassTypeClass", "Class"
        .ColumnHeaders.Add , "ClassTypeID", "ClassTypeID" 'Hidden
        .ColumnHeaders.Add , "ClassTypeDescription", "Description"
        .ColumnHeaders.Add , "LimitAmount", "Limit Amount"
        .ColumnHeaders.Add , "LimitAmountSort", "Sort Limit Amount" 'Hidden
        .ColumnHeaders.Add , "RCSaidProp", "RC Said"
        .ColumnHeaders.Add , "RCSaidPropSort", "Sort RC Said" 'Hidden
        .ColumnHeaders.Add , "Reserves", "Reserves"
        .ColumnHeaders.Add , "ReservesSort", "Sort Reserves" 'Hidden
        .ColumnHeaders.Add , "IsDeleted", "Is Deleted" 'Hidden
        .ColumnHeaders.Add , "UpLoadMe", "UpLoad Me" 'Hidden
        .ColumnHeaders.Add , "DateLastUpdated", "Date Last Updated"
        .ColumnHeaders.Add , "DateLastUpdatedSort", "Date Last Updated"
        .ColumnHeaders.Add , "AdminComments", "Admin Comments" 'Hidden
        .ColumnHeaders.Add , "ID", "ID" 'Hidden
        .ColumnHeaders.Add , "IDAssignments", "IDAssignments" 'Hidden
        
        .Sorted = False
        .SortOrder = lvwAscending
        'Class
        .ColumnHeaders.Item(GuiPolicyLimits.ClassTypeClass).Width = 700
        .ColumnHeaders.Item(GuiPolicyLimits.ClassTypeClass).Alignment = lvwColumnLeft
        'ClassTypeID
        .ColumnHeaders.Item(GuiPolicyLimits.ClassTypeID).Width = 0 ' hidden
        .ColumnHeaders.Item(GuiPolicyLimits.ClassTypeID).Alignment = lvwColumnLeft
        'Description
        .ColumnHeaders.Item(GuiPolicyLimits.ClassTypeDescription).Width = 1800
        .ColumnHeaders.Item(GuiPolicyLimits.ClassTypeDescription).Alignment = lvwColumnLeft
        'Limit Amount
        .ColumnHeaders.Item(GuiPolicyLimits.LimitAmount).Width = 1400
        .ColumnHeaders.Item(GuiPolicyLimits.LimitAmount).Alignment = lvwColumnRight
        'Sort Limit Amount
        .ColumnHeaders.Item(GuiPolicyLimits.LimitAmountSort).Width = 0 'hidden
        .ColumnHeaders.Item(GuiPolicyLimits.LimitAmountSort).Alignment = lvwColumnRight
        'RC Said
        .ColumnHeaders.Item(GuiPolicyLimits.RCSaidProp).Width = 1400
        .ColumnHeaders.Item(GuiPolicyLimits.RCSaidProp).Alignment = lvwColumnRight
        'Sort RC Said
        .ColumnHeaders.Item(GuiPolicyLimits.RCSaidPropSort).Width = 0  'hidden
        .ColumnHeaders.Item(GuiPolicyLimits.RCSaidPropSort).Alignment = lvwColumnRight
        'Reserves
        .ColumnHeaders.Item(GuiPolicyLimits.Reserves).Width = 1400
        .ColumnHeaders.Item(GuiPolicyLimits.Reserves).Alignment = lvwColumnRight
        'Sort Reserves
        .ColumnHeaders.Item(GuiPolicyLimits.ReservesSort).Width = 0  'hidden
        .ColumnHeaders.Item(GuiPolicyLimits.ReservesSort).Alignment = lvwColumnRight
        'Is Deleted
        If bHideDeleted Then
            .ColumnHeaders.Item(GuiPolicyLimits.IsDeleted).Width = 0 'hidden 400
        Else
            .ColumnHeaders.Item(GuiPolicyLimits.IsDeleted).Width = 400
        End If
        .ColumnHeaders.Item(GuiPolicyLimits.IsDeleted).Alignment = lvwColumnCenter
        .ColumnHeaders.Item(GuiPolicyLimits.IsDeleted).Icon = GuiPolicyLimitsPic.IsDeleted
        'UpLoad Me
        If bHideUploadFlags Then
            .ColumnHeaders.Item(GuiPolicyLimits.UpLoadMe).Width = 0 'hidden 400
        Else
            .ColumnHeaders.Item(GuiPolicyLimits.UpLoadMe).Width = 400
        End If
        .ColumnHeaders.Item(GuiPolicyLimits.UpLoadMe).Alignment = lvwColumnCenter
        .ColumnHeaders.Item(GuiPolicyLimits.UpLoadMe).Icon = GuiPolicyLimitsPic.UpLoadMe
        'DateLastUpdated
        .ColumnHeaders.Item(GuiPolicyLimits.DateLastUpdated).Width = 2200
        .ColumnHeaders.Item(GuiPolicyLimits.DateLastUpdated).Alignment = lvwColumnLeft
        'DateLastUpdatedSort
        .ColumnHeaders.Item(GuiPolicyLimits.DateLastUpdatedSort).Width = 0  'Hidden Sort Column
        .ColumnHeaders.Item(GuiPolicyLimits.DateLastUpdatedSort).Alignment = lvwColumnLeft
        'AdminComments
        .ColumnHeaders.Item(GuiPolicyLimits.AdminComments).Width = 0 'hidden 10000
        .ColumnHeaders.Item(GuiPolicyLimits.AdminComments).Alignment = lvwColumnLeft
        'ID
        .ColumnHeaders.Item(GuiPolicyLimits.ID).Width = 0  'hidden
        .ColumnHeaders.Item(GuiPolicyLimits.ID).Alignment = lvwColumnLeft
        'IDAssignments
        .ColumnHeaders.Item(GuiPolicyLimits.IDAssignments).Width = 0  'Hidden
        .ColumnHeaders.Item(GuiPolicyLimits.IDAssignments).Alignment = lvwColumnLeft
    End With
    
    bGridOn = CBool(GetSetting(App.EXEName, "GENERAL", "GRID_ON", False))
    lvwPLClassTypeID.GridLines = bGridOn
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub LoadHeaderPLClassTypeID"
End Sub


Private Function PopulatePolicyLimitControls() As Boolean
    On Error GoTo EH
    
    If mitmXPLSelected Is Nothing Then
        Exit Function
    End If
    
    lblPLCaption.Caption = mitmXPLSelected.SubItems(GuiPolicyLimits.ClassTypeDescription - 1)
    txtPLLimitAmount.Text = mitmXPLSelected.SubItems(GuiPolicyLimits.LimitAmount - 1)
    txtPLRCSaidProp.Text = mitmXPLSelected.SubItems(GuiPolicyLimits.RCSaidProp - 1)
    txtPLReserves.Text = mitmXPLSelected.SubItems(GuiPolicyLimits.Reserves - 1)
    If goUtil.GetFlagFromText(mitmXPLSelected.SubItems(GuiPolicyLimits.IsDeleted - 1)) Then
        chkPLIsDeleted.Value = vbChecked
    Else
        chkPLIsDeleted.Value = vbUnchecked
    End If
    txtPLAdminComments.Text = mitmXPLSelected.SubItems(GuiPolicyLimits.AdminComments - 1)
    
    
    PopulatePolicyLimitControls = True
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Function PopulatePolicyLimitControls"
End Function

Public Function BuildTextLossReport() As String
    On Error GoTo EH
    Dim sTextLossReport As String
    Dim sPolicyLimit As String
    Dim sTab As String
    Dim sAdjuster As String
    Dim sFName As String
    Dim sLName As String
    Dim sUserName As String
    Dim itmX As ListItem
    
    sTab = "    "
    sAdjuster = ""
    sFName = GetSetting(goUtil.gsMainAppEXEName, "GENERAL", "ADJUSTOR_FIRST_NAME", vbNullString)
    sLName = GetSetting(goUtil.gsMainAppEXEName, "GENERAL", "ADJUSTOR_LAST_NAME", vbNullString)
    sUserName = goUtil.utGetECSCryptSetting("ECS", "WEB_SECURITY", "USER_NAME", vbNullString)
    sAdjuster = sLName & ", " & sFName & " - " & sUserName
    
    'Title
    sTextLossReport = sTextLossReport & " " & vbCrLf
    sTextLossReport = sTextLossReport & "-------------------------------------LOSS REPORT-------------------------------------" & vbCrLf
    sTextLossReport = sTextLossReport & "Reported By: " & txtReportedBy.Text & vbCrLf
    sTextLossReport = sTextLossReport & "Phone:       " & txtReportedByPhone.Text & vbCrLf
    sTextLossReport = sTextLossReport & "-------------------------------------------------------------------------------------" & vbCrLf
    sTextLossReport = sTextLossReport & vbCrLf
    'Assigned Date and Adjuster
    sTextLossReport = sTextLossReport & "Assigned Date:      " & txtAssignedDate.Text & sTab & "Adjuster: " & sAdjuster & vbCrLf
    'Cat, Claim and Policy Info
    sTextLossReport = sTextLossReport & "Cat Code:           " & cboCatCode.Text & vbCrLf
    sTextLossReport = sTextLossReport & "Assignment Type:    " & cboAssignmentType.Text & sTab & "Type Of Loss: " & cboTypeOfLoss.Text & vbCrLf
    sTextLossReport = sTextLossReport & "Claim Number:       " & txtCLIENTNUM.Text & vbCrLf
    sTextLossReport = sTextLossReport & "Policy Number:      " & txtPolicyNo.Text & vbCrLf
    sTextLossReport = sTextLossReport & "Policy Description: " & txtPolicyDescription.Text & vbCrLf
    sTextLossReport = sTextLossReport & "Deductible:         " & txtDeductible.Text & vbCrLf
    'Policy Limits Need to loop through them all
    sTextLossReport = sTextLossReport & "Policy Limits: " & vbCrLf
    For Each itmX In lvwPLClassTypeID.ListItems
        If Not goUtil.GetFlagFromText(itmX.SubItems(GuiPolicyLimits.IsDeleted - 1)) Then
            sPolicyLimit = itmX.Text & " - "
            sPolicyLimit = sPolicyLimit & "(" & itmX.SubItems(GuiPolicyLimits.ClassTypeDescription - 1) & ")"
            sPolicyLimit = sPolicyLimit & " = "
            sPolicyLimit = sPolicyLimit & itmX.SubItems(GuiPolicyLimits.LimitAmount - 1)
            sTextLossReport = sTextLossReport & sPolicyLimit & vbCrLf
        End If
    Next
    'Insured and Addresses
    sTextLossReport = sTextLossReport & "Insured:        " & txtInsured.Text & vbCrLf
    sTextLossReport = sTextLossReport & "Home Phone:     " & txtHomePhone.Text & vbCrLf
    sTextLossReport = sTextLossReport & "Business Phone: " & txtBusinessPhone.Text & vbCrLf & vbCrLf
    
    sTextLossReport = sTextLossReport & "Mortgagee Name: " & txtMortgageeName.Text & vbCrLf & vbCrLf
    
    sTextLossReport = sTextLossReport & "Property Address: " & vbCrLf
    sTextLossReport = sTextLossReport & txtPAStreet.Text & vbCrLf
    sTextLossReport = sTextLossReport & txtPACity.Text & ", " & left(cboPAState.Text, 2) & " " & txtPAZIP.Text & "-" & txtPAZIP4.Text & vbCrLf & vbCrLf
    sTextLossReport = sTextLossReport & "Mailing Address:  " & vbCrLf
    sTextLossReport = sTextLossReport & txtMAStreet.Text & vbCrLf
    sTextLossReport = sTextLossReport & txtMACity.Text & ", " & left(cboMAState.Text, 2) & " " & txtMAZIP.Text & "-" & txtMAZIP4.Text & vbCrLf
    sTextLossReport = sTextLossReport & vbCrLf
    sTextLossReport = sTextLossReport & "-----------------------------------ADDITIONAL INFO-----------------------------------" & vbCrLf
    sTextLossReport = sTextLossReport & vbCrLf
    
    BuildTextLossReport = sTextLossReport
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function BuildTextLossReport"
End Function

Private Sub txtReportedByPhone_LostFocus()
    goUtil.utValidate , txtReportedByPhone
End Sub
