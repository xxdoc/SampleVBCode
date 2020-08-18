VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFixClaims 
   Caption         =   "Preview And Validate Claims"
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
   Icon            =   "frmFixClaims.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6585
   ScaleWidth      =   9480
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame framMain 
      Appearance      =   0  'Flat
      Caption         =   "Claim &Items (Uploaded Claims)"
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
      Begin VB.CheckBox chkWaitForUserOK 
         Caption         =   "OK items when sending"
         Height          =   220
         Left            =   5160
         TabIndex        =   9
         Top             =   2880
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.CommandButton cmdViewFailed 
         Caption         =   "Vie&w Failed"
         Height          =   375
         Left            =   3600
         TabIndex        =   7
         Top             =   2680
         Width           =   1455
      End
      Begin VB.CommandButton cmdFindNext 
         Caption         =   "Find &Next"
         Height          =   375
         Left            =   2400
         TabIndex        =   6
         Top             =   2680
         Width           =   1100
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Find"
         Height          =   375
         Left            =   1200
         TabIndex        =   5
         Top             =   2680
         Width           =   1100
      End
      Begin VB.CheckBox chkShowGrid 
         Caption         =   "Show Grid"
         Height          =   240
         Left            =   7920
         TabIndex        =   10
         Top             =   2670
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.Timer Timer_Resize 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   8880
         Top             =   0
      End
      Begin MSComctlLib.ImageList imglstClaims 
         Left            =   8520
         Top             =   240
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   7
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFixClaims.frx":030A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFixClaims.frx":075C
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFixClaims.frx":0BAE
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFixClaims.frx":0D08
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFixClaims.frx":115A
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFixClaims.frx":15AC
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFixClaims.frx":19FE
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdDown 
         Caption         =   "&."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   1.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   660
         Picture         =   "frmFixClaims.frx":1D18
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Next Record"
         Top             =   2680
         Width           =   400
      End
      Begin VB.CommandButton cmdUp 
         Caption         =   "&,"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   1.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Picture         =   "frmFixClaims.frx":215A
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Previous Record"
         Top             =   2680
         Width           =   400
      End
      Begin VB.CheckBox chkBrowseBadData 
         Caption         =   "Find failed validation only"
         Height          =   220
         Left            =   5160
         TabIndex        =   8
         Top             =   2670
         Value           =   1  'Checked
         Width           =   2655
      End
      Begin MSComctlLib.ListView lvwClaims 
         Height          =   2415
         Left            =   120
         TabIndex        =   2
         Tag             =   "Enable"
         Top             =   240
         Width           =   8985
         _ExtentX        =   15849
         _ExtentY        =   4260
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "imgFixXact"
         SmallIcons      =   "imglstClaims"
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
      TabIndex        =   15
      Top             =   3240
      Width           =   6015
      Begin VB.CommandButton cmdPrintList 
         Caption         =   "&View List"
         Height          =   855
         Left            =   3840
         MaskColor       =   &H00000000&
         Picture         =   "frmFixClaims.frx":259C
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Exit"
         Top             =   360
         Width           =   975
      End
      Begin VB.ComboBox cboBillingsVS 
         Height          =   360
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   420
         Width           =   2535
      End
      Begin VB.ComboBox cboBillingsSpeed 
         Height          =   360
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   790
         Width           =   2535
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   855
         Left            =   4920
         MaskColor       =   &H00000000&
         Picture         =   "frmFixClaims.frx":2E66
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Exit"
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdSendAll 
         Caption         =   "Send &All"
         Enabled         =   0   'False
         Height          =   855
         Left            =   120
         Picture         =   "frmFixClaims.frx":3170
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Send All"
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblBillingsVS 
         Caption         =   "Billings Settings"
         Height          =   255
         Left            =   1200
         TabIndex        =   17
         Top             =   165
         Width           =   2175
      End
   End
   Begin VB.Frame framSendStatus 
      Appearance      =   0  'Flat
      Caption         =   "Item Status"
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
      TabIndex        =   11
      Top             =   3240
      Width           =   3135
      Begin VB.OptionButton optSent 
         Caption         =   "Sent"
         Enabled         =   0   'False
         Height          =   855
         Left            =   2160
         Picture         =   "frmFixClaims.frx":347A
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   360
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.OptionButton optSkip 
         Caption         =   "Ski&p"
         Height          =   855
         Left            =   1140
         Picture         =   "frmFixClaims.frx":38BC
         Style           =   1  'Graphical
         TabIndex        =   13
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
         Picture         =   "frmFixClaims.frx":3CFE
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Mark to Send"
         Top             =   360
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin VB.Frame fram03 
      Appearance      =   0  'Flat
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
      Left            =   5280
      TabIndex        =   40
      Top             =   4680
      Width           =   4095
      Begin VB.TextBox txtb17sComments 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   0
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   41
         Top             =   360
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton cmdComments 
         Cancel          =   -1  'True
         Caption         =   "&Comments"
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
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   120
         Width           =   1095
      End
      Begin VB.TextBox txtb14sPROPADDR 
         Height          =   375
         Left            =   1320
         MaxLength       =   254
         TabIndex        =   52
         Top             =   1320
         Width           =   2655
      End
      Begin VB.TextBox txtb07cADMINFEE 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   2640
         MaxLength       =   40
         TabIndex        =   50
         Top             =   960
         Width           =   1328
      End
      Begin VB.TextBox txtb06cEXPENSEREIM 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   1320
         MaxLength       =   40
         TabIndex        =   49
         Top             =   960
         Width           =   1328
      End
      Begin VB.TextBox txtb05cSERVICEFEE 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   2640
         MaxLength       =   40
         TabIndex        =   47
         Top             =   600
         Width           =   1328
      End
      Begin VB.TextBox txtb04cGROSSLOSS 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   1320
         MaxLength       =   40
         TabIndex        =   46
         Top             =   600
         Width           =   1328
      End
      Begin VB.TextBox txtb11sCLAIMSTATE 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   3480
         MaxLength       =   2
         TabIndex        =   44
         Top             =   240
         Width           =   495
      End
      Begin VB.ComboBox cbob10sCLAIMCITY 
         Height          =   360
         Left            =   1320
         TabIndex        =   43
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label lblFixClaim 
         Caption         =   "Property Addr."
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
         Index           =   11
         Left            =   60
         TabIndex        =   51
         Top             =   1425
         Width           =   3015
      End
      Begin VB.Label lblFixClaim 
         Caption         =   "Gross | Srv Fee"
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
         Index           =   9
         Left            =   60
         TabIndex        =   45
         Top             =   720
         Width           =   3015
      End
      Begin VB.Label lblFixClaim 
         Caption         =   "Exp Rm | Admin"
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
         Index           =   10
         Left            =   60
         TabIndex        =   48
         Top             =   1080
         Width           =   3015
      End
      Begin VB.Label lblFixClaim 
         Caption         =   "Claim City | State"
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
         Index           =   8
         Left            =   60
         TabIndex        =   42
         Top             =   360
         Width           =   3015
      End
   End
   Begin VB.Frame fram02 
      Appearance      =   0  'Flat
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
      Left            =   2520
      TabIndex        =   31
      Top             =   4680
      Width           =   2655
      Begin VB.ComboBox cbob09sADJNAME 
         Height          =   360
         Left            =   720
         TabIndex        =   37
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox txtb08sFULLNAME 
         Height          =   375
         Left            =   720
         MaxLength       =   100
         TabIndex        =   39
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox txtb02sCLAIMNO 
         Height          =   375
         Left            =   720
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   35
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txtb03sIB 
         Height          =   375
         Left            =   720
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   33
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lblFixClaim 
         Caption         =   "Full Nam"
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
         Left            =   60
         TabIndex        =   38
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label lblFixClaim 
         Caption         =   "IB"
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
         Left            =   60
         TabIndex        =   32
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label lblFixClaim 
         Caption         =   "Adjuster"
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
         Index           =   6
         Left            =   60
         TabIndex        =   36
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label lblFixClaim 
         Caption         =   "Claim No"
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
         Left            =   60
         TabIndex        =   34
         Top             =   720
         Width           =   2415
      End
   End
   Begin VB.Frame fram01 
      Appearance      =   0  'Flat
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
      TabIndex        =   22
      Top             =   4680
      Width           =   2295
      Begin VB.ComboBox cbob01sCATNO 
         Height          =   360
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtb12dtFILESRECD 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   960
         MaxLength       =   10
         TabIndex        =   30
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txtb13dtDATECLOSED 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   960
         MaxLength       =   10
         TabIndex        =   28
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtEntryDate 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   960
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   26
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblFixClaim 
         Caption         =   "Date Loss"
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
         Index           =   7
         Left            =   60
         TabIndex        =   29
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label lblFixClaim 
         Caption         =   "Entry Date"
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
         Left            =   60
         TabIndex        =   25
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label lblFixClaim 
         Caption         =   "Date Closed"
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
         Left            =   60
         TabIndex        =   27
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label lblFixClaim 
         Caption         =   "Cat No"
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
         Left            =   60
         TabIndex        =   23
         Top             =   360
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmFixClaims"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const EDIT_HEIGHT As Long = 6960
Private Const VIEW_HEIGHT As Long = 5050
Private Const RED_BG As Long = &HC0C0FF
Private Const WHITE_BG As Long = &H80000005
Private mitmX As ListItem
Private mClaimItem As udtClaimItem
Private mcolClaimItems As Collection
Private mcolValidClaimItems As Collection
Private mbItemsValidated As Boolean
Private moSend2Billings As clsSend2Billings
Private mbAlreadySent As Boolean
Private mbLoadingEdit As Boolean
Private mbSkipAll As Boolean
Private mbModalFlag As Boolean
Private mcolCATNO As Collection
Private mcolADJNAME As Collection
Private mcolCLAIMCITY As Collection
Private mcolTAXSTATES As Collection
Private msFindText As String
Private mlLastFindIndex As Long

'Control Size in Ref to Form Diff Constants
Private Const FORM_W As Long = 9700
Private Const FORM_H As Long = 6990
Private Const framMain_H As Long = 3825
Private Const framMain_W As Long = 315
Private Const lvwClaims_H As Long = 720
Private Const lvwClaims_W As Long = 270
Private Const cmdUp_T As Long = 455
Private Const cmdDown_T As Long = 455
Private Const cmdFind_T As Long = 455
Private Const cmdFindNext_T As Long = 455
Private Const cmdViewFailed_T As Long = 455
Private Const chkBrowseBadData_T As Long = 465
Private Const chkWaitForUserOK_T As Long = 255
Private Const chkShowGrid_T As Long = 465
Private Const framSendStatus_T As Long = 3720
Private Const framCommands_T As Long = 3720
Private Const framCommands_W As Long = 3585
Private Const cmdPrintList_L As Long = 2175
Private Const cmdExit_L As Long = 1095
'Fram 1
Private Const fram01_T As Long = 2280
'fram 2
Private Const fram02_L As Long = 2520
Private Const fram02_T As Long = 2280
Private Const fram02_W As Long = 6945
Private Const txtb03sIB_W As Long = 840
Private Const txtb02sCLAIMNO_W As Long = 840
Private Const cbob09sADJNAME_W As Long = 840
Private Const txtb08sFULLNAME_W As Long = 840
'Fram 3
Private Const fram03_T As Long = 2280
Private Const fram03_L As Long = 105
Private Const fram03_W As Long = 225
Private Const cbob10sCLAIMCITY_W As Long = 1920
Private Const txtb11sCLAIMSTATE_L As Long = 15
Private Const txtb04cGROSSLOSS_W As Long = 2767
Private Const txtb05cSERVICEFEE_L As Long = 8
Private Const txtb06cEXPENSEREIM_W As Long = 2767
Private Const txtb07cADMINFEE_L As Long = 8
Private Const txtb14sPROPADDR_W As Long = 1440


Private mbResize As Boolean


Public Property Let colCATNO(pcolObject As Collection)
    Set mcolCATNO = pcolObject
End Property
Public Property Set colCATNO(pcolObject As Collection)
    Set mcolCATNO = pcolObject
End Property
Public Property Get colCATNO() As Collection
    Set colCATNO = mcolCATNO
End Property

Public Property Let colADJNAME(pcolObject As Collection)
    Set mcolADJNAME = pcolObject
End Property
Public Property Set colADJNAME(pcolObject As Collection)
    Set mcolADJNAME = pcolObject
End Property
Public Property Get colADJNAME() As Collection
    Set colADJNAME = mcolADJNAME
End Property

Public Property Let colCLAIMCITY(pcolObject As Collection)
    Set mcolCLAIMCITY = pcolObject
End Property
Public Property Set colCLAIMCITY(pcolObject As Collection)
    Set mcolCLAIMCITY = pcolObject
End Property
Public Property Get colCLAIMCITY() As Collection
    Set colCLAIMCITY = mcolCLAIMCITY
End Property

Public Property Let colTAXSTATES(pcolObject As Collection)
    Set mcolTAXSTATES = pcolObject
End Property
Public Property Set colTAXSTATES(pcolObject As Collection)
    Set mcolTAXSTATES = pcolObject
End Property
Public Property Get colTAXSTATES() As Collection
    Set colTAXSTATES = mcolTAXSTATES
End Property

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
Public Property Get ItemsValidated() As Boolean
    ItemsValidated = mbItemsValidated
End Property

Public Property Let ClaimItems(pcolClaimItems As Collection)
    Set mcolClaimItems = pcolClaimItems
End Property
Public Property Set ClaimItems(pcolClaimItems As Collection)
    Set mcolClaimItems = pcolClaimItems
End Property
Public Property Get ClaimItems() As Collection
    Set ClaimItems = mcolClaimItems
End Property
Public Property Let Send2Billings(poSend2Billings As clsSend2Billings)
    Set moSend2Billings = poSend2Billings
End Property
Public Property Set Send2Billings(poSend2Billings As clsSend2Billings)
    Set moSend2Billings = poSend2Billings
End Property

Private Sub LoadHeaderClaimItems()
    On Error GoTo EH
    Dim bGridOn As Boolean
    Dim bOKOn As Boolean
    
    'set the columnheaders
    With lvwClaims
      
        .ColumnHeaders.Add , "Status", "Status"
        .ColumnHeaders.Add , "Data", "Data"
        .ColumnHeaders.Add , "BillingDup", "Bill Dup"
        .ColumnHeaders.Add , "b01sCATNO", "CAT.NO"
        .ColumnHeaders.Add , "b01sCATNOSort", "CAT.NOSort"
        .ColumnHeaders.Add , "b03sIB", "IB"
        .ColumnHeaders.Add , "b03sIBSort", "IBSort"
        .ColumnHeaders.Add , "b02sCLAIMNO", "CLAIM.NO"
        .ColumnHeaders.Add , "b02sCLAIMNOSort", "CLAIM.NOSort"
        .ColumnHeaders.Add , "b15dtDateUploaded", "DATE UL"
        .ColumnHeaders.Add , "b15dtDateUploadedSort", "DATE ULSort"
        .ColumnHeaders.Add , "b16dtDateEntered", "ENTRY.DATE"
        .ColumnHeaders.Add , "b16dtDateEnteredSort", "ENTRY.DATESort"
        .ColumnHeaders.Add , "b13dtDATECLOSED", "DATE.CLOSED"
        .ColumnHeaders.Add , "b13dtDATECLOSEDSort", "DATE.CLOSEDSort"
        .ColumnHeaders.Add , "b12dtFILESRECD", "DATE OF LOSS"
        .ColumnHeaders.Add , "b12dtFILESRECDSort", "DATE OF LOSSSort"
        .ColumnHeaders.Add , "b09sADJNAME", "ADJUSTER.NAME"
        .ColumnHeaders.Add , "b08sFULLNAME", "FULL.NAME"
        .ColumnHeaders.Add , "b10sCLAIMCITY", "CLAIM.CITY"
        .ColumnHeaders.Add , "b11sCLAIMSTATE", "CLAIM.STATE"
        .ColumnHeaders.Add , "b04cGROSSLOSS", "GROSS.LOSS"
        .ColumnHeaders.Add , "b04cGROSSLOSSSort", "GROSS.LOSSSort"
        .ColumnHeaders.Add , "b05cSERVICEFEE", "SERVICE.FEE"
        .ColumnHeaders.Add , "b05cSERVICEFEESort", "SERVICE.FEESort"
        .ColumnHeaders.Add , "b06cEXPENSEREIM", "ADJ 100% Exp.Reim"
        .ColumnHeaders.Add , "b06cEXPENSEREIMSort", "ADJ 100% Exp.ReimSort"
        .ColumnHeaders.Add , "b07cADMINFEE", "Admin_Fee"
        .ColumnHeaders.Add , "b07cADMINFEESort", "Admin_FeeSort"
        .ColumnHeaders.Add , "b14sPROPADDR", "Property Address"
        .ColumnHeaders.Add , "b14sPROPADDRSort", "Property AddressSort"
        .ColumnHeaders.Add , "b17sComments", "Comments"

        .Sorted = False
        
        .ColumnHeaders.Item(FixClaimItem.Status).Width = 800
        .ColumnHeaders.Item(FixClaimItem.Status).Alignment = lvwColumnLeft
        .ColumnHeaders.Item(FixClaimItem.Data).Width = 600
        .ColumnHeaders.Item(FixClaimItem.Data).Alignment = lvwColumnLeft
        .ColumnHeaders.Item(FixClaimItem.BillingDup).Width = 600
        .ColumnHeaders.Item(FixClaimItem.BillingDup).Alignment = lvwColumnLeft
        .ColumnHeaders.Item(FixClaimItem.b01sCATNO).Width = 1500
        .ColumnHeaders.Item(FixClaimItem.b01sCATNO).Alignment = lvwColumnLeft
        .ColumnHeaders.Item(FixClaimItem.b01sCATNOSort).Width = 0
        .ColumnHeaders.Item(FixClaimItem.b01sCATNOSort).Alignment = lvwColumnLeft
        .ColumnHeaders.Item(FixClaimItem.b03sIB).Width = 2300
        .ColumnHeaders.Item(FixClaimItem.b03sIB).Alignment = lvwColumnLeft
        .ColumnHeaders.Item(FixClaimItem.b03sIBSort).Width = 0
        .ColumnHeaders.Item(FixClaimItem.b03sIBSort).Alignment = lvwColumnLeft
        .ColumnHeaders.Item(FixClaimItem.b02sCLAIMNO).Width = 1200
        .ColumnHeaders.Item(FixClaimItem.b02sCLAIMNO).Alignment = lvwColumnLeft
        .ColumnHeaders.Item(FixClaimItem.b02sCLAIMNOSort).Width = 0
        .ColumnHeaders.Item(FixClaimItem.b02sCLAIMNOSort).Alignment = lvwColumnLeft
        .ColumnHeaders.Item(FixClaimItem.b15dtDateUploaded).Width = 1300
        .ColumnHeaders.Item(FixClaimItem.b15dtDateUploaded).Alignment = lvwColumnLeft
        .ColumnHeaders.Item(FixClaimItem.b15dtDateUploadedSort).Width = 0
        .ColumnHeaders.Item(FixClaimItem.b15dtDateUploadedSort).Alignment = lvwColumnLeft
        .ColumnHeaders.Item(FixClaimItem.b16dtDateEntered).Width = 1300
        .ColumnHeaders.Item(FixClaimItem.b16dtDateEntered).Alignment = lvwColumnLeft
        .ColumnHeaders.Item(FixClaimItem.b16dtDateEnteredSort).Width = 0
        .ColumnHeaders.Item(FixClaimItem.b16dtDateEnteredSort).Alignment = lvwColumnLeft
        .ColumnHeaders.Item(FixClaimItem.b13dtDATECLOSED).Width = 1300
        .ColumnHeaders.Item(FixClaimItem.b13dtDATECLOSED).Alignment = lvwColumnLeft
        .ColumnHeaders.Item(FixClaimItem.b13dtDATECLOSEDSort).Width = 0
        .ColumnHeaders.Item(FixClaimItem.b13dtDATECLOSEDSort).Alignment = lvwColumnLeft
        .ColumnHeaders.Item(FixClaimItem.b12dtFILESRECD).Width = 1300
        .ColumnHeaders.Item(FixClaimItem.b12dtFILESRECD).Alignment = lvwColumnLeft
        .ColumnHeaders.Item(FixClaimItem.b12dtFILESRECDSort).Width = 0
        .ColumnHeaders.Item(FixClaimItem.b12dtFILESRECDSort).Alignment = lvwColumnLeft
        .ColumnHeaders.Item(FixClaimItem.b09sADJNAME).Width = 2000
        .ColumnHeaders.Item(FixClaimItem.b09sADJNAME).Alignment = lvwColumnLeft
        .ColumnHeaders.Item(FixClaimItem.b08sFULLNAME).Width = 2000
        .ColumnHeaders.Item(FixClaimItem.b08sFULLNAME).Alignment = lvwColumnLeft
        .ColumnHeaders.Item(FixClaimItem.b10sCLAIMCITY).Width = 2000
        .ColumnHeaders.Item(FixClaimItem.b10sCLAIMCITY).Alignment = lvwColumnLeft
        .ColumnHeaders.Item(FixClaimItem.b11sCLAIMSTATE).Width = 1200
        .ColumnHeaders.Item(FixClaimItem.b11sCLAIMSTATE).Alignment = lvwColumnLeft
        .ColumnHeaders.Item(FixClaimItem.b04cGROSSLOSS).Width = 2000
        .ColumnHeaders.Item(FixClaimItem.b04cGROSSLOSS).Alignment = lvwColumnRight
        .ColumnHeaders.Item(FixClaimItem.b04cGROSSLOSSSort).Width = 0
        .ColumnHeaders.Item(FixClaimItem.b04cGROSSLOSSSort).Alignment = lvwColumnRight
        .ColumnHeaders.Item(FixClaimItem.b05cSERVICEFEE).Width = 2000
        .ColumnHeaders.Item(FixClaimItem.b05cSERVICEFEE).Alignment = lvwColumnRight
        .ColumnHeaders.Item(FixClaimItem.b05cSERVICEFEESort).Width = 0
        .ColumnHeaders.Item(FixClaimItem.b05cSERVICEFEESort).Alignment = lvwColumnRight
        .ColumnHeaders.Item(FixClaimItem.b06cEXPENSEREIM).Width = 2000
        .ColumnHeaders.Item(FixClaimItem.b06cEXPENSEREIM).Alignment = lvwColumnRight
        .ColumnHeaders.Item(FixClaimItem.b06cEXPENSEREIMSort).Width = 0
        .ColumnHeaders.Item(FixClaimItem.b06cEXPENSEREIMSort).Alignment = lvwColumnRight
        .ColumnHeaders.Item(FixClaimItem.b07cADMINFEE).Width = 2000
        .ColumnHeaders.Item(FixClaimItem.b07cADMINFEE).Alignment = lvwColumnRight
        .ColumnHeaders.Item(FixClaimItem.b07cADMINFEESort).Width = 0
        .ColumnHeaders.Item(FixClaimItem.b07cADMINFEESort).Alignment = lvwColumnRight
        .ColumnHeaders.Item(FixClaimItem.b14sPROPADDR).Width = 4000
        .ColumnHeaders.Item(FixClaimItem.b14sPROPADDR).Alignment = lvwColumnLeft
        .ColumnHeaders.Item(FixClaimItem.b14sPROPADDRSort).Width = 0
        .ColumnHeaders.Item(FixClaimItem.b14sPROPADDRSort).Alignment = lvwColumnLeft
        .ColumnHeaders.Item(FixClaimItem.b17sComments).Width = 10000
        .ColumnHeaders.Item(FixClaimItem.b17sComments).Alignment = lvwColumnLeft
    End With
    
    bGridOn = CBool(GetSetting(App.EXEName, "GENERAL", "GRID_ON", False))
    If bGridOn Then
        chkShowGrid.Value = vbChecked
    Else
        chkShowGrid.Value = vbUnchecked
    End If
    
    lvwClaims.GridLines = bGridOn
    
    bOKOn = CBool(GetSetting(App.EXEName, "GENERAL", "OK_ON", False))
    If bOKOn Then
        chkWaitForUserOK = vbChecked
    Else
        chkWaitForUserOK = vbUnchecked
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub LoadHeaderClaimItems"
End Sub

Private Sub PopulatelvwClaims()
    On Error GoTo EH
    'Source Variables
    Dim vClaimItem  As Variant   'udtClaimItem '
    Dim ClaimItem As udtClaimItem
    Dim itmX As ListItem
    Dim iStatus As Integer
    Dim sStatus As String
    Dim resetcolClaimItems As Collection
    Dim lCount As Long
    Dim bPB As Boolean
    
    lvwClaims.ListItems.Clear

    If Not mcolClaimItems Is Nothing Then
        If Not moSend2Billings.ProgBar Is Nothing Then
            moSend2Billings.ProgBar.Max = mcolClaimItems.Count
            moSend2Billings.ProgBar.Value = 0
            bPB = True
        End If
        'Use this to Store Validated Claim Items
        'claims that go through inital validation from
        'collection of claim items.  It will be used to
        'reset the collection of claim items.
        Set resetcolClaimItems = New Collection
        
        For Each vClaimItem In mcolClaimItems
            ClaimItem = vClaimItem
            If ClaimItem.SentToBillings Then
                iStatus = FixClaimPic.AlreadySentToBilling
                sStatus = "Sent"
            Else
                iStatus = FixClaimPic.NeedsSentToBilling
                sStatus = "Send"
            End If
            
            If ClaimItem.SkipThisItem Then
                iStatus = FixClaimPic.Skip
                sStatus = "Skip"
            End If
                
            Set itmX = lvwClaims.ListItems.Add(, , sStatus, , iStatus)
            
            'Can't point out Rebilled items until the List View is Sorted by Ibnumber Desc
            'So pass in true to ignore prev client claim No
            'validate this claim item
            moSend2Billings.ValidateClaimItem ClaimItem, itmX, True
            
            itmX.Selected = False
            'add it to the reset collection
            resetcolClaimItems.Add ClaimItem, ClaimItem.Main.b03sIB
            
            lCount = lCount + 1
            If bPB Then
                If lCount <= moSend2Billings.ProgBar.Max Then
                    moSend2Billings.ProgBar.Value = lCount
                End If
            End If
        Next
        
        'Sort the List View by Ibnumber sort
        lvwClaims.SortOrder = lvwDescending
        lvwClaims.SortKey = FixClaimItem.b03sIBSort - 1
        lvwClaims.Sorted = True
        'Loop through again to point out the rebills
        For Each itmX In lvwClaims.ListItems
            ClaimItem = mcolClaimItems(itmX.ListSubItems(FixClaimItem.b03sIB - 1))
            'This time pass in false to not ignore the prev client claim No
            'That will be used to indicate if the Rebill Pointer should be used.
            moSend2Billings.ValidateClaimItem ClaimItem, itmX, False
        Next
        
        If bPB Then
            moSend2Billings.ProgBar.Value = 0
        End If
        
        'clear the claim item collection
        Set mcolClaimItems = Nothing
        'reset it to a collection of claims that have been
        'initialy validated, i.e. any claim that failed validation will
        'be flagged as such.
        Set mcolClaimItems = resetcolClaimItems
        EnableSend
        
    End If
    
CLEANUP:
    'Cleanup

    Set itmX = Nothing
    Set resetcolClaimItems = Nothing
    Exit Sub
EH:
    Set itmX = Nothing
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub PopulatelvwClaims"
End Sub

Private Sub cbob01sCATNO_Change()
    CATNO
End Sub

Private Sub cbob01sCATNO_click()
    CATNO
End Sub

Private Sub cbob01sCATNO_LostFocus()
    
    LoadEdit
End Sub

Private Sub CATNO()
    ValidateControl cbob01sCATNO, b01sCATNO, True
End Sub

Private Sub ADJNAME()
    ValidateControl cbob09sADJNAME, b09sADJNAME, True
End Sub

Private Sub cbob09sADJNAME_Change()
    ADJNAME
End Sub

Private Sub cbob09sADJNAME_Click()
    ADJNAME
End Sub

Private Sub cbob09sADJNAME_LostFocus()
    
    LoadEdit
End Sub

Private Sub CLAIMCITY()
    ValidateControl cbob10sCLAIMCITY, b10sCLAIMCITY, True
End Sub

Private Sub cbob10sCLAIMCITY_Change()
    CLAIMCITY
End Sub

Private Sub cbob10sCLAIMCITY_Click()
    CLAIMCITY
End Sub

Private Sub cbob10sCLAIMCITY_LostFocus()
    LoadEdit
End Sub

Private Sub cboBillingsSpeed_Click()
    On Error GoTo EH
    
    SaveSetting "ECS", "KEYBOARD", "SPEED", cboBillingsSpeed.Text
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cboBillingsSpeed_Click"
End Sub

Private Sub cboBillingsVS_Click()
    On Error GoTo EH
    
    SaveSetting goUtil.gsAppEXEName, "Version", "LatestVS", cboBillingsVS.Text
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cboBillingsVS_Click"
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
    lvwClaims.GridLines = bGridOn
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub chkShowGrid_Click"
End Sub


Private Sub chkWaitForUserOK_Click()
    On Error GoTo EH
    Dim bOKOn As Boolean
    
    If chkWaitForUserOK.Value = vbChecked Then
        bOKOn = True
    Else
        bOKOn = False
    End If
    
    SaveSetting App.EXEName, "GENERAL", "OK_ON", bOKOn
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub chkWaitForUserOK_Click"
End Sub

Private Sub cmdComments_Click()
    On Error GoTo EH
    If fram03.Enabled Then
        If cmdComments.Caption = "&Comments" Then
            txtb17sComments.Visible = True
            txtb17sComments.SetFocus
            cmdComments.Caption = "&OK"
            
        ElseIf cmdComments.Caption = "&OK" Then
            txtb17sComments.Visible = False
            cmdComments.Caption = "&Comments"
            lvwClaims.SetFocus
        End If
    End If
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdComments_Click"
End Sub

Private Sub cmdDown_Click()
    On Error GoTo EH
    Dim lCount As Long
    
    If Not mitmX Is Nothing Then
        lCount = mitmX.Index + 1
        Do Until lCount > lvwClaims.ListItems.Count
            Set mitmX = lvwClaims.ListItems(lCount)
            If chkBrowseBadData.Value = vbChecked Then
                If mitmX.ListSubItems(FixClaimItem.Data - 1).ReportIcon = FixClaimPic.BadData Then
                    If mitmX.SmallIcon <> FixClaimPic.Skip Then
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
    End If
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdDown_Click"
End Sub

Private Sub cmdExit_Click()
    On Error GoTo EH
    DoEvents
    moSend2Billings.ClaimItems = mcolClaimItems
    mbSkipAll = True
    LoadValidClaimItems
    moSend2Billings.CheckIsDirtyAndAdd mcolValidClaimItems
    Me.Hide
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdExit_Click"
    Me.Hide
End Sub

Private Sub cmdFind_Click()
    On Error GoTo EH
    If lvwClaims.ListItems.Count > 0 Then
        mlLastFindIndex = 0
        goUtil.utFindListItem Me, lvwClaims, msFindText, mlLastFindIndex
    End If
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdFind_Click"
End Sub

Private Sub cmdFindNext_Click()
    On Error GoTo EH
    
    If msFindText <> vbNullString And lvwClaims.ListItems.Count > 0 Then
        goUtil.utFindListItem Me, lvwClaims, msFindText, mlLastFindIndex
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdFindNext_Click"
End Sub

Private Sub cmdPrintList_Click()
    On Error GoTo EH
    Dim sTitle As String
    sTitle = "Claims "
    sTitle = sTitle & "(" & moSend2Billings.StartDate & " --> " & moSend2Billings.EndDate & ")"
    sTitle = sTitle & "  " & mcolClaimItems.Count & " Items"
    
    goUtil.utPrintListView App.EXEName, lvwClaims, "Claim Items"
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdPrintList_Click"
End Sub

Private Sub cmdViewFailed_Click()
    On Error GoTo EH
    Dim itmX As ListItem
    Dim lCount As Long
    
    
    'Select Alll the Failed Items in the List View
    If lvwClaims.ListItems.Count > 0 Then
        For lCount = 1 To lvwClaims.ListItems.Count
            Set itmX = lvwClaims.ListItems(lCount)
            If itmX.ListSubItems(FixClaimItem.Data - 1).ReportIcon = FixClaimPic.BadData Then
                itmX.Selected = True
            End If
        Next
    End If
    
    Set itmX = Nothing
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdViewFailed_Click"
End Sub

Private Sub cmdSendAll_Click()
    On Error GoTo EH
    cmdSendAll.SetFocus
    moSend2Billings.ClaimItems = mcolClaimItems
    mbSkipAll = False
    LoadValidClaimItems
    moSend2Billings.CheckIsDirtyAndAdd mcolValidClaimItems
    Me.Hide
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdSendAll_Click"
End Sub

Private Sub cmdUp_Click()
    On Error GoTo EH
    Dim lCount As Long
    
    If Not mitmX Is Nothing Then
        lCount = mitmX.Index - 1
        Do Until lCount <= 0
            Set mitmX = lvwClaims.ListItems(lCount)
            If chkBrowseBadData.Value = vbChecked Then
                If mitmX.ListSubItems(FixClaimItem.Data - 1).ReportIcon = FixClaimPic.BadData Then
                    If mitmX.SmallIcon <> FixClaimPic.Skip Then
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
    End If
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdUp_Click"
End Sub

Private Sub Form_Load()
    On Error GoTo EH
    mbResize = True
    goUtil.utFormWinRegPos App.EXEName, Me, , , , False
    mbResize = False
    
    LoadHeaderClaimItems
    PopulatelvwClaims
    PopulateLookupCbo
    PopulateSend2BillingsSettings
    
    goUtil.utSuffixLabels lblFixClaim
    
    Me.Caption = Me.Caption & " (" & moSend2Billings.StartDate & " --> " & moSend2Billings.EndDate & ")"
    Me.Caption = Me.Caption & "  " & mcolClaimItems.Count & " Items"
    Exit Sub
EH:
    mbResize = False
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_Load"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo EH
    
    'Hide instead of allowing unload if user closes window
    If UnloadMode = vbFormControlMenu Then
        cmdExit.SetFocus
        Cancel = True
        moSend2Billings.ClaimItems = mcolClaimItems
        mbSkipAll = True
        LoadValidClaimItems
        moSend2Billings.CheckIsDirtyAndAdd mcolValidClaimItems
        Me.Hide
    End If

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

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo EH
    mbResize = True
    goUtil.utFormWinRegPos App.EXEName, Me, True, , , False
    mbResize = False
    
    If Not goUtil.gfrmECTray Is Nothing Then
        goUtil.gfrmECTray.ShowMe False
    End If
    
    Exit Sub
EH:
    mbResize = False
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_Unload"
End Sub

Public Function CLEANUP() As Boolean
    On Error GoTo EH
    
    mbModalFlag = False
    
    Set mcolClaimItems = Nothing
    Set mcolValidClaimItems = Nothing
    Set moSend2Billings = Nothing
    Set mitmX = Nothing
    
    If Not gofrmMain Is Nothing Then
        If gofrmMain.UnloadFlag Then
            Set mcolCATNO = Nothing
            Set mcolADJNAME = Nothing
            Set mcolCLAIMCITY = Nothing
            Set mcolTAXSTATES = Nothing
        End If
    End If
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function CLEANUP"
End Function

Private Sub PopulateLookupCbo()
    On Error GoTo EH
    Dim vString As Variant
    Dim lCount As Long
    Dim bPB As Boolean
    
    If Not moSend2Billings.ProgBar Is Nothing Then
        bPB = True
    End If
    'CATNO
    If Not mcolCATNO Is Nothing Then
        cbob01sCATNO.Clear
        If bPB Then
            moSend2Billings.ProgBar.Max = mcolCATNO.Count
            moSend2Billings.ProgBar.Value = 0
            lCount = 0
        End If
        For Each vString In mcolCATNO
            cbob01sCATNO.AddItem vString
            If bPB Then
                lCount = lCount + 1
                If lCount <= moSend2Billings.ProgBar.Max Then
                    moSend2Billings.ProgBar.Value = lCount
                End If
            End If
        Next
    End If
    
    'ADJNAME
    If Not mcolADJNAME Is Nothing Then
        cbob09sADJNAME.Clear
        If bPB Then
            moSend2Billings.ProgBar.Max = mcolADJNAME.Count
            moSend2Billings.ProgBar.Value = 0
            lCount = 0
        End If
        For Each vString In mcolADJNAME
            cbob09sADJNAME.AddItem vString
            If bPB Then
                lCount = lCount + 1
                moSend2Billings.ProgBar.Value = lCount
            End If
        Next
    End If
    
    
    'CLAIMCITY
    If Not mcolCLAIMCITY Is Nothing Then
        cbob10sCLAIMCITY.Clear
        If bPB Then
            moSend2Billings.ProgBar.Max = mcolCLAIMCITY.Count
            moSend2Billings.ProgBar.Value = 0
            lCount = 0
        End If
        For Each vString In mcolCLAIMCITY
            cbob10sCLAIMCITY.AddItem vString
            If bPB Then
                lCount = lCount + 1
                moSend2Billings.ProgBar.Value = lCount
            End If
        Next
    End If
    
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub PopulateLookupCbo"
End Sub

Private Sub lblFixClaim_DblClick(Index As Integer)
    On Error GoTo EH
        cbob01sCATNO.Locked = False
        cbob01sCATNO.SetFocus
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lblFixClaim_DblClick"
End Sub

Private Sub lvwClaims_Click()
    LoadEdit True
End Sub

Private Sub LoadEdit(Optional pbUseSelected As Boolean)
    On Error GoTo EH
    Dim itmX As ListItem
    Dim bFound As Boolean
    Dim sZip5 As String
    Dim sZip4 As String
    
    mbLoadingEdit = True
    For Each itmX In lvwClaims.ListItems
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
        fram01.Enabled = True
        fram02.Enabled = True
        fram03.Enabled = True
        framSendStatus.Enabled = True
        optSend.Visible = True
        optSkip.Visible = True
        optSent.Visible = True
        
        With itmX
            Select Case .SmallIcon
                Case FixClaimPic.AlreadySentToBilling
                    optSent.Value = True
                    mbAlreadySent = True
                    mClaimItem.SentToBillings = True
                Case FixClaimPic.NeedsSentToBilling
                    optSend.Value = True
                    mbAlreadySent = False
                    mClaimItem.SentToBillings = False
                Case FixClaimPic.Skip
                    optSkip.Value = True
                    mbAlreadySent = False
                    mClaimItem.SkipThisItem = True
            End Select
            'Also check for Billing dup
            If .ListSubItems(FixClaimItem.BillingDup - 1).ReportIcon = FixClaimPic.BillingDup Then
                mClaimItem.BillingDup = True
            Else
                mClaimItem.BillingDup = False
            End If
            
            'Frame 01
            cbob01sCATNO.Text = .SubItems(FixClaimItem.b01sCATNO - 1)
            mClaimItem.Main.b01sCATNO = cbob01sCATNO.Text
            cbob01sCATNO.BackColor = IIf(.ListSubItems(FixClaimItem.b01sCATNO - 1).ReportIcon = FixClaimPic.BadData, RED_BG, WHITE_BG)
            '2.17.2003 Also need to unlock the item if it failed so user can change it.
            If cbob01sCATNO.BackColor = RED_BG Then
                cbob01sCATNO.Locked = False
            Else
                cbob01sCATNO.Locked = True
            End If
            
            'Date uploaded does not have a user entry view
            If .SubItems(FixClaimItem.b15dtDateUploaded - 1) <> vbNullString Then
                mClaimItem.Main.b15dtDateUploaded = .SubItems(FixClaimItem.b15dtDateUploaded - 1)
            Else
                mClaimItem.Main.b15dtDateUploaded = NULL_DATE
            End If
            'Date Entered has a user entry view but it is locked
            txtEntryDate.Text = .SubItems(FixClaimItem.b16dtDateEntered - 1)
            If txtEntryDate.Text <> vbNullString Then
                mClaimItem.Main.b16dtDateEntered = txtEntryDate.Text
            Else
                mClaimItem.Main.b16dtDateEntered = NULL_DATE
            End If
            
            txtb13dtDATECLOSED.Text = .SubItems(FixClaimItem.b13dtDATECLOSED - 1)
            If txtb13dtDATECLOSED.Text <> vbNullString Then
                mClaimItem.Main.b13dtDATECLOSED = txtb13dtDATECLOSED.Text
            Else
                mClaimItem.Main.b13dtDATECLOSED = NULL_DATE
            End If
            txtb13dtDATECLOSED.BackColor = IIf(.ListSubItems(FixClaimItem.b13dtDATECLOSED - 1).ReportIcon = FixClaimPic.BadData, RED_BG, WHITE_BG)
            
            txtb12dtFILESRECD.Text = .SubItems(FixClaimItem.b12dtFILESRECD - 1)
            If txtb12dtFILESRECD.Text <> vbNullString Then
                mClaimItem.Main.b12dtFILESRECD = txtb12dtFILESRECD.Text
            Else
                mClaimItem.Main.b12dtFILESRECD = NULL_DATE
            End If
            txtb12dtFILESRECD.BackColor = IIf(.ListSubItems(FixClaimItem.b12dtFILESRECD - 1).ReportIcon = FixClaimPic.BadData, RED_BG, WHITE_BG)
            
            'Frame 02
            'IB has a user entry view but it is locked
            txtb03sIB.Text = .SubItems(FixClaimItem.b03sIB - 1)
            mClaimItem.Main.b03sIB = txtb03sIB.Text
            txtb03sIB.BackColor = IIf(.ListSubItems(FixClaimItem.b03sIB - 1).ReportIcon = FixClaimPic.BadData, RED_BG, WHITE_BG)
            
            'ClaimNO Has a user entry view but it is locked
            txtb02sCLAIMNO.Text = .SubItems(FixClaimItem.b02sCLAIMNO - 1)
            mClaimItem.Main.b02sCLAIMNO = txtb02sCLAIMNO.Text
            txtb02sCLAIMNO.BackColor = IIf(.ListSubItems(FixClaimItem.b02sCLAIMNO - 1).ReportIcon = FixClaimPic.BadData, RED_BG, WHITE_BG)
            
            cbob09sADJNAME.Text = .SubItems(FixClaimItem.b09sADJNAME - 1)
            mClaimItem.Main.b09sADJNAME = cbob09sADJNAME.Text
            cbob09sADJNAME.BackColor = IIf(.ListSubItems(FixClaimItem.b09sADJNAME - 1).ReportIcon = FixClaimPic.BadData, RED_BG, WHITE_BG)
            
            txtb08sFULLNAME.Text = .SubItems(FixClaimItem.b08sFULLNAME - 1)
            mClaimItem.Main.b08sFULLNAME = txtb08sFULLNAME.Text
            txtb08sFULLNAME.BackColor = IIf(.ListSubItems(FixClaimItem.b08sFULLNAME - 1).ReportIcon = FixClaimPic.BadData, RED_BG, WHITE_BG)
            
            'Frame 03
            
            cbob10sCLAIMCITY.Text = .SubItems(FixClaimItem.b10sCLAIMCITY - 1)
            mClaimItem.Main.b10sCLAIMCITY = cbob10sCLAIMCITY.Text
            cbob10sCLAIMCITY.BackColor = IIf(.ListSubItems(FixClaimItem.b10sCLAIMCITY - 1).ReportIcon = FixClaimPic.BadData, RED_BG, WHITE_BG)
            
            txtb11sCLAIMSTATE.Text = .SubItems(FixClaimItem.b11sCLAIMSTATE - 1)
            mClaimItem.Main.b11sCLAIMSTATE = txtb11sCLAIMSTATE.Text
            txtb11sCLAIMSTATE.BackColor = IIf(.ListSubItems(FixClaimItem.b11sCLAIMSTATE - 1).ReportIcon = FixClaimPic.BadData, RED_BG, WHITE_BG)
            
            txtb04cGROSSLOSS.Text = .SubItems(FixClaimItem.b04cGROSSLOSS - 1)
            mClaimItem.Main.b04cGROSSLOSS = txtb04cGROSSLOSS.Text
            txtb04cGROSSLOSS.BackColor = IIf(.ListSubItems(FixClaimItem.b04cGROSSLOSS - 1).ReportIcon = FixClaimPic.BadData, RED_BG, WHITE_BG)
            
            txtb05cSERVICEFEE.Text = .SubItems(FixClaimItem.b05cSERVICEFEE - 1)
            mClaimItem.Main.b05cSERVICEFEE = txtb05cSERVICEFEE.Text
            txtb05cSERVICEFEE.BackColor = IIf(.ListSubItems(FixClaimItem.b05cSERVICEFEE - 1).ReportIcon = FixClaimPic.BadData, RED_BG, WHITE_BG)
            
            txtb06cEXPENSEREIM.Text = .SubItems(FixClaimItem.b06cEXPENSEREIM - 1)
            mClaimItem.Main.b06cEXPENSEREIM = txtb06cEXPENSEREIM.Text
            txtb06cEXPENSEREIM.BackColor = IIf(.ListSubItems(FixClaimItem.b06cEXPENSEREIM - 1).ReportIcon = FixClaimPic.BadData, RED_BG, WHITE_BG)
            
            txtb07cADMINFEE.Text = .SubItems(FixClaimItem.b07cADMINFEE - 1)
            mClaimItem.Main.b07cADMINFEE = txtb07cADMINFEE.Text
            txtb07cADMINFEE.BackColor = IIf(.ListSubItems(FixClaimItem.b07cADMINFEE - 1).ReportIcon = FixClaimPic.BadData, RED_BG, WHITE_BG)
            
            txtb14sPROPADDR.Text = .SubItems(FixClaimItem.b14sPROPADDR - 1)
            mClaimItem.Main.b14sPROPADDR = txtb14sPROPADDR.Text
            txtb14sPROPADDR.BackColor = IIf(.ListSubItems(FixClaimItem.b14sPROPADDR - 1).ReportIcon = FixClaimPic.BadData, RED_BG, WHITE_BG)
            
            txtb17sComments.Text = .SubItems(FixClaimItem.b17sComments - 1)
            mClaimItem.Main.b17sComments = txtb17sComments.Text
            txtb17sComments.BackColor = IIf(.ListSubItems(FixClaimItem.b17sComments - 1).ReportIcon = FixClaimPic.BadData, RED_BG, WHITE_BG)
            'If there are comments then make comments button bold
            If txtb17sComments.Text <> vbNullString Then
                cmdComments.Font.Bold = True
            Else
                cmdComments.Font.Bold = False
            End If
            
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

Private Sub lvwClaims_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    On Error GoTo EH
    If lvwClaims.SortOrder = lvwAscending Then
        lvwClaims.SortOrder = lvwDescending
    Else
        lvwClaims.SortOrder = lvwAscending
    End If
    
    'Need to See if this column has a hidden sort column
    Select Case ColumnHeader.Index
        Case FixClaimItem.b01sCATNO
            lvwClaims.SortKey = ColumnHeader.Index
        Case FixClaimItem.b02sCLAIMNO
            lvwClaims.SortKey = ColumnHeader.Index
        Case FixClaimItem.b03sIB
            lvwClaims.SortKey = ColumnHeader.Index
        Case FixClaimItem.b04cGROSSLOSS
            lvwClaims.SortKey = ColumnHeader.Index
        Case FixClaimItem.b05cSERVICEFEE
            lvwClaims.SortKey = ColumnHeader.Index
        Case FixClaimItem.b06cEXPENSEREIM
            lvwClaims.SortKey = ColumnHeader.Index
        Case FixClaimItem.b07cADMINFEE
            lvwClaims.SortKey = ColumnHeader.Index
        Case FixClaimItem.b12dtFILESRECD
            lvwClaims.SortKey = ColumnHeader.Index
        Case FixClaimItem.b13dtDATECLOSED
            lvwClaims.SortKey = ColumnHeader.Index
        Case FixClaimItem.b14sPROPADDR
            lvwClaims.SortKey = ColumnHeader.Index
        Case FixClaimItem.b15dtDateUploaded
            lvwClaims.SortKey = ColumnHeader.Index
        Case FixClaimItem.b16dtDateEntered
            lvwClaims.SortKey = ColumnHeader.Index
        Case Else
            lvwClaims.SortKey = ColumnHeader.Index - 1
    End Select
    
    lvwClaims.Sorted = True
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lvwClaims_ColumnClick"
End Sub

Private Sub lvwClaims_KeyUp(KeyCode As Integer, Shift As Integer)
    LoadEdit True
End Sub

Private Sub optSend_Click()
Dim itmX As ListItem
On Error GoTo EH
    If Not mitmX Is Nothing Then
        For Each itmX In lvwClaims.ListItems
            If itmX.Selected Then
                itmX.SmallIcon = FixClaimPic.NeedsSentToBilling
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
        lvwClaims.SetFocus
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub optSend_KeyUp"
End Sub

Private Sub optSkip_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo EH
    
    If KeyCode = vbKeyP Then
        lvwClaims.SetFocus
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub optSkip_KeyUp"
End Sub

Private Sub optSkip_Click()
Dim itmX As ListItem
On Error GoTo EH
    If Not mitmX Is Nothing Then
        For Each itmX In lvwClaims.ListItems
            If itmX.Selected Then
                itmX.SmallIcon = FixClaimPic.Skip
            End If
        Next
        EnableSend
    End If
    
    Set itmX = Nothing
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub optSkip_Click"
End Sub

Private Sub ValidateControl(pControl As Control, piClaimItem As FixClaimItem, pbMain As Boolean)
    On Error GoTo EH
    Dim sText As String
    
    If mbLoadingEdit Then
        Exit Sub
    End If
    'ucase the Text in the Control
    goUtil.utUCText pControl
    sText = pControl.Text
       
    With mitmX
        If pbMain Then
            Select Case piClaimItem
                Case FixClaimItem.b01sCATNO
                    mClaimItem.Main.b01sCATNO = sText
                    cbob01sCATNO.ToolTipText = sText
                Case FixClaimItem.b02sCLAIMNO
                    mClaimItem.Main.b02sCLAIMNO = sText
                Case FixClaimItem.b03sIB
                    mClaimItem.Main.b03sIB = sText
                Case FixClaimItem.b04cGROSSLOSS
                    If Not IsNumeric(sText) Then
                        sText = 0
                    End If
                    mClaimItem.Main.b04cGROSSLOSS = sText
                 Case FixClaimItem.b05cSERVICEFEE
                    If Not IsNumeric(sText) Then
                        sText = 0
                    End If
                    mClaimItem.Main.b05cSERVICEFEE = sText
                Case FixClaimItem.b06cEXPENSEREIM
                    If Not IsNumeric(sText) Then
                        sText = 0
                    End If
                    mClaimItem.Main.b06cEXPENSEREIM = sText
                 Case FixClaimItem.b07cADMINFEE
                    If Not IsNumeric(sText) Then
                        sText = 0
                    End If
                    mClaimItem.Main.b07cADMINFEE = sText
                Case FixClaimItem.b08sFULLNAME
                    mClaimItem.Main.b08sFULLNAME = sText
                Case FixClaimItem.b09sADJNAME
                    mClaimItem.Main.b09sADJNAME = sText
                Case FixClaimItem.b10sCLAIMCITY
                    mClaimItem.Main.b10sCLAIMCITY = sText
                Case FixClaimItem.b11sCLAIMSTATE
                    mClaimItem.Main.b11sCLAIMSTATE = sText
                Case FixClaimItem.b12dtFILESRECD
                    If Not IsDate(sText) Then
                        sText = NULL_DATE
                    End If
                    mClaimItem.Main.b12dtFILESRECD = sText
                Case FixClaimItem.b13dtDATECLOSED
                    If Not IsDate(sText) Then
                        sText = NULL_DATE
                    End If
                    mClaimItem.Main.b13dtDATECLOSED = sText
                Case FixClaimItem.b14sPROPADDR
                    mClaimItem.Main.b14sPROPADDR = sText
                Case FixClaimItem.b17sComments
                    mClaimItem.Main.b17sComments = sText
                    
            End Select
        Else
            'Nothing here yet ?
        End If
        
        moSend2Billings.ValidateClaimItem mClaimItem, mitmX, True
        
        pControl.BackColor = IIf(.ListSubItems(piClaimItem - 1).ReportIcon = FixClaimPic.BadData, RED_BG, WHITE_BG)
       
        EnableSend
    End With
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub ValidateControl"
End Sub

Private Sub EnableSend()
    On Error GoTo EH
    Dim itmX As ListItem
    Dim bFoundSend As Boolean
    
    mbItemsValidated = True
    For Each itmX In lvwClaims.ListItems
        If itmX.SmallIcon = FixClaimPic.NeedsSentToBilling Then
            bFoundSend = True
            If itmX.ListSubItems(FixClaimItem.Data - 1).ReportIcon = FixClaimPic.BadData Then
                mbItemsValidated = False
                Exit For
            End If
        End If
    Next
    If mbItemsValidated And bFoundSend Then
        cmdSendAll.Enabled = True
    Else
        cmdSendAll.Enabled = False
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub EnableSend"
End Sub

Private Sub LoadValidClaimItems()
    On Error GoTo EH
    Dim itmX As ListItem
    Dim lCount As Long
    Dim ClaimItem As V2EberlsBillings.udtClaimItem
    
    If lvwClaims.ListItems.Count = 0 Then
        Exit Sub
    End If
    
    Set mcolValidClaimItems = New Collection
    
    For lCount = 1 To lvwClaims.ListItems.Count
        Set itmX = lvwClaims.ListItems(lCount)
        If itmX.SmallIcon = FixClaimPic.NeedsSentToBilling Then
            ClaimItem.SentToBillings = False
        End If
        If itmX.SmallIcon = FixClaimPic.AlreadySentToBilling Then
            ClaimItem.SentToBillings = True
        End If
        If itmX.SmallIcon = FixClaimPic.Skip Or mbSkipAll Then
            'Issue 248 9.23.2002 Send To Send2Billings Status marked as Sent when it is not.
            'need to be sure set this flag is False
            If itmX.SmallIcon = FixClaimPic.Skip Then
                ClaimItem.SkipThisItem = True
                ClaimItem.SentToBillings = False
            Else
                ClaimItem.SkipThisItem = False
            End If
            ClaimItem.SentToBillings = False
        Else
            ClaimItem.SkipThisItem = False
        End If
        If itmX.ListSubItems(FixClaimItem.Data - 1).ReportIcon = FixClaimPic.BadData Then
            ClaimItem.ValidData = False
        Else
            ClaimItem.ValidData = True
        End If
        
        With ClaimItem.Main
            .b01sCATNO = itmX.SubItems(FixClaimItem.b01sCATNO - 1)
            .b02sCLAIMNO = itmX.SubItems(FixClaimItem.b02sCLAIMNO - 1)
            .b03sIB = itmX.SubItems(FixClaimItem.b03sIB - 1)
            .b04cGROSSLOSS = itmX.SubItems(FixClaimItem.b04cGROSSLOSS - 1)
            .b05cSERVICEFEE = itmX.SubItems(FixClaimItem.b05cSERVICEFEE - 1)
            .b06cEXPENSEREIM = itmX.SubItems(FixClaimItem.b06cEXPENSEREIM - 1)
            .b07cADMINFEE = itmX.SubItems(FixClaimItem.b07cADMINFEE - 1)
            .b08sFULLNAME = itmX.SubItems(FixClaimItem.b08sFULLNAME - 1)
            .b09sADJNAME = itmX.SubItems(FixClaimItem.b09sADJNAME - 1)
            .b10sCLAIMCITY = itmX.SubItems(FixClaimItem.b10sCLAIMCITY - 1)
            .b11sCLAIMSTATE = itmX.SubItems(FixClaimItem.b11sCLAIMSTATE - 1)
            If itmX.SubItems(FixClaimItem.b12dtFILESRECD - 1) <> vbNullString Then
                .b12dtFILESRECD = itmX.SubItems(FixClaimItem.b12dtFILESRECD - 1)
            Else
                .b12dtFILESRECD = NULL_DATE
            End If
            If itmX.SubItems(FixClaimItem.b13dtDATECLOSED - 1) <> vbNullString Then
                .b13dtDATECLOSED = itmX.SubItems(FixClaimItem.b13dtDATECLOSED - 1)
            Else
                .b13dtDATECLOSED = NULL_DATE
            End If
            .b14sPROPADDR = itmX.SubItems(FixClaimItem.b14sPROPADDR - 1)
            If itmX.SubItems(FixClaimItem.b15dtDateUploaded - 1) <> vbNullString Then
                .b15dtDateUploaded = itmX.SubItems(FixClaimItem.b15dtDateUploaded - 1)
            Else
                .b15dtDateUploaded = NULL_DATE
            End If
            If itmX.SubItems(FixClaimItem.b16dtDateEntered - 1) <> vbNullString Then
                .b16dtDateEntered = itmX.SubItems(FixClaimItem.b16dtDateEntered - 1)
            Else
                .b16dtDateEntered = NULL_DATE
            End If
            .b17sComments = itmX.SubItems(FixClaimItem.b17sComments - 1)
            
        End With

        mcolValidClaimItems.Add ClaimItem, ClaimItem.Main.b03sIB
    Next
    
    'clean up
    Set itmX = Nothing
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub LoadValidClaimItems"
End Sub

Private Sub PopulateSend2BillingsSettings()
    'Issue 224  9.23.2002 Send to Send2Billings no work with Send2Billings 2002
    On Error GoTo EH
    
    'Version Info
    cboBillingsVS.AddItem "1"
    cboBillingsVS.Text = GetSetting(goUtil.gsAppEXEName, "Version", "LatestVS", "1")
        
    'Speed Settings
    cboBillingsSpeed.AddItem "Entry Speed - FAST"
    cboBillingsSpeed.AddItem "Entry Speed - MEDIUM"
    cboBillingsSpeed.AddItem "Entry Speed - SLOW"
    
    cboBillingsSpeed.Text = GetSetting("ECS", "KEYBOARD", "SPEED", "Entry Speed - FAST")
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub PopulateSend2BillingsSettings"
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
    lvwClaims.Width = framMain.Width - lvwClaims_W
    lvwClaims.Height = framMain.Height - lvwClaims_H
    cmdUp.Top = framMain.Height - cmdUp_T
    cmdDown.Top = framMain.Height - cmdDown_T
    cmdFind.Top = framMain.Height - cmdFind_T
    cmdFindNext.Top = framMain.Height - cmdFindNext_T
    cmdViewFailed.Top = framMain.Height - cmdViewFailed_T
    chkBrowseBadData.Top = framMain.Height - chkBrowseBadData_T
    chkWaitForUserOK.Top = framMain.Height - chkWaitForUserOK_T
    chkShowGrid.Top = framMain.Height - chkShowGrid_T
    
    'Fram Status
    framSendStatus.Top = lH - framSendStatus_T
    
    'Fram Commands
    framCommands.Top = lH - framCommands_T
    framCommands.Width = lW - framCommands_W
    cmdPrintList.Left = framCommands.Width - cmdPrintList_L
    cmdExit.Left = framCommands.Width - cmdExit_L
    
    'Fram 1 - 3
    'Tops
    fram01.Top = lH - fram01_T
    fram02.Top = lH - fram02_T
    fram03.Top = lH - fram03_T
    
    'Widths
    lNewWidth = lW - fram02_W
    If fram02.Width < lNewWidth Then
        lNewWidth = lNewWidth - fram02.Width
        fram02.Width = fram02.Width + (lNewWidth / 2)
    Else
        fram02.Width = lNewWidth
    End If
    
    fram03.Left = fram02_L + fram02.Width + fram03_L
    fram03.Width = lW - fram03.Left - fram03_W
    
    'Fram 02
    txtb03sIB.Width = fram02.Width - txtb03sIB_W
    txtb02sCLAIMNO.Width = fram02.Width - txtb02sCLAIMNO_W
    cbob09sADJNAME.Width = fram02.Width - cbob09sADJNAME_W
    txtb08sFULLNAME.Width = fram02.Width - txtb08sFULLNAME_W
    
    'Fram03
    cbob10sCLAIMCITY.Width = fram03.Width - cbob10sCLAIMCITY_W
    txtb11sCLAIMSTATE.Left = cbob10sCLAIMCITY.Left + cbob10sCLAIMCITY.Width + txtb11sCLAIMSTATE_L
    txtb04cGROSSLOSS.Width = fram03.Width - txtb04cGROSSLOSS_W
    txtb05cSERVICEFEE.Left = txtb04cGROSSLOSS.Left + txtb04cGROSSLOSS.Width + txtb05cSERVICEFEE_L
    txtb06cEXPENSEREIM.Width = fram03.Width - txtb06cEXPENSEREIM_W
    txtb07cADMINFEE.Left = txtb06cEXPENSEREIM.Left + txtb06cEXPENSEREIM.Width + txtb07cADMINFEE_L
    txtb14sPROPADDR.Width = fram03.Width - txtb14sPROPADDR_W
    
    
    VisibleFrames True
    
    mbResize = False
    Exit Sub
EH:
    Err.Clear
    Resume Next
End Sub


Private Sub txtb02sCLAIMNO_Change()
    ValidateControl txtb02sCLAIMNO, b02sCLAIMNO, True
End Sub

Private Sub txtb02sCLAIMNO_GotFocus()
    goUtil.utSelText txtb02sCLAIMNO
End Sub

Private Sub txtb02sCLAIMNO_LostFocus()
    
    LoadEdit
End Sub

Private Sub txtb03sIB_Change()
    ValidateControl txtb03sIB, b03sIB, True
End Sub

Private Sub txtb03sIB_GotFocus()
    goUtil.utSelText txtb03sIB
End Sub

Private Sub txtb03sIB_LostFocus()
    
    LoadEdit
End Sub

Private Sub txtb04cGROSSLOSS_Change()
    ValidateControl txtb04cGROSSLOSS, b04cGROSSLOSS, True
End Sub

Private Sub txtb04cGROSSLOSS_GotFocus()
    goUtil.utSelText txtb04cGROSSLOSS
End Sub

Private Sub txtb04cGROSSLOSS_LostFocus()
    
    LoadEdit
End Sub

Private Sub txtb05cSERVICEFEE_Change()
    ValidateControl txtb05cSERVICEFEE, b05cSERVICEFEE, True
End Sub

Private Sub txtb05cSERVICEFEE_GotFocus()
    goUtil.utSelText txtb05cSERVICEFEE
End Sub

Private Sub txtb05cSERVICEFEE_LostFocus()
    
    LoadEdit
End Sub

Private Sub txtb06cEXPENSEREIM_Change()
    ValidateControl txtb06cEXPENSEREIM, b06cEXPENSEREIM, True
End Sub

Private Sub txtb06cEXPENSEREIM_GotFocus()
    goUtil.utSelText txtb06cEXPENSEREIM
End Sub

Private Sub txtb06cEXPENSEREIM_LostFocus()
    
    LoadEdit
End Sub

Private Sub txtb07cADMINFEE_Change()
    ValidateControl txtb07cADMINFEE, b07cADMINFEE, True
End Sub

Private Sub txtb07cADMINFEE_GotFocus()
    goUtil.utSelText txtb07cADMINFEE
End Sub

Private Sub txtb07cADMINFEE_LostFocus()
    
    LoadEdit
End Sub

Private Sub txtb08sFULLNAME_Change()
    ValidateControl txtb08sFULLNAME, b08sFULLNAME, True
End Sub

Private Sub txtb08sFULLNAME_GotFocus()
    goUtil.utSelText txtb08sFULLNAME
End Sub

Private Sub txtb08sFULLNAME_LostFocus()
    
    LoadEdit
End Sub

Private Sub txtb11sCLAIMSTATE_Change()
    ValidateControl txtb11sCLAIMSTATE, b11sCLAIMSTATE, True
End Sub

Private Sub txtb11sCLAIMSTATE_GotFocus()
    goUtil.utSelText txtb11sCLAIMSTATE
End Sub

Private Sub txtb11sCLAIMSTATE_LostFocus()
    LoadEdit
End Sub

Private Sub txtb12dtFILESRECD_Change()
    ValidateControl txtb12dtFILESRECD, b12dtFILESRECD, True
End Sub

Private Sub txtb12dtFILESRECD_GotFocus()
    goUtil.utSelText txtb12dtFILESRECD
End Sub

Private Sub txtb12dtFILESRECD_LostFocus()
    
    LoadEdit
End Sub

Private Sub txtb13dtDATECLOSED_Change()
    ValidateControl txtb13dtDATECLOSED, b13dtDATECLOSED, True
End Sub

Private Sub txtb13dtDATECLOSED_GotFocus()
    goUtil.utSelText txtb13dtDATECLOSED
End Sub

Private Sub txtb13dtDATECLOSED_LostFocus()
     
    LoadEdit
End Sub

Private Sub txtb14sPROPADDR_Change()
    ValidateControl txtb14sPROPADDR, b14sPROPADDR, True
End Sub

Private Sub txtb14sPROPADDR_GotFocus()
    goUtil.utSelText txtb14sPROPADDR
End Sub

Private Sub txtb14sPROPADDR_LostFocus()
    LoadEdit
End Sub

Private Sub txtb17sComments_Change()
    ValidateControl txtb17sComments, b17sComments, True
End Sub

Private Sub txtb17sComments_GotFocus()
    goUtil.utSelText txtb17sComments
End Sub

Private Sub txtb17sComments_LostFocus()
    LoadEdit
End Sub

Private Sub txtEntryDate_Change()
    ValidateControl txtEntryDate, b16dtDateEntered, True
End Sub

Private Sub txtEntryDate_GotFocus()
    goUtil.utSelText txtEntryDate
End Sub

Private Sub txtEntryDate_LostFocus()
    LoadEdit
End Sub

Public Sub VisibleFrames(pbVisible As Boolean)
    On Error GoTo EH
    framMain.Visible = pbVisible
    framSendStatus.Visible = pbVisible
    framCommands.Visible = pbVisible
    fram01.Visible = pbVisible
    fram02.Visible = pbVisible
    fram03.Visible = pbVisible
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub VisibleFrames"
End Sub

