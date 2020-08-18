VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmIndemnity 
   AutoRedraw      =   -1  'True
   Caption         =   " "
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
   LockControls    =   -1  'True
   ScaleHeight     =   6555
   ScaleWidth      =   11820
   StartUpPosition =   3  'Windows Default
   Tag             =   "Indemnity"
   Begin VB.Frame framPayReqs 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   240
      TabIndex        =   49
      Top             =   480
      Width           =   11295
      Begin VB.CommandButton cmdFindNextPayReqs 
         Caption         =   "Find &Next"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   53
         Top             =   180
         Width           =   1100
      End
      Begin VB.CommandButton cmdFindPayReqs 
         Caption         =   "&Find"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   52
         Top             =   180
         Width           =   1100
      End
      Begin VB.CommandButton cmdSelAllPayReqs 
         Caption         =   "&Select A&ll"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   51
         Top             =   180
         Width           =   1100
      End
      Begin VB.CommandButton cmdEditPayReqs 
         Caption         =   "&Edit"
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
         Left            =   120
         TabIndex        =   50
         Top             =   180
         Width           =   975
      End
      Begin VB.Frame framPayReqsMaint 
         Caption         =   "Payment Request Maintenance"
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
         TabIndex        =   55
         Top             =   3840
         Width           =   10995
         Begin VB.CheckBox chkIncludeMortgagee 
            Caption         =   "Include Mortgagee on Draft"
            Height          =   255
            Left            =   2280
            TabIndex        =   69
            Top             =   360
            Visible         =   0   'False
            Width           =   2775
         End
         Begin VB.CommandButton cmdRefreshPayReqs 
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
            TabIndex        =   56
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmdDelPayReqs 
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
            Left            =   9360
            TabIndex        =   59
            TabStop         =   0   'False
            Top             =   240
            Width           =   1455
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
            Index           =   1
            Left            =   8160
            Style           =   1  'Graphical
            TabIndex        =   58
            Top             =   240
            Visible         =   0   'False
            Width           =   1100
         End
         Begin VB.CommandButton cmdAddPayReqs 
            Caption         =   "&Add"
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
            Left            =   1200
            TabIndex        =   57
            Top             =   240
            Visible         =   0   'False
            Width           =   975
         End
      End
      Begin MSComctlLib.ImageList imgPayReqsStatus 
         Left            =   10080
         Top             =   720
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
               Picture         =   "frmIndemnity.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmIndemnity.frx":015A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmIndemnity.frx":0546
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lvwPayReqs 
         Height          =   3135
         Left            =   120
         TabIndex        =   54
         Tag             =   "Enable"
         ToolTipText     =   "Right Click for Menu"
         Top             =   600
         Width           =   10995
         _ExtentX        =   19394
         _ExtentY        =   5530
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         OLEDragMode     =   1
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "imgPayReqsStatus"
         ColHdrIcons     =   "imgPayReqsStatus"
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
   Begin VB.Frame framIndemnity 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   240
      TabIndex        =   10
      Top             =   480
      Width           =   11295
      Begin VB.CommandButton cmdFindNextlIndemnity 
         Caption         =   "Find &Next"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         TabIndex        =   15
         Top             =   200
         Width           =   1100
      End
      Begin VB.CommandButton cmdFindlIndemnity 
         Caption         =   "&Find"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         TabIndex        =   14
         Top             =   200
         Width           =   1100
      End
      Begin VB.CommandButton cmdSelAllIndemnity 
         Caption         =   "&Select A&ll"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   13
         Top             =   200
         Width           =   1100
      End
      Begin VB.CommandButton cmdEditIndemnity 
         Caption         =   "&Edit"
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
         Left            =   1080
         TabIndex        =   12
         Top             =   200
         Width           =   975
      End
      Begin VB.CommandButton cmdAddIndemnity 
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
         Left            =   120
         TabIndex        =   11
         Top             =   200
         Width           =   975
      End
      Begin VB.Frame framIndemTotals 
         Caption         =   "Indemnity Totals (Does not include previous payment amounts.)"
         Height          =   1335
         Left            =   120
         TabIndex        =   17
         Top             =   2520
         Width           =   10995
         Begin VB.TextBox txtLessExcessLimitsAbsorbDed 
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
            Left            =   8160
            Locked          =   -1  'True
            TabIndex        =   33
            ToolTipText     =   "Amount of Excess Limits Allowed to Absorb Deductible"
            Top             =   480
            Width           =   1335
         End
         Begin VB.TextBox txtDeductibleValue 
            Alignment       =   1  'Right Justify
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
            Left            =   8160
            Locked          =   -1  'True
            TabIndex        =   29
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox txtNetActualCashValueClaimValue 
            Alignment       =   1  'Right Justify
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
            Left            =   6360
            Locked          =   -1  'True
            TabIndex        =   37
            Top             =   960
            Width           =   1455
         End
         Begin VB.TextBox txtLessMiscValue 
            Alignment       =   1  'Right Justify
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
            Left            =   6360
            Locked          =   -1  'True
            TabIndex        =   35
            Top             =   720
            Width           =   1455
         End
         Begin VB.TextBox txtLessExcessLimitsValue 
            Alignment       =   1  'Right Justify
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
            Left            =   6360
            Locked          =   -1  'True
            TabIndex        =   31
            Top             =   480
            Width           =   1455
         End
         Begin VB.TextBox txtAppliedDeductibleValue 
            Alignment       =   1  'Right Justify
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
            Left            =   6360
            Locked          =   -1  'True
            TabIndex        =   27
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox txtPayReqsValue 
            Alignment       =   1  'Right Justify
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
            Locked          =   -1  'True
            TabIndex        =   39
            Top             =   960
            Width           =   1335
         End
         Begin VB.TextBox txtACVLossValue 
            Alignment       =   1  'Right Justify
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
            Left            =   3000
            Locked          =   -1  'True
            TabIndex        =   25
            Top             =   960
            Width           =   1455
         End
         Begin VB.TextBox txtNonRecovDeprValue 
            Alignment       =   1  'Right Justify
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
            Left            =   3000
            Locked          =   -1  'True
            TabIndex        =   23
            Top             =   720
            Width           =   1455
         End
         Begin VB.TextBox txtRecoverableDepreciationValue 
            Alignment       =   1  'Right Justify
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
            Left            =   3000
            Locked          =   -1  'True
            TabIndex        =   21
            Top             =   480
            Width           =   1455
         End
         Begin VB.TextBox txtFullCostOfRepairValue 
            Alignment       =   1  'Right Justify
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
            Left            =   3000
            Locked          =   -1  'True
            TabIndex        =   19
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label lblXsAbsorb 
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
            Left            =   7920
            TabIndex        =   32
            Top             =   480
            Width           =   135
         End
         Begin VB.Label lblOf 
            Alignment       =   2  'Center
            Caption         =   "of"
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
            Left            =   7920
            TabIndex        =   28
            Top             =   240
            Width           =   135
         End
         Begin VB.Label lblAppliedDeductible 
            Alignment       =   1  'Right Justify
            Caption         =   "Applied Deductible:"
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
            Left            =   4560
            TabIndex        =   26
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label lblLessMisc 
            Alignment       =   1  'Right Justify
            Caption         =   "Less Miscellaneous:"
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
            Left            =   4320
            TabIndex        =   34
            Top             =   720
            Width           =   1935
         End
         Begin VB.Label lblPayReqs 
            Alignment       =   1  'Right Justify
            Caption         =   "Payment Request:"
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
            Left            =   7800
            TabIndex        =   38
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label lblLessExcessLimits 
            Alignment       =   1  'Right Justify
            Caption         =   "Less Excess Limits:"
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
            Left            =   4560
            TabIndex        =   30
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label lblACVLoss 
            Alignment       =   1  'Right Justify
            Caption         =   "Actual Cash Value:"
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
            Left            =   1200
            TabIndex        =   24
            Top             =   960
            Width           =   1695
         End
         Begin VB.Label lblNonRecovDepr 
            Alignment       =   1  'Right Justify
            Caption         =   "Non-Recoverable Depreciation:"
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
            Left            =   120
            TabIndex        =   22
            Top             =   720
            Width           =   2775
         End
         Begin VB.Label lblRecoverableDepreciation 
            Alignment       =   1  'Right Justify
            Caption         =   "Recoverable Depreciation:"
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
            Left            =   120
            TabIndex        =   20
            Top             =   480
            Width           =   2775
         End
         Begin VB.Label lblFullCostOfRepair 
            Alignment       =   1  'Right Justify
            Caption         =   "Full Cost of Repair/Replacement:"
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
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Width           =   2775
         End
         Begin VB.Label lblNetActualCashValueClaim 
            Alignment       =   1  'Right Justify
            Caption         =   "Net ACVC:"
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
            TabIndex        =   36
            Top             =   960
            Width           =   2655
         End
         Begin VB.Label lblAmountWarning 
            ForeColor       =   &H008080FF&
            Height          =   735
            Left            =   9480
            TabIndex        =   68
            Top             =   240
            Width           =   1335
         End
      End
      Begin MSComctlLib.ImageList imgIndemStatus 
         Left            =   10080
         Top             =   720
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
               Picture         =   "frmIndemnity.frx":0998
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmIndemnity.frx":0AF2
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmIndemnity.frx":0EDE
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Frame framIndemnityMaint 
         Caption         =   "Indemnity Maintenance"
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
         TabIndex        =   40
         Top             =   3840
         Width           =   10995
         Begin VB.CommandButton cmdPrintIndem 
            Caption         =   "&Print All"
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
            TabIndex        =   42
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
            Index           =   0
            Left            =   8160
            Style           =   1  'Graphical
            TabIndex        =   43
            Top             =   240
            Visible         =   0   'False
            Width           =   1100
         End
         Begin VB.CommandButton cmdDelIndemnity 
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
            Left            =   9360
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   240
            Width           =   1455
         End
         Begin VB.CommandButton cmdRefreshIndemnity 
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
            TabIndex        =   41
            Top             =   240
            Width           =   975
         End
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
         Height          =   1455
         Left            =   1320
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   720
         Visible         =   0   'False
         Width           =   3855
      End
      Begin MSComctlLib.ListView lvwIndemnity 
         Height          =   1935
         Left            =   120
         TabIndex        =   16
         Tag             =   "Enable"
         ToolTipText     =   "Right Click for Menu"
         Top             =   600
         Width           =   10995
         _ExtentX        =   19394
         _ExtentY        =   3413
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         OLEDragMode     =   1
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "imgIndemStatus"
         ColHdrIcons     =   "imgIndemStatus"
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
   Begin VB.Frame framDeductible 
      Height          =   4695
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   11295
      Begin MSComctlLib.ImageList imgPolicyLimits 
         Left            =   8880
         Top             =   960
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
               Picture         =   "frmIndemnity.frx":1330
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmIndemnity.frx":1782
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmIndemnity.frx":1B6E
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdDelDed 
         Caption         =   "&DELETE"
         Height          =   375
         Left            =   9720
         TabIndex        =   9
         Top             =   1920
         Width           =   1455
      End
      Begin VB.CommandButton CmdReNumberSort 
         Caption         =   "&Save Sort Order"
         Height          =   975
         Left            =   10320
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Renumber and Save"
         Top             =   840
         Width           =   855
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
         Left            =   9720
         Picture         =   "frmIndemnity.frx":1E7D
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Move Selected Item UP"
         Top             =   855
         Width           =   480
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
         Left            =   9720
         Picture         =   "frmIndemnity.frx":22BF
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Move Selected Item DOWN"
         Top             =   1335
         Width           =   480
      End
      Begin VB.Frame framDedMaint 
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
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   11295
         Begin VB.ComboBox cboAddPLAppClassTypeID 
            Height          =   360
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   240
            Width           =   8415
         End
         Begin VB.CommandButton cmdAddPLAppClassTypeID 
            Caption         =   "&ADD"
            Height          =   375
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   975
         End
      End
      Begin MSComctlLib.ListView lvwDed 
         Height          =   3735
         Left            =   120
         TabIndex        =   5
         Tag             =   "Enable"
         Top             =   840
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   6588
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         OLEDragMode     =   1
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "imgListPhotos"
         SmallIcons      =   "imgPolicyLimits"
         ColHdrIcons     =   "imgPolicyLimits"
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
   Begin MSComctlLib.TabStrip TSIndem 
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
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Indemnity Entry"
            Object.Tag             =   "framIndemnity"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Payment Re&quests"
            Object.Tag             =   "framPayReqs"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Appl&y Deductible Order"
            Object.Tag             =   "framDeductible"
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
   Begin VB.Frame framAssocRTChecksID 
      Caption         =   "Associate Payment Request to selected Indemnity Items:"
      Height          =   1215
      Left            =   120
      TabIndex        =   45
      Top             =   5280
      Width           =   8055
      Begin VB.CommandButton cmdAddPayReqAssIndem 
         Caption         =   "Add Pa&y Request"
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
         Left            =   5280
         TabIndex        =   47
         Top             =   240
         Width           =   2655
      End
      Begin VB.ComboBox cboAssocRTChecksID 
         Height          =   360
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Top             =   720
         Width           =   5055
      End
      Begin VB.CommandButton cmdAssocRTChecksID 
         Caption         =   "Associate &Pay Request"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   713
         Width           =   2655
      End
   End
   Begin VB.Frame framAssocBillingID 
      Caption         =   "Associate IB (Internal Billing) to selected Payment Request(s):"
      Height          =   1215
      Left            =   120
      TabIndex        =   60
      Top             =   5280
      Width           =   8055
      Begin VB.CommandButton cmdAssocBillingID 
         Caption         =   "Associate I&B"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   713
         Width           =   2655
      End
      Begin VB.ComboBox cboAssocBillingID 
         Height          =   360
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   61
         Top             =   720
         Width           =   5055
      End
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
      Left            =   8280
      TabIndex        =   63
      Top             =   5280
      Width           =   3375
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
         Left            =   120
         MaskColor       =   &H00000000&
         Picture         =   "frmIndemnity.frx":2701
         Style           =   1  'Graphical
         TabIndex        =   64
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
         Left            =   2280
         MaskColor       =   &H00000000&
         Picture         =   "frmIndemnity.frx":2B43
         Style           =   1  'Graphical
         TabIndex        =   66
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
         Picture         =   "frmIndemnity.frx":2E4D
         Style           =   1  'Graphical
         TabIndex        =   65
         ToolTipText     =   "Exit"
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Menu PopUpmnuPayReqs 
      Caption         =   "PopUpPayReqs"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuEditPayReqs 
         Caption         =   "&Edit"
      End
      Begin VB.Menu mnuDeletePayReqs 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuSelectAllPayReqs 
         Caption         =   "&Select All"
      End
   End
   Begin VB.Menu PopUpMnuIndem 
      Caption         =   "PopUpIndem"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuEditIndem 
         Caption         =   "&Edit"
      End
      Begin VB.Menu mnuDeleteIndem 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuSelectAllIndem 
         Caption         =   "&Select All"
      End
   End
End
Attribute VB_Name = "frmIndemnity"
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
Private mbLoading As Boolean 'Loading Form
Private mbLoadingMe As Boolean 'Just loading data
Private msFindText As String
Private mlLastFindIndex As Long
Private mArv As V2ARViewer.clsARViewer
Private moForm As Form
Private mcolSelIndemID As Collection

Public Property Get colSelIndemID() As Collection
    Set colSelIndemID = mcolSelIndemID
End Property

Public Property Let CurrentTextBox(poTextBox As Object)
    On Error GoTo EH
    Set moCurrentTextBox = poTextBox
    If moCurrentTextBox Is Nothing Then
'        cmdSpelling.Enabled = False
    Else
'        cmdSpelling.Enabled = True
    End If
    Exit Property
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Property Let CurrentTextBox"
End Property
Public Property Set CurrentTextBox(poTextBox As Object)
    On Error GoTo EH
    Set moCurrentTextBox = poTextBox
    If moCurrentTextBox Is Nothing Then
'        cmdSpelling.Enabled = False
    Else
'        cmdSpelling.Enabled = True
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



Private Sub cboAssocRTChecksID_Click()
    On Error GoTo EH
    
    If cboAssocRTChecksID.ListIndex > -1 Then
        cmdAssocRTChecksID.Enabled = True
    Else
        cmdAssocRTChecksID.Enabled = False
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cboAssocRTChecksID_Click"
End Sub


Private Sub chkHideDeleted_Click(Index As Integer)
    On Error GoTo EH
    Dim bHideDeleted As Boolean
    Static bChangeMe As Boolean
    
    If bChangeMe Then
        Exit Sub
    End If
    bChangeMe = True
    
    If Index = 1 Then
        If chkHideDeleted(1).Value = vbChecked Then
            chkHideDeleted(0).Caption = "&Hide Deleted"
            chkHideDeleted(0).Value = vbChecked
            chkHideDeleted(1).Caption = "&Hide Deleted"
            bHideDeleted = True
        Else
            chkHideDeleted(0).Caption = "Sho&w Deleted"
            chkHideDeleted(0).Value = vbUnchecked
            chkHideDeleted(1).Caption = "Sho&w Deleted"
            bHideDeleted = False
        End If
    Else
        If chkHideDeleted(0).Value = vbChecked Then
            chkHideDeleted(0).Caption = "&Hide Deleted"
            chkHideDeleted(1).Caption = "&Hide Deleted"
            chkHideDeleted(1).Value = vbChecked
            bHideDeleted = True
        Else
            chkHideDeleted(0).Caption = "Sho&w Deleted"
            chkHideDeleted(1).Caption = "Sho&w Deleted"
            chkHideDeleted(1).Value = vbUnchecked
            bHideDeleted = False
        End If
    End If
    
    SaveSetting App.EXEName, "GENERAL", "HIDE_DELETED", bHideDeleted
    If Not mbLoading Then
        LoadMe
    End If
    bChangeMe = False
    Exit Sub
EH:
    bChangeMe = False
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub chkHideDeleted_Click"
End Sub



Private Sub cmdAddIndemnity_Click()
    On Error GoTo EH
    Dim MyIndem As GuiIndemItem
    Dim sSQL As String
    Dim RS As ADODB.Recordset
    Dim oConn As ADODB.Connection
    Dim itmX As ListItem
    Dim sID As String
    Dim MyID As String
    
    cmdAddIndemnity.Enabled = False
    
    Set oConn = New ADODB.Connection
    Set RS = New ADODB.Recordset
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    'Get some info from Assignments Record
    Set RS = mfrmClaim.adoRSAssignments
    
    If RS.RecordCount = 1 Then
        RS.MoveFirst
        With MyIndem
            .ACVClaim = "0.00"
            .ACVLessExcessLimits = "0.00"
            .SpecialLimits = "0.00"
            .ExcessLimits = "0.00"
            .Miscellaneous = "0.00"
            .MiscDescription = vbNullString
            .IsAddAmountOfInsurance = False
            .ExcessAbsorbsDeductible = True
            .AppliedDeductible = "0.00"
            .NonRecoverableDep = "0.00"
            .RecoverableDep = "0.00"
            .ReplacementCost = "0.00"
            .TypeOfLossID = goUtil.IsNullIsVbNullString(RS.Fields("TypeOfLossID"))
            .ClassOfLossID = "0"
            .Description = vbNullString
            .IsPreviousPayment = False
            .PPayDatePaid = "Null"
            .PPayAmountPaid = "0.00"
            .PPayCheckNumber = "Null"
            .IsDeleted = False
            .DownLoadMe = False
            .UpLoadMe = True
            .AdminComments = vbNullString
            .DateLastUpdated = Format(Now(), "MM/DD/YYYY HH:MM:SS")
            .UpdateByUserID = goUtil.gsCurUsersID
            If AddIndemnityItem(MyIndem, sID) Then
                mfrmClaim.RefreshMe
                'Select the newly added Pay req
                For Each itmX In lvwIndemnity.ListItems
                    itmX.Selected = False
                Next
                For Each itmX In lvwIndemnity.ListItems
                    MyID = itmX.SubItems(GuiIndemListView.ID - 1)
                    If MyID = sID Then
                        itmX.Selected = True
                        Exit For
                    End If
                Next
                
                EditIndemnity
                
            End If
        End With
    End If
    
     cmdAddIndemnity.Enabled = True
    
    'cleanup
    Set RS = Nothing
    Set oConn = Nothing
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdAddIndemnity_Click"
End Sub

Private Sub cmdAddPayReqs_Click()
    On Error GoTo EH
    
    cmdAddPayReqs.Enabled = False
    
    If AddPayReqs Then
        cmdAddPayReqs.Enabled = True
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdAddPayReqs_Click"
End Sub

Private Function AddPayReqs(Optional pbAssociateFromIndem As Boolean = False) As Boolean
    On Error GoTo EH
    Dim MyPayReq As GuiPayReqsItem
    Dim sSQL As String
    Dim RS As ADODB.Recordset
    Dim oConn As ADODB.Connection
    Dim sMaxCheckNum As String
    Dim lMaxCheckNum As Long
    Dim lNextCheckNum As Long
    Dim itmX As ListItem
    Dim sID As String
    Dim MyID As String
    
    
    Set oConn = New ADODB.Connection
    Set RS = New ADODB.Recordset
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    
    'Need to get the Max of CheckNums for this Assignment and Add One
    sSQL = "SELECT Max(RTC.[CheckNum]) As MaxCheckNum "
    sSQL = sSQL & "FROM RTChecks RTC "
    sSQL = sSQL & "WHERE RTC.[IDAssignments] = " & msAssignmentsID & " "
    
    RS.CursorLocation = adUseClient
    RS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set RS.ActiveConnection = Nothing
    
    If RS.RecordCount = 1 Then
        sMaxCheckNum = goUtil.IsNullIsVbNullString(RS.Fields("MaxCheckNum"))
        If sMaxCheckNum = vbNullString Then
            lMaxCheckNum = 0
        Else
            lMaxCheckNum = CLng(sMaxCheckNum)
        End If
        lNextCheckNum = lMaxCheckNum + 1
    End If
    
    'Get some info from Assignments Record
    Set RS = Nothing
    Set RS = mfrmClaim.adoRSAssignments
    
    If RS.RecordCount = 1 Then
        RS.MoveFirst
        With MyPayReq
            .CheckNum = lNextCheckNum
            .RT42_ClassOfLossID = "null"
            .RT43_TypeOfLossID = goUtil.IsNullIsVbNullString(RS.Fields("TypeOfLossID"))
            If .RT43_TypeOfLossID = vbNullString Then
                .RT43_TypeOfLossID = "null"
            End If
            .RT50_sInsuredPayeeName = goUtil.IsNullIsVbNullString(RS.Fields("Insured"))
            If chkIncludeMortgagee.Value = vbChecked Then
                .RT51_sPayeeNames = goUtil.IsNullIsVbNullString(RS.Fields("MortgageeName"))
            Else
                .RT51_sPayeeNames = vbNullString
            End If
            .RT52_sAddress = goUtil.IsNullIsVbNullString(RS.Fields("MAStreet")) & "    "
            .RT52_sAddress = .RT52_sAddress & goUtil.IsNullIsVbNullString(RS.Fields("MACity")) & ", "
            .RT52_sAddress = .RT52_sAddress & goUtil.IsNullIsVbNullString(RS.Fields("MAState")) & " "
            .RT52_sAddress = .RT52_sAddress & Format(goUtil.IsNullIsVbNullString(RS.Fields("MAZIP")), "00000")
'            & "-"
'            .RT52_sAddress = .RT52_sAddress & Format(goUtil.IsNullIsVbNullString(RS.Fields("MAZIP4")), "0000") & " "
            .RT53_cAmountOfCheck = "0"
            .AppliedDeductible = "0"
            .RT54_CompanyCatSpecID = goUtil.IsNullIsVbNullString(RS.Fields("ClientCompanyCatSpecID"))
            .tempCHeckName = vbNullString
            If lNextCheckNum = 1 Then
                .PrintOnIB = "False" 'Turned off at this time!! 6/4/2005
            Else
                .PrintOnIB = "False"
            End If
            .IsDeleted = "False"
            .DownLoadMe = "False"
            .UpLoadMe = "True"
            .AdminComments = vbNullString
            .DateLastUpdated = Format(Now(), "MM/DD/YYYY HH:MM:SS")
            .UpdateByUserID = goUtil.gsCurUsersID
            If AddPayReqItem(MyPayReq, sID) Then
                mfrmClaim.RefreshMe
                'Select the newly added Pay req
                For Each itmX In lvwPayReqs.ListItems
                    itmX.Selected = False
                Next
                For Each itmX In lvwPayReqs.ListItems
                    MyID = itmX.SubItems(GuiPayReqsListView.ID - 1)
                    If MyID = sID Then
                        itmX.Selected = True
                        Exit For
                    End If
                Next
                
                AddPayReqs = EditPayReqs(pbAssociateFromIndem)
                
            End If
        End With
    End If
    
    cmdAddPayReqs.Enabled = True
    
    'cleanup
    Set RS = Nothing
    Set oConn = Nothing
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Function AddPayReqs"
End Function

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
    
    cmdAddPLAppClassTypeID.Enabled = False
    If ADDDedApply() Then
        cmdAddPLAppClassTypeID.Enabled = True
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdAddPLAppClassTypeID_Click"
End Sub

Private Sub cmdAssocBillingID_Click()
    On Error GoTo EH
    Dim sMess As String
    Dim sIBDesc As String
    Dim sTitle As String
    Dim bItemSelected As Boolean
    Dim itmX As MSComctlLib.ListItem
    Dim sRTChecksID As String
    Dim vRTChecksID As Variant
    Dim colRTChecksID As Collection
    
    
    If lvwPayReqs.ListItems.Count > 0 Then
        
        If cboAssocBillingID.ListIndex = -1 Then
            sMess = "You must select an IB from the Drop down List!"
            MsgBox sMess, vbExclamation + vbOKOnly, "Select an IB."
            cmdAssocBillingID.Enabled = False
            Exit Sub
        End If
        
        For Each itmX In lvwPayReqs.ListItems
            If itmX.Selected Then
                bItemSelected = True
                Exit For
            End If
        Next
        
        'See if there is a selected item
        If Not bItemSelected Then
            sMess = "You must select at least one Payment Request item from the View!"
            MsgBox sMess, vbExclamation + vbOKOnly, "Select at least one Payment Request item."
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
        
        lvwPayReqs.Visible = False
        
        sMess = vbNullString
        
        Set colRTChecksID = New Collection
        For Each itmX In lvwPayReqs.ListItems
            If itmX.Selected Then
                'Do not allow association to of Checks that are of Other Class OF Loss Code
                If StrComp(mfrmClaim.GetClassOfLossCode(itmX.SubItems(GuiPayReqsListView.RT42_ClassOfLossID - 1)), "Other", vbTextCompare) <> 0 Then
                    colRTChecksID.Add itmX.SubItems(GuiPayReqsListView.ID - 1), itmX.SubItems(GuiPayReqsListView.ID - 1)
                Else
                    sMess = sMess & itmX.SubItems(GuiPayReqsListView.ClassOfLoss - 1)
                    sMess = sMess & " (" & itmX.SubItems(GuiPayReqsListView.ClassOfLossCode - 1) & ") - "
                    sMess = sMess & itmX.SubItems(GuiPayReqsListView.TypeOfLoss - 1)
                    sMess = sMess & " (" & itmX.SubItems(GuiPayReqsListView.TypeOfLossCode - 1) & ") - "
                    sMess = sMess & " [" & itmX.SubItems(GuiPayReqsListView.RT53_cAmountOfCheck - 1) & "]"
                    sMess = sMess & "Could not update the above item... " & vbCrLf & vbCrLf
                    sMess = sMess & "The Class Of Loss is not valid for this operation."
                End If
            End If
        Next
        For Each vRTChecksID In colRTChecksID
            sRTChecksID = vRTChecksID
            If Not AssocRTChecksItemToBillingID(sRTChecksID) Then
                Exit Sub
            End If
        Next
    End If
    
    RefreshPayReqs
    
    lvwPayReqs.Visible = True
    cmdAssocBillingID.Enabled = True
    
    If sMess <> vbNullString Then
        MsgBox sMess, vbExclamation + vbOKOnly, "Associate IB"
    End If
    
    Set itmX = Nothing
    Set colRTChecksID = Nothing
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdAssocBillingID_Click"
End Sub

Public Function AssocRTChecksItemToBillingID(psID As String) As Boolean
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
    
    sSQL = "UPDATE RTChecks SET "
    sSQL = sSQL & "[BillingCountID] = " & sIDBillingCount & ", "
    sSQL = sSQL & "[IDBillingCount] = " & sIDBillingCount & ", "
    sSQL = sSQL & "[UpLoadMe] = True, "
    sSQL = sSQL & "[DateLastUpdated] = #" & Format(Now(), "MM/DD/YYYY HH:MM:SS") & "#, "
    sSQL = sSQL & "[UpdateByUserID] = " & goUtil.gsCurUsersID & " "
    sSQL = sSQL & "WHERE [ID] = " & psID & " "
    sSQL = sSQL & "AND IDAssignments = " & msAssignmentsID & " "

    oConn.Execute sSQL
    
    AssocRTChecksItemToBillingID = True
    
    'clean up
    Set oConn = Nothing
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function AssocRTChecksItemToBillingID"
End Function


Private Sub cmdAssocRTChecksID_Click()
    On Error GoTo EH
    Dim sMess As String
    Dim sPayReqDesc As String
    Dim sIndemDesc As String
    Dim sTitle As String
    Dim bItemSelected As Boolean
    Dim itmX As MSComctlLib.ListItem
    Dim sIndemID As String
    Dim vIndemID As Variant
    Dim colIndemID As Collection
    Dim sFlagText As String
    Dim sRTIndemnityID As String
    Dim sRTChecksID As String
    Dim sThisRTChecksID As String
    
    
    If lvwIndemnity.ListItems.Count > 0 Then
        
        If cboAssocRTChecksID.ListIndex = -1 Then
            sMess = "You must select a Payment Request from the Drop down List!"
            MsgBox sMess, vbExclamation + vbOKOnly, "Select a Payment Request."
            cmdAssocRTChecksID.Enabled = False
            Exit Sub
        End If
        
        For Each itmX In lvwIndemnity.ListItems
            If itmX.Selected Then
                bItemSelected = True
                Exit For
            End If
        Next
        
        'See if there is a selected item
        If Not bItemSelected Then
            sMess = "You must select at least one Indemnity item from the View!"
            MsgBox sMess, vbExclamation + vbOKOnly, "Select at least one Indemnity item."
            Exit Sub
        End If
    
        
        sPayReqDesc = cboAssocRTChecksID.Text
        
        lvwIndemnity.Visible = False
        sMess = vbNullString
        Set colIndemID = New Collection
        For Each itmX In lvwIndemnity.ListItems
            If itmX.Selected Then
                'Do not allow association to indemnities marked as Previous Payment
                sFlagText = itmX.SubItems(GuiIndemListView.IsPreviousPayment - 1)
                If Not goUtil.GetFlagFromText(sFlagText) Then
                    colIndemID.Add itmX.SubItems(GuiIndemListView.ID - 1), itmX.SubItems(GuiIndemListView.ID - 1)
                Else
                    sIndemDesc = itmX.SubItems(GuiIndemListView.ClassOfLoss - 1)
                    sIndemDesc = sIndemDesc & " (" & itmX.SubItems(GuiIndemListView.ClassOfLossCode - 1) & ") - "
                    sIndemDesc = sIndemDesc & itmX.SubItems(GuiIndemListView.TypeOfLoss - 1)
                    sIndemDesc = sIndemDesc & " (" & itmX.SubItems(GuiIndemListView.TypeOfLossCode - 1) & ") - "
                    sIndemDesc = sIndemDesc & " [" & itmX.SubItems(GuiIndemListView.ReplacementCost - 1) & "]"
                    sMess = sMess & "Could not update " & sPayReqDesc & " for... " & vbCrLf
                    sMess = sMess & sIndemDesc & vbCrLf
                    sMess = sMess & "This item is marked as Previous Payment."
                End If
            End If
        Next
        
        If sMess <> vbNullString Then
            sTitle = "Could not update some items"
            MsgBox sMess, vbInformation + vbOKOnly, sTitle
        End If
        
        For Each vIndemID In colIndemID
            sIndemID = vIndemID
            If Not AssocIndemItemToRTChecksID(sIndemID, cboAssocRTChecksID) Then
                Exit Sub
            End If
        Next
    End If
    
    'Edit the Payment request if one is selected
    If cboAssocRTChecksID.ListIndex > 0 Then
        'Always reset this here
        Set mcolSelIndemID = New Collection
        For Each itmX In lvwIndemnity.ListItems
            If itmX.Selected Then
                sRTIndemnityID = itmX.SubItems(GuiIndemListView.RTIndemnityID - 1)
                mcolSelIndemID.Add sRTIndemnityID, """" & sRTIndemnityID & """"
            End If
        Next
        
        'First unselect all items
        For Each itmX In lvwPayReqs.ListItems
            itmX.Selected = False
        Next
        
        'Need to Select the Payment Request to Edit first
        sRTChecksID = cboAssocRTChecksID.ItemData(cboAssocRTChecksID.ListIndex)
        'Select the one needed
        For Each itmX In lvwPayReqs.ListItems
            sThisRTChecksID = itmX.SubItems(GuiPayReqsListView.RTChecksID - 1)
            If StrComp(sRTChecksID, sThisRTChecksID, vbTextCompare) = 0 Then
                itmX.Selected = True
                Exit For
            End If
        Next
        
        EditPayReqs , True
        
    End If
    
    RefreshIndemnity
    
    lvwIndemnity.Visible = True
    cmdAssocRTChecksID.Enabled = True
    
    Set itmX = Nothing
    Set colIndemID = Nothing
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdAssocRTChecksID_Click"
End Sub

Public Function AssocIndemItemToRTChecksID(psID As String, poCBO As Object) As Boolean
    On Error GoTo EH
    Dim sSQL As String
    Dim sPath As String
    Dim sRTChecksID As String
    Dim oConn As ADODB.Connection
    Dim MycboAssocRTChecksID As ComboBox
    
    Set MycboAssocRTChecksID = poCBO
    
    'Set the IDBillingCOunt to Drop down item data
    
    sRTChecksID = MycboAssocRTChecksID.ItemData(MycboAssocRTChecksID.ListIndex)
    
    'Check to See if it is Null value
    
    If sRTChecksID = 0 Then
        sRTChecksID = "null"
    End If
    
    Set oConn = New ADODB.Connection
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    sSQL = "UPDATE RTIndemnity SET "
    sSQL = sSQL & "[RTChecksID] = " & sRTChecksID & ", "
    sSQL = sSQL & "[IDRTChecks] = " & sRTChecksID & ", "
    sSQL = sSQL & "[UpLoadMe] = True, "
    sSQL = sSQL & "[DateLastUpdated] = #" & Format(Now(), "MM/DD/YYYY HH:MM:SS") & "#, "
    sSQL = sSQL & "[UpdateByUserID] = " & goUtil.gsCurUsersID & " "
    sSQL = sSQL & "WHERE [ID] = " & psID & " "
    sSQL = sSQL & "AND IDAssignments = " & msAssignmentsID & " "

    oConn.Execute sSQL
    
    AssocIndemItemToRTChecksID = True
    
    'clean up
    Set oConn = Nothing
    Set MycboAssocRTChecksID = Nothing
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function AssocIndemItemToRTChecksID"
End Function

Public Sub RefreshDed()
    On Error GoTo EH
    
    'Set pbUseAppDedClassTypeIDOrder flag
    mfrmClaim.SetadoRSPolicyLimits msAssignmentsID, True
    
    PopulatelvwDed
    PopulateAddPLAppClassTypeIDLookUp
    
    'Set policy limits back
    mfrmClaim.SetadoRSPolicyLimits msAssignmentsID, False
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub RefreshDed"
End Sub

Public Sub RefreshIndemnity()
    On Error GoTo EH
    'populate the Totals on the Indem screen
    If Not mfrmClaim.SetadoRSRTIndemnity(msAssignmentsID) Then
        Exit Sub
    End If
    PopulatelvwIndemnity lvwIndemnity
    
    'Load Payment Req RS (RTChecks)
    mfrmClaim.SetadoRSPayment msAssignmentsID, True
    cboAssocRTChecksID.Clear
    cboAssocRTChecksID.AddItem "(--Disassociate Payment--)"
    '0 indicates Null ID since ID must be >= 1 or <= -1
    '>=1    : WEB Server Synched
    '<=-1   : Client has yet to synch Data ID with Web Server.
    cboAssocRTChecksID.ItemData(cboAssocRTChecksID.NewIndex) = 0
    mfrmClaim.PopulateLookUp mfrmClaim.adoRSPayment, _
                        Nothing, _
                        cboAssocRTChecksID, _
                        "ID", _
                        vbNullString, _
                        "CheckNum", _
                        "PayReqDescription"
                        
    PopulateIndemTotals
                        
    Exit Sub
EH:
   goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub RefreshIndemnity"
End Sub

Public Sub RefreshPayReqs()
    On Error GoTo EH
    If Not mfrmClaim.SetadoRSRTChecks(msAssignmentsID) Then
        Exit Sub
    End If
    
    PopulatelvwPayReqs
    
    'Load Billing RS
    mfrmClaim.SetadoRSBillingCount msAssignmentsID, , , True
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
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub RefreshPayReqs"
End Sub

Private Sub cmdDelDed_Click()
    On Error GoTo EH
    Dim sMess As String
    Dim sID As String
    Dim itmX As ListItem
    
    If lvwDed.SelectedItem Is Nothing Then
        sMess = "You must select an item from the list!"
        MsgBox sMess, vbExclamation + vbOKOnly, "Nothing Selected"
        Exit Sub
    Else
        Set itmX = lvwDed.SelectedItem
        sID = itmX.SubItems(GuiPolicyLimits.ID - 1)
    End If
    
    sMess = "Are you sure you want to remove the selected item? "
    
    If MsgBox(sMess, vbQuestion + vbYesNo, "Remove Item From list") = vbYes Then
        'Remove the Selected item from the List View
        lvwDed.ListItems.Remove ("""" & sID & """")

        'And then Save the Sort Order
        CmdReNumberSort_Click
    End If
    
    'clean up
    Set itmX = Nothing
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdDelDed_Click"
End Sub

Private Sub cmdDelIndemnity_Click()
    On Error GoTo EH
    Dim itmX As MSComctlLib.ListItem
    Dim sIndemID As String
    Dim vIndemID As Variant
    Dim colIndemID As Collection
    
    
    If lvwIndemnity.ListItems.Count > 0 Then
        If MsgBox("Are you sure ?", vbYesNo, "DELETE SELECTED ITEMS") = vbYes Then
            lvwIndemnity.Visible = False
            Set colIndemID = New Collection
            For Each itmX In lvwIndemnity.ListItems
                If itmX.Selected Then
                    colIndemID.Add itmX.SubItems(GuiIndemListView.ID - 1), itmX.SubItems(GuiIndemListView.ID - 1)
                End If
            Next
            For Each vIndemID In colIndemID
                sIndemID = vIndemID
                If DeleteIndemnityItem(sIndemID) Then
                    lvwIndemnity.ListItems.Remove ("""" & sIndemID & """")
                End If
            Next
            mfrmClaim.RefreshMe
        End If
    End If
    
    lvwIndemnity.Visible = True
    Set itmX = Nothing
    Set colIndemID = Nothing
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdDelIndemnity_Click"
End Sub

Private Sub cmdDelPayReqs_Click()
    On Error GoTo EH
    Dim itmX As MSComctlLib.ListItem
    Dim sPayReqID As String
    Dim vPayReqID As Variant
    Dim colPayReqID As Collection
    
    
    If lvwPayReqs.ListItems.Count > 0 Then
        If MsgBox("Are you sure ?", vbYesNo, "DELETE SELECTED ITEMS") = vbYes Then
            lvwPayReqs.Visible = False
            Set colPayReqID = New Collection
            For Each itmX In lvwPayReqs.ListItems
                If itmX.Selected Then
                    colPayReqID.Add itmX.SubItems(GuiPayReqsListView.ID - 1), itmX.SubItems(GuiPayReqsListView.ID - 1)
                End If
            Next
            For Each vPayReqID In colPayReqID
                sPayReqID = vPayReqID
                If DeletePayReqItem(sPayReqID) Then
                    lvwPayReqs.ListItems.Remove ("""" & sPayReqID & """")
                End If
            Next
            mfrmClaim.RefreshMe
        End If
    End If
    
    lvwPayReqs.Visible = True
    Set itmX = Nothing
    Set colPayReqID = Nothing
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdDelPayReqs_Click"
End Sub

Private Sub cmdDown_Click()
    goUtil.utMoveListItem lvwDed, MoveDown
End Sub

Private Sub cmdEditIndemnity_Click()
    On Error GoTo EH
    
    cmdEditIndemnity.Enabled = False
    
    EditIndemnity
    
    cmdEditIndemnity.Enabled = True
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdEditIndemnity_Click"
End Sub

Public Function EditIndemnity() As Boolean
    On Error GoTo EH
    Dim oListView As MSComctlLib.ListView
    Dim itmX As MSComctlLib.ListItem
    Dim sMyIndemID As String

    If lvwIndemnity.ListItems.Count = 0 Then
        Exit Function
    Else
        Set oListView = lvwIndemnity
    End If

    Set itmX = oListView.SelectedItem
    
    With EditIndem
        .MyIndemnity = Me
        .MyfrmClaim = Me.MyfrmClaim
        .MyGUI = Me.MyGUI
        .AssignmentsID = itmX.SubItems(GuiIndemListView.IDAssignments - 1)
        sMyIndemID = itmX.SubItems(GuiIndemListView.ID - 1)
        .IndemID = sMyIndemID
         Load EditIndem
        .Caption = "Edit Indemnity"
        .cmdSave.Enabled = False
        .WindowState = vbNormal
        .Show vbModal
    End With
   
    'If the current indemnity does not have a class of loss selected
    'then need to Delete that item from Indemnity List
    
    If EditIndem.COLOrigListIndex = -1 And EditIndem.cmdSave.Enabled Then
        If DeleteIndemnityItem(sMyIndemID) Then
            lvwIndemnity.ListItems.Remove ("""" & sMyIndemID & """")
        End If
    End If
    
    EditIndem.CLEANUP
    Unload EditIndem
    Set EditIndem = Nothing
    If oListView.Visible Then
        oListView.SetFocus
    End If
    
    Sleep 500
    PopulateAppliedDeductible
    RefreshPayReqs
    PopulateIndemTotals
    'Reports  mfrmReports
    If Not mfrmClaim.MyReports Is Nothing Then
        mfrmClaim.MyReports.LoadMe
    End If
    
    EditIndemnity = True
    
    Set oListView = Nothing
    Set itmX = Nothing
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function EditIndemnity"
    Unload EditPayReq
End Function

Public Function EditIdemnityItem(pudtIndem As GuiIndemItem) As Boolean
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    
    With pudtIndem
        If .RTIndemnityID = vbNullString Or .RTIndemnityID = "0" Then
            .RTIndemnityID = "Null"
        End If
        .AssignmentsID = msAssignmentsID
        If .RTChecksID = vbNullString Or .RTChecksID = "0" Then
            .RTChecksID = "Null"
        End If
        If .ID = vbNullString Or .ID = "0" Then
            .ID = "Null"
        End If
        .IDAssignments = msAssignmentsID
        If .IDRTChecks = vbNullString Or .IDRTChecks = "0" Then
            .IDRTChecks = "Null"
        End If
    End With

    sSQL = "UPDATE RTIndemnity Set "
    sSQL = sSQL & "[RTIndemnityID] = " & pudtIndem.RTIndemnityID & ", "
    sSQL = sSQL & "[AssignmentsID] = " & pudtIndem.AssignmentsID & ", "
    sSQL = sSQL & "[RTChecksID] = " & pudtIndem.RTChecksID & ", "
    sSQL = sSQL & "[ID] = " & pudtIndem.ID & ", "
    sSQL = sSQL & "[IDAssignments] = " & pudtIndem.IDAssignments & ", "
    sSQL = sSQL & "[IDRTChecks] = " & pudtIndem.IDRTChecks & ", "
    sSQL = sSQL & "[ACVClaim] = " & CCur(pudtIndem.ACVClaim) & ", "
    sSQL = sSQL & "[ACVLessExcessLimits] = " & CCur(pudtIndem.ACVLessExcessLimits) & ", "
    sSQL = sSQL & "[SpecialLimits] = " & CCur(pudtIndem.SpecialLimits) & ", "
    sSQL = sSQL & "[ExcessLimits] = " & CCur(pudtIndem.ExcessLimits) & ", "
    sSQL = sSQL & "[Miscellaneous] = " & CCur(pudtIndem.Miscellaneous) & ", "
    sSQL = sSQL & "[MiscDescription] = '" & goUtil.utCleanSQLString(pudtIndem.MiscDescription) & "', "
    sSQL = sSQL & "[IsAddAmountOfInsurance] = " & pudtIndem.IsAddAmountOfInsurance & ", "
    sSQL = sSQL & "[ExcessAbsorbsDeductible] = " & pudtIndem.ExcessAbsorbsDeductible & ", "
    sSQL = sSQL & "[AppliedDeductible] = " & CCur(pudtIndem.AppliedDeductible) & ", "
    sSQL = sSQL & "[NonRecoverableDep] = " & CCur(pudtIndem.NonRecoverableDep) & ", "
    sSQL = sSQL & "[RecoverableDep] = " & CCur(pudtIndem.RecoverableDep) & ", "
    sSQL = sSQL & "[ReplacementCost] = " & CCur(pudtIndem.ReplacementCost) & ", "
    sSQL = sSQL & "[TypeOfLossID]  = " & pudtIndem.TypeOfLossID & ", "
    sSQL = sSQL & "[ClassOfLossID] = " & pudtIndem.ClassOfLossID & ", "
    sSQL = sSQL & "[Description] = '" & goUtil.utCleanSQLString(pudtIndem.Description) & "', "
    sSQL = sSQL & "[IsPreviousPayment] = " & pudtIndem.IsPreviousPayment & ", "
    If IsDate(pudtIndem.PPayDatePaid) Then
        sSQL = sSQL & "[PPayDatePaid] = #" & pudtIndem.PPayDatePaid & "#, "
        sSQL = sSQL & "[PPayAmountPaid] = " & CCur(pudtIndem.PPayAmountPaid) & ", "
        sSQL = sSQL & "[PPayCheckNumber] = '" & goUtil.utCleanSQLString(pudtIndem.PPayCheckNumber) & "', "
    Else
        sSQL = sSQL & "[PPayDatePaid] = Null, "
        sSQL = sSQL & "[PPayAmountPaid] = Null, "
        sSQL = sSQL & "[PPayCheckNumber] = Null, "
    End If
    sSQL = sSQL & "[IsDeleted] = " & pudtIndem.IsDeleted & ", "
    sSQL = sSQL & "[DownLoadMe] = " & pudtIndem.DownLoadMe & ", "
    sSQL = sSQL & "[UpLoadMe] = " & pudtIndem.UpLoadMe & ", "
    sSQL = sSQL & "[AdminComments] = '" & goUtil.utCleanSQLString(pudtIndem.AdminComments) & "', "
    sSQL = sSQL & "[DateLastUpdated] = #" & pudtIndem.DateLastUpdated & "#, "
    sSQL = sSQL & "[UpdateByUserID]  = " & pudtIndem.UpdateByUserID & " "
    sSQL = sSQL & "WHERE IDAssignments = " & pudtIndem.IDAssignments & " "
    sSQL = sSQL & "AND ID = " & pudtIndem.ID & " "

    Set oConn = New ADODB.Connection
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    oConn.Execute sSQL
    
    EditIdemnityItem = True
    'Clean up
    Set oConn = Nothing
    Exit Function

EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function EditIdemnityItem"
End Function

Private Sub cmdEditPayReqs_Click()
    On Error GoTo EH
    
    cmdEditPayReqs.Enabled = False
    
    EditPayReqs
    
    cmdEditPayReqs.Enabled = True
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdEditPayReqs_Click"
End Sub

Public Function EditPayReqs(Optional pbAssociateFromIndem As Boolean = False, Optional pbEditGetAmount As Boolean = False) As Boolean
    On Error GoTo EH
    Dim oListView As MSComctlLib.ListView
    Dim itmX As MSComctlLib.ListItem
    Dim itmX2 As MSComctlLib.ListItem
    Dim lMyPayReqID As Long
    Dim RS As ADODB.Recordset
    Dim sInsured As String
    Dim sMortgagee As String
    Dim sAddress As String
     'Get some info from Assignments Rcord
    Set RS = mfrmClaim.adoRSAssignments
    
    If RS.RecordCount = 1 Then
        RS.MoveFirst
        sInsured = goUtil.IsNullIsVbNullString(RS.Fields("Insured"))
        sMortgagee = goUtil.IsNullIsVbNullString(RS.Fields("MortgageeName"))
        sAddress = goUtil.IsNullIsVbNullString(RS.Fields("MAStreet")) & "    "
        sAddress = sAddress & goUtil.IsNullIsVbNullString(RS.Fields("MACity")) & ", "
        sAddress = sAddress & goUtil.IsNullIsVbNullString(RS.Fields("MAState")) & " "
        sAddress = sAddress & Format(goUtil.IsNullIsVbNullString(RS.Fields("MAZIP")), "00000")
'        & "-"
'        sAddress = sAddress & Format(goUtil.IsNullIsVbNullString(RS.Fields("MAZIP4")), "0000") & " "
    End If

    If lvwPayReqs.ListItems.Count = 0 Then
        Exit Function
    Else
        Set oListView = lvwPayReqs
    End If

    Set itmX = oListView.SelectedItem
    
    With EditPayReq
        .MyIndemnity = Me
        .MyfrmClaim = Me.MyfrmClaim
        .MyGUI = Me.MyGUI
        .AssignmentsID = itmX.SubItems(GuiPayReqsListView.IDAssignments - 1)
        lMyPayReqID = itmX.SubItems(GuiPayReqsListView.ID - 1)
        .PayReqID = lMyPayReqID
        .Insured = sInsured
        .Mortgagee = sMortgagee
        .Address = sAddress
        .AssociateFromIndem = pbAssociateFromIndem
        .EditGetAmount = pbEditGetAmount
        Load EditPayReq
        .Caption = "Edit Payment Request"
        .cmdSave.Enabled = False
'        .WindowState = vbMaximized
        .ShowEditRptParam
        .Show vbModal
    End With
    
    'If the current indemnity does not have a class of loss selected
    'then need to Delete that item from Indemnity List
    
    If EditPayReq.COLOrigListIndex = -1 And EditPayReq.cmdSave.Enabled Then
        If DeletePayReqItem(CStr(lMyPayReqID)) Then
            lvwPayReqs.ListItems.Remove ("""" & CStr(lMyPayReqID) & """")
        End If
    Else
        For Each itmX2 In EditPayReq.lvwRptParams.ListItems
            'Snag the NoOfRequests Param
            If StrComp("f_p00_sNumberOfRequests", itmX2.SubItems(GuiRptParamsListView.ParamName - 1), vbTextCompare) = 0 Then
                'Set the Number of Requests here !
                itmX.SubItems(GuiPayReqsListView.NoOfRequests - 1) = itmX2.SubItems(GuiRptParamsListView.ParamValue - 1)
                itmX.SubItems(GuiPayReqsListView.NoOfRequestsSort - 1) = itmX2.SubItems(GuiRptParamsListView.ParamValue - 1)
                Exit For
            End If
        Next
    End If
    
    EditPayReq.CLEANUP
    Unload EditPayReq
    Set EditPayReq = Nothing
    If oListView.Visible Then
        oListView.SetFocus
    End If
    
    RefreshIndemnity
    RefreshPayReqs
    'Reports  mfrmReports
    If Not mfrmClaim.MyReports Is Nothing Then
        mfrmClaim.MyReports.LoadMe
    End If
    
    EditPayReqs = True
    
    Set oListView = Nothing
    Set itmX = Nothing
    Set RS = Nothing
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function EditPayReqs"
    Unload EditPayReq
End Function

Private Sub cmdExit_Click()
    On Error GoTo EH
    
    mbUnloadMe = True
    Me.Visible = False
    mfrmClaim.Timer_UnloadForm.Enabled = True
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdExit_Click"
End Sub

Private Sub cmdFindlIndemnity_Click()
    On Error GoTo EH
    If lvwIndemnity.ListItems.Count > 0 Then
        mlLastFindIndex = 0
        goUtil.utFindListItem Me, lvwIndemnity, msFindText, mlLastFindIndex
    End If
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdFindlIndemnity_Click"
End Sub

Private Sub cmdFindNextlIndemnity_Click()
    On Error GoTo EH
    
    If msFindText <> vbNullString And lvwIndemnity.ListItems.Count > 0 Then
        goUtil.utFindListItem Me, lvwIndemnity, msFindText, mlLastFindIndex
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdFindNextlIndemnity_Click"
End Sub

Private Sub cmdFindNextPayReqs_Click()
    On Error GoTo EH
    
    If msFindText <> vbNullString And lvwPayReqs.ListItems.Count > 0 Then
        goUtil.utFindListItem Me, lvwPayReqs, msFindText, mlLastFindIndex
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdFindNextPayReqs_Click"
End Sub

Private Sub cmdFindPayReqs_Click()
    On Error GoTo EH
    If lvwPayReqs.ListItems.Count > 0 Then
        mlLastFindIndex = 0
        goUtil.utFindListItem Me, lvwPayReqs, msFindText, mlLastFindIndex
    End If
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdFindPayReqs_Click"
End Sub


Private Sub cmdPrintIndem_Click()
    On Error GoTo EH
    Screen.MousePointer = vbHourglass
    
    cmdPrintIndem.Enabled = False
    If PrintIndem(msAssignmentsID) Then
        If Not mbUnloadMe Then
            cmdPrintIndem.Enabled = True
        End If
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
EH:
    Screen.MousePointer = vbDefault
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdPrintIndem_Click"
End Sub

Public Function PrintIndem(psIDAssignments As String) As Boolean
    On Error GoTo EH
    Dim MyIndem As Object
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
                                    "_arIndem", _
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
    
    sPrintPreview = GetSetting(goUtil.gsMainAppEXEName, "GENERAL", "PRINT_PREVIEW", "USE_ADOBE")
    
    Select Case UCase(sPrintPreview)
        Case "USE_ADOBE"
            bUseAdobeReader = True
    End Select
    
    lMainSPVersion = mfrmClaim.adoRSAssignments.Fields("SPVersion").Value
    Set oConn = New ADODB.Connection
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    lrptVersion = goUtil.GetApplicationVersionNumber(lMainSPVersion, sProjectName, oConn)
    sParams = sParams & "psAssignmentsID=" & psIDAssignments & "|"
    sParams = sParams & "pbPreview=" & "True" & "|"
    
    If bUseAdobeReader Then
        sPDFFilePath = goUtil.gsInstallDir & "\TempActiveReport" & goUtil.utGetTickCount & ".pdf"
        sParams = sParams & "psXportPath=" & sPDFFilePath & "|"
        sParams = sParams & "pPDFJPEGQuality=" & "50" & "|"
        sParams = sParams & "pXportType=" & ExportType.ARPdf & "|"
    Else
        sParams = sParams & "pbGetObjectOnly=" & "True" & "|"
    End If

    Set oCarList = CreateObject(goUtil.goCurCarList.ClassName)
    
    If bUseAdobeReader Then
        oCarList.GetARReport sReportName, lrptVersion, sParams
        If goUtil.utFileExists(sPDFFilePath) Then
            goUtil.utShellExecute , "OPEN", sPDFFilePath, , , vbNormalFocus, True, True, True, "Indemnity Report"
            DoEvents
            Sleep 1000
            goUtil.utDeleteFile sPDFFilePath
            oCarList.CLEANUP
            Set oCarList = Nothing
        End If
    Else
    
        Set MyIndem = oCarList.GetARReport(sReportName, lrptVersion, sParams)
    
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
            MyIndem.Run False 'True
            .objARvReport = MyIndem
            .sRptTitle = "Indemnity Report"
            .HidePrintButton = False
            .ShowReportOnForm moForm, vbModeless
            Unload .objARvReport
            Set .objARvReport = Nothing
            Unload MyIndem
            Set MyIndem = Nothing
            oCarList.CLEANUP
            Set oCarList = Nothing
        End With
    End If
    
    'Clear the Local ref to this report object only
    'The actual cleanup of this active report object will occur within gARV
    Set oConn = Nothing
    Set oCarList = Nothing
    Set adoRSApplication = Nothing
    
    PrintIndem = True
    
    Exit Function
EH:
    Screen.MousePointer = vbDefault
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function PrintIndem"
End Function

Private Sub cmdPrintList_Click()
    On Error GoTo EH
    
    If lvwPayReqs.Visible Then
        goUtil.utPrintListView App.EXEName, lvwPayReqs, "Payment Requests"
    End If
    
    If lvwDed.Visible Then
        goUtil.utPrintListView App.EXEName, lvwDed, "Apply Deductible Order"
    End If
    
    If lvwIndemnity.Visible Then
        goUtil.utPrintListView App.EXEName, lvwIndemnity, "Indemnity Entry"
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdPrintList_Click"
End Sub

Private Sub cmdRefreshIndemnity_Click()
    On Error GoTo EH
    
    cmdRefreshIndemnity.Enabled = False
    Screen.MousePointer = vbHourglass
    If LoadMe Then
        cmdRefreshIndemnity.Enabled = True
    End If
    Screen.MousePointer = vbDefault
    
    Exit Sub
EH:
    Screen.MousePointer = vbDefault
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdRefreshIndemnity_Click"
End Sub

Private Sub cmdRefreshPayReqs_Click()
    On Error GoTo EH
    
    cmdRefreshPayReqs.Enabled = False
    Screen.MousePointer = vbHourglass
    If LoadMe Then
        cmdRefreshPayReqs.Enabled = True
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
EH:
    Screen.MousePointer = vbDefault
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdRefreshPayReqs_Click"
End Sub

Private Sub CmdReNumberSort_Click()
    On Error GoTo EH
    
    CmdReNumberSort.Enabled = False
    
    If ReOrderDedApply() Then
        CmdReNumberSort.Enabled = True
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub CmdReNumberSort_Click"
End Sub

Public Function ADDDedApply() As Boolean
    On Error GoTo EH
    Dim sAppDedClassTypeIDOrder As String
    Dim sSQL As String
    Dim oConn As ADODB.Connection
    
    sAppDedClassTypeIDOrder = cboAddPLAppClassTypeID.ItemData(cboAddPLAppClassTypeID.ListIndex)
    'Need to Update the The Assignments RS
    
    sSQL = "UPDATE Assignments SET "
    sSQL = sSQL & "AppDedClassTypeIDOrder = IIF(AppDedClassTypeIDOrder Is Null Or AppDedClassTypeIDOrder = '','" & sAppDedClassTypeIDOrder & "',AppDedClassTypeIDOrder & '," & sAppDedClassTypeIDOrder & "' ), "
    sSQL = sSQL & "UpLoadMe = True, "
    sSQL = sSQL & "DateLastUpdated = #" & Now() & "#, "
    sSQL = sSQL & "UpdateByUserID = " & goUtil.gsCurUsersID & " "
    sSQL = sSQL & "WHERE ID = " & msAssignmentsID & " "
    
    Set oConn = New ADODB.Connection
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    oConn.Execute sSQL
    
    Sleep 500
    
    'Refresh Entire CLaim
    mfrmClaim.RefreshMe
    
    ADDDedApply = True
    
    Set oConn = Nothing
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function ADDDedApply"
End Function

Public Function ReOrderDedApply() As Boolean
    On Error GoTo EH
    Dim oListView As MSComctlLib.ListView
    Dim itmX  As MSComctlLib.ListItem
    Dim sAppDedClassTypeIDOrder As String
    Dim sSQL As String
    Dim oConn As ADODB.Connection
    
    Screen.MousePointer = vbHourglass
    
    Set oListView = lvwDed
    For Each itmX In oListView.ListItems
        'Build the Apply Deductible Class Type ID Order
        'This is the order in which Deductible will be applied
        'to each Line of coverage
        If sAppDedClassTypeIDOrder = vbNullString Then
            sAppDedClassTypeIDOrder = itmX.SubItems(GuiPolicyLimits.ClassTypeID - 1)
        Else
            sAppDedClassTypeIDOrder = sAppDedClassTypeIDOrder & "," & itmX.SubItems(GuiPolicyLimits.ClassTypeID - 1)
        End If
    Next
    
    'Need to Update the The Assignments RS
    
    sSQL = "UPDATE Assignments SET "
    sSQL = sSQL & "AppDedClassTypeIDOrder = '" & sAppDedClassTypeIDOrder & "', "
    sSQL = sSQL & "UpLoadMe = True, "
    sSQL = sSQL & "DateLastUpdated = #" & Now() & "#, "
    sSQL = sSQL & "UpdateByUserID = " & goUtil.gsCurUsersID & " "
    sSQL = sSQL & "WHERE ID = " & msAssignmentsID & " "
    
    Set oConn = New ADODB.Connection
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    oConn.Execute sSQL
    
    Sleep 500

    'Refresh Entire CLaim
    mfrmClaim.RefreshMe
    ReOrderDedApply = True
    Screen.MousePointer = vbDefault
    Set itmX = Nothing
    Set oListView = Nothing
    Set oConn = Nothing
    Exit Function
EH:
    Screen.MousePointer = vbDefault
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function ReOrderDedApply"
End Function

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

Private Sub cmdSelAllIndemnity_Click()
    On Error GoTo EH
    Dim itmX As ListItem
    For Each itmX In lvwIndemnity.ListItems
        itmX.Selected = True
    Next
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdSelAllIndemnity_Click"
End Sub

Private Sub cmdSelAllPayReqs_Click()
    On Error GoTo EH
    Dim itmX As ListItem
    For Each itmX In lvwPayReqs.ListItems
        itmX.Selected = True
    Next
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdSelAllPayReqs_Click"
End Sub

Private Sub cmdSpelling_Click()

End Sub

Private Sub cmdUp_Click()
    goUtil.utMoveListItem lvwDed, MoveUp
End Sub

Private Sub cmdAddPayReqAssIndem_Click()
    On Error GoTo EH
    Dim sMess As String
    Dim bItemSelected As Boolean
    Dim itmX As MSComctlLib.ListItem
    Dim sRTIndemnityID As String
    
    'Always reset this here
    Set mcolSelIndemID = New Collection
    
    If lvwIndemnity.ListItems.Count > 0 Then
        For Each itmX In lvwIndemnity.ListItems
            If itmX.Selected Then
                'Make sure that there is no payment already assoicated
                If itmX.Text = vbNullString Then
                    bItemSelected = True
                    'Add this item to collecion
                    'Only add ones that are not already associated
                    'User must first disassociate the item if it is already associated before
                    'allowing the association to occur when adding a payment request.
                    sRTIndemnityID = itmX.SubItems(GuiIndemListView.RTIndemnityID - 1)
                    mcolSelIndemID.Add sRTIndemnityID, """" & sRTIndemnityID & """"
                End If
            End If
        Next
        'See if there is a selected item
        If Not bItemSelected Then
            sMess = "To " & Replace(cmdAddPayReqAssIndem.Caption, "&", vbNullString) & " : " & vbCrLf & vbCrLf
            sMess = sMess & "First, you must select at least one or more Indemnity item(s) " & vbCrLf
            sMess = sMess & "that are not already associated to a Payment Request. " & vbCrLf
            MsgBox sMess, vbExclamation + vbOKOnly, "Select Indemnity item."
            Exit Sub
        End If
        
        On Error GoTo EH
    
        cmdAddPayReqAssIndem.Enabled = False
        
        If AddPayReqs(True) Then
            cmdAddPayReqAssIndem.Enabled = True
        End If
    End If
    
    Set itmX = Nothing
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdAddPayReqAssIndem_Click"
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
    Dim bHideDeleted As Boolean
    
    mbLoading = True
    'Set the Form Icon
    Me.Icon = mfrmClaim.optClaim(GuiClaimOptions.opt03_Indemnity).Picture
    LoadHeaderlvwDed
    LoadHeaderlvwIndemnity lvwIndemnity
    LoadHeaderlvwPayReqs
    bHideDeleted = CBool(GetSetting(App.EXEName, "GENERAL", "HIDE_DELETED", True))
    If bHideDeleted Then
        chkHideDeleted(0).Value = vbChecked
        chkHideDeleted(1).Value = vbChecked
    Else
        chkHideDeleted(0).Value = vbUnchecked
        chkHideDeleted(1).Value = vbUnchecked
    End If
    LoadMe
    CheckStatus
    
    ShowFrame
    cmdSave.Enabled = False
    mbLoading = False
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_Load"
End Sub

Public Sub LoadHeaderlvwDed()
    On Error GoTo EH
    Dim bGridOn As Boolean
    Dim bHideDeleted As Boolean
    Dim bHideUploadFlags As Boolean
    
    bHideDeleted = CBool(GetSetting(App.EXEName, "GENERAL", "HIDE_DELETED", True))
    bHideUploadFlags = CBool(GetSetting(App.EXEName, "GENERAL", "HIDE_UPLOAD_FLAGS", True))
    
    'set the columnheaders
    With lvwDed
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
        .ColumnHeaders.Add , "SortOrder", "SortOrder" 'Hidden
        
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
            .ColumnHeaders.Item(GuiPolicyLimits.IsDeleted).Width = 0  'Hidden 400
        Else
            .ColumnHeaders.Item(GuiPolicyLimits.IsDeleted).Width = 400
        End If
        .ColumnHeaders.Item(GuiPolicyLimits.IsDeleted).Alignment = lvwColumnCenter
        .ColumnHeaders.Item(GuiPolicyLimits.IsDeleted).Icon = GuiPolicyLimitsPic.IsDeleted
        'UpLoad Me
        If bHideUploadFlags Then
            .ColumnHeaders.Item(GuiPolicyLimits.UpLoadMe).Width = 0  'Hidden 400
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
        .ColumnHeaders.Item(GuiPolicyLimits.AdminComments).Width = 0  'Hidden 10000
        .ColumnHeaders.Item(GuiPolicyLimits.AdminComments).Alignment = lvwColumnLeft
        'ID
        .ColumnHeaders.Item(GuiPolicyLimits.ID).Width = 0  'hidden
        .ColumnHeaders.Item(GuiPolicyLimits.ID).Alignment = lvwColumnLeft
        'IDAssignments
        .ColumnHeaders.Item(GuiPolicyLimits.IDAssignments).Width = 0  'Hidden
        .ColumnHeaders.Item(GuiPolicyLimits.IDAssignments).Alignment = lvwColumnLeft
        'SortOrder
        .ColumnHeaders.Item(GuiPolicyLimits.SortOrder).Width = 0  'Hidden
        .ColumnHeaders.Item(GuiPolicyLimits.SortOrder).Alignment = lvwColumnLeft
    End With
    
    bGridOn = CBool(GetSetting(App.EXEName, "GENERAL", "GRID_ON", False))
    lvwDed.GridLines = bGridOn
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub LoadHeaderlvwDed"
End Sub

Public Sub LoadHeaderlvwIndemnity(poLvw As ListView)
    On Error GoTo EH
    Dim bGridOn As Boolean
    Dim oListView As ListView
    Dim bHideDeleted As Boolean
    Dim bHideUploadFlags As Boolean
    
    bHideDeleted = CBool(GetSetting(App.EXEName, "GENERAL", "HIDE_DELETED", True))
    bHideUploadFlags = CBool(GetSetting(App.EXEName, "GENERAL", "HIDE_UPLOAD_FLAGS", True))
    
    Set oListView = poLvw
    'set the columnheaders
    With oListView
        .ColumnHeaders.Add , "PaymentRequest", "Payment"
        .ColumnHeaders.Add , "ClassOfLoss", "Class"
        .ColumnHeaders.Add , "ClassOfLossCode", "Code"
        .ColumnHeaders.Add , "TypeOfLoss", "Type Of Loss"
        .ColumnHeaders.Add , "TypeOfLossCode", "Code"
        .ColumnHeaders.Add , "Description", "Description"
        .ColumnHeaders.Add , "ReplacementCost", "Rep Cost"
        .ColumnHeaders.Add , "ReplacementCostSort", "ReplacementCostSort" 'Hidden
        .ColumnHeaders.Add , "RecoverableDep", "Rec Dep"
        .ColumnHeaders.Add , "RecoverableDepSort", "RecoverableDepSort" 'Hidden
        .ColumnHeaders.Add , "NonRecoverableDep", "Non Rec Dep"
        .ColumnHeaders.Add , "NonRecoverableDepSort", "NonRecoverableDepSort" 'Hidden
        .ColumnHeaders.Add , "ACVClaim", "Acv Claim"
        .ColumnHeaders.Add , "ACVClaimSort", "ACVClaimSort" 'Hidden
        .ColumnHeaders.Add , "SpecialLimits", "Special Limits"
        .ColumnHeaders.Add , "SpecialLimitsSort", "SpecialLimitsSort" 'Hidden
        .ColumnHeaders.Add , "IsAddAmountOfInsurance", "Is Add Amount Of Insurance" 'Hidden
        .ColumnHeaders.Add , "ExcessAbsorbsDeductible", "ExcessAbsorbsDeductible" 'Hidden
        .ColumnHeaders.Add , "AppliedDeductible", "App Ded"
        .ColumnHeaders.Add , "AppliedDeductibleSort", "AppliedDeductibleSort" 'Hidden
        .ColumnHeaders.Add , "ExcessLimits", "Excess Lim"
        .ColumnHeaders.Add , "ExcessLimitsSort", "ExcessLimitsSort" 'Hidden
        .ColumnHeaders.Add , "Miscellaneous", "Misc."
        .ColumnHeaders.Add , "MiscellaneousSort", "MiscellaneousSort" 'Hidden
        .ColumnHeaders.Add , "MiscellaneousDesc", "Misc. Desc"
        .ColumnHeaders.Add , "ACVLessExcessLimits", "ACV-Excess Lim"
        .ColumnHeaders.Add , "ACVLessExcessLimitsSort", "ACVLessExcessLimitsSort" 'Hidden
        .ColumnHeaders.Add , "IsPreviousPayment", "Prev Payment"
        .ColumnHeaders.Add , "PPayDatePaid", "Date Paid"
        .ColumnHeaders.Add , "PPayDatePaidSort", "PPayDatePaidSort" 'Hidden
        .ColumnHeaders.Add , "PPayAmountPaid", "Amount Paid"
        .ColumnHeaders.Add , "PPayAmountPaidSort", "PPayAmountPaidSort" 'hidden
        .ColumnHeaders.Add , "PPayCheckNumber", "Check No."
        .ColumnHeaders.Add , "PPayCheckNumberSort", "PPayCheckNumberSort" 'Hidenn"
        .ColumnHeaders.Add , "IsDeleted", "Is Deleted" 'Hidden
        .ColumnHeaders.Add , "UpLoadMe", "Up LoadMe" 'Hidden
        .ColumnHeaders.Add , "DateLastUpdated", "Date Last Updated"
        .ColumnHeaders.Add , "DateLastUpdatedSort", "DateLastUpdatedSort" 'Hidden
        .ColumnHeaders.Add , "AdminComments", "Admin Comments" 'Hidden
        .ColumnHeaders.Add , "ClassOfLossID", "ClassOfLossID" 'Hidden
        .ColumnHeaders.Add , "TypeOfLossID", "TypeOfLossID" 'Hidden
        .ColumnHeaders.Add , "ID", "ID" 'Hidden
        .ColumnHeaders.Add , "IDAssignments", "IDAssignments" 'Hidden
        .ColumnHeaders.Add , "RTIndemnityID", "RTIndemnityID" 'Hidden
        .ColumnHeaders.Add , "AssignmentsID", "AssignmentsID" 'Hidden
        .ColumnHeaders.Add , "RTChecksID", "RTChecksID" 'Hidden
        .ColumnHeaders.Add , "IDRTChecks", "IDRTChecks" 'Hidden
        .ColumnHeaders.Add , "DownLoadMe", "DownLoadMe" 'Hidden
        .ColumnHeaders.Add , "UpdateByUserID", "UpdateByUserID" 'Hidden
        
        .Sorted = False
        .SortOrder = lvwAscending
        
        'PaymentRequest
        .ColumnHeaders.Item(GuiIndemListView.PaymentRequest).Width = 1000
        .ColumnHeaders.Item(GuiIndemListView.PaymentRequest).Alignment = lvwColumnLeft
        'ClassOfLoss
        .ColumnHeaders.Item(GuiIndemListView.ClassOfLoss).Width = 3000
        .ColumnHeaders.Item(GuiIndemListView.ClassOfLoss).Alignment = lvwColumnLeft
        'ClassOfLossCode
        .ColumnHeaders.Item(GuiIndemListView.ClassOfLossCode).Width = 700
        .ColumnHeaders.Item(GuiIndemListView.ClassOfLossCode).Alignment = lvwColumnLeft
        'TypeOfLoss
        .ColumnHeaders.Item(GuiIndemListView.TypeOfLoss).Width = 3000
        .ColumnHeaders.Item(GuiIndemListView.TypeOfLoss).Alignment = lvwColumnLeft
        'TypeOfLossCode
        .ColumnHeaders.Item(GuiIndemListView.TypeOfLossCode).Width = 700
        .ColumnHeaders.Item(GuiIndemListView.TypeOfLossCode).Alignment = lvwColumnLeft
        'Description
        .ColumnHeaders.Item(GuiIndemListView.Description).Width = 5000
        .ColumnHeaders.Item(GuiIndemListView.Description).Alignment = lvwColumnLeft
        'ReplacementCost
        .ColumnHeaders.Item(GuiIndemListView.ReplacementCost).Width = 1200
        .ColumnHeaders.Item(GuiIndemListView.ReplacementCost).Alignment = lvwColumnRight
        'ReplacementCostSort
        .ColumnHeaders.Item(GuiIndemListView.ReplacementCostSort).Width = 0 'hidden
        .ColumnHeaders.Item(GuiIndemListView.ReplacementCostSort).Alignment = lvwColumnLeft
        'RecoverableDep
        .ColumnHeaders.Item(GuiIndemListView.RecoverableDep).Width = 1200
        .ColumnHeaders.Item(GuiIndemListView.RecoverableDep).Alignment = lvwColumnRight
        'RecoverableDepSort
        .ColumnHeaders.Item(GuiIndemListView.RecoverableDepSort).Width = 0 'hidden
        .ColumnHeaders.Item(GuiIndemListView.RecoverableDepSort).Alignment = lvwColumnLeft
        'NonRecoverableDep
        .ColumnHeaders.Item(GuiIndemListView.NonRecoverableDep).Width = 1200
        .ColumnHeaders.Item(GuiIndemListView.NonRecoverableDep).Alignment = lvwColumnRight
        'NonRecoverableDepSort
        .ColumnHeaders.Item(GuiIndemListView.NonRecoverableDepSort).Width = 0 'hidden
        .ColumnHeaders.Item(GuiIndemListView.NonRecoverableDepSort).Alignment = lvwColumnLeft
        'ACVClaim
        .ColumnHeaders.Item(GuiIndemListView.ACVClaim).Width = 1200
        .ColumnHeaders.Item(GuiIndemListView.ACVClaim).Alignment = lvwColumnRight
        'ACVClaimSort
        .ColumnHeaders.Item(GuiIndemListView.ACVClaimSort).Width = 0 'hidden
        .ColumnHeaders.Item(GuiIndemListView.ACVClaimSort).Alignment = lvwColumnLeft
        'SpecialLimits
        .ColumnHeaders.Item(GuiIndemListView.SpecialLimits).Width = 1500
        .ColumnHeaders.Item(GuiIndemListView.SpecialLimits).Alignment = lvwColumnRight
        'SpecialLimitsSort
        .ColumnHeaders.Item(GuiIndemListView.SpecialLimitsSort).Width = 0 'hidden
        .ColumnHeaders.Item(GuiIndemListView.SpecialLimitsSort).Alignment = lvwColumnLeft
        'IsAddAmountOfInsurance
        .ColumnHeaders.Item(GuiIndemListView.IsAddAmountOfInsurance).Width = 0 'hidden 400
        .ColumnHeaders.Item(GuiIndemListView.IsAddAmountOfInsurance).Alignment = lvwColumnCenter
        'ExcessAbsorbsDeductible
        .ColumnHeaders.Item(GuiIndemListView.ExcessAbsorbsDeductible).Width = 0 'hidden 400
        .ColumnHeaders.Item(GuiIndemListView.ExcessAbsorbsDeductible).Alignment = lvwColumnCenter
        'AppliedDeductible
        .ColumnHeaders.Item(GuiIndemListView.AppliedDeductible).Width = 1500
        .ColumnHeaders.Item(GuiIndemListView.AppliedDeductible).Alignment = lvwColumnRight
        'AppliedDeductibleSort
        .ColumnHeaders.Item(GuiIndemListView.AppliedDeductibleSort).Width = 0 'hidden
        .ColumnHeaders.Item(GuiIndemListView.AppliedDeductibleSort).Alignment = lvwColumnLeft
        'ExcessLimits
        .ColumnHeaders.Item(GuiIndemListView.ExcessLimits).Width = 1500
        .ColumnHeaders.Item(GuiIndemListView.ExcessLimits).Alignment = lvwColumnRight
        'ExcessLimitsSort
        .ColumnHeaders.Item(GuiIndemListView.ExcessLimitsSort).Width = 0 'hidden
        .ColumnHeaders.Item(GuiIndemListView.ExcessLimitsSort).Alignment = lvwColumnLeft
        'Miscellaneous
        .ColumnHeaders.Item(GuiIndemListView.Miscellaneous).Width = 1500
        .ColumnHeaders.Item(GuiIndemListView.Miscellaneous).Alignment = lvwColumnRight
        'MiscellaneousSort
        .ColumnHeaders.Item(GuiIndemListView.MiscellaneousSort).Width = 0 'hidden
        .ColumnHeaders.Item(GuiIndemListView.MiscellaneousSort).Alignment = lvwColumnLeft
        'MiscellaneousDesc
        .ColumnHeaders.Item(GuiIndemListView.MiscellaneousDesc).Width = 5000
        .ColumnHeaders.Item(GuiIndemListView.MiscellaneousDesc).Alignment = lvwColumnLeft
        'ACVLessExcessLimits
        .ColumnHeaders.Item(GuiIndemListView.ACVLessExcessLimits).Width = 1700
        .ColumnHeaders.Item(GuiIndemListView.ACVLessExcessLimits).Alignment = lvwColumnRight
        'ACVLessExcessLimitsSort
        .ColumnHeaders.Item(GuiIndemListView.ACVLessExcessLimitsSort).Width = 0 'hidden
        .ColumnHeaders.Item(GuiIndemListView.ACVLessExcessLimitsSort).Alignment = lvwColumnLeft
        'IsPreviousPayment
        .ColumnHeaders.Item(GuiIndemListView.IsPreviousPayment).Width = 400
        .ColumnHeaders.Item(GuiIndemListView.IsPreviousPayment).Alignment = lvwColumnCenter
        'PPayDatePaid
        .ColumnHeaders.Item(GuiIndemListView.PPayDatePaid).Width = 1900
        .ColumnHeaders.Item(GuiIndemListView.PPayDatePaid).Alignment = lvwColumnLeft
        'PPayDatePaidSort
        .ColumnHeaders.Item(GuiIndemListView.PPayDatePaidSort).Width = 0 'hidden
        .ColumnHeaders.Item(GuiIndemListView.PPayDatePaidSort).Alignment = lvwColumnLeft
        'PPayAmountPaid
        .ColumnHeaders.Item(GuiIndemListView.PPayAmountPaid).Width = 1200
        .ColumnHeaders.Item(GuiIndemListView.PPayAmountPaid).Alignment = lvwColumnRight
        'PPayAmountPaidSort
        .ColumnHeaders.Item(GuiIndemListView.PPayAmountPaidSort).Width = 0 'hidden
        .ColumnHeaders.Item(GuiIndemListView.PPayAmountPaidSort).Alignment = lvwColumnRight
        'PPayCheckNumber
        .ColumnHeaders.Item(GuiIndemListView.PPayCheckNumber).Width = 1700
        .ColumnHeaders.Item(GuiIndemListView.PPayCheckNumber).Alignment = lvwColumnLeft
        'PPayCheckNumberSort
        .ColumnHeaders.Item(GuiIndemListView.PPayCheckNumberSort).Width = 0 'hidden
        .ColumnHeaders.Item(GuiIndemListView.PPayCheckNumberSort).Alignment = lvwColumnLeft
        'IsDeleted
        If bHideDeleted Then
            .ColumnHeaders.Item(GuiIndemListView.IsDeleted).Width = 0 'hidden 400
        Else
            .ColumnHeaders.Item(GuiIndemListView.IsDeleted).Width = 400
        End If
        .ColumnHeaders.Item(GuiIndemListView.IsDeleted).Alignment = lvwColumnCenter
        .ColumnHeaders.Item(GuiIndemListView.IsDeleted).Icon = GuiIndemStatusList.IsDeleted
        'UpLoadMe
        If bHideUploadFlags Then
            .ColumnHeaders.Item(GuiIndemListView.UpLoadMe).Width = 0 'hidden 400
        Else
            .ColumnHeaders.Item(GuiIndemListView.UpLoadMe).Width = 400
        End If
        .ColumnHeaders.Item(GuiIndemListView.UpLoadMe).Alignment = lvwColumnCenter
        .ColumnHeaders.Item(GuiIndemListView.UpLoadMe).Icon = GuiIndemStatusList.UpLoadMe
        'DateLastUpdated
        .ColumnHeaders.Item(GuiIndemListView.DateLastUpdated).Width = 2200
        .ColumnHeaders.Item(GuiIndemListView.DateLastUpdated).Alignment = lvwColumnLeft
        'DateLastUpdatedSort
        .ColumnHeaders.Item(GuiIndemListView.DateLastUpdatedSort).Width = 0 'hidden
        .ColumnHeaders.Item(GuiIndemListView.DateLastUpdatedSort).Alignment = lvwColumnLeft
        'AdminComments
        .ColumnHeaders.Item(GuiIndemListView.AdminComments).Width = 0 'hidden 10000
        .ColumnHeaders.Item(GuiIndemListView.AdminComments).Alignment = lvwColumnLeft
        'ClassOfLossID
        .ColumnHeaders.Item(GuiIndemListView.ClassOfLossID).Width = 0 'hidden
        .ColumnHeaders.Item(GuiIndemListView.ClassOfLossID).Alignment = lvwColumnLeft
        'TypeOfLossID
        .ColumnHeaders.Item(GuiIndemListView.TypeOfLossID).Width = 0 'hidden
        .ColumnHeaders.Item(GuiIndemListView.TypeOfLossID).Alignment = lvwColumnLeft
        'ID
        .ColumnHeaders.Item(GuiIndemListView.ID).Width = 0 'hidden
        .ColumnHeaders.Item(GuiIndemListView.ID).Alignment = lvwColumnLeft
        'IDAssignments
        .ColumnHeaders.Item(GuiIndemListView.IDAssignments).Width = 0 'hidden
        .ColumnHeaders.Item(GuiIndemListView.IDAssignments).Alignment = lvwColumnLeft
        'RTIndemnityID
        .ColumnHeaders.Item(GuiIndemListView.RTIndemnityID).Width = 0 'hidden
        .ColumnHeaders.Item(GuiIndemListView.RTIndemnityID).Alignment = lvwColumnLeft
        'AssignmentsID
        .ColumnHeaders.Item(GuiIndemListView.AssignmentsID).Width = 0 'hidden
        .ColumnHeaders.Item(GuiIndemListView.AssignmentsID).Alignment = lvwColumnLeft
        'RTChecksID
        .ColumnHeaders.Item(GuiIndemListView.RTChecksID).Width = 0 'hidden
        .ColumnHeaders.Item(GuiIndemListView.RTChecksID).Alignment = lvwColumnLeft
        'IDRTChecks
        .ColumnHeaders.Item(GuiIndemListView.IDRTChecks).Width = 0 'hidden
        .ColumnHeaders.Item(GuiIndemListView.IDRTChecks).Alignment = lvwColumnLeft
        'DownLoadMe
        .ColumnHeaders.Item(GuiIndemListView.DownLoadMe).Width = 0 'hidden
        .ColumnHeaders.Item(GuiIndemListView.DownLoadMe).Alignment = lvwColumnLeft
        'UpdateByUserID
        .ColumnHeaders.Item(GuiIndemListView.UpdateByUserID).Width = 0 'hidden
        .ColumnHeaders.Item(GuiIndemListView.UpdateByUserID).Alignment = lvwColumnLeft
        
    End With
    
    bGridOn = CBool(GetSetting(App.EXEName, "GENERAL", "GRID_ON", False))
    oListView.GridLines = bGridOn
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub LoadHeaderlvwIndemnity"
End Sub

Public Sub LoadHeaderlvwPayReqs()
    On Error GoTo EH
    Dim bGridOn As Boolean
    Dim bHideDeleted As Boolean
    Dim bHideUploadFlags As Boolean
    
    bHideDeleted = CBool(GetSetting(App.EXEName, "GENERAL", "HIDE_DELETED", True))
    bHideUploadFlags = CBool(GetSetting(App.EXEName, "GENERAL", "HIDE_UPLOAD_FLAGS", True))
    
    'set the columnheaders
    With lvwPayReqs
        .ColumnHeaders.Add , "CheckNum", "No."
        .ColumnHeaders.Add , "CheckNumSort", "CheckNumSort" 'hidden
        .ColumnHeaders.Add , "PrintedDate", "Printed Date" 'hidden
        .ColumnHeaders.Add , "PrintedDateSort", "PrintedDateSort" 'hidden
        .ColumnHeaders.Add , "NoOfRequests", "No. Of Requests"
        .ColumnHeaders.Add , "NoOfRequestsSort", "NoOfRequestsSort" 'hidden
        .ColumnHeaders.Add , "IB", "IB"
        .ColumnHeaders.Add , "PrintOnIB", "Print On IB" 'hidden
        .ColumnHeaders.Add , "ClassOfLoss", "Class"
        .ColumnHeaders.Add , "ClassOfLossCode", "Code"
        .ColumnHeaders.Add , "TypeOfLoss", "Type Of Loss"
        .ColumnHeaders.Add , "TypeOfLossCode", "Code"
        .ColumnHeaders.Add , "RT50_sInsuredPayeeName", "Insured Payee Name(s)"
        .ColumnHeaders.Add , "RT51_sPayeeNames", "Other Payee Name(s)"
        .ColumnHeaders.Add , "RT52_sAddress", "Address"
        .ColumnHeaders.Add , "RT52_sAddressSort", "AddressSort" 'Hidden
        .ColumnHeaders.Add , "RT53_cAmountOfCheck", "Amount Of Check"
        .ColumnHeaders.Add , "RT53_cAmountOfCheckSort", "AmountOfCheckSort" 'Hidden
        .ColumnHeaders.Add , "AppliedDeductible", "App Ded"
        .ColumnHeaders.Add , "AppliedDeductibleSort", "AppliedDeductible" 'Hidden
        .ColumnHeaders.Add , "CatCode", "Cat Code"
        .ColumnHeaders.Add , "IsDeleted", "Is Deleted" ' Hidden
        .ColumnHeaders.Add , "UpLoadMe", "UpLoadMe" ' Hidden
        .ColumnHeaders.Add , "DateLastUpdated", "Date Last Updated"
        .ColumnHeaders.Add , "DateLastUpdatedSort", "DateLastUpdatedSort" ' Hidden
        .ColumnHeaders.Add , "AdminComments", "Admin Comments" ' Hidden
        .ColumnHeaders.Add , "tempCHeckName", "tempCHeckName" 'Hidden
        .ColumnHeaders.Add , "RT42_ClassOfLossID", "RT42_ClassOfLossID" 'Hidden
        .ColumnHeaders.Add , "RT43_TypeOfLossID", "RT43_TypeOfLossID" 'Hidden
        .ColumnHeaders.Add , "RT54_CompanyCatSpecID", "RT54_CompanyCatSpecID" 'Hidden
        .ColumnHeaders.Add , "ID", "ID" ' Hidden
        .ColumnHeaders.Add , "IDAssignments", "IDAssignments" ' Hidden
        .ColumnHeaders.Add , "RTChecksID", "RTChecksID" ' Hidden
        .ColumnHeaders.Add , "AssignmentsID", "AssignmentsID" ' Hidden
        .ColumnHeaders.Add , "BillingCountID", "BillingCountID" ' Hidden
        .ColumnHeaders.Add , "IDBillingCount", "IDBillingCount" 'Hidden
        .ColumnHeaders.Add , "DownLoadMe", "DownLoadMe" ' hidden
        .ColumnHeaders.Add , "UpdateByUserID", "UpdateByUserID" ' Hidden
           
        .Sorted = False
        .SortOrder = lvwAscending
        
        
        'CheckNum
        .ColumnHeaders.Item(GuiPayReqsListView.CheckNum).Width = 600
        .ColumnHeaders.Item(GuiPayReqsListView.CheckNum).Alignment = lvwColumnLeft
        'CheckNumSort
        .ColumnHeaders.Item(GuiPayReqsListView.CheckNumSort).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiPayReqsListView.CheckNumSort).Alignment = lvwColumnLeft
        'PrintedDate
        .ColumnHeaders.Item(GuiPayReqsListView.PrintedDate).Width = 0 'Hidden 2200
        .ColumnHeaders.Item(GuiPayReqsListView.PrintedDate).Alignment = lvwColumnLeft
        'PrintedDateSort
        .ColumnHeaders.Item(GuiPayReqsListView.PrintedDateSort).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiPayReqsListView.PrintedDateSort).Alignment = lvwColumnLeft
        'NoOfRequests
        .ColumnHeaders.Item(GuiPayReqsListView.NoOfRequests).Width = 1700 'Hidden
        .ColumnHeaders.Item(GuiPayReqsListView.NoOfRequests).Alignment = lvwColumnLeft
        'NoOfRequestsSort
        .ColumnHeaders.Item(GuiPayReqsListView.NoOfRequestsSort).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiPayReqsListView.NoOfRequestsSort).Alignment = lvwColumnLeft
        'IB
        .ColumnHeaders.Item(GuiPayReqsListView.IB).Width = 1200
        .ColumnHeaders.Item(GuiPayReqsListView.IB).Alignment = lvwColumnLeft
        'PrintOnIB
        .ColumnHeaders.Item(GuiPayReqsListView.PrintOnIB).Width = 0 'Hidden 400
        .ColumnHeaders.Item(GuiPayReqsListView.PrintOnIB).Alignment = lvwColumnLeft
        'ClassOfLoss
        .ColumnHeaders.Item(GuiPayReqsListView.ClassOfLoss).Width = 3000
        .ColumnHeaders.Item(GuiPayReqsListView.ClassOfLoss).Alignment = lvwColumnLeft
        'ClassOfLossCode
        .ColumnHeaders.Item(GuiPayReqsListView.ClassOfLossCode).Width = 700
        .ColumnHeaders.Item(GuiPayReqsListView.ClassOfLossCode).Alignment = lvwColumnLeft
        'TypeOfLoss
        .ColumnHeaders.Item(GuiPayReqsListView.TypeOfLoss).Width = 3000
        .ColumnHeaders.Item(GuiPayReqsListView.TypeOfLoss).Alignment = lvwColumnLeft
        'TypeOfLossCode
        .ColumnHeaders.Item(GuiPayReqsListView.TypeOfLossCode).Width = 700
        .ColumnHeaders.Item(GuiPayReqsListView.TypeOfLossCode).Alignment = lvwColumnLeft
        'RT50_sInsuredPayeeName
        .ColumnHeaders.Item(GuiPayReqsListView.RT50_sInsuredPayeeName).Width = 3000
        .ColumnHeaders.Item(GuiPayReqsListView.RT50_sInsuredPayeeName).Alignment = lvwColumnLeft
        'RT51_sPayeeNames
        .ColumnHeaders.Item(GuiPayReqsListView.RT51_sPayeeNames).Width = 3000
        .ColumnHeaders.Item(GuiPayReqsListView.RT51_sPayeeNames).Alignment = lvwColumnLeft
        'RT52_sAddress
        .ColumnHeaders.Item(GuiPayReqsListView.RT52_sAddress).Width = 6000
        .ColumnHeaders.Item(GuiPayReqsListView.RT52_sAddress).Alignment = lvwColumnLeft
        'RT52_sAddressSort
        .ColumnHeaders.Item(GuiPayReqsListView.RT52_sAddressSort).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiPayReqsListView.RT52_sAddressSort).Alignment = lvwColumnLeft
        'RT53_cAmountOfCheck
        .ColumnHeaders.Item(GuiPayReqsListView.RT53_cAmountOfCheck).Width = 2000
        .ColumnHeaders.Item(GuiPayReqsListView.RT53_cAmountOfCheck).Alignment = lvwColumnRight
        'RT53_cAmountOfCheckSort
        .ColumnHeaders.Item(GuiPayReqsListView.RT53_cAmountOfCheckSort).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiPayReqsListView.RT53_cAmountOfCheckSort).Alignment = lvwColumnLeft
        'AppliedDeductible
        .ColumnHeaders.Item(GuiPayReqsListView.AppliedDeductible).Width = 2000
        .ColumnHeaders.Item(GuiPayReqsListView.AppliedDeductible).Alignment = lvwColumnRight
        'AppliedDeductibleSort
        .ColumnHeaders.Item(GuiPayReqsListView.AppliedDeductibleSort).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiPayReqsListView.AppliedDeductibleSort).Alignment = lvwColumnLeft
        'CatCode
        .ColumnHeaders.Item(GuiPayReqsListView.CatCode).Width = 1000
        .ColumnHeaders.Item(GuiPayReqsListView.CatCode).Alignment = lvwColumnLeft
        'IsDeleted
        If bHideDeleted Then
            .ColumnHeaders.Item(GuiPayReqsListView.IsDeleted).Width = 0 'Hidden 400
        Else
            .ColumnHeaders.Item(GuiPayReqsListView.IsDeleted).Width = 400
        End If
        .ColumnHeaders.Item(GuiPayReqsListView.IsDeleted).Alignment = lvwColumnLeft
        .ColumnHeaders.Item(GuiPayReqsListView.IsDeleted).Icon = GuiPayReqsStatusList.IsDeleted
        'UpLoadMe
        If bHideUploadFlags Then
            .ColumnHeaders.Item(GuiPayReqsListView.UpLoadMe).Width = 0 'Hidden 400
        Else
            .ColumnHeaders.Item(GuiPayReqsListView.UpLoadMe).Width = 400
        End If
        .ColumnHeaders.Item(GuiPayReqsListView.UpLoadMe).Alignment = lvwColumnLeft
        .ColumnHeaders.Item(GuiPayReqsListView.UpLoadMe).Icon = GuiPayReqsStatusList.UpLoadMe
        'DateLastUpdated
        .ColumnHeaders.Item(GuiPayReqsListView.DateLastUpdated).Width = 2200
        .ColumnHeaders.Item(GuiPayReqsListView.DateLastUpdated).Alignment = lvwColumnLeft
        'DateLastUpdatedSort
        .ColumnHeaders.Item(GuiPayReqsListView.DateLastUpdatedSort).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiPayReqsListView.DateLastUpdatedSort).Alignment = lvwColumnLeft
        'AdminComments
        .ColumnHeaders.Item(GuiPayReqsListView.AdminComments).Width = 0 'Hidden 10000
        .ColumnHeaders.Item(GuiPayReqsListView.AdminComments).Alignment = lvwColumnLeft
        'tempCHeckName
        .ColumnHeaders.Item(GuiPayReqsListView.tempCHeckName).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiPayReqsListView.tempCHeckName).Alignment = lvwColumnLeft
        'RT42_ClassOfLossID
        .ColumnHeaders.Item(GuiPayReqsListView.RT42_ClassOfLossID).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiPayReqsListView.RT42_ClassOfLossID).Alignment = lvwColumnLeft
        'RT43_TypeOfLossID
        .ColumnHeaders.Item(GuiPayReqsListView.RT43_TypeOfLossID).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiPayReqsListView.RT43_TypeOfLossID).Alignment = lvwColumnLeft
        'RT54_CompanyCatSpecID
        .ColumnHeaders.Item(GuiPayReqsListView.RT54_CompanyCatSpecID).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiPayReqsListView.RT54_CompanyCatSpecID).Alignment = lvwColumnLeft
        'ID
        .ColumnHeaders.Item(GuiPayReqsListView.ID).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiPayReqsListView.ID).Alignment = lvwColumnLeft
        'IDAssignments
        .ColumnHeaders.Item(GuiPayReqsListView.IDAssignments).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiPayReqsListView.IDAssignments).Alignment = lvwColumnLeft
        'RTChecksID
        .ColumnHeaders.Item(GuiPayReqsListView.RTChecksID).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiPayReqsListView.RTChecksID).Alignment = lvwColumnLeft
        'AssignmentsID
        .ColumnHeaders.Item(GuiPayReqsListView.AssignmentsID).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiPayReqsListView.AssignmentsID).Alignment = lvwColumnLeft
        'BillingCountID
        .ColumnHeaders.Item(GuiPayReqsListView.BillingCountID).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiPayReqsListView.BillingCountID).Alignment = lvwColumnLeft
        'IDBillingCount
        .ColumnHeaders.Item(GuiPayReqsListView.IDBillingCount).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiPayReqsListView.IDBillingCount).Alignment = lvwColumnLeft
        'DownLoadMe
        .ColumnHeaders.Item(GuiPayReqsListView.DownLoadMe).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiPayReqsListView.DownLoadMe).Alignment = lvwColumnLeft
        'UpdateByUserID
        .ColumnHeaders.Item(GuiPayReqsListView.UpdateByUserID).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiPayReqsListView.UpdateByUserID).Alignment = lvwColumnLeft
        
    End With
    
    bGridOn = CBool(GetSetting(App.EXEName, "GENERAL", "GRID_ON", False))
    lvwPayReqs.GridLines = bGridOn
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub LoadHeaderlvwPayReqs"
End Sub

Public Function LoadMe() As Boolean
    On Error GoTo EH
    
    mbLoadingMe = True
    
    RefreshDed
    RefreshIndemnity
    PopulateAppliedDeductible
    RefreshIndemnity
    RefreshPayReqs
    
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
    
 
    
    cmdSave.Enabled = False
    SaveMe = True
    
    'cleanup
    Set oConn = Nothing
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
    TSIndem.Width = Me.Width - 405
    
    'Fram Deductible
    framDeductible.Width = Me.Width - 645
    framDedMaint.Width = Me.Width - 645
    cboAddPLAppClassTypeID.Width = Me.Width - 3525
    lvwDed.Width = Me.Width - 2445
    cmdUp.left = Me.Width - 2220
    cmdDown.left = Me.Width - 2220
    CmdReNumberSort.left = Me.Width - 1620
    cmdDelDed.left = Me.Width - 2220
    
    'PayReq
    framPayReqs.Width = Me.Width - 645
    lvwPayReqs.Width = Me.Width - 945
    framPayReqsMaint.Width = Me.Width - 945
    chkHideDeleted(1).left = Me.Width - 3780
    cmdDelPayReqs.left = Me.Width - 2580
    
    
    'Indem
    framIndemnity.Width = Me.Width - 645
    lvwIndemnity.Width = Me.Width - 945
    framIndemTotals.Width = Me.Width - 945
    lblPayReqs.left = Me.Width - 4260
    txtPayReqsValue.left = Me.Width - 2460
    lblAmountWarning.left = Me.Width - 2500
    framIndemnityMaint.Width = Me.Width - 945
    chkHideDeleted(0).left = Me.Width - 3780
    cmdDelIndemnity.left = Me.Width - 2580
    
    'framCommands
    framCommands.left = Me.Width - 3660
    
    
    'Heights and Tops
    TSIndem.Height = Me.Height - 1785
    
    'fram Deductible
    framDeductible.Height = Me.Height - 2265
    lvwDed.Height = Me.Height - 3225
    
    'PayReq
    framPayReqs.Height = Me.Height - 2265
    lvwPayReqs.Height = Me.Height - 3825
    framPayReqsMaint.top = Me.Height - 3120
    framAssocBillingID.top = Me.Height - 1680
    
    'Indem
    framIndemnity.Height = Me.Height - 2265
    lvwIndemnity.Height = Me.Height - 5025
    framIndemTotals.top = Me.Height - 4440
    framIndemnityMaint.top = Me.Height - 3120
    framAssocRTChecksID.top = Me.Height - 1680
    
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
    
    If Not moForm Is Nothing Then
        Unload moForm
        Set moForm = Nothing
    End If
    
    If Not mArv Is Nothing Then
        mArv.CLEANUP
        Set mArv = Nothing
    End If
    
    Set mcolSelIndemID = Nothing
    
    CLEANUP = True
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function CLEANUP"
End Function





Private Sub lvwIndemnity_DblClick()
    On Error GoTo EH
    If Not lvwIndemnity.SelectedItem Is Nothing Then
        EditIndemnity
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lvwIndemnity_DblClick"
End Sub

Private Sub lvwIndemnity_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo EH
    Select Case KeyCode
        Case KeyCodeConstants.vbKeyReturn
            EditIndemnity
        Case vbKeyDelete
            cmdDelIndemnity_Click
    End Select
Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lvwIndemnity_KeyDown"
End Sub

Private Sub lvwIndemnity_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo EH
    If Button = vbRightButton Then
        PopupMenu PopUpMnuIndem
    End If
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lvwIndemnity_MouseUp"
End Sub

Private Sub lvwPayReqs_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    On Error GoTo EH
    Dim sMess As String
    If lvwPayReqs.SortOrder = lvwAscending Then
        lvwPayReqs.SortOrder = lvwDescending
        sMess = "Descending"
    Else
        lvwPayReqs.SortOrder = lvwAscending
        sMess = "Ascending"
    End If
    
    'Set the Tool Tip
    lvwPayReqs.ToolTipText = "Sort By " & ColumnHeader.Text & " " & sMess
    
    'Need to see if the Column clicked was a NON text sort
    'Like Date or number.  If a non textual column was clicked
    'need to use the next column as the sort since this next column
    'will be hidden and contain a sort friendly format.
    'Sort Key is Base 0 where CoulmnHeader is not
    Select Case ColumnHeader.Index
        Case GuiPayReqsListView.DateLastUpdated, GuiPayReqsListView.CheckNum, GuiPayReqsListView.RT52_sAddress, GuiPayReqsListView.RT53_cAmountOfCheck, GuiPayReqsListView.AppliedDeductible
            lvwPayReqs.SortKey = ColumnHeader.Index
        Case GuiPayReqsListView.NoOfRequests
            lvwPayReqs.SortKey = ColumnHeader.Index
        Case Else
            lvwPayReqs.SortKey = ColumnHeader.Index - 1
    End Select
    
    lvwPayReqs.Sorted = True
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lvwPayReqs_ColumnClick"
End Sub

Private Sub lvwIndemnity_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    On Error GoTo EH
    Dim sMess As String
    If lvwIndemnity.SortOrder = lvwAscending Then
        lvwIndemnity.SortOrder = lvwDescending
        sMess = "Descending"
    Else
        lvwIndemnity.SortOrder = lvwAscending
        sMess = "Ascending"
    End If
    
    'Set the Tool Tip
    lvwIndemnity.ToolTipText = "Sort By " & ColumnHeader.Text & " " & sMess
    
    'Need to see if the Column clicked was a NON text sort
    'Like Date or number.  If a non textual column was clicked
    'need to use the next column as the sort since this next column
    'will be hidden and contain a sort friendly format.
    'Sort Key is Base 0 where CoulmnHeader is not
    Select Case ColumnHeader.Index
        Case GuiIndemListView.DateLastUpdated
            lvwIndemnity.SortKey = ColumnHeader.Index
        Case GuiIndemListView.ACVClaim
            lvwIndemnity.SortKey = ColumnHeader.Index
        Case GuiIndemListView.ACVLessExcessLimits
            lvwIndemnity.SortKey = ColumnHeader.Index
        Case GuiIndemListView.ExcessLimits
            lvwIndemnity.SortKey = ColumnHeader.Index
        Case GuiIndemListView.NonRecoverableDep
            lvwIndemnity.SortKey = ColumnHeader.Index
        Case GuiIndemListView.RecoverableDep
            lvwIndemnity.SortKey = ColumnHeader.Index
        Case GuiIndemListView.ReplacementCost
            lvwIndemnity.SortKey = ColumnHeader.Index
        Case GuiIndemListView.SpecialLimits
            lvwIndemnity.SortKey = ColumnHeader.Index
        Case Else
            lvwIndemnity.SortKey = ColumnHeader.Index - 1
    End Select
    
    lvwIndemnity.Sorted = True
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lvwIndemnity_ColumnClick"
End Sub

Private Sub lvwPayReqs_DblClick()
    On Error GoTo EH
    If Not lvwPayReqs.SelectedItem Is Nothing Then
        EditPayReqs
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lvwPayReqs_DblClick"
End Sub

Private Sub lvwPayReqs_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo EH
    Select Case KeyCode
        Case KeyCodeConstants.vbKeyReturn
            EditPayReqs
        Case vbKeyDelete
            cmdDelPayReqs_Click
    End Select
Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lvwPayReqs_KeyDown"
End Sub

Private Sub lvwPayReqs_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo EH
    If Button = vbRightButton Then
        PopupMenu PopUpmnuPayReqs
    End If
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lvwPayReqs_MouseUp"
End Sub

Private Sub mnuDeleteIndem_Click()
    On Error GoTo EH
    Dim itmX As MSComctlLib.ListItem
    Dim sIndemID As String
    
    Set itmX = lvwIndemnity.SelectedItem
    
    If Not itmX Is Nothing Then
        sIndemID = itmX.SubItems(GuiIndemListView.ID - 1)
        If MsgBox("Are you sure you want to delete this Indemnity Item?", vbYesNo, "DELETE SELECTED ITEM") = vbYes Then
            If DeleteIndemnityItem(sIndemID) Then
                lvwIndemnity.ListItems.Remove ("""" & sIndemID & """")
            End If
            mfrmClaim.RefreshMe
            lvwIndemnity.SetFocus
        End If
    End If
    
    Set itmX = Nothing
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub mnuDeleteIndem_Click"
End Sub

Private Sub mnuDeletePayReqs_Click()
    On Error GoTo EH
    Dim itmX As MSComctlLib.ListItem
    Dim sPayReqID As String
    
    Set itmX = lvwPayReqs.SelectedItem
    
    If Not itmX Is Nothing Then
        sPayReqID = itmX.SubItems(GuiPayReqsListView.ID - 1)
        If MsgBox("Are you sure you want to delete this Payment Request Item?", vbYesNo, "DELETE SELECTED ITEM") = vbYes Then
            If DeletePayReqItem(sPayReqID) Then
                lvwPayReqs.ListItems.Remove ("""" & sPayReqID & """")
            End If
            mfrmClaim.RefreshMe
            lvwPayReqs.SetFocus
        End If
    End If
    
    Set itmX = Nothing
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub mnuDeletePayReqs_Click"
End Sub

Private Sub mnuEditIndem_Click()
    On Error GoTo EH
    
    EditIndemnity
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub mnuEditIndem_Click"
End Sub

Private Sub mnuEditPayReqs_Click()
    On Error GoTo EH
    
    EditPayReqs
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub mnuEditPayReqs_Click"
End Sub

Private Sub mnuSelectAllIndem_Click()
    On Error GoTo EH
    Dim itmX As ListItem
    
    For Each itmX In lvwIndemnity.ListItems
        itmX.Selected = True
    Next
    
    Set itmX = Nothing
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub mnuSelectAllIndem_Click"
End Sub

Private Sub mnuSelectAllPayReqs_Click()
    On Error GoTo EH
    Dim itmX As ListItem
    
    For Each itmX In lvwPayReqs.ListItems
        itmX.Selected = True
    Next
    
    Set itmX = Nothing
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub mnuSelectAllPayReqs_Click"
End Sub

Private Sub TSIndem_Click()
    ShowFrame
End Sub

Public Function ShowFrame() As Boolean
    On Error GoTo EH
    Dim sFrameName As String
    Dim oFrame As Control
    Dim MyFrame As Frame
    Dim oControl As Control
    
    sFrameName = TSIndem.SelectedItem.Tag
    
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
            ElseIf StrComp(MyFrame.Name, framDedMaint.Name, vbTextCompare) = 0 Then
                MyFrame.Visible = True
            ElseIf StrComp(MyFrame.Name, framPayReqsMaint.Name, vbTextCompare) = 0 Then
                MyFrame.Visible = True
             ElseIf StrComp(MyFrame.Name, framIndemTotals.Name, vbTextCompare) = 0 Then
                MyFrame.Visible = True
            ElseIf StrComp(MyFrame.Name, framIndemnityMaint.Name, vbTextCompare) = 0 Then
                MyFrame.Visible = True
            ElseIf StrComp(MyFrame.Name, framCommands.Name, vbTextCompare) = 0 Then
                MyFrame.Visible = True
            ElseIf StrComp(sFrameName, framPayReqs.Name, vbTextCompare) = 0 And StrComp(MyFrame.Name, framAssocBillingID.Name, vbTextCompare) = 0 Then
                MyFrame.Visible = True
            ElseIf StrComp(sFrameName, framIndemnity.Name, vbTextCompare) = 0 And StrComp(MyFrame.Name, framAssocRTChecksID.Name, vbTextCompare) = 0 Then
                MyFrame.Visible = True
            Else
                MyFrame.Visible = False
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

Public Sub PopulatelvwIndemnity(poLvw As ListView)
    On Error GoTo EH
    Dim iMyIcon As Long
    Dim sFlagText As String
    Dim sTemp As String
    Dim itmX As ListItem
    Dim oListView As MSComctlLib.ListView
    Dim RS As ADODB.Recordset
    
    'Clear the List view
    Set oListView = poLvw

    oListView.Visible = False
    oListView.ListItems.Clear
    oListView.Sorted = False
    
'    If Not mfrmClaim.SetadoRSRTIndemnity(msAssignmentsID) Then
'        Exit Sub
'    End If
    Set RS = mfrmClaim.adoRSRTIndemnity

    If Not RS.EOF Then
        RS.MoveFirst
        Do Until RS.EOF
            'PaymentRequest
            Set itmX = oListView.ListItems.Add(, """" & goUtil.IsNullIsVbNullString(RS.Fields("ID")) & """", goUtil.IsNullIsVbNullString(RS.Fields("PaymentRequest")))
            'ClassOfLoss
            itmX.SubItems(GuiIndemListView.ClassOfLoss - 1) = goUtil.IsNullIsVbNullString(RS.Fields("ClassOfLoss"))
            'ClassOfLossCode
            itmX.SubItems(GuiIndemListView.ClassOfLossCode - 1) = goUtil.IsNullIsVbNullString(RS.Fields("ClassOfLossCode"))
            'TypeOfLoss
            itmX.SubItems(GuiIndemListView.TypeOfLoss - 1) = goUtil.IsNullIsVbNullString(RS.Fields("TypeOfLoss"))
            'TypeOfLossCode
            itmX.SubItems(GuiIndemListView.TypeOfLossCode - 1) = goUtil.IsNullIsVbNullString(RS.Fields("TypeOfLossCode"))
            'Description
            itmX.SubItems(GuiIndemListView.Description - 1) = goUtil.IsNullIsVbNullString(RS.Fields("Description"))
            'ReplacementCost
            itmX.SubItems(GuiIndemListView.ReplacementCost - 1) = Format(goUtil.IsNullIsVbNullString(RS.Fields("ReplacementCost")), "#,###,###,##0.00")
            'ReplacementCostSort 'Hidden
            itmX.SubItems(GuiIndemListView.ReplacementCostSort - 1) = goUtil.utNumInTextSortFormat(Format(goUtil.IsNullIsVbNullString(RS.Fields("ReplacementCost")), "#,###,###,##0.00"))
            'RecoverableDep
            itmX.SubItems(GuiIndemListView.RecoverableDep - 1) = Format(goUtil.IsNullIsVbNullString(RS.Fields("RecoverableDep")), "#,###,###,##0.00")
            'RecoverableDepSort 'Hidden
            itmX.SubItems(GuiIndemListView.RecoverableDepSort - 1) = goUtil.utNumInTextSortFormat(Format(goUtil.IsNullIsVbNullString(RS.Fields("RecoverableDep")), "#,###,###,##0.00"))
            'NonRecoverableDep
            itmX.SubItems(GuiIndemListView.NonRecoverableDep - 1) = Format(goUtil.IsNullIsVbNullString(RS.Fields("NonRecoverableDep")), "#,###,###,##0.00")
            'NonRecoverableDepSort 'Hidden
            itmX.SubItems(GuiIndemListView.NonRecoverableDepSort - 1) = goUtil.utNumInTextSortFormat(Format(goUtil.IsNullIsVbNullString(RS.Fields("NonRecoverableDep")), "#,###,###,##0.00"))
            'ACVClaim
            itmX.SubItems(GuiIndemListView.ACVClaim - 1) = Format(goUtil.IsNullIsVbNullString(RS.Fields("ACVClaim")), "#,###,###,##0.00")
            'ACVClaimSort 'Hidden
            itmX.SubItems(GuiIndemListView.ACVClaimSort - 1) = goUtil.utNumInTextSortFormat(Format(goUtil.IsNullIsVbNullString(RS.Fields("ACVClaim")), "#,###,###,##0.00"))
            'SpecialLimits
            itmX.SubItems(GuiIndemListView.SpecialLimits - 1) = Format(goUtil.IsNullIsVbNullString(RS.Fields("SpecialLimits")), "#,###,###,##0.00")
            'SpecialLimitsSort 'Hidden
            itmX.SubItems(GuiIndemListView.SpecialLimitsSort - 1) = goUtil.utNumInTextSortFormat(Format(goUtil.IsNullIsVbNullString(RS.Fields("SpecialLimits")), "#,###,###,##0.00"))
            'IsAddAmountOfInsurance
            If CBool(RS.Fields("IsAddAmountOfInsurance")) Then
                iMyIcon = GuiIndemStatusList.IsChecked
            Else
                iMyIcon = Empty
            End If
            sFlagText = goUtil.GetFlagText(RS.Fields("IsAddAmountOfInsurance"))
            itmX.SubItems(GuiIndemListView.IsAddAmountOfInsurance - 1) = sFlagText
            itmX.ListSubItems(GuiIndemListView.IsAddAmountOfInsurance - 1).ReportIcon = iMyIcon
            'ExcessAbsorbsDeductible
            If CBool(RS.Fields("ExcessAbsorbsDeductible")) Then
                iMyIcon = GuiIndemStatusList.IsChecked
            Else
                iMyIcon = Empty
            End If
            sFlagText = goUtil.GetFlagText(RS.Fields("ExcessAbsorbsDeductible"))
            itmX.SubItems(GuiIndemListView.ExcessAbsorbsDeductible - 1) = sFlagText
            itmX.ListSubItems(GuiIndemListView.ExcessAbsorbsDeductible - 1).ReportIcon = iMyIcon
            'AppliedDeductible
            itmX.SubItems(GuiIndemListView.AppliedDeductible - 1) = Format(goUtil.IsNullIsVbNullString(RS.Fields("AppliedDeductible")), "#,###,###,##0.00")
            'AppliedDeductibleSort
            itmX.SubItems(GuiIndemListView.AppliedDeductibleSort - 1) = goUtil.utNumInTextSortFormat(Format(goUtil.IsNullIsVbNullString(RS.Fields("AppliedDeductible")), "#,###,###,##0.00"))
            'ExcessLimits
            itmX.SubItems(GuiIndemListView.ExcessLimits - 1) = Format(goUtil.IsNullIsVbNullString(RS.Fields("ExcessLimits")), "#,###,###,##0.00")
            'ExcessLimitsSort 'Hidden
            itmX.SubItems(GuiIndemListView.ExcessLimitsSort - 1) = goUtil.utNumInTextSortFormat(Format(goUtil.IsNullIsVbNullString(RS.Fields("ExcessLimits")), "#,###,###,##0.00"))
            'Miscellaneous
            itmX.SubItems(GuiIndemListView.Miscellaneous - 1) = Format(goUtil.IsNullIsVbNullString(RS.Fields("Miscellaneous")), "#,###,###,##0.00")
            'MiscellaneousSort 'Hidden
            itmX.SubItems(GuiIndemListView.MiscellaneousSort - 1) = goUtil.utNumInTextSortFormat(Format(goUtil.IsNullIsVbNullString(RS.Fields("Miscellaneous")), "#,###,###,##0.00"))
            'MiscellaneousDesc
            itmX.SubItems(GuiIndemListView.MiscellaneousDesc - 1) = goUtil.IsNullIsVbNullString(RS.Fields("MiscDescription"))
            'ACVLessExcessLimits
            itmX.SubItems(GuiIndemListView.ACVLessExcessLimits - 1) = Format(goUtil.IsNullIsVbNullString(RS.Fields("ACVLessExcessLimits")), "#,###,###,##0.00")
            'ACVLessExcessLimitsSort 'Hidden
            itmX.SubItems(GuiIndemListView.ACVLessExcessLimitsSort - 1) = goUtil.utNumInTextSortFormat(Format(goUtil.IsNullIsVbNullString(RS.Fields("ACVLessExcessLimits")), "#,###,###,##0.00"))
            'IsPreviousPayment
            If CBool(RS.Fields("IsPreviousPayment")) Then
                iMyIcon = GuiIndemStatusList.IsChecked
            Else
                iMyIcon = Empty
            End If
            sFlagText = goUtil.GetFlagText(RS.Fields("IsPreviousPayment"))
            itmX.SubItems(GuiIndemListView.IsPreviousPayment - 1) = sFlagText
            itmX.ListSubItems(GuiIndemListView.IsPreviousPayment - 1).ReportIcon = iMyIcon
            'PPayDatePaid
            If Not IsNull(RS.Fields("PPayDatePaid").Value) Then
                If IsDate(RS.Fields("PPayDatePaid").Value) Then
                    itmX.SubItems(GuiIndemListView.PPayDatePaid - 1) = Format(RS.Fields("PPayDatePaid").Value, "MM/DD/YYYY")
                    itmX.SubItems(GuiIndemListView.PPayDatePaidSort - 1) = Format(RS.Fields("PPayDatePaid").Value, "YYYY/MM/DD")
                Else
                    itmX.SubItems(GuiIndemListView.PPayDatePaid - 1) = vbNullString
                    itmX.SubItems(GuiIndemListView.PPayDatePaidSort - 1) = vbNullString
                End If
            Else
                itmX.SubItems(GuiIndemListView.PPayDatePaid - 1) = vbNullString
                itmX.SubItems(GuiIndemListView.PPayDatePaidSort - 1) = vbNullString
            End If
            'PPayAmountPaid
            itmX.SubItems(GuiIndemListView.PPayAmountPaid - 1) = Format(goUtil.IsNullIsVbNullString(RS.Fields("PPayAmountPaid")), "#,###,###,##0.00")
            'PPayAmountPaidSort 'hidden
            itmX.SubItems(GuiIndemListView.PPayAmountPaidSort - 1) = goUtil.utNumInTextSortFormat(Format(goUtil.IsNullIsVbNullString(RS.Fields("PPayAmountPaid")), "#,###,###,##0.00"))
            'PPayCheckNumber
            itmX.SubItems(GuiIndemListView.PPayCheckNumber - 1) = goUtil.IsNullIsVbNullString(RS.Fields("PPayCheckNumber"))
            'PPayCheckNumberSort 'Hidden
            itmX.SubItems(GuiIndemListView.PPayCheckNumberSort - 1) = goUtil.utNumInTextSortFormat(goUtil.IsNullIsVbNullString(RS.Fields("PPayCheckNumber")))
            'IsDeleted
            If CBool(RS.Fields("IsDeleted")) Then
                iMyIcon = GuiIndemStatusList.IsDeleted
            Else
                iMyIcon = Empty
            End If
            sFlagText = goUtil.GetFlagText(RS.Fields("IsDeleted"))
            itmX.SubItems(GuiIndemListView.IsDeleted - 1) = sFlagText
            itmX.ListSubItems(GuiIndemListView.IsDeleted - 1).ReportIcon = iMyIcon
            'UpLoadMe
            If CBool(RS.Fields("UpLoadMe")) Then
                iMyIcon = GuiIndemStatusList.UpLoadMe
            Else
                iMyIcon = Empty
            End If
            sFlagText = goUtil.GetFlagText(RS.Fields("UpLoadMe"))
            itmX.SubItems(GuiIndemListView.UpLoadMe - 1) = sFlagText
            itmX.ListSubItems(GuiIndemListView.UpLoadMe - 1).ReportIcon = iMyIcon
            'DateLastUpdated
            If Not IsNull(RS.Fields("DateLastUpdated").Value) Then
                If IsDate(RS.Fields("DateLastUpdated").Value) Then
                    itmX.SubItems(GuiIndemListView.DateLastUpdated - 1) = Format(RS.Fields("DateLastUpdated").Value, "MM/DD/YYYY HH:MM:SS")
                    itmX.SubItems(GuiIndemListView.DateLastUpdatedSort - 1) = Format(RS.Fields("DateLastUpdated").Value, "YYYY/MM/DD HH:MM:SS")
                Else
                    itmX.SubItems(GuiIndemListView.DateLastUpdated - 1) = vbNullString
                    itmX.SubItems(GuiIndemListView.DateLastUpdatedSort - 1) = vbNullString
                End If
            Else
                itmX.SubItems(GuiIndemListView.DateLastUpdated - 1) = vbNullString
                itmX.SubItems(GuiIndemListView.DateLastUpdatedSort - 1) = vbNullString
            End If
            
            'AdminComments
            itmX.SubItems(GuiIndemListView.AdminComments - 1) = goUtil.IsNullIsVbNullString(RS.Fields("AdminComments"))
            'ClassOfLossID 'Hidden
            itmX.SubItems(GuiIndemListView.ClassOfLossID - 1) = goUtil.IsNullIsVbNullString(RS.Fields("ClassOfLossID"))
            'TypeOfLossID 'Hidden
            itmX.SubItems(GuiIndemListView.TypeOfLossID - 1) = goUtil.IsNullIsVbNullString(RS.Fields("TypeOfLossID"))
            'ID 'Hidden
            itmX.SubItems(GuiIndemListView.ID - 1) = goUtil.IsNullIsVbNullString(RS.Fields("ID"))
            'IDAssignments 'Hidden
            itmX.SubItems(GuiIndemListView.IDAssignments - 1) = goUtil.IsNullIsVbNullString(RS.Fields("IDAssignments"))
            'RTIndemnityID 'Hidden
            itmX.SubItems(GuiIndemListView.RTIndemnityID - 1) = goUtil.IsNullIsVbNullString(RS.Fields("RTIndemnityID"))
            'AssignmentsID 'Hidden
            itmX.SubItems(GuiIndemListView.AssignmentsID - 1) = goUtil.IsNullIsVbNullString(RS.Fields("AssignmentsID"))
            'RTChecksID 'Hidden
            itmX.SubItems(GuiIndemListView.RTChecksID - 1) = goUtil.IsNullIsVbNullString(RS.Fields("RTChecksID"))
            'IDRTChecks 'Hidden
            itmX.SubItems(GuiIndemListView.IDRTChecks - 1) = goUtil.IsNullIsVbNullString(RS.Fields("IDRTChecks"))
            'DownLoadMe 'Hidden
            itmX.SubItems(GuiIndemListView.DownLoadMe - 1) = goUtil.IsNullIsVbNullString(RS.Fields("DownLoadMe"))
            'UpdateByUserID 'Hidden
            itmX.SubItems(GuiIndemListView.UpdateByUserID - 1) = goUtil.IsNullIsVbNullString(RS.Fields("UpdateByUserID"))
        
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
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub PopulatelvwIndemnity"
    oListView.Visible = True
End Sub


Private Sub PopulatelvwPayReqs()
    On Error GoTo EH
    Dim iMyIcon As Long
    Dim sFlagText As String
    Dim sTemp As String
    Dim itmX As ListItem
    Dim oListView As MSComctlLib.ListView
    Dim RS As ADODB.Recordset
    
    'Clear the List view
    Set oListView = lvwPayReqs

    oListView.Visible = False
    oListView.ListItems.Clear

    Set RS = mfrmClaim.adoRSRTChecks

    If Not RS.EOF Then
        RS.MoveFirst
        Do Until RS.EOF
            'CheckNum
             Set itmX = oListView.ListItems.Add(, """" & goUtil.IsNullIsVbNullString(RS.Fields("ID")) & """", goUtil.IsNullIsVbNullString(RS.Fields("CheckNum")))
            'CheckNumSort 'hidden
            itmX.SubItems(GuiPayReqsListView.CheckNumSort - 1) = goUtil.utNumInTextSortFormat(goUtil.IsNullIsVbNullString(RS.Fields("CheckNum")))
            If Not IsNull(RS.Fields("PrintedDate").Value) Then
                If IsDate(RS.Fields("PrintedDate").Value) Then
                    itmX.SubItems(GuiPayReqsListView.PrintedDate - 1) = Format(RS.Fields("PrintedDate").Value, "MM/DD/YYYY HH:MM:SS")
                    itmX.SubItems(GuiPayReqsListView.PrintedDateSort - 1) = Format(RS.Fields("PrintedDate").Value, "YYYY/MM/DD HH:MM:SS")
                Else
                    itmX.SubItems(GuiPayReqsListView.PrintedDate - 1) = vbNullString
                    itmX.SubItems(GuiPayReqsListView.PrintedDateSort - 1) = vbNullString
                End If
            Else
                itmX.SubItems(GuiPayReqsListView.PrintedDate - 1) = vbNullString
                itmX.SubItems(GuiPayReqsListView.PrintedDateSort - 1) = vbNullString
            End If
            'IB
            itmX.SubItems(GuiPayReqsListView.IB - 1) = goUtil.IsNullIsVbNullString(RS.Fields("IB"))
            'NoOfRequests
            itmX.SubItems(GuiPayReqsListView.NoOfRequests - 1) = goUtil.IsNullIsVbNullString(RS.Fields("NoOfRequests"))
            'NoOfRequestsSort
            itmX.SubItems(GuiPayReqsListView.NoOfRequestsSort - 1) = goUtil.utNumInTextSortFormat(goUtil.IsNullIsVbNullString(RS.Fields("NoOfRequests")))
            'PrintOnIB
            If CBool(RS.Fields("PrintOnIB")) Then
                iMyIcon = GuiPayReqsStatusList.IsChecked
            Else
                iMyIcon = Empty
            End If
            sFlagText = goUtil.GetFlagText(RS.Fields("PrintOnIB"))
            itmX.SubItems(GuiPayReqsListView.PrintOnIB - 1) = sFlagText
            itmX.ListSubItems(GuiPayReqsListView.PrintOnIB - 1).ReportIcon = iMyIcon
            'ClassOfLoss
            itmX.SubItems(GuiPayReqsListView.ClassOfLoss - 1) = goUtil.IsNullIsVbNullString(RS.Fields("ClassOfLoss"))
            'ClassOfLossCode
            itmX.SubItems(GuiPayReqsListView.ClassOfLossCode - 1) = goUtil.IsNullIsVbNullString(RS.Fields("ClassOfLossCode"))
            'TypeOfLoss
            itmX.SubItems(GuiPayReqsListView.TypeOfLoss - 1) = goUtil.IsNullIsVbNullString(RS.Fields("TypeOfLoss"))
            'TypeOfLossCode
            itmX.SubItems(GuiPayReqsListView.TypeOfLossCode - 1) = goUtil.IsNullIsVbNullString(RS.Fields("TypeOfLossCode"))
            'RT50_sInsuredPayeeName
            itmX.SubItems(GuiPayReqsListView.RT50_sInsuredPayeeName - 1) = goUtil.IsNullIsVbNullString(RS.Fields("RT50_sInsuredPayeeName"))
            'RT51_sPayeeNames
            itmX.SubItems(GuiPayReqsListView.RT51_sPayeeNames - 1) = goUtil.IsNullIsVbNullString(RS.Fields("RT51_sPayeeNames"))
            'RT52_sAddress
            itmX.SubItems(GuiPayReqsListView.RT52_sAddress - 1) = Replace(goUtil.IsNullIsVbNullString(RS.Fields("RT52_sAddress")), F_VBCRLF, "    ")
            'RT52_sAddressSort 'Hidden
            itmX.SubItems(GuiPayReqsListView.RT52_sAddressSort - 1) = Replace(goUtil.utNumInTextSortFormat(goUtil.IsNullIsVbNullString(RS.Fields("RT52_sAddress"))), F_VBCRLF, "    ")
            'RT53_cAmountOfCheck
            itmX.SubItems(GuiPayReqsListView.RT53_cAmountOfCheck - 1) = Format(goUtil.IsNullIsVbNullString(RS.Fields("RT53_cAmountOfCheck")), "#,###,###,##0.00")
            'RT53_cAmountOfCheckSort 'Hidden
            itmX.SubItems(GuiPayReqsListView.RT53_cAmountOfCheckSort - 1) = goUtil.utNumInTextSortFormat(Format(goUtil.IsNullIsVbNullString(RS.Fields("RT53_cAmountOfCheck")), "#,###,###,##0.00"))
            'AppliedDeductible
            itmX.SubItems(GuiPayReqsListView.AppliedDeductible - 1) = Format(goUtil.IsNullIsVbNullString(RS.Fields("AppliedDeductible")), "#,###,###,##0.00")
            'AppliedDeductibleSort 'Hidden
            itmX.SubItems(GuiPayReqsListView.AppliedDeductibleSort - 1) = goUtil.utNumInTextSortFormat(Format(goUtil.IsNullIsVbNullString(RS.Fields("AppliedDeductible")), "#,###,###,##0.00"))
            'CatCode
            itmX.SubItems(GuiPayReqsListView.CatCode - 1) = goUtil.IsNullIsVbNullString(RS.Fields("CatCode"))
            'IsDeleted
            If CBool(RS.Fields("IsDeleted")) Then
                iMyIcon = GuiPayReqsStatusList.IsDeleted
            Else
                iMyIcon = Empty
            End If
            sFlagText = goUtil.GetFlagText(RS.Fields("IsDeleted"))
            itmX.SubItems(GuiPayReqsListView.IsDeleted - 1) = sFlagText
            itmX.ListSubItems(GuiPayReqsListView.IsDeleted - 1).ReportIcon = iMyIcon
            'UpLoadMe
            If CBool(RS.Fields("UpLoadMe")) Then
                iMyIcon = GuiPayReqsStatusList.UpLoadMe
            Else
                iMyIcon = Empty
            End If
            sFlagText = goUtil.GetFlagText(RS.Fields("UpLoadMe"))
            itmX.SubItems(GuiPayReqsListView.UpLoadMe - 1) = sFlagText
            itmX.ListSubItems(GuiPayReqsListView.UpLoadMe - 1).ReportIcon = iMyIcon
            'DateLastUpdated
            If Not IsNull(RS.Fields("DateLastUpdated").Value) Then
                If IsDate(RS.Fields("DateLastUpdated").Value) Then
                    itmX.SubItems(GuiPayReqsListView.DateLastUpdated - 1) = Format(RS.Fields("DateLastUpdated").Value, "MM/DD/YYYY HH:MM:SS")
                    itmX.SubItems(GuiPayReqsListView.DateLastUpdatedSort - 1) = Format(RS.Fields("DateLastUpdated").Value, "YYYY/MM/DD HH:MM:SS")
                Else
                    itmX.SubItems(GuiPayReqsListView.DateLastUpdated - 1) = vbNullString
                    itmX.SubItems(GuiPayReqsListView.DateLastUpdatedSort - 1) = vbNullString
                End If
            Else
                itmX.SubItems(GuiPayReqsListView.DateLastUpdated - 1) = vbNullString
                itmX.SubItems(GuiPayReqsListView.DateLastUpdatedSort - 1) = vbNullString
            End If
            'AdminComments
            itmX.SubItems(GuiPayReqsListView.AdminComments - 1) = goUtil.IsNullIsVbNullString(RS.Fields("AdminComments"))
            'tempCHeckName 'Hidden
            itmX.SubItems(GuiPayReqsListView.tempCHeckName - 1) = goUtil.IsNullIsVbNullString(RS.Fields("tempCHeckName"))
            'RT42_ClassOfLossID 'Hidden
            itmX.SubItems(GuiPayReqsListView.RT42_ClassOfLossID - 1) = goUtil.IsNullIsVbNullString(RS.Fields("RT42_ClassOfLossID"))
            'RT43_TypeOfLossID 'Hidden
            itmX.SubItems(GuiPayReqsListView.RT43_TypeOfLossID - 1) = goUtil.IsNullIsVbNullString(RS.Fields("RT43_TypeOfLossID"))
            'RT54_CompanyCatSpecID 'Hidden
            itmX.SubItems(GuiPayReqsListView.RT54_CompanyCatSpecID - 1) = goUtil.IsNullIsVbNullString(RS.Fields("RT54_CompanyCatSpecID"))
            'ID ' Hidden
            itmX.SubItems(GuiPayReqsListView.ID - 1) = goUtil.IsNullIsVbNullString(RS.Fields("ID"))
            'IDAssignments ' Hidden
            itmX.SubItems(GuiPayReqsListView.IDAssignments - 1) = goUtil.IsNullIsVbNullString(RS.Fields("IDAssignments"))
            'RTChecksID ' Hidden
            itmX.SubItems(GuiPayReqsListView.RTChecksID - 1) = goUtil.IsNullIsVbNullString(RS.Fields("RTChecksID"))
            'AssignmentsID ' Hidden
            itmX.SubItems(GuiPayReqsListView.AssignmentsID - 1) = goUtil.IsNullIsVbNullString(RS.Fields("AssignmentsID"))
            'BillingCountID ' Hidden
            itmX.SubItems(GuiPayReqsListView.BillingCountID - 1) = goUtil.IsNullIsVbNullString(RS.Fields("BillingCountID"))
            'IDBillingCount 'Hidden
            itmX.SubItems(GuiPayReqsListView.IDBillingCount - 1) = goUtil.IsNullIsVbNullString(RS.Fields("IDBillingCount"))
            'DownLoadMe ' hidden
            itmX.SubItems(GuiPayReqsListView.DownLoadMe - 1) = goUtil.IsNullIsVbNullString(RS.Fields("DownLoadMe"))
            'UpdateByUserID ' Hidden
            itmX.SubItems(GuiPayReqsListView.UpdateByUserID - 1) = goUtil.IsNullIsVbNullString(RS.Fields("UpdateByUserID"))
        
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
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub PopulatelvwPayReqs"
    oListView.Visible = True
End Sub

Private Sub PopulateIndemTotals()
    On Error GoTo EH
    Dim RS As ADODB.Recordset
    Dim cFullCostOfRepair As Currency
    Dim cRecoverableDepreciation As Currency
    Dim cNonRecovDepr As Currency
    Dim cACVLoss As Currency
    Dim cDeductible As Currency
    Dim cAppliedDeductible As Currency
    Dim cLessExcessLimits As Currency
    Dim cLessExcessLimitsAbsorbDed As Currency
    Dim cLessMiscellaneous As Currency
    Dim cNetActualCashValueClaim As Currency
    Dim cPayReqs As Currency
    
     
    'Get the Assignments RS to get Deductible
    If Not mfrmClaim.SetadoRSAssignments(msAssignmentsID) Then
        GoTo CLEAN_UP
    End If
    
    Set RS = mfrmClaim.adoRSAssignments
    
    If RS.RecordCount = 1 Then
        RS.MoveFirst
        cDeductible = CCur(Format(goUtil.IsNullIsVbNullString(RS.Fields("Deductible")), "#,###,###,##0.00"))
    End If
    
    txtFullCostOfRepairValue.Text = "0.00"
    txtRecoverableDepreciationValue.Text = "0.00"
    txtNonRecovDeprValue.Text = "0.00"
    txtACVLossValue.Text = "0.00"
    txtDeductibleValue.Text = Format(cDeductible, "#,###,###,##0.00")
    txtAppliedDeductibleValue = "0.00"
    txtLessExcessLimitsValue.Text = "0.00"
    txtLessExcessLimitsAbsorbDed = "0.00"
    txtLessMiscValue.Text = "0.00"
    txtNetActualCashValueClaimValue.Text = "0.00"
    
    Set RS = Nothing
    'Get the Assignments RS to get Total Pay Reqs
    Set RS = mfrmClaim.adoRSPayment
    
    If RS.RecordCount > 0 Then
        RS.MoveFirst
        Do Until RS.EOF
            If StrComp(goUtil.IsNullIsVbNullString(RS.Fields("ClassOfLossCode")), "Other", vbTextCompare) <> 0 And Not RS.Fields("IsDeleted") Then
                cPayReqs = cPayReqs + CCur(Format(goUtil.IsNullIsVbNullString(RS.Fields("RT53_cAmountOfCheck")), "#,###,###,##0.00"))
            End If
            RS.MoveNext
        Loop
    End If
    txtPayReqsValue.Text = Format(cPayReqs, "#,###,###,##0.00")
    
    Set RS = Nothing
    
    'Set the Indemnity
    If Not mfrmClaim.SetadoRSRTIndemnity(msAssignmentsID) Then
        Exit Sub
    End If
    'Get the Indemnity RS
    Set RS = mfrmClaim.adoRSRTIndemnity
    
    If RS.RecordCount > 0 Then
        RS.MoveFirst
        Do Until RS.EOF
            If Not RS.Fields("IsPreviousPayment") And Not RS.Fields("IsDeleted") And StrComp(goUtil.IsNullIsVbNullString(RS.Fields("ClassOfLossCode")), "Other", vbTextCompare) <> 0 Then
                cFullCostOfRepair = cFullCostOfRepair + CCur(Format(goUtil.IsNullIsVbNullString(RS.Fields("ReplacementCost")), "#,###,###,##0.00"))
                cRecoverableDepreciation = cRecoverableDepreciation + CCur(Format(goUtil.IsNullIsVbNullString(RS.Fields("RecoverableDep")), "#,###,###,##0.00"))
                cNonRecovDepr = cNonRecovDepr + CCur(Format(goUtil.IsNullIsVbNullString(RS.Fields("NonRecoverableDep")), "#,###,###,##0.00"))
                cACVLoss = cACVLoss + CCur(Format(goUtil.IsNullIsVbNullString(RS.Fields("ACVClaim")), "#,###,###,##0.00"))
                cAppliedDeductible = cAppliedDeductible + CCur(Format(goUtil.IsNullIsVbNullString(RS.Fields("AppliedDeductible")), "#,###,###,##0.00"))
                cLessExcessLimits = cLessExcessLimits + CCur(Format(goUtil.IsNullIsVbNullString(RS.Fields("ExcessLimits")), "#,###,###,##0.00"))
                If RS.Fields("ExcessAbsorbsDeductible") Then
                    cLessExcessLimitsAbsorbDed = cLessExcessLimitsAbsorbDed + CCur(Format(goUtil.IsNullIsVbNullString(RS.Fields("ExcessLimits")), "#,###,###,##0.00"))
                End If
                cLessMiscellaneous = cLessMiscellaneous + CCur(Format(goUtil.IsNullIsVbNullString(RS.Fields("Miscellaneous")), "#,###,###,##0.00"))
            End If
            RS.MoveNext
        Loop
        
        If cLessExcessLimitsAbsorbDed > 0 Then
            If cAppliedDeductible <> cDeductible - cLessExcessLimitsAbsorbDed Then
                cLessExcessLimitsAbsorbDed = cDeductible - cAppliedDeductible
            End If
        End If
        
        txtFullCostOfRepairValue.Text = Format(cFullCostOfRepair, "#,###,###,##0.00")
        txtRecoverableDepreciationValue.Text = Format(cRecoverableDepreciation, "#,###,###,##0.00")
        txtNonRecovDeprValue.Text = Format(cNonRecovDepr, "#,###,###,##0.00")
        txtACVLossValue.Text = Format(cACVLoss, "#,###,###,##0.00")
        txtAppliedDeductibleValue = Format(cAppliedDeductible, "#,###,###,##0.00")
        txtLessExcessLimitsValue.Text = Format(cLessExcessLimits, "#,###,###,##0.00")
        txtLessExcessLimitsAbsorbDed = Format(cLessExcessLimitsAbsorbDed, "#,###,###,##0.00")
        txtLessMiscValue.Text = Format(cLessMiscellaneous, "#,###,###,##0.00")
        txtNetActualCashValueClaimValue.Text = Format(cACVLoss - (cAppliedDeductible + cLessExcessLimits + cLessMiscellaneous), "#,###,###,##0.00")
    End If
    
    'If the Net ACVC does not match payment totals updated the Warning Label
    If cPayReqs <> CCur(txtNetActualCashValueClaimValue.Text) Then
        lblAmountWarning.Caption = "Indemnity - Payreq Not balanced!"
        lblAmountWarning.Visible = True
    Else
        lblAmountWarning.Visible = False
    End If
    
CLEAN_UP:
    'Cleanup
    
    Set RS = Nothing
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub PopulateIndemTotals"
End Sub

Public Function PopulateAppliedDeductible() As Boolean
    On Error GoTo EH
    Dim cDeductible As Currency
    Dim cAppliedDeductible As Currency
    Dim cAppDedItem As Currency
    Dim cACVC As Currency
    Dim cACVCItem As Currency
    Dim cTemp As Currency
    Dim cLessExcessLimits As Currency
    Dim cLessExcessLimitsItem As Currency
    Dim RS As ADODB.Recordset
    Dim RS2 As ADODB.Recordset
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    Dim oListView As ListView
    Dim itmX As ListItem
    Dim IndemitmX As ListItem
    Dim sClassTypeID As String
    Dim sClassOfLossID As String
    Dim cRemainDed As Currency
    Dim cNewAppDedItem As Currency
    Dim sIndemID As String
    Dim MyIndemUdt As GuiIndemItem
        
    'Need to Get the Amount of deductible applied
    'Get the Latest Updated Indem
    mfrmClaim.SetadoRSRTIndemnity msAssignmentsID
    Set RS = mfrmClaim.adoRSRTIndemnity
    If RS.RecordCount > 0 Then
        
        Set oConn = New ADODB.Connection
        goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
        
        'Get Deductible
        If IsNumeric(txtDeductibleValue.Text) Then
            cDeductible = txtDeductibleValue.Text
        End If
        
        'First Figure Out the Deductible after
        'Allowing Excess limits to Absorb Deductible
        RS.MoveFirst
        Do Until RS.EOF
            If Not RS.Fields("IsPreviousPayment") _
                And Not RS.Fields("IsDeleted") _
                And StrComp(goUtil.IsNullIsVbNullString(RS.Fields("ClassOfLossCode")), "Other", vbTextCompare) <> 0 Then
                If RS.Fields("ExcessAbsorbsDeductible") Then
                    cLessExcessLimitsItem = CCur(Format(goUtil.IsNullIsVbNullString(RS.Fields("ExcessLimits")), "#,###,###,##0.00"))
                    cLessExcessLimits = cLessExcessLimits + cLessExcessLimitsItem
                End If
            End If
            RS.MoveNext
        Loop
        
        If cLessExcessLimits > 0 Then
            If cLessExcessLimits >= cDeductible Then
                cDeductible = 0
            Else
                cDeductible = cDeductible - cLessExcessLimits
            End If
        End If
        
        cACVC = 0
        cLessExcessLimits = 0
        cAppliedDeductible = 0
        'Loop throught the Listview for Apply Ded
        For Each itmX In lvwDed.ListItems
            sClassTypeID = itmX.SubItems(GuiPolicyLimits.ClassTypeID - 1)
            sSQL = "SELECT COL.[ClassOfLossID], "
            sSQL = sSQL & "(CT.Description + ' (' + COL.Description + ')') As ClassOfLoss, "
            sSQL = sSQL & "COL.Code As ClassOfLossCode "
            sSQL = sSQL & "FROM ClassOfLoss COL "
            sSQL = sSQL & "INNER JOIN ClassType CT ON CT.ClassTypeID = COL.ClassTypeID "
            sSQL = sSQL & "WHERE COL.[ClientCompanyID] = " & goUtil.gsCurCar & " "
            sSQL = sSQL & "AND COL.[ClassTypeID] = " & sClassTypeID & " "
            Set RS2 = New ADODB.Recordset
            RS2.CursorLocation = adUseClient
            RS2.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
            Set RS2.ActiveConnection = Nothing
            If RS2.RecordCount = 1 Then
                RS2.MoveFirst
                sClassOfLossID = goUtil.IsNullIsVbNullString(RS2.Fields("ClassOfLossID"))
            End If
            
            RS.MoveFirst
            Do Until RS.EOF
                cACVCItem = 0
                cLessExcessLimitsItem = 0
                cAppDedItem = 0
                If Not RS.Fields("IsPreviousPayment") _
                    And Not RS.Fields("IsDeleted") _
                    And StrComp(goUtil.IsNullIsVbNullString(RS.Fields("ClassOfLossCode")), "Other") <> 0 _
                    And goUtil.IsNullIsVbNullString(RS.Fields("ClassOfLossID")) = sClassOfLossID Then
                    sIndemID = goUtil.IsNullIsVbNullString(RS.Fields("ID"))
                    cACVCItem = CCur(Format(goUtil.IsNullIsVbNullString(RS.Fields("ACVLessExcessLimits")), "#,###,###,##0.00"))
                    cACVC = cACVC + cACVCItem
                    'Need to See if the Applied Deductible for this Item needs to Change
                    cAppDedItem = CCur(Format(goUtil.IsNullIsVbNullString(RS.Fields("AppliedDeductible")), "#,###,###,##0.00"))
                    
                    'Get the Remaining Deductible
                    cRemainDed = (cDeductible - cAppliedDeductible)
                    
                    'Figure out what the applicable Ded for this item should be
                    If cRemainDed <= 0 Then
                        cNewAppDedItem = 0
                    Else
                        'Get RemainDeductible
                        If cRemainDed > 0 Then
                            If cACVCItem >= cRemainDed Then
                                cNewAppDedItem = cRemainDed
                            Else
                                cNewAppDedItem = cRemainDed - (cRemainDed - cACVCItem)
                            End If
                        Else
                            cNewAppDedItem = 0
                        End If
                        
                    End If
                    
                    cAppliedDeductible = cAppliedDeductible + cNewAppDedItem
                    
                    'If the New Applied Ded For this Item Does Not match what
                    'was previously stored, need to Update the DB
                    If cNewAppDedItem <> cAppDedItem Then
                        Set IndemitmX = lvwIndemnity.ListItems("""" & sIndemID & """")
                        With MyIndemUdt
                            .RTIndemnityID = IIf(IsNull(RS.Fields("RTIndemnityID")), "NULL", RS.Fields("RTIndemnityID"))
                            .AssignmentsID = IIf(IsNull(RS.Fields("AssignmentsID")), "NULL", RS.Fields("AssignmentsID"))
                            .RTChecksID = IIf(IsNull(RS.Fields("RTChecksID")), "NULL", RS.Fields("RTChecksID"))
                            .ID = IIf(IsNull(RS.Fields("ID")), "NULL", RS.Fields("ID"))
                            .IDAssignments = IIf(IsNull(RS.Fields("IDAssignments")), "NULL", RS.Fields("IDAssignments"))
                            .IDRTChecks = IIf(IsNull(RS.Fields("IDRTChecks")), "NULL", RS.Fields("IDRTChecks"))
                            .ACVClaim = IIf(IsNull(RS.Fields("ACVClaim")), "NULL", RS.Fields("ACVClaim"))
                            .ACVLessExcessLimits = IIf(IsNull(RS.Fields("ACVLessExcessLimits")), "NULL", RS.Fields("ACVLessExcessLimits"))
                            .SpecialLimits = IIf(IsNull(RS.Fields("SpecialLimits")), "NULL", RS.Fields("SpecialLimits"))
                            .ExcessLimits = IIf(IsNull(RS.Fields("ExcessLimits")), "NULL", RS.Fields("ExcessLimits"))
                            .Miscellaneous = IIf(IsNull(RS.Fields("Miscellaneous")), "NULL", RS.Fields("Miscellaneous"))
                            .MiscDescription = IIf(IsNull(RS.Fields("MiscDescription")), "NULL", RS.Fields("MiscDescription"))
                            .IsAddAmountOfInsurance = IIf(IsNull(RS.Fields("IsAddAmountOfInsurance")), "NULL", RS.Fields("IsAddAmountOfInsurance"))
                            .ExcessAbsorbsDeductible = IIf(IsNull(RS.Fields("ExcessAbsorbsDeductible")), "NULL", RS.Fields("ExcessAbsorbsDeductible"))
                            .AppliedDeductible = cNewAppDedItem
                            IndemitmX.SubItems(GuiIndemListView.AppliedDeductible - 1) = cNewAppDedItem
                            IndemitmX.SubItems(GuiIndemListView.AppliedDeductibleSort - 1) = goUtil.utNumInTextSortFormat(CStr(cNewAppDedItem))
                            .NonRecoverableDep = IIf(IsNull(RS.Fields("NonRecoverableDep")), "NULL", RS.Fields("NonRecoverableDep"))
                            .RecoverableDep = IIf(IsNull(RS.Fields("RecoverableDep")), "NULL", RS.Fields("RecoverableDep"))
                            .ReplacementCost = IIf(IsNull(RS.Fields("ReplacementCost")), "NULL", RS.Fields("ReplacementCost"))
                            .TypeOfLossID = IIf(IsNull(RS.Fields("TypeOfLossID")), "NULL", RS.Fields("TypeOfLossID"))
                            .ClassOfLossID = IIf(IsNull(RS.Fields("ClassOfLossID")), "NULL", RS.Fields("ClassOfLossID"))
                            .Description = IIf(IsNull(RS.Fields("Description")), "NULL", RS.Fields("Description"))
                            .IsPreviousPayment = IIf(IsNull(RS.Fields("IsPreviousPayment")), "NULL", RS.Fields("IsPreviousPayment"))
                            .PPayDatePaid = IIf(IsNull(RS.Fields("PPayDatePaid")), "NULL", RS.Fields("PPayDatePaid"))
                            .PPayAmountPaid = IIf(IsNull(RS.Fields("PPayAmountPaid")), "NULL", RS.Fields("PPayAmountPaid"))
                            .PPayCheckNumber = IIf(IsNull(RS.Fields("PPayCheckNumber")), "NULL", RS.Fields("PPayCheckNumber"))
                            .IsDeleted = IIf(IsNull(RS.Fields("IsDeleted")), "NULL", RS.Fields("IsDeleted"))
                            .DownLoadMe = IIf(IsNull(RS.Fields("DownLoadMe")), "NULL", RS.Fields("DownLoadMe"))
                            IndemitmX.ListSubItems(GuiIndemListView.UpLoadMe - 1).ReportIcon = GuiIndemStatusList.UpLoadMe
                            .UpLoadMe = True
                            .AdminComments = IIf(IsNull(RS.Fields("AdminComments")), "NULL", RS.Fields("AdminComments"))
                            .DateLastUpdated = Format(Now(), "MM/DD/YYYY HH:MM:SS")
                            IndemitmX.SubItems(GuiIndemListView.DateLastUpdated - 1) = .DateLastUpdated
                            IndemitmX.SubItems(GuiIndemListView.DateLastUpdatedSort - 1) = Format(.DateLastUpdated, "YYYY/MM/DD HH:MM:SS")
                            IndemitmX.SubItems(GuiIndemListView.UpdateByUserID - 1) = goUtil.gsCurUsersID
                            .UpdateByUserID = IndemitmX.SubItems(GuiIndemListView.UpdateByUserID - 1)
                        End With
                        EditIdemnityItem MyIndemUdt
                        lvwIndemnity.SortKey = GuiIndemListView.PaymentRequest
                        lvwIndemnity.Sorted = True
                    End If
                    
                End If
                RS.MoveNext
            Loop
        Next
    End If
    
    PopulateAppliedDeductible = True
    
    'cleanup
    Set RS = Nothing
    Set oConn = Nothing
    Set itmX = Nothing
    Set IndemitmX = Nothing
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function PopulateAppliedDeductible"
End Function

Public Function AddPayReqItem(pudtPayReq As GuiPayReqsItem, psID As String) As Boolean
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    Dim sID As String
    
    Screen.MousePointer = MousePointerConstants.vbHourglass
    
    sID = goUtil.GetAccessDBUID("ID", "RTChecks")
    
    With pudtPayReq
        .RTChecksID = sID
        .AssignmentsID = msAssignmentsID
        'Use Current Billing Item if Available
        'Do not Associate IB to ClassOfLossCode that is Other
        If StrComp(mfrmClaim.GetClassOfLossCode(.RT42_ClassOfLossID), "Other", vbTextCompare) <> 0 Then
            .BillingCountID = mfrmClaim.GetCurrentBillingCountID(False)
            .IDBillingCount = mfrmClaim.GetCurrentBillingCountID(True)
        Else
            .BillingCountID = "Null"
            .IDBillingCount = "Null"
        End If
        .ID = sID
        .IDAssignments = msAssignmentsID 'not set here
    End With
    
    sSQL = "INSERT INTO RTChecks "
    sSQL = sSQL & "( "
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
    sSQL = sSQL & "[UpdateByUserID] "
    sSQL = sSQL & ") "
    sSQL = sSQL & "SELECT "
    sSQL = sSQL & pudtPayReq.RTChecksID & " As [RTChecksID], "
    sSQL = sSQL & pudtPayReq.AssignmentsID & " As [AssignmentsID], "
    sSQL = sSQL & pudtPayReq.BillingCountID & " As [BillingCountID] , "
    sSQL = sSQL & pudtPayReq.ID & " As [ID], "
    sSQL = sSQL & pudtPayReq.IDAssignments & " As [IDAssignments], "
    sSQL = sSQL & pudtPayReq.IDBillingCount & " As [IDBillingCount], "
    sSQL = sSQL & pudtPayReq.CheckNum & " As [CheckNum], "
    sSQL = sSQL & pudtPayReq.RT42_ClassOfLossID & " As [RT42_ClassOfLossID], "
    sSQL = sSQL & pudtPayReq.RT43_TypeOfLossID & " As [RT43_TypeOfLossID], "
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtPayReq.RT50_sInsuredPayeeName) & "'" & " As [RT50_sInsuredPayeeName], "
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtPayReq.RT51_sPayeeNames) & "'" & " As [RT51_sPayeeNames], "
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtPayReq.RT52_sAddress) & "'" & " As [RT52_sAddress], "
    sSQL = sSQL & pudtPayReq.RT53_cAmountOfCheck & " As [RT53_cAmountOfCheck], "
    sSQL = sSQL & pudtPayReq.AppliedDeductible & " As [AppliedDeductible], "
    sSQL = sSQL & pudtPayReq.RT54_CompanyCatSpecID & " As [RT54_CompanyCatSpecID], "
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtPayReq.tempCHeckName) & "'" & " As [tempCHeckName], "
    sSQL = sSQL & pudtPayReq.PrintOnIB & " As [PrintOnIB], "
    sSQL = sSQL & pudtPayReq.IsDeleted & " As [IsDeleted], "
    sSQL = sSQL & pudtPayReq.DownLoadMe & " As [DownLoadMe], "
    sSQL = sSQL & pudtPayReq.UpLoadMe & " As [UpLoadMe], "
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtPayReq.AdminComments) & "'" & " As [AdminComments], "
    sSQL = sSQL & "#" & pudtPayReq.DateLastUpdated & "#" & " As [DateLastUpdated], "
    sSQL = sSQL & pudtPayReq.UpdateByUserID & " As [UpdateByUserID] "
   
    Set oConn = New ADODB.Connection
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    oConn.Execute sSQL
    
    Screen.MousePointer = MousePointerConstants.vbDefault
    psID = sID
    AddPayReqItem = True
    
    'Clean up
    Set oConn = Nothing
    Exit Function
EH:
    Screen.MousePointer = vbDefault
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function AddPayReqItem"
End Function

Public Function AddIndemnityItem(pudtIndem As GuiIndemItem, psID As String) As Boolean
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    Dim sID As String
    
    Screen.MousePointer = MousePointerConstants.vbHourglass
    
    sID = goUtil.GetAccessDBUID("ID", "RTIndemnity")
    
    With pudtIndem
        .RTIndemnityID = sID
        .AssignmentsID = msAssignmentsID
        .RTChecksID = "Null"
        .ID = sID
        .IDAssignments = msAssignmentsID 'not set here
        .IDRTChecks = "Null"
    End With
    
    sSQL = "INSERT INTO RTIndemnity "
    sSQL = sSQL & "( "
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
    sSQL = sSQL & ") "
    sSQL = sSQL & "SELECT "
    sSQL = sSQL & pudtIndem.RTIndemnityID & " As [RTIndemnityID], "
    sSQL = sSQL & pudtIndem.AssignmentsID & " As [AssignmentsID], "
    sSQL = sSQL & pudtIndem.RTChecksID & " As [RTChecksID] , "
    sSQL = sSQL & pudtIndem.ID & " As [ID], "
    sSQL = sSQL & pudtIndem.IDAssignments & " As [IDAssignments], "
    sSQL = sSQL & pudtIndem.IDRTChecks & " As [IDRTChecks], "
    sSQL = sSQL & pudtIndem.ACVClaim & " As [ACVClaim], "
    sSQL = sSQL & pudtIndem.ACVLessExcessLimits & " As [ACVLessExcessLimits], "
    sSQL = sSQL & pudtIndem.SpecialLimits & " As [SpecialLimits], "
    sSQL = sSQL & pudtIndem.ExcessLimits & " As [ExcessLimits], "
    sSQL = sSQL & pudtIndem.Miscellaneous & " As [Miscellaneous], "
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtIndem.MiscDescription) & "'" & " As [MiscDescription], "
    sSQL = sSQL & pudtIndem.IsAddAmountOfInsurance & " As [IsAddAmountOfInsurance], "
    sSQL = sSQL & pudtIndem.ExcessAbsorbsDeductible & " As [ExcessAbsorbsDeductible], "
    sSQL = sSQL & pudtIndem.AppliedDeductible & " As [AppliedDeductible], "
    sSQL = sSQL & pudtIndem.NonRecoverableDep & " As [NonRecoverableDep], "
    sSQL = sSQL & pudtIndem.RecoverableDep & " As [RecoverableDep], "
    sSQL = sSQL & pudtIndem.ReplacementCost & " As [ReplacementCost], "
    sSQL = sSQL & pudtIndem.TypeOfLossID & " As [TypeOfLossID], "
    sSQL = sSQL & pudtIndem.ClassOfLossID & " As [ClassOfLossID], "
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtIndem.Description) & "'" & " As [Description], "
    sSQL = sSQL & pudtIndem.IsPreviousPayment & " As [IsPreviousPayment], "
    sSQL = sSQL & "Null" & " As [PPayDatePaid], "
    sSQL = sSQL & "Null" & " As [PPayAmountPaid], "
    sSQL = sSQL & "Null" & " As [PPayCheckNumber], "
    sSQL = sSQL & pudtIndem.IsDeleted & " As [IsDeleted], "
    sSQL = sSQL & pudtIndem.DownLoadMe & " As [DownLoadMe], "
    sSQL = sSQL & pudtIndem.UpLoadMe & " As [UpLoadMe], "
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtIndem.AdminComments) & "'" & " As [AdminComments], "
    sSQL = sSQL & "#" & pudtIndem.DateLastUpdated & "#" & " As [DateLastUpdated], "
    sSQL = sSQL & pudtIndem.UpdateByUserID & " As [UpdateByUserID] "
   
    Set oConn = New ADODB.Connection
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    oConn.Execute sSQL
    
    Sleep 500
    Screen.MousePointer = MousePointerConstants.vbDefault
    psID = sID
    AddIndemnityItem = True
    
    'Clean up
    Set oConn = Nothing
    Exit Function
EH:
    Screen.MousePointer = vbDefault
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function AddIndemnityItem("
End Function


Public Function EditPayReqItem(pudtPayReq As GuiPayReqsItem) As Boolean
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    
    With pudtPayReq
        If .RTChecksID = vbNullString Or .RTChecksID = "0" Then
            .RTChecksID = "Null"
        End If
        .AssignmentsID = msAssignmentsID
        'Use Current Billing Item if Available
        'Do not Associate IB to ClassOfLossCode that is Other
        If StrComp(mfrmClaim.GetClassOfLossCode(.RT42_ClassOfLossID), "Other", vbTextCompare) <> 0 Then
            .BillingCountID = mfrmClaim.GetCurrentBillingCountID(False)
            .IDBillingCount = mfrmClaim.GetCurrentBillingCountID(True)
        Else
            .BillingCountID = "Null"
            .IDBillingCount = "Null"
        End If
        If .ID = vbNullString Or .ID = "0" Then
            .ID = "Null"
        End If
        .IDAssignments = msAssignmentsID 'not set here
    End With

    sSQL = "UPDATE RTChecks Set "
    sSQL = sSQL & "[RTChecksID] = " & pudtPayReq.RTChecksID & ", "
    sSQL = sSQL & "[AssignmentsID] = " & pudtPayReq.AssignmentsID & ", "
    'When editing a Payreq, Only update the BillingcountID if it is Currently Null
    sSQL = sSQL & "[BillingCountID] = IIF(IsNull([BillingCountID]), " & pudtPayReq.BillingCountID & ", [BillingCountID]), "
    sSQL = sSQL & "[ID] = " & pudtPayReq.ID & ", "
    sSQL = sSQL & "[IDAssignments] = " & pudtPayReq.IDAssignments & ", "
    'When editing a Payreq, Only update the BillingcountID if it is Currently Null
    sSQL = sSQL & "[IDBillingCount] = IIF(IsNull([IDBillingCount]), " & pudtPayReq.IDBillingCount & ", [IDBillingCount]), "
    sSQL = sSQL & "[CheckNum] = " & pudtPayReq.CheckNum & ", "
    sSQL = sSQL & "[RT42_ClassOfLossID] = " & pudtPayReq.RT42_ClassOfLossID & ", "
    sSQL = sSQL & "[RT43_TypeOfLossID] = " & pudtPayReq.RT43_TypeOfLossID & ", "
    sSQL = sSQL & "[RT50_sInsuredPayeeName] = '" & goUtil.utCleanSQLString(pudtPayReq.RT50_sInsuredPayeeName) & "', "
    sSQL = sSQL & "[RT51_sPayeeNames] = '" & goUtil.utCleanSQLString(pudtPayReq.RT51_sPayeeNames) & "', "
    sSQL = sSQL & "[RT52_sAddress] = '" & goUtil.utCleanSQLString(pudtPayReq.RT52_sAddress) & "', "
    sSQL = sSQL & "[RT53_cAmountOfCheck] = " & CCur(pudtPayReq.RT53_cAmountOfCheck) & ", "
    sSQL = sSQL & "[AppliedDeductible] = " & CCur(pudtPayReq.AppliedDeductible) & ", "
    sSQL = sSQL & "[RT54_CompanyCatSpecID] = " & pudtPayReq.RT54_CompanyCatSpecID & ", "
    sSQL = sSQL & "[tempCHeckName] = '" & goUtil.utCleanSQLString(pudtPayReq.tempCHeckName) & "', "
    sSQL = sSQL & "[PrintOnIB] = " & pudtPayReq.PrintOnIB & ", "
    sSQL = sSQL & "[IsDeleted] = " & pudtPayReq.IsDeleted & ", "
    sSQL = sSQL & "[DownLoadMe] = " & pudtPayReq.DownLoadMe & ", "
    sSQL = sSQL & "[UpLoadMe] = " & pudtPayReq.UpLoadMe & ", "
    sSQL = sSQL & "[AdminComments] = '" & goUtil.utCleanSQLString(pudtPayReq.AdminComments) & "', "
    sSQL = sSQL & "[DateLastUpdated] = #" & pudtPayReq.DateLastUpdated & "#, "
    sSQL = sSQL & "[UpdateByUserID]  = " & pudtPayReq.UpdateByUserID & " "
    sSQL = sSQL & "WHERE IDAssignments = " & pudtPayReq.IDAssignments & " "
    sSQL = sSQL & "AND ID = " & pudtPayReq.ID & " "

    Set oConn = New ADODB.Connection
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    oConn.Execute sSQL

    EditPayReqItem = True
    'Clean up
    Set oConn = Nothing
    
    Exit Function

EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function EditPayReqItem"
End Function

Public Function DeletePayReqItem(psID As String) As Boolean
    On Error GoTo EH
    Dim sSQL As String
    Dim bUpdateAsDeletedOnly As Boolean
    Dim oConn As ADODB.Connection
    Dim RS As ADODB.Recordset
    Dim RSPackageItem As ADODB.Recordset
    Dim bIsDeleted As Boolean
    Dim sPackageItemID As String
    Dim sCheckNum As String
    
    Set oConn = New ADODB.Connection
    Set RS = New ADODB.Recordset
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name


    sSQL = "SELECT A.[ID] "
    sSQL = sSQL & "FROM RTChecks A "
    sSQL = sSQL & "WHERE A.[ID] = " & psID & " "
    'Only allow actual deletion of PayReq (RTChecks) that have never been uploaded
    'Negative number for the Main Table Indentity will be negative number
    'if this is true.
    sSQL = sSQL & "AND (A.[RTChecksID] Is Null Or A.[RTChecksID] < 0)  "
    
    'Use Disconnected Record Set on asUseClient Cursor ONLY !
    RS.CursorLocation = adUseClient
    RS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set RS.ActiveConnection = Nothing
    
    If RS.RecordCount = 1 Then
        bUpdateAsDeletedOnly = False
    Else
        bUpdateAsDeletedOnly = True
    End If
    
    
    '---------------------------Package Item Update------------------
    Set RS = New ADODB.Recordset
    
    sSQL = "SELECT [CheckNum] "
    sSQL = sSQL & "FROM RTChecks "
    sSQL = sSQL & "WHERE [ID] = " & psID & " "
    sSQL = sSQL & "AND [IDAssignments] = " & msAssignmentsID & " "
    
    RS.CursorLocation = adUseClient
    RS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set RS.ActiveConnection = Nothing
    
    If RS.RecordCount = 1 Then
        RS.MoveFirst
        If Not IsNull(RS.Fields("CheckNum")) Then
            sCheckNum = RS.Fields("CheckNum")
        End If
    End If
    '---------------------------package Item Update^^^^^^^^^^^^^^^^^^^^^
    
    'Since this Item May Be Associated in other tables...
    'Need to Set the Associations to Null
    
    'Indemnity Table
    sSQL = "UPDATE RTIndemnity SET "
    sSQL = sSQL & "[RTChecksID] = null, "
    sSQL = sSQL & "[IDRTChecks] = null, "
    sSQL = sSQL & "[DateLastUpdated] = #" & Format(Now(), "MM/DD/YYYY HH:MM:SS") & "#, "
    sSQL = sSQL & "[UpdateByUserID] = " & goUtil.gsCurUsersID & " "
    sSQL = sSQL & "WHERE [IDAssignments] = " & msAssignmentsID & " "
    sSQL = sSQL & "AND [IDRTChecks] = " & psID & " "
    oConn.Execute sSQL
    
    If bUpdateAsDeletedOnly Then
        sSQL = "UPDATE RTChecks SET "
        sSQL = sSQL & "[IsDeleted] = IIF([IsDeleted], False, True), "
        sSQL = sSQL & "[UpLoadMe] = True, "
        sSQL = sSQL & "[DateLastUpdated] = #" & Format(Now(), "MM/DD/YYYY HH:MM:SS") & "#, "
        sSQL = sSQL & "[UpdateByUserID] = " & goUtil.gsCurUsersID & " "
        sSQL = sSQL & "WHERE [ID] = " & psID & " "
        sSQL = sSQL & "AND [IDAssignments] = " & msAssignmentsID & " "
    Else
        'If Removing the record need to remove any Misc Params as well !
        '2/8/2004 MiscReportParam , and MiscReportParam01 to MiscReportParam30
        '4/13/2005 BGS Issue 316  frmIndemnity ERROR # -2147217865
        '...cannot find the input table or query 'MRP'...
        'Pay Req Items are currently stored under MiscReportParam only
        sSQL = "DELETE * FROM MiscReportParam MRP "
        '2/8/2004 MiscReportParam , and MiscReportParam01 to MiscReportParam30
        sSQL = sSQL & "WHERE MRP.[IDAssignments] = " & msAssignmentsID & " "
        sSQL = sSQL & "AND MRP.Number IN ("
                            sSQL = sSQL & "SELECT A.[CheckNum] "
                            sSQL = sSQL & "FROM RTChecks A "
                            sSQL = sSQL & "WHERE A.[ID] = " & psID & " "
                            sSQL = sSQL & "AND A.[IDAssignments] = " & msAssignmentsID & " "
        sSQL = sSQL & ") "
        sSQL = sSQL & "AND ProjectName = '" & goUtil.utCleanSQLString("ECRpt" & goUtil.gsCurCarDBName & "_arRptAddlChk") & "' "
        oConn.Execute sSQL
        sSQL = "DELETE * FROM RTChecks A "
        sSQL = sSQL & "WHERE A.[ID] = " & psID & " "
        sSQL = sSQL & "AND A.[IDAssignments] = " & msAssignmentsID & " "
    End If
    
    oConn.Execute sSQL
  
    '---------------------------Package Item Update------------------
    'if so need to update the package item if it also happens to be in there and is not already  flagged as deleted
    If sCheckNum = vbNullString Then
        GoTo CLEAN_UP
    End If
    
    mfrmClaim.SetadoRSPackageItem msAssignmentsID, vbNullString, , sCheckNum
    Set RSPackageItem = mfrmClaim.adoRSPackageItem
    If RSPackageItem.RecordCount > 0 Then
        Do Until RSPackageItem.EOF
            bIsDeleted = goUtil.IsNullIsVbNullString(RSPackageItem.Fields("IsDeleted"))
            If Not bIsDeleted Then
                sPackageItemID = goUtil.IsNullIsVbNullString(RSPackageItem.Fields("PackageItemID"))
                If bUpdateAsDeletedOnly Then
                    sSQL = "UPDATE PackageItem SET "
                    sSQL = sSQL & "[IsDeleted] = True, "
                    sSQL = sSQL & "[UploadMe] = True, "
                    sSQL = sSQL & "[DateLastUpdated] = #" & Now() & "#, "
                    sSQL = sSQL & "[UpdateByUserID] = " & goUtil.gsCurUsersID & " "
                    sSQL = sSQL & "WHERE [AssignmentsID] = " & msAssignmentsID & " "
                    sSQL = sSQL & "AND [PackageItemID] = " & sPackageItemID & " "
                    'BGS 6.13.2005 Important!!!
                    'Do not Allow Delete from Package Table if it's Sent Date
                    'has been set.
                    sSQL = sSQL & "AND [SentDate] Is Null  "
                Else
                    sSQL = "DELETE * FROM PackageItem "
                    sSQL = sSQL & "WHERE [AssignmentsID] = " & msAssignmentsID & " "
                    sSQL = sSQL & "AND [PackageItemID] = " & sPackageItemID & " "
                End If
                oConn.Execute sSQL
                Sleep 100
            End If
            RSPackageItem.MoveNext
        Loop
    End If
    '---------------------------package Item Update^^^^^^^^^^^^^^^^^^^^^

    DeletePayReqItem = True
    
CLEAN_UP:
    'clean up
    
    Set RS = Nothing
    Set RSPackageItem = Nothing
    Set oConn = Nothing
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function DeletePayReqItem"
End Function

Public Function DeleteIndemnityItem(psID As String) As Boolean
    On Error GoTo EH
    Dim sSQL As String
    Dim bUpdateAsDeletedOnly As Boolean
    Dim oConn As ADODB.Connection
    Dim RS As ADODB.Recordset
    
    Set oConn = New ADODB.Connection
    Set RS = New ADODB.Recordset
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name


    sSQL = "SELECT A.[ID] "
    sSQL = sSQL & "FROM RTIndemnity A "
    sSQL = sSQL & "WHERE A.[ID] = " & psID & " "
    'Only allow actual deletion of PayReq (RTChecks) that have never been uploaded
    'Negative number for the Main Table Indentity will be negative number
    'if this is true.
    sSQL = sSQL & "AND (A.[RTIndemnityID] Is Null Or A.[RTIndemnityID] < 0)  "
    
    'Use Disconnected Record Set on asUseClient Cursor ONLY !
    RS.CursorLocation = adUseClient
    RS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set RS.ActiveConnection = Nothing
    
    If RS.RecordCount = 1 Then
        bUpdateAsDeletedOnly = False
    Else
        bUpdateAsDeletedOnly = True
    End If

    
    If bUpdateAsDeletedOnly Then
        sSQL = "UPDATE RTIndemnity SET "
        sSQL = sSQL & "[IsDeleted] = IIF([IsDeleted], False, True), "
        sSQL = sSQL & "[UpLoadMe] = True, "
        sSQL = sSQL & "[DateLastUpdated] = #" & Format(Now(), "MM/DD/YYYY HH:MM:SS") & "#, "
        sSQL = sSQL & "[UpdateByUserID] = " & goUtil.gsCurUsersID & " "
        sSQL = sSQL & "WHERE [ID] = " & psID & " "
        sSQL = sSQL & "AND [IDAssignments] = " & msAssignmentsID & " "
    Else
        sSQL = "DELETE * FROM RTIndemnity A "
        sSQL = sSQL & "WHERE A.[ID] = " & psID & " "
        sSQL = sSQL & "AND A.[IDAssignments] = " & msAssignmentsID & " "
    End If
    

    oConn.Execute sSQL
    
    DeleteIndemnityItem = True
    'clean up
    
    Set RS = Nothing
    Set oConn = Nothing
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function DeleteIndemnityItem"
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
    'Set policy limits back
    mfrmClaim.SetadoRSPolicyLimits msAssignmentsID, False
    Set MyadoRSClassType = mfrmClaim.adoRSPolicyLimits
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
                sTemp = MyadoRSClassType.Fields("ClassTypeClass").Value
                sTemp = sTemp & " ("
                sTemp = sTemp & MyadoRSClassType.Fields("ClassTypeDescription").Value
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

Private Sub PopulatelvwDed()
    On Error GoTo EH
    Dim itmX As ListItem
    Dim iMyIcon As Long
    Dim sFlagText As String
    Dim sTemp As String
    Dim MyadoRSPolicyLimits As ADODB.Recordset
    Dim sAppDedClassTypeIDOrder As String
    Dim RS As ADODB.Recordset
    Dim saryDedOrder() As String
    Dim lPos As Long
    Dim sThisClassTypeID As String
    
    
    'Need to get Apply Deductible Order
    Set RS = mfrmClaim.adoRSAssignments
    If RS.RecordCount = 1 Then
        RS.MoveFirst
        sAppDedClassTypeIDOrder = goUtil.IsNullIsVbNullString(RS.Fields("AppDedClassTypeIDOrder"))
        Set RS = Nothing
    End If
    
    If sAppDedClassTypeIDOrder <> vbNullString Then
        saryDedOrder() = Split(sAppDedClassTypeIDOrder, ",")
    End If
   
    If Not mfrmClaim.adoRSPolicyLimits Is Nothing Then
        Set MyadoRSPolicyLimits = mfrmClaim.adoRSPolicyLimits
    Else
        Exit Sub
    End If
    
    'Clear Any Existing Items
    lvwDed.ListItems.Clear
    
    If Not MyadoRSPolicyLimits.EOF Then
        MyadoRSPolicyLimits.MoveFirst
        Do Until MyadoRSPolicyLimits.EOF
            'Class
            sTemp = goUtil.IsNullIsVbNullString(MyadoRSPolicyLimits.Fields("ID"))
            Set itmX = lvwDed.ListItems.Add(, """" & sTemp & """", goUtil.IsNullIsVbNullString(MyadoRSPolicyLimits.Fields("ClassTypeClass")))
            'ClassTypeID
            sThisClassTypeID = goUtil.IsNullIsVbNullString(MyadoRSPolicyLimits.Fields("ClassTypeID"))
            itmX.SubItems(GuiPolicyLimits.ClassTypeID - 1) = sThisClassTypeID
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
            'SortOrder
            If sAppDedClassTypeIDOrder <> vbNullString Then
                For lPos = LBound(saryDedOrder, 1) To UBound(saryDedOrder, 1)
                    If sThisClassTypeID = saryDedOrder(lPos) Then
                        Exit For
                    End If
                Next
            End If
            itmX.SubItems(GuiPolicyLimits.SortOrder - 1) = goUtil.utNumInTextSortFormat(CStr(lPos))
            
            itmX.Selected = False
            MyadoRSPolicyLimits.MoveNext
        Loop
    End If
    
    'make sure the itmes are sorted according to
    'Class Type Order
    lvwDed.SortKey = GuiPolicyLimits.SortOrder - 1
    lvwDed.Sorted = True
    
    'Cleanup
    Set itmX = Nothing
    Set MyadoRSPolicyLimits = Nothing
    
    Exit Sub
EH:
    Set itmX = Nothing
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub PopulatelvwDed"
End Sub

Private Sub txtACVLossValue_GotFocus()
    goUtil.utSelText txtACVLossValue
End Sub

Private Sub txtAppliedDeductibleValue_GotFocus()
    goUtil.utSelText txtAppliedDeductibleValue
End Sub

Private Sub txtDeductibleValue_GotFocus()
    goUtil.utSelText txtDeductibleValue
End Sub

Private Sub txtFullCostOfRepairValue_GotFocus()
    goUtil.utSelText txtFullCostOfRepairValue
End Sub

Private Sub txtLessExcessLimitsAbsorbDed_GotFocus()
    goUtil.utSelText txtLessExcessLimitsAbsorbDed
End Sub

Private Sub txtLessExcessLimitsValue_GotFocus()
    goUtil.utSelText txtLessExcessLimitsValue
End Sub

Private Sub txtLessMiscValue_GotFocus()
    goUtil.utSelText txtLessMiscValue
End Sub

Private Sub txtNetActualCashValueClaimValue_GotFocus()
    goUtil.utSelText txtNetActualCashValueClaimValue
End Sub

Private Sub txtNonRecovDeprValue_GotFocus()
    goUtil.utSelText txtNonRecovDeprValue
End Sub

Private Sub txtPayReqsValue_GotFocus()
    goUtil.utSelText txtPayReqsValue
End Sub

Private Sub txtRecoverableDepreciationValue_GotFocus()
    goUtil.utSelText txtRecoverableDepreciationValue
End Sub

