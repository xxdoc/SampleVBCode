VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBillingInfo 
   AutoRedraw      =   -1  'True
   Caption         =   "Billing Information"
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
   StartUpPosition =   3  'Windows Default
   Tag             =   "Billing Information"
   Begin VB.ComboBox cboCopy 
      Height          =   360
      ItemData        =   "frmBillingInfo.frx":0000
      Left            =   1920
      List            =   "frmBillingInfo.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   600
      Width           =   2535
   End
   Begin VB.CommandButton cmdPrintIBReport 
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
      Left            =   4440
      TabIndex        =   12
      Top             =   600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtServiceFee 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E9E9E9&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   6990
      Locked          =   -1  'True
      MaxLength       =   14
      TabIndex        =   25
      Tag             =   "Currency"
      Top             =   1365
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton cmdCalcFeeSched 
      Caption         =   "Base &Fee >>"
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
      Height          =   255
      Left            =   5520
      TabIndex        =   24
      Top             =   1380
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtServiceFeeComment 
      Enabled         =   0   'False
      Height          =   1335
      Left            =   9120
      MaxLength       =   254
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   27
      Top             =   240
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton cmdAddEditIB 
      Caption         =   "&Add IB"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3360
      TabIndex        =   10
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton cmdViewIB 
      Caption         =   "Print I&B"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3360
      TabIndex        =   9
      Top             =   720
      Width           =   1935
   End
   Begin VB.ComboBox cboBillingID 
      Height          =   360
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   5175
   End
   Begin VB.Frame framServiceFee 
      Caption         =   "Billing Selection"
      Height          =   1695
      Left            =   75
      TabIndex        =   0
      Top             =   0
      Width           =   5295
      Begin VB.TextBox txtFeeServiceHourlyRate 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E9E9E9&
         DataField       =   "BillingHours"
         DataSource      =   "Claims"
         Enabled         =   0   'False
         Height          =   360
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   8
         Tag             =   "Currency|NO_SHOW_FRAME"
         Top             =   1200
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtActLogHours 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E9E9E9&
         DataField       =   "BillingHours"
         DataSource      =   "Claims"
         Enabled         =   0   'False
         Height          =   360
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   5
         Tag             =   "HoursInDecimal|NO_SHOW_FRAME"
         Top             =   1200
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkFeeByTime 
         Caption         =   "Fee By Time"
         Enabled         =   0   'False
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
         TabIndex        =   2
         Tag             =   "BillingItem"
         Top             =   675
         Width           =   1335
      End
      Begin VB.CheckBox chkUseActivityTime 
         Caption         =   "Activity Log"
         DataField       =   "UseActivityTime"
         DataSource      =   "Claims"
         Enabled         =   0   'False
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
         Left            =   1560
         TabIndex        =   3
         Tag             =   "BillingItem"
         Top             =   675
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.Label lblFeeServiceHourlyRate 
         Caption         =   "$ Per hour"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1320
         TabIndex        =   7
         Tag             =   "NO_SHOW_FRAME"
         Top             =   960
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblActLogHours 
         Caption         =   "Time:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Tag             =   "NO_SHOW_FRAME"
         Top             =   960
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lblxPerHour 
         Alignment       =   2  'Center
         Caption         =   "X"
         Enabled         =   0   'False
         Height          =   255
         Left            =   840
         TabIndex        =   6
         Tag             =   "NO_SHOW_FRAME"
         Top             =   1260
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin VB.Frame framIBOptions 
      Height          =   1695
      Left            =   5400
      TabIndex        =   13
      Top             =   0
      Width           =   6375
      Begin VB.CheckBox IBOptions 
         Caption         =   "View &Details"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   180
         Width           =   2535
      End
      Begin VB.CheckBox IBOptions 
         Caption         =   "View &Overriding Fees"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   480
         Width           =   2535
      End
      Begin VB.TextBox txtOverrideMiscellaneous 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1590
         Locked          =   -1  'True
         MaxLength       =   14
         TabIndex        =   23
         Tag             =   "Currency"
         ToolTipText     =   "Any amount other than 0.00 will Override the automatically calculated Depreciation"
         Top             =   1095
         Width           =   1935
      End
      Begin VB.TextBox txtOverrideExcessLimit 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1590
         Locked          =   -1  'True
         MaxLength       =   14
         TabIndex        =   21
         Tag             =   "Currency"
         ToolTipText     =   "This amount will override the automatically calculated Gross Loss"
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox txtOverrideDepreciation 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1590
         Locked          =   -1  'True
         MaxLength       =   14
         TabIndex        =   19
         Tag             =   "Currency"
         ToolTipText     =   "Any amount other than 0.00 will Override the automatically calculated Depreciation"
         Top             =   620
         Width           =   1935
      End
      Begin VB.TextBox txtOverrideGrossLoss 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1590
         Locked          =   -1  'True
         MaxLength       =   14
         TabIndex        =   17
         Tag             =   "Currency"
         ToolTipText     =   "This amount will override the automatically calculated Gross Loss"
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox txtComments 
         Enabled         =   0   'False
         Height          =   735
         Left            =   4440
         MaxLength       =   150
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   26
         Top             =   840
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CheckBox ChkOverrideAmounts 
         Alignment       =   1  'Right Justify
         Caption         =   "OFF"
         DataField       =   "UseActivityTime"
         DataSource      =   "Claims"
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
         Height          =   210
         Left            =   2880
         TabIndex        =   15
         Tag             =   "BillingItem"
         Top             =   150
         Width           =   615
      End
      Begin VB.Label lblMiscellaneous 
         Caption         =   "Miscellaneous"
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
         Height          =   195
         Left            =   120
         TabIndex        =   22
         ToolTipText     =   "Any amount other than 0.00 will Override the automatically calculated Depreciation"
         Top             =   1133
         Width           =   1455
      End
      Begin VB.Label lblExcessLimit 
         Caption         =   "Excess Limit"
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
         Height          =   195
         Left            =   120
         TabIndex        =   20
         ToolTipText     =   "Any amount other than 0.00 will Override the automatically calculated Depreciation"
         Top             =   878
         Width           =   1455
      End
      Begin VB.Label lblGrossLoss 
         Caption         =   "Gross Loss"
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
         Height          =   195
         Left            =   120
         TabIndex        =   16
         ToolTipText     =   "Any amount other than 0.00 will Override the automatically calculated Gross Loss"
         Top             =   398
         Width           =   1335
      End
      Begin VB.Label lblDepreciation 
         Caption         =   "Depreciation"
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
         Height          =   195
         Left            =   120
         TabIndex        =   18
         ToolTipText     =   "Any amount other than 0.00 will Override the automatically calculated Depreciation"
         Top             =   658
         Width           =   1455
      End
      Begin VB.Label lblOverrideIndemnityAmounts 
         Caption         =   "Override Amounts"
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
         Height          =   195
         Left            =   120
         TabIndex        =   14
         ToolTipText     =   "Any amount other than 0.00 will Override the automatically calculated Gross Loss"
         Top             =   180
         Width           =   1815
      End
   End
   Begin VB.Frame framServiceFees 
      Caption         =   "Service Fees"
      Enabled         =   0   'False
      Height          =   1575
      Left            =   80
      TabIndex        =   28
      Top             =   1680
      Width           =   9015
      Begin VB.ComboBox lstAddServiceFeeItemNumberOfItems 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmBillingInfo.frx":0004
         Left            =   5160
         List            =   "frmBillingInfo.frx":0006
         Style           =   1  'Simple Combo
         TabIndex        =   31
         Tag             =   "NO_SHOW_FRAME"
         Top             =   260
         Width           =   1755
      End
      Begin VB.TextBox txtAddServiceFeeItemAmount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Left            =   6945
         MaxLength       =   14
         TabIndex        =   32
         Tag             =   "Currency|NO_SHOW_FRAME"
         Top             =   260
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txtAddServiceFeeItemComment 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Left            =   480
         MaxLength       =   50
         TabIndex        =   30
         Tag             =   "NO_SHOW_FRAME"
         Top             =   260
         Visible         =   0   'False
         Width           =   4640
      End
      Begin VB.TextBox txtTtlServiceFee 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E9E9E9&
         Enabled         =   0   'False
         Height          =   375
         Left            =   6945
         Locked          =   -1  'True
         MaxLength       =   14
         TabIndex        =   36
         Tag             =   "Currency"
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox txtMiscServiceFee 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   375
         Left            =   5160
         MaxLength       =   14
         TabIndex        =   35
         Tag             =   "Currency"
         Top             =   1080
         Width           =   1700
      End
      Begin VB.TextBox txtMiscServiceFeeComment 
         Enabled         =   0   'False
         Height          =   375
         Left            =   2400
         MaxLength       =   50
         TabIndex        =   34
         Top             =   1080
         Width           =   2655
      End
      Begin MSComctlLib.ListView lvwServiceFees 
         Height          =   735
         Left            =   120
         TabIndex        =   29
         Tag             =   "Enable"
         Top             =   240
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   1296
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         OLEDragMode     =   1
         FlatScrollBar   =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         Icons           =   "imgBillingStatus"
         SmallIcons      =   "imgBillingStatus"
         ColHdrIcons     =   "imgBillingStatus"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
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
      Begin VB.Label lblMiscServiceFee 
         Alignment       =   1  'Right Justify
         Caption         =   "Other Misc Service Fee:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   1140
         Width           =   2175
      End
   End
   Begin VB.Frame framExpenses 
      Caption         =   "Expense Fees"
      Enabled         =   0   'False
      Height          =   1575
      Left            =   80
      TabIndex        =   37
      Top             =   3240
      Width           =   9015
      Begin VB.ComboBox lstAddExpenseFeeItemNumberOfItems 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5160
         Style           =   1  'Simple Combo
         TabIndex        =   40
         Tag             =   "NO_SHOW_FRAME"
         Top             =   255
         Width           =   1755
      End
      Begin VB.TextBox txtAddExpenseFeeItemComment 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Left            =   480
         MaxLength       =   50
         TabIndex        =   39
         Tag             =   "NO_SHOW_FRAME"
         Top             =   260
         Visible         =   0   'False
         Width           =   4640
      End
      Begin VB.TextBox txtAddExpenseFeeItemAmount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Left            =   6945
         MaxLength       =   14
         TabIndex        =   41
         Tag             =   "Currency|NO_SHOW_FRAME"
         Top             =   260
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txtTtlExpenses 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E9E9E9&
         Enabled         =   0   'False
         Height          =   375
         Left            =   6945
         Locked          =   -1  'True
         MaxLength       =   14
         TabIndex        =   45
         Tag             =   "Currency"
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox txtMiscExpenseFee 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   375
         Left            =   5160
         MaxLength       =   14
         TabIndex        =   44
         Tag             =   "Currency"
         Top             =   1080
         Width           =   1700
      End
      Begin VB.TextBox txtMiscExpenseFeeComment 
         Enabled         =   0   'False
         Height          =   375
         Left            =   2400
         MaxLength       =   50
         TabIndex        =   43
         Top             =   1080
         Width           =   2655
      End
      Begin MSComctlLib.ListView lvwExpenseFees 
         Height          =   735
         Left            =   120
         TabIndex        =   38
         Tag             =   "Enable"
         Top             =   240
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   1296
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         OLEDragMode     =   1
         FlatScrollBar   =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         Icons           =   "imgBillingStatus"
         SmallIcons      =   "imgBillingStatus"
         ColHdrIcons     =   "imgBillingStatus"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
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
      Begin VB.Label lblMiscExpenseFee 
         Alignment       =   1  'Right Justify
         Caption         =   "Other Misc Expense Fee:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   1140
         Width           =   2295
      End
   End
   Begin VB.Frame framTotals 
      Caption         =   "Invoice Total"
      Enabled         =   0   'False
      Height          =   1695
      Left            =   80
      TabIndex        =   46
      Top             =   4800
      Width           =   9015
      Begin VB.Timer Timer_PickIB 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   1560
         Top             =   840
      End
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
         Picture         =   "frmBillingInfo.frx":0008
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Exit"
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox txtTaxPercent 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   375
         Left            =   5160
         MaxLength       =   7
         TabIndex        =   51
         Tag             =   "TaxPercent"
         Top             =   720
         Width           =   1700
      End
      Begin VB.TextBox txtTaxAmount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E9E9E9&
         Enabled         =   0   'False
         Height          =   375
         Left            =   6945
         Locked          =   -1  'True
         MaxLength       =   14
         TabIndex        =   52
         Tag             =   "Currency"
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox txtTotalAdjustingFee 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E9E9E9&
         Enabled         =   0   'False
         Height          =   375
         Left            =   6945
         Locked          =   -1  'True
         MaxLength       =   14
         TabIndex        =   54
         Tag             =   "Currency"
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox txtTtlServiceExp 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E9E9E9&
         Enabled         =   0   'False
         Height          =   375
         Left            =   6945
         Locked          =   -1  'True
         MaxLength       =   14
         TabIndex        =   49
         Tag             =   "Currency"
         Top             =   240
         Width           =   1935
      End
      Begin MSComctlLib.ImageList imgBillingStatus 
         Left            =   2640
         Top             =   360
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
               Picture         =   "frmBillingInfo.frx":044A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBillingInfo.frx":05A4
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBillingInfo.frx":0990
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBillingInfo.frx":0B01
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBillingInfo.frx":0F2F
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBillingInfo.frx":12E0
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label lblTaxPercent 
         Alignment       =   1  'Right Justify
         Caption         =   "Tax Percent:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3840
         TabIndex        =   50
         Top             =   780
         Width           =   1215
      End
      Begin VB.Label lblTotalAdjustingFee 
         Alignment       =   1  'Right Justify
         Caption         =   "Total Invoice:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3960
         TabIndex        =   53
         Top             =   1260
         Width           =   2895
      End
      Begin VB.Label lblTtlServiceExp 
         Alignment       =   1  'Right Justify
         Caption         =   "Total Expense and Service Fee:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3960
         TabIndex        =   48
         Top             =   300
         Width           =   2895
      End
   End
   Begin VB.Frame framCommands 
      Height          =   1695
      Left            =   9120
      TabIndex        =   70
      Top             =   4800
      Width           =   2655
      Begin VB.CommandButton cmdDateClosed 
         Enabled         =   0   'False
         Height          =   375
         Left            =   2160
         Picture         =   "frmBillingInfo.frx":16F2
         Style           =   1  'Graphical
         TabIndex        =   72
         Tag             =   "Date"
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtdtDateClosed 
         Enabled         =   0   'False
         Height          =   360
         Left            =   525
         TabIndex        =   71
         Tag             =   "Date"
         Top             =   360
         Width           =   1680
      End
      Begin VB.CommandButton cmdExit 
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
         Left            =   1560
         MaskColor       =   &H00000000&
         Picture         =   "frmBillingInfo.frx":1B34
         Style           =   1  'Graphical
         TabIndex        =   74
         ToolTipText     =   "Exit"
         Top             =   720
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
         Left            =   525
         MaskColor       =   &H00000000&
         Picture         =   "frmBillingInfo.frx":1E3E
         Style           =   1  'Graphical
         TabIndex        =   73
         ToolTipText     =   "Exit"
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lbldtDateClosed 
         Caption         =   "Close date for this IB"
         Enabled         =   0   'False
         Height          =   255
         Left            =   525
         TabIndex        =   75
         Top             =   150
         Width           =   1935
      End
   End
   Begin VB.Frame framOverrideFees 
      Caption         =   "Overriding Fees"
      Enabled         =   0   'False
      Height          =   3135
      Left            =   9120
      TabIndex        =   66
      Top             =   1680
      Visible         =   0   'False
      Width           =   2655
      Begin VB.TextBox txtOverrideFeeItemAmount 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   600
         MaxLength       =   14
         TabIndex        =   68
         Tag             =   "Currency|NO_SHOW_FRAME"
         Top             =   260
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txtOverrideFeeItemComment 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         MaxLength       =   50
         TabIndex        =   67
         Tag             =   "NO_SHOW_FRAME"
         Top             =   260
         Visible         =   0   'False
         Width           =   2235
      End
      Begin MSComctlLib.ListView lvwOverrideFees 
         Height          =   2775
         Left            =   120
         TabIndex        =   69
         Tag             =   "Enable"
         Top             =   240
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   4895
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         OLEDragMode     =   1
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         Icons           =   "imgBillingStatus"
         SmallIcons      =   "imgBillingStatus"
         ColHdrIcons     =   "imgBillingStatus"
         ForeColor       =   -2147483640
         BackColor       =   -2147483624
         BorderStyle     =   1
         Appearance      =   0
         Enabled         =   0   'False
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
   Begin VB.Frame framIBDetails 
      Caption         =   "IB Details"
      Enabled         =   0   'False
      Height          =   3135
      Left            =   9120
      TabIndex        =   57
      Top             =   1680
      Width           =   2655
      Begin VB.TextBox txtDateLastUpdated 
         BackColor       =   &H00E9E9E9&
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
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   59
         Top             =   1080
         Width           =   2400
      End
      Begin VB.CommandButton cmdChangeFeeSchedule 
         Caption         =   "Change Fee Schedule"
         Enabled         =   0   'False
         Height          =   360
         Left            =   120
         TabIndex        =   65
         Top             =   2640
         Width           =   2415
      End
      Begin VB.ComboBox cboFeeSchedule 
         Height          =   360
         ItemData        =   "frmBillingInfo.frx":1F88
         Left            =   120
         List            =   "frmBillingInfo.frx":1F8A
         Style           =   2  'Dropdown List
         TabIndex        =   64
         Top             =   2640
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Image imgAdminComments 
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   255
         Left            =   2280
         Stretch         =   -1  'True
         Tag             =   "NO_SHOW_FRAME"
         Top             =   1920
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image imgUploadMe 
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   255
         Left            =   2280
         Stretch         =   -1  'True
         Tag             =   "NO_SHOW_FRAME"
         Top             =   1680
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image imgVoid 
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   255
         Left            =   2280
         Stretch         =   -1  'True
         Tag             =   "NO_SHOW_FRAME"
         Top             =   1440
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblVoid 
         Alignment       =   1  'Right Justify
         Caption         =   "VOID IB:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   60
         Tag             =   "NO_SHOW_FRAME"
         Top             =   1440
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label lblUploadMe 
         Alignment       =   1  'Right Justify
         Caption         =   "Upload Me:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   61
         Tag             =   "NO_SHOW_FRAME"
         Top             =   1680
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label lblAdminComments 
         Alignment       =   1  'Right Justify
         Caption         =   "Admin Comments:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   62
         Tag             =   "NO_SHOW_FRAME"
         Top             =   1920
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label lblDateLastUpdated 
         Caption         =   "Date Last Updated"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   58
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label lblFeeSchedule 
         BackColor       =   &H00E9E9E9&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   120
         TabIndex        =   63
         Top             =   2280
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frmBillingInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Overidfees Got focus and Lost focus size
Private Const ORFEES_FRAM_HEIGHT_ONFOCUS As Long = 3695
Private Const ORFEES_LVW_HEIGHT_ONFOCUS As Long = 3335
Private Const ORFEES_FRAM_HEIGHT_LOSTFOCUS As Long = 1215
Private Const ORFEES_LVW_HEIGHT_LOSTFOCUS As Long = 855
'Overidfees Got focus and Lost focus size

'OverrideFees VBFormula
Private Const VBFORMULA_OVERRIDES_ALL As String = "OVERRIDES_ALL"
Private Const VBFORMULA_OVERRIDES_SERVICE_FEE As String = "OVERRIDES_SERVICE_FEE"
Private Const VBFORMULA_OVERRIDES_FEEBYTIME_FEE As String = "OVERRIDES_FEEBYTIME_FEE"
'OverrideFees VBFormula

'Option View
Private Const OPT_VIEW_IB_DETAILS As Long = 0
Private Const OPT_VIEW_IB_OVERRIDING As Long = 1
'Option View

Private mlIgnoreMouseMove As Long

Private mfrmClaim As frmClaim
Private mbUnloadMe As Boolean
Private msAssignmentsID As String
Private moGUI As V2ECKeyBoard.clsCarGUI
Private mbEditMode As Boolean
Private mbLoading As Boolean 'Loading Form
Private mbLoadingMe As Boolean 'Just loading data
Private mbPopulateIBInfo As Boolean
Private msBillingCountID As String
Private mbClosedIB As Boolean
Private mSelectedOverrideFeeItemX As ListItem
Private mbEditOverrideFeesItemX As Boolean
Private mSelectedAddServiceFeeItemX As ListItem
Private mbEditServiceFeeItemX As Boolean
Private mSelectedAddExpenseFeeItemX As ListItem
Private mbEditExpenseFeeItemX As Boolean
Private mbOverridesServiceFee As Boolean 'True only on loading an IB that is set to manual override Service fee (That means Fee Schedule Calc is Disabled)
Private mcOverridesServiceFeeAmount As Currency
Private mbOverridesFeeByTimeFee As Boolean 'True only on loading an IB that is set to manual override Fee BybTime (That means that activity log time is not used)
Private mcOverridesFeeByTimeFeeAmount As Currency
Private mdblOverridesActLogTime As Double
Private mbOverridesALL As Boolean
Private mbServiceFeeCommentIsVisible As Boolean
'Current Values
Private mcFeeServiceHourlyRate As Currency
Private mdblCurrentActLogTime As Double
Private mlCurrentPhotoCount As Long
Private mcCurrentAmountOfCheck As Currency
Private mcCurrentServiceFee As Currency
Private mdblCurrentTaxPercent As Double
Private msAdminComments As String
'Current Items not shown on Form but will show on printed IB
Private mcOverrideGrossLoss As Currency
Private mcCurrentGrossLoss As Currency
Private mcCurrentDepreciation As Currency
Private mcCurrentDeductible As Currency
Private mcCurrentLessExcessLimits As Currency
Private mcCurrentLessMiscellaneous As Currency
Private mcCurrnetNetClaim As Currency
'Current Admin items not shown on from but will show on printed IB
Private msCurrentSubToCarrier As String
Private msCurrentIBNumber As String
Private msCurrentLocation As String
Private msCurrentState As String
Private msCurrentAdjusterName As String
Private msCurrentSALN As String
Private msCurrentPolicyNo As String
Private msCurrentInsuredName As String
Private msCurrentLossLocation As String
Private msCurrentDateOfLoss As String
Private msFeeScheduleID As String




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

Private Sub cboBillingID_Click()
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
            sMess = "Do you want to Save Changes?" & vbCrLf & vbCrLf & Me.Caption
'            If MsgBox(sMess, vbQuestion + vbYesNo, "Save Changes") = vbYes Then
                Screen.MousePointer = MousePointerConstants.vbHourglass
                SaveMe
                Screen.MousePointer = MousePointerConstants.vbDefault
                cmdSave.Enabled = False
                lbldtDateClosed.Visible = cmdSave.Enabled
                txtdtDateClosed.Visible = cmdSave.Enabled
                cmdDateClosed.Visible = cmdSave.Enabled
                cmdChangeFeeSchedule.Enabled = False
'            End If
        End If
    End If
    
    For lCount = 0 To cboBillingID.ListCount - 1
        sIBListText = sIBListText & cboBillingID.List(lCount) & vbCrLf
    Next
    
    If cboBillingID.ListIndex = -1 Or cboBillingID.ListIndex = 0 Then
        msBillingCountID = 0
        'Only show the Add button if there is nothing currently
        'open.  Adjuster must close the currnt bill before adding another.
        If InStr(1, sIBListText, "Current", vbTextCompare) = 0 Then
            cmdAddEditIB.Visible = True
            cmdAddEditIB.Caption = "&Add IB"
            cmdAddEditIB.Enabled = True
        Else
            cmdAddEditIB.Visible = False
        End If
        cmdViewIB.Visible = False
    Else
        msBillingCountID = cboBillingID.ItemData(cboBillingID.ListIndex)
        sIBText = cboBillingID.List(cboBillingID.ListIndex)
        If InStr(1, sIBText, "Current", vbTextCompare) > 0 Then
            mbClosedIB = False
            cmdAddEditIB.Visible = True
            cmdAddEditIB.Enabled = True
            cmdAddEditIB.Caption = "&Edit IB"
        ElseIf InStr(1, sIBText, "Closed", vbTextCompare) > 0 Then
            mbClosedIB = True
            If InStr(1, sIBListText, "Current", vbTextCompare) = 0 Then
                cmdAddEditIB.Visible = True
                cmdAddEditIB.Caption = "&Rebill IB"
                cmdAddEditIB.Enabled = True
            Else
                cmdAddEditIB.Visible = True
                cmdAddEditIB.Caption = "&Rebill IB"
                cmdAddEditIB.Enabled = False
            End If
        End If
        cmdViewIB.Enabled = True
        cmdViewIB.Visible = True
    End If
    
    InitValues
    
    Exit Sub
EH:
    Screen.MousePointer = MousePointerConstants.vbDefault
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cboBillingID_Change"
End Sub

Private Sub InitValues()
    On Error GoTo EH
    
    'Reset these
    mbPopulateIBInfo = True
    chkFeeByTime.Value = vbUnchecked
    chkUseActivityTime.Value = vbUnchecked
    ChkOverrideAmounts.Value = vbUnchecked
    mbPopulateIBInfo = False
    txtActLogHours.Text = vbNullString
    txtFeeServiceHourlyRate.Text = vbNullString
    lvwOverrideFees.ListItems.Clear
    lblFeeSchedule.Caption = vbNullString
    txtServiceFeeComment.Text = vbNullString
    txtServiceFee.Text = vbNullString
    lvwServiceFees.ListItems.Clear
    txtMiscServiceFeeComment.Text = vbNullString
    txtMiscServiceFee.Text = vbNullString
    txtTtlServiceFee.Text = vbNullString
    lvwExpenseFees.ListItems.Clear
    txtMiscExpenseFeeComment.Text = vbNullString
    txtMiscExpenseFee.Text = vbNullString
    txtTtlExpenses.Text = vbNullString
    txtdtDateClosed.Text = vbNullString
    imgUploadMe.Picture = LoadPicture()
    imgVoid.Picture = LoadPicture()
    imgAdminComments.Picture = LoadPicture()
    txtDateLastUpdated.Text = vbNullString
    txtTtlServiceExp.Text = vbNullString
    txtTaxPercent.Text = vbNullString
    txtTaxAmount.Text = vbNullString
    txtTotalAdjustingFee.Text = vbNullString
    HideAddExpenseFeeItem
    HideAddServiceFeeItem
    HideOverrideFeeItem
'    If Not txtOverrideFeeItemComment.Visible Then
'        framOverrideFees.Height = ORFEES_FRAM_HEIGHT_LOSTFOCUS
'        lvwOverrideFees.Height = ORFEES_LVW_HEIGHT_LOSTFOCUS
'    End If
    EnableEditFrames False
    cmdSave.Enabled = False
    lbldtDateClosed.Visible = cmdSave.Enabled
    txtdtDateClosed.Visible = cmdSave.Enabled
    cmdDateClosed.Visible = cmdSave.Enabled
    cmdChangeFeeSchedule.Enabled = False
    cmdPrintIBReport.Visible = False
    cboCopy.Visible = False
    txtComments.Visible = False
    txtComments.Text = vbNullString
    
    'Reset Memeber flags
    mbEditOverrideFeesItemX = False
    mbEditServiceFeeItemX = False
    mbEditExpenseFeeItemX = False
    mbOverridesServiceFee = False
    mbOverridesFeeByTimeFee = False
    mbOverridesALL = False
    msFeeScheduleID = vbNullString
    
    'reset Current values
    'Current Values
    mcFeeServiceHourlyRate = 0
    mdblCurrentActLogTime = 0
    mlCurrentPhotoCount = 0
    mcCurrentAmountOfCheck = 0
    mcCurrentServiceFee = 0
    mdblCurrentTaxPercent = 0
    msAdminComments = vbNullString
    'Current Items not shown on Form but will show on printed IB
    mcCurrentGrossLoss = 0
    mcCurrentDepreciation = 0
    mcCurrentDeductible = 0
    mcCurrentLessExcessLimits = 0
    mcCurrentLessMiscellaneous = 0
    mcCurrnetNetClaim = 0
    'Current Admin items not shown on from but will show on printed IB
    msCurrentSubToCarrier = vbNullString
    msCurrentIBNumber = vbNullString
    msCurrentLocation = vbNullString
    msCurrentState = vbNullString
    msCurrentAdjusterName = vbNullString
    msCurrentSALN = vbNullString
    msCurrentPolicyNo = vbNullString
    msCurrentInsuredName = vbNullString
    msCurrentLossLocation = vbNullString
    msCurrentDateOfLoss = vbNullString
    msFeeScheduleID = vbNullString
        
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub InitValues"
End Sub

Private Sub EnableEditFrames(pbEnabled As Boolean)
    On Error GoTo EH
    
    framServiceFee.Enabled = pbEnabled
    framIBOptions.Enabled = pbEnabled
'    framOverrideFees.Enabled = pbEnabled
    If Not mbOverridesServiceFee And Not mbOverridesFeeByTimeFee And Not mbOverridesALL Then
        chkFeeByTime.Enabled = pbEnabled
        cmdCalcFeeSched.Enabled = pbEnabled
        ChkOverrideAmounts.Enabled = pbEnabled
    Else
        cmdCalcFeeSched.Enabled = False
        chkFeeByTime.Enabled = False
        ChkOverrideAmounts.Enabled = False
    End If
    chkFeeByTime.Enabled = pbEnabled
    txtServiceFeeComment.Enabled = pbEnabled
    txtServiceFeeComment.Visible = pbEnabled
    If chkFeeByTime.Value = vbChecked Then
        cmdCalcFeeSched.Visible = False
    Else
        cmdCalcFeeSched.Visible = pbEnabled
    End If
    txtServiceFee.Enabled = pbEnabled
    txtServiceFee.Visible = pbEnabled
    framServiceFees.Enabled = pbEnabled
    framExpenses.Enabled = pbEnabled
    framTotals.Enabled = pbEnabled
    cmdChangeFeeSchedule.Enabled = pbEnabled
    
    ShowFrame
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub EnableEditFrames"
End Sub

Private Sub cboBillingID_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo EH
    
    Select Case KeyCode
        Case KeyCodeConstants.vbKeyReturn, KeyCodeConstants.vbKeySpace
            If cmdAddEditIB.Enabled And cmdAddEditIB.Visible And StrComp(cmdAddEditIB.Caption, "&Edit IB", vbTextCompare) = 0 Then
                cmdAddEditIB_Click
            End If
    End Select
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cboBillingID_KeyDown"
End Sub

Private Sub cboFeeSchedule_Click()
    On Error GoTo EH
    Dim sFeeScheduleID As String
    
    If cboFeeSchedule.ListIndex <= 0 Then
        cboFeeSchedule.Visible = False
        If framIBDetails.Enabled Then
            cmdChangeFeeSchedule.Visible = True
        End If
        Exit Sub
    End If
    
    Screen.MousePointer = MousePointerConstants.vbHourglass
    cboFeeSchedule.Visible = False
    
    sFeeScheduleID = cboFeeSchedule.ItemData(cboFeeSchedule.ListIndex)
    
    LoadEdit sFeeScheduleID
    
    cmdChangeFeeSchedule.Visible = True
    Screen.MousePointer = MousePointerConstants.vbDefault
    Exit Sub
EH:
    Screen.MousePointer = MousePointerConstants.vbDefault
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cboFeeSchedule_Click"
End Sub

Private Sub chkFeeByTime_Click()
    On Error GoTo EH
    Dim sMess As String
    Dim sServiceFeeComment As String
    
    If chkFeeByTime.Value = vbChecked Then
        cmdCalcFeeSched.Visible = False
        txtActLogHours.Visible = True
        lblActLogHours.Visible = True
        lblxPerHour.Visible = True
        txtFeeServiceHourlyRate.Visible = True
        lblFeeServiceHourlyRate.Visible = True
    Else
        cmdCalcFeeSched.Visible = True
        txtActLogHours.Visible = False
        lblActLogHours.Visible = False
        lblActLogHours.Visible = False
        lblxPerHour.Visible = False
        txtFeeServiceHourlyRate.Visible = False
        lblFeeServiceHourlyRate.Visible = False
    End If
    
    If mbPopulateIBInfo Then
        Exit Sub
    End If
    
    If mbOverridesFeeByTimeFee Then
        If chkFeeByTime.Value = vbUnchecked Then
            sMess = "The " & framOverrideFees.Caption & " item:  ""Override Fee By Time"" is checked!" & vbCrLf
            sMess = sMess & "This item can not be unchecked until the above item is unchecked."
            MsgBox sMess, vbExclamation, "Fee By Time"
            chkFeeByTime.Value = vbChecked
            Exit Sub
        End If
    End If
    
    If chkFeeByTime.Value = vbChecked Then
        If mbEditMode Then
            If mbOverridesFeeByTimeFee Then
                txtServiceFee.Text = Format(mdblOverridesActLogTime * mcFeeServiceHourlyRate, "#,###,###,##0.00")
                SumServiceFees
                sServiceFeeComment = Format(mdblOverridesActLogTime, "##0.00") & " @ " & Format(mcFeeServiceHourlyRate, "##0.00")
                txtServiceFeeComment.Text = sServiceFeeComment
                txtServiceFeeComment.ToolTipText = sServiceFeeComment
            Else
                txtServiceFee.Text = Format(mdblCurrentActLogTime * mcFeeServiceHourlyRate, "#,###,###,##0.00")
                chkUseActivityTime.Value = vbChecked
                SumServiceFees
                sServiceFeeComment = Format(mdblCurrentActLogTime, "##0.00") & " @ " & Format(mcFeeServiceHourlyRate, "##0.00")
                txtServiceFeeComment.Text = sServiceFeeComment
                txtServiceFeeComment.ToolTipText = sServiceFeeComment
            End If
        End If
    Else
        If mbEditMode Then
            If Not mbOverridesFeeByTimeFee Then
                txtServiceFee.Text = "0.00"
                SumServiceFees
                txtServiceFeeComment.Text = vbNullString
                txtServiceFeeComment.ToolTipText = vbNullString
            End If
        End If
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub chkFeeByTime_Click"
End Sub

Private Sub ChkOverrideAmounts_Click()
    On Error GoTo EH
    
        
    If ChkOverrideAmounts.Value = vbChecked Then
        ChkOverrideAmounts.Caption = "ON"
        txtOverrideGrossLoss.ForeColor = &H8080FF
        txtOverrideGrossLoss.Locked = False
        txtOverrideDepreciation.ForeColor = &H8080FF
        txtOverrideDepreciation.Locked = False
        txtOverrideExcessLimit.ForeColor = &H8080FF
        txtOverrideExcessLimit.Locked = False
        txtOverrideMiscellaneous.ForeColor = &H8080FF
        txtOverrideMiscellaneous.Locked = False
    Else
        ChkOverrideAmounts.Caption = "OFF"
        txtOverrideGrossLoss.ForeColor = &H80000008
        txtOverrideGrossLoss.Locked = True
        txtOverrideDepreciation.ForeColor = &H80000008
        txtOverrideDepreciation.Locked = True
        txtOverrideExcessLimit.ForeColor = &H80000008
        txtOverrideExcessLimit.Locked = True
        txtOverrideMiscellaneous.ForeColor = &H80000008
        txtOverrideMiscellaneous.Locked = True
    End If
    
        
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub ChkOverrideAmounts_Click"
End Sub

'Private Sub chkToggleOverride_Click()
'    On Error GoTo EH
'
'    If chkToggleOverride.Value = vbChecked Then
'        framOverrideFees.Height = ORFEES_FRAM_HEIGHT_ONFOCUS
'        lvwOverrideFees.Height = ORFEES_LVW_HEIGHT_ONFOCUS
'        txtServiceFeeComment.Visible = False
'    Else
'        LooseOverridesFocus
'    End If
'
'
'    Exit Sub
'EH:
'    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub chkToggleOverride_Click"
'End Sub

Private Sub cmdChangeFeeSchedule_Click()
    On Error GoTo EH
    Dim sMess As String
    
    sMess = "Are you sure you want to change the current Fee Schedule ?" & vbCrLf & vbCrLf
    sMess = sMess & "All fee items will be reset after you select the new Fee Schedule."
    
    If MsgBox(sMess, vbQuestion + vbYesNo, "Change Fee Schedule") = vbNo Then
        Exit Sub
    End If
    
    cmdChangeFeeSchedule.Visible = False
    cboFeeSchedule.Visible = True
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub chkFeeByTime_Click"
End Sub

Private Sub cmdDateClosed_Click()
    On Error GoTo EH
    Dim sMess As String
    
'    sMess = "Are you sure you are ready to close this IB?" & vbCrLf & vbCrLf
'    sMess = sMess & "Once Closed, any future corrections will require a Rebill. "
'
'    If MsgBox(sMess, vbYesNo + vbQuestion, "Close IB Date") = vbYes Then
        mfrmClaim.ShowCalendar txtdtDateClosed
'    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdDateClosed_Click"
End Sub

Private Sub cmdAddEditIB_Click()
    On Error GoTo EH
    Dim sCapText As String
    Dim sMess As String
    Dim lRet As Long
    Dim sRet As String
    'Used for Edit IB
    Dim bSelectTheCurrentIB As Boolean
    Dim lCount As Long
    Dim sData As String
    Dim sFeeScheduleID As String
    
    
    cmdViewIB.Visible = False
    cmdAddEditIB.Visible = False
    
    Me.Refresh
    
    sCapText = cmdAddEditIB.Caption
        
    If StrComp(sCapText, "&Rebill IB", vbTextCompare) = 0 Then
        sMess = "Are you sure you want to Rebill this IB ?" & vbCrLf & vbCrLf
        sMess = sMess & "Select ""Yes"" ONLY IF you need to make a correction to this item!"
        lRet = MsgBox(sMess, vbQuestion + vbYesNo, "REBILL IB")
        If lRet = vbNo Then
            Exit Sub
        End If
        Screen.MousePointer = MousePointerConstants.vbHourglass
        msFeeScheduleID = mfrmClaim.GetFeeScheduleID(msAssignmentsID, msBillingCountID, mbClosedIB)
        RebillIB
        Screen.MousePointer = MousePointerConstants.vbDefault
        bSelectTheCurrentIB = True
    ElseIf StrComp(sCapText, "&Add IB", vbTextCompare) = 0 Then
        'Only give Supplemental Question if there is one or more
        'existing bills.  anything over Count 1 is an existing bill.
        If cboBillingID.ListCount > 1 Then
            sMess = "Are you sure you want to do a Supplemental Bill ?" & vbCrLf & vbCrLf
            sMess = sMess & "Select ""Yes"" ONLY IF you need to Add an INTERIM IB!"
            lRet = MsgBox(sMess, vbQuestion + vbYesNo, "INTERIM IB")
            If lRet = vbNo Then
                Exit Sub
            Else
                sMess = "You must enter a password from your Manager to Supplement this Claim!"
                sRet = InputBox(sMess, "Enter Password to Supplement this claim!")
                If sRet <> Format(Now(), "DDYYMM") Then
                    If sRet = vbNullString Then
                        Exit Sub
                    End If
                    sMess = "Invalid Password!"
                    MsgBox sMess, vbExclamation, "Invalid Password"
                    Exit Sub
                End If
            End If
        End If
        Screen.MousePointer = MousePointerConstants.vbHourglass
        SupplementIB
        Screen.MousePointer = MousePointerConstants.vbDefault
        bSelectTheCurrentIB = True
    ElseIf StrComp(sCapText, "&Edit IB", vbTextCompare) = 0 Then
        sFeeScheduleID = mfrmClaim.GetFeeScheduleID(msAssignmentsID, msBillingCountID, mbClosedIB)
        Screen.MousePointer = MousePointerConstants.vbHourglass
        LoadEdit sFeeScheduleID
        Screen.MousePointer = MousePointerConstants.vbDefault
    End If
    
    'if the user just click add or Rebill...
    'If so then select the Billing that was just made current... save a few clicks
    If bSelectTheCurrentIB Then
        For lCount = 1 To cboBillingID.ListCount - 1
            sData = cboBillingID.List(lCount)
            If InStr(1, sData, "Current", vbTextCompare) > 0 Then
                cboBillingID.ListIndex = lCount
                Exit For
            End If
        Next
        sCapText = cmdAddEditIB.Caption
        If StrComp(sCapText, "&Edit IB", vbTextCompare) = 0 Then
            cmdAddEditIB_Click
        End If
    End If
    
    Exit Sub
EH:
    Screen.MousePointer = MousePointerConstants.vbDefault
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdAddEditIB_Click"
End Sub

Private Sub LoadEdit(psFeeScheduleID As String)
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim RSFeeScheduleFeeTypes As ADODB.Recordset
    Dim MyRTIBFeeItem As GuiRTIBFeeItem
    Dim sSQL As String
    Dim sRTIBFeeID As String
    
    'Load current values to get anything such as Activity Log times that have changed,
    'Photo count changes
    'Tax percent on fee schedule change
    'Hourly Rate for fee by time change
    msFeeScheduleID = psFeeScheduleID
    LoadCurrentValues
    mbEditMode = True
    'First need to add any new FeeItems that have not already been added.
    Set oConn = New ADODB.Connection
    Set RSFeeScheduleFeeTypes = New ADODB.Recordset
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    sSQL = "SELECT RetFeeScheduleFeeTypes.* "
    sSQL = sSQL & "FROM( "
    sSQL = sSQL & "SELECT "
    sSQL = sSQL & "[FeeScheduleFeeTypesID], "
    sSQL = sSQL & "[FeeScheduleID], "
    sSQL = sSQL & "[TypeNum], "
    sSQL = sSQL & "[Name], "
    sSQL = sSQL & "[Description], "
    sSQL = sSQL & "[FeeAmount], "
    sSQL = sSQL & "[IsExpense], "
    sSQL = sSQL & "[MaxNumberOfItems], "
    sSQL = sSQL & "[MaxFeeAmount], "
    sSQL = sSQL & "[IsMiscAmount], "
    sSQL = sSQL & "[UseFormula], "
    sSQL = sSQL & "[VBFormula], "
    sSQL = sSQL & "[IsDeleted], "
    sSQL = sSQL & "[DateLastUpdated], "
    sSQL = sSQL & "(SELECT    USERNAME "
    sSQL = sSQL & "FROM   USERS "
    sSQL = sSQL & "WHERE  USERSID = F.[UpdateByUserID]) As [UpdateByUserName],  "
    sSQL = sSQL & "[UpdateByUserID] "
    sSQL = sSQL & "FROM FeeScheduleFeeTypes F "
    sSQL = sSQL & ") RetFeeScheduleFeeTypes "
    sSQL = sSQL & "WHERE FeeScheduleFeeTypesID Not IN ( "
                                    sSQL = sSQL & "SELECT   FeeScheduleFeeTypesID "
                                    sSQL = sSQL & "FROM     RTIBfee "
                                    sSQL = sSQL & "WHERE    AssignmentsID = " & msAssignmentsID & " "
                                    sSQL = sSQL & ") "
    If msFeeScheduleID = vbNullString Then
        sSQL = sSQL & "AND [FeeScheduleID] = ( "
                                sSQL = sSQL & "SELECT   [FeeScheduleID] "
                                sSQL = sSQL & "FROM     ClientCompanyCat "
                                sSQL = sSQL & "WHERE    [ClientCompanyID] = " & goUtil.gsCurCar & " "
                                sSQL = sSQL & "AND      [CATID] = " & goUtil.gsCurCat & " "
                                sSQL = sSQL & ") "
    Else
        sSQL = sSQL & "AND [FeeScheduleID] = ( "
                                sSQL = sSQL & "SELECT   " & msFeeScheduleID & " As [FeeScheduleID] "
                                sSQL = sSQL & "FROM     ClientCompanyCat "
                                sSQL = sSQL & "WHERE    [ClientCompanyID] = " & goUtil.gsCurCar & " "
                                sSQL = sSQL & "AND      [CATID] = " & goUtil.gsCurCat & " "
                                sSQL = sSQL & ") "
    End If
    'Exclude any feetypes that already exist in the RTIBFee table for this Assignment
    sSQL = sSQL & "ORDER BY [TypeNum] "
    
    RSFeeScheduleFeeTypes.CursorLocation = adUseClient
    RSFeeScheduleFeeTypes.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set RSFeeScheduleFeeTypes.ActiveConnection = Nothing
    
    If RSFeeScheduleFeeTypes.RecordCount > 0 Then
        RSFeeScheduleFeeTypes.MoveFirst
    End If
    Do Until RSFeeScheduleFeeTypes.EOF
        With MyRTIBFeeItem
            .RTIBFeeID = "[RTIBFeeID]"
            .AssignmentsID = msAssignmentsID
            .ID = "[ID]"
            .IDAssignments = msAssignmentsID
            .FeeScheduleFeeTypesID = goUtil.IsNullIsVbNullString(RSFeeScheduleFeeTypes.Fields("FeeScheduleFeeTypesID"))
            .NumberOfItems = "0"
            .Amount = "0.00"
            .Comment = vbNullString
            .DownLoadMe = "False"
            .UpLoadMe = "True"
            .AdminComments = vbNullString
            .DateLastUpdated = Now()
            .UpdateByUserID = goUtil.gsCurUsersID
        End With
        'Since this Feetypw does not exist yet need to insert it
        If Not AddRTIBFee(MyRTIBFeeItem, sRTIBFeeID) Then
            GoTo CLEAN_UP
        End If
        
        RSFeeScheduleFeeTypes.MoveNext
    Loop
    
    
    If mbClosedIB Then
        PopulateClosedBillingInfo
    Else
        PopulateOpenBillingInfo
    End If
    EnableEditFrames True
    If txtComments.Visible And txtComments.Enabled Then
        txtComments.SetFocus
    End If
    
CLEAN_UP:

    'cleanup
    Set oConn = Nothing
    Set RSFeeScheduleFeeTypes = Nothing
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub LoadEdit"
End Sub

Private Sub cmdCalcFeeSched_Click()
    On Error GoTo EH
    'Calculate FeeSchedule
    If Not CalcFeeSched() Then
        Exit Sub
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdCalcFeeSched_Click"
End Sub

Private Sub cmdExit_Click()
    On Error GoTo EH
    Dim sMess As String
    
    If cmdSave.Enabled Then
'        sMess = "Do you want to Save Changes?" & vbCrLf & vbCrLf & Me.Caption
'        If MsgBox(sMess, vbQuestion + vbYesNo, "Save Changes") = vbNo Then
'            cmdSave.Enabled = False
'        End If
    End If
    
    mbUnloadMe = True
    Me.Visible = False
    mfrmClaim.Timer_UnloadForm.Enabled = True
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdExit_Click"
End Sub

Private Sub cmdPrintIBReport_Click()
    On Error GoTo EH
        
    Screen.MousePointer = MousePointerConstants.vbHourglass
    'First be sure control are valid
    
    cmdPrintIBReport.Enabled = False
    cboCopy.Enabled = False
    If mfrmClaim.PrintActiveReport(cboBillingID, , cboCopy.Text) Then
        If Not mbUnloadMe Then
            cmdPrintIBReport.Enabled = True
            cboCopy.Enabled = True
        End If
    End If
    Screen.MousePointer = MousePointerConstants.vbDefault
    
    Exit Sub
EH:
    Screen.MousePointer = MousePointerConstants.vbDefault
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdPrintIBReport_Click"
End Sub

Private Sub cmdPrintList_Click()
    On Error GoTo EH
    
    If lvwServiceFees.Visible Then
        goUtil.utPrintListView App.EXEName, lvwServiceFees, "Service Fees"
    End If
    
    If lvwExpenseFees.Visible Then
        goUtil.utPrintListView App.EXEName, lvwExpenseFees, "Expense Fees"
    End If
    
    If lvwOverrideFees.Visible Then
        goUtil.utPrintListView App.EXEName, lvwOverrideFees, "Overriding Fees"
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdPrintList_Click"
End Sub

Private Sub cmdSave_Click()
    On Error GoTo EH
    cmdSave.Enabled = False
    Screen.MousePointer = MousePointerConstants.vbArrowHourglass
    If SaveMe() Then
        mfrmClaim.RefreshMe
    End If
    Screen.MousePointer = MousePointerConstants.vbDefault
    Exit Sub
EH:
    Screen.MousePointer = MousePointerConstants.vbDefault
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdSave_Click"
End Sub

Private Sub cmdViewIB_Click()
    On Error GoTo EH
    
    cmdViewIB.Visible = False
    cmdAddEditIB.Visible = False
    
    Me.Refresh
    
    Screen.MousePointer = MousePointerConstants.vbHourglass
    
    mbEditMode = False
    
    'Load current values to get anything such as Activity Log times that have changed,
    'Photo count changes
    'Tax percent on fee schedule change
    'Hourly Rate for fee by time change
    msFeeScheduleID = mfrmClaim.GetFeeScheduleID(msAssignmentsID, msBillingCountID, mbClosedIB)
    LoadCurrentValues
    If mbClosedIB Then
        PopulateClosedBillingInfo
    Else
        PopulateOpenBillingInfo
    End If
    
    lblTtlServiceExp.Visible = True
    txtTtlServiceExp.Visible = True
    lblTaxPercent.Visible = True
    txtTaxPercent.Visible = True
    txtTaxAmount.Visible = True
    lblTotalAdjustingFee.Visible = True
    txtTotalAdjustingFee.Visible = True
    txtServiceFeeComment.Visible = True
    txtServiceFee.Visible = True
    
    
    cmdSave.Enabled = False
    lbldtDateClosed.Visible = cmdSave.Enabled
    txtdtDateClosed.Visible = cmdSave.Enabled
    cmdDateClosed.Visible = cmdSave.Enabled
    cmdChangeFeeSchedule.Enabled = False
    cmdPrintIBReport.Visible = True
    LoadCopy
    cboCopy.Visible = True
'    txtComments.Visible = True   'NOT USING COMMENTS AT THIS TIME (Could be used for Billing Admin purposes)
    
    If cmdPrintIBReport.Visible And cmdPrintIBReport.Enabled Then
        cmdPrintIBReport.SetFocus
    End If
    
    Screen.MousePointer = MousePointerConstants.vbDefault
    
    Exit Sub
EH:
    Screen.MousePointer = MousePointerConstants.vbDefault
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdViewIB_Click"
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
    Me.Icon = mfrmClaim.optClaim(GuiClaimOptions.opt04_BillingInformation).Picture
    
    LoadHeaderlvwOverrideFees
    LoadHeaderlvwServiceFees
    LoadHeaderlvwExpenseFees
    msBillingCountID = 0
    
    LoadMe
    
    ShowFrame
    
    cmdSave.Enabled = False
    lbldtDateClosed.Visible = cmdSave.Enabled
    txtdtDateClosed.Visible = cmdSave.Enabled
    cmdDateClosed.Visible = cmdSave.Enabled
    cmdChangeFeeSchedule.Enabled = False
    mbLoading = False
    Timer_PickIB.Enabled = True
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_Load"
End Sub

Public Function LoadMe() As Boolean
    On Error GoTo EH
    
    mbLoadingMe = True
    
    RefreshBilling
    RefreshFeeSchedule
    
    
    LoadMe = True
    mbLoadingMe = False
    Exit Function
EH:
    mbLoadingMe = False
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function LoadMe"
End Function

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

Public Function SaveMe() As Boolean
    On Error GoTo EH
    'RTIB Items
    Dim MyRTIBItem As GuiRTIBItem
    Dim MyIBItem As GuiIBItem
    Dim sIDIB As String
    Dim RSAssgn As ADODB.Recordset
    Dim RSACID As ADODB.Recordset
    Dim RSCatCode As ADODB.Recordset
    Dim RSClientCoCat As ADODB.Recordset
    Dim RSFeeSchedule As ADODB.Recordset
    Dim RSBillingCountItem As ADODB.Recordset
    Dim sFeeScheduleID As String
    Dim sClientCompanyID As String
    Dim sSection As String
    Dim sDateClosed As String
    Dim lCurVoidImage As Long
    Dim sVoid As String
    Dim sFeeByTime As String
    Dim sUseActLogTime As String
    Dim sFlagText As String
    'RTIBFee Items
    Dim MyRTIBFeeItem As GuiRTIBFeeItem
    Dim MyIBFeeItem As GuiIBFeeItem
    Dim itmX As ListItem
    Dim sRTIBFeeID As String
    Dim sIBFeeID As String
    'PackageItem
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    Dim RSPackageItem As ADODB.Recordset
    Dim sPackageItemName As String
    Dim sNewPackageItemName As String
    Dim sPackageItemID As String
    
    'be sure to finish any serve fee items that were left in Editi mode
    If txtOverrideFeeItemAmount.Visible Then
        goUtil.utValidate , txtOverrideFeeItemAmount
        CalcOverrideFeeItem
    End If
    If txtAddServiceFeeItemAmount.Visible Then
        goUtil.utValidate , txtAddServiceFeeItemAmount
        CalcAddServiceFeeItem
    End If
    If txtAddExpenseFeeItemAmount.Visible Then
        goUtil.utValidate , txtAddExpenseFeeItemAmount
        CalcAddExpenseFeeItem
    End If

    'validate all the fields on this form
    goUtil.utValidate Me
    
    'Need to get updated info to populate stuff on the IB
    If Not mfrmClaim.SetadoRSAssignments(msAssignmentsID) Then
        GoTo CLEAN_UP
    End If
    If Not moGUI.SetadoRSACID Then
        GoTo CLEAN_UP
    End If
    If Not moGUI.SetadoRSCatCode Then
        GoTo CLEAN_UP
    End If
    If Not mfrmClaim.SetadoRSClientCOCat Then
        GoTo CLEAN_UP
    End If
    If Not moGUI.SetadoRSFeeSchedule Then
        GoTo CLEAN_UP
    End If
    If Not mfrmClaim.SetadoRSBillingCountItem(msAssignmentsID, msBillingCountID, False) Then
        GoTo CLEAN_UP
    End If
    
    Set RSAssgn = mfrmClaim.adoRSAssignments
    Set RSACID = moGUI.adoRSACID
    Set RSCatCode = moGUI.adoRSCatCode
    Set RSClientCoCat = mfrmClaim.adoRSClientCOCat
    Set RSFeeSchedule = moGUI.adoFeeSchedule
    Set RSBillingCountItem = mfrmClaim.adoRSBillingCountItem
    
    'Used for Getting Reg Setting init options
    If msFeeScheduleID = vbNullString Then
        sFeeScheduleID = goUtil.IsNullIsVbNullString(RSFeeSchedule.Fields("FeeScheduleID"))
    Else
        sFeeScheduleID = msFeeScheduleID
    End If
    sClientCompanyID = goUtil.IsNullIsVbNullString(RSFeeSchedule.Fields("ClientCompanyID"))
    sSection = sFeeScheduleID & "_" & sClientCompanyID
    
    'Date CLosed
    If txtdtDateClosed.Text = vbNullString Then
        sDateClosed = "Null"
    Else
        sDateClosed = txtdtDateClosed.Text
    End If
    
    'Check for VOID Flag
    lCurVoidImage = imgBillingStatus.ListImages(GuiIBStatusList.IsDeleted).Picture
    
    If imgVoid.Picture = lCurVoidImage Then
        sVoid = "True"
    Else
        sVoid = "False"
    End If
    
    'Fee By Time
    If chkFeeByTime.Value = vbChecked Then
        sFeeByTime = "True"
    Else
        sFeeByTime = "False"
    End If
    'use Act Log Time
    If chkUseActivityTime.Value = vbChecked Then
        sUseActLogTime = "True"
    Else
        sUseActLogTime = "False"
    End If
    
    'Need to Save to RTIB First
    With MyRTIBItem
        .AssignmentsID = msAssignmentsID
        .BillingCountID = msBillingCountID
        .IDAssignments = msAssignmentsID
        .IDBillingCount = msBillingCountID
        .RT00_lSSN = goUtil.utGetECSCryptSetting("ECS", "WEB_SECURITY", "SSN")
        .RT01_sSubToCarrier = goUtil.IsNullIsVbNullString(RSACID.Fields("ClientCompanyDesc"))
        .RT02_sIBNumber = goUtil.IsNullIsVbNullString(RSAssgn.Fields("IBNUM"))
        .RT05_sLocation = goUtil.IsNullIsVbNullString(RSClientCoCat.Fields("SACity"))
        .RT05a_sState = goUtil.IsNullIsVbNullString(RSClientCoCat.Fields("SAState"))
        .RT06_dtDateClosed = sDateClosed
        .RT07_sAdjusterName = goUtil.IsNullIsVbNullString(RSACID.Fields("LFName"))
        .RT09_sSALN = goUtil.IsNullIsVbNullString(RSAssgn.Fields("CLIENTNUM"))
        .RT09a_sPolicyNo = goUtil.IsNullIsVbNullString(RSAssgn.Fields("PolicyNo"))
        .RT10_sInsuredName = goUtil.IsNullIsVbNullString(RSAssgn.Fields("Insured"))
        .RT11_sLossLocation = goUtil.IsNullIsVbNullString(RSAssgn.Fields("PAStreet")) & vbCrLf
        .RT11_sLossLocation = .RT11_sLossLocation & goUtil.IsNullIsVbNullString(RSAssgn.Fields("PACity")) & ", "
        .RT11_sLossLocation = .RT11_sLossLocation & goUtil.IsNullIsVbNullString(RSAssgn.Fields("PAState")) & " "
        .RT11_sLossLocation = .RT11_sLossLocation & Format(goUtil.IsNullIsVbNullString(RSAssgn.Fields("PAZIP")), "00000") & " - "
        .RT11_sLossLocation = .RT11_sLossLocation & Format(goUtil.IsNullIsVbNullString(RSAssgn.Fields("PAZIP4")), "0000")
        .RT12_dtDateOfLoss = IIf(goUtil.IsNullIsVbNullString(RSAssgn.Fields("LossDate")) = vbNullString, "Null", goUtil.IsNullIsVbNullString(RSAssgn.Fields("LossDate")))
        If ChkOverrideAmounts.Value = vbChecked Then
            .RT13_cGrossLoss = txtOverrideGrossLoss.Text
            .RT14_cDepreciation = txtOverrideDepreciation.Text
        Else
            .RT13_cGrossLoss = mfrmClaim.GetFullCostOfRepair(CLng(msBillingCountID))
            .RT14_cDepreciation = mfrmClaim.GetDepreciation(CLng(msBillingCountID))
        End If
        .RT14a_sSupplement = goUtil.IsNullIsVbNullString(RSBillingCountItem.Fields("Supplement"))
        .RT14b_sRebilled = goUtil.IsNullIsVbNullString(RSBillingCountItem.Fields("Rebill"))
        .RT15_cDeductible = goUtil.IsNullIsVbNullString(RSAssgn.Fields("Deductible"))
        If ChkOverrideAmounts.Value = vbChecked Then
            .RT15a_cLessExcessLimits = txtOverrideExcessLimit.Text
            .RT15c_cLessMiscellaneous = txtOverrideMiscellaneous
        Else
            .RT15a_cLessExcessLimits = mfrmClaim.GetLessExcessLimits(CLng(msBillingCountID))
            .RT15c_cLessMiscellaneous = mfrmClaim.GetLessMiscellaneous(CLng(msBillingCountID))
        End If
        .RT15b_sExcessLimDesc = vbNullString
        .RT15d_cMiscellaneousDesc = vbNullString
        If ChkOverrideAmounts.Value = vbChecked Then
            .RT16_cNetClaim = mfrmClaim.GetNetActualCashValueClaim(CLng(msBillingCountID), True, CCur(.RT13_cGrossLoss), CCur(.RT14_cDepreciation), CCur(.RT15a_cLessExcessLimits), CCur(.RT15c_cLessMiscellaneous))
        Else
            .RT16_cNetClaim = mfrmClaim.GetNetActualCashValueClaim(CLng(msBillingCountID))
        End If
        .RT17_cServiceFee = txtServiceFee.Text
        .RT17a_cMiscServiceFee = txtMiscServiceFee.Text
        .RT18_sServiceFeeComment = txtServiceFeeComment.Text
        .RT18a_sMiscServiceFeeComment = txtMiscServiceFeeComment.Text
        .RT25_cServiceFeeSubTotal = txtTtlServiceFee.Text
        .RT29a_sMiscExpenseFeeComment = txtMiscExpenseFeeComment.Text
        .RT29b_cMiscExpenseFee = txtMiscExpenseFee.Text
        .RT30_cTotalExpenses = txtTtlExpenses.Text
        .RT31_dTaxPercent = txtTaxPercent.Text
        .RT32_cTaxAmount = txtTaxAmount.Text
        .RT33_cTotalAdjustingFee = txtTotalAdjustingFee.Text
        .RT33a_sAccountCode = CStr(ChkOverrideAmounts.Value)
        .FeeScheduleID = msFeeScheduleID
        .Void = sVoid
        .FeeByTime = sFeeByTime
        .UseActivityTime = sUseActLogTime
        .DownLoadMe = "False"
        .UpLoadMe = "True"
        .Comments = txtComments.Text
        .AdminComments = msAdminComments
        .DateLastUpdated = Now()
        .UpdateByUserID = goUtil.gsCurUsersID
    End With
    
    'Save the RTIB
    If EditRTIB(MyRTIBItem) Then
        'Need to Update RTIbfee items
        'Overriding Items
        For Each itmX In lvwOverrideFees.ListItems
            With MyRTIBFeeItem
                .RTIBFeeID = "[RTIBFeeID]"
                .AssignmentsID = msAssignmentsID
                .ID = "[ID]"
                .IDAssignments = msAssignmentsID
                .FeeScheduleFeeTypesID = itmX.ListSubItems(GuiOverrideFeesListView.FeeScheduleFeeTypesID - 1)
                .NumberOfItems = itmX.ListSubItems(GuiOverrideFeesListView.NumberOfItems - 1)
                .Amount = CCur(itmX.ListSubItems(GuiOverrideFeesListView.Amount - 1))
                .Comment = itmX.ListSubItems(GuiOverrideFeesListView.Comment - 1)
                .DownLoadMe = "False"
                sFlagText = itmX.ListSubItems(GuiOverrideFeesListView.UpLoadMe - 1)
                .UpLoadMe = goUtil.GetFlagFromText(sFlagText)
                .AdminComments = itmX.ListSubItems(GuiOverrideFeesListView.AdminComments - 1)
                .DateLastUpdated = itmX.ListSubItems(GuiOverrideFeesListView.DateLastUpdated - 1)
                .UpdateByUserID = itmX.ListSubItems(GuiOverrideFeesListView.UpdateByUserID - 1)
            End With
            If Not EditRTIBFee(MyRTIBFeeItem) Then
                If Not AddRTIBFee(MyRTIBFeeItem, sRTIBFeeID) Then
                    GoTo CLEAN_UP
                End If
            End If
        Next
        'ServiceFees Items
        For Each itmX In lvwServiceFees.ListItems
            With MyRTIBFeeItem
                .RTIBFeeID = "[RTIBFeeID]"
                .AssignmentsID = msAssignmentsID
                .ID = "[ID]"
                .IDAssignments = msAssignmentsID
                .FeeScheduleFeeTypesID = itmX.ListSubItems(GuiServiceFeesListView.FeeScheduleFeeTypesID - 1)
                .NumberOfItems = itmX.ListSubItems(GuiServiceFeesListView.NumberOfItems - 1)
                .Amount = CCur(itmX.ListSubItems(GuiServiceFeesListView.Amount - 1))
                .Comment = itmX.ListSubItems(GuiServiceFeesListView.Comment - 1)
                .DownLoadMe = "False"
                sFlagText = itmX.ListSubItems(GuiServiceFeesListView.UpLoadMe - 1)
                .UpLoadMe = goUtil.GetFlagFromText(sFlagText)
                .AdminComments = itmX.ListSubItems(GuiServiceFeesListView.AdminComments - 1)
                .DateLastUpdated = itmX.ListSubItems(GuiServiceFeesListView.DateLastUpdated - 1)
                .UpdateByUserID = itmX.ListSubItems(GuiServiceFeesListView.UpdateByUserID - 1)
            End With
            If Not EditRTIBFee(MyRTIBFeeItem) Then
                If Not AddRTIBFee(MyRTIBFeeItem, sRTIBFeeID) Then
                    GoTo CLEAN_UP
                End If
            End If
        Next
        'ExpenseFees Items
        For Each itmX In lvwExpenseFees.ListItems
            With MyRTIBFeeItem
                .RTIBFeeID = "[RTIBFeeID]"
                .AssignmentsID = msAssignmentsID
                .ID = "[ID]"
                .IDAssignments = msAssignmentsID
                .FeeScheduleFeeTypesID = itmX.ListSubItems(GuiExpenseFeesListView.FeeScheduleFeeTypesID - 1)
                .NumberOfItems = itmX.ListSubItems(GuiExpenseFeesListView.NumberOfItems - 1)
                .Amount = CCur(itmX.ListSubItems(GuiExpenseFeesListView.Amount - 1))
                .Comment = itmX.ListSubItems(GuiExpenseFeesListView.Comment - 1)
                .DownLoadMe = "False"
                sFlagText = itmX.ListSubItems(GuiExpenseFeesListView.UpLoadMe - 1)
                .UpLoadMe = goUtil.GetFlagFromText(sFlagText)
                .AdminComments = itmX.ListSubItems(GuiExpenseFeesListView.AdminComments - 1)
                .DateLastUpdated = itmX.ListSubItems(GuiExpenseFeesListView.DateLastUpdated - 1)
                .UpdateByUserID = itmX.ListSubItems(GuiExpenseFeesListView.UpdateByUserID - 1)
            End With
            If Not EditRTIBFee(MyRTIBFeeItem) Then
                If Not AddRTIBFee(MyRTIBFeeItem, sRTIBFeeID) Then
                    GoTo CLEAN_UP
                End If
            End If
        Next
    End If
    
    'If the close date is set then need to Update the IB table
    'If its a new IB then Insert is required. If its Rebilled IB then
    'edit of the Old IB is Required
    
    If sDateClosed = "Null" Then
        GoTo SKIP_IB_SAVE
    End If
    
       'Need to Save Closed IB First
    With MyIBItem
        .IBID = "[IBID]" 'Set in AddIB
        .AssignmentsID = MyRTIBItem.AssignmentsID
        .BillingCountID = MyRTIBItem.BillingCountID
        .ID = "[ID]" 'Set in AddIB
        .IDAssignments = MyRTIBItem.IDAssignments
        .IDBillingCount = MyRTIBItem.IDBillingCount
        .IB00_lssn = MyRTIBItem.RT00_lSSN
        .IB01_sSubToCarrier = MyRTIBItem.RT01_sSubToCarrier
        .IB02_sIBNumber = MyRTIBItem.RT02_sIBNumber
        .IB05_sLocation = MyRTIBItem.RT05_sLocation
        .IB05a_sState = MyRTIBItem.RT05a_sState
        .IB06_dtDateClosed = MyRTIBItem.RT06_dtDateClosed
        .IB07_sAdjusterName = MyRTIBItem.RT07_sAdjusterName
        .IB09_sSALN = MyRTIBItem.RT09_sSALN
        .IB09a_sPolicyNo = MyRTIBItem.RT09a_sPolicyNo
        .IB10_sInsuredName = MyRTIBItem.RT10_sInsuredName
        .IB11_sLossLocation = MyRTIBItem.RT11_sLossLocation
        .IB12_dtDateOfLoss = MyRTIBItem.RT12_dtDateOfLoss
        .IB13_cGrossLoss = MyRTIBItem.RT13_cGrossLoss
        .IB14_cDepreciation = MyRTIBItem.RT14_cDepreciation
        .IB14a_sSupplement = MyRTIBItem.RT14a_sSupplement
        .IB14b_sRebilled = MyRTIBItem.RT14b_sRebilled
        .IB15_cDeductible = MyRTIBItem.RT15_cDeductible
        .IB15a_cLessExcessLimits = MyRTIBItem.RT15a_cLessExcessLimits
        .IB15b_sExcessLimDesc = MyRTIBItem.RT15b_sExcessLimDesc
        .IB15c_cLessMiscellaneous = MyRTIBItem.RT15c_cLessMiscellaneous
        .IB15d_cMiscellaneousDesc = MyRTIBItem.RT15d_cMiscellaneousDesc
        .IB16_cNetClaim = MyRTIBItem.RT16_cNetClaim
        .IB17_cServiceFee = MyRTIBItem.RT17_cServiceFee
        .IB17a_cMiscServiceFee = MyRTIBItem.RT17a_cMiscServiceFee
        .IB18_sServiceFeeComment = MyRTIBItem.RT18_sServiceFeeComment
        .IB18a_sMiscServiceFeeComment = MyRTIBItem.RT18a_sMiscServiceFeeComment
        .IB25_cServiceFeeSubTotal = MyRTIBItem.RT25_cServiceFeeSubTotal
        .IB29a_sMiscExpenseFeeComment = MyRTIBItem.RT29a_sMiscExpenseFeeComment
        .IB29b_cMiscExpenseFee = MyRTIBItem.RT29b_cMiscExpenseFee
        .IB30_cTotalExpenses = MyRTIBItem.RT30_cTotalExpenses
        .IB31_dTaxPercent = MyRTIBItem.RT31_dTaxPercent
        .IB32_cTaxAmount = MyRTIBItem.RT32_cTaxAmount
        .IB33_cTotalAdjustingFee = MyRTIBItem.RT33_cTotalAdjustingFee
        .IB33a_sAccountCode = MyRTIBItem.RT33a_sAccountCode
        .FeeScheduleID = MyRTIBItem.FeeScheduleID
        .Void = MyRTIBItem.Void
        .FeeByTime = MyRTIBItem.FeeByTime
        .UseActivityTime = MyRTIBItem.UseActivityTime
        .DownLoadMe = MyRTIBItem.DownLoadMe
        .UpLoadMe = MyRTIBItem.UpLoadMe
        .Comments = MyRTIBItem.Comments
        .AdminComments = MyRTIBItem.AdminComments
        .DateLastUpdated = MyRTIBItem.DateLastUpdated
        .UpdateByUserID = MyRTIBItem.UpdateByUserID
    End With
    
    'Save the Closed IB
    
    If Not EditIB(MyIBItem, sIDIB) Then
        If Not AddIB(MyIBItem, sIDIB) Then
            GoTo CLEAN_UP
        End If
    End If
    
    'After saving check to see if this was a rebilled item
    'if it was need to check the Current packageitem for this Fee bill item
    'Need to update the Name to reflect the Rebilleed suffix
    If MyRTIBItem.RT14b_sRebilled > 0 Then
        mfrmClaim.SetadoRSPackageItem msAssignmentsID, vbNullString, MyRTIBItem.RT14a_sSupplement
        Set RSPackageItem = mfrmClaim.adoRSPackageItem
        If RSPackageItem.RecordCount > 0 Then
            Do Until RSPackageItem.EOF
                sPackageItemName = goUtil.IsNullIsVbNullString(RSPackageItem.Fields("Name"))
                If sPackageItemName <> vbNullString Then
                    sNewPackageItemName = MyRTIBItem.RT02_sIBNumber
                    sNewPackageItemName = sNewPackageItemName & IIf(MyRTIBItem.RT14a_sSupplement = "0", vbNullString, "S" & MyRTIBItem.RT14a_sSupplement)
                    sNewPackageItemName = sNewPackageItemName & IIf(MyRTIBItem.RT14b_sRebilled = "0", vbNullString, "R" & MyRTIBItem.RT14b_sRebilled)
                    'Need to update the Format of the Feebill name
                    If StrComp(sPackageItemName, sNewPackageItemName, vbTextCompare) <> 0 Then
                        If oConn Is Nothing Then
                            Set oConn = New ADODB.Connection
                            goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
                        End If
                        sPackageItemID = goUtil.IsNullIsVbNullString(RSPackageItem.Fields("PackageItemID"))
                        sSQL = "UPDATE PackageItem SET "
                        sSQL = sSQL & "[Name] = '" & sNewPackageItemName & "', "
                        sSQL = sSQL & "[UploadMe] = True, "
                        sSQL = sSQL & "[DateLastUpdated] = #" & Now() & "#, "
                        sSQL = sSQL & "[UpdateByUserID] = " & goUtil.gsCurUsersID & " "
                        sSQL = sSQL & "WHERE [AssignmentsID] = " & msAssignmentsID & " "
                        sSQL = sSQL & "AND [PackageItemID] = " & sPackageItemID & " "
                        oConn.Execute sSQL
                        Sleep 100
                    End If
                End If
                RSPackageItem.MoveNext
            Loop
        End If
    End If
    
    
    'Need to Update the Closed Ibfee items
    'Overriding Items
    For Each itmX In lvwOverrideFees.ListItems
        With MyIBFeeItem
            .IBFeeID = "[IBFeeID]"
            .AssignmentsID = msAssignmentsID
            .IBID = sIDIB
            .ID = "[ID]"
            .IDAssignments = msAssignmentsID
            .IDIB = sIDIB
            .FeeScheduleFeeTypesID = itmX.ListSubItems(GuiOverrideFeesListView.FeeScheduleFeeTypesID - 1)
            .NumberOfItems = itmX.ListSubItems(GuiOverrideFeesListView.NumberOfItems - 1)
            .Amount = CCur(itmX.ListSubItems(GuiOverrideFeesListView.Amount - 1))
            .Comment = itmX.ListSubItems(GuiOverrideFeesListView.Comment - 1)
            .DownLoadMe = "False"
            sFlagText = itmX.ListSubItems(GuiOverrideFeesListView.UpLoadMe - 1)
            .UpLoadMe = goUtil.GetFlagFromText(sFlagText)
            .AdminComments = itmX.ListSubItems(GuiOverrideFeesListView.AdminComments - 1)
            .DateLastUpdated = itmX.ListSubItems(GuiOverrideFeesListView.DateLastUpdated - 1)
            .UpdateByUserID = itmX.ListSubItems(GuiOverrideFeesListView.UpdateByUserID - 1)
        End With
        If Not EditIBFee(MyIBFeeItem) Then
            If Not AddIBFee(MyIBFeeItem, sIBFeeID) Then
                GoTo CLEAN_UP
            End If
        End If
    Next
    'ServiceFees Items
    For Each itmX In lvwServiceFees.ListItems
        With MyIBFeeItem
            .IBFeeID = "[IBFeeID]"
            .AssignmentsID = msAssignmentsID
            .IBID = sIDIB
            .ID = "[ID]"
            .IDAssignments = msAssignmentsID
            .IDIB = sIDIB
            .FeeScheduleFeeTypesID = itmX.ListSubItems(GuiServiceFeesListView.FeeScheduleFeeTypesID - 1)
            .NumberOfItems = itmX.ListSubItems(GuiServiceFeesListView.NumberOfItems - 1)
            .Amount = CCur(itmX.ListSubItems(GuiServiceFeesListView.Amount - 1))
            .Comment = itmX.ListSubItems(GuiServiceFeesListView.Comment - 1)
            .DownLoadMe = "False"
            sFlagText = itmX.ListSubItems(GuiServiceFeesListView.UpLoadMe - 1)
            .UpLoadMe = goUtil.GetFlagFromText(sFlagText)
            .AdminComments = itmX.ListSubItems(GuiServiceFeesListView.AdminComments - 1)
            .DateLastUpdated = itmX.ListSubItems(GuiServiceFeesListView.DateLastUpdated - 1)
            .UpdateByUserID = itmX.ListSubItems(GuiServiceFeesListView.UpdateByUserID - 1)
        End With
        If Not EditIBFee(MyIBFeeItem) Then
            If Not AddIBFee(MyIBFeeItem, sIBFeeID) Then
                GoTo CLEAN_UP
            End If
        End If
    Next
    'ExpenseFees Items
    For Each itmX In lvwExpenseFees.ListItems
        With MyIBFeeItem
            .IBFeeID = "[IBFeeID]"
            .AssignmentsID = msAssignmentsID
            .IBID = sIDIB
            .ID = "[ID]"
            .IDAssignments = msAssignmentsID
            .IDIB = sIDIB
            .FeeScheduleFeeTypesID = itmX.ListSubItems(GuiExpenseFeesListView.FeeScheduleFeeTypesID - 1)
            .NumberOfItems = itmX.ListSubItems(GuiExpenseFeesListView.NumberOfItems - 1)
            .Amount = CCur(itmX.ListSubItems(GuiExpenseFeesListView.Amount - 1))
            .Comment = itmX.ListSubItems(GuiExpenseFeesListView.Comment - 1)
            .DownLoadMe = "False"
            sFlagText = itmX.ListSubItems(GuiExpenseFeesListView.UpLoadMe - 1)
            .UpLoadMe = goUtil.GetFlagFromText(sFlagText)
            .AdminComments = itmX.ListSubItems(GuiExpenseFeesListView.AdminComments - 1)
            .DateLastUpdated = itmX.ListSubItems(GuiExpenseFeesListView.DateLastUpdated - 1)
            .UpdateByUserID = itmX.ListSubItems(GuiExpenseFeesListView.UpdateByUserID - 1)
        End With
        If Not EditIBFee(MyIBFeeItem) Then
            If Not AddIBFee(MyIBFeeItem, sIBFeeID) Then
                GoTo CLEAN_UP
            End If
        End If
    Next
    
SKIP_IB_SAVE:
    
    cmdSave.Enabled = False
    lbldtDateClosed.Visible = cmdSave.Enabled
    txtdtDateClosed.Visible = cmdSave.Enabled
    cmdDateClosed.Visible = cmdSave.Enabled
    cmdChangeFeeSchedule.Enabled = False
    SaveMe = True
    
CLEAN_UP:

    Set RSAssgn = Nothing
    Set RSACID = Nothing
    Set RSCatCode = Nothing
    Set RSClientCoCat = Nothing
    Set RSFeeSchedule = Nothing
    Set RSBillingCountItem = Nothing
    Set RSPackageItem = Nothing
    Set oConn = Nothing
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function SaveMe"
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo EH
    Dim sMess As String
    
    Select Case UnloadMode
        Case vbFormControlMenu
            sMess = "Are you sure you want to Exit Billing Information?"
'            If MsgBox(sMess, vbQuestion + vbOKCancel, "Exit Billing Information") = vbCancel Then
'                Cancel = True
'                Exit Sub
'            End If
            If cmdSave.Enabled Then
'                sMess = "Do you want to Save Changes?" & vbCrLf & vbCrLf & Me.Caption
'                If MsgBox(sMess, vbQuestion + vbYesNo, "Save Changes") = vbNo Then
'                    cmdSave.Enabled = False
'                End If
            End If
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
    Dim lHeightDiff As Long
    Dim lWidthDiff As Long
    
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
    'Width
    'IB Options
    framIBOptions.Width = Me.Width - 5565
    
    'Fee Overrides frame
    framOverrideFees.Width = Me.Width - 9285
    lvwOverrideFees.Width = Me.Width - 9525
    
    'IB Details Frame
    framIBDetails.Width = Me.Width - 9285
   
    txtDateLastUpdated.Width = Me.Width - 9540
    lblVoid.Width = Me.Width - 9885
    imgVoid.left = Me.Width - 9660
    lblUploadMe.Width = Me.Width - 9885
    imgUploadMe.left = Me.Width - 9660
    lblAdminComments.Width = Me.Width - 9885
    imgAdminComments.left = Me.Width - 9660
    lblFeeSchedule.Width = Me.Width - 9525
    cmdChangeFeeSchedule.Width = Me.Width - 9525
    cboFeeSchedule.Width = Me.Width - 9525
    
    'Commands Save, Exit
    framCommands.Width = Me.Width - 9285
    
    IBOptions(0).left = Me.Width - 8220
    IBOptions(1).left = Me.Width - 8220
    
    'Left
    cmdSave.left = Me.Width - 11415
    cmdExit.left = Me.Width - 10380
    lbldtDateClosed.left = Me.Width - 11415
    txtdtDateClosed.left = Me.Width - 11415
    cmdDateClosed.left = Me.Width - 9780
    
    'Check Height Diff and divide to figure Tops and Heights
    framTotals.top = Me.Height - 2160
    If Me.Height - 3720 > 0 Then
        'Tops
        lHeightDiff = Me.Height - 3720
        lHeightDiff = lHeightDiff - 3720
        framExpenses.top = 3720 + (lHeightDiff / 2)
        'Heights
        framServiceFees.Height = framExpenses.top - 1665
        framExpenses.Height = 1575 + (lHeightDiff / 2)
    Else
        'Default Tops
        framExpenses.top = 3240
        'Default heights
        framServiceFees = 1575
        framExpenses.Height = 1575
    End If
    
    'tops
    framCommands.top = Me.Height - 2160
    lblMiscServiceFee.top = framServiceFees.Height - 435
    txtMiscServiceFeeComment.top = framServiceFees.Height - 435
    txtMiscServiceFee.top = framServiceFees.Height - 435
    txtTtlServiceFee.top = framServiceFees.Height - 435
    lblMiscExpenseFee.top = framExpenses.Height - 435
    txtMiscExpenseFeeComment.top = framExpenses.Height - 435
    txtMiscExpenseFee.top = framExpenses.Height - 435
    txtTtlExpenses.top = framExpenses.Height - 435
    
    'Heights
    framOverrideFees.Height = Me.Height - 3825
    lvwOverrideFees.Height = Me.Height - 4185
    framIBDetails.Height = Me.Height - 3825
    lvwServiceFees.Height = framServiceFees.Height - 840
    lvwExpenseFees.Height = framExpenses.Height - 840
    
    Set mSelectedOverrideFeeItemX = Nothing
    Set mSelectedAddServiceFeeItemX = Nothing
    Set mSelectedAddExpenseFeeItemX = Nothing
    HideAddExpenseFeeItem
    HideAddServiceFeeItem
    HideOverrideFeeItem
    lvwOverrideFees.Enabled = True
    lvwServiceFees.Enabled = True
    lvwExpenseFees.Enabled = True

    If Me.WindowState <> vbMinimized Then
        Me.Visible = True
    End If
End Sub

Public Function CLEANUP() As Boolean
    On Error GoTo EH
    
    Set mSelectedAddServiceFeeItemX = Nothing
    Set mSelectedAddExpenseFeeItemX = Nothing
    Set mSelectedOverrideFeeItemX = Nothing
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

Public Sub LoadHeaderlvwOverrideFees()
    On Error GoTo EH
    Dim bGridOn As Boolean
    Dim bHideDeleted As Boolean
    Dim bHideUploadFlags As Boolean
    
    bHideDeleted = CBool(GetSetting(App.EXEName, "GENERAL", "HIDE_DELETED", True))
    bHideUploadFlags = CBool(GetSetting(App.EXEName, "GENERAL", "HIDE_UPLOAD_FLAGS", True))
    
    'set the columnheaders
    With lvwOverrideFees
        .ColumnHeaders.Add , "fsftDescription", "Description"
        .ColumnHeaders.Add , "Comment", "Comment" 'hidden
        .ColumnHeaders.Add , "NumberOfItems", ""
        .ColumnHeaders.Add , "Amount", "Amount"
        .ColumnHeaders.Add , "UpLoadMe", "UpLoadMe" 'hidden
        .ColumnHeaders.Add , "DateLastUpdated", "Date Last Updated" 'hidden
        .ColumnHeaders.Add , "AdminComments", "Admin Comments" 'hidden
        .ColumnHeaders.Add , "DownLoadMe", "DownLoadMe" 'hidden"
        .ColumnHeaders.Add , "RTIBFeeID", "RTIBFeeID" 'hidden"
        .ColumnHeaders.Add , "AssignmentsID", "AssignmentsID" 'hidden"
        .ColumnHeaders.Add , "ID", "ID" 'hidden"
        .ColumnHeaders.Add , "IDAssignments", "IDAssignments" 'hidden"
        .ColumnHeaders.Add , "UpdateByUserID", "UpdateByUserID" 'hidden"
        .ColumnHeaders.Add , "FeeScheduleFeeTypesID", "FeeScheduleFeeTypesID" 'hidden"
        .ColumnHeaders.Add , "fsftFeeScheduleID", "fsftFeeScheduleID" 'hidden"
        .ColumnHeaders.Add , "fsftTypeNum", "fsftTypeNum" 'hidden"
        .ColumnHeaders.Add , "fsftName", "fsftName" 'hidden"
        .ColumnHeaders.Add , "fsftFeeAmount", "fsftFeeAmount" 'hidden"
        .ColumnHeaders.Add , "fsftIsExpense", "fsftIsExpense" 'hidden"
        .ColumnHeaders.Add , "fsftMaxNumberOfItems", "fsftMaxNumberOfItems" 'hidden"
        .ColumnHeaders.Add , "fsftMaxFeeAmount", "fsftMaxFeeAmount" 'hidden"
        .ColumnHeaders.Add , "fsftIsMiscAmount", "fsftIsMiscAmount" 'hidden"
        .ColumnHeaders.Add , "fsftUseFormula", "fsftUseFormula" 'hidden"
        .ColumnHeaders.Add , "fsftVBFormula", "fsftVBFormula" 'hidden"
        .ColumnHeaders.Add , "fsftIsDeleted", "fsftIsDeleted" 'hidden"
        .ColumnHeaders.Add , "fsftDateLastUpdated", "fsftDateLastUpdated" 'hidden"
        .ColumnHeaders.Add , "fsftUpdateByUserID", "fsftUpdateByUserID" 'hidden"
    
        .Sorted = False
        .SortOrder = lvwAscending
'        GuiOverrideFeesListView.fsftDescription = 2324.977
'GuiOverrideFeesListView.Comment = 1170.142

        'fsftDescription
        .ColumnHeaders.Item(GuiOverrideFeesListView.fsftDescription).Width = 2500 'Me.Width - 11850 '3500
        .ColumnHeaders.Item(GuiOverrideFeesListView.fsftDescription).Alignment = lvwColumnLeft
        'Comment
        .ColumnHeaders.Item(GuiOverrideFeesListView.Comment).Width = 0 'Hidden Me.Width - 11850 '3500
        .ColumnHeaders.Item(GuiOverrideFeesListView.Comment).Alignment = lvwColumnLeft
        'NumberOfItems
        .ColumnHeaders.Item(GuiOverrideFeesListView.NumberOfItems).Width = 400
        .ColumnHeaders.Item(GuiOverrideFeesListView.NumberOfItems).Alignment = lvwColumnRight
'        .ColumnHeaders.Item(GuiOverrideFeesListView.NumberOfItems).Icon = GuiIBStatusList.IsUnchecked
        'Amount
        .ColumnHeaders.Item(GuiOverrideFeesListView.Amount).Width = 1000
        .ColumnHeaders.Item(GuiOverrideFeesListView.Amount).Alignment = lvwColumnRight
        'UpLoadMe
        If bHideUploadFlags Then
            .ColumnHeaders.Item(GuiOverrideFeesListView.UpLoadMe).Width = 0 'Hidden 400
        Else
            .ColumnHeaders.Item(GuiOverrideFeesListView.UpLoadMe).Width = 400
        End If
        .ColumnHeaders.Item(GuiOverrideFeesListView.UpLoadMe).Alignment = lvwColumnLeft
        .ColumnHeaders.Item(GuiOverrideFeesListView.UpLoadMe).Icon = GuiIBStatusList.UpLoadMe
        'DateLastUpdated
        .ColumnHeaders.Item(GuiOverrideFeesListView.DateLastUpdated).Width = 0 'Hidden 2200
        .ColumnHeaders.Item(GuiOverrideFeesListView.DateLastUpdated).Alignment = lvwColumnLeft
        'AdminComments
        .ColumnHeaders.Item(GuiOverrideFeesListView.AdminComments).Width = 0 'Hidden 5000
        .ColumnHeaders.Item(GuiOverrideFeesListView.AdminComments).Alignment = lvwColumnLeft
        'DownLoadMe
        .ColumnHeaders.Item(GuiOverrideFeesListView.DownLoadMe).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiOverrideFeesListView.DownLoadMe).Alignment = lvwColumnLeft
        'RTIBFeeID
        .ColumnHeaders.Item(GuiOverrideFeesListView.RTIBFeeID).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiOverrideFeesListView.RTIBFeeID).Alignment = lvwColumnLeft
        'AssignmentsID
        .ColumnHeaders.Item(GuiOverrideFeesListView.AssignmentsID).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiOverrideFeesListView.AssignmentsID).Alignment = lvwColumnLeft
        'ID
        .ColumnHeaders.Item(GuiOverrideFeesListView.ID).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiOverrideFeesListView.ID).Alignment = lvwColumnLeft
        'IDAssignments
        .ColumnHeaders.Item(GuiOverrideFeesListView.IDAssignments).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiOverrideFeesListView.IDAssignments).Alignment = lvwColumnLeft
        'UpdateByUserID
        .ColumnHeaders.Item(GuiOverrideFeesListView.UpdateByUserID).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiOverrideFeesListView.UpdateByUserID).Alignment = lvwColumnLeft
        'FeeScheduleFeeTypesID
        .ColumnHeaders.Item(GuiOverrideFeesListView.FeeScheduleFeeTypesID).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiOverrideFeesListView.FeeScheduleFeeTypesID).Alignment = lvwColumnLeft
        'fsftFeeScheduleID
        .ColumnHeaders.Item(GuiOverrideFeesListView.fsftFeeScheduleID).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiOverrideFeesListView.fsftFeeScheduleID).Alignment = lvwColumnLeft
        'fsftTypeNum
        .ColumnHeaders.Item(GuiOverrideFeesListView.fsftTypeNum).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiOverrideFeesListView.fsftTypeNum).Alignment = lvwColumnLeft
        'fsftName
        .ColumnHeaders.Item(GuiOverrideFeesListView.fsftName).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiOverrideFeesListView.fsftName).Alignment = lvwColumnLeft
        'fsftFeeAmount
        .ColumnHeaders.Item(GuiOverrideFeesListView.fsftFeeAmount).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiOverrideFeesListView.fsftFeeAmount).Alignment = lvwColumnLeft
        'fsftIsExpense
        .ColumnHeaders.Item(GuiOverrideFeesListView.fsftIsExpense).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiOverrideFeesListView.fsftIsExpense).Alignment = lvwColumnLeft
        'fsftMaxNumberOfItems
        .ColumnHeaders.Item(GuiOverrideFeesListView.fsftMaxNumberOfItems).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiOverrideFeesListView.fsftMaxNumberOfItems).Alignment = lvwColumnLeft
        'fsftMaxFeeAmount
        .ColumnHeaders.Item(GuiOverrideFeesListView.fsftMaxFeeAmount).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiOverrideFeesListView.fsftMaxFeeAmount).Alignment = lvwColumnLeft
        'fsftIsMiscAmount
        .ColumnHeaders.Item(GuiOverrideFeesListView.fsftIsMiscAmount).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiOverrideFeesListView.fsftIsMiscAmount).Alignment = lvwColumnLeft
        'fsftUseFormula
        .ColumnHeaders.Item(GuiOverrideFeesListView.fsftUseFormula).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiOverrideFeesListView.fsftUseFormula).Alignment = lvwColumnLeft
        'fsftVBFormula
        .ColumnHeaders.Item(GuiOverrideFeesListView.fsftVBFormula).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiOverrideFeesListView.fsftVBFormula).Alignment = lvwColumnLeft
        'fsftIsDeleted
        .ColumnHeaders.Item(GuiOverrideFeesListView.fsftIsDeleted).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiOverrideFeesListView.fsftIsDeleted).Alignment = lvwColumnLeft
        'fsftDateLastUpdated
        .ColumnHeaders.Item(GuiOverrideFeesListView.fsftDateLastUpdated).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiOverrideFeesListView.fsftDateLastUpdated).Alignment = lvwColumnLeft
        'fsftUpdateByUserID
        .ColumnHeaders.Item(GuiOverrideFeesListView.fsftUpdateByUserID).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiOverrideFeesListView.fsftUpdateByUserID).Alignment = lvwColumnLeft
    End With
    
    bGridOn = CBool(GetSetting(App.EXEName, "GENERAL", "GRID_ON", False))
    lvwOverrideFees.GridLines = bGridOn
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub LoadHeaderlvwOverrideFees"
End Sub

Public Sub LoadHeaderlvwServiceFees()
    On Error GoTo EH
    Dim bGridOn As Boolean
    Dim bHideDeleted As Boolean
    Dim bHideUploadFlags As Boolean
    
    bHideDeleted = CBool(GetSetting(App.EXEName, "GENERAL", "HIDE_DELETED", True))
    bHideUploadFlags = CBool(GetSetting(App.EXEName, "GENERAL", "HIDE_UPLOAD_FLAGS", True))
    
    'set the columnheaders
    With lvwServiceFees
        .ColumnHeaders.Add , "fsftDescription", "Description"
        .ColumnHeaders.Add , "Comment", "Comment" 'hidden
        .ColumnHeaders.Add , "NumberOfItems", "Items"
        .ColumnHeaders.Add , "Amount", "Amount"
        .ColumnHeaders.Add , "UpLoadMe", "UpLoadMe" 'hidden
        .ColumnHeaders.Add , "DateLastUpdated", "Date Last Updated" 'hidden
        .ColumnHeaders.Add , "AdminComments", "Admin Comments" 'hidden
        .ColumnHeaders.Add , "DownLoadMe", "DownLoadMe" 'hidden"
        .ColumnHeaders.Add , "RTIBFeeID", "RTIBFeeID" 'hidden"
        .ColumnHeaders.Add , "AssignmentsID", "AssignmentsID" 'hidden"
        .ColumnHeaders.Add , "ID", "ID" 'hidden"
        .ColumnHeaders.Add , "IDAssignments", "IDAssignments" 'hidden"
        .ColumnHeaders.Add , "UpdateByUserID", "UpdateByUserID" 'hidden"
        .ColumnHeaders.Add , "FeeScheduleFeeTypesID", "FeeScheduleFeeTypesID" 'hidden"
        .ColumnHeaders.Add , "fsftFeeScheduleID", "fsftFeeScheduleID" 'hidden"
        .ColumnHeaders.Add , "fsftTypeNum", "fsftTypeNum" 'hidden"
        .ColumnHeaders.Add , "fsftName", "fsftName" 'hidden"
        .ColumnHeaders.Add , "fsftFeeAmount", "fsftFeeAmount" 'hidden"
        .ColumnHeaders.Add , "fsftIsExpense", "fsftIsExpense" 'hidden"
        .ColumnHeaders.Add , "fsftMaxNumberOfItems", "fsftMaxNumberOfItems" 'hidden"
        .ColumnHeaders.Add , "fsftMaxFeeAmount", "fsftMaxFeeAmount" 'hidden"
        .ColumnHeaders.Add , "fsftIsMiscAmount", "fsftIsMiscAmount" 'hidden"
        .ColumnHeaders.Add , "fsftUseFormula", "fsftUseFormula" 'hidden"
        .ColumnHeaders.Add , "fsftVBFormula", "fsftVBFormula" 'hidden"
        .ColumnHeaders.Add , "fsftIsDeleted", "fsftIsDeleted" 'hidden"
        .ColumnHeaders.Add , "fsftDateLastUpdated", "fsftDateLastUpdated" 'hidden"
        .ColumnHeaders.Add , "fsftUpdateByUserID", "fsftUpdateByUserID" 'hidden"
    
        .Sorted = False
        .SortOrder = lvwAscending
'        GuiServiceFeesListView.fsftDescription = 4320
'GuiServiceFeesListView.Comment = 3435.024
        'fsftDescription
        .ColumnHeaders.Item(GuiServiceFeesListView.fsftDescription).Width = 5000 ' Me.Width - 9320 '6030
        .ColumnHeaders.Item(GuiServiceFeesListView.fsftDescription).Alignment = lvwColumnLeft
        'Comment
        .ColumnHeaders.Item(GuiServiceFeesListView.Comment).Width = 0 'Hidden Me.Width - 10320 '5030
        .ColumnHeaders.Item(GuiServiceFeesListView.Comment).Alignment = lvwColumnLeft
        'NumberOfItems
        .ColumnHeaders.Item(GuiServiceFeesListView.NumberOfItems).Width = 1750
        .ColumnHeaders.Item(GuiServiceFeesListView.NumberOfItems).Alignment = lvwColumnRight
'        .ColumnHeaders.Item(GuiServiceFeesListView.NumberOfItems).Icon = GuiIBStatusList.showDropDown
        'Amount
        .ColumnHeaders.Item(GuiServiceFeesListView.Amount).Width = 2000
        .ColumnHeaders.Item(GuiServiceFeesListView.Amount).Alignment = lvwColumnRight
        'UpLoadMe
        If bHideUploadFlags Then
            .ColumnHeaders.Item(GuiServiceFeesListView.UpLoadMe).Width = 0 'Hidden 400
        Else
            .ColumnHeaders.Item(GuiServiceFeesListView.UpLoadMe).Width = 400
        End If
        .ColumnHeaders.Item(GuiServiceFeesListView.UpLoadMe).Alignment = lvwColumnLeft
        .ColumnHeaders.Item(GuiServiceFeesListView.UpLoadMe).Icon = GuiIBStatusList.UpLoadMe
        'DateLastUpdated
        .ColumnHeaders.Item(GuiServiceFeesListView.DateLastUpdated).Width = 0 'Hidden 2200
        .ColumnHeaders.Item(GuiServiceFeesListView.DateLastUpdated).Alignment = lvwColumnLeft
        'AdminComments
        .ColumnHeaders.Item(GuiServiceFeesListView.AdminComments).Width = 0 'Hidden 5000
        .ColumnHeaders.Item(GuiServiceFeesListView.AdminComments).Alignment = lvwColumnLeft
        'DownLoadMe
        .ColumnHeaders.Item(GuiServiceFeesListView.DownLoadMe).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiServiceFeesListView.DownLoadMe).Alignment = lvwColumnLeft
        'RTIBFeeID
        .ColumnHeaders.Item(GuiServiceFeesListView.RTIBFeeID).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiServiceFeesListView.RTIBFeeID).Alignment = lvwColumnLeft
        'AssignmentsID
        .ColumnHeaders.Item(GuiServiceFeesListView.AssignmentsID).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiServiceFeesListView.AssignmentsID).Alignment = lvwColumnLeft
        'ID
        .ColumnHeaders.Item(GuiServiceFeesListView.ID).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiServiceFeesListView.ID).Alignment = lvwColumnLeft
        'IDAssignments
        .ColumnHeaders.Item(GuiServiceFeesListView.IDAssignments).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiServiceFeesListView.IDAssignments).Alignment = lvwColumnLeft
        'UpdateByUserID
        .ColumnHeaders.Item(GuiServiceFeesListView.UpdateByUserID).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiServiceFeesListView.UpdateByUserID).Alignment = lvwColumnLeft
        'FeeScheduleFeeTypesID
        .ColumnHeaders.Item(GuiServiceFeesListView.FeeScheduleFeeTypesID).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiServiceFeesListView.FeeScheduleFeeTypesID).Alignment = lvwColumnLeft
        'fsftFeeScheduleID
        .ColumnHeaders.Item(GuiServiceFeesListView.fsftFeeScheduleID).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiServiceFeesListView.fsftFeeScheduleID).Alignment = lvwColumnLeft
        'fsftTypeNum
        .ColumnHeaders.Item(GuiServiceFeesListView.fsftTypeNum).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiServiceFeesListView.fsftTypeNum).Alignment = lvwColumnLeft
        'fsftName
        .ColumnHeaders.Item(GuiServiceFeesListView.fsftName).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiServiceFeesListView.fsftName).Alignment = lvwColumnLeft
        'fsftFeeAmount
        .ColumnHeaders.Item(GuiServiceFeesListView.fsftFeeAmount).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiServiceFeesListView.fsftFeeAmount).Alignment = lvwColumnLeft
        'fsftIsExpense
        .ColumnHeaders.Item(GuiServiceFeesListView.fsftIsExpense).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiServiceFeesListView.fsftIsExpense).Alignment = lvwColumnLeft
        'fsftMaxNumberOfItems
        .ColumnHeaders.Item(GuiServiceFeesListView.fsftMaxNumberOfItems).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiServiceFeesListView.fsftMaxNumberOfItems).Alignment = lvwColumnLeft
        'fsftMaxFeeAmount
        .ColumnHeaders.Item(GuiServiceFeesListView.fsftMaxFeeAmount).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiServiceFeesListView.fsftMaxFeeAmount).Alignment = lvwColumnLeft
        'fsftIsMiscAmount
        .ColumnHeaders.Item(GuiServiceFeesListView.fsftIsMiscAmount).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiServiceFeesListView.fsftIsMiscAmount).Alignment = lvwColumnLeft
        'fsftUseFormula
        .ColumnHeaders.Item(GuiServiceFeesListView.fsftUseFormula).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiServiceFeesListView.fsftUseFormula).Alignment = lvwColumnLeft
        'fsftVBFormula
        .ColumnHeaders.Item(GuiServiceFeesListView.fsftVBFormula).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiServiceFeesListView.fsftVBFormula).Alignment = lvwColumnLeft
        'fsftIsDeleted
        .ColumnHeaders.Item(GuiServiceFeesListView.fsftIsDeleted).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiServiceFeesListView.fsftIsDeleted).Alignment = lvwColumnLeft
        'fsftDateLastUpdated
        .ColumnHeaders.Item(GuiServiceFeesListView.fsftDateLastUpdated).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiServiceFeesListView.fsftDateLastUpdated).Alignment = lvwColumnLeft
        'fsftUpdateByUserID
        .ColumnHeaders.Item(GuiServiceFeesListView.fsftUpdateByUserID).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiServiceFeesListView.fsftUpdateByUserID).Alignment = lvwColumnLeft
    End With
    
    bGridOn = CBool(GetSetting(App.EXEName, "GENERAL", "GRID_ON", False))
    lvwServiceFees.GridLines = bGridOn
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub LoadHeaderlvwServiceFees"
End Sub

Public Sub LoadHeaderlvwExpenseFees()
    On Error GoTo EH
    Dim bGridOn As Boolean
    Dim bHideDeleted As Boolean
    Dim bHideUploadFlags As Boolean
    
    bHideDeleted = CBool(GetSetting(App.EXEName, "GENERAL", "HIDE_DELETED", True))
    bHideUploadFlags = CBool(GetSetting(App.EXEName, "GENERAL", "HIDE_UPLOAD_FLAGS", True))

    'set the columnheaders
    With lvwExpenseFees
        .ColumnHeaders.Add , "fsftDescription", "Description"
        .ColumnHeaders.Add , "Comment", "Comment" 'hidden
        .ColumnHeaders.Add , "NumberOfItems", "Items"
        .ColumnHeaders.Add , "Amount", "Amount"
        .ColumnHeaders.Add , "UpLoadMe", "UpLoadMe" 'hidden
        .ColumnHeaders.Add , "DateLastUpdated", "Date Last Updated" 'hidden
        .ColumnHeaders.Add , "AdminComments", "Admin Comments" 'hidden
        .ColumnHeaders.Add , "DownLoadMe", "DownLoadMe" 'hidden"
        .ColumnHeaders.Add , "RTIBFeeID", "RTIBFeeID" 'hidden"
        .ColumnHeaders.Add , "AssignmentsID", "AssignmentsID" 'hidden"
        .ColumnHeaders.Add , "ID", "ID" 'hidden"
        .ColumnHeaders.Add , "IDAssignments", "IDAssignments" 'hidden"
        .ColumnHeaders.Add , "UpdateByUserID", "UpdateByUserID" 'hidden"
        .ColumnHeaders.Add , "FeeScheduleFeeTypesID", "FeeScheduleFeeTypesID" 'hidden"
        .ColumnHeaders.Add , "fsftFeeScheduleID", "fsftFeeScheduleID" 'hidden"
        .ColumnHeaders.Add , "fsftTypeNum", "fsftTypeNum" 'hidden"
        .ColumnHeaders.Add , "fsftName", "fsftName" 'hidden"
        .ColumnHeaders.Add , "fsftFeeAmount", "fsftFeeAmount" 'hidden"
        .ColumnHeaders.Add , "fsftIsExpense", "fsftIsExpense" 'hidden"
        .ColumnHeaders.Add , "fsftMaxNumberOfItems", "fsftMaxNumberOfItems" 'hidden"
        .ColumnHeaders.Add , "fsftMaxFeeAmount", "fsftMaxFeeAmount" 'hidden"
        .ColumnHeaders.Add , "fsftIsMiscAmount", "fsftIsMiscAmount" 'hidden"
        .ColumnHeaders.Add , "fsftUseFormula", "fsftUseFormula" 'hidden"
        .ColumnHeaders.Add , "fsftVBFormula", "fsftVBFormula" 'hidden"
        .ColumnHeaders.Add , "fsftIsDeleted", "fsftIsDeleted" 'hidden"
        .ColumnHeaders.Add , "fsftDateLastUpdated", "fsftDateLastUpdated" 'hidden"
        .ColumnHeaders.Add , "fsftUpdateByUserID", "fsftUpdateByUserID" 'hidden"
    
        .Sorted = False
        .SortOrder = lvwAscending
'        GuiExpenseFeesListView.fsftDescription = 4334.74
'GuiExpenseFeesListView.Comment = 3420.284

        'fsftDescription
        .ColumnHeaders.Item(GuiExpenseFeesListView.fsftDescription).Width = 5000 'Me.Width - 9320 '6030
        .ColumnHeaders.Item(GuiExpenseFeesListView.fsftDescription).Alignment = lvwColumnLeft
        'Comment
        .ColumnHeaders.Item(GuiExpenseFeesListView.Comment).Width = 0 'Hidden Me.Width - 10320 '5030
        .ColumnHeaders.Item(GuiExpenseFeesListView.Comment).Alignment = lvwColumnLeft
        'NumberOfItems
        .ColumnHeaders.Item(GuiExpenseFeesListView.NumberOfItems).Width = 1750
        .ColumnHeaders.Item(GuiExpenseFeesListView.NumberOfItems).Alignment = lvwColumnRight
'        .ColumnHeaders.Item(GuiExpenseFeesListView.NumberOfItems).Icon = GuiIBStatusList.showDropDown
        'Amount
        .ColumnHeaders.Item(GuiExpenseFeesListView.Amount).Width = 2000
        .ColumnHeaders.Item(GuiExpenseFeesListView.Amount).Alignment = lvwColumnRight
        'UpLoadMe
        If bHideUploadFlags Then
            .ColumnHeaders.Item(GuiExpenseFeesListView.UpLoadMe).Width = 0 'Hidden 400
        Else
            .ColumnHeaders.Item(GuiExpenseFeesListView.UpLoadMe).Width = 400
        End If
        .ColumnHeaders.Item(GuiExpenseFeesListView.UpLoadMe).Alignment = lvwColumnLeft
        .ColumnHeaders.Item(GuiExpenseFeesListView.UpLoadMe).Icon = GuiIBStatusList.UpLoadMe
        'DateLastUpdated
        .ColumnHeaders.Item(GuiExpenseFeesListView.DateLastUpdated).Width = 0 'Hidden 2200
        .ColumnHeaders.Item(GuiExpenseFeesListView.DateLastUpdated).Alignment = lvwColumnLeft
        'AdminComments
        .ColumnHeaders.Item(GuiExpenseFeesListView.AdminComments).Width = 0 'Hidden 5000
        .ColumnHeaders.Item(GuiExpenseFeesListView.AdminComments).Alignment = lvwColumnLeft
        'DownLoadMe
        .ColumnHeaders.Item(GuiExpenseFeesListView.DownLoadMe).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiExpenseFeesListView.DownLoadMe).Alignment = lvwColumnLeft
        'RTIBFeeID
        .ColumnHeaders.Item(GuiExpenseFeesListView.RTIBFeeID).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiExpenseFeesListView.RTIBFeeID).Alignment = lvwColumnLeft
        'AssignmentsID
        .ColumnHeaders.Item(GuiExpenseFeesListView.AssignmentsID).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiExpenseFeesListView.AssignmentsID).Alignment = lvwColumnLeft
        'ID
        .ColumnHeaders.Item(GuiExpenseFeesListView.ID).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiExpenseFeesListView.ID).Alignment = lvwColumnLeft
        'IDAssignments
        .ColumnHeaders.Item(GuiExpenseFeesListView.IDAssignments).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiExpenseFeesListView.IDAssignments).Alignment = lvwColumnLeft
        'UpdateByUserID
        .ColumnHeaders.Item(GuiExpenseFeesListView.UpdateByUserID).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiExpenseFeesListView.UpdateByUserID).Alignment = lvwColumnLeft
        'FeeScheduleFeeTypesID
        .ColumnHeaders.Item(GuiExpenseFeesListView.FeeScheduleFeeTypesID).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiExpenseFeesListView.FeeScheduleFeeTypesID).Alignment = lvwColumnLeft
        'fsftFeeScheduleID
        .ColumnHeaders.Item(GuiExpenseFeesListView.fsftFeeScheduleID).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiExpenseFeesListView.fsftFeeScheduleID).Alignment = lvwColumnLeft
        'fsftTypeNum
        .ColumnHeaders.Item(GuiExpenseFeesListView.fsftTypeNum).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiExpenseFeesListView.fsftTypeNum).Alignment = lvwColumnLeft
        'fsftName
        .ColumnHeaders.Item(GuiExpenseFeesListView.fsftName).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiExpenseFeesListView.fsftName).Alignment = lvwColumnLeft
        'fsftFeeAmount
        .ColumnHeaders.Item(GuiExpenseFeesListView.fsftFeeAmount).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiExpenseFeesListView.fsftFeeAmount).Alignment = lvwColumnLeft
        'fsftIsExpense
        .ColumnHeaders.Item(GuiExpenseFeesListView.fsftIsExpense).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiExpenseFeesListView.fsftIsExpense).Alignment = lvwColumnLeft
        'fsftMaxNumberOfItems
        .ColumnHeaders.Item(GuiExpenseFeesListView.fsftMaxNumberOfItems).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiExpenseFeesListView.fsftMaxNumberOfItems).Alignment = lvwColumnLeft
        'fsftMaxFeeAmount
        .ColumnHeaders.Item(GuiExpenseFeesListView.fsftMaxFeeAmount).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiExpenseFeesListView.fsftMaxFeeAmount).Alignment = lvwColumnLeft
        'fsftIsMiscAmount
        .ColumnHeaders.Item(GuiExpenseFeesListView.fsftIsMiscAmount).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiExpenseFeesListView.fsftIsMiscAmount).Alignment = lvwColumnLeft
        'fsftUseFormula
        .ColumnHeaders.Item(GuiExpenseFeesListView.fsftUseFormula).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiExpenseFeesListView.fsftUseFormula).Alignment = lvwColumnLeft
        'fsftVBFormula
        .ColumnHeaders.Item(GuiExpenseFeesListView.fsftVBFormula).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiExpenseFeesListView.fsftVBFormula).Alignment = lvwColumnLeft
        'fsftIsDeleted
        .ColumnHeaders.Item(GuiExpenseFeesListView.fsftIsDeleted).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiExpenseFeesListView.fsftIsDeleted).Alignment = lvwColumnLeft
        'fsftDateLastUpdated
        .ColumnHeaders.Item(GuiExpenseFeesListView.fsftDateLastUpdated).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiExpenseFeesListView.fsftDateLastUpdated).Alignment = lvwColumnLeft
        'fsftUpdateByUserID
        .ColumnHeaders.Item(GuiExpenseFeesListView.fsftUpdateByUserID).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiExpenseFeesListView.fsftUpdateByUserID).Alignment = lvwColumnLeft
    End With
    
    bGridOn = CBool(GetSetting(App.EXEName, "GENERAL", "GRID_ON", False))
    lvwExpenseFees.GridLines = bGridOn
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub LoadHeaderlvwExpenseFees"
End Sub


Private Sub framExpenses_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo EH
    
'    LooseOverridesFocus
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub framExpenses_MouseMove"
End Sub


Private Sub framOverrideFees_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 On Error GoTo EH
    
'    'Need to expand on focus lack of Realestate in 800X 600 restraint
'    framOverrideFees.Height = ORFEES_FRAM_HEIGHT_ONFOCUS
'    lvwOverrideFees.Height = ORFEES_LVW_HEIGHT_ONFOCUS
'    txtServiceFeeComment.Visible = False
    
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub framOverrideFees_MouseMove"
End Sub

Private Sub framServiceFee_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo EH
'    LooseOverridesFocus
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub framServiceFee_MouseMove"
End Sub

Private Sub framServiceFees_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo EH
    
'    LooseOverridesFocus
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub framServiceFees_MouseMove"
End Sub


Private Sub framTotals_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    LooseOverridesFocus
End Sub

Private Sub IBOptions_Click(Index As Integer)
    On Error GoTo EH
    
    Select Case Index
        Case OPT_VIEW_IB_DETAILS
            If IBOptions(OPT_VIEW_IB_DETAILS).Value = vbChecked Then
                framIBDetails.Enabled = True
                framIBDetails.Visible = True
                framOverrideFees.Enabled = False
                framOverrideFees.Visible = False
                IBOptions(OPT_VIEW_IB_OVERRIDING).Value = vbUnchecked
            Else
                framIBDetails.Enabled = False
                framIBDetails.Visible = False
            End If
        Case OPT_VIEW_IB_OVERRIDING
            If IBOptions(OPT_VIEW_IB_OVERRIDING).Value = vbChecked Then
                framOverrideFees.Enabled = True
                framOverrideFees.Visible = True
                framIBDetails.Enabled = False
                framIBDetails.Visible = False
                IBOptions(OPT_VIEW_IB_DETAILS).Value = vbUnchecked
            Else
                framOverrideFees.Enabled = False
                framOverrideFees.Visible = False
            End If
    End Select
    
    ShowFrame
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub IBOptions_Click"
End Sub

Private Sub imgAdminComments_Click()
    On Error GoTo EH
    Dim oLR As V2ECKeyBoard.clsLossReports
    Dim sLRFormat As String
    Dim sLRData As String
    Dim sCaption As String
    Dim sPDFFilePath As String
    Dim sMess As String
    Dim lCurPDFImage As Long

    'check to see if this claim is currenlty unloading
    'if it is don' allow this event to occur
    If mfrmClaim.MyClaimsList.UnloadingClaim Then
        Exit Sub
    End If
    
    lCurPDFImage = imgBillingStatus.ListImages(GuiIBStatusList.showPDF).Picture
    
    If imgAdminComments.Picture <> lCurPDFImage Then
        sMess = "No Admin Comments available."
        MsgBox sMess, vbInformation + vbOKOnly, "Admin Comments"
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    sCaption = "Admin Comments "

    sLRFormat = "TEXT"
    sLRData = msAdminComments
    
   
    sPDFFilePath = goUtil.gsInstallDir & "\TempAdminComments" & goUtil.utGetTickCount & ".pdf"
    Set oLR = New V2ECKeyBoard.clsLossReports
    If StrComp(sLRFormat, "TEXT", vbTextCompare) <> 0 Then
        sLRData = sLRFormat & vbCrLf & sLRData
    End If
    oLR.CreateExport sLRData, sPDFFilePath, ARPdf
    If goUtil.utFileExists(sPDFFilePath) Then
        goUtil.utShellExecute , "OPEN", sPDFFilePath, , , vbNormalFocus, True, True, True, sCaption
        DoEvents
        Sleep 1000
        goUtil.utDeleteFile sPDFFilePath
    End If
    
    Set oLR = Nothing
    
    Screen.MousePointer = MousePointerConstants.vbDefault
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub imgAdminComments_Click"
End Sub

Private Sub imgUploadMe_Click()
    On Error GoTo EH
    Dim sMess As String
    Dim lCurUploadImage As Long
    
    lCurUploadImage = imgBillingStatus.ListImages(GuiIBStatusList.UpLoadMe).Picture
    
    If imgUploadMe.Picture = lCurUploadImage Then
        sMess = "Recent changes require IB data to be uploaded the next time you connect."
    Else
        sMess = "No IB data is currently scheduled to be uploaded."
    End If

    MsgBox sMess, vbInformation + vbOKOnly, "Upload Me"
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub imgUploadMe_Click"
End Sub


Private Sub imgVoid_Click()
    On Error GoTo EH
    Dim sMess As String
    Dim lCurVoidImage As Long
    
    lCurVoidImage = imgBillingStatus.ListImages(GuiIBStatusList.IsDeleted).Picture
    
    If imgVoid.Picture = lCurVoidImage Then
        sMess = "Are you sure you want to UNVOID this IB?"
    Else
        sMess = "Are you sure you want to VOID this IB?"
    End If

    If MsgBox(sMess, vbYesNo + vbQuestion, "VOID IB") = vbYes Then
        If imgVoid.Picture = lCurVoidImage Then
            imgVoid.Picture = LoadPicture()
        Else
            imgVoid.Picture = imgBillingStatus.ListImages(GuiIBStatusList.IsDeleted).Picture
        End If
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub imgVoid_Click"
End Sub

Private Sub lstAddExpenseFeeItemNumberOfItems_Change()
    On Error GoTo EH
    Dim lMaxItems As Long
    Dim lNumberOfItems As Long

    If Not IsNumeric(lstAddExpenseFeeItemNumberOfItems.Text) Then
        lstAddExpenseFeeItemNumberOfItems.Text = 0
        goUtil.utSelText lstAddExpenseFeeItemNumberOfItems
    Else
        lMaxItems = CLng(mSelectedAddExpenseFeeItemX.ListSubItems(GuiExpenseFeesListView.fsftMaxNumberOfItems - 1))
        lNumberOfItems = lstAddExpenseFeeItemNumberOfItems.Text
        If lNumberOfItems > lMaxItems Then
            lstAddExpenseFeeItemNumberOfItems.Text = lMaxItems
            goUtil.utSelText lstAddExpenseFeeItemNumberOfItems
        ElseIf lNumberOfItems < 0 Then
            lstAddExpenseFeeItemNumberOfItems.Text = 0
            goUtil.utSelText lstAddExpenseFeeItemNumberOfItems
        End If
    End If
    
    If lstAddExpenseFeeItemNumberOfItems.Text = 0 Then
        txtAddExpenseFeeItemAmount.Text = "0.00"
        lstAddExpenseFeeItemNumberOfItems.Text = 0
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lstAddExpenseFeeItemNumberOfItems_Change"
End Sub

Private Sub lstAddExpenseFeeItemNumberOfItems_LostFocus()
    CalcAddExpenseFeeItem
End Sub

Private Sub lstAddServiceFeeItemNumberOfItems_Change()
    On Error GoTo EH
    Dim lMaxItems As Long
    Dim lNumberOfItems As Long

    If Not IsNumeric(lstAddServiceFeeItemNumberOfItems.Text) Then
        lstAddServiceFeeItemNumberOfItems.Text = 0
        goUtil.utSelText lstAddServiceFeeItemNumberOfItems
    Else
        lMaxItems = CLng(mSelectedAddServiceFeeItemX.ListSubItems(GuiServiceFeesListView.fsftMaxNumberOfItems - 1))
        lNumberOfItems = lstAddServiceFeeItemNumberOfItems.Text
        If lNumberOfItems > lMaxItems Then
            lstAddServiceFeeItemNumberOfItems.Text = lMaxItems
            goUtil.utSelText lstAddServiceFeeItemNumberOfItems
        ElseIf lNumberOfItems < 0 Then
            lstAddServiceFeeItemNumberOfItems.Text = 0
            goUtil.utSelText lstAddServiceFeeItemNumberOfItems
        End If
    End If
    
    If lstAddServiceFeeItemNumberOfItems.Text = 0 Then
        txtAddServiceFeeItemAmount = "0.00"
        lstAddServiceFeeItemNumberOfItems.Text = 0
    End If
    
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lstAddServiceFeeItemNumberOfItems_Change"
End Sub

Public Sub RefreshFeeSchedule()
    On Error GoTo EH
        
    'Load Fee Schedule RS
    mfrmClaim.SetadoRSFeeScheduleList
    cboFeeSchedule.Clear
    cboFeeSchedule.AddItem "(--Select Fee Schedule--)"
    '0 indicates Null ID since ID must be >= 1 or <= -1
    '>=1    : WEB Server Synched
    '<=-1   : Client has yet to synch Data ID with Web Server.
    cboFeeSchedule.ItemData(cboFeeSchedule.NewIndex) = 0
    mfrmClaim.PopulateLookUp mfrmClaim.adoRSFeeScheduleList, _
                        Nothing, _
                        cboFeeSchedule, _
                        "FeeScheduleID", _
                        vbNullString, _
                        "ScheduleName", _
                        "Description"
                        
    cboFeeSchedule.ListIndex = 0
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub RefreshFeeSchedule"
End Sub

Public Sub RefreshBilling()
    On Error GoTo EH
    Dim sData As String
    Dim RS As ADODB.Recordset
    Dim lCount As Long
    Dim sIBID As String
    Dim sSupplement As String
    Dim sReportTitle As String
    
    'Load Billing RS
    mfrmClaim.SetadoRSBillingCount msAssignmentsID, , , True
    cboBillingID.Clear
    cboBillingID.AddItem "(--Select Billing--)"
    '0 indicates Null ID since ID must be >= 1 or <= -1
    '>=1    : WEB Server Synched
    '<=-1   : Client has yet to synch Data ID with Web Server.
    cboBillingID.ItemData(cboBillingID.NewIndex) = 0
    mfrmClaim.PopulateLookUp mfrmClaim.adoRSBillingCount, _
                        Nothing, _
                        cboBillingID, _
                        "ID", _
                        vbNullString, _
                        "IB", _
                        "IBDescription", , , True, "IBDescription2"
    'Need to Add some Data items to Text so that
    'the IB can be printed... Wether it is Closed or Current
    'this contains the Software info needed to Print IB
    mfrmClaim.SetadoRSBillingReports ' software for Bills
    mfrmClaim.SetadoRSBillingReportsHistory 'software history for Bills
    Set RS = mfrmClaim.adoRSBillingReports
    
    If RS.RecordCount = 1 Then
        RS.MoveFirst
        For lCount = 1 To cboBillingID.ListCount - 1
            sData = cboBillingID.List(lCount)
            sReportTitle = Trim(left(sData, InStr(1, sData, "-", vbTextCompare) - 1)) & ")"
            If InStr(1, sData, "Current", vbTextCompare) > 0 Then
                'Set the IBID to "" if this is Current Billing
                'that way the IB Report will use the RTTable instead of the
                'Closed IB Table
                sIBID = vbNullString
                sSupplement = mfrmClaim.GetSupplement(cboBillingID.ItemData(lCount))
            Else
                sIBID = mfrmClaim.GetIBID(cboBillingID.ItemData(lCount))
                sSupplement = mfrmClaim.GetSupplement(cboBillingID.ItemData(lCount))
            End If
            sData = sData & String(200, " ")
            sData = sData & RS.Fields("ProjectName").Value & "|" & RS.Fields("ClassName").Value & "|" & RS.Fields("Version").Value
            sData = sData & "|" & "pIBID=" & sIBID
            sData = sData & "|" & "pSupplement=" & sSupplement
            sData = sData & "|" & "psReportTitle=" & sReportTitle
            'Add the Software Data to this IB item
            cboBillingID.List(lCount) = sData
        Next
    Else
        Set RS = mfrmClaim.adoRSBillingReportsHistory
        If RS.RecordCount > 0 Then
            RS.MoveFirst
            For lCount = 1 To cboBillingID.ListCount - 1
                sData = cboBillingID.List(lCount)
                sReportTitle = Trim(left(sData, InStr(1, sData, "-", vbTextCompare) - 1)) & ")"
                If InStr(1, sData, "Current", vbTextCompare) > 0 Then
                    'Set the IBID to "" if this is Current Billing
                    'that way the IB Report will use the RTTable instead of the
                    'Closed IB Table
                    sIBID = vbNullString
                    sSupplement = mfrmClaim.GetSupplement(cboBillingID.ItemData(lCount))
                Else
                    sIBID = mfrmClaim.GetIBID(cboBillingID.ItemData(lCount))
                    sSupplement = mfrmClaim.GetSupplement(cboBillingID.ItemData(lCount))
                End If
                sData = sData & String(200, " ")
                sData = sData & RS.Fields("ProjectName").Value & "|" & RS.Fields("ClassName").Value & "|" & RS.Fields("Version").Value
                sData = sData & "|" & "pIBID=" & sIBID
                sData = sData & "|" & "pSupplement=" & sSupplement
                sData = sData & "|" & "psReportTitle=" & sReportTitle
                'Add the Software Data to this IB item
                cboBillingID.List(lCount) = sData
            Next
        End If
    End If
    
    'select the first Element
    cboBillingID.ListIndex = 0
    
    'cleanup
    
    Set RS = Nothing
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub RefreshBilling"
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
                        If StrComp(oControl.Name, chkUseActivityTime.Name, vbTextCompare) = 0 Then
                            oControl.Enabled = False 'This item must always be disabled
                        ElseIf StrComp(oControl.Name, chkFeeByTime.Name, vbTextCompare) = 0 Then
                            If Not mbOverridesServiceFee And Not mbOverridesFeeByTimeFee And Not mbOverridesALL Then
                                oControl.Enabled = MyFrame.Enabled
                            Else
                                oControl.Enabled = False
                            End If
                        ElseIf StrComp(oControl.Name, ChkOverrideAmounts.Name, vbTextCompare) = 0 Then
                            If Not mbOverridesServiceFee And Not mbOverridesFeeByTimeFee And Not mbOverridesALL Then
                                oControl.Enabled = MyFrame.Enabled
                            Else
                                oControl.Enabled = False
                            End If
                        Else
                            oControl.Enabled = MyFrame.Enabled
                        End If
                        
                        If InStr(1, oControl.Tag, "NO_SHOW_FRAME", vbTextCompare) = 0 Then
                            oControl.Visible = MyFrame.Enabled
                            If StrComp(oControl.Name, IBOptions(0).Name, vbTextCompare) = 0 Then
                                If Not oControl.Visible Then
                                    framIBDetails.Enabled = False
                                    framOverrideFees.Enabled = False
                                    IBOptions(0).Value = False
                                    IBOptions(1).Value = False
                                End If
                            End If
                        End If
                        
                        If StrComp(oControl.Name, txtComments.Name, vbTextCompare) = 0 Then
                            If framServiceFee.Enabled Then
                                txtComments.Visible = False  'NOT USING COMMENTS AT THIS TIME (Could be used for Billing Admin purposes)
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

Public Function PopulateClosedBillingInfo() As Boolean
    On Error GoTo EH
    Dim RSCLIB As ADODB.Recordset
    Dim RSCLIBFee As ADODB.Recordset
    
    If Not mfrmClaim.SetadoRSCLIB(msAssignmentsID, msBillingCountID) Then
        GoTo CLEAN_UP
    End If
    If Not mfrmClaim.SetadoRSCLIBFee(msAssignmentsID, msBillingCountID, msFeeScheduleID) Then
        GoTo CLEAN_UP
    End If
    
    Set RSCLIB = mfrmClaim.adoRSCLIB
    Set RSCLIBFee = mfrmClaim.adoRSCLIBFee
    
    If RSCLIBFee.RecordCount > 0 Then
        RSCLIBFee.MoveFirst
    End If
    PopulatelvwOverrideFees lvwOverrideFees, RSCLIBFee
    
    If RSCLIBFee.RecordCount > 0 Then
        RSCLIBFee.MoveFirst
    End If
    PopulatelvwServiceFees lvwServiceFees, RSCLIBFee
    
    If RSCLIBFee.RecordCount > 0 Then
        RSCLIBFee.MoveFirst
    End If
    PopulatelvwExpenseFees lvwExpenseFees, RSCLIBFee
    
    PopulateIBInfo RSCLIB
    
CLEAN_UP:
    Set RSCLIB = Nothing
    Set RSCLIBFee = Nothing
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function PopulateClosedBillingInfo"
End Function

Public Function PopulateOpenBillingInfo() As Boolean
    On Error GoTo EH
    Dim RSRTIB As ADODB.Recordset
    Dim RSRTIBFee As ADODB.Recordset
    
    If Not mfrmClaim.SetadoRSRTIB(msAssignmentsID) Then
        GoTo CLEAN_UP
    End If
    If Not mfrmClaim.SetadoRSRTIBFee(msAssignmentsID, msFeeScheduleID) Then
        GoTo CLEAN_UP
    End If
    
    Set RSRTIB = mfrmClaim.adoRSRTIB
    Set RSRTIBFee = mfrmClaim.adoRSRTIBFee
    
    If RSRTIBFee.RecordCount > 0 Then
        RSRTIBFee.MoveFirst
    End If
    PopulatelvwOverrideFees lvwOverrideFees, RSRTIBFee
    
    If RSRTIBFee.RecordCount > 0 Then
        RSRTIBFee.MoveFirst
    End If
    PopulatelvwServiceFees lvwServiceFees, RSRTIBFee
    
    If RSRTIBFee.RecordCount > 0 Then
        RSRTIBFee.MoveFirst
    End If
    PopulatelvwExpenseFees lvwExpenseFees, RSRTIBFee
    
    PopulateIBInfo RSRTIB
    
CLEAN_UP:
    Set RSRTIB = Nothing
    Set RSRTIBFee = Nothing
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function PopulateOpenBillingInfo"
End Function

Public Sub PopulatelvwOverrideFees(poLvw As ListView, pRSFee As ADODB.Recordset)
    On Error GoTo EH
    Dim iMyIcon As Long
    Dim sFlagText As String
    Dim sTemp As String
    Dim itmX As ListItem
    Dim oListView As MSComctlLib.ListView
    Dim RS As ADODB.Recordset
    Dim sVBFormula As String
    Dim bHideDeleted As Boolean
    Dim bHideUploadFlags As Boolean
    
    bHideDeleted = CBool(GetSetting(App.EXEName, "GENERAL", "HIDE_DELETED", True))
    bHideUploadFlags = CBool(GetSetting(App.EXEName, "GENERAL", "HIDE_UPLOAD_FLAGS", True))
    
    'Clear the List view
    Set oListView = poLvw

    oListView.Visible = False
    oListView.ListItems.Clear

    Set RS = pRSFee

    If Not RS.EOF Then
        RS.MoveFirst
        Do Until RS.EOF
            'ONLY populate VBFORMULA THAT CONTAINS "Overrides_"
            sVBFormula = goUtil.IsNullIsVbNullString(RS.Fields("fsftVBFormula"))
            If InStr(1, sVBFormula, "Overrides_", vbTextCompare) = 0 Then
                GoTo SKIP_ITEM
            ElseIf StrComp(sVBFormula, VBFORMULA_OVERRIDES_SERVICE_FEE, vbTextCompare) = 0 Then
                If CBool(goUtil.IsNullIsVbNullString(RS.Fields("NumberOfItems"))) Then
                    'True only on loading an IB that is set to manual override Service fee (That means Fee Schedule Calc is Disabled)
                    mbOverridesServiceFee = True
                    mcOverridesServiceFeeAmount = goUtil.IsNullIsVbNullString(RS.Fields("Amount"))
                    cmdCalcFeeSched.Enabled = False
                Else
                    mbOverridesServiceFee = False
                    mcOverridesServiceFeeAmount = 0#
                End If
            ElseIf StrComp(sVBFormula, VBFORMULA_OVERRIDES_FEEBYTIME_FEE, vbTextCompare) = 0 Then
                If CBool(goUtil.IsNullIsVbNullString(RS.Fields("NumberOfItems"))) Then
                    'True only on loading an IB that is set to manual override Fee By Time (That means that activity log time is not used)
                    mbOverridesFeeByTimeFee = True
                    mdblOverridesActLogTime = Format(goUtil.IsNullIsVbNullString(RS.Fields("Amount")), "##0.00")
                    mcOverridesFeeByTimeFeeAmount = CCur(Format(mdblOverridesActLogTime * mcFeeServiceHourlyRate, "#,###,###,##0.00"))
                    cmdCalcFeeSched.Enabled = False
                Else
                    mbOverridesFeeByTimeFee = False
                    mcOverridesFeeByTimeFeeAmount = 0#
                    mdblOverridesActLogTime = 0#
                End If
            ElseIf StrComp(sVBFormula, VBFORMULA_OVERRIDES_ALL, vbTextCompare) = 0 Then
                If CBool(goUtil.IsNullIsVbNullString(RS.Fields("NumberOfItems"))) Then
                    mbOverridesALL = True
                    framServiceFees.Enabled = False
                    framExpenses.Enabled = False
                    ShowFrame
                    cmdCalcFeeSched.Enabled = False
                End If
            End If
            'fsftDescription = 1
            Set itmX = oListView.ListItems.Add(, """" & goUtil.IsNullIsVbNullString(RS.Fields("ID")) & """", goUtil.IsNullIsVbNullString(RS.Fields("fsftDescription")))
            'Comment
            itmX.SubItems(GuiOverrideFeesListView.Comment - 1) = goUtil.IsNullIsVbNullString(RS.Fields("Comment"))
            'NumberOfItems
            itmX.SubItems(GuiOverrideFeesListView.NumberOfItems - 1) = goUtil.IsNullIsVbNullString(RS.Fields("NumberOfItems"))
            'This item may only have a max of 1 so if it is 0 its unchecked
            'If its 1 then its checked
            'As well if any one item is checked all the other items must be unchecked
            'But that is not figured here.
            If CBool(itmX.SubItems(GuiOverrideFeesListView.NumberOfItems - 1)) Then
                iMyIcon = GuiIBStatusList.IsChecked
            Else
                iMyIcon = GuiIBStatusList.IsUnchecked
            End If
            itmX.ListSubItems(GuiOverrideFeesListView.NumberOfItems - 1).ReportIcon = iMyIcon
            
            'Amount
            itmX.SubItems(GuiOverrideFeesListView.Amount - 1) = Format(goUtil.IsNullIsVbNullString(RS.Fields("Amount")), "#,###,###,##0.00")
            'UpLoadMe
            If CBool(RS.Fields("UpLoadMe")) Then
                iMyIcon = GuiIBStatusList.UpLoadMe
            Else
                iMyIcon = Empty
            End If
            sFlagText = goUtil.GetFlagText(RS.Fields("UpLoadMe"))
            itmX.SubItems(GuiOverrideFeesListView.UpLoadMe - 1) = sFlagText
            If Not bHideUploadFlags Then
                itmX.ListSubItems(GuiOverrideFeesListView.UpLoadMe - 1).ReportIcon = iMyIcon
            End If
            'DateLastUpdated
            If Not IsNull(RS.Fields("DateLastUpdated").Value) Then
                If IsDate(RS.Fields("DateLastUpdated").Value) Then
                    itmX.SubItems(GuiOverrideFeesListView.DateLastUpdated - 1) = Format(RS.Fields("DateLastUpdated").Value, "MM/DD/YYYY HH:MM:SS")
                Else
                    itmX.SubItems(GuiOverrideFeesListView.DateLastUpdated - 1) = vbNullString
                End If
            Else
                itmX.SubItems(GuiOverrideFeesListView.DateLastUpdated - 1) = vbNullString
            End If
            'AdminComments
            itmX.SubItems(GuiOverrideFeesListView.AdminComments - 1) = goUtil.IsNullIsVbNullString(RS.Fields("AdminComments"))
            'DownLoadMe 'hidden
            itmX.SubItems(GuiOverrideFeesListView.DownLoadMe - 1) = goUtil.IsNullIsVbNullString(RS.Fields("DownLoadMe"))
            'RTIBFeeID 'hidden
            If Not mbClosedIB Then
                itmX.SubItems(GuiOverrideFeesListView.RTIBFeeID - 1) = goUtil.IsNullIsVbNullString(RS.Fields("RTIBFeeID"))
            Else
                itmX.SubItems(GuiOverrideFeesListView.RTIBFeeID - 1) = vbNullString
            End If
            'AssignmentsID 'hidden
            itmX.SubItems(GuiOverrideFeesListView.AssignmentsID - 1) = goUtil.IsNullIsVbNullString(RS.Fields("AssignmentsID"))
            'ID 'hidden
            itmX.SubItems(GuiOverrideFeesListView.ID - 1) = goUtil.IsNullIsVbNullString(RS.Fields("ID"))
            'IDAssignments 'hidden
            itmX.SubItems(GuiOverrideFeesListView.IDAssignments - 1) = goUtil.IsNullIsVbNullString(RS.Fields("IDAssignments"))
            'UpdateByUserID 'hidden
            itmX.SubItems(GuiOverrideFeesListView.UpdateByUserID - 1) = goUtil.IsNullIsVbNullString(RS.Fields("UpdateByUserID"))
            'FeeScheduleFeeTypesID 'hidden
            itmX.SubItems(GuiOverrideFeesListView.FeeScheduleFeeTypesID - 1) = goUtil.IsNullIsVbNullString(RS.Fields("FeeScheduleFeeTypesID"))
            'fsftFeeScheduleID 'hidden
            itmX.SubItems(GuiOverrideFeesListView.fsftFeeScheduleID - 1) = goUtil.IsNullIsVbNullString(RS.Fields("fsftFeeScheduleID"))
            'fsftTypeNum 'hidden
            itmX.SubItems(GuiOverrideFeesListView.fsftTypeNum - 1) = goUtil.IsNullIsVbNullString(RS.Fields("fsftTypeNum"))
            'fsftName 'hidden
            itmX.SubItems(GuiOverrideFeesListView.fsftName - 1) = goUtil.IsNullIsVbNullString(RS.Fields("fsftName"))
            'fsftFeeAmount 'hidden
            itmX.SubItems(GuiOverrideFeesListView.fsftFeeAmount - 1) = goUtil.IsNullIsVbNullString(RS.Fields("fsftFeeAmount"))
            'fsftIsExpense 'hidden
            itmX.SubItems(GuiOverrideFeesListView.fsftIsExpense - 1) = goUtil.IsNullIsVbNullString(RS.Fields("fsftIsExpense"))
            'fsftMaxNumberOfItems 'hidden
            itmX.SubItems(GuiOverrideFeesListView.fsftMaxNumberOfItems - 1) = goUtil.IsNullIsVbNullString(RS.Fields("fsftMaxNumberOfItems"))
            'fsftMaxFeeAmount 'hidden
            itmX.SubItems(GuiOverrideFeesListView.fsftMaxFeeAmount - 1) = goUtil.IsNullIsVbNullString(RS.Fields("fsftMaxFeeAmount"))
            'fsftIsMiscAmount 'hidden
            itmX.SubItems(GuiOverrideFeesListView.fsftIsMiscAmount - 1) = goUtil.IsNullIsVbNullString(RS.Fields("fsftIsMiscAmount"))
            'fsftUseFormula 'hidden
            itmX.SubItems(GuiOverrideFeesListView.fsftUseFormula - 1) = goUtil.IsNullIsVbNullString(RS.Fields("fsftUseFormula"))
            'fsftVBFormula 'hidden
            itmX.SubItems(GuiOverrideFeesListView.fsftVBFormula - 1) = goUtil.IsNullIsVbNullString(RS.Fields("fsftVBFormula"))
            'fsftIsDeleted 'hidden
            itmX.SubItems(GuiOverrideFeesListView.fsftIsDeleted - 1) = goUtil.IsNullIsVbNullString(RS.Fields("fsftIsDeleted"))
            'fsftDateLastUpdated 'hidden
            itmX.SubItems(GuiOverrideFeesListView.fsftDateLastUpdated - 1) = goUtil.IsNullIsVbNullString(RS.Fields("fsftDateLastUpdated"))
            'fsftUpdateByUserID 'hidden
            itmX.SubItems(GuiOverrideFeesListView.fsftUpdateByUserID - 1) = goUtil.IsNullIsVbNullString(RS.Fields("fsftUpdateByUserID"))
        
            itmX.Selected = False
SKIP_ITEM:
            RS.MoveNext
        Loop
    End If
    oListView.Visible = True
    'Cleanup
    Set itmX = Nothing
    Set RS = Nothing
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub PopulatelvwOverrideFees"
    oListView.Visible = True
End Sub


Public Sub PopulatelvwServiceFees(poLvw As ListView, pRSFee As ADODB.Recordset)
    On Error GoTo EH
    Dim iMyIcon As Long
    Dim sFlagText As String
    Dim sTemp As String
    Dim itmX As ListItem
    Dim oListView As MSComctlLib.ListView
    Dim RS As ADODB.Recordset
    Dim sVBFormula As String
    Dim sIsExpense As String
    Dim lMaxNumberOfItems As Long
    Dim bHideDeleted As Boolean
    Dim bHideUploadFlags As Boolean
    
    bHideDeleted = CBool(GetSetting(App.EXEName, "GENERAL", "HIDE_DELETED", True))
    bHideUploadFlags = CBool(GetSetting(App.EXEName, "GENERAL", "HIDE_UPLOAD_FLAGS", True))
    
    'Clear the List view
    Set oListView = poLvw

    oListView.Visible = False
    oListView.ListItems.Clear

    Set RS = pRSFee

    If Not RS.EOF Then
        RS.MoveFirst
        Do Until RS.EOF
            'ONLY populate Those items that are not Overiding fee and is not an expense"
            sVBFormula = goUtil.IsNullIsVbNullString(RS.Fields("fsftVBFormula"))
            sIsExpense = goUtil.IsNullIsVbNullString(RS.Fields("fsftIsExpense"))
            lMaxNumberOfItems = goUtil.IsNullIsVbNullString(RS.Fields("fsftMaxNumberOfItems"))
            If InStr(1, sVBFormula, "OVERRIDES_", vbTextCompare) > 0 Then
                GoTo SKIP_ITEM
            End If
            If CBool(sIsExpense) Then
                GoTo SKIP_ITEM
            End If
            'Don't show those items where the Max number of items has been Zeroed out
            If lMaxNumberOfItems = 0 Then
                GoTo SKIP_ITEM
            End If
            'fsftDescription = 1
            Set itmX = oListView.ListItems.Add(, """" & goUtil.IsNullIsVbNullString(RS.Fields("ID")) & """", goUtil.IsNullIsVbNullString(RS.Fields("fsftDescription")))
            'Comment
            itmX.SubItems(GuiServiceFeesListView.Comment - 1) = goUtil.IsNullIsVbNullString(RS.Fields("Comment"))
            'NumberOfItems
            itmX.SubItems(GuiServiceFeesListView.NumberOfItems - 1) = goUtil.IsNullIsVbNullString(RS.Fields("NumberOfItems"))
            'Amount
            itmX.SubItems(GuiServiceFeesListView.Amount - 1) = Format(goUtil.IsNullIsVbNullString(RS.Fields("Amount")), "#,###,###,##0.00")
            iMyIcon = GuiIBStatusList.showDropDown
'            itmX.ListSubItems(GuiServiceFeesListView.Amount - 1).ReportIcon = iMyIcon
            'UpLoadMe
            If CBool(RS.Fields("UpLoadMe")) Then
                iMyIcon = GuiIBStatusList.UpLoadMe
            Else
                iMyIcon = Empty
            End If
            sFlagText = goUtil.GetFlagText(RS.Fields("UpLoadMe"))
            itmX.SubItems(GuiServiceFeesListView.UpLoadMe - 1) = sFlagText
            If Not bHideUploadFlags Then
                itmX.ListSubItems(GuiServiceFeesListView.UpLoadMe - 1).ReportIcon = iMyIcon
            End If
            'DateLastUpdated
            If Not IsNull(RS.Fields("DateLastUpdated").Value) Then
                If IsDate(RS.Fields("DateLastUpdated").Value) Then
                    itmX.SubItems(GuiServiceFeesListView.DateLastUpdated - 1) = Format(RS.Fields("DateLastUpdated").Value, "MM/DD/YYYY HH:MM:SS")
                Else
                    itmX.SubItems(GuiServiceFeesListView.DateLastUpdated - 1) = vbNullString
                End If
            Else
                itmX.SubItems(GuiServiceFeesListView.DateLastUpdated - 1) = vbNullString
            End If
            'AdminComments
            itmX.SubItems(GuiServiceFeesListView.AdminComments - 1) = goUtil.IsNullIsVbNullString(RS.Fields("AdminComments"))
            'DownLoadMe 'hidden
            itmX.SubItems(GuiServiceFeesListView.DownLoadMe - 1) = goUtil.IsNullIsVbNullString(RS.Fields("DownLoadMe"))
            'RTIBFeeID 'hidden
            If Not mbClosedIB Then
                itmX.SubItems(GuiServiceFeesListView.RTIBFeeID - 1) = goUtil.IsNullIsVbNullString(RS.Fields("RTIBFeeID"))
            Else
                itmX.SubItems(GuiServiceFeesListView.RTIBFeeID - 1) = vbNullString
            End If
            'AssignmentsID 'hidden
            itmX.SubItems(GuiServiceFeesListView.AssignmentsID - 1) = goUtil.IsNullIsVbNullString(RS.Fields("AssignmentsID"))
            'ID 'hidden
            itmX.SubItems(GuiServiceFeesListView.ID - 1) = goUtil.IsNullIsVbNullString(RS.Fields("ID"))
            'IDAssignments 'hidden
            itmX.SubItems(GuiServiceFeesListView.IDAssignments - 1) = goUtil.IsNullIsVbNullString(RS.Fields("IDAssignments"))
            'UpdateByUserID 'hidden
            itmX.SubItems(GuiServiceFeesListView.UpdateByUserID - 1) = goUtil.IsNullIsVbNullString(RS.Fields("UpdateByUserID"))
            'FeeScheduleFeeTypesID 'hidden
            itmX.SubItems(GuiServiceFeesListView.FeeScheduleFeeTypesID - 1) = goUtil.IsNullIsVbNullString(RS.Fields("FeeScheduleFeeTypesID"))
            'fsftFeeScheduleID 'hidden
            itmX.SubItems(GuiServiceFeesListView.fsftFeeScheduleID - 1) = goUtil.IsNullIsVbNullString(RS.Fields("fsftFeeScheduleID"))
            'fsftTypeNum 'hidden
            itmX.SubItems(GuiServiceFeesListView.fsftTypeNum - 1) = goUtil.IsNullIsVbNullString(RS.Fields("fsftTypeNum"))
            'fsftName 'hidden
            itmX.SubItems(GuiServiceFeesListView.fsftName - 1) = goUtil.IsNullIsVbNullString(RS.Fields("fsftName"))
            'fsftFeeAmount 'hidden
            itmX.SubItems(GuiServiceFeesListView.fsftFeeAmount - 1) = goUtil.IsNullIsVbNullString(RS.Fields("fsftFeeAmount"))
            'fsftIsExpense 'hidden
            itmX.SubItems(GuiServiceFeesListView.fsftIsExpense - 1) = goUtil.IsNullIsVbNullString(RS.Fields("fsftIsExpense"))
            'fsftMaxNumberOfItems 'hidden
            itmX.SubItems(GuiServiceFeesListView.fsftMaxNumberOfItems - 1) = goUtil.IsNullIsVbNullString(RS.Fields("fsftMaxNumberOfItems"))
            'fsftMaxFeeAmount 'hidden
            itmX.SubItems(GuiServiceFeesListView.fsftMaxFeeAmount - 1) = goUtil.IsNullIsVbNullString(RS.Fields("fsftMaxFeeAmount"))
            'fsftIsMiscAmount 'hidden
            itmX.SubItems(GuiServiceFeesListView.fsftIsMiscAmount - 1) = goUtil.IsNullIsVbNullString(RS.Fields("fsftIsMiscAmount"))
            'fsftUseFormula 'hidden
            itmX.SubItems(GuiServiceFeesListView.fsftUseFormula - 1) = goUtil.IsNullIsVbNullString(RS.Fields("fsftUseFormula"))
            'fsftVBFormula 'hidden
            itmX.SubItems(GuiServiceFeesListView.fsftVBFormula - 1) = goUtil.IsNullIsVbNullString(RS.Fields("fsftVBFormula"))
            'fsftIsDeleted 'hidden
            itmX.SubItems(GuiServiceFeesListView.fsftIsDeleted - 1) = goUtil.IsNullIsVbNullString(RS.Fields("fsftIsDeleted"))
            'fsftDateLastUpdated 'hidden
            itmX.SubItems(GuiServiceFeesListView.fsftDateLastUpdated - 1) = goUtil.IsNullIsVbNullString(RS.Fields("fsftDateLastUpdated"))
            'fsftUpdateByUserID 'hidden
            itmX.SubItems(GuiServiceFeesListView.fsftUpdateByUserID - 1) = goUtil.IsNullIsVbNullString(RS.Fields("fsftUpdateByUserID"))
        
            itmX.Selected = False
SKIP_ITEM:
            RS.MoveNext
        Loop
    End If
    oListView.Visible = True
    'Cleanup
    Set itmX = Nothing
    Set RS = Nothing
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub PopulatelvwServiceFees"
    oListView.Visible = True
End Sub

Public Sub PopulatelvwExpenseFees(poLvw As ListView, pRSFee As ADODB.Recordset)
    On Error GoTo EH
    Dim iMyIcon As Long
    Dim sFlagText As String
    Dim sTemp As String
    Dim itmX As ListItem
    Dim oListView As MSComctlLib.ListView
    Dim RS As ADODB.Recordset
    Dim sVBFormula As String
    Dim sIsExpense As String
    Dim bHideDeleted As Boolean
    Dim bHideUploadFlags As Boolean
    
    bHideDeleted = CBool(GetSetting(App.EXEName, "GENERAL", "HIDE_DELETED", True))
    bHideUploadFlags = CBool(GetSetting(App.EXEName, "GENERAL", "HIDE_UPLOAD_FLAGS", True))
    
    'Clear the List view
    Set oListView = poLvw

    oListView.Visible = False
    oListView.ListItems.Clear

    Set RS = pRSFee

    If Not RS.EOF Then
        RS.MoveFirst
        Do Until RS.EOF
            'ONLY populate Those items that are not Overiding fee and is an expense"
            sVBFormula = goUtil.IsNullIsVbNullString(RS.Fields("fsftVBFormula"))
            sIsExpense = goUtil.IsNullIsVbNullString(RS.Fields("fsftIsExpense"))
            If InStr(1, sVBFormula, "OverrideS_", vbTextCompare) > 0 Then
                GoTo SKIP_ITEM
            End If
            If Not CBool(sIsExpense) Then
                GoTo SKIP_ITEM
            End If
            'fsftDescription = 1
            Set itmX = oListView.ListItems.Add(, """" & goUtil.IsNullIsVbNullString(RS.Fields("ID")) & """", goUtil.IsNullIsVbNullString(RS.Fields("fsftDescription")))
            'Comment
            itmX.SubItems(GuiServiceFeesListView.Comment - 1) = goUtil.IsNullIsVbNullString(RS.Fields("Comment"))
            'NumberOfItems
            itmX.SubItems(GuiServiceFeesListView.NumberOfItems - 1) = goUtil.IsNullIsVbNullString(RS.Fields("NumberOfItems"))
            'Amount
            itmX.SubItems(GuiServiceFeesListView.Amount - 1) = Format(goUtil.IsNullIsVbNullString(RS.Fields("Amount")), "#,###,###,##0.00")
            iMyIcon = GuiIBStatusList.showDropDown
'            itmX.ListSubItems(GuiServiceFeesListView.Amount - 1).ReportIcon = iMyIcon
            'UpLoadMe
            If CBool(RS.Fields("UpLoadMe")) Then
                iMyIcon = GuiIBStatusList.UpLoadMe
            Else
                iMyIcon = Empty
            End If
            sFlagText = goUtil.GetFlagText(RS.Fields("UpLoadMe"))
            itmX.SubItems(GuiServiceFeesListView.UpLoadMe - 1) = sFlagText
            If Not bHideUploadFlags Then
                itmX.ListSubItems(GuiServiceFeesListView.UpLoadMe - 1).ReportIcon = iMyIcon
            End If
            'DateLastUpdated
            If Not IsNull(RS.Fields("DateLastUpdated").Value) Then
                If IsDate(RS.Fields("DateLastUpdated").Value) Then
                    itmX.SubItems(GuiServiceFeesListView.DateLastUpdated - 1) = Format(RS.Fields("DateLastUpdated").Value, "MM/DD/YYYY HH:MM:SS")
                Else
                    itmX.SubItems(GuiServiceFeesListView.DateLastUpdated - 1) = vbNullString
                End If
            Else
                itmX.SubItems(GuiServiceFeesListView.DateLastUpdated - 1) = vbNullString
            End If
            'AdminComments
            itmX.SubItems(GuiServiceFeesListView.AdminComments - 1) = goUtil.IsNullIsVbNullString(RS.Fields("AdminComments"))
            'DownLoadMe 'hidden
            itmX.SubItems(GuiServiceFeesListView.DownLoadMe - 1) = goUtil.IsNullIsVbNullString(RS.Fields("DownLoadMe"))
            'RTIBFeeID 'hidden
            If Not mbClosedIB Then
                itmX.SubItems(GuiServiceFeesListView.RTIBFeeID - 1) = goUtil.IsNullIsVbNullString(RS.Fields("RTIBFeeID"))
            Else
                itmX.SubItems(GuiServiceFeesListView.RTIBFeeID - 1) = vbNullString
            End If
            'AssignmentsID 'hidden
            itmX.SubItems(GuiServiceFeesListView.AssignmentsID - 1) = goUtil.IsNullIsVbNullString(RS.Fields("AssignmentsID"))
            'ID 'hidden
            itmX.SubItems(GuiServiceFeesListView.ID - 1) = goUtil.IsNullIsVbNullString(RS.Fields("ID"))
            'IDAssignments 'hidden
            itmX.SubItems(GuiServiceFeesListView.IDAssignments - 1) = goUtil.IsNullIsVbNullString(RS.Fields("IDAssignments"))
            'UpdateByUserID 'hidden
            itmX.SubItems(GuiServiceFeesListView.UpdateByUserID - 1) = goUtil.IsNullIsVbNullString(RS.Fields("UpdateByUserID"))
            'FeeScheduleFeeTypesID 'hidden
            itmX.SubItems(GuiServiceFeesListView.FeeScheduleFeeTypesID - 1) = goUtil.IsNullIsVbNullString(RS.Fields("FeeScheduleFeeTypesID"))
            'fsftFeeScheduleID 'hidden
            itmX.SubItems(GuiServiceFeesListView.fsftFeeScheduleID - 1) = goUtil.IsNullIsVbNullString(RS.Fields("fsftFeeScheduleID"))
            'fsftTypeNum 'hidden
            itmX.SubItems(GuiServiceFeesListView.fsftTypeNum - 1) = goUtil.IsNullIsVbNullString(RS.Fields("fsftTypeNum"))
            'fsftName 'hidden
            itmX.SubItems(GuiServiceFeesListView.fsftName - 1) = goUtil.IsNullIsVbNullString(RS.Fields("fsftName"))
            'fsftFeeAmount 'hidden
            itmX.SubItems(GuiServiceFeesListView.fsftFeeAmount - 1) = goUtil.IsNullIsVbNullString(RS.Fields("fsftFeeAmount"))
            'fsftIsExpense 'hidden
            itmX.SubItems(GuiServiceFeesListView.fsftIsExpense - 1) = goUtil.IsNullIsVbNullString(RS.Fields("fsftIsExpense"))
            'fsftMaxNumberOfItems 'hidden
            itmX.SubItems(GuiServiceFeesListView.fsftMaxNumberOfItems - 1) = goUtil.IsNullIsVbNullString(RS.Fields("fsftMaxNumberOfItems"))
            'fsftMaxFeeAmount 'hidden
            itmX.SubItems(GuiServiceFeesListView.fsftMaxFeeAmount - 1) = goUtil.IsNullIsVbNullString(RS.Fields("fsftMaxFeeAmount"))
            'fsftIsMiscAmount 'hidden
            itmX.SubItems(GuiServiceFeesListView.fsftIsMiscAmount - 1) = goUtil.IsNullIsVbNullString(RS.Fields("fsftIsMiscAmount"))
            'fsftUseFormula 'hidden
            itmX.SubItems(GuiServiceFeesListView.fsftUseFormula - 1) = goUtil.IsNullIsVbNullString(RS.Fields("fsftUseFormula"))
            'fsftVBFormula 'hidden
            itmX.SubItems(GuiServiceFeesListView.fsftVBFormula - 1) = goUtil.IsNullIsVbNullString(RS.Fields("fsftVBFormula"))
            'fsftIsDeleted 'hidden
            itmX.SubItems(GuiServiceFeesListView.fsftIsDeleted - 1) = goUtil.IsNullIsVbNullString(RS.Fields("fsftIsDeleted"))
            'fsftDateLastUpdated 'hidden
            itmX.SubItems(GuiServiceFeesListView.fsftDateLastUpdated - 1) = goUtil.IsNullIsVbNullString(RS.Fields("fsftDateLastUpdated"))
            'fsftUpdateByUserID 'hidden
            itmX.SubItems(GuiServiceFeesListView.fsftUpdateByUserID - 1) = goUtil.IsNullIsVbNullString(RS.Fields("fsftUpdateByUserID"))
        
            itmX.Selected = False
SKIP_ITEM:
            RS.MoveNext
        Loop
    End If
    oListView.Visible = True
    'Cleanup
    Set itmX = Nothing
    Set RS = Nothing
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub PopulatelvwExpenseFees"
    oListView.Visible = True
End Sub


Public Sub PopulateIBInfo(pRSIB As ADODB.Recordset)
    On Error GoTo EH
    Dim sIB As String
    Dim sValue As String
    Dim sMess As String
    Dim sTemp As String
    Dim sCaption As String
    Dim sServiceFeeComment As String
    Dim itmX As ListItem
    Dim sfsftName As String
    Dim sfsftDescription As String
    Dim lNumberOfItems As Long
    'Check Some Idemnity Items
    Dim sOverrideAmounts As String
    Dim cGrossLoss As Currency
    Dim cDepreciation As Currency
    Dim cDeductible As Currency
    Dim cLessExcessLimits As Currency
    Dim cLessMiscellaneous As Currency
    Dim cNetClaim As Currency
    'check som Admin Items
    Dim sSubToCarrier As String
    Dim sIbnumber As String
    Dim sLocation As String
    Dim sState As String
    Dim sAdjusterName As String
    Dim sSALN As String
    Dim sPolicyNo As String
    Dim sInsuredName As String
    Dim sLossLocation As String
    Dim sDateOfLoss As String
    
    
    mbPopulateIBInfo = True
    
    If mbClosedIB Then
        sIB = "IB"
    Else
        sIB = "RT"
    End If
    
    sOverrideAmounts = goUtil.IsNullIsVbNullString(pRSIB.Fields(sIB & "33a_sAccountCode"))
    If CStr(vbChecked) = sOverrideAmounts Then
        ChkOverrideAmounts.Value = vbChecked
    Else
        ChkOverrideAmounts.Value = vbUnchecked
    End If
    
    
    If ChkOverrideAmounts.Value = vbChecked Then
        cGrossLoss = goUtil.IsNullIsVbNullString(pRSIB.Fields(sIB & "13_cGrossLoss"))
        cDepreciation = goUtil.IsNullIsVbNullString(pRSIB.Fields(sIB & "14_cDepreciation"))
        cLessExcessLimits = goUtil.IsNullIsVbNullString(pRSIB.Fields(sIB & "15a_cLessExcessLimits"))
        cLessMiscellaneous = goUtil.IsNullIsVbNullString(pRSIB.Fields(sIB & "15c_cLessMiscellaneous"))
        cNetClaim = goUtil.IsNullIsVbNullString(pRSIB.Fields(sIB & "16_cNetClaim"))
    Else
        cGrossLoss = mfrmClaim.GetFullCostOfRepair(CLng(msBillingCountID))
        cDepreciation = mfrmClaim.GetDepreciation(CLng(msBillingCountID))
        cLessExcessLimits = mfrmClaim.GetLessExcessLimits(CLng(msBillingCountID))
        cLessMiscellaneous = mfrmClaim.GetLessMiscellaneous(CLng(msBillingCountID))
        cNetClaim = mfrmClaim.GetNetActualCashValueClaim(CLng(msBillingCountID))
    End If
    cDeductible = goUtil.IsNullIsVbNullString(pRSIB.Fields(sIB & "15_cDeductible"))
    
    txtOverrideGrossLoss.Text = Format(cGrossLoss, "0.00")
    txtOverrideDepreciation.Text = Format(cDepreciation, "0.00")
    txtOverrideExcessLimit.Text = Format(cLessExcessLimits, "0.00")
    txtOverrideMiscellaneous.Text = Format(cLessMiscellaneous, "0.00")
    
    
    
    
    'lblFeeSchedule
    
    sCaption = mfrmClaim.GetCurrentFeeScheduleItem("ScheduleName", msFeeScheduleID)
    sCaption = sCaption & mfrmClaim.GetCurrentFeeScheduleItem("Description", msFeeScheduleID)
    lblFeeSchedule.Caption = sCaption
    
    'Service Fee
    'Coment
    sValue = goUtil.IsNullIsVbNullString(pRSIB.Fields(sIB & "18_sServiceFeeComment"))
    txtServiceFeeComment.Text = Trim(sValue)
    txtServiceFeeComment.ToolTipText = Trim(sValue)
    
    sValue = goUtil.IsNullIsVbNullString(pRSIB.Fields(sIB & "17_cServiceFee"))
    txtServiceFee.Text = Format(sValue, "#,###,###,##0.00")
    
    'Billing Selection
    sValue = goUtil.IsNullIsVbNullString(pRSIB.Fields("FeeByTime"))
    If CBool(sValue) Then
        chkFeeByTime.Value = vbChecked
    Else
        chkFeeByTime.Value = vbUnchecked
    End If
    sValue = goUtil.IsNullIsVbNullString(pRSIB.Fields("UseActivityTime"))
    If CBool(sValue) Then
        chkUseActivityTime.Value = vbChecked
    Else
        chkUseActivityTime.Value = vbUnchecked
    End If
    
    'Comments
    sValue = goUtil.IsNullIsVbNullString(pRSIB.Fields("Comments"))
    txtComments.Text = sValue
    
    If mbClosedIB Then
        '--------------------------------------BEGIN CLOSED IB---------------------------------
        '    mbOverridesFeeByTimeFee As Boolean 'True only on loading an IB that is set to manual override Fee BybTime (That means that activity log time is not used)
        If mbOverridesFeeByTimeFee Then
            'Check to See if the currenct Acttime Selection is different from
            'the current service fee amount.  If it Is need to inform user that
            'Rebill is necesarry to sycnh up the amounts
            If chkFeeByTime.Value = vbChecked Then
                If Format(mdblOverridesActLogTime * mcFeeServiceHourlyRate, "#,###,###,##0.00") <> txtServiceFee.Text Then
                    sMess = "The Current Overriding Fee By Time values do not match the Service Fee Amount!" & vbCrLf
                    sMess = sMess & "You must rebill this IB to correct."
                    MsgBox sMess, vbExclamation + vbOKOnly, "Service Fee not balanced"
                End If
            End If
            txtActLogHours.Text = mdblOverridesActLogTime
        Else
            'Check to See if the currenct Acttime Selection is different from
            'the current service fee amount.  If it Is need to inform user that
            'Rebill is necesarry to sycnh up the amounts
            If chkFeeByTime.Value = vbChecked Then
                If Format(mdblCurrentActLogTime * mcFeeServiceHourlyRate, "#,###,###,##0.00") <> txtServiceFee.Text Then
                    sMess = "The Current Activity Log Time for Fee By Time values do not match the Service Fee Amount!" & vbCrLf
                    sMess = sMess & "You must rebill this IB to correct."
                    MsgBox sMess, vbExclamation + vbOKOnly, "Service Fee not balanced"
                End If
            Else
                'If not overriding the Fee by Time
                'And not doing Fee By time
                'and not Overriding the Service Fee
                'and not overriding ALL fees
                'Check to see if the Current Service Fee Matches the Last calculated Schedule
                If Not mbOverridesServiceFee And Not mbOverridesALL Then
                    'Only do this if the txtServiceFee > 0
                    If CCur(txtServiceFee.Text) > 0 Then
                        If CCur(txtServiceFee.Text) <> Format(mcCurrentServiceFee, "#,###,###,##0.00") Then
                            sMess = "The Service Fee Amount (" & txtServiceFee.Text & ")  does not match the Calculated Feeschedule Amount (" & Format(mcCurrentServiceFee, "#,###,###,##0.00") & ")!" & vbCrLf
                            sMess = sMess & "You must rebill this IB to correct."
                            MsgBox sMess, vbExclamation + vbOKOnly, "Service Fee not balanced"
                        End If
                    End If
                End If
            End If
            txtActLogHours.Text = Format(mdblCurrentActLogTime, "##0.00")
        End If
        '--------------------------------------END CLOSED IB---------------------------------
    Else
        '--------------------------------------BEGIN OPEN IB---------------------------------
        If mbOverridesFeeByTimeFee Then
            'Check to See if the currenct Acttime Selection is different from
            'the current service fee amount.  If it Is need to inform user that
            'Rebill is necesarry to synch up the amounts
            If chkFeeByTime.Value = vbChecked Then
                If Format(mdblOverridesActLogTime * mcFeeServiceHourlyRate, "#,###,###,##0.00") <> txtServiceFee.Text Then
                    sMess = "The Current Overriding Fee By Time values do not match the Service Fee Amount!" & vbCrLf
                    If mbEditMode Then
                        sMess = sMess & "Do you want to update the Service Fee Amount?"
                    Else
                        sMess = sMess & "You must Edit this IB to update the Service Fee Amount."
                    End If
                    If mbEditMode Then
                        If MsgBox(sMess, vbYesNo + vbQuestion, "Update Service Fee Amount") = vbYes Then
                            txtServiceFee.Text = Format(mdblOverridesActLogTime * mcFeeServiceHourlyRate, "#,###,###,##0.00")
                            sServiceFeeComment = Format(mdblOverridesActLogTime, "##0.00") & " @ " & Format(mcFeeServiceHourlyRate, "##0.00")
                            txtServiceFeeComment.Text = sServiceFeeComment
                            txtServiceFeeComment.ToolTipText = sServiceFeeComment
                        End If
                    Else
                        MsgBox sMess, vbExclamation + vbOKOnly, "Service Fee not balanced"
                    End If
                End If
            End If
            txtActLogHours.Text = mdblOverridesActLogTime
        Else
            'Check to See if the currenct Acttime Selection is different from
            'the current service fee amount.  If it Is need to inform user that
            'Rebill is necesarry to sycnh up the amounts
            If chkFeeByTime.Value = vbChecked Then
                If Format(mdblCurrentActLogTime * mcFeeServiceHourlyRate, "#,###,###,##0.00") <> txtServiceFee.Text Then
                    sMess = "The Current Activity Log Time for Fee By Time values do not match the Service Fee Amount!" & vbCrLf
                    If mbEditMode Then
                        sMess = sMess & "Do you want to update the Service Fee Amount?"
                    Else
                        sMess = sMess & "You must Edit this IB to update the Service Fee Amount."
                    End If
                    If mbEditMode Then
                        If MsgBox(sMess, vbYesNo + vbQuestion, "Update Service Fee Amount") = vbYes Then
                            txtServiceFee.Text = Format(mdblCurrentActLogTime * mcFeeServiceHourlyRate, "#,###,###,##0.00")
                            sServiceFeeComment = mdblCurrentActLogTime & " @ " & mcFeeServiceHourlyRate
                            txtServiceFeeComment.Text = sServiceFeeComment
                            txtServiceFeeComment.ToolTipText = sServiceFeeComment
                        End If
                    Else
                        MsgBox sMess, vbExclamation + vbOKOnly, "Service Fee not balanced"
                    End If
                End If
            Else
                'If not overriding the Fee by Time
                'And not doing Fee By time
                'and not Overriding the Service Fee
                'and not overriding ALL fees
                'Check to see if the Current Service Fee Matches the Last calculated Schedule
                If Not mbOverridesServiceFee And Not mbOverridesALL Then
                    If CCur(txtServiceFee.Text) > 0 Then
                        If CCur(txtServiceFee.Text) <> Format(mcCurrentServiceFee, "#,###,###,##0.00") Then
                            If ChkOverrideAmounts.Value = vbUnchecked Then
                                sMess = "The Service Fee Amount (" & txtServiceFee.Text & ")  does not match the Calculated Feeschedule Amount (" & Format(mcCurrentServiceFee, "#,###,###,##0.00") & ")!" & vbCrLf
                                If mbEditMode Then
                                    sMess = sMess & "Do you want to update the Service Fee Amount?"
                                Else
                                    sMess = sMess & "You must Edit this IB to update the Service Fee Amount."
                                End If
                                If mbEditMode Then
                                    If MsgBox(sMess, vbYesNo + vbQuestion, "Update Service Fee Amount") = vbYes Then
                                        CalcFeeSched
                                    End If
                                Else
                                    MsgBox sMess, vbExclamation + vbOKOnly, "Service Fee not balanced"
                                End If
                            Else
                                CalcFeeSched
                            End If
                        End If
                    Else
                        'If service fee is currently zero and not overriding, and in edit mode
                        'Go ahead and update the amount
                        If mbEditMode Then
                            CalcFeeSched
                        End If
                    End If
                End If
            End If
            txtActLogHours.Text = Format(mdblCurrentActLogTime, "##0.00")
        End If
        '--------------------------------------END OPEN IB---------------------------------
    End If
    
    'Check some Indemnity totals that do not match items that will
    'Print on IB.  Give appropriate message
    'Only show this if the IB is Closed or Open and not in Edit mode.
    sMess = vbNullString
    If Not mbEditMode Then
        sSubToCarrier = goUtil.IsNullIsVbNullString(pRSIB.Fields(sIB & "01_sSubToCarrier"))
        sIbnumber = goUtil.IsNullIsVbNullString(pRSIB.Fields(sIB & "02_sIBNumber"))
        sLocation = goUtil.IsNullIsVbNullString(pRSIB.Fields(sIB & "05_sLocation"))
        sState = goUtil.IsNullIsVbNullString(pRSIB.Fields(sIB & "05a_sState"))
        sAdjusterName = goUtil.IsNullIsVbNullString(pRSIB.Fields(sIB & "07_sAdjusterName"))
        sSALN = goUtil.IsNullIsVbNullString(pRSIB.Fields(sIB & "09_sSALN"))
        sPolicyNo = goUtil.IsNullIsVbNullString(pRSIB.Fields(sIB & "09a_sPolicyNo"))
        sInsuredName = goUtil.IsNullIsVbNullString(pRSIB.Fields(sIB & "10_sInsuredName"))
        sLossLocation = goUtil.IsNullIsVbNullString(pRSIB.Fields(sIB & "11_sLossLocation"))
        sDateOfLoss = goUtil.IsNullIsVbNullString(pRSIB.Fields(sIB & "12_dtDateOfLoss"))
        If StrComp(sSubToCarrier, msCurrentSubToCarrier, vbTextCompare) <> 0 Then
            sMess = sMess & "Submitted to (" & sSubToCarrier & ") "
            sMess = sMess & " <> updated Submitted to (" & msCurrentSubToCarrier & ") "
            sMess = sMess & vbCrLf
        End If
        If StrComp(sIbnumber, msCurrentIBNumber, vbTextCompare) <> 0 Then
            sMess = sMess & "IB # (" & sIbnumber & ") "
            sMess = sMess & " <> updated IB # (" & msCurrentIBNumber & ") "
            sMess = sMess & vbCrLf
        End If
        If StrComp(sLocation, msCurrentLocation, vbTextCompare) <> 0 Then
            sMess = sMess & "Site City (" & sLocation & ") "
            sMess = sMess & " <> updated Site City (" & msCurrentLocation & ") "
            sMess = sMess & vbCrLf
        End If
        If StrComp(sState, msCurrentState, vbTextCompare) <> 0 Then
            sMess = sMess & "Site State (" & sState & ") "
            sMess = sMess & " <> updated Site State  (" & msCurrentState & ") "
            sMess = sMess & vbCrLf
        End If
        If StrComp(sAdjusterName, msCurrentAdjusterName, vbTextCompare) <> 0 Then
            sMess = sMess & "Adjuster (" & sAdjusterName & ") "
            sMess = sMess & " <> updated Adjuster (" & msCurrentAdjusterName & ") "
            sMess = sMess & vbCrLf
        End If
        If StrComp(sSALN, msCurrentSALN, vbTextCompare) <> 0 Then
            sMess = sMess & "Claim # (" & sSALN & ") "
            sMess = sMess & " <> updated Claim # (" & msCurrentSALN & ") "
            sMess = sMess & vbCrLf
        End If
        If StrComp(sPolicyNo, msCurrentPolicyNo, vbTextCompare) <> 0 Then
            sMess = sMess & "Policy No # (" & sPolicyNo & ") "
            sMess = sMess & " <> updated Policy No # (" & msCurrentPolicyNo & ") "
            sMess = sMess & vbCrLf
        End If
        If StrComp(sInsuredName, msCurrentInsuredName, vbTextCompare) <> 0 Then
            sMess = sMess & "Insured (" & sInsuredName & ") "
            sMess = sMess & " <> updated Insured  (" & msCurrentInsuredName & ") "
            sMess = sMess & vbCrLf
        End If
        If StrComp(sLossLocation, msCurrentLossLocation, vbTextCompare) <> 0 Then
            sMess = sMess & "Loss Location (" & sLossLocation & ") "
            sMess = sMess & " <> updated Loss Location (" & msCurrentLossLocation & ") "
            sMess = sMess & vbCrLf
        End If
        If Format(sDateOfLoss, "MM/DD/YYYY") <> Format(msCurrentDateOfLoss, "MM/DD/YYYY") Then
            sMess = sMess & "Date Of Loss (" & Format(sDateOfLoss, "MM/DD/YYYY") & ") "
            sMess = sMess & " <> updated Date Of Loss (" & Format(msCurrentDateOfLoss, "MM/DD/YYYY") & ") "
            sMess = sMess & vbCrLf
        End If
    End If
    
    If sMess <> vbNullString Then
        sTemp = sMess
        sMess = "Some Administrative items listed below have changed." & vbCrLf
        sMess = sMess & "(These items are displayed on the IB, but are hidden on this form.)" & vbCrLf & vbCrLf & sTemp & vbCrLf
        If mbClosedIB Then
            'If the IB is closed inform the user that a rebill is necessary to update this item
            sMess = sMess & "You must Rebill this item to update."
        Else
            sMess = sMess & "You must Edit and Save this item to update."
        End If
        MsgBox sMess, vbExclamation + vbOKOnly, "Administrative Items have changed."
    End If
    
    'Check some Administrative info that do not match items that will
    'Print on IB.  Give appropriate message
    'Only show this if the IB is Closed or Open and not in Edit mode.
        sMess = vbNullString
     If Not mbEditMode Then
        If cGrossLoss <> mcCurrentGrossLoss And ChkOverrideAmounts.Value = vbUnchecked Then
            sMess = sMess & "Gross Loss (" & Format(cGrossLoss, "#,###,###,##0.00") & ") "
            sMess = sMess & " <> updated Gross Loss(" & Format(mcCurrentGrossLoss, "#,###,###,##0.00") & ") "
            sMess = sMess & vbCrLf
        End If
        If cDepreciation <> mcCurrentDepreciation And ChkOverrideAmounts.Value = vbUnchecked Then
            sMess = sMess & "Depreciation (" & Format(cDepreciation, "#,###,###,##0.00") & ") "
            sMess = sMess & " <> updated Depreciation(" & Format(mcCurrentDepreciation, "#,###,###,##0.00") & ") "
            sMess = sMess & vbCrLf
        End If
        If cDeductible <> mcCurrentDeductible Then
            sMess = sMess & "Deductible (" & Format(cDeductible, "#,###,###,##0.00") & ") "
            sMess = sMess & " <> updated Deductible(" & Format(mcCurrentDeductible, "#,###,###,##0.00") & ") "
            sMess = sMess & vbCrLf
        End If
        If cLessExcessLimits <> mcCurrentLessExcessLimits And ChkOverrideAmounts.Value = vbUnchecked Then
            sMess = sMess & "ExcessLimits (" & Format(cLessExcessLimits, "#,###,###,##0.00") & ") "
            sMess = sMess & " <> updated ExcessLimits(" & Format(mcCurrentLessExcessLimits, "#,###,###,##0.00") & ") "
            sMess = sMess & vbCrLf
        End If
        If cLessMiscellaneous <> mcCurrentLessMiscellaneous And ChkOverrideAmounts.Value = vbUnchecked Then
            sMess = sMess & "Miscellaneous (" & Format(cLessMiscellaneous, "#,###,###,##0.00") & ") "
            sMess = sMess & " <> updated Miscellaneous(" & Format(mcCurrentLessMiscellaneous, "#,###,###,##0.00") & ") "
            sMess = sMess & vbCrLf
        End If
        If cNetClaim <> mcCurrnetNetClaim And ChkOverrideAmounts.Value = vbUnchecked Then
            sMess = sMess & "Net Claim (" & Format(cNetClaim, "#,###,###,##0.00") & ") "
            sMess = sMess & " <> updated Net Claim(" & Format(mcCurrnetNetClaim, "#,###,###,##0.00") & ") "
            sMess = sMess & vbCrLf
        End If
    End If
    
    If sMess <> vbNullString Then
        sTemp = sMess
        sMess = "Some indemnity items listed below have changed." & vbCrLf
        sMess = sMess & "(These items are displayed on the IB, but are hidden on this form.)" & vbCrLf & vbCrLf & sTemp & vbCrLf
        If mbClosedIB Then
            'If the IB is closed inform the user that a rebill is necessary to update this item
            sMess = sMess & "You must Rebill this item to update."
        Else
            sMess = sMess & "You must Edit and Save this item to update."
        End If
        MsgBox sMess, vbExclamation + vbOKOnly, "Indemnity Items have changed."
    End If
    

    txtFeeServiceHourlyRate.Text = Format(mcFeeServiceHourlyRate, "##0.00")
    
    'Additional Service Fees
    sValue = goUtil.IsNullIsVbNullString(pRSIB.Fields(sIB & "18a_sMiscServiceFeeComment"))
    txtMiscServiceFeeComment.Text = sValue
    
    sValue = goUtil.IsNullIsVbNullString(pRSIB.Fields(sIB & "17a_cMiscServiceFee"))
    txtMiscServiceFee.Text = Format(sValue, "#,###,###,##0.00")
    
    sValue = goUtil.IsNullIsVbNullString(pRSIB.Fields(sIB & "25_cServiceFeeSubTotal"))
    txtTtlServiceFee.Text = Format(sValue, "#,###,###,##0.00")
    
    'Additional Expenses
    
    For Each itmX In lvwExpenseFees.ListItems
        sfsftName = itmX.ListSubItems(GuiExpenseFeesListView.fsftName - 1)
        lNumberOfItems = itmX.ListSubItems(GuiExpenseFeesListView.NumberOfItems - 1)
        sfsftDescription = itmX.Text
        If StrComp(sfsftName, "FEE_PER_PHOTO", vbTextCompare) = 0 Then
            'Check the current photo count against what is
            'recorded under the additional Expens item
            'If it does not match give opportunity to update
            If lNumberOfItems <> mlCurrentPhotoCount Then
                sMess = "The current number of photos (" & mlCurrentPhotoCount & ") does not match "
                sMess = sMess & "( " & lNumberOfItems & ") under the " & framExpenses.Caption & "  " & sfsftDescription & " Item!" & vbCrLf
                If mbClosedIB Then
                    sMess = sMess & "You must rebill this IB to correct."
                    MsgBox sMess, vbExclamation + vbOKOnly, sfsftDescription
                Else
                    If mbEditMode Then
                        sMess = sMess & "Do you want to update this item?"
                        If MsgBox(sMess, vbQuestion + vbYesNo, sfsftDescription) = vbYes Then
                            ShowEditExpenseFeesItem itmX
                            lstAddExpenseFeeItemNumberOfItems.Text = mlCurrentPhotoCount
                            CalcAddExpenseFeeItem
                        End If
                    Else
                        sMess = sMess & "You must EDIT this IB to correct."
                        MsgBox sMess, vbExclamation + vbOKOnly, sfsftDescription
                    End If
                End If
            End If
        End If
    Next
    
    sValue = goUtil.IsNullIsVbNullString(pRSIB.Fields(sIB & "29a_sMiscExpenseFeeComment"))
    txtMiscExpenseFeeComment.Text = sValue
    
    sValue = goUtil.IsNullIsVbNullString(pRSIB.Fields(sIB & "29b_cMiscExpenseFee"))
    txtMiscExpenseFee.Text = Format(sValue, "#,###,###,##0.00")
    
    sValue = goUtil.IsNullIsVbNullString(pRSIB.Fields(sIB & "30_cTotalExpenses"))
    txtTtlExpenses.Text = Format(sValue, "#,###,###,##0.00")
    
    'Invoice Total
    sValue = goUtil.IsNullIsVbNullString(pRSIB.Fields(sIB & "06_dtDateClosed"))
    txtdtDateClosed.Text = Format(sValue, "MM/DD/YYYY")
    
    sValue = goUtil.IsNullIsVbNullString(pRSIB.Fields("UpLoadMe"))
    If CBool(sValue) Then
        imgUploadMe.Picture = imgBillingStatus.ListImages(GuiIBStatusList.UpLoadMe).Picture
    Else
        imgUploadMe.Picture = LoadPicture()
    End If
    
    sValue = goUtil.IsNullIsVbNullString(pRSIB.Fields("Void"))
    If CBool(sValue) Then
        imgVoid.Picture = imgBillingStatus.ListImages(GuiIBStatusList.IsDeleted).Picture
    Else
        imgVoid.Picture = LoadPicture()
    End If
    
    sValue = goUtil.IsNullIsVbNullString(pRSIB.Fields("AdminComments"))
    If sValue <> vbNullString Then
        imgAdminComments.Picture = imgBillingStatus.ListImages(GuiIBStatusList.showPDF).Picture
        msAdminComments = sValue
    Else
        imgAdminComments.Picture = LoadPicture()
        msAdminComments = vbNullString
    End If
    
    sValue = goUtil.IsNullIsVbNullString(pRSIB.Fields("DateLastUpdated"))
    txtDateLastUpdated.Text = sValue
    
    sValue = CCur(txtTtlServiceFee.Text) + CCur(txtTtlExpenses.Text)
    txtTtlServiceExp.Text = Format(sValue, "#,###,###,##0.00")
    
    sValue = goUtil.IsNullIsVbNullString(pRSIB.Fields(sIB & "31_dTaxPercent"))
    txtTaxPercent.Text = Format(sValue, "#0.000")
    
    sValue = goUtil.IsNullIsVbNullString(pRSIB.Fields(sIB & "32_cTaxAmount"))
    txtTaxAmount.Text = Format(sValue, "#,###,###,##0.00")
    
    sValue = goUtil.IsNullIsVbNullString(pRSIB.Fields(sIB & "33_cTotalAdjustingFee"))
    txtTotalAdjustingFee.Text = Format(sValue, "#,###,###,##0.00")
    
    
    mbPopulateIBInfo = False
    
    SumServiceFees
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub PopulateIBInfo"
End Sub

Public Sub LoadCurrentValues()
    On Error GoTo EH
    Dim lBillingCountID As Long
    Dim RSAssgn As ADODB.Recordset
    Dim RSACID As ADODB.Recordset
    Dim RSCatCode As ADODB.Recordset
    Dim RSClientCoCat As ADODB.Recordset
    Dim RSFeeSchedule As ADODB.Recordset
    Dim RSBillingCountItem As ADODB.Recordset
    'Used for Getting Reg Setting init options
    Dim sFeeScheduleID As String
    Dim sClientCompanyID As String
    Dim sSection As String
    
    lBillingCountID = CLng(msBillingCountID)
    
    mcCurrentServiceFee = mfrmClaim.GetCurrentServiceFee(lBillingCountID, msFeeScheduleID)
    mdblCurrentActLogTime = mfrmClaim.GetCurrentActLogTime(msAssignmentsID, msBillingCountID)
    mlCurrentPhotoCount = mfrmClaim.GetCurrentPhotoCount(msAssignmentsID, msBillingCountID)
    mcCurrentAmountOfCheck = mfrmClaim.GetCurrentAmountOfCheck(msAssignmentsID, msBillingCountID)
    mcFeeServiceHourlyRate = mfrmClaim.GetCurrentFeeScheduleItem("FeeServiceHourlyRate", msFeeScheduleID)
    mdblCurrentTaxPercent = mfrmClaim.GetCurrentFeeScheduleItem("TaxPercent", msFeeScheduleID)
    
    mfrmClaim.SetadoRSAssignments msAssignmentsID
    'Current Items not shown on Form but will show on printed form
    'Need to be compared
    
    'Set the member variables here
    mcCurrentGrossLoss = mfrmClaim.GetFullCostOfRepair(CLng(msBillingCountID))
    mcCurrentDepreciation = mfrmClaim.GetDepreciation(CLng(msBillingCountID))
    mcCurrentDeductible = goUtil.IsNullIsVbNullString(mfrmClaim.adoRSAssignments.Fields("Deductible"))
    mcCurrentLessExcessLimits = mfrmClaim.GetLessExcessLimits(CLng(msBillingCountID))
    mcCurrentLessMiscellaneous = mfrmClaim.GetLessMiscellaneous(CLng(msBillingCountID))
    mcCurrnetNetClaim = mfrmClaim.GetNetActualCashValueClaim(CLng(msBillingCountID))
    
    'Need to get updated info to compare to stuff on the IB
    If Not mfrmClaim.SetadoRSAssignments(msAssignmentsID) Then
        GoTo CLEAN_UP
    End If
    If Not moGUI.SetadoRSACID Then
        GoTo CLEAN_UP
    End If
    If Not moGUI.SetadoRSCatCode Then
        GoTo CLEAN_UP
    End If
    If Not mfrmClaim.SetadoRSClientCOCat Then
        GoTo CLEAN_UP
    End If
    If Not moGUI.SetadoRSFeeSchedule Then
        GoTo CLEAN_UP
    End If
    If Not mfrmClaim.SetadoRSBillingCountItem(msAssignmentsID, msBillingCountID, False) Then
        GoTo CLEAN_UP
    End If
    
    Set RSAssgn = mfrmClaim.adoRSAssignments
    Set RSACID = moGUI.adoRSACID
    Set RSCatCode = moGUI.adoRSCatCode
    Set RSClientCoCat = mfrmClaim.adoRSClientCOCat
    Set RSFeeSchedule = moGUI.adoFeeSchedule
    Set RSBillingCountItem = mfrmClaim.adoRSBillingCountItem
    
    'Used for Getting Reg Setting init options
    sFeeScheduleID = goUtil.IsNullIsVbNullString(RSFeeSchedule.Fields("FeeScheduleID"))
    sClientCompanyID = goUtil.IsNullIsVbNullString(RSFeeSchedule.Fields("ClientCompanyID"))
    sSection = sFeeScheduleID & "_" & sClientCompanyID
    
    
    'Set the memebr variables here
    msCurrentSubToCarrier = goUtil.IsNullIsVbNullString(RSACID.Fields("ClientCompanyDesc"))
    msCurrentIBNumber = goUtil.IsNullIsVbNullString(RSAssgn.Fields("IBNUM"))
    msCurrentLocation = goUtil.IsNullIsVbNullString(RSClientCoCat.Fields("SACity"))
    msCurrentState = goUtil.IsNullIsVbNullString(RSClientCoCat.Fields("SAState"))
    msCurrentAdjusterName = goUtil.IsNullIsVbNullString(RSACID.Fields("LFName"))
    msCurrentSALN = goUtil.IsNullIsVbNullString(RSAssgn.Fields("CLIENTNUM"))
    msCurrentPolicyNo = goUtil.IsNullIsVbNullString(RSAssgn.Fields("PolicyNo"))
    msCurrentInsuredName = goUtil.IsNullIsVbNullString(RSAssgn.Fields("Insured"))
    msCurrentLossLocation = goUtil.IsNullIsVbNullString(RSAssgn.Fields("PAStreet")) & vbCrLf
    msCurrentLossLocation = msCurrentLossLocation & goUtil.IsNullIsVbNullString(RSAssgn.Fields("PACity")) & ", "
    msCurrentLossLocation = msCurrentLossLocation & goUtil.IsNullIsVbNullString(RSAssgn.Fields("PAState")) & " "
    msCurrentLossLocation = msCurrentLossLocation & Format(goUtil.IsNullIsVbNullString(RSAssgn.Fields("PAZIP")), "00000") & " - "
    msCurrentLossLocation = msCurrentLossLocation & Format(goUtil.IsNullIsVbNullString(RSAssgn.Fields("PAZIP4")), "0000")
    msCurrentDateOfLoss = goUtil.IsNullIsVbNullString(RSAssgn.Fields("LossDate"))
    
CLEAN_UP:
    'cleanup
    Set RSAssgn = Nothing
    Set RSACID = Nothing
    Set RSCatCode = Nothing
    Set RSClientCoCat = Nothing
    Set RSFeeSchedule = Nothing
    Set RSBillingCountItem = Nothing
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub LoadCurrentValues"
End Sub

Public Sub RebillIB()
    On Error GoTo EH
    Dim RSBillCountItem As ADODB.Recordset
    Dim RSRTIB As ADODB.Recordset
    Dim bInsertRecord As Boolean
    Dim RSRTIBFee As ADODB.Recordset
    Dim sRTIBFeeID As String
    Dim RSCLIBFee As ADODB.Recordset
    Dim sSQL As String
    Dim RSFeeScheduleFeeTypes As ADODB.Recordset
    Dim oConn As ADODB.Connection
    Dim MyRTIBFeeItem As GuiRTIBFeeItem
    
    If Not mfrmClaim.SetadoRSBillingCountItem(msAssignmentsID, msBillingCountID) Then
        GoTo CLEAN_UP
    End If
    If Not mfrmClaim.SetadoRSRTIB(msAssignmentsID) Then
        GoTo CLEAN_UP
    End If
    If Not mfrmClaim.SetadoRSRTIBFee(msAssignmentsID) Then
        GoTo CLEAN_UP
    End If
    If Not mfrmClaim.SetadoRSCLIBFee(msAssignmentsID, msBillingCountID) Then
        GoTo CLEAN_UP
    End If
    
    Set RSBillCountItem = mfrmClaim.adoRSBillingCountItem
    Set RSRTIB = mfrmClaim.adoRSRTIB
    Set RSRTIBFee = mfrmClaim.adoRSRTIBFee
    Set RSCLIBFee = mfrmClaim.adoRSCLIBFee
    
    'Need to initilize the RTIB table for this Item
    'If the RSRTIB recordcount is 0 then need to insert a record for it
    If RSRTIB.RecordCount = 0 Then
        bInsertRecord = True
    End If
    If Not InitRTIB(bInsertRecord, True, RSBillCountItem, msFeeScheduleID) Then
        GoTo CLEAN_UP
    End If
    
    'First need to see if the current RTIBFee table has all the lastest
    'fee types in it.  If not then need to insert them
    
    Set oConn = New ADODB.Connection
    Set RSFeeScheduleFeeTypes = New ADODB.Recordset
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    sSQL = "SELECT RetFeeScheduleFeeTypes.* "
    sSQL = sSQL & "FROM( "
    sSQL = sSQL & "SELECT "
    sSQL = sSQL & "[FeeScheduleFeeTypesID], "
    sSQL = sSQL & "[FeeScheduleID], "
    sSQL = sSQL & "[TypeNum], "
    sSQL = sSQL & "[Name], "
    sSQL = sSQL & "[Description], "
    sSQL = sSQL & "[FeeAmount], "
    sSQL = sSQL & "[IsExpense], "
    sSQL = sSQL & "[MaxNumberOfItems], "
    sSQL = sSQL & "[MaxFeeAmount], "
    sSQL = sSQL & "[IsMiscAmount], "
    sSQL = sSQL & "[UseFormula], "
    sSQL = sSQL & "[VBFormula], "
    sSQL = sSQL & "[IsDeleted], "
    sSQL = sSQL & "[DateLastUpdated], "
    sSQL = sSQL & "(SELECT    USERNAME "
    sSQL = sSQL & "FROM   USERS "
    sSQL = sSQL & "WHERE  USERSID = F.[UpdateByUserID]) As [UpdateByUserName],  "
    sSQL = sSQL & "[UpdateByUserID] "
    sSQL = sSQL & "FROM FeeScheduleFeeTypes F "
    sSQL = sSQL & ") RetFeeScheduleFeeTypes "
    If msFeeScheduleID = vbNullString Then
        sSQL = sSQL & "WHERE [FeeScheduleID] = ( "
                                sSQL = sSQL & "SELECT   [FeeScheduleID] "
                                sSQL = sSQL & "FROM     ClientCompanyCat "
                                sSQL = sSQL & "WHERE    [ClientCompanyID] = " & goUtil.gsCurCar & " "
                                sSQL = sSQL & "AND      [CATID] = " & goUtil.gsCurCat & " "
                                sSQL = sSQL & ") "
    Else
        sSQL = sSQL & "WHERE [FeeScheduleID] = ( "
                                sSQL = sSQL & "SELECT   " & msFeeScheduleID & " As [FeeScheduleID] "
                                sSQL = sSQL & "FROM     ClientCompanyCat "
                                sSQL = sSQL & "WHERE    [ClientCompanyID] = " & goUtil.gsCurCar & " "
                                sSQL = sSQL & "AND      [CATID] = " & goUtil.gsCurCat & " "
                                sSQL = sSQL & ") "
    End If
    'Exclude any feetypes that already exist in the RTIBFee table for this Assignment
    sSQL = sSQL & "AND FeeScheduleFeeTypesID Not IN ( "
                                    sSQL = sSQL & "SELECT   FeeScheduleFeeTypesID "
                                    sSQL = sSQL & "FROM     RTIBfee "
                                    sSQL = sSQL & "WHERE    AssignmentsID = " & msAssignmentsID & " "
                                    sSQL = sSQL & ") "
    sSQL = sSQL & "ORDER BY [TypeNum] "
    
    RSFeeScheduleFeeTypes.CursorLocation = adUseClient
    RSFeeScheduleFeeTypes.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set RSFeeScheduleFeeTypes.ActiveConnection = Nothing
        
    If RSFeeScheduleFeeTypes.RecordCount > 0 Then
        RSFeeScheduleFeeTypes.MoveFirst
    End If
    Do Until RSFeeScheduleFeeTypes.EOF
        With MyRTIBFeeItem
            .RTIBFeeID = "[RTIBFeeID]"
            .AssignmentsID = msAssignmentsID
            .ID = "[ID]"
            .IDAssignments = msAssignmentsID
            .FeeScheduleFeeTypesID = goUtil.IsNullIsVbNullString(RSFeeScheduleFeeTypes.Fields("FeeScheduleFeeTypesID"))
            .NumberOfItems = "0"
            .Amount = "0.00"
            .Comment = vbNullString
            .DownLoadMe = "False"
            .UpLoadMe = "True"
            .AdminComments = vbNullString
            .DateLastUpdated = Now()
            .UpdateByUserID = goUtil.gsCurUsersID
        End With
        If Not EditRTIBFee(MyRTIBFeeItem) Then
            If Not AddRTIBFee(MyRTIBFeeItem, sRTIBFeeID) Then
                GoTo CLEAN_UP
            End If
        End If
        RSFeeScheduleFeeTypes.MoveNext
    Loop
    
    'Need to Update RTIbfee items with previous amounts
    'since this is a rebill
    'If there are items in the OLD RS then loop through them
    'and add
    If RSCLIBFee.RecordCount > 0 Then
        RSCLIBFee.MoveFirst
    End If
    Do Until RSCLIBFee.EOF
        With MyRTIBFeeItem
            .RTIBFeeID = "[RTIBFeeID]"
            .AssignmentsID = msAssignmentsID
            .ID = "[ID]"
            .IDAssignments = msAssignmentsID
            .FeeScheduleFeeTypesID = goUtil.IsNullIsVbNullString(RSCLIBFee.Fields("FeeScheduleFeeTypesID"))
            .NumberOfItems = goUtil.IsNullIsVbNullString(RSCLIBFee.Fields("NumberOfItems"))
            .Amount = goUtil.IsNullIsVbNullString(RSCLIBFee.Fields("Amount"))
            .Comment = goUtil.IsNullIsVbNullString(RSCLIBFee.Fields("Comment"))
            .DownLoadMe = "False"
            .UpLoadMe = "True"
            .AdminComments = goUtil.IsNullIsVbNullString(RSCLIBFee.Fields("AdminComments"))
            .DateLastUpdated = Now()
            .UpdateByUserID = goUtil.gsCurUsersID
        End With
        If Not EditRTIBFee(MyRTIBFeeItem) Then
            If Not AddRTIBFee(MyRTIBFeeItem, sRTIBFeeID) Then
                GoTo CLEAN_UP
            End If
        End If
        RSCLIBFee.MoveNext
    Loop

    'After Rebilling need to Refresh Entire Claim
    mfrmClaim.RefreshMe
        
CLEAN_UP:
    Set RSBillCountItem = Nothing
    Set RSRTIB = Nothing
    Set RSRTIBFee = Nothing
    Set RSCLIBFee = Nothing
    Set RSFeeScheduleFeeTypes = Nothing
    Set oConn = Nothing
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub RebillIB"
End Sub

Public Sub SupplementIB()
    On Error GoTo EH
    Dim RSBillCountItem As ADODB.Recordset
    Dim RSRTIB As ADODB.Recordset
    Dim bInsertRecord As Boolean
    Dim RSRTIBFee As ADODB.Recordset
    Dim sRTIBFeeID As String
    Dim RSFeeScheduleFeeTypes As ADODB.Recordset
    Dim MyRTIBFeeItem As GuiRTIBFeeItem
    
    If Not mfrmClaim.SetadoRSBillingCountItem(msAssignmentsID, msBillingCountID, True) Then
        GoTo CLEAN_UP
    End If
    If Not mfrmClaim.SetadoRSRTIB(msAssignmentsID) Then
        GoTo CLEAN_UP
    End If
    If Not mfrmClaim.SetadoRSRTIBFee(msAssignmentsID) Then
        GoTo CLEAN_UP
    End If
    If Not moGUI.SetadoRSFeeScheduleFeeTypes(msFeeScheduleID) Then
        GoTo CLEAN_UP
    End If
    
    Set RSBillCountItem = mfrmClaim.adoRSBillingCountItem
    Set RSRTIB = mfrmClaim.adoRSRTIB
    Set RSRTIBFee = mfrmClaim.adoRSRTIBFee
    Set RSFeeScheduleFeeTypes = moGUI.adoFeeScheduleFeeTypes
    
    'Need to initilize the RTIB table for this Item
    'If the RSRTIB recordcount is 0 then need to insert a record for it
    If RSRTIB.RecordCount = 0 Then
        bInsertRecord = True
    End If
    If Not InitRTIB(bInsertRecord, False, RSBillCountItem) Then
        GoTo CLEAN_UP
    End If

    'First need to see if the current RTIBFee table has all the lastest
    'fee types in it.  If not then need to insert them
    If RSFeeScheduleFeeTypes.RecordCount > 0 Then
        RSFeeScheduleFeeTypes.MoveFirst
    End If
    Do Until RSFeeScheduleFeeTypes.EOF
        With MyRTIBFeeItem
            .RTIBFeeID = "[RTIBFeeID]"
            .AssignmentsID = msAssignmentsID
            .ID = "[ID]"
            .IDAssignments = msAssignmentsID
            .FeeScheduleFeeTypesID = goUtil.IsNullIsVbNullString(RSFeeScheduleFeeTypes.Fields("FeeScheduleFeeTypesID"))
            .NumberOfItems = "0"
            .Amount = "0.00"
            .Comment = vbNullString
            .DownLoadMe = "False"
            .UpLoadMe = "True"
            .AdminComments = vbNullString
            .DateLastUpdated = Now()
            .UpdateByUserID = goUtil.gsCurUsersID
        End With
        If Not EditRTIBFee(MyRTIBFeeItem) Then
            If Not AddRTIBFee(MyRTIBFeeItem, sRTIBFeeID) Then
                GoTo CLEAN_UP
            End If
        End If
        RSFeeScheduleFeeTypes.MoveNext
    Loop
    
    'After Rebilling need to Refresh Entire Claim
    mfrmClaim.RefreshMe
        
CLEAN_UP:
    Set RSBillCountItem = Nothing
    Set RSRTIB = Nothing
    Set RSRTIBFee = Nothing
    Set RSFeeScheduleFeeTypes = Nothing
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub SupplementIB"
End Sub


Public Function InitRTIB(pbInsertRecord As Boolean, pbRebill As Boolean, pRSBillCountItem As ADODB.Recordset, Optional psFeeScheduleID As String) As Boolean
    On Error GoTo EH
    Dim RSCLIB As ADODB.Recordset
    Dim RSAssgn As ADODB.Recordset
    Dim RSACID As ADODB.Recordset
    Dim RSCatCode As ADODB.Recordset
    Dim RSClientCoCat As ADODB.Recordset
    Dim RSFeeSchedule As ADODB.Recordset
    Dim sFeeScheduleID As String
    Dim sClientCompanyID As String
    Dim sSection As String
    Dim lCurRebill As Long
    Dim lRebill As Long
    Dim lCurSupplement As Long
    Dim lSupplement As Long
    Dim sIDBillingCount As String
    Dim MyudtBillingCountItem As GuiBillingCountItem
    Dim MyudtRTIBItem As GuiRTIBItem
    Dim MyudtRTIBfeeItem As GuiRTIBFeeItem
    
    If pbRebill Then
        'Need to update the BillingCount Rebill
        lCurRebill = pRSBillCountItem.Fields("Rebill")
        lRebill = lCurRebill + 1
        With MyudtBillingCountItem
            .AdminComments = pRSBillCountItem.Fields("AdminComments")
            .AssignmentsID = pRSBillCountItem.Fields("AssignmentsID")
            .BillingCountID = pRSBillCountItem.Fields("BillingCountID")
            .DateLastUpdated = Now()
            .DownLoadMe = pRSBillCountItem.Fields("DownLoadMe")
            .ID = pRSBillCountItem.Fields("ID")
            .IDAssignments = pRSBillCountItem.Fields("IDAssignments")
            .Rebill = lRebill
            .Supplement = pRSBillCountItem.Fields("Supplement")
            .UpdateByUserID = goUtil.gsCurUsersID
            .UpLoadMe = "True"
        End With
        'Update Assignment status to indicate rebill inprogress
        If Not mfrmClaim.UpdateAssgnStatus(iAssignmentsStatus_REOPEN) Then
            GoTo CLEAN_UP
        End If
        If Not EditBillingCount(MyudtBillingCountItem) Then
            GoTo CLEAN_UP
        End If
    Else
        'Supplementing
        'Need to insert another Billing Count
        'The pRSBillCountItem should either the last Billingcount
        'Or none at all
        If pRSBillCountItem.RecordCount = 0 Then
            ' The next supplement count will be 0
            '0 supplemtn is the very first Bill
            lCurSupplement = -1
        Else
            lCurSupplement = pRSBillCountItem.Fields("Supplement")
        End If
        lSupplement = lCurSupplement + 1
        With MyudtBillingCountItem
            .AdminComments = ""
            .AssignmentsID = msAssignmentsID
            .BillingCountID = "null" 'since adding can't set here
            .DateLastUpdated = Now()
            .DownLoadMe = "False"
            .ID = "null" 'since adding can't set here
            .IDAssignments = msAssignmentsID
            .Rebill = "0" 'Every New Bill starts rebill at 0
            .Supplement = lSupplement
            .UpdateByUserID = goUtil.gsCurUsersID
            .UpLoadMe = "True"
        End With
        'Update Assignment status to indicate Supplement (INTERIM Billing) inprogress
        'Only update to Interim if the supplement is > 0
        If lSupplement > 0 Then
            If Not mfrmClaim.UpdateAssgnStatus(iAssignmentsStatus_INTERIM) Then
                GoTo CLEAN_UP
            End If
        End If
        If Not AddBillingCount(MyudtBillingCountItem, sIDBillingCount) Then
            GoTo CLEAN_UP
        Else
            'After Adding the BillingCOunt Item need to update any Null BillingcountID and IDBilling Count
            'with the new sIDBillingCount.
            UpdateFkeyBillingCountID sIDBillingCount
        End If
    End If
    
    'Need to get updated info to populate stuff on the IB
    If Not mfrmClaim.SetadoRSAssignments(msAssignmentsID) Then
        GoTo CLEAN_UP
    End If
    If Not moGUI.SetadoRSACID Then
        GoTo CLEAN_UP
    End If
    If Not moGUI.SetadoRSCatCode Then
        GoTo CLEAN_UP
    End If
    If Not mfrmClaim.SetadoRSClientCOCat Then
        GoTo CLEAN_UP
    End If
    If Not moGUI.SetadoRSFeeSchedule Then
        GoTo CLEAN_UP
    End If
    
    Set RSAssgn = mfrmClaim.adoRSAssignments
    Set RSACID = moGUI.adoRSACID
    Set RSCatCode = moGUI.adoRSCatCode
    Set RSClientCoCat = mfrmClaim.adoRSClientCOCat
    Set RSFeeSchedule = moGUI.adoFeeSchedule
    
    'Used for Getting Reg Setting init options
    If psFeeScheduleID = vbNullString Then
        sFeeScheduleID = goUtil.IsNullIsVbNullString(RSFeeSchedule.Fields("FeeScheduleID"))
    Else
        sFeeScheduleID = psFeeScheduleID
    End If
    sClientCompanyID = goUtil.IsNullIsVbNullString(RSFeeSchedule.Fields("ClientCompanyID"))
    sSection = sFeeScheduleID & "_" & sClientCompanyID
    
    'If rebilling then need to Get the previous values
    If pbRebill Then
        If Not mfrmClaim.SetadoRSCLIB(msAssignmentsID, msBillingCountID) Then
            GoTo CLEAN_UP
        End If
        Set RSCLIB = mfrmClaim.adoRSCLIB
        With MyudtRTIBItem
            .AssignmentsID = msAssignmentsID
            .BillingCountID = goUtil.IsNullIsVbNullString(RSCLIB.Fields("BillingCountID"))
            .IDAssignments = msAssignmentsID
            .IDBillingCount = goUtil.IsNullIsVbNullString(RSCLIB.Fields("IDBillingCount"))
            .RT00_lSSN = goUtil.utGetECSCryptSetting("ECS", "WEB_SECURITY", "SSN")
            .RT01_sSubToCarrier = goUtil.IsNullIsVbNullString(RSACID.Fields("ClientCompanyDesc"))
            .RT02_sIBNumber = goUtil.IsNullIsVbNullString(RSAssgn.Fields("IBNUM"))
            .RT05_sLocation = goUtil.IsNullIsVbNullString(RSClientCoCat.Fields("SACity"))
            .RT05a_sState = goUtil.IsNullIsVbNullString(RSClientCoCat.Fields("SAState"))
            .RT06_dtDateClosed = "Null"
            .RT07_sAdjusterName = goUtil.IsNullIsVbNullString(RSACID.Fields("LFName"))
            .RT09_sSALN = goUtil.IsNullIsVbNullString(RSAssgn.Fields("CLIENTNUM"))
            .RT09a_sPolicyNo = goUtil.IsNullIsVbNullString(RSAssgn.Fields("PolicyNo"))
            .RT10_sInsuredName = goUtil.IsNullIsVbNullString(RSAssgn.Fields("Insured"))
            .RT11_sLossLocation = goUtil.IsNullIsVbNullString(RSAssgn.Fields("PAStreet")) & vbCrLf
            .RT11_sLossLocation = .RT11_sLossLocation & goUtil.IsNullIsVbNullString(RSAssgn.Fields("PACity")) & ", "
            .RT11_sLossLocation = .RT11_sLossLocation & goUtil.IsNullIsVbNullString(RSAssgn.Fields("PAState")) & " "
            .RT11_sLossLocation = .RT11_sLossLocation & Format(goUtil.IsNullIsVbNullString(RSAssgn.Fields("PAZIP")), "00000") & " - "
            .RT11_sLossLocation = .RT11_sLossLocation & Format(goUtil.IsNullIsVbNullString(RSAssgn.Fields("PAZIP4")), "0000")
            .RT12_dtDateOfLoss = IIf(goUtil.IsNullIsVbNullString(RSAssgn.Fields("LossDate")) = vbNullString, "Null", goUtil.IsNullIsVbNullString(RSAssgn.Fields("LossDate")))
            .RT13_cGrossLoss = goUtil.IsNullIsVbNullString(RSCLIB.Fields("IB13_cGrossLoss"))
            .RT14_cDepreciation = goUtil.IsNullIsVbNullString(RSCLIB.Fields("IB14_cDepreciation"))
            .RT14a_sSupplement = goUtil.IsNullIsVbNullString(RSCLIB.Fields("IB14a_sSupplement"))
            .RT14b_sRebilled = lRebill
            .RT15_cDeductible = goUtil.IsNullIsVbNullString(RSAssgn.Fields("Deductible"))
            .RT15a_cLessExcessLimits = goUtil.IsNullIsVbNullString(RSCLIB.Fields("IB15a_cLessExcessLimits"))
            .RT15b_sExcessLimDesc = vbNullString
            .RT15c_cLessMiscellaneous = goUtil.IsNullIsVbNullString(RSCLIB.Fields("IB15c_cLessMiscellaneous"))
            .RT15d_cMiscellaneousDesc = goUtil.IsNullIsVbNullString(RSCLIB.Fields("IB15d_cMiscellaneousDesc"))
            .RT16_cNetClaim = goUtil.IsNullIsVbNullString(RSCLIB.Fields("IB16_cNetClaim"))
            .RT17_cServiceFee = goUtil.IsNullIsVbNullString(RSCLIB.Fields("IB17_cServiceFee"))
            .RT17a_cMiscServiceFee = goUtil.IsNullIsVbNullString(RSCLIB.Fields("IB17a_cMiscServiceFee"))
            .RT18_sServiceFeeComment = goUtil.IsNullIsVbNullString(RSCLIB.Fields("IB18_sServiceFeeComment"))
            .RT18a_sMiscServiceFeeComment = goUtil.IsNullIsVbNullString(RSCLIB.Fields("IB18a_sMiscServiceFeeComment"))
            .RT25_cServiceFeeSubTotal = goUtil.IsNullIsVbNullString(RSCLIB.Fields("IB25_cServiceFeeSubTotal"))
            .RT29a_sMiscExpenseFeeComment = goUtil.IsNullIsVbNullString(RSCLIB.Fields("IB29a_sMiscExpenseFeeComment"))
            .RT29b_cMiscExpenseFee = goUtil.IsNullIsVbNullString(RSCLIB.Fields("IB29b_cMiscExpenseFee"))
            .RT30_cTotalExpenses = goUtil.IsNullIsVbNullString(RSCLIB.Fields("IB30_cTotalExpenses"))
            .RT31_dTaxPercent = goUtil.IsNullIsVbNullString(RSFeeSchedule.Fields("TaxPercent"))
            .RT32_cTaxAmount = goUtil.IsNullIsVbNullString(RSCLIB.Fields("IB32_cTaxAmount"))
            .RT33_cTotalAdjustingFee = goUtil.IsNullIsVbNullString(RSCLIB.Fields("IB33_cTotalAdjustingFee"))
            .RT33a_sAccountCode = goUtil.IsNullIsVbNullString(RSCLIB.Fields("IB33a_sAccountCode"))
            .FeeScheduleID = goUtil.IsNullIsVbNullString(RSCLIB.Fields("FeeScheduleID"))
            If .FeeScheduleID = "0" Then
                .FeeScheduleID = sFeeScheduleID
            End If
            psFeeScheduleID = .FeeScheduleID
            .Void = goUtil.IsNullIsVbNullString(RSCLIB.Fields("Void"))
            .FeeByTime = goUtil.IsNullIsVbNullString(RSCLIB.Fields("FeeByTime"))
            .UseActivityTime = goUtil.IsNullIsVbNullString(RSCLIB.Fields("UseActivityTime"))
            .DownLoadMe = goUtil.IsNullIsVbNullString(RSCLIB.Fields("DownLoadMe"))
            .UpLoadMe = "True"
            .AdminComments = goUtil.IsNullIsVbNullString(RSCLIB.Fields("AdminComments"))
            .DateLastUpdated = Now()
            .UpdateByUserID = goUtil.gsCurUsersID
        End With
    Else
        'Supplementing
        With MyudtRTIBItem
            .AssignmentsID = msAssignmentsID
            .BillingCountID = sIDBillingCount
            .IDAssignments = msAssignmentsID
            .IDBillingCount = sIDBillingCount
            .RT00_lSSN = goUtil.utGetECSCryptSetting("ECS", "WEB_SECURITY", "SSN")
            .RT01_sSubToCarrier = goUtil.IsNullIsVbNullString(RSACID.Fields("ClientCompanyDesc"))
            .RT02_sIBNumber = goUtil.IsNullIsVbNullString(RSAssgn.Fields("IBNUM"))
            .RT05_sLocation = goUtil.IsNullIsVbNullString(RSClientCoCat.Fields("SACity"))
            .RT05a_sState = goUtil.IsNullIsVbNullString(RSClientCoCat.Fields("SAState"))
            .RT06_dtDateClosed = "Null"
            .RT07_sAdjusterName = goUtil.IsNullIsVbNullString(RSACID.Fields("LFName"))
            .RT09_sSALN = goUtil.IsNullIsVbNullString(RSAssgn.Fields("CLIENTNUM"))
            .RT09a_sPolicyNo = goUtil.IsNullIsVbNullString(RSAssgn.Fields("PolicyNo"))
            .RT10_sInsuredName = goUtil.IsNullIsVbNullString(RSAssgn.Fields("Insured"))
            .RT11_sLossLocation = goUtil.IsNullIsVbNullString(RSAssgn.Fields("PAStreet")) & vbCrLf
            .RT11_sLossLocation = .RT11_sLossLocation & goUtil.IsNullIsVbNullString(RSAssgn.Fields("PACity")) & ", "
            .RT11_sLossLocation = .RT11_sLossLocation & goUtil.IsNullIsVbNullString(RSAssgn.Fields("PAState")) & " "
            .RT11_sLossLocation = .RT11_sLossLocation & Format(goUtil.IsNullIsVbNullString(RSAssgn.Fields("PAZIP")), "00000") & " - "
            .RT11_sLossLocation = .RT11_sLossLocation & Format(goUtil.IsNullIsVbNullString(RSAssgn.Fields("PAZIP4")), "0000")
            .RT12_dtDateOfLoss = IIf(goUtil.IsNullIsVbNullString(RSAssgn.Fields("LossDate")) = vbNullString, "Null", goUtil.IsNullIsVbNullString(RSAssgn.Fields("LossDate")))
            .RT13_cGrossLoss = "0.00"
            .RT14_cDepreciation = "0.00"
            .RT14a_sSupplement = lSupplement
            .RT14b_sRebilled = "0"
            .RT15_cDeductible = goUtil.IsNullIsVbNullString(RSAssgn.Fields("Deductible"))
            .RT15a_cLessExcessLimits = "0.00"
            .RT15b_sExcessLimDesc = vbNullString
            .RT15c_cLessMiscellaneous = "0.00"
            .RT15d_cMiscellaneousDesc = vbNullString
            .RT16_cNetClaim = "0.00"
            .RT17_cServiceFee = "0.00"
            .RT17a_cMiscServiceFee = "0.00"
            .RT18_sServiceFeeComment = vbNullString
            .RT18a_sMiscServiceFeeComment = vbNullString
            .RT25_cServiceFeeSubTotal = "0.00"
            .RT29a_sMiscExpenseFeeComment = vbNullString
            .RT29b_cMiscExpenseFee = "0.00"
            .RT30_cTotalExpenses = "0.00"
            .RT31_dTaxPercent = goUtil.IsNullIsVbNullString(RSFeeSchedule.Fields("TaxPercent"))
            .RT32_cTaxAmount = "0.00"
            .RT33_cTotalAdjustingFee = "0.00"
            .RT33a_sAccountCode = vbUnchecked
            .FeeScheduleID = sFeeScheduleID
            .Void = "False"
            .FeeByTime = GetSetting(App.EXEName, "FeeSchedule\" & sSection, "INIT_OPT_USE_SERVICE_FEE_BYTIME", "False")
            .UseActivityTime = GetSetting(App.EXEName, "FeeSchedule\" & sSection, "INIT_OPT_USE_ACTIVITY_TIME", "False")
            .DownLoadMe = "False"
            .UpLoadMe = "True"
            .AdminComments = vbNullString
            .DateLastUpdated = Now()
            .UpdateByUserID = goUtil.gsCurUsersID
        End With
    End If
    
    If pbInsertRecord Then
        If Not AddRTIB(MyudtRTIBItem) Then
            GoTo CLEAN_UP
        End If
    Else
        If Not EditRTIB(MyudtRTIBItem) Then
            GoTo CLEAN_UP
        End If
    End If
    
    Sleep 100
    InitRTIB = True
    
CLEAN_UP:
    Set RSCLIB = Nothing
    Set RSAssgn = Nothing
    Set RSACID = Nothing
    Set RSCatCode = Nothing
    Set RSClientCoCat = Nothing
    Set RSFeeSchedule = Nothing
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function InitRTIB"
End Function

Public Function EditBillingCount(pudtBillingCountItem As GuiBillingCountItem) As Boolean
    On Error GoTo EH
    Dim sSQL As String
    Dim oConn As ADODB.Connection
    
    sSQL = "UPDATE BillingCount Set "
    sSQL = sSQL & "[BillingCountID] = " & pudtBillingCountItem.BillingCountID & ", "
    sSQL = sSQL & "[AssignmentsID] = " & pudtBillingCountItem.AssignmentsID & ", "
    sSQL = sSQL & "[ID] = " & pudtBillingCountItem.ID & ", "
    sSQL = sSQL & "[IDAssignments] = " & pudtBillingCountItem.IDAssignments & ", "
    sSQL = sSQL & "[Rebill] = " & pudtBillingCountItem.Rebill & ", "
    sSQL = sSQL & "[Supplement] = " & pudtBillingCountItem.Supplement & ", "
    sSQL = sSQL & "[DownLoadMe] = " & pudtBillingCountItem.DownLoadMe & ", "
    sSQL = sSQL & "[UpLoadMe] = " & pudtBillingCountItem.UpLoadMe & ", "
    sSQL = sSQL & "[AdminComments] = '" & goUtil.utCleanSQLString(pudtBillingCountItem.AdminComments) & "', "
    sSQL = sSQL & "[DateLastUpdated] = #" & pudtBillingCountItem.DateLastUpdated & "#, "
    sSQL = sSQL & "[UpdateByUserID]  = " & pudtBillingCountItem.UpdateByUserID & " "
    sSQL = sSQL & "WHERE [IDAssignments] = " & pudtBillingCountItem.IDAssignments & " "
    sSQL = sSQL & "AND [ID] = " & pudtBillingCountItem.ID & " "
    
    Set oConn = New ADODB.Connection
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    oConn.Execute sSQL

    Sleep 100
    
    EditBillingCount = True
    
CLEAN_UP:
    Set oConn = Nothing
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function EditBillingCount"
End Function

Public Function AddBillingCount(pudtBillingCountItem As GuiBillingCountItem, psID As String) As Boolean
    On Error GoTo EH
    Dim sSQL As String
    Dim oConn As ADODB.Connection
    Dim sID As String
    
    sID = goUtil.GetAccessDBUID("ID", "BillingCount")
    
    With pudtBillingCountItem
        .BillingCountID = sID
        .AssignmentsID = msAssignmentsID
        .ID = sID
        .IDAssignments = msAssignmentsID 'not set here
    End With
    
    sSQL = "INSERT INTO BillingCount "
    sSQL = sSQL & "( "
    sSQL = sSQL & "[BillingCountID], "
    sSQL = sSQL & "[AssignmentsID], "
    sSQL = sSQL & "[ID], "
    sSQL = sSQL & "[IDAssignments], "
    sSQL = sSQL & "[Rebill], "
    sSQL = sSQL & "[Supplement], "
    sSQL = sSQL & "[DownLoadMe], "
    sSQL = sSQL & "[UpLoadMe], "
    sSQL = sSQL & "[AdminComments], "
    sSQL = sSQL & "[DateLastUpdated], "
    sSQL = sSQL & "[UpdateByUserID] "
    sSQL = sSQL & ") "
    sSQL = sSQL & "SELECT "
    sSQL = sSQL & pudtBillingCountItem.BillingCountID & " As [BillingCountID], "
    sSQL = sSQL & pudtBillingCountItem.AssignmentsID & " As [AssignmentsID], "
    sSQL = sSQL & pudtBillingCountItem.ID & " As [ID] , "
    sSQL = sSQL & pudtBillingCountItem.IDAssignments & " As [IDAssignments], "
    sSQL = sSQL & pudtBillingCountItem.Rebill & " As [Rebill], "
    sSQL = sSQL & pudtBillingCountItem.Supplement & " As [Supplement], "
    sSQL = sSQL & pudtBillingCountItem.DownLoadMe & " As [DownLoadMe], "
    sSQL = sSQL & pudtBillingCountItem.UpLoadMe & " As [UpLoadMe], "
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtBillingCountItem.AdminComments) & "'" & " As [AdminComments], "
    sSQL = sSQL & "#" & pudtBillingCountItem.DateLastUpdated & "#" & " As [DateLastUpdated], "
    sSQL = sSQL & pudtBillingCountItem.UpdateByUserID & " As [UpdateByUserID] "
    
    Set oConn = New ADODB.Connection
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    oConn.Execute sSQL
    
    psID = sID
    
    Sleep 100
    
    AddBillingCount = True
    
CLEAN_UP:
    Set oConn = Nothing
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function AddBillingCount"
End Function

'---------------------------------RTIB DB-------------------------------------------
Public Function EditRTIB(pudtGuiRTIBItem As GuiRTIBItem) As Boolean
    On Error GoTo EH
    Dim sSQL As String
    Dim oConn As ADODB.Connection
    Dim lRecordsAffected As Long
    
    sSQL = "UPDATE RTIB Set "
    sSQL = sSQL & "[AssignmentsID] = " & pudtGuiRTIBItem.AssignmentsID & ", "
    sSQL = sSQL & "[BillingCountID] = " & pudtGuiRTIBItem.BillingCountID & ", "
    sSQL = sSQL & "[IDAssignments] = " & pudtGuiRTIBItem.IDAssignments & ", "
    sSQL = sSQL & "[IDBillingCount] = " & pudtGuiRTIBItem.IDBillingCount & ", "
    sSQL = sSQL & "[RT00_lSSN] = " & pudtGuiRTIBItem.RT00_lSSN & ", "
    sSQL = sSQL & "[RT01_sSubToCarrier] = '" & goUtil.utCleanSQLString(pudtGuiRTIBItem.RT01_sSubToCarrier) & "', "
    sSQL = sSQL & "[RT02_sIBNumber] = '" & goUtil.utCleanSQLString(pudtGuiRTIBItem.RT02_sIBNumber) & "', "
    sSQL = sSQL & "[RT05_sLocation] = '" & goUtil.utCleanSQLString(pudtGuiRTIBItem.RT05_sLocation) & "', "
    sSQL = sSQL & "[RT05a_sState] = '" & goUtil.utCleanSQLString(pudtGuiRTIBItem.RT05a_sState) & "', "
    If StrComp(pudtGuiRTIBItem.RT06_dtDateClosed, "Null", vbTextCompare) = 0 Then
        sSQL = sSQL & "[RT06_dtDateClosed] = Null, "
    Else
        sSQL = sSQL & "[RT06_dtDateClosed] = #" & pudtGuiRTIBItem.RT06_dtDateClosed & "#, "
    End If
    sSQL = sSQL & "[RT07_sAdjusterName] = '" & goUtil.utCleanSQLString(pudtGuiRTIBItem.RT07_sAdjusterName) & "', "
    sSQL = sSQL & "[RT09_sSALN] = '" & goUtil.utCleanSQLString(pudtGuiRTIBItem.RT09_sSALN) & "', "
    sSQL = sSQL & "[RT09a_sPolicyNo] = '" & goUtil.utCleanSQLString(pudtGuiRTIBItem.RT09a_sPolicyNo) & "', "
    sSQL = sSQL & "[RT10_sInsuredName] = '" & goUtil.utCleanSQLString(pudtGuiRTIBItem.RT10_sInsuredName) & "', "
    sSQL = sSQL & "[RT11_sLossLocation] = '" & goUtil.utCleanSQLString(pudtGuiRTIBItem.RT11_sLossLocation) & "', "
    If StrComp(pudtGuiRTIBItem.RT12_dtDateOfLoss, "Null", vbTextCompare) = 0 Then
        sSQL = sSQL & "[RT12_dtDateOfLoss] = Null, "
    Else
        sSQL = sSQL & "[RT12_dtDateOfLoss] = #" & pudtGuiRTIBItem.RT12_dtDateOfLoss & "#, "
    End If
    sSQL = sSQL & "[RT13_cGrossLoss] = " & CCur(pudtGuiRTIBItem.RT13_cGrossLoss) & ", "
    sSQL = sSQL & "[RT14_cDepreciation] = " & CCur(pudtGuiRTIBItem.RT14_cDepreciation) & ", "
    sSQL = sSQL & "[RT14a_sSupplement] = '" & goUtil.utCleanSQLString(pudtGuiRTIBItem.RT14a_sSupplement) & "', "
    sSQL = sSQL & "[RT14b_sRebilled] = '" & goUtil.utCleanSQLString(pudtGuiRTIBItem.RT14b_sRebilled) & "', "
    sSQL = sSQL & "[RT15_cDeductible] = " & CCur(pudtGuiRTIBItem.RT15_cDeductible) & ", "
    sSQL = sSQL & "[RT15a_cLessExcessLimits] = " & CCur(pudtGuiRTIBItem.RT15a_cLessExcessLimits) & ", "
    sSQL = sSQL & "[RT15b_sExcessLimDesc] = '" & goUtil.utCleanSQLString(pudtGuiRTIBItem.RT15b_sExcessLimDesc) & "', "
    sSQL = sSQL & "[RT15c_cLessMiscellaneous] = " & CCur(pudtGuiRTIBItem.RT15c_cLessMiscellaneous) & ", "
    sSQL = sSQL & "[RT15d_cMiscellaneousDesc] = '" & goUtil.utCleanSQLString(pudtGuiRTIBItem.RT15d_cMiscellaneousDesc) & "', "
    sSQL = sSQL & "[RT16_cNetClaim] = " & CCur(pudtGuiRTIBItem.RT16_cNetClaim) & ", "
    sSQL = sSQL & "[RT17_cServiceFee] = " & CCur(pudtGuiRTIBItem.RT17_cServiceFee) & ", "
    sSQL = sSQL & "[RT17a_cMiscServiceFee] = " & CCur(pudtGuiRTIBItem.RT17a_cMiscServiceFee) & ", "
    sSQL = sSQL & "[RT18_sServiceFeeComment] = '" & goUtil.utCleanSQLString(pudtGuiRTIBItem.RT18_sServiceFeeComment) & "', "
    sSQL = sSQL & "[RT18a_sMiscServiceFeeComment] = '" & goUtil.utCleanSQLString(pudtGuiRTIBItem.RT18a_sMiscServiceFeeComment) & "', "
    sSQL = sSQL & "[RT25_cServiceFeeSubTotal] = " & CCur(pudtGuiRTIBItem.RT25_cServiceFeeSubTotal) & ", "
    sSQL = sSQL & "[RT29a_sMiscExpenseFeeComment] = '" & goUtil.utCleanSQLString(pudtGuiRTIBItem.RT29a_sMiscExpenseFeeComment) & "', "
    sSQL = sSQL & "[RT29b_cMiscExpenseFee] = " & CCur(pudtGuiRTIBItem.RT29b_cMiscExpenseFee) & ", "
    sSQL = sSQL & "[RT30_cTotalExpenses] = " & CCur(pudtGuiRTIBItem.RT30_cTotalExpenses) & ", "
    sSQL = sSQL & "[RT31_dTaxPercent] = " & pudtGuiRTIBItem.RT31_dTaxPercent & ", "
    sSQL = sSQL & "[RT32_cTaxAmount] = " & CCur(pudtGuiRTIBItem.RT32_cTaxAmount) & ", "
    sSQL = sSQL & "[RT33_cTotalAdjustingFee] = " & CCur(pudtGuiRTIBItem.RT33_cTotalAdjustingFee) & ", "
    sSQL = sSQL & "[RT33a_sAccountCode] = '" & goUtil.utCleanSQLString(pudtGuiRTIBItem.RT33a_sAccountCode) & "', "
    sSQL = sSQL & "[FeeScheduleID] = " & pudtGuiRTIBItem.FeeScheduleID & ", "
    sSQL = sSQL & "[Void] = " & pudtGuiRTIBItem.Void & ", "
    sSQL = sSQL & "[FeeByTime] = " & pudtGuiRTIBItem.FeeByTime & ", "
    sSQL = sSQL & "[UseActivityTime] = " & pudtGuiRTIBItem.UseActivityTime & ", "
    sSQL = sSQL & "[DownLoadMe] = " & pudtGuiRTIBItem.DownLoadMe & ", "
    sSQL = sSQL & "[UpLoadMe] = " & pudtGuiRTIBItem.UpLoadMe & ", "
    sSQL = sSQL & "[Comments] = '" & goUtil.utCleanSQLString(pudtGuiRTIBItem.Comments) & "', "
    sSQL = sSQL & "[AdminComments] = '" & goUtil.utCleanSQLString(pudtGuiRTIBItem.AdminComments) & "', "
    sSQL = sSQL & "[DateLastUpdated] = #" & pudtGuiRTIBItem.DateLastUpdated & "#, "
    sSQL = sSQL & "[UpdateByUserID] = " & pudtGuiRTIBItem.UpdateByUserID & " "
    sSQL = sSQL & "WHERE AssignmentsID = " & msAssignmentsID & " "
    
    Set oConn = New ADODB.Connection
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    oConn.Execute sSQL, lRecordsAffected

    Sleep 200
    
    EditRTIB = CBool(lRecordsAffected)
    
CLEAN_UP:
    Set oConn = Nothing
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function EditRTIB"
End Function

Public Function AddRTIB(pudtGuiRTIBItem As GuiRTIBItem) As Boolean
    On Error GoTo EH
    Dim sSQL As String
    Dim oConn As ADODB.Connection
    Dim sID As String
    Dim lRecordsAffected As Long
    
    sSQL = "INSERT INTO RTIB "
    sSQL = sSQL & "( "
    sSQL = sSQL & "[AssignmentsID], "
    sSQL = sSQL & "[BillingCountID], "
    sSQL = sSQL & "[IDAssignments], "
    sSQL = sSQL & "[IDBillingCount], "
    sSQL = sSQL & "[RT00_lSSN], "
    sSQL = sSQL & "[RT01_sSubToCarrier], "
    sSQL = sSQL & "[RT02_sIBNumber], "
    sSQL = sSQL & "[RT05_sLocation], "
    sSQL = sSQL & "[RT05a_sState], "
    sSQL = sSQL & "[RT06_dtDateClosed], "
    sSQL = sSQL & "[RT07_sAdjusterName], "
    sSQL = sSQL & "[RT09_sSALN], "
    sSQL = sSQL & "[RT09a_sPolicyNo], "
    sSQL = sSQL & "[RT10_sInsuredName], "
    sSQL = sSQL & "[RT11_sLossLocation], "
    sSQL = sSQL & "[RT12_dtDateOfLoss], "
    sSQL = sSQL & "[RT13_cGrossLoss], "
    sSQL = sSQL & "[RT14_cDepreciation], "
    sSQL = sSQL & "[RT14a_sSupplement], "
    sSQL = sSQL & "[RT14b_sRebilled], "
    sSQL = sSQL & "[RT15_cDeductible], "
    sSQL = sSQL & "[RT15a_cLessExcessLimits], "
    sSQL = sSQL & "[RT15b_sExcessLimDesc], "
    sSQL = sSQL & "[RT15c_cLessMiscellaneous], "
    sSQL = sSQL & "[RT15d_cMiscellaneousDesc], "
    sSQL = sSQL & "[RT16_cNetClaim], "
    sSQL = sSQL & "[RT17_cServiceFee], "
    sSQL = sSQL & "[RT17a_cMiscServiceFee], "
    sSQL = sSQL & "[RT18_sServiceFeeComment], "
    sSQL = sSQL & "[RT18a_sMiscServiceFeeComment], "
    sSQL = sSQL & "[RT25_cServiceFeeSubTotal], "
    sSQL = sSQL & "[RT29a_sMiscExpenseFeeComment], "
    sSQL = sSQL & "[RT29b_cMiscExpenseFee], "
    sSQL = sSQL & "[RT30_cTotalExpenses], "
    sSQL = sSQL & "[RT31_dTaxPercent], "
    sSQL = sSQL & "[RT32_cTaxAmount], "
    sSQL = sSQL & "[RT33_cTotalAdjustingFee], "
    sSQL = sSQL & "[RT33a_sAccountCode], "
    sSQL = sSQL & "[FeeScheduleID], "
    sSQL = sSQL & "[Void], "
    sSQL = sSQL & "[FeeByTime], "
    sSQL = sSQL & "[UseActivityTime], "
    sSQL = sSQL & "[DownLoadMe], "
    sSQL = sSQL & "[UpLoadMe], "
    sSQL = sSQL & "[Comments], "
    sSQL = sSQL & "[AdminComments], "
    sSQL = sSQL & "[DateLastUpdated], "
    sSQL = sSQL & "[UpdateByUserID] "
    sSQL = sSQL & ") "
    sSQL = sSQL & "SELECT "
    sSQL = sSQL & pudtGuiRTIBItem.AssignmentsID & " As [AssignmentsID], "
    sSQL = sSQL & pudtGuiRTIBItem.BillingCountID & " As [BillingCountID] , "
    sSQL = sSQL & pudtGuiRTIBItem.IDAssignments & " As [IDAssignments], "
    sSQL = sSQL & pudtGuiRTIBItem.IDBillingCount & " As [IDBillingCount] , "
    sSQL = sSQL & pudtGuiRTIBItem.RT00_lSSN & " As [RT00_lSSN], "
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtGuiRTIBItem.RT01_sSubToCarrier) & "'" & " As [RT01_sSubToCarrier], "
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtGuiRTIBItem.RT02_sIBNumber) & "'" & " As [RT02_sIBNumber], "
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtGuiRTIBItem.RT05_sLocation) & "'" & " As [RT05_sLocation], "
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtGuiRTIBItem.RT05a_sState) & "'" & " As [RT05a_sState], "
    If StrComp(pudtGuiRTIBItem.RT06_dtDateClosed, "Null", vbTextCompare) = 0 Then
        sSQL = sSQL & "Null As [RT06_dtDateClosed], "
    Else
        sSQL = sSQL & "#" & pudtGuiRTIBItem.RT06_dtDateClosed & "#" & " As [RT06_dtDateClosed], "
    End If
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtGuiRTIBItem.RT07_sAdjusterName) & "'" & " As [RT07_sAdjusterName], "
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtGuiRTIBItem.RT09_sSALN) & "'" & " As [RT09_sSALN], "
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtGuiRTIBItem.RT09a_sPolicyNo) & "'" & " As [RT09a_sPolicyNo], "
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtGuiRTIBItem.RT10_sInsuredName) & "'" & " As [RT10_sInsuredName], "
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtGuiRTIBItem.RT11_sLossLocation) & "'" & " As [RT11_sLossLocation], "
    If StrComp(pudtGuiRTIBItem.RT12_dtDateOfLoss, "Null", vbTextCompare) = 0 Then
        sSQL = sSQL & "Null As [RT12_dtDateOfLoss], "
    Else
        sSQL = sSQL & "#" & pudtGuiRTIBItem.RT12_dtDateOfLoss & "#" & " As [RT12_dtDateOfLoss], "
    End If
    sSQL = sSQL & CCur(pudtGuiRTIBItem.RT13_cGrossLoss) & " As [RT13_cGrossLoss], "
    sSQL = sSQL & CCur(pudtGuiRTIBItem.RT14_cDepreciation) & " As [RT14_cDepreciation], "
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtGuiRTIBItem.RT14a_sSupplement) & "'" & " As [RT14a_sSupplement], "
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtGuiRTIBItem.RT14b_sRebilled) & "'" & " As [RT14b_sRebilled], "
    sSQL = sSQL & CCur(pudtGuiRTIBItem.RT15_cDeductible) & " As [RT15_cDeductible], "
    sSQL = sSQL & CCur(pudtGuiRTIBItem.RT15a_cLessExcessLimits) & " As [RT15a_cLessExcessLimits], "
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtGuiRTIBItem.RT15b_sExcessLimDesc) & "'" & " As [RT15b_sExcessLimDesc], "
    sSQL = sSQL & CCur(pudtGuiRTIBItem.RT15c_cLessMiscellaneous) & " As [RT15c_cLessMiscellaneous], "
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtGuiRTIBItem.RT15d_cMiscellaneousDesc) & "'" & " As [RT15d_cMiscellaneousDesc], "
    sSQL = sSQL & CCur(pudtGuiRTIBItem.RT16_cNetClaim) & " As [RT16_cNetClaim], "
    sSQL = sSQL & CCur(pudtGuiRTIBItem.RT17_cServiceFee) & " As [RT17_cServiceFee], "
    sSQL = sSQL & CCur(pudtGuiRTIBItem.RT17a_cMiscServiceFee) & " As [RT17a_cMiscServiceFee], "
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtGuiRTIBItem.RT18_sServiceFeeComment) & "'" & " As [RT18_sServiceFeeComment], "
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtGuiRTIBItem.RT18a_sMiscServiceFeeComment) & "'" & " As [RT18a_sMiscServiceFeeComment], "
    sSQL = sSQL & CCur(pudtGuiRTIBItem.RT25_cServiceFeeSubTotal) & " As [RT25_cServiceFeeSubTotal], "
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtGuiRTIBItem.RT29a_sMiscExpenseFeeComment) & "'" & " As [RT29a_sMiscExpenseFeeComment], "
    sSQL = sSQL & CCur(pudtGuiRTIBItem.RT29b_cMiscExpenseFee) & " As [RT29b_cMiscExpenseFee], "
    sSQL = sSQL & CCur(pudtGuiRTIBItem.RT30_cTotalExpenses) & " As [RT30_cTotalExpenses], "
    sSQL = sSQL & CCur(pudtGuiRTIBItem.RT31_dTaxPercent) & " As [RT31_dTaxPercent], "
    sSQL = sSQL & CCur(pudtGuiRTIBItem.RT32_cTaxAmount) & " As [RT32_cTaxAmount], "
    sSQL = sSQL & CCur(pudtGuiRTIBItem.RT33_cTotalAdjustingFee) & " As [RT33_cTotalAdjustingFee], "
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtGuiRTIBItem.RT33a_sAccountCode) & "'" & " As [RT33a_sAccountCode], "
    sSQL = sSQL & pudtGuiRTIBItem.FeeScheduleID & " As [FeeScheduleID], "
    sSQL = sSQL & pudtGuiRTIBItem.Void & " As [Void], "
    sSQL = sSQL & pudtGuiRTIBItem.FeeByTime & " As [FeeByTime], "
    sSQL = sSQL & pudtGuiRTIBItem.UseActivityTime & " As [UseActivityTime], "
    sSQL = sSQL & pudtGuiRTIBItem.DownLoadMe & " As [DownLoadMe], "
    sSQL = sSQL & pudtGuiRTIBItem.UpLoadMe & " As [UpLoadMe], "
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtGuiRTIBItem.Comments) & "'" & " As [Comments], "
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtGuiRTIBItem.AdminComments) & "'" & " As [AdminComments], "
    sSQL = sSQL & "#" & pudtGuiRTIBItem.DateLastUpdated & "#" & " As [DateLastUpdated], "
    sSQL = sSQL & pudtGuiRTIBItem.UpdateByUserID & " As [UpdateByUserID] "
    
    Set oConn = New ADODB.Connection
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    oConn.Execute sSQL, lRecordsAffected
        
    Sleep 200
    
    AddRTIB = CBool(lRecordsAffected)
    
CLEAN_UP:
    Set oConn = Nothing
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function AddRTIB"
End Function


Public Function EditRTIBFee(pudtRTIBFeeItem As GuiRTIBFeeItem) As Boolean
    On Error GoTo EH
    Dim sSQL As String
    Dim oConn As ADODB.Connection
    Dim lRecordsAffected As Long
    
    sSQL = "UPDATE RTIBFee Set "
    sSQL = sSQL & "[RTIBFeeID] = " & pudtRTIBFeeItem.RTIBFeeID & ", "
    sSQL = sSQL & "[AssignmentsID] = " & pudtRTIBFeeItem.AssignmentsID & ", "
    sSQL = sSQL & "[ID] = " & pudtRTIBFeeItem.ID & ", "
    sSQL = sSQL & "[IDAssignments] = " & pudtRTIBFeeItem.IDAssignments & ", "
    sSQL = sSQL & "[FeeScheduleFeeTypesID] = " & pudtRTIBFeeItem.FeeScheduleFeeTypesID & ", "
    sSQL = sSQL & "[NumberOfItems] = " & pudtRTIBFeeItem.NumberOfItems & ", "
    sSQL = sSQL & "[Amount] = " & CCur(pudtRTIBFeeItem.Amount) & ", "
    sSQL = sSQL & "[Comment]  = '" & goUtil.utCleanSQLString(pudtRTIBFeeItem.Comment) & "', "
    sSQL = sSQL & "[DownLoadMe] = " & pudtRTIBFeeItem.DownLoadMe & ", "
    sSQL = sSQL & "[UpLoadMe] = " & pudtRTIBFeeItem.UpLoadMe & ", "
    sSQL = sSQL & "[AdminComments] = '" & goUtil.utCleanSQLString(pudtRTIBFeeItem.AdminComments) & "', "
    sSQL = sSQL & "[DateLastUpdated] = #" & pudtRTIBFeeItem.DateLastUpdated & "#, "
    sSQL = sSQL & "[UpdateByUserID] = " & pudtRTIBFeeItem.UpdateByUserID & " "
    sSQL = sSQL & "WHERE AssignmentsID = " & pudtRTIBFeeItem.AssignmentsID & " "
    If pudtRTIBFeeItem.ID = "[ID]" Then
        sSQL = sSQL & "AND FeeScheduleFeeTypesID = " & pudtRTIBFeeItem.FeeScheduleFeeTypesID & " "
    Else
        sSQL = sSQL & "AND ID = " & pudtRTIBFeeItem.ID & " "
    End If
    
    Set oConn = New ADODB.Connection
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    oConn.Execute sSQL, lRecordsAffected

    Sleep 10
    
    EditRTIBFee = CBool(lRecordsAffected)
    
CLEAN_UP:
    Set oConn = Nothing
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function EditRTIBFee"
End Function

Public Function AddRTIBFee(pudtRTIBFeeItem As GuiRTIBFeeItem, psID As String) As Boolean
    On Error GoTo EH
    Dim sSQL As String
    Dim oConn As ADODB.Connection
    Dim sID As String
    Dim lRecordsAffected As Long
    
    sID = goUtil.GetAccessDBUID("ID", "RTIBFee")
    
    With pudtRTIBFeeItem
        .RTIBFeeID = sID
        .AssignmentsID = msAssignmentsID
        .ID = sID
        .IDAssignments = msAssignmentsID 'not set here
    End With
    
    sSQL = "INSERT INTO RTIBFee "
    sSQL = sSQL & "( "
    sSQL = sSQL & "[RTIBFeeID], "
    sSQL = sSQL & "[AssignmentsID], "
    sSQL = sSQL & "[ID], "
    sSQL = sSQL & "[IDAssignments], "
    sSQL = sSQL & "[FeeScheduleFeeTypesID], "
    sSQL = sSQL & "[NumberOfItems], "
    sSQL = sSQL & "[Amount], "
    sSQL = sSQL & "[Comment], "
    sSQL = sSQL & "[DownLoadMe], "
    sSQL = sSQL & "[UpLoadMe], "
    sSQL = sSQL & "[AdminComments], "
    sSQL = sSQL & "[DateLastUpdated], "
    sSQL = sSQL & "[UpdateByUserID] "
    sSQL = sSQL & ") "
    sSQL = sSQL & "SELECT "
    sSQL = sSQL & pudtRTIBFeeItem.RTIBFeeID & " As [RTIBFeeID], "
    sSQL = sSQL & pudtRTIBFeeItem.AssignmentsID & " As [AssignmentsID], "
    sSQL = sSQL & pudtRTIBFeeItem.ID & " As [ID] , "
    sSQL = sSQL & pudtRTIBFeeItem.IDAssignments & " As [IDAssignments], "
    sSQL = sSQL & pudtRTIBFeeItem.FeeScheduleFeeTypesID & " As [FeeScheduleFeeTypesID], "
    sSQL = sSQL & pudtRTIBFeeItem.NumberOfItems & " As [NumberOfItems], "
    sSQL = sSQL & CCur(pudtRTIBFeeItem.Amount) & " As [Amount], "
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtRTIBFeeItem.Comment) & "'" & " As [Comment], "
    sSQL = sSQL & pudtRTIBFeeItem.DownLoadMe & " As [DownLoadMe], "
    sSQL = sSQL & pudtRTIBFeeItem.UpLoadMe & " As [UpLoadMe], "
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtRTIBFeeItem.AdminComments) & "'" & " As [AdminComments], "
    sSQL = sSQL & "#" & pudtRTIBFeeItem.DateLastUpdated & "#" & " As [DateLastUpdated], "
    sSQL = sSQL & pudtRTIBFeeItem.UpdateByUserID & " As [UpdateByUserID] "
    
    Set oConn = New ADODB.Connection
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    oConn.Execute sSQL, lRecordsAffected
    
    psID = sID
    
    Sleep 10
    
    AddRTIBFee = CBool(lRecordsAffected)
    
CLEAN_UP:
    Set oConn = Nothing
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function AddRTIBFee"
End Function
'--------------------------------END RTIB DB---------------------------------------


'---------------------------------IB DB-------------------------------------------
Public Function EditIB(pudtGuiIBItem As GuiIBItem, psID As String) As Boolean
    On Error GoTo EH
    Dim sSQL As String
    Dim oConn As ADODB.Connection
    Dim lRecordsAffected As Long
    Dim RS As ADODB.Recordset
    
    sSQL = "UPDATE IB Set "
    sSQL = sSQL & "[IBID] = " & pudtGuiIBItem.IBID & ", "
    sSQL = sSQL & "[AssignmentsID] = " & pudtGuiIBItem.AssignmentsID & ", "
    sSQL = sSQL & "[BillingCountID] = " & pudtGuiIBItem.BillingCountID & ", "
    sSQL = sSQL & "[ID] = " & pudtGuiIBItem.ID & ", "
    sSQL = sSQL & "[IDAssignments] = " & pudtGuiIBItem.IDAssignments & ", "
    sSQL = sSQL & "[IDBillingCount] = " & pudtGuiIBItem.IDBillingCount & ", "
    sSQL = sSQL & "[IB00_lSSN] = " & pudtGuiIBItem.IB00_lssn & ", "
    sSQL = sSQL & "[IB01_sSubToCarrier] = '" & goUtil.utCleanSQLString(pudtGuiIBItem.IB01_sSubToCarrier) & "', "
    sSQL = sSQL & "[IB02_sIBNumber] = '" & goUtil.utCleanSQLString(pudtGuiIBItem.IB02_sIBNumber) & "', "
    sSQL = sSQL & "[IB05_sLocation] = '" & goUtil.utCleanSQLString(pudtGuiIBItem.IB05_sLocation) & "', "
    sSQL = sSQL & "[IB05a_sState] = '" & goUtil.utCleanSQLString(pudtGuiIBItem.IB05a_sState) & "', "
    If StrComp(pudtGuiIBItem.IB06_dtDateClosed, "Null", vbTextCompare) = 0 Then
        sSQL = sSQL & "[IB06_dtDateClosed] = Null, "
    Else
        sSQL = sSQL & "[IB06_dtDateClosed] = #" & pudtGuiIBItem.IB06_dtDateClosed & "#, "
    End If
    sSQL = sSQL & "[IB07_sAdjusterName] = '" & goUtil.utCleanSQLString(pudtGuiIBItem.IB07_sAdjusterName) & "', "
    sSQL = sSQL & "[IB09_sSALN] = '" & goUtil.utCleanSQLString(pudtGuiIBItem.IB09_sSALN) & "', "
    sSQL = sSQL & "[IB09a_sPolicyNo] = '" & goUtil.utCleanSQLString(pudtGuiIBItem.IB09a_sPolicyNo) & "', "
    sSQL = sSQL & "[IB10_sInsuredName] = '" & goUtil.utCleanSQLString(pudtGuiIBItem.IB10_sInsuredName) & "', "
    sSQL = sSQL & "[IB11_sLossLocation] = '" & goUtil.utCleanSQLString(pudtGuiIBItem.IB11_sLossLocation) & "', "
    If StrComp(pudtGuiIBItem.IB12_dtDateOfLoss, "Null", vbTextCompare) = 0 Then
        sSQL = sSQL & "[IB12_dtDateOfLoss] = Null, "
    Else
        sSQL = sSQL & "[IB12_dtDateOfLoss] = #" & pudtGuiIBItem.IB12_dtDateOfLoss & "#, "
    End If
    sSQL = sSQL & "[IB13_cGrossLoss] = " & CCur(pudtGuiIBItem.IB13_cGrossLoss) & ", "
    sSQL = sSQL & "[IB14_cDepreciation] = " & CCur(pudtGuiIBItem.IB14_cDepreciation) & ", "
    sSQL = sSQL & "[IB14a_sSupplement] = '" & goUtil.utCleanSQLString(pudtGuiIBItem.IB14a_sSupplement) & "', "
    sSQL = sSQL & "[IB14b_sRebilled] = '" & goUtil.utCleanSQLString(pudtGuiIBItem.IB14b_sRebilled) & "', "
    sSQL = sSQL & "[IB15_cDeductible] = " & CCur(pudtGuiIBItem.IB15_cDeductible) & ", "
    sSQL = sSQL & "[IB15a_cLessExcessLimits] = " & CCur(pudtGuiIBItem.IB15a_cLessExcessLimits) & ", "
    sSQL = sSQL & "[IB15b_sExcessLimDesc] = '" & goUtil.utCleanSQLString(pudtGuiIBItem.IB15b_sExcessLimDesc) & "', "
    sSQL = sSQL & "[IB15c_cLessMiscellaneous] = " & CCur(pudtGuiIBItem.IB15c_cLessMiscellaneous) & ", "
    sSQL = sSQL & "[IB15d_cMiscellaneousDesc] = '" & goUtil.utCleanSQLString(pudtGuiIBItem.IB15d_cMiscellaneousDesc) & "', "
    sSQL = sSQL & "[IB16_cNetClaim] = " & CCur(pudtGuiIBItem.IB16_cNetClaim) & ", "
    sSQL = sSQL & "[IB17_cServiceFee] = " & CCur(pudtGuiIBItem.IB17_cServiceFee) & ", "
    sSQL = sSQL & "[IB17a_cMiscServiceFee] = " & CCur(pudtGuiIBItem.IB17a_cMiscServiceFee) & ", "
    sSQL = sSQL & "[IB18_sServiceFeeComment] = '" & goUtil.utCleanSQLString(pudtGuiIBItem.IB18_sServiceFeeComment) & "', "
    sSQL = sSQL & "[IB18a_sMiscServiceFeeComment] = '" & goUtil.utCleanSQLString(pudtGuiIBItem.IB18a_sMiscServiceFeeComment) & "', "
    sSQL = sSQL & "[IB25_cServiceFeeSubTotal] = " & CCur(pudtGuiIBItem.IB25_cServiceFeeSubTotal) & ", "
    sSQL = sSQL & "[IB29a_sMiscExpenseFeeComment] = '" & goUtil.utCleanSQLString(pudtGuiIBItem.IB29a_sMiscExpenseFeeComment) & "', "
    sSQL = sSQL & "[IB29b_cMiscExpenseFee] = " & CCur(pudtGuiIBItem.IB29b_cMiscExpenseFee) & ", "
    sSQL = sSQL & "[IB30_cTotalExpenses] = " & CCur(pudtGuiIBItem.IB30_cTotalExpenses) & ", "
    sSQL = sSQL & "[IB31_dTaxPercent] = " & pudtGuiIBItem.IB31_dTaxPercent & ", "
    sSQL = sSQL & "[IB32_cTaxAmount] = " & CCur(pudtGuiIBItem.IB32_cTaxAmount) & ", "
    sSQL = sSQL & "[IB33_cTotalAdjustingFee] = " & CCur(pudtGuiIBItem.IB33_cTotalAdjustingFee) & ", "
    sSQL = sSQL & "[IB33a_sAccountCode] = '" & goUtil.utCleanSQLString(pudtGuiIBItem.IB33a_sAccountCode) & "', "
    sSQL = sSQL & "[FeeScheduleID] = " & pudtGuiIBItem.FeeScheduleID & ", "
    sSQL = sSQL & "[Void] = " & pudtGuiIBItem.Void & ", "
    sSQL = sSQL & "[FeeByTime] = " & pudtGuiIBItem.FeeByTime & ", "
    sSQL = sSQL & "[UseActivityTime] = " & pudtGuiIBItem.UseActivityTime & ", "
    sSQL = sSQL & "[DownLoadMe] = " & pudtGuiIBItem.DownLoadMe & ", "
    sSQL = sSQL & "[UpLoadMe] = " & pudtGuiIBItem.UpLoadMe & ", "
    sSQL = sSQL & "[Comments] = '" & goUtil.utCleanSQLString(pudtGuiIBItem.Comments) & "', "
    sSQL = sSQL & "[AdminComments] = '" & goUtil.utCleanSQLString(pudtGuiIBItem.AdminComments) & "', "
    sSQL = sSQL & "[DateLastUpdated] = #" & pudtGuiIBItem.DateLastUpdated & "#, "
    sSQL = sSQL & "[UpdateByUserID] = " & pudtGuiIBItem.UpdateByUserID & " "
    sSQL = sSQL & "WHERE AssignmentsID = " & msAssignmentsID & " "
    sSQL = sSQL & "AND [IDBillingCount] = " & pudtGuiIBItem.IDBillingCount & " "
    Set oConn = New ADODB.Connection
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    oConn.Execute sSQL, lRecordsAffected

    Sleep 200
    
    If CBool(lRecordsAffected) Then
        sSQL = "SELECT [ID] "
        sSQL = sSQL & "FROM IB "
        sSQL = sSQL & "WHERE AssignmentsID = " & msAssignmentsID & " "
        sSQL = sSQL & "AND [IDBillingCount] = " & pudtGuiIBItem.IDBillingCount & " "
        
        Set RS = New ADODB.Recordset
        
        'Use Disconnected Record Set on asUseClient Cusor ONLY !
        RS.CursorLocation = adUseClient
        RS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
        Set RS.ActiveConnection = Nothing
        If RS.RecordCount = 1 Then
            RS.MoveFirst
            psID = goUtil.IsNullIsVbNullString(RS.Fields("ID"))
            EditIB = True
        Else
            EditIB = False
        End If
    End If

CLEAN_UP:
    Set oConn = Nothing
    Set RS = Nothing
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function EditIB"
End Function

Public Function AddIB(pudtGuiIBItem As GuiIBItem, psID As String) As Boolean
    On Error GoTo EH
    Dim sSQL As String
    Dim oConn As ADODB.Connection
    Dim sID As String
    Dim lRecordsAffected As Long
    
    sID = goUtil.GetAccessDBUID("ID", "IB")
    
    With pudtGuiIBItem
        .IBID = sID
        .ID = sID
    End With
    
    sSQL = "INSERT INTO IB "
    sSQL = sSQL & "( "
    sSQL = sSQL & "[IBID], "
    sSQL = sSQL & "[AssignmentsID], "
    sSQL = sSQL & "[BillingCountID], "
    sSQL = sSQL & "[ID], "
    sSQL = sSQL & "[IDAssignments], "
    sSQL = sSQL & "[IDBillingCount], "
    sSQL = sSQL & "[IB00_lSSN], "
    sSQL = sSQL & "[IB01_sSubToCarrier], "
    sSQL = sSQL & "[IB02_sIBNumber], "
    sSQL = sSQL & "[IB05_sLocation], "
    sSQL = sSQL & "[IB05a_sState], "
    sSQL = sSQL & "[IB06_dtDateClosed], "
    sSQL = sSQL & "[IB07_sAdjusterName], "
    sSQL = sSQL & "[IB09_sSALN], "
    sSQL = sSQL & "[IB09a_sPolicyNo], "
    sSQL = sSQL & "[IB10_sInsuredName], "
    sSQL = sSQL & "[IB11_sLossLocation], "
    sSQL = sSQL & "[IB12_dtDateOfLoss], "
    sSQL = sSQL & "[IB13_cGrossLoss], "
    sSQL = sSQL & "[IB14_cDepreciation], "
    sSQL = sSQL & "[IB14a_sSupplement], "
    sSQL = sSQL & "[IB14b_sRebilled], "
    sSQL = sSQL & "[IB15_cDeductible], "
    sSQL = sSQL & "[IB15a_cLessExcessLimits], "
    sSQL = sSQL & "[IB15b_sExcessLimDesc], "
    sSQL = sSQL & "[IB15c_cLessMiscellaneous], "
    sSQL = sSQL & "[IB15d_cMiscellaneousDesc], "
    sSQL = sSQL & "[IB16_cNetClaim], "
    sSQL = sSQL & "[IB17_cServiceFee], "
    sSQL = sSQL & "[IB17a_cMiscServiceFee], "
    sSQL = sSQL & "[IB18_sServiceFeeComment], "
    sSQL = sSQL & "[IB18a_sMiscServiceFeeComment], "
    sSQL = sSQL & "[IB25_cServiceFeeSubTotal], "
    sSQL = sSQL & "[IB29a_sMiscExpenseFeeComment], "
    sSQL = sSQL & "[IB29b_cMiscExpenseFee], "
    sSQL = sSQL & "[IB30_cTotalExpenses], "
    sSQL = sSQL & "[IB31_dTaxPercent], "
    sSQL = sSQL & "[IB32_cTaxAmount], "
    sSQL = sSQL & "[IB33_cTotalAdjustingFee], "
    sSQL = sSQL & "[IB33a_sAccountCode], "
    sSQL = sSQL & "[FeeScheduleID], "
    sSQL = sSQL & "[Void], "
    sSQL = sSQL & "[FeeByTime], "
    sSQL = sSQL & "[UseActivityTime], "
    sSQL = sSQL & "[DownLoadMe], "
    sSQL = sSQL & "[UpLoadMe], "
    sSQL = sSQL & "[Comments], "
    sSQL = sSQL & "[AdminComments], "
    sSQL = sSQL & "[DateLastUpdated], "
    sSQL = sSQL & "[UpdateByUserID] "
    sSQL = sSQL & ") "
    sSQL = sSQL & "SELECT "
    sSQL = sSQL & pudtGuiIBItem.IBID & " As [IBID], "
    sSQL = sSQL & pudtGuiIBItem.AssignmentsID & " As [AssignmentsID], "
    sSQL = sSQL & pudtGuiIBItem.BillingCountID & " As [BillingCountID] , "
    sSQL = sSQL & pudtGuiIBItem.ID & " As [ID] , "
    sSQL = sSQL & pudtGuiIBItem.IDAssignments & " As [IDAssignments], "
    sSQL = sSQL & pudtGuiIBItem.IDBillingCount & " As [IDBillingCount] , "
    sSQL = sSQL & pudtGuiIBItem.IB00_lssn & " As [IB00_lSSN], "
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtGuiIBItem.IB01_sSubToCarrier) & "'" & " As [IB01_sSubToCarrier], "
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtGuiIBItem.IB02_sIBNumber) & "'" & " As [IB02_sIBNumber], "
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtGuiIBItem.IB05_sLocation) & "'" & " As [IB05_sLocation], "
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtGuiIBItem.IB05a_sState) & "'" & " As [IB05a_sState], "
    If StrComp(pudtGuiIBItem.IB06_dtDateClosed, "Null", vbTextCompare) = 0 Then
        sSQL = sSQL & "Null As [IB06_dtDateClosed], "
    Else
        sSQL = sSQL & "#" & pudtGuiIBItem.IB06_dtDateClosed & "#" & " As [IB06_dtDateClosed], "
    End If
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtGuiIBItem.IB07_sAdjusterName) & "'" & " As [IB07_sAdjusterName], "
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtGuiIBItem.IB09_sSALN) & "'" & " As [IB09_sSALN], "
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtGuiIBItem.IB09a_sPolicyNo) & "'" & " As [IB09a_sPolicyNo], "
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtGuiIBItem.IB10_sInsuredName) & "'" & " As [IB10_sInsuredName], "
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtGuiIBItem.IB11_sLossLocation) & "'" & " As [IB11_sLossLocation], "
    If StrComp(pudtGuiIBItem.IB12_dtDateOfLoss, "Null", vbTextCompare) = 0 Then
        sSQL = sSQL & "Null As [IB12_dtDateOfLoss], "
    Else
        sSQL = sSQL & "#" & pudtGuiIBItem.IB12_dtDateOfLoss & "#" & " As [IB12_dtDateOfLoss], "
    End If
    sSQL = sSQL & CCur(pudtGuiIBItem.IB13_cGrossLoss) & " As [IB13_cGrossLoss], "
    sSQL = sSQL & CCur(pudtGuiIBItem.IB14_cDepreciation) & " As [IB14_cDepreciation], "
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtGuiIBItem.IB14a_sSupplement) & "'" & " As [IB14a_sSupplement], "
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtGuiIBItem.IB14b_sRebilled) & "'" & " As [IB14b_sRebilled], "
    sSQL = sSQL & CCur(pudtGuiIBItem.IB15_cDeductible) & " As [IB15_cDeductible], "
    sSQL = sSQL & CCur(pudtGuiIBItem.IB15a_cLessExcessLimits) & " As [IB15a_cLessExcessLimits], "
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtGuiIBItem.IB15b_sExcessLimDesc) & "'" & " As [IB15b_sExcessLimDesc], "
    sSQL = sSQL & CCur(pudtGuiIBItem.IB15c_cLessMiscellaneous) & " As [IB15c_cLessMiscellaneous], "
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtGuiIBItem.IB15d_cMiscellaneousDesc) & "'" & " As [IB15d_cMiscellaneousDesc], "
    sSQL = sSQL & CCur(pudtGuiIBItem.IB16_cNetClaim) & " As [IB16_cNetClaim], "
    sSQL = sSQL & CCur(pudtGuiIBItem.IB17_cServiceFee) & " As [IB17_cServiceFee], "
    sSQL = sSQL & CCur(pudtGuiIBItem.IB17a_cMiscServiceFee) & " As [IB17a_cMiscServiceFee], "
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtGuiIBItem.IB18_sServiceFeeComment) & "'" & " As [IB18_sServiceFeeComment], "
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtGuiIBItem.IB18a_sMiscServiceFeeComment) & "'" & " As [IB18a_sMiscServiceFeeComment], "
    sSQL = sSQL & CCur(pudtGuiIBItem.IB25_cServiceFeeSubTotal) & " As [IB25_cServiceFeeSubTotal], "
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtGuiIBItem.IB29a_sMiscExpenseFeeComment) & "'" & " As [IB29a_sMiscExpenseFeeComment], "
    sSQL = sSQL & CCur(pudtGuiIBItem.IB29b_cMiscExpenseFee) & " As [IB29b_cMiscExpenseFee], "
    sSQL = sSQL & CCur(pudtGuiIBItem.IB30_cTotalExpenses) & " As [IB30_cTotalExpenses], "
    sSQL = sSQL & CCur(pudtGuiIBItem.IB31_dTaxPercent) & " As [IB31_dTaxPercent], "
    sSQL = sSQL & CCur(pudtGuiIBItem.IB32_cTaxAmount) & " As [IB32_cTaxAmount], "
    sSQL = sSQL & CCur(pudtGuiIBItem.IB33_cTotalAdjustingFee) & " As [IB33_cTotalAdjustingFee], "
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtGuiIBItem.IB33a_sAccountCode) & "'" & " As [IB33a_sAccountCode], "
    sSQL = sSQL & pudtGuiIBItem.FeeScheduleID & " As [FeeScheduleID], "
    sSQL = sSQL & pudtGuiIBItem.Void & " As [Void], "
    sSQL = sSQL & pudtGuiIBItem.FeeByTime & " As [FeeByTime], "
    sSQL = sSQL & pudtGuiIBItem.UseActivityTime & " As [UseActivityTime], "
    sSQL = sSQL & pudtGuiIBItem.DownLoadMe & " As [DownLoadMe], "
    sSQL = sSQL & pudtGuiIBItem.UpLoadMe & " As [UpLoadMe], "
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtGuiIBItem.Comments) & "'" & " As [Comments], "
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtGuiIBItem.AdminComments) & "'" & " As [AdminComments], "
    sSQL = sSQL & "#" & pudtGuiIBItem.DateLastUpdated & "#" & " As [DateLastUpdated], "
    sSQL = sSQL & pudtGuiIBItem.UpdateByUserID & " As [UpdateByUserID] "
    
    Set oConn = New ADODB.Connection
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    oConn.Execute sSQL, lRecordsAffected
    
    psID = sID
    
    Sleep 200
    
    AddIB = CBool(lRecordsAffected)
    
CLEAN_UP:
    Set oConn = Nothing
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function AddIB"
End Function


Public Function EditIBFee(pudtIBFeeItem As GuiIBFeeItem) As Boolean
    On Error GoTo EH
    Dim sSQL As String
    Dim oConn As ADODB.Connection
    Dim lRecordsAffected As Long
    
    sSQL = "UPDATE IBFee Set "
    sSQL = sSQL & "[IBFeeID] = " & pudtIBFeeItem.IBFeeID & ", "
    sSQL = sSQL & "[AssignmentsID] = " & pudtIBFeeItem.AssignmentsID & ", "
    sSQL = sSQL & "[IBID] = " & pudtIBFeeItem.IBID & ", "
    sSQL = sSQL & "[ID] = " & pudtIBFeeItem.ID & ", "
    sSQL = sSQL & "[IDAssignments] = " & pudtIBFeeItem.IDAssignments & ", "
    sSQL = sSQL & "[IDIB] = " & pudtIBFeeItem.IDIB & ", "
    sSQL = sSQL & "[FeeScheduleFeeTypesID] = " & pudtIBFeeItem.FeeScheduleFeeTypesID & ", "
    sSQL = sSQL & "[NumberOfItems] = " & pudtIBFeeItem.NumberOfItems & ", "
    sSQL = sSQL & "[Amount] = " & CCur(pudtIBFeeItem.Amount) & ", "
    sSQL = sSQL & "[Comment]  = '" & goUtil.utCleanSQLString(pudtIBFeeItem.Comment) & "', "
    sSQL = sSQL & "[DownLoadMe] = " & pudtIBFeeItem.DownLoadMe & ", "
    sSQL = sSQL & "[UpLoadMe] = " & pudtIBFeeItem.UpLoadMe & ", "
    sSQL = sSQL & "[AdminComments] = '" & goUtil.utCleanSQLString(pudtIBFeeItem.AdminComments) & "', "
    sSQL = sSQL & "[DateLastUpdated] = #" & pudtIBFeeItem.DateLastUpdated & "#, "
    sSQL = sSQL & "[UpdateByUserID] = " & pudtIBFeeItem.UpdateByUserID & " "
    sSQL = sSQL & "WHERE AssignmentsID = " & pudtIBFeeItem.AssignmentsID & " "
    If pudtIBFeeItem.ID = "[ID]" Then
        sSQL = sSQL & "AND FeeScheduleFeeTypesID = " & pudtIBFeeItem.FeeScheduleFeeTypesID & " "
        sSQL = sSQL & "AND [IDIB] = " & pudtIBFeeItem.IDIB & " "
    Else
        sSQL = sSQL & "AND ID = " & pudtIBFeeItem.ID & " "
    End If
    
    Set oConn = New ADODB.Connection
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    oConn.Execute sSQL, lRecordsAffected

    Sleep 10
    
    EditIBFee = CBool(lRecordsAffected)
    
CLEAN_UP:
    Set oConn = Nothing
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function EditIBFee"
End Function

Public Function AddIBFee(pudtIBFeeItem As GuiIBFeeItem, psID As String) As Boolean
    On Error GoTo EH
    Dim sSQL As String
    Dim oConn As ADODB.Connection
    Dim sID As String
    Dim lRecordsAffected As Long
    
    sID = goUtil.GetAccessDBUID("ID", "IBFee")
    
    With pudtIBFeeItem
        .IBFeeID = sID
        .ID = sID
    End With
    
    sSQL = "INSERT INTO IBFee "
    sSQL = sSQL & "( "
    sSQL = sSQL & "[IBFeeID], "
    sSQL = sSQL & "[AssignmentsID], "
    sSQL = sSQL & "[IBID], "
    sSQL = sSQL & "[ID], "
    sSQL = sSQL & "[IDAssignments], "
    sSQL = sSQL & "[IDIB], "
    sSQL = sSQL & "[FeeScheduleFeeTypesID], "
    sSQL = sSQL & "[NumberOfItems], "
    sSQL = sSQL & "[Amount], "
    sSQL = sSQL & "[Comment], "
    sSQL = sSQL & "[DownLoadMe], "
    sSQL = sSQL & "[UpLoadMe], "
    sSQL = sSQL & "[AdminComments], "
    sSQL = sSQL & "[DateLastUpdated], "
    sSQL = sSQL & "[UpdateByUserID] "
    sSQL = sSQL & ") "
    sSQL = sSQL & "SELECT "
    sSQL = sSQL & pudtIBFeeItem.IBFeeID & " As [IBFeeID], "
    sSQL = sSQL & pudtIBFeeItem.AssignmentsID & " As [AssignmentsID], "
    sSQL = sSQL & pudtIBFeeItem.IBID & " As [IBID], "
    sSQL = sSQL & pudtIBFeeItem.ID & " As [ID] , "
    sSQL = sSQL & pudtIBFeeItem.IDAssignments & " As [IDAssignments], "
    sSQL = sSQL & pudtIBFeeItem.IDIB & " As [IDIB], "
    sSQL = sSQL & pudtIBFeeItem.FeeScheduleFeeTypesID & " As [FeeScheduleFeeTypesID], "
    sSQL = sSQL & pudtIBFeeItem.NumberOfItems & " As [NumberOfItems], "
    sSQL = sSQL & CCur(pudtIBFeeItem.Amount) & " As [Amount], "
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtIBFeeItem.Comment) & "'" & " As [Comment], "
    sSQL = sSQL & pudtIBFeeItem.DownLoadMe & " As [DownLoadMe], "
    sSQL = sSQL & pudtIBFeeItem.UpLoadMe & " As [UpLoadMe], "
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtIBFeeItem.AdminComments) & "'" & " As [AdminComments], "
    sSQL = sSQL & "#" & pudtIBFeeItem.DateLastUpdated & "#" & " As [DateLastUpdated], "
    sSQL = sSQL & pudtIBFeeItem.UpdateByUserID & " As [UpdateByUserID] "
    
    Set oConn = New ADODB.Connection
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    oConn.Execute sSQL, lRecordsAffected
    
    psID = sID
    
    Sleep 10
    
    AddIBFee = CBool(lRecordsAffected)
    
CLEAN_UP:
    Set oConn = Nothing
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function AddIBFee"
End Function



Private Sub lstAddServiceFeeItemNumberOfItems_LostFocus()
    CalcAddServiceFeeItem
End Sub

Private Sub lvwExpenseFees_DblClick()
    On Error GoTo EH
    
    ShowEditExpenseFeesItem lvwExpenseFees.SelectedItem
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lvwExpenseFees_DblClick"
End Sub

Private Sub lvwExpenseFees_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error GoTo EH
    
    lvwExpenseFees.ToolTipText = lvwExpenseFees.SelectedItem.Text
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lvwExpenseFees_ItemClick"
End Sub

'--------------------------------END IB DB----------------------------------------

Private Sub lvwExpenseFees_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    LooseOverridesFocus
End Sub



Private Sub lvwOverrideFees_DblClick()
    On Error GoTo EH

    ShowEditOverrideFeesItem lvwOverrideFees.SelectedItem

    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lvwOverrideFees_DblClick"
End Sub

Private Sub lvwOverrideFees_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error GoTo EH

    lvwOverrideFees.ToolTipText = lvwOverrideFees.SelectedItem.Text

    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lvwOverrideFees_ItemClick"
End Sub

Private Sub lvwOverrideFees_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo EH
    
'    'Need to expand on focus lack of Realestate in 800X 600 restraint
'    framOverrideFees.Height = ORFEES_FRAM_HEIGHT_ONFOCUS
'    lvwOverrideFees.Height = ORFEES_LVW_HEIGHT_ONFOCUS
'    txtServiceFeeComment.Visible = False
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lvwOverrideFees_MouseMove"
End Sub

Private Sub lvwOverrideFees_LostFocus()
    On Error GoTo EH
'    If txtOverrideFeeItemComment.Visible Then
'        Exit Sub
'    End If
'    'Need to expand on focus lack of Realestate in 800X 600 restraint
'    framOverrideFees.Height = ORFEES_FRAM_HEIGHT_LOSTFOCUS
'    lvwOverrideFees.Height = ORFEES_LVW_HEIGHT_LOSTFOCUS
'    txtServiceFeeComment.Visible = True
'
'    txtServiceFeeComment.Visible = True
    
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lvwOverrideFees_LostFocus"
End Sub

Private Sub lvwOverrideFees_Click()
    On Error GoTo EH

    lvwOverrideFees.ToolTipText = lvwOverrideFees.SelectedItem.Text
'    ShowEditOverrideFeesItem lvwOverrideFees.SelectedItem
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lvwOverrideFees_Click"
End Sub

Private Sub lvwOverrideFees_GotFocus()
    On Error GoTo EH
    
'    mlIgnoreMouseMove = 5
'
'    'Need to expand on focus lack of Realestate in 800X 600 restraint
'    framOverrideFees.Height = ORFEES_FRAM_HEIGHT_ONFOCUS
'    lvwOverrideFees.Height = ORFEES_LVW_HEIGHT_ONFOCUS
'
'    txtServiceFeeComment.Visible = False
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lvwOverrideFees_GotFocus"
End Sub

Private Sub lvwOverrideFees_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo EH
    Select Case KeyCode
        Case KeyCodeConstants.vbKeyReturn, KeyCodeConstants.vbKeySpace
            If Not lvwOverrideFees.SelectedItem Is Nothing Then
                ShowEditOverrideFeesItem lvwOverrideFees.SelectedItem
            End If
    End Select
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lvwOverrideFees_KeyDown"
End Sub

Private Sub lvwOverrideFees_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo EH
    
    Select Case KeyCode
        Case KeyCodeConstants.vbKeyTab
            SelectedNextItmX lvwOverrideFees, CBool(Shift)
    End Select
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lvwOverrideFees_KeyUp"
End Sub

Public Sub ShowEditOverrideFeesItem(pitmX As ListItem)
    On Error GoTo EH
    Dim lCount As Long
    Dim sCount As String
    Dim lMaxItems As Long
    Dim sNumberOfItems As String
    Dim lMyLeft As Long
    Dim lMyWidth As Long
    Dim lMyHeight As Long
    Dim sVBFormula As String
    Dim sVBFormula2 As String
    Dim bItemChecked As Boolean
    Dim sMess As String
    Dim sDescription As String
    Dim itmX As ListItem
    Dim sServiceFeeComment As String
    Dim sfsftFeeAmount As String
    
    'Set the member varibales for this item
    Set mSelectedOverrideFeeItemX = pitmX
    
    mbEditOverrideFeesItemX = True
    lvwOverrideFees.Enabled = False
    cmdExit.Cancel = False
    
    sVBFormula = pitmX.ListSubItems(GuiOverrideFeesListView.fsftVBFormula - 1)
    bItemChecked = CBool(pitmX.ListSubItems(GuiOverrideFeesListView.NumberOfItems - 1))
    sDescription = pitmX.Text
    'If the item is checked then need to ask if really want to uncheck Item
    If bItemChecked Then
'        sMess = "Do you really want to uncheck " & sDescription & " ?"
'
'        If MsgBox(sMess, vbYesNo + vbQuestion, framOverrideFees.Caption) = vbYes Then
            CalcOverrideFeeItem True
'        Else
'            HideOverrideFeeItem
'        End If
        Exit Sub
    Else
'        sMess = "Do you really want to Check " & sDescription & " ?"
'        sMess = sMess & vbCrLf & vbCrLf
'        If StrComp(sVBFormula, VBFORMULA_OVERRIDES_ALL, vbTextCompare) = 0 Then
'            sMess = sMess & "This item will override all other Service Fees!" & vbCrLf & vbCrLf
'        ElseIf StrComp(sVBFormula, VBFORMULA_OVERRIDES_SERVICE_FEE, vbTextCompare) = 0 Then
'            sMess = sMess & "Only check this item if you need to manually override the Service Fee Amount!"
'        ElseIf StrComp(sVBFormula, VBFORMULA_OVERRIDES_FEEBYTIME_FEE, vbTextCompare) = 0 Then
'            sMess = sMess & "Only check this item if you need to manually override the Fee By Time Amount!" & vbCrLf
'            sMess = sMess & "***Note Enter a time amount not a dollar amount!" & vbCrLf & vbCrLf
'            sMess = sMess & "The amount you enter will be calculated to figure the Service Fee dollar amount."
'        End If
'        sMess = sMess & vbCrLf & vbCrLf & "(Any other " & framOverrideFees.Caption & " will be unchecked.)"
'        If MsgBox(sMess, vbYesNo + vbQuestion, framOverrideFees.Caption) = vbNo Then
'            HideOverrideFeeItem
'            Exit Sub
'        End If
    End If

    txtOverrideFeeItemComment.top = pitmX.top + 280
    lMyLeft = lvwOverrideFees.ColumnHeaders.Item(GuiOverrideFeesListView.fsftDescription).Width
    lMyLeft = lMyLeft + lvwOverrideFees.left
    lMyLeft = lMyLeft + 40
    txtOverrideFeeItemComment.left = lMyLeft
    lMyWidth = lvwOverrideFees.ColumnHeaders.Item(GuiOverrideFeesListView.Comment).Width
    txtOverrideFeeItemComment.Width = lMyWidth
    lMyLeft = lMyLeft + lMyWidth
    lMyLeft = lMyLeft + lvwOverrideFees.ColumnHeaders.Item(GuiOverrideFeesListView.NumberOfItems).Width
'    txtOverrideFeeItemAmount.left = lMyLeft
    txtOverrideFeeItemAmount.left = lvwOverrideFees.left
    txtOverrideFeeItemAmount.Width = lvwOverrideFees.ColumnHeaders.Item(GuiOverrideFeesListView.fsftDescription).Width
    txtOverrideFeeItemAmount.top = pitmX.top + 280

    'populate the Comment
    txtOverrideFeeItemComment.Text = pitmX.ListSubItems(GuiOverrideFeesListView.Comment - 1)

    'First have to uncheck all other items
    For Each itmX In lvwOverrideFees.ListItems
        bItemChecked = CBool(itmX.ListSubItems(GuiOverrideFeesListView.NumberOfItems - 1))
        sVBFormula2 = itmX.ListSubItems(GuiOverrideFeesListView.fsftVBFormula - 1)
        sDescription = itmX.Text
        If bItemChecked Then
            itmX.ListSubItems(GuiOverrideFeesListView.NumberOfItems - 1) = "0"
            itmX.ListSubItems(GuiOverrideFeesListView.NumberOfItems - 1).ReportIcon = GuiIBStatusList.IsUnchecked
            itmX.ListSubItems(GuiOverrideFeesListView.Comment - 1) = vbNullString
            itmX.ListSubItems(GuiOverrideFeesListView.Amount - 1) = "0.00"
            itmX.ListSubItems(GuiOverrideFeesListView.UpLoadMe - 1) = goUtil.GetFlagText(True)
            itmX.ListSubItems(GuiOverrideFeesListView.DateLastUpdated - 1) = Format(Now(), "MM/DD/YYYY HH:MM:SS")
            itmX.ListSubItems(GuiOverrideFeesListView.UpdateByUserID - 1) = goUtil.gsCurUsersID
            'if the Item checked happens to be the fee by time ovveride then
            'need to clear some things
            If StrComp(sVBFormula2, VBFORMULA_OVERRIDES_FEEBYTIME_FEE, vbTextCompare) = 0 Then
                mbOverridesFeeByTimeFee = False
                txtActLogHours.Text = Format(mdblCurrentActLogTime, "##0.00")
                chkUseActivityTime.Value = vbChecked
                txtServiceFee.Text = Format(mdblCurrentActLogTime * mcFeeServiceHourlyRate, "#,###,###,##0.00")
                SumServiceFees
                sServiceFeeComment = mdblCurrentActLogTime & " @ " & mcFeeServiceHourlyRate
                txtServiceFeeComment.Text = sServiceFeeComment
                txtServiceFeeComment.ToolTipText = sServiceFeeComment
            ElseIf StrComp(sVBFormula2, VBFORMULA_OVERRIDES_SERVICE_FEE, vbTextCompare) = 0 Then
                mbOverridesServiceFee = False
            ElseIf StrComp(sVBFormula2, VBFORMULA_OVERRIDES_ALL, vbTextCompare) = 0 Then
                framServiceFees.Enabled = True
                framExpenses.Enabled = True
                ShowFrame
            End If
        End If
    Next
    pitmX.ListSubItems(GuiServiceFeesListView.NumberOfItems - 1) = "1"
    pitmX.ListSubItems(GuiServiceFeesListView.NumberOfItems - 1).ReportIcon = GuiIBStatusList.IsChecked

    'populate Amount
    If CLng(pitmX.ListSubItems(GuiOverrideFeesListView.Amount - 1)) = 0 Then
        sfsftFeeAmount = pitmX.ListSubItems(GuiOverrideFeesListView.fsftFeeAmount - 1)
        txtOverrideFeeItemAmount.Text = Format(sfsftFeeAmount, "#,###,###,##0.00")
    Else
        txtOverrideFeeItemAmount.Text = Format(pitmX.ListSubItems(GuiOverrideFeesListView.Amount - 1), "#,###,###,##0.00")
    End If
    

    'Make them visible
'    txtOverrideFeeItemComment.Visible = True
    If CLng(txtOverrideFeeItemAmount.Text) > 0 Then
        txtOverrideFeeItemAmount.left = -5000
    End If
    txtOverrideFeeItemAmount.Visible = True

    mbEditOverrideFeesItemX = False
'    txtOverrideFeeItemComment.SetFocus
    txtOverrideFeeItemAmount.SetFocus
    If CLng(txtOverrideFeeItemAmount.Text) > 0 Then
        lvwOverrideFees.Enabled = True
        lvwOverrideFees.SetFocus
    End If
    Exit Sub
EH:
   goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub ShowEditOverrideFeesItem"
End Sub

Private Sub CalcOverrideFeeItem(Optional pbUncheckItem As Boolean = False)
    On Error GoTo EH
    Dim lNumItems As Long
    Dim cFeeAmount As Currency
    Dim cAmount As Currency
    Dim cMaxFeeAmount As Currency
    Dim sSQL As String
    Dim oConn As ADODB.Connection
    Dim RS As ADODB.Recordset
    Dim sVBFormula As String
    Dim sRTIBFeeID As String
    Dim bUseFormula As Boolean
    Dim sServiceFeeComment As String
    Dim sFlagText As String
    Dim sMess As String

    If mSelectedOverrideFeeItemX Is Nothing Then
        Exit Sub
    End If
    
    If pbUncheckItem Then
        lNumItems = 0
    Else
        lNumItems = 1
    End If
    
    If lNumItems = 0 Then
        mSelectedOverrideFeeItemX.ListSubItems(GuiOverrideFeesListView.NumberOfItems - 1) = 0
        mSelectedOverrideFeeItemX.ListSubItems(GuiOverrideFeesListView.NumberOfItems - 1).ReportIcon = GuiIBStatusList.IsUnchecked
        mSelectedOverrideFeeItemX.ListSubItems(GuiOverrideFeesListView.Comment - 1) = vbNullString
    Else
        mSelectedOverrideFeeItemX.ListSubItems(GuiOverrideFeesListView.NumberOfItems - 1) = 1
        mSelectedOverrideFeeItemX.ListSubItems(GuiOverrideFeesListView.NumberOfItems - 1).ReportIcon = GuiIBStatusList.IsChecked
        mSelectedOverrideFeeItemX.ListSubItems(GuiOverrideFeesListView.Comment - 1) = txtOverrideFeeItemComment.Text
    End If
    
    sRTIBFeeID = mSelectedOverrideFeeItemX.ListSubItems(GuiOverrideFeesListView.RTIBFeeID - 1)
    bUseFormula = CBool(mSelectedOverrideFeeItemX.ListSubItems(GuiOverrideFeesListView.fsftUseFormula - 1))
    sVBFormula = mSelectedOverrideFeeItemX.ListSubItems(GuiOverrideFeesListView.fsftVBFormula - 1)
    cMaxFeeAmount = CCur(mSelectedOverrideFeeItemX.ListSubItems(GuiOverrideFeesListView.fsftMaxFeeAmount - 1))
    cFeeAmount = CCur(mSelectedOverrideFeeItemX.ListSubItems(GuiOverrideFeesListView.fsftFeeAmount - 1))

    'If using formula need to make DB Connection
    If bUseFormula Then
        If StrComp(sVBFormula, VBFORMULA_OVERRIDES_ALL, vbTextCompare) = 0 Then
            If lNumItems = 0 Then
                cAmount = 0#
                chkFeeByTime.Enabled = True
                cmdCalcFeeSched.Enabled = True
                sServiceFeeComment = vbNullString
                mbOverridesALL = False
            Else
                cAmount = CCur(txtOverrideFeeItemAmount.Text)
                chkFeeByTime.Enabled = False
                cmdCalcFeeSched.Enabled = False
                sServiceFeeComment = mSelectedOverrideFeeItemX.Text
                mbOverridesALL = True
            End If
            mbOverridesFeeByTimeFee = False
            mbOverridesServiceFee = False
            chkFeeByTime.Value = vbUnchecked
            txtServiceFee.Text = Format(cAmount, "#,###,###,##0.00")
            
            txtServiceFeeComment.Text = sServiceFeeComment
            txtServiceFeeComment.ToolTipText = sServiceFeeComment
        ElseIf StrComp(sVBFormula, VBFORMULA_OVERRIDES_SERVICE_FEE, vbTextCompare) = 0 Then
            If lNumItems = 0 Then
                cAmount = 0#
                chkFeeByTime.Enabled = True
                cmdCalcFeeSched.Enabled = True
                sServiceFeeComment = vbNullString
            Else
                cAmount = CCur(txtOverrideFeeItemAmount.Text)
                chkFeeByTime.Enabled = False
                cmdCalcFeeSched.Enabled = False
                sServiceFeeComment = mSelectedOverrideFeeItemX.Text
            End If
            mbOverridesServiceFee = True
            mbOverridesFeeByTimeFee = False
            mbOverridesALL = False
            chkFeeByTime.Value = vbUnchecked
            txtServiceFee.Text = Format(cAmount, "#,###,###,##0.00")
            txtServiceFeeComment.Text = sServiceFeeComment
            txtServiceFeeComment.ToolTipText = sServiceFeeComment
        ElseIf StrComp(sVBFormula, VBFORMULA_OVERRIDES_FEEBYTIME_FEE, vbTextCompare) = 0 Then
            If lNumItems = 0 Then
                cAmount = 0#
                chkFeeByTime.Enabled = True
                cmdCalcFeeSched.Enabled = True
                mbOverridesFeeByTimeFee = False
                mbOverridesServiceFee = False
                mbOverridesALL = False
                txtActLogHours.Text = Format(mdblCurrentActLogTime, "##0.00")
                chkUseActivityTime.Value = vbChecked
                txtServiceFee.Text = Format(mdblCurrentActLogTime * mcFeeServiceHourlyRate, "#,###,###,##0.00")
                sServiceFeeComment = Format(mdblCurrentActLogTime, "##0.00") & " @ " & Format(mcFeeServiceHourlyRate, "##0.00")
                txtServiceFeeComment.Text = sServiceFeeComment
                txtServiceFeeComment.ToolTipText = sServiceFeeComment
            Else
                cAmount = CCur(txtOverrideFeeItemAmount.Text)
                chkFeeByTime.Enabled = False
                cmdCalcFeeSched.Enabled = False
                mbOverridesFeeByTimeFee = True
                mbOverridesServiceFee = False
                mbOverridesALL = False
                If mcFeeServiceHourlyRate > 0 Then
                mdblOverridesActLogTime = Format((cAmount), "##0.00")
                Else
                    mdblOverridesActLogTime = 0#
                End If
                txtActLogHours.Text = Format(mdblOverridesActLogTime, "##0.00")
                chkUseActivityTime.Value = vbUnchecked
                txtServiceFee.Text = Format(mdblOverridesActLogTime * mcFeeServiceHourlyRate, "#,###,###,##0.00")
                sServiceFeeComment = Format(mdblOverridesActLogTime, "##0.00") & " @ " & Format(mcFeeServiceHourlyRate, "##0.00")
                txtServiceFeeComment.Text = sServiceFeeComment
                txtServiceFeeComment.ToolTipText = sServiceFeeComment
                mbPopulateIBInfo = True
                chkFeeByTime.Value = vbChecked
                mbPopulateIBInfo = False
            End If
        Else
            Set oConn = New ADODB.Connection
            Set RS = New ADODB.Recordset
            goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
            sSQL = "SELECT "
            sSQL = sSQL & "( " & sVBFormula & ") As MyCalcFeeAmount "
            sSQL = sSQL & "FROM( "
            sSQL = sSQL & "SELECT RTIBFee.[RTIBFeeID], "
            sSQL = sSQL & "RTIBFee.[AssignmentsID], "
            sSQL = sSQL & "RTIBFee.[ID], "
            sSQL = sSQL & "RTIBFee.[IDAssignments], "
            sSQL = sSQL & "RTIBFee.[FeeScheduleFeeTypesID], "
            sSQL = sSQL & "FSFT.[FeeScheduleID] As [fsftFeeScheduleID], "
            sSQL = sSQL & "FSFT.[TypeNum] As [fsftTypeNum], "
            sSQL = sSQL & "FSFT.[Name] As [fsftName], "
            sSQL = sSQL & "FSFT.[Description] As [fsftDescription], "
            sSQL = sSQL & "FSFT.[FeeAmount] As [fsftFeeAmount], "
            sSQL = sSQL & "FSFT.[IsExpense] As [fsftIsExpense], "
            sSQL = sSQL & "FSFT.[MaxNumberOfItems] As [fsftMaxNumberOfItems], "
            sSQL = sSQL & "FSFT.[MaxFeeAmount] As [fsftMaxFeeAmount], "
            sSQL = sSQL & "FSFT.[IsMiscAmount] As [fsftIsMiscAmount], "
            sSQL = sSQL & "FSFT.[UseFormula] As [fsftUseFormula], "
            sSQL = sSQL & "FSFT.[VBFormula] As [fsftVBFormula], "
            sSQL = sSQL & "FSFT.[IsDeleted] As [fsftIsDeleted], "
            sSQL = sSQL & "FSFT.[DateLastUpdated] As [fsftDateLastUpdated], "
            sSQL = sSQL & "FSFT.[UpdateByUserID] As [fsftUpdateByUserID], "
            sSQL = sSQL & lNumItems & " As [NumberOfItems], "
            sSQL = sSQL & "RTIBFee.[Amount], "
            sSQL = sSQL & "RTIBFee.[Comment], "
            sSQL = sSQL & "RTIBFee.[DownLoadMe], "
            sSQL = sSQL & "RTIBFee.[UpLoadMe], "
            sSQL = sSQL & "RTIBFee.[AdminComments], "
            sSQL = sSQL & "RTIBFee.[DateLastUpdated], "
            sSQL = sSQL & "RTIBFee.[UpdateByUserID] "
            sSQL = sSQL & "FROM RTIBFee INNER JOIN FeeScheduleFeeTypes FSFT ON RTIBFee.[FeeScheduleFeeTypesID] = FSFT.[FeeScheduleFeeTypesID] "
            sSQL = sSQL & "WHERE RTIBFee.[AssignmentsID] = " & msAssignmentsID & " "
            sSQL = sSQL & "AND  [RTIBFeeID] = " & sRTIBFeeID & " "
            sSQL = sSQL & ") RetRTIBfee "
    
            RS.CursorLocation = adUseClient
            RS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
            Set RS.ActiveConnection = Nothing
    
            If RS.RecordCount > 0 Then
                cAmount = goUtil.IsNullIsVbNullString(RS.Fields("MyCalcFeeAmount"))
            Else
                cAmount = 0
            End If
        End If
    Else
        If lNumItems = 0 Then
            cAmount = txtOverrideFeeItemAmount.Text
        Else
            cAmount = lNumItems * cFeeAmount
        End If

    End If

    If cMaxFeeAmount > 0 Then
        If cAmount > cMaxFeeAmount Then
            cAmount = cMaxFeeAmount
        End If
    End If

    mSelectedOverrideFeeItemX.ListSubItems(GuiOverrideFeesListView.Amount - 1) = Format(cAmount, "#,###,###,##0.00")
    sFlagText = goUtil.GetFlagText(True)
    mSelectedOverrideFeeItemX.ListSubItems(GuiOverrideFeesListView.UpLoadMe - 1) = sFlagText
    mSelectedOverrideFeeItemX.ListSubItems(GuiOverrideFeesListView.DateLastUpdated - 1) = Format(Now(), "MM/DD/YYYY HH:MM:SS")
    mSelectedOverrideFeeItemX.ListSubItems(GuiOverrideFeesListView.UpdateByUserID - 1) = goUtil.gsCurUsersID
    
    
    'Also Need to Zero out fee lists if the override items is ALL
    If StrComp(sVBFormula, VBFORMULA_OVERRIDES_ALL, vbTextCompare) = 0 Then
        'Only zero out if this is Checking the override all not unchecking !
        If Not pbUncheckItem Then
'            sMess = "Do you want to clear All other Service Fees ?"
'            If MsgBox(sMess, vbQuestion + vbYesNo, "Clear Service Fees") = vbYes Then
                ZeroOutFeeList lvwServiceFees
'                ZeroOutFeeList lvwExpenseFees
                framServiceFees.Enabled = False
'                framExpenses.Enabled = False
'            Else
'                framServiceFees.Enabled = True
'                framExpenses.Enabled = True
'            End If
            ShowFrame
        Else
            framServiceFees.Enabled = True
            framExpenses.Enabled = True
            ShowFrame
        End If
    End If
    
    HideOverrideFeeItem
    mlIgnoreMouseMove = 0
'    LooseOverridesFocus
    SumServiceFees
    'cleanup
    Set RS = Nothing
    Set oConn = Nothing
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub CalcOverrideFeeItem"
End Sub

Private Sub HideOverrideFeeItem()
    On Error GoTo EH
    
    txtOverrideFeeItemComment.Visible = False
    txtOverrideFeeItemAmount.Visible = False
    
    If Not mSelectedOverrideFeeItemX Is Nothing Then
        lvwOverrideFees.Enabled = True
        mSelectedOverrideFeeItemX.EnsureVisible
        Set mSelectedOverrideFeeItemX = Nothing
'        If txtServiceFee.Enabled Then
'            txtServiceFee.SetFocus
'        End If
    End If
    
    cmdExit.Cancel = True
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub HideOverrideFeeItem"
End Sub

Private Sub lvwServiceFees_Click()
    On Error GoTo EH
    
    lvwServiceFees.ToolTipText = lvwServiceFees.SelectedItem.Text
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lvwServiceFees_Click"
End Sub

Private Sub lvwServiceFees_DblClick()
    On Error GoTo EH
    
    ShowEditServiceFeesItem lvwServiceFees.SelectedItem
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lvwServiceFees_DblClick"
End Sub


Private Sub lvwServiceFees_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error GoTo EH
    
    lvwServiceFees.ToolTipText = lvwServiceFees.SelectedItem.Text
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lvwServiceFees_ItemClick"
End Sub

Private Sub lvwServiceFees_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo EH
    Select Case KeyCode
        Case KeyCodeConstants.vbKeyReturn, KeyCodeConstants.vbKeySpace
            If Not lvwExpenseFees.SelectedItem Is Nothing Then
                ShowEditServiceFeesItem lvwServiceFees.SelectedItem
            End If
        Case KeyCodeConstants.vbKeyDelete
            ShowEditServiceFeesItem lvwServiceFees.SelectedItem, CInt(KeyCodeConstants.vbKey0 - 48)
        Case Else
            If KeyCode >= KeyCodeConstants.vbKey0 And KeyCode <= KeyCodeConstants.vbKey9 Then
                ShowEditServiceFeesItem lvwServiceFees.SelectedItem, CInt(KeyCode - 48)
            End If
            If KeyCode >= KeyCodeConstants.vbKeyNumpad0 And KeyCode <= KeyCodeConstants.vbKeyNumpad9 Then
                ShowEditServiceFeesItem lvwServiceFees.SelectedItem, CInt(KeyCode - 96)
            End If
            
    End Select
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lvwServiceFees_KeyDown"
End Sub


Private Sub SelectedNextItmX(poLvw As ListView, Optional pbMoveUp As Boolean = True)
    On Error GoTo EH
    
    If Not poLvw.SelectedItem Is Nothing And poLvw.Enabled And poLvw.Visible Then
        If Not pbMoveUp Then
            If poLvw.SelectedItem.Index < poLvw.ListItems.Count Then
                poLvw.ListItems.Item(poLvw.SelectedItem.Index + 1).Selected = True
                poLvw.SetFocus
                poLvw.SelectedItem.EnsureVisible
            End If
        Else
            If poLvw.SelectedItem.Index > 1 Then
                poLvw.ListItems.Item(poLvw.SelectedItem.Index - 1).Selected = True
                poLvw.SetFocus
                poLvw.SelectedItem.EnsureVisible
            End If
        End If
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub SelectedNextItmX"
End Sub

Public Sub ShowEditServiceFeesItem(pitmX As ListItem, Optional NumKey As Integer = -1)
    On Error GoTo EH
    Dim lCount As Long
    Dim sCount As String
    Dim lMaxItems As Long
    Dim sNumberOfItems As String
    Dim lMyLeft As Long
    Dim lMyWidth As Long
    Dim lMyHeight As Long
    
    mbEditServiceFeeItemX = True
    lvwServiceFees.Enabled = False
    cmdExit.Cancel = False
    
    'Set the member varibales for this item
    Set mSelectedAddServiceFeeItemX = pitmX
    
    txtAddServiceFeeItemComment.top = pitmX.top + 280
    lMyLeft = lvwServiceFees.ColumnHeaders.Item(GuiServiceFeesListView.fsftDescription).Width
    lMyLeft = lMyLeft + lvwServiceFees.left
    txtAddServiceFeeItemComment.left = lMyLeft + 40
    lMyWidth = lvwServiceFees.ColumnHeaders.Item(GuiServiceFeesListView.Comment).Width
    txtAddServiceFeeItemComment.Width = lMyWidth
    lstAddServiceFeeItemNumberOfItems.top = pitmX.top + 280
    lstAddServiceFeeItemNumberOfItems.left = txtMiscServiceFee.left - 80
    lMyHeight = (framServiceFees.Height - lstAddServiceFeeItemNumberOfItems.top) / 2
    lstAddServiceFeeItemNumberOfItems.Height = lMyHeight
    txtAddServiceFeeItemAmount.top = pitmX.top + 280
    txtAddServiceFeeItemAmount.left = txtTtlServiceFee.left - 80
    
    'populate the Comment
    txtAddServiceFeeItemComment.Text = pitmX.ListSubItems(GuiServiceFeesListView.Comment - 1)
    
    'Populate and select the correct number of items
    lstAddServiceFeeItemNumberOfItems.Clear
    lMaxItems = CLng(pitmX.ListSubItems(GuiServiceFeesListView.fsftMaxNumberOfItems - 1))
    sNumberOfItems = pitmX.ListSubItems(GuiServiceFeesListView.NumberOfItems - 1)
    'Populate the max number of items
    For lCount = 0 To lMaxItems
        sCount = lCount
        lstAddServiceFeeItemNumberOfItems.AddItem sCount
        If lstAddServiceFeeItemNumberOfItems.List(lstAddServiceFeeItemNumberOfItems.NewIndex) = sNumberOfItems Then
            lstAddServiceFeeItemNumberOfItems.Text = lstAddServiceFeeItemNumberOfItems.List(lstAddServiceFeeItemNumberOfItems.NewIndex)
        End If
    Next
    
    'populate Amount
    txtAddServiceFeeItemAmount.Text = Format(pitmX.ListSubItems(GuiServiceFeesListView.Amount - 1), "#,###,###,##0.00")
    
    'Make them visible
    If NumKey > -1 Then
        If lMaxItems < 10 Then
            lstAddServiceFeeItemNumberOfItems.left = -5000
        End If
        lstAddServiceFeeItemNumberOfItems.Text = NumKey
    End If
    
'    txtAddServiceFeeItemComment.Visible = True
    lstAddServiceFeeItemNumberOfItems.Visible = True
'    txtAddServiceFeeItemAmount.Visible = True
    
    
    mbEditServiceFeeItemX = False
'    txtAddServiceFeeItemComment.SetFocus
    If lstAddServiceFeeItemNumberOfItems.Enabled Then
        lstAddServiceFeeItemNumberOfItems.SetFocus
    End If
    If NumKey > -1 Then
        If lMaxItems > 9 And NumKey <> 0 Then
            lstAddServiceFeeItemNumberOfItems.SelStart = 1
        Else
            lvwServiceFees.Enabled = True
            lvwServiceFees.SetFocus
        End If
    End If
    
    
    Exit Sub
EH:
   goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub ShowEditServiceFeesItem"
End Sub

Private Sub lstAddServiceFeeItemNumberOfItems_DblClick()
    On Error GoTo EH
    
    If mbEditServiceFeeItemX Then
        Exit Sub
    End If
    
    If lstAddServiceFeeItemNumberOfItems.Text = 0 Then
        txtAddServiceFeeItemAmount = "0.00"
    End If
    
    CalcAddServiceFeeItem
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lstAddServiceFeeItemNumberOfItems_DblClick"
End Sub

Private Sub lstAddServiceFeeItemNumberOfItems_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo EH
    
    Select Case KeyCode
        Case KeyCodeConstants.vbKeyReturn
            If lstAddServiceFeeItemNumberOfItems.Text = 0 Then
                txtAddServiceFeeItemAmount = "0.00"
            End If
            CalcAddServiceFeeItem
        Case KeyCodeConstants.vbKeyEscape
            HideAddServiceFeeItem
    End Select
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lstAddServiceFeeItemNumberOfItems_KeyDown"
End Sub

Private Sub lvwServiceFees_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo EH
    
'    LooseOverridesFocus
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lvwServiceFees_MouseMove"
End Sub



Private Sub Timer_PickIB_Timer()
    On Error GoTo EH
    Dim sIBItem As String
    Dim lPos As Long
    Dim bPickItem As Boolean
    Dim bAddIB As Boolean
    Dim bEditIB As Boolean
    
    Timer_PickIB.Enabled = False
    'If the Add Button is visible then see if
    'there are any IBs already added.
    'If there are then don't try to add a supplement.
    'Just Pick the IB... if it is Current (That means it is still open)
    'Then go straight to edit mode on it
    If cmdAddEditIB.Visible And StrComp(cmdAddEditIB.Caption, "&Add IB", vbTextCompare) = 0 Then
        If cboBillingID.ListCount = 1 Then
            'if there is only one Count (The Select item)
            'then add an IB
            bAddIB = True
        Else
            'if there is more than one item
            'Select the Last one, but don't Rebill it, let the user decide
            lPos = cboBillingID.ListCount - 1
            bPickItem = True
        End If
    ElseIf Not cmdAddEditIB.Visible Then
        'Select the Last available item
        'Loop through all the available IBs looking for
        'a current one, select it and Edit it
        'Otherwise just select it
        If cboBillingID.ListCount > 1 Then
            bPickItem = True
            For lPos = 0 To cboBillingID.ListCount
                sIBItem = cboBillingID.List(lPos)
                If InStr(1, sIBItem, "Current", vbTextCompare) > 0 Then
                    bEditIB = True
                    Exit For
                End If
            Next
        End If
    End If
    
    If bAddIB Then
        cmdAddEditIB_Click
    ElseIf bPickItem Then
        cboBillingID.ListIndex = lPos
        If bEditIB Then
            If cmdAddEditIB.Visible And StrComp(cmdAddEditIB.Caption, "&Edit IB", vbTextCompare) = 0 Then
                cmdAddEditIB_Click
            End If
        End If
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Timer_PickIB_Timer"
End Sub

Private Sub txtActLogHours_GotFocus()
    goUtil.utSelText txtActLogHours
End Sub

Private Sub txtAddExpenseFeeItemAmount_DblClick()
    CalcAddExpenseFeeItem
End Sub

Private Sub txtAddExpenseFeeItemAmount_GotFocus()
    goUtil.utSelText txtAddExpenseFeeItemAmount
End Sub

Private Sub txtAddExpenseFeeItemAmount_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo EH
    
    Select Case KeyCode
        Case KeyCodeConstants.vbKeyReturn
            CalcAddExpenseFeeItem
        Case KeyCodeConstants.vbKeyEscape
            HideAddExpenseFeeItem
    End Select
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub txtAddExpenseFeeItemAmount_KeyDown"
End Sub

Private Sub txtAddExpenseFeeItemAmount_LostFocus()
    goUtil.utValidate , txtAddExpenseFeeItemAmount
    CalcAddExpenseFeeItem
End Sub

Private Sub txtAddExpenseFeeItemComment_DblClick()
    CalcAddExpenseFeeItem
End Sub

Private Sub txtAddExpenseFeeItemComment_GotFocus()
    goUtil.utSelText txtAddExpenseFeeItemComment
End Sub


Private Sub txtAddServiceFeeItemAmount_DblClick()
    CalcAddServiceFeeItem
End Sub

Private Sub txtAddServiceFeeItemComment_DblClick()
    CalcAddServiceFeeItem
End Sub

Private Sub txtComments_Change()
    On Error GoTo EH
    Dim lPos As Long
    
    If InStr(1, txtComments.Text, vbCrLf, vbBinaryCompare) > 0 Then
        lPos = txtComments.SelStart
        txtComments.Text = Replace(txtComments.Text, vbCrLf, vbNullString)
        txtComments.SelStart = lPos
    End If
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub txtComments_Change"
End Sub

Private Sub txtComments_GotFocus()
    goUtil.utSelText txtComments
End Sub

Private Sub txtdtDateClosed_LostFocus()
    goUtil.utValidate , txtdtDateClosed
End Sub

Private Sub txtFeeServiceHourlyRate_GotFocus()
    goUtil.utSelText txtFeeServiceHourlyRate
End Sub

Private Sub txtMiscExpenseFee_Change()
    SumExpenseFees
End Sub

Private Sub txtMiscExpenseFee_LostFocus()
    goUtil.utValidate , txtMiscExpenseFee
End Sub

Private Sub txtMiscExpenseFeeComment_Click()
    On Error GoTo EH
    If lvwExpenseFees.ListItems.Count > 0 Then
        If lvwExpenseFees.SelectedItem.Index < lvwExpenseFees.ListItems.Count Then
            lvwExpenseFees.ListItems(lvwExpenseFees.ListItems.Count).Selected = True
            txtMiscExpenseFeeComment.SetFocus
        End If
    End If
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub txtMiscExpenseFeeComment_Click"
End Sub

Private Sub txtMiscServiceFee_Change()
    SumServiceFees
End Sub

Private Sub txtMiscServiceFee_LostFocus()
    goUtil.utValidate , txtMiscServiceFee
End Sub

Private Sub txtMiscServiceFeeComment_Click()
    On Error GoTo EH
    If lvwServiceFees.ListItems.Count > 0 Then
        If lvwServiceFees.SelectedItem.Index < lvwServiceFees.ListItems.Count Then
            lvwServiceFees.ListItems(lvwServiceFees.ListItems.Count).Selected = True
            txtMiscServiceFeeComment.SetFocus
        End If
    End If
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub txtMiscServiceFeeComment_Click"
End Sub


Private Sub txtOverrideDepreciation_Change()
    If Not mbOverridesServiceFee _
        And Not mbOverridesFeeByTimeFee _
        And Not mbOverridesALL _
        And chkFeeByTime.Value = vbUnchecked Then
        CalcFeeSched
        SumServiceFees
    End If
End Sub

Private Sub txtOverrideDepreciation_GotFocus()
    goUtil.utSelText txtOverrideDepreciation
End Sub

Private Sub txtOverrideDepreciation_LostFocus()
    goUtil.utValidate , txtOverrideDepreciation
End Sub

Private Sub txtOverrideExcessLimit_Change()
    If Not mbOverridesServiceFee _
        And Not mbOverridesFeeByTimeFee _
        And Not mbOverridesALL _
        And chkFeeByTime.Value = vbUnchecked Then
        CalcFeeSched
        SumServiceFees
    End If
End Sub

Private Sub txtOverrideExcessLimit_GotFocus()
    goUtil.utSelText txtOverrideExcessLimit
End Sub

Private Sub txtOverrideExcessLimit_LostFocus()
    goUtil.utValidate , txtOverrideExcessLimit
End Sub

Private Sub txtOverrideFeeItemAmount_DblClick()
    CalcOverrideFeeItem
End Sub

Private Sub txtOverrideFeeItemAmount_GotFocus()
    goUtil.utSelText txtOverrideFeeItemAmount
End Sub

Private Sub txtOverrideFeeItemAmount_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo EH
    
    Select Case KeyCode
        Case KeyCodeConstants.vbKeyReturn
            CalcOverrideFeeItem
        Case KeyCodeConstants.vbKeyEscape
            'Since Escaping need to uncheck the item
            mSelectedOverrideFeeItemX.ListSubItems(GuiServiceFeesListView.NumberOfItems - 1) = "0"
            mSelectedOverrideFeeItemX.ListSubItems(GuiServiceFeesListView.NumberOfItems - 1).ReportIcon = GuiIBStatusList.IsUnchecked
            HideOverrideFeeItem
            'Check the Service Fee comment for a Previous Overrides entry
            'if there was one in there then zero out the amount and null the comment
            'only do this if the previous comment was an overrding one.
            If cmdCalcFeeSched.Visible Then
                If Not cmdCalcFeeSched.Enabled Then
                    txtServiceFeeComment.Text = vbNullString
                    txtServiceFee.Text = "0.00"
                    SumServiceFees
                End If
            End If
    End Select
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub txtOverrideFeeItemAmount_KeyDown"
End Sub

Private Sub txtOverrideFeeItemAmount_LostFocus()
    goUtil.utValidate , txtOverrideFeeItemAmount
    CalcOverrideFeeItem
End Sub

Private Sub txtAddServiceFeeItemAmount_GotFocus()
    goUtil.utSelText txtAddServiceFeeItemAmount
End Sub

Private Sub txtAddServiceFeeItemAmount_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo EH
    
    Select Case KeyCode
        Case KeyCodeConstants.vbKeyReturn
            CalcAddServiceFeeItem
        Case KeyCodeConstants.vbKeyEscape
            HideAddServiceFeeItem
    End Select
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub txtAddServiceFeeItemAmount_KeyDown"
End Sub

Private Sub txtAddServiceFeeItemAmount_LostFocus()
    goUtil.utValidate , txtAddServiceFeeItemAmount
    CalcAddServiceFeeItem
End Sub

Private Sub txtAddServiceFeeItemComment_GotFocus()
    goUtil.utSelText txtAddServiceFeeItemComment
End Sub

Private Sub txtAddServiceFeeItemComment_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo EH
    
    Select Case KeyCode
        Case KeyCodeConstants.vbKeyReturn
            CalcAddServiceFeeItem
        Case KeyCodeConstants.vbKeyEscape
            HideAddServiceFeeItem
    End Select
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub txtAddServiceFeeItemComment_KeyDown"
End Sub

Private Sub CalcAddServiceFeeItem()
    On Error GoTo EH
    Dim lNumItems As Long
    Dim cFeeAmount As Currency
    Dim cAmount As Currency
    Dim cMaxFeeAmount As Currency
    Dim sSQL As String
    Dim oConn As ADODB.Connection
    Dim RS As ADODB.Recordset
    Dim sVBFormula As String
    Dim sRTIBFeeID As String
    Dim bUseFormula As Boolean
    Dim sFlagText As String
    
    If mSelectedAddServiceFeeItemX Is Nothing Then
        Exit Sub
    End If
    
    mSelectedAddServiceFeeItemX.ListSubItems(GuiServiceFeesListView.Comment - 1) = txtAddServiceFeeItemComment.Text
    mSelectedAddServiceFeeItemX.ListSubItems(GuiServiceFeesListView.NumberOfItems - 1) = lstAddServiceFeeItemNumberOfItems.Text
    lNumItems = CLng(lstAddServiceFeeItemNumberOfItems.Text)
    sRTIBFeeID = mSelectedAddServiceFeeItemX.ListSubItems(GuiServiceFeesListView.RTIBFeeID - 1)
    bUseFormula = CBool(mSelectedAddServiceFeeItemX.ListSubItems(GuiServiceFeesListView.fsftUseFormula - 1))
    sVBFormula = mSelectedAddServiceFeeItemX.ListSubItems(GuiServiceFeesListView.fsftVBFormula - 1)
    cMaxFeeAmount = CCur(mSelectedAddServiceFeeItemX.ListSubItems(GuiServiceFeesListView.fsftMaxFeeAmount - 1))
    cFeeAmount = CCur(mSelectedAddServiceFeeItemX.ListSubItems(GuiServiceFeesListView.fsftFeeAmount - 1))
    
    'If using formula need to make DB Connection
    If bUseFormula And lNumItems > 0 Then
        Set oConn = New ADODB.Connection
        Set RS = New ADODB.Recordset
        goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
        
        sSQL = "SELECT "
        sSQL = sSQL & "( " & sVBFormula & ") As MyCalcFeeAmount "
        sSQL = sSQL & "FROM( "
        sSQL = sSQL & "SELECT RTIBFee.[RTIBFeeID], "
        sSQL = sSQL & "RTIBFee.[AssignmentsID], "
        sSQL = sSQL & "RTIBFee.[ID], "
        sSQL = sSQL & "RTIBFee.[IDAssignments], "
        sSQL = sSQL & "RTIBFee.[FeeScheduleFeeTypesID], "
        sSQL = sSQL & "FSFT.[FeeScheduleID] As [fsftFeeScheduleID], "
        sSQL = sSQL & "FSFT.[TypeNum] As [fsftTypeNum], "
        sSQL = sSQL & "FSFT.[Name] As [fsftName], "
        sSQL = sSQL & "FSFT.[Description] As [fsftDescription], "
        sSQL = sSQL & "FSFT.[FeeAmount] As [fsftFeeAmount], "
        sSQL = sSQL & "FSFT.[IsExpense] As [fsftIsExpense], "
        sSQL = sSQL & "FSFT.[MaxNumberOfItems] As [fsftMaxNumberOfItems], "
        sSQL = sSQL & "FSFT.[MaxFeeAmount] As [fsftMaxFeeAmount], "
        sSQL = sSQL & "FSFT.[IsMiscAmount] As [fsftIsMiscAmount], "
        sSQL = sSQL & "FSFT.[UseFormula] As [fsftUseFormula], "
        sSQL = sSQL & "FSFT.[VBFormula] As [fsftVBFormula], "
        sSQL = sSQL & "FSFT.[IsDeleted] As [fsftIsDeleted], "
        sSQL = sSQL & "FSFT.[DateLastUpdated] As [fsftDateLastUpdated], "
        sSQL = sSQL & "FSFT.[UpdateByUserID] As [fsftUpdateByUserID], "
        sSQL = sSQL & lNumItems & " As [NumberOfItems], "
        sSQL = sSQL & "RTIBFee.[Amount], "
        sSQL = sSQL & "RTIBFee.[Comment], "
        sSQL = sSQL & "RTIBFee.[DownLoadMe], "
        sSQL = sSQL & "RTIBFee.[UpLoadMe], "
        sSQL = sSQL & "RTIBFee.[AdminComments], "
        sSQL = sSQL & "RTIBFee.[DateLastUpdated], "
        sSQL = sSQL & "RTIBFee.[UpdateByUserID] "
        sSQL = sSQL & "FROM RTIBFee INNER JOIN FeeScheduleFeeTypes FSFT ON RTIBFee.[FeeScheduleFeeTypesID] = FSFT.[FeeScheduleFeeTypesID] "
        sSQL = sSQL & "WHERE RTIBFee.[AssignmentsID] = " & msAssignmentsID & " "
        sSQL = sSQL & "AND  [RTIBFeeID] = " & sRTIBFeeID & " "
        sSQL = sSQL & ") RetRTIBfee "
        
        RS.CursorLocation = adUseClient
        RS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
        Set RS.ActiveConnection = Nothing
        
        If RS.RecordCount > 0 Then
            cAmount = goUtil.IsNullIsVbNullString(RS.Fields("MyCalcFeeAmount"))
        Else
            cAmount = 0
        End If
    Else
        If lNumItems = 0 Then
            If IsNumeric(txtAddServiceFeeItemAmount.Text) Then
                cAmount = txtAddServiceFeeItemAmount.Text
            Else
                cAmount = 0
            End If
        Else
            cAmount = lNumItems * cFeeAmount
        End If
        
    End If
    
    If cMaxFeeAmount > 0 Then
        If cAmount > cMaxFeeAmount Then
            cAmount = cMaxFeeAmount
        End If
    End If
    
    mSelectedAddServiceFeeItemX.ListSubItems(GuiServiceFeesListView.Amount - 1) = Format(cAmount, "#,###,###,##0.00")
    sFlagText = goUtil.GetFlagText(True)
    mSelectedAddServiceFeeItemX.ListSubItems(GuiServiceFeesListView.UpLoadMe - 1) = sFlagText
    mSelectedAddServiceFeeItemX.ListSubItems(GuiServiceFeesListView.DateLastUpdated - 1) = Format(Now(), "MM/DD/YYYY HH:MM:SS")
    mSelectedAddServiceFeeItemX.ListSubItems(GuiServiceFeesListView.UpdateByUserID - 1) = goUtil.gsCurUsersID
    
    HideAddServiceFeeItem
    
    SumServiceFees
    'cleanup
    Set RS = Nothing
    Set oConn = Nothing
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub CalcAddServiceFeeItem"
End Sub

Private Sub HideAddServiceFeeItem()
    On Error GoTo EH
    
    txtAddServiceFeeItemComment.Visible = False
    lstAddServiceFeeItemNumberOfItems.Visible = False
    txtAddServiceFeeItemAmount.Visible = False
    
    If Not mSelectedAddServiceFeeItemX Is Nothing Then
        lvwServiceFees.Enabled = True
        mSelectedAddServiceFeeItemX.EnsureVisible
        Set mSelectedAddServiceFeeItemX = Nothing
        lvwServiceFees.SetFocus
    End If
    
    cmdExit.Cancel = True
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub HideAddServiceFeeItem"
End Sub


Private Sub lvwExpenseFees_Click()
    On Error GoTo EH
    
    lvwExpenseFees.ToolTipText = lvwExpenseFees.SelectedItem.Text
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lvwExpenseFees_Click"
End Sub

Private Sub lvwExpenseFees_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo EH
    Select Case KeyCode
        Case KeyCodeConstants.vbKeyReturn, KeyCodeConstants.vbKeySpace
            If Not lvwExpenseFees.SelectedItem Is Nothing Then
                ShowEditExpenseFeesItem lvwExpenseFees.SelectedItem
            End If
        Case KeyCodeConstants.vbKeyDelete
            ShowEditExpenseFeesItem lvwExpenseFees.SelectedItem, CInt(KeyCodeConstants.vbKey0 - 48)
        Case Else
            If KeyCode >= KeyCodeConstants.vbKey0 And KeyCode <= KeyCodeConstants.vbKey9 Then
                ShowEditExpenseFeesItem lvwExpenseFees.SelectedItem, CInt(KeyCode - 48)
            End If
            If KeyCode >= KeyCodeConstants.vbKeyNumpad0 And KeyCode <= KeyCodeConstants.vbKeyNumpad9 Then
                ShowEditExpenseFeesItem lvwExpenseFees.SelectedItem, CInt(KeyCode - 96)
            End If
    End Select
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lvwExpenseFees_KeyDown"
End Sub


Private Sub lvwExpenseFees_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo EH
    
    Select Case KeyCode
        Case KeyCodeConstants.vbKeyTab
            SelectedNextItmX lvwExpenseFees, CBool(Shift)
    End Select
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lvwExpenseFees_KeyUp"
End Sub

Public Sub ShowEditExpenseFeesItem(pitmX As ListItem, Optional NumKey As Integer = -1)
    On Error GoTo EH
    Dim lCount As Long
    Dim sCount As String
    Dim lMaxItems As Long
    Dim sNumberOfItems As String
    Dim lMyLeft As Long
    Dim lMyWidth As Long
    Dim lMyHeight As Long
    
    mbEditExpenseFeeItemX = True
    lvwExpenseFees.Enabled = False
    cmdExit.Cancel = False
    
    'Set the member varibales for this item
    Set mSelectedAddExpenseFeeItemX = pitmX
    
    txtAddExpenseFeeItemComment.top = pitmX.top + 280
    lMyLeft = lvwExpenseFees.ColumnHeaders.Item(GuiExpenseFeesListView.fsftDescription).Width
    lMyLeft = lMyLeft + lvwExpenseFees.left
    txtAddExpenseFeeItemComment.left = lMyLeft + 40
    lMyWidth = lvwExpenseFees.ColumnHeaders.Item(GuiExpenseFeesListView.Comment).Width
    txtAddExpenseFeeItemComment.Width = lMyWidth
    lstAddExpenseFeeItemNumberOfItems.top = pitmX.top + 280
    lstAddExpenseFeeItemNumberOfItems.left = txtMiscExpenseFee.left - 80
    lMyHeight = (framExpenses.Height - lstAddExpenseFeeItemNumberOfItems.top) / 2
    lstAddExpenseFeeItemNumberOfItems.Height = lMyHeight
    txtAddExpenseFeeItemAmount.top = pitmX.top + 280
    txtAddExpenseFeeItemAmount.left = txtTtlExpenses.left - 80
    
    'populate the Comment
    txtAddExpenseFeeItemComment.Text = pitmX.ListSubItems(GuiExpenseFeesListView.Comment - 1)
    
    'Populate and select the correct number of items
    lstAddExpenseFeeItemNumberOfItems.Clear
    lMaxItems = CLng(pitmX.ListSubItems(GuiExpenseFeesListView.fsftMaxNumberOfItems - 1))
    sNumberOfItems = pitmX.ListSubItems(GuiExpenseFeesListView.NumberOfItems - 1)
    'Populate the max number of items
    For lCount = 0 To lMaxItems
        sCount = lCount
        lstAddExpenseFeeItemNumberOfItems.AddItem sCount
        If lstAddExpenseFeeItemNumberOfItems.List(lstAddExpenseFeeItemNumberOfItems.NewIndex) = sNumberOfItems Then
            lstAddExpenseFeeItemNumberOfItems.Text = lstAddExpenseFeeItemNumberOfItems.List(lstAddExpenseFeeItemNumberOfItems.NewIndex)
        End If
    Next
    
    'Populate amount
    txtAddExpenseFeeItemAmount.Text = pitmX.ListSubItems(GuiExpenseFeesListView.Amount - 1)
    
    'Make them visible
    If NumKey > -1 Then
        If lMaxItems < 10 Then
            lstAddExpenseFeeItemNumberOfItems.left = -5000
        End If
        lstAddExpenseFeeItemNumberOfItems.Text = NumKey
    End If
    
'    txtAddExpenseFeeItemComment.Visible = True
    lstAddExpenseFeeItemNumberOfItems.Visible = True
'    txtAddExpenseFeeItemAmount.Visible = True
    
    mbEditExpenseFeeItemX = False
'    If txtAddExpenseFeeItemComment.Enabled Then
'        txtAddExpenseFeeItemComment.SetFocus
'    End If
    If lstAddExpenseFeeItemNumberOfItems.Enabled Then
        lstAddExpenseFeeItemNumberOfItems.SetFocus
    End If
    
    If NumKey > -1 Then
        If lMaxItems > 9 And NumKey <> 0 Then
            lstAddExpenseFeeItemNumberOfItems.SelStart = 1
        Else
            lvwExpenseFees.Enabled = True
            lvwExpenseFees.SetFocus
        End If
    End If
    
    Exit Sub
EH:
   goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub ShowEditExpenseFeesItem"
End Sub

Private Sub lstAddExpenseFeeItemNumberOfItems_DblClick()
    On Error GoTo EH
    
    If mbEditExpenseFeeItemX Then
       Exit Sub
    End If
    
    If lstAddExpenseFeeItemNumberOfItems.Text = 0 Then
        txtAddExpenseFeeItemAmount.Text = "0.00"
    End If
    
    CalcAddExpenseFeeItem
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lstAddExpenseFeeItemNumberOfItems_DblClick"
End Sub

Private Sub lstAddExpenseFeeItemNumberOfItems_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo EH
    
    Select Case KeyCode
        Case KeyCodeConstants.vbKeyReturn
            If lstAddExpenseFeeItemNumberOfItems.Text = 0 Then
                txtAddExpenseFeeItemAmount.Text = "0.00"
            End If
            CalcAddExpenseFeeItem
        Case KeyCodeConstants.vbKeyEscape
            HideAddExpenseFeeItem
    End Select
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lstAddExpenseFeeItemNumberOfItems_KeyDown"
End Sub

Private Sub txtAddExpenseFeeItemComment_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo EH
    
    Select Case KeyCode
        Case KeyCodeConstants.vbKeyReturn
            CalcAddExpenseFeeItem
        Case KeyCodeConstants.vbKeyEscape
            HideAddExpenseFeeItem
    End Select
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub txtAddExpenseFeeItemComment_KeyDown"
End Sub


Private Sub CalcAddExpenseFeeItem()
    On Error GoTo EH
    Dim lNumItems As Long
    Dim cFeeAmount As Currency
    Dim cAmount As Currency
    Dim cMaxFeeAmount As Currency
    Dim sSQL As String
    Dim oConn As ADODB.Connection
    Dim RS As ADODB.Recordset
    Dim sVBFormula As String
    Dim sRTIBFeeID As String
    Dim bUseFormula As Boolean
    Dim sFlagText As String
    
    If mSelectedAddExpenseFeeItemX Is Nothing Then
        Exit Sub
    End If
    
    mSelectedAddExpenseFeeItemX.ListSubItems(GuiExpenseFeesListView.Comment - 1) = txtAddExpenseFeeItemComment.Text
    mSelectedAddExpenseFeeItemX.ListSubItems(GuiExpenseFeesListView.NumberOfItems - 1) = lstAddExpenseFeeItemNumberOfItems.Text
    If IsNumeric(lstAddExpenseFeeItemNumberOfItems.Text) Then
        lNumItems = CLng(lstAddExpenseFeeItemNumberOfItems.Text)
    Else
        lNumItems = 0
    End If
    sRTIBFeeID = mSelectedAddExpenseFeeItemX.ListSubItems(GuiExpenseFeesListView.RTIBFeeID - 1)
    bUseFormula = CBool(mSelectedAddExpenseFeeItemX.ListSubItems(GuiExpenseFeesListView.fsftUseFormula - 1))
    sVBFormula = mSelectedAddExpenseFeeItemX.ListSubItems(GuiExpenseFeesListView.fsftVBFormula - 1)
    cMaxFeeAmount = CCur(mSelectedAddExpenseFeeItemX.ListSubItems(GuiExpenseFeesListView.fsftMaxFeeAmount - 1))
    cFeeAmount = CCur(mSelectedAddExpenseFeeItemX.ListSubItems(GuiExpenseFeesListView.fsftFeeAmount - 1))
    
    'If using formula need to make DB Connection
    If bUseFormula And lNumItems > 0 Then
        Set oConn = New ADODB.Connection
        Set RS = New ADODB.Recordset
        goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
        
        sSQL = "SELECT "
        sSQL = sSQL & "( " & sVBFormula & ") As MyCalcFeeAmount "
        sSQL = sSQL & "FROM( "
        sSQL = sSQL & "SELECT RTIBFee.[RTIBFeeID], "
        sSQL = sSQL & "RTIBFee.[AssignmentsID], "
        sSQL = sSQL & "RTIBFee.[ID], "
        sSQL = sSQL & "RTIBFee.[IDAssignments], "
        sSQL = sSQL & "RTIBFee.[FeeScheduleFeeTypesID], "
        sSQL = sSQL & "FSFT.[FeeScheduleID] As [fsftFeeScheduleID], "
        sSQL = sSQL & "FSFT.[TypeNum] As [fsftTypeNum], "
        sSQL = sSQL & "FSFT.[Name] As [fsftName], "
        sSQL = sSQL & "FSFT.[Description] As [fsftDescription], "
        sSQL = sSQL & "FSFT.[FeeAmount] As [fsftFeeAmount], "
        sSQL = sSQL & "FSFT.[IsExpense] As [fsftIsExpense], "
        sSQL = sSQL & "FSFT.[MaxNumberOfItems] As [fsftMaxNumberOfItems], "
        sSQL = sSQL & "FSFT.[MaxFeeAmount] As [fsftMaxFeeAmount], "
        sSQL = sSQL & "FSFT.[IsMiscAmount] As [fsftIsMiscAmount], "
        sSQL = sSQL & "FSFT.[UseFormula] As [fsftUseFormula], "
        sSQL = sSQL & "FSFT.[VBFormula] As [fsftVBFormula], "
        sSQL = sSQL & "FSFT.[IsDeleted] As [fsftIsDeleted], "
        sSQL = sSQL & "FSFT.[DateLastUpdated] As [fsftDateLastUpdated], "
        sSQL = sSQL & "FSFT.[UpdateByUserID] As [fsftUpdateByUserID], "
        sSQL = sSQL & lNumItems & " As [NumberOfItems], "
        sSQL = sSQL & "RTIBFee.[Amount], "
        sSQL = sSQL & "RTIBFee.[Comment], "
        sSQL = sSQL & "RTIBFee.[DownLoadMe], "
        sSQL = sSQL & "RTIBFee.[UpLoadMe], "
        sSQL = sSQL & "RTIBFee.[AdminComments], "
        sSQL = sSQL & "RTIBFee.[DateLastUpdated], "
        sSQL = sSQL & "RTIBFee.[UpdateByUserID] "
        sSQL = sSQL & "FROM RTIBFee INNER JOIN FeeScheduleFeeTypes FSFT ON RTIBFee.[FeeScheduleFeeTypesID] = FSFT.[FeeScheduleFeeTypesID] "
        sSQL = sSQL & "WHERE RTIBFee.[AssignmentsID] = " & msAssignmentsID & " "
        sSQL = sSQL & "AND  [RTIBFeeID] = " & sRTIBFeeID & " "
        sSQL = sSQL & ") RetRTIBfee "
        
        RS.CursorLocation = adUseClient
        RS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
        Set RS.ActiveConnection = Nothing
        
        If RS.RecordCount > 0 Then
            cAmount = goUtil.IsNullIsVbNullString(RS.Fields("MyCalcFeeAmount"))
        Else
            cAmount = 0
        End If
    Else
        If lNumItems = 0 Then
            If IsNumeric(txtAddExpenseFeeItemAmount.Text) Then
                cAmount = txtAddExpenseFeeItemAmount.Text
            Else
                cAmount = 0
            End If
        Else
            cAmount = lNumItems * cFeeAmount
        End If
        
    End If
    
    If cMaxFeeAmount > 0 Then
        If cAmount > cMaxFeeAmount Then
            cAmount = cMaxFeeAmount
        End If
    End If
    
    mSelectedAddExpenseFeeItemX.ListSubItems(GuiExpenseFeesListView.Amount - 1) = Format(cAmount, "#,###,###,##0.00")
    sFlagText = goUtil.GetFlagText(True)
    mSelectedAddExpenseFeeItemX.ListSubItems(GuiServiceFeesListView.UpLoadMe - 1) = sFlagText
    mSelectedAddExpenseFeeItemX.ListSubItems(GuiServiceFeesListView.DateLastUpdated - 1) = Format(Now(), "MM/DD/YYYY HH:MM:SS")
    mSelectedAddExpenseFeeItemX.ListSubItems(GuiServiceFeesListView.UpdateByUserID - 1) = goUtil.gsCurUsersID
    
    HideAddExpenseFeeItem
    SumExpenseFees
    
    'cleanup
    
    Set RS = Nothing
    Set oConn = Nothing
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub CalcAddExpenseFeeItem"
End Sub

Private Sub HideAddExpenseFeeItem()
    On Error GoTo EH
    
    txtAddExpenseFeeItemComment.Visible = False
    lstAddExpenseFeeItemNumberOfItems.Visible = False
    txtAddExpenseFeeItemAmount.Visible = False
    
    If Not mSelectedAddExpenseFeeItemX Is Nothing Then
        lvwExpenseFees.Enabled = True
        mSelectedAddExpenseFeeItemX.EnsureVisible
        Set mSelectedAddExpenseFeeItemX = Nothing
        If lvwExpenseFees.Enabled Then
            lvwExpenseFees.SetFocus
        End If
    End If
    cmdExit.Cancel = True
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub HideAddExpenseFeeItem"
End Sub

Private Sub txtDateLastUpdated_GotFocus()
    goUtil.utSelText txtDateLastUpdated
End Sub

Private Sub txtdtDateClosed_GotFocus()
    goUtil.utSelText txtdtDateClosed
End Sub

Private Sub txtMiscExpenseFee_GotFocus()
    goUtil.utSelText txtMiscExpenseFee
End Sub

Private Sub txtMiscExpenseFeeComment_GotFocus()
    goUtil.utSelText txtMiscExpenseFeeComment
    SelectedNextItmX lvwExpenseFees, False
End Sub

Private Sub txtMiscServiceFee_GotFocus()
    goUtil.utSelText txtMiscServiceFee
End Sub

Private Sub txtMiscServiceFeeComment_GotFocus()
    goUtil.utSelText txtMiscServiceFeeComment
    SelectedNextItmX lvwServiceFees, False
End Sub

Private Sub txtOverrideFeeItemComment_DblClick()
    CalcOverrideFeeItem
End Sub

Private Sub txtOverrideFeeItemComment_GotFocus()
    goUtil.utSelText txtOverrideFeeItemComment
End Sub

Private Sub txtOverrideFeeItemComment_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo EH
    
    Select Case KeyCode
        Case KeyCodeConstants.vbKeyReturn
            CalcOverrideFeeItem
        Case KeyCodeConstants.vbKeyEscape
            'Since Escaping need to uncheck the item
            mSelectedOverrideFeeItemX.ListSubItems(GuiServiceFeesListView.NumberOfItems - 1) = "0"
            mSelectedOverrideFeeItemX.ListSubItems(GuiServiceFeesListView.NumberOfItems - 1).ReportIcon = GuiIBStatusList.IsUnchecked
            HideOverrideFeeItem
            If cmdCalcFeeSched.Visible Then
                If Not cmdCalcFeeSched.Enabled Then
                    txtServiceFeeComment.Text = vbNullString
                    txtServiceFee.Text = "0.00"
                    SumServiceFees
                End If
            End If
    End Select
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub txtOverrideFeeItemComment_KeyDown"
End Sub


Private Sub txtOverrideGrossLoss_Change()
    If IsNumeric(txtOverrideGrossLoss.Text) Then
        mcOverrideGrossLoss = CCur(txtOverrideGrossLoss.Text)
    End If
    If Not mbOverridesServiceFee _
        And Not mbOverridesFeeByTimeFee _
        And Not mbOverridesALL _
        And chkFeeByTime.Value = vbUnchecked Then
        CalcFeeSched
        SumServiceFees
    End If
End Sub

Private Sub txtOverrideGrossLoss_GotFocus()
    goUtil.utSelText txtOverrideGrossLoss
End Sub

Private Sub txtOverrideGrossLoss_LostFocus()
    goUtil.utValidate , txtOverrideGrossLoss
End Sub


Private Sub txtOverrideMiscellaneous_Change()
    If Not mbOverridesServiceFee _
        And Not mbOverridesFeeByTimeFee _
        And Not mbOverridesALL _
        And chkFeeByTime.Value = vbUnchecked Then
        CalcFeeSched
        SumServiceFees
    End If
End Sub

Private Sub txtOverrideMiscellaneous_GotFocus()
    goUtil.utSelText txtOverrideMiscellaneous
End Sub

Private Sub txtOverrideMiscellaneous_LostFocus()
    goUtil.utValidate , txtOverrideMiscellaneous
End Sub

Private Sub txtServiceFee_Change()
    SumServiceFees
End Sub

Private Sub txtServiceFee_Click()
    On Error GoTo EH
    If lvwServiceFees.ListItems.Count > 0 Then
        If lvwServiceFees.SelectedItem.Index > 1 Then
            lvwServiceFees.ListItems(1).Selected = True
            txtServiceFee.SetFocus
        End If
    End If
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub txtServiceFee_Click"
End Sub

Private Sub txtServiceFee_GotFocus()
    goUtil.utSelText txtServiceFee
    SelectedNextItmX lvwServiceFees, True
End Sub

Private Sub txtServiceFee_LostFocus()
    goUtil.utValidate , txtServiceFee
End Sub

Private Sub txtServiceFeeComment_Change()
    On Error GoTo EH
    Dim lPos As Long
    
    If InStr(1, txtServiceFeeComment.Text, vbCrLf, vbBinaryCompare) > 0 Then
        lPos = txtServiceFeeComment.SelStart
        txtServiceFeeComment.Text = Replace(txtServiceFeeComment.Text, vbCrLf, vbNullString)
        txtServiceFeeComment.SelStart = lPos
        txtServiceFee.SetFocus
    End If
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub txtServiceFeeComment_Change"
End Sub

Private Sub txtServiceFeeComment_GotFocus()
     goUtil.utSelText txtServiceFeeComment
'     txtServiceFeeComment.Width = 4575
'     txtServiceFeeComment.Height = 1935
End Sub

Private Sub txtServiceFeeComment_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo EH
    Select Case KeyCode
        Case KeyCodeConstants.vbKeyReturn, KeyCodeConstants.vbKeyEscape
            txtServiceFee.SetFocus
    End Select
Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub txtServiceFeeComment_KeyDown"
End Sub

Private Sub txtServiceFeeComment_LostFocus()
'    txtServiceFeeComment.Width = 2535
'    txtServiceFeeComment.Height = 495
End Sub


Private Sub txtTaxAmount_GotFocus()
    goUtil.utSelText txtTaxAmount
End Sub

Private Sub txtTaxPercent_Change()
    SumTax
End Sub

Private Sub txtTaxPercent_GotFocus()
    goUtil.utSelText txtTaxPercent
End Sub

Private Sub txtTaxPercent_LostFocus()
    goUtil.utValidate , txtTaxPercent
End Sub

Private Sub txtTotalAdjustingFee_GotFocus()
    goUtil.utSelText txtTotalAdjustingFee
End Sub

Private Sub txtTtlExpenses_GotFocus()
    goUtil.utSelText txtTtlExpenses
End Sub

Private Sub txtTtlServiceExp_GotFocus()
    goUtil.utSelText txtTtlServiceExp
End Sub

Private Sub txtTtlServiceFee_Click()
    On Error GoTo EH
    If lvwExpenseFees.ListItems.Count > 0 Then
        If lvwExpenseFees.SelectedItem.Index > 1 Then
            lvwExpenseFees.ListItems(1).Selected = True
            txtTtlServiceFee.SetFocus
        End If
    End If
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub txtTtlServiceFee_Click"
End Sub

Private Sub txtTtlServiceFee_GotFocus()
    goUtil.utSelText txtTtlServiceFee
    SelectedNextItmX lvwExpenseFees, True
End Sub

'Private Sub LooseOverridesFocus()
'    On Error GoTo EH
'    If txtOverrideFeeItemComment.Visible Then
'        Exit Sub
'    End If
'    If mlIgnoreMouseMove > 0 Then
'        mlIgnoreMouseMove = mlIgnoreMouseMove - 1
'    End If
'
'    If Not CBool(mlIgnoreMouseMove) Then
''        framOverrideFees.Height = ORFEES_FRAM_HEIGHT_LOSTFOCUS
''        lvwOverrideFees.Height = ORFEES_LVW_HEIGHT_LOSTFOCUS
''        txtServiceFeeComment.Visible = True
'    End If
'
'    Exit Sub
'EH:
'    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub LooseOverridesFocus"
'End Sub

Private Function ZeroOutFeeList(pFeeList As ListView) As Boolean
    On Error GoTo EH
    Dim sComment As String
    Dim lNumberOfItems As Long
    Dim cAmount As Currency
    Dim sfsftName As String
    Dim sfsftDescription As String
    Dim sFlagText As String
    Dim itmX As ListItem
    
    'this function zeroes out all fee list except for the overriding fees
    'Since this will be called if an overriding fee is selected
    
    
    For Each itmX In pFeeList.ListItems
        If pFeeList.Name = lvwServiceFees.Name Then
            sComment = itmX.ListSubItems(GuiServiceFeesListView.Comment - 1)
            lNumberOfItems = itmX.ListSubItems(GuiServiceFeesListView.NumberOfItems - 1)
            cAmount = CCur(itmX.ListSubItems(GuiServiceFeesListView.Amount - 1))
            sfsftName = itmX.ListSubItems(GuiServiceFeesListView.fsftName - 1)
            sfsftDescription = itmX.Text
        ElseIf pFeeList.Name = lvwExpenseFees.Name Then
            sComment = itmX.ListSubItems(GuiExpenseFeesListView.Comment - 1)
            lNumberOfItems = itmX.ListSubItems(GuiExpenseFeesListView.NumberOfItems - 1)
            cAmount = CCur(itmX.ListSubItems(GuiExpenseFeesListView.Amount - 1))
            sfsftName = itmX.ListSubItems(GuiExpenseFeesListView.fsftName - 1)
            sfsftDescription = itmX.Text
        Else
            Exit Function
        End If
        
        If lNumberOfItems > 0 Or cAmount > 0 Or sComment <> vbNullString Then
            If pFeeList.Name = lvwServiceFees.Name Then
                ShowEditServiceFeesItem itmX
                txtAddServiceFeeItemComment.Text = vbNullString
                lstAddServiceFeeItemNumberOfItems.Text = "0"
                txtAddServiceFeeItemAmount.Text = "0.00"
                CalcAddServiceFeeItem
            ElseIf pFeeList.Name = lvwExpenseFees.Name Then
                ShowEditExpenseFeesItem itmX
                txtAddExpenseFeeItemComment.Text = vbNullString
                lstAddExpenseFeeItemNumberOfItems.Text = "0"
                txtAddExpenseFeeItemAmount.Text = "0.00"
                CalcAddExpenseFeeItem
            End If
        End If
    Next
    
    If pFeeList.Name = lvwServiceFees.Name Then
        txtMiscServiceFeeComment.Text = vbNullString
        txtMiscServiceFee.Text = "0.00"
        txtTtlServiceFee.Text = "0.00"
    ElseIf pFeeList.Name = lvwExpenseFees.Name Then
        txtMiscExpenseFeeComment.Text = vbNullString
        txtMiscExpenseFee.Text = "0.00"
        txtTtlExpenses.Text = "0.00"
    End If
    
    ZeroOutFeeList = True
    
    'cleanup
    Set itmX = Nothing
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Function ZeroOutFeeList"
End Function

Private Function CalcFeeSched() As Boolean
    On Error GoTo EH
    Dim cCurrentServieFee As Currency
    Dim lBillingCountID As Long
    Dim RSFeeSched As ADODB.Recordset
    Dim sServiceFeeComment As String
    
    lBillingCountID = CLng(msBillingCountID)
    
    If ChkOverrideAmounts.Value = vbUnchecked Then
        cCurrentServieFee = mfrmClaim.GetCurrentServiceFee(lBillingCountID, msFeeScheduleID)
    Else
        cCurrentServieFee = mfrmClaim.GetCurrentServiceFee(lBillingCountID, msFeeScheduleID, True, mcOverrideGrossLoss)
    End If
    
    txtServiceFee.Text = Format(cCurrentServieFee, "#,###,###,##0.00")
    
    SumServiceFees
    
    'Set the Service Fee Comment to the Selected FeeSchedule and Time Calculated
    moGUI.SetadoRSFeeSchedule msFeeScheduleID
    Set RSFeeSched = moGUI.adoFeeSchedule
    
    If RSFeeSched.RecordCount = 0 Then
        GoTo CLEAN_UP
    End If
    
    RSFeeSched.MoveFirst
    sServiceFeeComment = goUtil.IsNullIsVbNullString(RSFeeSched.Fields("ScheduleName"))
    sServiceFeeComment = sServiceFeeComment & " " & Now()
    txtServiceFeeComment.Text = sServiceFeeComment
    txtServiceFeeComment.ToolTipText = sServiceFeeComment
    
    CalcFeeSched = True
CLEAN_UP:
    'cleanup
    Set RSFeeSched = Nothing
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Function CalcFeeSched"
End Function

Private Sub SumServiceFees()
    On Error GoTo EH
    Dim itmX As ListItem
    Dim cAmount As Currency
    Dim cServiceFeesSubTotal As Currency
    
    'Get the main Service Fee
    If IsNumeric(txtServiceFee.Text) Then
        cServiceFeesSubTotal = CCur(txtServiceFee.Text)
    Else
        cServiceFeesSubTotal = 0
    End If
    
    'Add Additional Service Fees
    For Each itmX In lvwServiceFees.ListItems
        cAmount = CCur(itmX.ListSubItems(GuiServiceFeesListView.Amount - 1))
        cServiceFeesSubTotal = cServiceFeesSubTotal + cAmount
    Next
    
    Set itmX = Nothing
    
    'Add misc Service fee
    If IsNumeric(txtMiscServiceFee.Text) Then
        cAmount = CCur(txtMiscServiceFee.Text)
    Else
        cAmount = 0
    End If
    
    cServiceFeesSubTotal = cServiceFeesSubTotal + cAmount
    
    txtTtlServiceFee.Text = Format(cServiceFeesSubTotal, "#,###,###,##0.00")
    
    
    SumExpenseFees
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub SumServiceFees"
End Sub

Private Sub SumExpenseFees()
    On Error GoTo EH
    Dim itmX As ListItem
    Dim cAmount As Currency
    Dim cExpenseFeesSubTotal As Currency
    
    
    cExpenseFeesSubTotal = 0
        
    'Add Additional Service Fees
    For Each itmX In lvwExpenseFees.ListItems
        cAmount = CCur(itmX.ListSubItems(GuiExpenseFeesListView.Amount - 1))
        cExpenseFeesSubTotal = cExpenseFeesSubTotal + cAmount
    Next
    
    Set itmX = Nothing
    
    'Add misc Expense fee
    If IsNumeric(txtMiscExpenseFee.Text) Then
        cAmount = CCur(txtMiscExpenseFee.Text)
    Else
        cAmount = 0
    End If
    
    cExpenseFeesSubTotal = cExpenseFeesSubTotal + cAmount
    
    txtTtlExpenses.Text = Format(cExpenseFeesSubTotal, "#,###,###,##0.00")
    
    
    SumTtlExpenseAndServiceFees
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub SumServiceFees"
End Sub

Private Sub SumTtlExpenseAndServiceFees()
    On Error GoTo EH
    Dim cServiceFees As Currency
    Dim cExpenseFees As Currency
    Dim cTtlServiceExp As Currency
    
    'Get the Service Fees Subtotal
    If IsNumeric(txtTtlServiceFee.Text) Then
        cServiceFees = CCur(txtTtlServiceFee.Text)
    Else
        cServiceFees = 0
    End If
    
    'Get the Expense Fees Subtotal
    If IsNumeric(txtTtlExpenses.Text) Then
        cExpenseFees = CCur(txtTtlExpenses.Text)
    Else
        cExpenseFees = 0
    End If
    
    cTtlServiceExp = cServiceFees + cExpenseFees
    
    txtTtlServiceExp.Text = Format(cTtlServiceExp, "#,###,###,##0.00")
    
    SumTax
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub SumTtlExpenseAndServiceFees"
End Sub

Private Sub SumTax()
    On Error GoTo EH
    Dim cTtlServiceExp As Currency
    Dim dblTaxPercent As Double
    Dim cTaxAmount As Currency
    
    'Get the Service and Expense Fees Subtotal
    If IsNumeric(txtTtlServiceExp.Text) Then
        cTtlServiceExp = CCur(txtTtlServiceExp.Text)
    Else
        cTtlServiceExp = 0
    End If
    
    'Get the TaxPercent
    If IsNumeric(txtTaxPercent.Text) Then
        dblTaxPercent = CDbl(txtTaxPercent.Text)
    Else
        dblTaxPercent = 0
    End If
    
    cTaxAmount = cTtlServiceExp * (dblTaxPercent / 100)
    
    txtTaxAmount.Text = Format(cTaxAmount, "#,###,###,##0.00")
    
    SumTtlInvoice
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub SumTax"
End Sub

Private Sub SumTtlInvoice()
    On Error GoTo EH
    Dim cTtlServiceExp As Currency
    Dim cTaxAmount As Currency
    Dim cTotalAdjustingFee As Currency
    
   'Get the Service and Expense Fees Subtotal
    If IsNumeric(txtTtlServiceExp.Text) Then
        cTtlServiceExp = CCur(txtTtlServiceExp.Text)
    Else
        cTtlServiceExp = 0
    End If
    
    'Get Tax Amount
    If IsNumeric(txtTaxAmount.Text) Then
        cTaxAmount = CCur(txtTaxAmount.Text)
    Else
        cTaxAmount = 0
    End If
    
    cTotalAdjustingFee = cTtlServiceExp + cTaxAmount
    
    txtTotalAdjustingFee.Text = Format(cTotalAdjustingFee, "#,###,###,##0.00")
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub SumTtlInvoice"
End Sub

Public Sub UpdateFkeyBillingCountID(psIDBillingCount As String)
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    Dim lRecordsAffected As Long
    
    Set oConn = New ADODB.Connection
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    'Update Photo Table Fkey
    sSQL = "UPDATE RTPhotoLog Set "
    sSQL = sSQL & "[BillingCountID] = " & psIDBillingCount & ", "
    sSQL = sSQL & "[IDBillingCount] = " & psIDBillingCount & ", "
    sSQL = sSQL & "[UpLoadMe] = True, "
    sSQL = sSQL & "[DateLastUpdated] = #" & Now() & "#, "
    sSQL = sSQL & "[UpdateByUserID]  = " & goUtil.gsCurUsersID & " "
    sSQL = sSQL & "WHERE IDAssignments = " & msAssignmentsID & " "
    sSQL = sSQL & "AND IDBillingCount Is Null "
    
    oConn.Execute sSQL, lRecordsAffected
    Sleep 100
    
    sSQL = "UPDATE RTActivityLog Set "
    sSQL = sSQL & "[BillingCountID] = " & psIDBillingCount & ", "
    sSQL = sSQL & "[IDBillingCount] = " & psIDBillingCount & ", "
    sSQL = sSQL & "[UpLoadMe] = True, "
    sSQL = sSQL & "[DateLastUpdated] = #" & Now() & "#, "
    sSQL = sSQL & "[UpdateByUserID]  = " & goUtil.gsCurUsersID & " "
    sSQL = sSQL & "WHERE IDAssignments = " & msAssignmentsID & " "
    sSQL = sSQL & "AND IDBillingCount Is Null "
    
    oConn.Execute sSQL, lRecordsAffected
    Sleep 100
    
    sSQL = "UPDATE RTChecks Set "
    sSQL = sSQL & "[BillingCountID] = " & psIDBillingCount & ", "
    sSQL = sSQL & "[IDBillingCount] = " & psIDBillingCount & ", "
    sSQL = sSQL & "[UpLoadMe] = True, "
    sSQL = sSQL & "[DateLastUpdated] = #" & Now() & "#, "
    sSQL = sSQL & "[UpdateByUserID]  = " & goUtil.gsCurUsersID & " "
    sSQL = sSQL & "WHERE IDAssignments = " & msAssignmentsID & " "
    sSQL = sSQL & "AND IDBillingCount Is Null "
    
    oConn.Execute sSQL, lRecordsAffected
    Sleep 100
    
    'cleanup
    Set oConn = Nothing
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub UpdateFkeyBillingCountID"
End Sub
