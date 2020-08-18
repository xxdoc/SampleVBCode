VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmActivityLog 
   AutoRedraw      =   -1  'True
   Caption         =   "Activity Log"
   ClientHeight    =   6705
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
   ScaleHeight     =   6705
   ScaleWidth      =   11820
   StartUpPosition =   3  'Windows Default
   Tag             =   "Activity Log"
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
      TabIndex        =   44
      Top             =   5400
      Width           =   4455
      Begin VB.CommandButton cmdSpelling 
         Caption         =   "Spellin&g"
         Height          =   855
         Left            =   120
         MaskColor       =   &H00000000&
         Picture         =   "frmActivityLog.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   45
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
         Picture         =   "frmActivityLog.frx":044A
         Style           =   1  'Graphical
         TabIndex        =   47
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
         Picture         =   "frmActivityLog.frx":0594
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Exit"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdPrintList 
         Caption         =   "Sho&w Item List"
         Height          =   855
         Left            =   1200
         MaskColor       =   &H00000000&
         Picture         =   "frmActivityLog.frx":089E
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Exit"
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame framActivityLog 
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   11580
      Begin VB.CommandButton cmdFindNext 
         Caption         =   "Find &Next"
         Height          =   375
         Left            =   4200
         TabIndex        =   34
         Top             =   240
         Width           =   1100
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Find"
         Height          =   375
         Left            =   3120
         TabIndex        =   33
         Top             =   240
         Width           =   1100
      End
      Begin VB.CommandButton cmdSelAll 
         Caption         =   "&Select A&ll"
         Height          =   375
         Left            =   2040
         TabIndex        =   32
         Top             =   240
         Width           =   1100
      End
      Begin VB.CommandButton cmdEditActLog 
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
         TabIndex        =   51
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdAddActLog 
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
         TabIndex        =   50
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtSpellMe 
         Height          =   1335
         Left            =   480
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   2520
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.Frame framActivityLogMaint 
         Caption         =   "Activity Log Maintenance"
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
         TabIndex        =   36
         Top             =   4440
         Width           =   11355
         Begin VB.CommandButton cmdPrintActLog 
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
            TabIndex        =   38
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmdRefreshActLog 
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
            TabIndex        =   37
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmdDelActLog 
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
            TabIndex        =   40
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
            Left            =   8640
            Style           =   1  'Graphical
            TabIndex        =   39
            Top             =   240
            Visible         =   0   'False
            Width           =   1100
         End
      End
      Begin MSComctlLib.ImageList imgActLogStatus 
         Left            =   10920
         Top             =   1680
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483633
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   16777215
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmActivityLog.frx":0CE0
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmActivityLog.frx":0E3A
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lstvActLog 
         Height          =   3735
         Left            =   120
         TabIndex        =   35
         Tag             =   "Enable"
         ToolTipText     =   "Right Click for Menu"
         Top             =   720
         Width           =   11355
         _ExtentX        =   20029
         _ExtentY        =   6588
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
         ColHdrIcons     =   "imgActLogStatus"
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
      Begin VB.Frame framAL03_sExplainedRCV 
         Appearance      =   0  'Flat
         Caption         =   "3. EXPLAINED && GAVE RCV FORM TO CUSTOMER?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   3675
         Visible         =   0   'False
         Width           =   4935
         Begin VB.OptionButton optAL03_sExplainedRCV 
            Appearance      =   0  'Flat
            Caption         =   "N/A"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   4080
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   200
            Width           =   735
         End
         Begin VB.OptionButton optAL03_sExplainedRCV 
            Appearance      =   0  'Flat
            Caption         =   "NO"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   3360
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   200
            Width           =   735
         End
         Begin VB.OptionButton optAL03_sExplainedRCV 
            Appearance      =   0  'Flat
            Caption         =   "YES"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   2640
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   200
            Width           =   735
         End
      End
      Begin VB.Frame framAL02_sExplainedEstimate 
         Appearance      =   0  'Flat
         Caption         =   "2. EXPLAINED && GAVE ESTIMATE TO CUSTOMER?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   3195
         Visible         =   0   'False
         Width           =   4935
         Begin VB.OptionButton optAL02_sExplainedEstimate 
            Appearance      =   0  'Flat
            Caption         =   "NO"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   3360
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   200
            Width           =   735
         End
         Begin VB.OptionButton optAL02_sExplainedEstimate 
            Appearance      =   0  'Flat
            Caption         =   "YES"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   2640
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   200
            Width           =   735
         End
      End
      Begin VB.Frame framAL01_sPresentDurringInspection 
         Appearance      =   0  'Flat
         Caption         =   "1. DOES INSURED WANT TO BE PRESENT FOR INSPECTION?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   2700
         Visible         =   0   'False
         Width           =   4935
         Begin VB.OptionButton optAL01_sPresentDurringInspection 
            Appearance      =   0  'Flat
            Caption         =   "NO"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   3360
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   200
            Width           =   735
         End
         Begin VB.OptionButton optAL01_sPresentDurringInspection 
            Appearance      =   0  'Flat
            Caption         =   "YES"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   2640
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   200
            Width           =   735
         End
      End
      Begin VB.Frame framAL06_sConfirmedCoverage 
         Appearance      =   0  'Flat
         Caption         =   "6. IS COVERAGE OK?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   5160
         TabIndex        =   19
         Top             =   3675
         Visible         =   0   'False
         Width           =   3015
         Begin VB.OptionButton optAL06_sConfirmedCoverage 
            Appearance      =   0  'Flat
            Caption         =   "NO"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   1440
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   200
            Width           =   735
         End
         Begin VB.OptionButton optAL06_sConfirmedCoverage 
            Appearance      =   0  'Flat
            Caption         =   "YES"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   720
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   200
            Width           =   735
         End
      End
      Begin VB.Frame framAL05_sExplainedMortgageeChecks 
         Appearance      =   0  'Flat
         Caption         =   "5. EXPLAINED MORTGAGE CHECKS?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   5160
         TabIndex        =   15
         Top             =   3195
         Visible         =   0   'False
         Width           =   3015
         Begin VB.OptionButton optAL05_sExplainedMortgageeChecks 
            Appearance      =   0  'Flat
            Caption         =   "N/A"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   2160
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   200
            Width           =   735
         End
         Begin VB.OptionButton optAL05_sExplainedMortgageeChecks 
            Appearance      =   0  'Flat
            Caption         =   "NO"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   1440
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   200
            Width           =   735
         End
         Begin VB.OptionButton optAL05_sExplainedMortgageeChecks 
            Appearance      =   0  'Flat
            Caption         =   "YES"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   720
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   200
            Width           =   735
         End
      End
      Begin VB.Frame framAL04_sConfirmMortgageeIsCorrect 
         Appearance      =   0  'Flat
         Caption         =   "4. CONFIRMED MORTGAGEE...?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   5160
         TabIndex        =   11
         Top             =   2700
         Visible         =   0   'False
         Width           =   3015
         Begin VB.OptionButton optAL04_sConfirmMortgageeIsCorrect 
            Appearance      =   0  'Flat
            Caption         =   "N/A"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   2160
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   200
            Width           =   735
         End
         Begin VB.OptionButton optAL04_sConfirmMortgageeIsCorrect 
            Appearance      =   0  'Flat
            Caption         =   "NO"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   1440
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   200
            Width           =   735
         End
         Begin VB.OptionButton optAL04_sConfirmMortgageeIsCorrect 
            Appearance      =   0  'Flat
            Caption         =   "YES"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   720
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   200
            Width           =   735
         End
      End
      Begin VB.Frame framAL09_sSubrogation 
         Appearance      =   0  'Flat
         Caption         =   "9. SUBROGATION?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   8280
         TabIndex        =   29
         Top             =   3675
         Visible         =   0   'False
         Width           =   3135
         Begin VB.OptionButton optAL09_sSubrogation 
            Appearance      =   0  'Flat
            Caption         =   "NO"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   1560
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   200
            Width           =   735
         End
         Begin VB.OptionButton optAL09_sSubrogation 
            Appearance      =   0  'Flat
            Caption         =   "YES"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   840
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   200
            Width           =   735
         End
      End
      Begin VB.Frame framAL08_sSalvage 
         Appearance      =   0  'Flat
         Caption         =   "8. SALVAGE?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   8280
         TabIndex        =   26
         Top             =   3195
         Visible         =   0   'False
         Width           =   3135
         Begin VB.OptionButton optAL08_sSalvage 
            Appearance      =   0  'Flat
            Caption         =   "NO"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   1560
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   200
            Width           =   735
         End
         Begin VB.OptionButton optAL08_sSalvage 
            Appearance      =   0  'Flat
            Caption         =   "YES"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   840
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   200
            Width           =   735
         End
      End
      Begin VB.Frame framAL07_sPriorLoss 
         Appearance      =   0  'Flat
         Caption         =   "7. APPLICABLE PRIOR LOSSES?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   8280
         TabIndex        =   22
         Top             =   2700
         Visible         =   0   'False
         Width           =   3135
         Begin VB.OptionButton optAL07_sPriorLoss 
            Appearance      =   0  'Flat
            Caption         =   "N/A"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   2280
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   200
            Width           =   735
         End
         Begin VB.OptionButton optAL07_sPriorLoss 
            Appearance      =   0  'Flat
            Caption         =   "NO"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   1560
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   200
            Width           =   735
         End
         Begin VB.OptionButton optAL07_sPriorLoss 
            Appearance      =   0  'Flat
            Caption         =   "YES"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   840
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   200
            Width           =   735
         End
      End
   End
   Begin VB.Frame framAssocBillingID 
      Caption         =   "Associate IB (Internal Billing) to selected Items:"
      Height          =   1215
      Left            =   120
      TabIndex        =   41
      Top             =   5400
      Width           =   6735
      Begin VB.CommandButton cmdAssocBillingID 
         Caption         =   "Associate I&B"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5280
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   713
         Width           =   1335
      End
      Begin VB.ComboBox cboAssocBillingID 
         Height          =   360
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   720
         Width           =   5055
      End
   End
   Begin VB.Menu PopUpmnuActLog 
      Caption         =   "PopUpActLog"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuEditActLog 
         Caption         =   "&Edit"
      End
      Begin VB.Menu mnuDeleteActLog 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuSelectAllActLog 
         Caption         =   "&Select All"
      End
   End
End
Attribute VB_Name = "frmActivityLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mfrmClaim As frmClaim
Private mbUnloadMe As Boolean
Private msAssignmentsID As String
Private moGUI As V2ECKeyBoard.clsCarGUI
Private mitmXSelected As ListItem 'Currently selected Photo Item
Private mbLoading As Boolean 'Loading Form
Private mbLoadingMe As Boolean 'Just loading data
Private msFindText As String
Private mlLastFindIndex As Long
Private mArv As V2ARViewer.clsARViewer
Private moForm As Form


Public Property Let itmXSelected(pitmX As ListItem)
    On Error GoTo EH
    Set mitmXSelected = pitmX
    If Not mitmXSelected Is Nothing Then
        cmdEditActLog.Enabled = True
    Else
        cmdEditActLog.Enabled = False
    End If
    Exit Property
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Property Let itmXSelected"
End Property
Public Property Set itmXSelected(pitmX As ListItem)
    On Error GoTo EH
    Set mitmXSelected = pitmX
    If Not mitmXSelected Is Nothing Then
        cmdEditActLog.Enabled = True
    Else
        cmdEditActLog.Enabled = False
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

Private Sub cmdAddActLog_Click()
    On Error GoTo EH
    cmdAddActLog.Enabled = False
    With AddActLog
        .MyActivityLog = Me
        .MyfrmClaim = Me.MyfrmClaim
        .Adding = True
        .AssignmentsID = msAssignmentsID
         Load AddActLog
        .Caption = "Add Activity Log"
        .txtActDate.Text = Format(Now(), "MM/DD/YYYY")
        .timeActTime.ecsTime = Format(Now, "HH:MM")
        .txtServiceTime.Text = "0.00"
        .txtActText = vbNullString
        .cmdSave.Enabled = False
        .Show vbModal
    End With
    
    'Clean Up
    Unload AddActLog
    Set AddActLog = Nothing
    cmdAddActLog.Enabled = True
    
    If lstvActLog.Visible Then
        lstvActLog.SetFocus
    End If
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdAddActLog_Click"
End Sub

Private Sub cmdAssocBillingID_Click()
    On Error GoTo EH
    Dim sMess As String
    Dim sIBDesc As String
    Dim sTitle As String
    Dim bItemSelected As Boolean
    Dim itmX As MSComctlLib.ListItem
    Dim sActLogID As String
    Dim vActLogID As Variant
    Dim colActLogID As Collection
    
    
    If lstvActLog.ListItems.Count > 0 Then
        
        If cboAssocBillingID.ListIndex = -1 Then
            sMess = "You must select an IB from the Drop down List!"
            MsgBox sMess, vbExclamation + vbOKOnly, "Select an IB."
            cmdAssocBillingID.Enabled = False
            Exit Sub
        End If
        
        For Each itmX In lstvActLog.ListItems
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
        
        lstvActLog.Visible = False
        Set colActLogID = New Collection
        For Each itmX In lstvActLog.ListItems
            If itmX.Selected Then
                colActLogID.Add itmX.SubItems(GuiActLogListView.ID - 1), itmX.SubItems(GuiActLogListView.ID - 1)
            End If
        Next
        For Each vActLogID In colActLogID
            sActLogID = vActLogID
            If Not AssocActLogItemToBillingID(sActLogID) Then
                Exit Sub
            End If
        Next
    End If
    
    RefreshActLog
    
    lstvActLog.Visible = True
    cmdAssocBillingID.Enabled = True
    
    Set itmX = Nothing
    Set colActLogID = Nothing
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdAssocBillingID_Click"
End Sub

Public Function AssocActLogItemToBillingID(psID As String) As Boolean
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
    
    sSQL = "UPDATE RTActivityLog SET "
    sSQL = sSQL & "[BillingCountID] = " & sIDBillingCount & ", "
    sSQL = sSQL & "[IDBillingCount] = " & sIDBillingCount & ", "
    sSQL = sSQL & "[UpLoadMe] = True, "
    sSQL = sSQL & "[DateLastUpdated] = #" & Format(Now(), "MM/DD/YYYY HH:MM:SS") & "#, "
    sSQL = sSQL & "[UpdateByUserID] = " & goUtil.gsCurUsersID & " "
    sSQL = sSQL & "WHERE [ID] = " & psID & " "
    sSQL = sSQL & "AND IDAssignments = " & msAssignmentsID & " "

    oConn.Execute sSQL
    
    AssocActLogItemToBillingID = True
    
    'clean up
    Set oConn = Nothing
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function AssocActLogItemToBillingID"
End Function


Private Sub cmdDelActLog_Click()
    On Error GoTo EH
    Dim itmX As MSComctlLib.ListItem
    Dim sActLogID As String
    Dim vActLogID As Variant
    Dim colActLogID As Collection
    
    
    If lstvActLog.ListItems.Count > 0 Then
        If MsgBox("Are you sure ?", vbYesNo, "DELETE SELECTED LOG ITEMS") = vbYes Then
            lstvActLog.Visible = False
            Set colActLogID = New Collection
            For Each itmX In lstvActLog.ListItems
                If itmX.Selected Then
                    colActLogID.Add itmX.SubItems(GuiActLogListView.ID - 1), itmX.SubItems(GuiActLogListView.ID - 1)
                End If
            Next
            For Each vActLogID In colActLogID
                sActLogID = vActLogID
                If DeleteActLogItem(sActLogID) Then
                    lstvActLog.ListItems.Remove ("""" & sActLogID & """")
                End If
            Next
        End If
    End If
    
    lstvActLog.Visible = True
    Set itmX = Nothing
    Set colActLogID = Nothing
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdDelActLog_Click"
End Sub

Private Sub cmdEditActLog_Click()
    On Error GoTo EH
    
    cmdEditActLog.Enabled = False
    
    EditActLog
    
    cmdEditActLog.Enabled = True
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdEditActLog_Click"
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
    If lstvActLog.ListItems.Count > 0 Then
        mlLastFindIndex = 0
        goUtil.utFindListItem Me, lstvActLog, msFindText, mlLastFindIndex
    End If
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdFind_Click"
End Sub

Private Sub cmdFindNext_Click()
    On Error GoTo EH
    
    If msFindText <> vbNullString And lstvActLog.ListItems.Count > 0 Then
        goUtil.utFindListItem Me, lstvActLog, msFindText, mlLastFindIndex
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdFindNext_Click"
End Sub

Private Sub cmdPrintActLog_Click()
    On Error GoTo EH
    Screen.MousePointer = vbHourglass
    
    'Save stuff first if applicable
    If cmdSave.Enabled Then
        If SaveMe Then
            mfrmClaim.RefreshMe
            cmdSave.Enabled = False
        Else
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    End If
    
    cmdPrintActLog.Enabled = False
    If PrintActLog(msAssignmentsID) Then
        If Not mbUnloadMe Then
            cmdPrintActLog.Enabled = True
        End If
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
EH:
    Screen.MousePointer = vbDefault
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdPrintActLog_Click"
End Sub

Private Sub cmdPrintList_Click()
    On Error GoTo EH
    goUtil.utPrintListView App.EXEName, lstvActLog, "Activity Log"
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdPrintList_Click"
End Sub

Private Sub cmdRefreshActLog_Click()
    On Error GoTo EH
    cmdRefreshActLog.Enabled = False
    Screen.MousePointer = vbHourglass
    RefreshActLog
    Screen.MousePointer = vbDefault
    cmdRefreshActLog.Enabled = True
Exit Sub
EH:
    Screen.MousePointer = vbDefault
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdRefreshActLog_Click"
End Sub

Public Sub RefreshActLog()
    LoadMe
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
    For Each itmX In lstvActLog.ListItems
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
    Dim udtActLog As GuiActLogItem
    Dim sDelim As String
    sDelim = Chr(160) & vbCrLf & Chr(160)
    
    txtSpellMe.Text = vbNullString
    
    For Each itmX In lstvActLog.ListItems
        sText = sText & itmX.SubItems(GuiActLogListView.ActText - 1) & sDelim
    Next
    'take off the last Delim
    If sText <> sDelim Then
        sText = left(sText, InStrRev(sText, sDelim, , vbBinaryCompare) - 1)
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
    saryText() = Split(sText, sDelim, , vbBinaryCompare)
    
    'check the spelling against the List view...
    'if any changes then need to save those changes to the db
    For lPos = LBound(saryText, 1) To UBound(saryText, 1)
        sText = saryText(lPos)
        Set itmX = lstvActLog.ListItems(lPos + 1)
        
        If StrComp(sText, itmX.SubItems(GuiActLogListView.ActText - 1), vbTextCompare) <> 0 Then
            With udtActLog
                .RTActivityLogID = itmX.SubItems(GuiActLogListView.RTActivityLogID - 1)
                .AssignmentsID = itmX.SubItems(GuiActLogListView.AssignmentsID - 1)
                .BillingCountID = itmX.SubItems(GuiActLogListView.BillingCountID - 1)
                .ID = itmX.SubItems(GuiActLogListView.ID - 1)
                .IDAssignments = itmX.SubItems(GuiActLogListView.IDAssignments - 1)
                .IDBillingCount = itmX.SubItems(GuiActLogListView.IDBillingCount - 1)
                .ActDate = itmX.Text
                'Set the text to the corrected text
                itmX.SubItems(GuiActLogListView.ActText - 1) = sText
                .ActText = sText
                .ActTime = itmX.Text & " " & itmX.SubItems(GuiActLogListView.ActTime - 1)
                .BlankPageAfter = itmX.SubItems(GuiActLogListView.BlankPageAfter - 1)
                .BlankRowsAfter = itmX.SubItems(GuiActLogListView.BlankRowsAfter - 1)
                .IsMgrEntry = itmX.SubItems(GuiActLogListView.IsMgrEntry - 1)
                .PageBreakAfter = itmX.SubItems(GuiActLogListView.PageBreakAfter - 1)
                .ServiceTime = itmX.SubItems(GuiActLogListView.ServiceTime - 1)
                .IsDeleted = goUtil.GetFlagFromText(itmX.SubItems(GuiActLogListView.IsDeleted - 1))
                .DownLoadMe = goUtil.GetFlagFromText(itmX.SubItems(GuiActLogListView.DownLoadMe - 1))
                .UpLoadMe = "True"
                .AdminComments = itmX.SubItems(GuiActLogListView.AdminComments - 1)
                .DateLastUpdated = Format(Now(), "MM/DD/YYYY HH:MM:SS")
                itmX.SubItems(GuiActLogListView.DateLastUpdated - 1) = .DateLastUpdated
                itmX.SubItems(GuiActLogListView.DateLastUpdatedSort - 1) = Format(.DateLastUpdated, "YYYY/MM/DD HH:MM:SS")
                .UpdateByUserID = goUtil.gsCurUsersID
            End With
            EditActLogItem udtActLog
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
    Me.Icon = mfrmClaim.optClaim(GuiClaimOptions.opt02_ActivityLog).Picture
    
    LoadHeaderlstvActLog
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
    
    mbLoadingMe = True

    If Not mfrmClaim.SetadoRSRTActivityLogInfo(msAssignmentsID) Then
        Exit Function
    End If
    PopulateActLogInfo
    
    If Not mfrmClaim.SetadoRSRTActivityLog(msAssignmentsID) Then
        Exit Function
    End If
    PopulatelstvActLog
    
    'Load Billing RS
    mfrmClaim.SetadoRSBillingCount msAssignmentsID, True
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
    Dim udtActLofInfo As GuiActLogInfoItem
       
    'Need to Edit the Act Log Info
    With udtActLofInfo
        .AssignmentsID = msAssignmentsID
        .IDAssignments = msAssignmentsID
        .AL01_sPresentDurringInspection = CStr(optAL01_sPresentDurringInspection(GuiOptValues.YES).Value)
        .AL02_sExplainedEstimate = CStr(optAL02_sExplainedEstimate(GuiOptValues.YES).Value)
        .AL03_sExplainedRCV = CStr(optAL03_sExplainedRCV(GuiOptValues.YES).Value)
        .AL03_sExplainedRCVNA = CStr(optAL03_sExplainedRCV(GuiOptValues.NA).Value)
        .AL04_sConfirmMortgageeIsCorrect = CStr(optAL04_sConfirmMortgageeIsCorrect(GuiOptValues.YES).Value)
        .AL04_sConfirmMortgageeIsCorrectNA = CStr(optAL04_sConfirmMortgageeIsCorrect(GuiOptValues.NA).Value)
        .AL05_sExplainedMortgageeChecks = CStr(optAL05_sExplainedMortgageeChecks(GuiOptValues.YES).Value)
        .AL05_sExplainedMortgageeChecksNA = CStr(optAL05_sExplainedMortgageeChecks(GuiOptValues.NA).Value)
        .AL06_sConfirmedCoverage = CStr(optAL06_sConfirmedCoverage(GuiOptValues.YES).Value)
        .AL07_sPriorLoss = CStr(optAL07_sPriorLoss(GuiOptValues.YES).Value)
        .AL07_sPriorLossNA = CStr(optAL07_sPriorLoss(GuiOptValues.NA).Value)
        .AL08_sSalvage = CStr(optAL08_sSalvage(GuiOptValues.YES).Value)
        .AL09_sSubrogation = CStr(optAL09_sSubrogation(GuiOptValues.YES).Value)
         'IsDeleted
        .IsDeleted = "False"
        'DownLoadMe
        .DownLoadMe = "False"
        'UpLoadMe
        .UpLoadMe = "True"
        'AdminComments
        .AdminComments = vbNullString
        'DateLastUpdated
        .DateLastUpdated = Format(Now(), "MM/DD/YYYY HH:MM:SS")
        'UpdateByUserID
        .UpdateByUserID = goUtil.gsCurUsersID
    End With
    
    EditActivityLogInfoItem udtActLofInfo
    
    SaveMe = True
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function SaveMe"
End Function

Public Function CheckStatus() As Boolean
    On Error GoTo EH
    
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
    framActivityLog.Width = Me.Width - 360
    lstvActLog.Width = framActivityLog.Width - 225
    framActivityLogMaint.Width = framActivityLog.Width - 225
    chkHideDeleted.left = framActivityLogMaint.Width - 2775
    cmdDelActLog.left = framActivityLogMaint.Width - 1575
    'framCommands
    framCommands.left = Me.Width - 4695
    
    'Heights and Tops
    framActivityLog.Height = Me.Height - 2100
    lstvActLog.Height = framActivityLog.Height - 1560
    framActivityLogMaint.top = framActivityLog.Height - 855
    
    'framAssocBillingID
    framAssocBillingID.top = Me.Height - 1995
    
    'framCommands
    framCommands.top = Me.Height - 1995
    
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


Public Sub LoadHeaderlstvActLog()
    On Error GoTo EH
    Dim bGridOn As Boolean
    Dim bHideDeleted As Boolean
    Dim bHideUploadFlags As Boolean
    
    bHideDeleted = CBool(GetSetting(App.EXEName, "GENERAL", "HIDE_DELETED", True))
    bHideUploadFlags = CBool(GetSetting(App.EXEName, "GENERAL", "HIDE_UPLOAD_FLAGS", True))
    
    'set the columnheaders
    With lstvActLog
        .ColumnHeaders.Add , "ActDate", "Date"
        .ColumnHeaders.Add , "ActDateSort", "Sort Date" 'hidden
        .ColumnHeaders.Add , "ActTime", "Time"
        .ColumnHeaders.Add , "ActTimeSort", "Sort Time" 'hidden
        .ColumnHeaders.Add , "ActText", "Activity Text"
        .ColumnHeaders.Add , "ServiceTime", "Service Time"
        .ColumnHeaders.Add , "ServiceTimeSort", "Sort Service Time" 'hidden
        .ColumnHeaders.Add , "IB", "IB" ' Shows the IB this ActLog Items is Asscoiated with
        .ColumnHeaders.Add , "PageBreakAfter", "PB After" 'hidden until future Use
        .ColumnHeaders.Add , "BlankPageAfter", "BP After" 'hidden until future Use
        .ColumnHeaders.Add , "BlankRowsAfter", "BR After" 'hidden until future Use
        .ColumnHeaders.Add , "BlankRowsAfterSort", "Sort BR After" ' Hidden
        .ColumnHeaders.Add , "IsMgrEntry", "Is Mgr Entry" 'hidden until future Use
        .ColumnHeaders.Add , "IsDeleted", "Is Deleted" 'hidden
        .ColumnHeaders.Add , "UpLoadMe", "UpLoad Me" 'hidden
        .ColumnHeaders.Add , "DateLastUpdated", "Date Last Updated"
        .ColumnHeaders.Add , "DateLastUpdatedSort", "Sort Date Last Updated" ' hidden
        .ColumnHeaders.Add , "AdminComments", "Admin Comments" 'hidden
        .ColumnHeaders.Add , "ID", "ID" 'Hidden
        .ColumnHeaders.Add , "IDAssignments", "IDAssignments" 'Hidden
        .ColumnHeaders.Add , "RTActivityLogID", "RTActivityLogID" ' Hidden
        .ColumnHeaders.Add , "AssignmentsID", "AssignmentsID"  ' Hidden
        .ColumnHeaders.Add , "BillingCountID", "BillingCountID"  ' Hidden
        .ColumnHeaders.Add , "IDBillingCount", "IDBillingCount" ' Hidden
        .ColumnHeaders.Add , "DownLoadMe", "DownLoadMe"  ' hidden
        .ColumnHeaders.Add , "UpdateByUserID", "UpdateByUserID"  ' Hidden
        
        .Sorted = False
        .SortOrder = lvwAscending
        
        'ActDate
        .ColumnHeaders.Item(GuiActLogListView.ActDate).Width = 1335
        .ColumnHeaders.Item(GuiActLogListView.ActDate).Alignment = lvwColumnLeft
        'ActDateSort
        .ColumnHeaders.Item(GuiActLogListView.ActDateSort).Width = 0 ' Hidden
        .ColumnHeaders.Item(GuiActLogListView.ActDateSort).Alignment = lvwColumnLeft
        'ActTime
        .ColumnHeaders.Item(GuiActLogListView.ActTime).Width = 1335
        .ColumnHeaders.Item(GuiActLogListView.ActTime).Alignment = lvwColumnLeft
        'ActTimeSort
        .ColumnHeaders.Item(GuiActLogListView.ActTimeSort).Width = 0 ' Hidden
        .ColumnHeaders.Item(GuiActLogListView.ActTimeSort).Alignment = lvwColumnLeft
        'ActText
        .ColumnHeaders.Item(GuiActLogListView.ActText).Width = 8000
        .ColumnHeaders.Item(GuiActLogListView.ActText).Alignment = lvwColumnLeft
        'ServiceTime
        .ColumnHeaders.Item(GuiActLogListView.ServiceTime).Width = 1335
        .ColumnHeaders.Item(GuiActLogListView.ServiceTime).Alignment = lvwColumnRight
        'ServiceTimeSort
        .ColumnHeaders.Item(GuiActLogListView.ServiceTimeSort).Width = 0 'hidden
        .ColumnHeaders.Item(GuiActLogListView.ServiceTimeSort).Alignment = lvwColumnLeft
        'IB
        .ColumnHeaders.Item(GuiActLogListView.IB).Width = 750
        .ColumnHeaders.Item(GuiActLogListView.IB).Alignment = lvwColumnLeft
        'PageBreakAfter
        .ColumnHeaders.Item(GuiActLogListView.PageBreakAfter).Width = 0 'hidden
        .ColumnHeaders.Item(GuiActLogListView.PageBreakAfter).Alignment = lvwColumnLeft
        'BlankPageAfter
        .ColumnHeaders.Item(GuiActLogListView.BlankPageAfter).Width = 0 'hidden
        .ColumnHeaders.Item(GuiActLogListView.BlankPageAfter).Alignment = lvwColumnLeft
        'BlankRowsAfter
        .ColumnHeaders.Item(GuiActLogListView.BlankRowsAfter).Width = 0 'hidden
        .ColumnHeaders.Item(GuiActLogListView.BlankRowsAfter).Alignment = lvwColumnLeft
        'BlankRowsAfterSort
        .ColumnHeaders.Item(GuiActLogListView.BlankRowsAfterSort).Width = 0 'hidden
        .ColumnHeaders.Item(GuiActLogListView.BlankRowsAfterSort).Alignment = lvwColumnLeft
        'IsMgrEntry
        .ColumnHeaders.Item(GuiActLogListView.IsMgrEntry).Width = 0 'hidden
        .ColumnHeaders.Item(GuiActLogListView.IsMgrEntry).Alignment = lvwColumnLeft
        'Is Deleted
        If bHideDeleted Then
            .ColumnHeaders.Item(GuiActLogListView.IsDeleted).Width = 0 ' Hidden 400
        Else
            .ColumnHeaders.Item(GuiActLogListView.IsDeleted).Width = 400
        End If
        .ColumnHeaders.Item(GuiActLogListView.IsDeleted).Alignment = lvwColumnCenter
        .ColumnHeaders.Item(GuiActLogListView.IsDeleted).Icon = GuiActLogStatusList.IsDeleted
        'UpLoad Me
        If bHideUploadFlags Then
            .ColumnHeaders.Item(GuiActLogListView.UpLoadMe).Width = 0 ' Hidden 400
        Else
            .ColumnHeaders.Item(GuiActLogListView.UpLoadMe).Width = 400
        End If
        .ColumnHeaders.Item(GuiActLogListView.UpLoadMe).Alignment = lvwColumnCenter
        .ColumnHeaders.Item(GuiActLogListView.UpLoadMe).Icon = GuiActLogStatusList.UpLoadMe
        'DateLastUpdated
        .ColumnHeaders.Item(GuiActLogListView.DateLastUpdated).Width = 2200
        .ColumnHeaders.Item(GuiActLogListView.DateLastUpdated).Alignment = lvwColumnLeft
        'DateLastUpdatedSort
        .ColumnHeaders.Item(GuiActLogListView.DateLastUpdatedSort).Width = 0  'Hidden Sort Column
        .ColumnHeaders.Item(GuiActLogListView.DateLastUpdatedSort).Alignment = lvwColumnLeft
        'AdminComments
        .ColumnHeaders.Item(GuiActLogListView.AdminComments).Width = 0 ' Hidden 10000
        .ColumnHeaders.Item(GuiActLogListView.AdminComments).Alignment = lvwColumnLeft
        'ID
        .ColumnHeaders.Item(GuiActLogListView.ID).Width = 0  'hidden
        .ColumnHeaders.Item(GuiActLogListView.ID).Alignment = lvwColumnLeft
        'IDAssignments
        .ColumnHeaders.Item(GuiActLogListView.IDAssignments).Width = 0  'Hidden
        .ColumnHeaders.Item(GuiActLogListView.IDAssignments).Alignment = lvwColumnLeft
        'RTActivityLogID
        .ColumnHeaders.Item(GuiActLogListView.RTActivityLogID).Width = 0   'Hidden
        .ColumnHeaders.Item(GuiActLogListView.RTActivityLogID).Alignment = lvwColumnLeft
        'AssignmentsID
        .ColumnHeaders.Item(GuiActLogListView.AssignmentsID).Width = 0  'Hidden
        .ColumnHeaders.Item(GuiActLogListView.AssignmentsID).Alignment = lvwColumnLeft
        'BillingCountID
        .ColumnHeaders.Item(GuiActLogListView.BillingCountID).Width = 0  'Hidden
        .ColumnHeaders.Item(GuiActLogListView.BillingCountID).Alignment = lvwColumnLeft
        'IDBillingCount
        .ColumnHeaders.Item(GuiActLogListView.IDBillingCount).Width = 0  'Hidden
        .ColumnHeaders.Item(GuiActLogListView.IDBillingCount).Alignment = lvwColumnLeft
        'DownLoadMe
        .ColumnHeaders.Item(GuiActLogListView.DownLoadMe).Width = 0  'Hidden
        .ColumnHeaders.Item(GuiActLogListView.DownLoadMe).Alignment = lvwColumnLeft
        'UpdateByUserID
        .ColumnHeaders.Item(GuiActLogListView.UpdateByUserID).Width = 0  'Hidden
        .ColumnHeaders.Item(GuiActLogListView.UpdateByUserID).Alignment = lvwColumnLeft
    End With
    
    bGridOn = CBool(GetSetting(App.EXEName, "GENERAL", "GRID_ON", False))
    lstvActLog.GridLines = bGridOn
    
   
    If bHideDeleted Then
        chkHideDeleted.Value = vbChecked
    Else
        chkHideDeleted.Value = vbUnchecked
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub LoadHeaderlstvActLog"
End Sub

Private Sub PopulateActLogInfo()
    On Error GoTo EH
    Dim RS As ADODB.Recordset
    Dim udfActLogInfo As GuiActLogInfoItem
    
    Set RS = mfrmClaim.adoRSRTActivityLogInfo
    
    If RS.RecordCount = 1 Then
        RS.MoveFirst
        '1. Present Durring Inspection
        If CBool(RS.Fields("AL01_sPresentDurringInspection")) Then
            optAL01_sPresentDurringInspection(GuiOptValues.YES).Value = True
        Else
            optAL01_sPresentDurringInspection(GuiOptValues.No).Value = True
        End If
        
        '2. Explained Estimate
        If CBool(RS.Fields("AL02_sExplainedEstimate")) Then
            optAL02_sExplainedEstimate(GuiOptValues.YES).Value = True
        Else
            optAL02_sExplainedEstimate(GuiOptValues.No).Value = True
        End If
        
        '3. Explained RCV
        If CBool(RS.Fields("AL03_sExplainedRCV")) Then
            optAL03_sExplainedRCV(GuiOptValues.YES).Value = True
        Else
            If CBool(RS.Fields("AL03_sExplainedRCVNA")) Then
                optAL03_sExplainedRCV(GuiOptValues.NA).Value = True
            Else
                optAL03_sExplainedRCV(GuiOptValues.No).Value = True
            End If
        End If
        '4. Confirm Mortgagee Is Correct
        If CBool(RS.Fields("AL04_sConfirmMortgageeIsCorrect")) Then
            optAL04_sConfirmMortgageeIsCorrect(GuiOptValues.YES).Value = True
        Else
            If CBool(RS.Fields("AL04_sConfirmMortgageeIsCorrectNA")) Then
                optAL04_sConfirmMortgageeIsCorrect(GuiOptValues.NA).Value = True
            Else
                optAL04_sConfirmMortgageeIsCorrect(GuiOptValues.No).Value = True
            End If
        End If
        
        '5. Explained Mortgagee Checks
        If CBool(RS.Fields("AL05_sExplainedMortgageeChecks")) Then
            optAL05_sExplainedMortgageeChecks(GuiOptValues.YES).Value = True
        Else
            If CBool(RS.Fields("AL05_sExplainedMortgageeChecksNA")) Then
                optAL05_sExplainedMortgageeChecks(GuiOptValues.NA).Value = True
            Else
                optAL05_sExplainedMortgageeChecks(GuiOptValues.No).Value = True
            End If
        End If
        
        '6. Confirmed Coverage
        If CBool(RS.Fields("AL06_sConfirmedCoverage")) Then
            optAL06_sConfirmedCoverage(GuiOptValues.YES).Value = True
        Else
            optAL06_sConfirmedCoverage(GuiOptValues.No).Value = True
        End If
        
        '7. Prior Loss
        If CBool(RS.Fields("AL07_sPriorLoss")) Then
            optAL07_sPriorLoss(GuiOptValues.YES).Value = True
        Else
            If CBool(RS.Fields("AL07_sPriorLossNA")) Then
                optAL07_sPriorLoss(GuiOptValues.NA).Value = True
            Else
                optAL07_sPriorLoss(GuiOptValues.No).Value = True
            End If
        End If
        
        '8. Salvage
        If CBool(RS.Fields("AL08_sSalvage")) Then
            optAL08_sSalvage(GuiOptValues.YES).Value = True
        Else
            optAL08_sSalvage(GuiOptValues.No).Value = True
        End If
        
        '9. Subrogation
        If CBool(RS.Fields("AL09_sSubrogation")) Then
            optAL09_sSubrogation(GuiOptValues.YES).Value = True
        Else
            optAL09_sSubrogation(GuiOptValues.No).Value = True
        End If
    Else
        
        'AssignmentsID
        udfActLogInfo.AssignmentsID = msAssignmentsID
        'IDAssignments
        udfActLogInfo.IDAssignments = msAssignmentsID
        
        'Need to make an entry in the Act log info Table
        'Deault Values for Options
        'Present Durring Inspection
        optAL01_sPresentDurringInspection(GuiOptValues.YES).Value = True
        optAL01_sPresentDurringInspection(GuiOptValues.No).Value = False
        udfActLogInfo.AL01_sPresentDurringInspection = "True"
        'Explained Estimate
        optAL02_sExplainedEstimate(GuiOptValues.YES).Value = True
        optAL02_sExplainedEstimate(GuiOptValues.No).Value = False
        udfActLogInfo.AL02_sExplainedEstimate = "True"
        'Explained RCV
        optAL03_sExplainedRCV(GuiOptValues.YES).Value = True
        optAL03_sExplainedRCV(GuiOptValues.No).Value = False
        udfActLogInfo.AL03_sExplainedRCV = "True"
        optAL03_sExplainedRCV(GuiOptValues.NA).Value = False
        udfActLogInfo.AL03_sExplainedRCVNA = "False"
        'Confirm Mortgagee Is Correct
        optAL04_sConfirmMortgageeIsCorrect(GuiOptValues.YES).Value = False
        optAL04_sConfirmMortgageeIsCorrect(GuiOptValues.No).Value = False
        udfActLogInfo.AL04_sConfirmMortgageeIsCorrect = "False"
        optAL04_sConfirmMortgageeIsCorrect(GuiOptValues.NA).Value = True
        udfActLogInfo.AL04_sConfirmMortgageeIsCorrectNA = "True"
        'Explained Mortgagee Checks
        optAL05_sExplainedMortgageeChecks(GuiOptValues.YES).Value = False
        optAL05_sExplainedMortgageeChecks(GuiOptValues.No).Value = False
        udfActLogInfo.AL05_sExplainedMortgageeChecks = "False"
        optAL05_sExplainedMortgageeChecks(GuiOptValues.NA).Value = True
        udfActLogInfo.AL05_sExplainedMortgageeChecksNA = "True"
        'Confirmed Coverage
        optAL06_sConfirmedCoverage(GuiOptValues.YES).Value = True
        optAL06_sConfirmedCoverage(GuiOptValues.No).Value = False
        udfActLogInfo.AL06_sConfirmedCoverage = "True"
        'Prior Loss
        optAL07_sPriorLoss(GuiOptValues.YES).Value = False
        optAL07_sPriorLoss(GuiOptValues.No).Value = True
        udfActLogInfo.AL07_sPriorLoss = "False"
        optAL07_sPriorLoss(GuiOptValues.NA).Value = False
        udfActLogInfo.AL07_sPriorLossNA = "False"
        'Salvage
        optAL08_sSalvage(GuiOptValues.YES).Value = False
        optAL08_sSalvage(GuiOptValues.No).Value = True
        udfActLogInfo.AL08_sSalvage = "False"
        'Subrogation
        optAL09_sSubrogation(GuiOptValues.YES).Value = False
        optAL09_sSubrogation(GuiOptValues.No).Value = True
        udfActLogInfo.AL09_sSubrogation = "False"
        
        'IsDeleted
        udfActLogInfo.IsDeleted = "False"
        'DownLoadMe
        udfActLogInfo.DownLoadMe = "False"
        'UpLoadMe
        udfActLogInfo.UpLoadMe = "True"
        'AdminComments
        udfActLogInfo.AdminComments = vbNullString
        'DateLastUpdated
        udfActLogInfo.DateLastUpdated = Format(Now(), "MM/DD/YYYY HH:MM:SS")
        'UpdateByUserID
        udfActLogInfo.UpdateByUserID = goUtil.gsCurUsersID
        'Add it Here
        AddActivityLogInfoItem udfActLogInfo
        'Reset the RS for ActLogInfo
        mfrmClaim.SetadoRSRTActivityLogInfo msAssignmentsID
    End If
    
    'Cleanup
    Set RS = Nothing
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub PopulateActLogInfo"
End Sub


Private Sub PopulatelstvActLog()
    On Error GoTo EH
    Dim iMyIcon As Long
    Dim sFlagText As String
    Dim sTemp As String
    Dim itmX As ListItem
    Dim oListView As MSComctlLib.ListView
    Dim RS As ADODB.Recordset
    
    'Clear the List view
    Set oListView = lstvActLog

    oListView.Visible = False
    oListView.ListItems.Clear

    Set RS = mfrmClaim.adoRSRTActivityLog

    If Not RS.EOF Then
        RS.MoveFirst
        Do Until RS.EOF
            '1. ActDate
            If Not IsNull(RS.Fields("ActDate").Value) Then
                If IsDate(RS.Fields("ActDate").Value) Then
                    Set itmX = oListView.ListItems.Add(, """" & goUtil.IsNullIsVbNullString(RS.Fields("ID")) & """", Format(goUtil.IsNullIsVbNullString(RS.Fields("ActDate")), "MM/DD/YYYY"))
                    'Be sure to use the ActTime With Date AND TIME when populating Date Sort
                    itmX.SubItems(GuiActLogListView.ActDateSort - 1) = Format(RS.Fields("ActTime").Value, "YYYY/MM/DD HH:MM")
                Else
                    Set itmX = oListView.ListItems.Add(, """" & goUtil.IsNullIsVbNullString(RS.Fields("ID")) & """", vbNullString)
                    itmX.SubItems(GuiActLogListView.ActDateSort - 1) = vbNullString
                End If
            Else
                Set itmX = oListView.ListItems.Add(, """" & goUtil.IsNullIsVbNullString(RS.Fields("ID")) & """", vbNullString)
                itmX.SubItems(GuiActLogListView.ActDateSort - 1) = vbNullString
            End If
            
            '2. ActTime
            If Not IsNull(RS.Fields("ActTime").Value) Then
                If IsDate(RS.Fields("ActTime").Value) Then
                    itmX.SubItems(GuiActLogListView.ActTime - 1) = Format(RS.Fields("ActTime").Value, "HH:MM")
                    'Be sure to use the ActTime With TIME ONLY when populating ActTime Sort
                    itmX.SubItems(GuiActLogListView.ActTimeSort - 1) = Format(RS.Fields("ActTime").Value, "HH:MM")
                Else
                    itmX.SubItems(GuiActLogListView.ActTime - 1) = vbNullString
                    itmX.SubItems(GuiActLogListView.ActTimeSort - 1) = vbNullString
                End If
            Else
                itmX.SubItems(GuiActLogListView.ActTime - 1) = vbNullString
                itmX.SubItems(GuiActLogListView.ActTimeSort - 1) = vbNullString
            End If

            '3. ActText
            itmX.SubItems(GuiActLogListView.ActText - 1) = goUtil.IsNullIsVbNullString(RS.Fields("ActText"))
            
            '4. ServiceTime
            itmX.SubItems(GuiActLogListView.ServiceTime - 1) = Format(goUtil.IsNullIsVbNullString(RS.Fields("ServiceTime")), "#,###,###,##0.00")
            itmX.SubItems(GuiActLogListView.ServiceTimeSort - 1) = goUtil.utNumInTextSortFormat(Format(goUtil.IsNullIsVbNullString(RS.Fields("ServiceTime")), "#,###,###,##0.00"))
            
            '5. IB Shows the IB this ActLog Items is Asscoiated with
            itmX.SubItems(GuiActLogListView.IB - 1) = goUtil.IsNullIsVbNullString(RS.Fields("IB"))
            
            '6. PageBreakAfter hidden until future Use
            itmX.SubItems(GuiActLogListView.PageBreakAfter - 1) = goUtil.IsNullIsVbNullString(RS.Fields("PageBreakAfter"))
            
            '7. BlankPageAfter hidden until future Use
            itmX.SubItems(GuiActLogListView.BlankPageAfter - 1) = goUtil.IsNullIsVbNullString(RS.Fields("BlankPageAfter"))
            
            '8. BlankRowsAfter
            itmX.SubItems(GuiActLogListView.BlankRowsAfter - 1) = goUtil.IsNullIsVbNullString(RS.Fields("BlankRowsAfter"))
            itmX.SubItems(GuiActLogListView.BlankRowsAfterSort - 1) = goUtil.utNumInTextSortFormat(goUtil.IsNullIsVbNullString(RS.Fields("BlankRowsAfter")))
            
            '9. IsMgrEntry hidden until future Use
            itmX.SubItems(GuiActLogListView.IsMgrEntry - 1) = goUtil.IsNullIsVbNullString(RS.Fields("IsMgrEntry"))
            
            '10. Is Deleted
            If CBool(RS.Fields("IsDeleted")) Then
                iMyIcon = GuiActLogStatusList.IsDeleted
            Else
                iMyIcon = Empty
            End If
            sFlagText = goUtil.GetFlagText(RS.Fields("IsDeleted"))
            itmX.SubItems(GuiActLogListView.IsDeleted - 1) = sFlagText
            itmX.ListSubItems(GuiActLogListView.IsDeleted - 1).ReportIcon = iMyIcon
            
            '11. UpLoad Me
            If CBool(RS.Fields("UpLoadMe")) Then
                iMyIcon = GuiActLogStatusList.UpLoadMe
            Else
                iMyIcon = Empty
            End If
            sFlagText = goUtil.GetFlagText(RS.Fields("UpLoadMe"))
            itmX.SubItems(GuiActLogListView.UpLoadMe - 1) = sFlagText
            itmX.ListSubItems(GuiActLogListView.UpLoadMe - 1).ReportIcon = iMyIcon
            
            '12. DateLastUpdated
            If Not IsNull(RS.Fields("DateLastUpdated").Value) Then
                If IsDate(RS.Fields("DateLastUpdated").Value) Then
                    itmX.SubItems(GuiActLogListView.DateLastUpdated - 1) = Format(RS.Fields("DateLastUpdated").Value, "MM/DD/YYYY HH:MM:SS")
                    itmX.SubItems(GuiActLogListView.DateLastUpdatedSort - 1) = Format(RS.Fields("DateLastUpdated").Value, "YYYY/MM/DD HH:MM:SS")
                Else
                    itmX.SubItems(GuiActLogListView.DateLastUpdated - 1) = vbNullString
                    itmX.SubItems(GuiActLogListView.DateLastUpdatedSort - 1) = vbNullString
                End If
            Else
                itmX.SubItems(GuiActLogListView.DateLastUpdated - 1) = vbNullString
                itmX.SubItems(GuiActLogListView.DateLastUpdatedSort - 1) = vbNullString
            End If
            
            '13. AdminComments
            itmX.SubItems(GuiActLogListView.AdminComments - 1) = goUtil.IsNullIsVbNullString(RS.Fields("AdminComments"))
            
            '14. ID hidden
            itmX.SubItems(GuiActLogListView.ID - 1) = goUtil.IsNullIsVbNullString(RS.Fields("ID"))
            
            '15. IDAssignments hidden
            itmX.SubItems(GuiActLogListView.IDAssignments - 1) = goUtil.IsNullIsVbNullString(RS.Fields("IDAssignments"))
            
            '16. RTActivityLogID hidden
            itmX.SubItems(GuiActLogListView.RTActivityLogID - 1) = goUtil.IsNullIsVbNullString(RS.Fields("RTActivityLogID"))
            
            '17. AssignmentsID hidden
            itmX.SubItems(GuiActLogListView.AssignmentsID - 1) = goUtil.IsNullIsVbNullString(RS.Fields("AssignmentsID"))
            
            '18. BillingCountID hidden
            itmX.SubItems(GuiActLogListView.BillingCountID - 1) = goUtil.IsNullIsVbNullString(RS.Fields("BillingCountID"))
            
            '19. IDBillingCount hidden
            itmX.SubItems(GuiActLogListView.IDBillingCount - 1) = goUtil.IsNullIsVbNullString(RS.Fields("IDBillingCount"))
            
            '20. DownLoadMe hidden
            itmX.SubItems(GuiActLogListView.DownLoadMe - 1) = goUtil.IsNullIsVbNullString(RS.Fields("DownLoadMe"))
            
            '21. UpdateByUserID hidden
            itmX.SubItems(GuiActLogListView.UpdateByUserID - 1) = goUtil.IsNullIsVbNullString(RS.Fields("UpdateByUserID"))
            
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
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub PopulatelstvActLog"
    oListView.Visible = True
End Sub

Private Sub lstvActLog_Click()
    On Error GoTo EH
    itmXSelected = lstvActLog.SelectedItem
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lstvActLog_Click"
End Sub

Private Sub lstvActLog_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    On Error GoTo EH
    Dim sMess As String
    If lstvActLog.SortOrder = lvwAscending Then
        lstvActLog.SortOrder = lvwDescending
        sMess = "Descending"
    Else
        lstvActLog.SortOrder = lvwAscending
        sMess = "Ascending"
    End If
    
    'Set the Tool Tip
    lstvActLog.ToolTipText = "Sort By " & ColumnHeader.Text & " " & sMess
    
    'Need to see if the Column clicked was a NON text sort
    'Like Date or number.  If a non textual column was clicked
    'need to use the next column as the sort since this next column
    'will be hidden and contain a sort friendly format.
    'Sort Key is Base 0 where CoulmnHeader is not
    Select Case ColumnHeader.Index
        Case GuiActLogListView.ActDate, GuiActLogListView.ActTime
            lstvActLog.SortKey = ColumnHeader.Index
        Case GuiActLogListView.BlankRowsAfter
            lstvActLog.SortKey = ColumnHeader.Index
        Case GuiActLogListView.DateLastUpdated
            lstvActLog.SortKey = ColumnHeader.Index
        Case GuiActLogListView.ServiceTime
            lstvActLog.SortKey = ColumnHeader.Index
        Case Else
            lstvActLog.SortKey = ColumnHeader.Index - 1
    End Select
    
    lstvActLog.Sorted = True
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lstvActLog_ColumnClick"
End Sub

Private Sub lstvActLog_DblClick()
    On Error GoTo EH
    itmXSelected = lstvActLog.SelectedItem
    If Not lstvActLog.SelectedItem Is Nothing Then
        EditActLog
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lstvActLog_DblClick"
End Sub

Public Function EditActLog() As Boolean
    On Error GoTo EH
    Dim oListView As MSComctlLib.ListView
    Dim itmX As MSComctlLib.ListItem

    If lstvActLog.ListItems.Count = 0 Then
        Exit Function
    Else
        Set oListView = lstvActLog
    End If

    Set itmX = oListView.SelectedItem
    
    With AddActLog
        .MyActivityLog = Me
        .MyfrmClaim = Me.MyfrmClaim
        .Adding = False
        .AssignmentsID = itmX.SubItems(GuiActLogListView.IDAssignments - 1)
        .ActLogID = itmX.SubItems(GuiActLogListView.ID - 1)
         Load AddActLog
        .Caption = "Edit Activity Log"
        .txtActDate.Text = itmX.Text
        .timeActTime.ecsTime = itmX.SubItems(GuiActLogListView.ActTime - 1)
        .txtServiceTime.Text = itmX.SubItems(GuiActLogListView.ServiceTime - 1)
        .txtActText = itmX.SubItems(GuiActLogListView.ActText - 1)
        .cmdSave.Enabled = False
        .Show vbModal
    End With
   

    Unload AddActLog
    Set AddActLog = Nothing
    If lstvActLog.Visible Then
        lstvActLog.SetFocus
    End If
    
    EditActLog = True
    
    Set oListView = Nothing
    Set itmX = Nothing
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function EditActLog"
    Unload AddPhoto
    Set AddPhoto = Nothing
End Function

Private Sub lstvActLog_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error GoTo EH
    itmXSelected = lstvActLog.SelectedItem
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lstvActLog_ItemClick"
End Sub

Private Sub lstvActLog_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo EH
    Select Case KeyCode
        Case KeyCodeConstants.vbKeyReturn
            EditActLog
        Case vbKeyDelete
            cmdDelActLog_Click
    End Select
Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lstvActLog_KeyDown"
End Sub

Private Sub lstvActLog_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo EH
    If Button = vbRightButton Then
        PopupMenu PopUpmnuActLog
    End If
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lstvActLog_MouseUp"
End Sub

Private Sub mnuDeleteActLog_Click()
    On Error GoTo EH
    Dim itmX As MSComctlLib.ListItem
    Dim sActLogID As String
    
    Set itmX = lstvActLog.SelectedItem
    
    If Not itmX Is Nothing Then
        sActLogID = itmX.SubItems(GuiActLogListView.ID - 1)
        If MsgBox("Are you sure you want to delete this Activity Log Item?", vbYesNo, "DELETE SELECTED ITEM") = vbYes Then
            If DeleteActLogItem(sActLogID) Then
                lstvActLog.ListItems.Remove ("""" & sActLogID & """")
            End If
        End If
        lstvActLog.SetFocus
    End If
    
    Set itmX = Nothing
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub mnuDeleteActLog_Click"
End Sub


Private Sub mnuEditActLog_Click()
    On Error GoTo EH
    
    EditActLog
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub mnuEditActLog_Click"
End Sub

Public Function DeleteActLogItem(psID As String) As Boolean
    On Error GoTo EH
    Dim sSQL As String
    Dim bUpdateAsDeletedOnly As Boolean
    Dim oConn As ADODB.Connection
    Dim RS As ADODB.Recordset
    
    Set oConn = New ADODB.Connection
    Set RS = New ADODB.Recordset
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name


    sSQL = "SELECT A.[ID] "
    sSQL = sSQL & "FROM RTActivityLog A "
    sSQL = sSQL & "WHERE A.[ID] = " & psID & " "
    'Only allow actual deletion of ActLog that have never been uploaded
    'Negative number for the Main Table Indentity will be negative number
    'if this is true.
    sSQL = sSQL & "AND (A.[RTActivityLogID] Is Null Or A.[RTActivityLogID] < 0)  "
    
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
        sSQL = "UPDATE RTActivityLog SET "
        sSQL = sSQL & "[IsDeleted] = IIF([IsDeleted], False, True), "
        sSQL = sSQL & "[UpLoadMe] = True, "
        sSQL = sSQL & "[DateLastUpdated] = #" & Format(Now(), "MM/DD/YYYY HH:MM:SS") & "#, "
        sSQL = sSQL & "[UpdateByUserID] = " & goUtil.gsCurUsersID & " "
        sSQL = sSQL & "WHERE [ID] = " & psID & " "
        sSQL = sSQL & "AND [IDAssignments] = " & msAssignmentsID & " "
    Else
        sSQL = "DELETE * FROM RTActivityLog A "
        sSQL = sSQL & "WHERE A.[ID] = " & psID & " "
        sSQL = sSQL & "AND A.[IDAssignments] = " & msAssignmentsID & " "
    End If

    oConn.Execute sSQL
    
    DeleteActLogItem = True
    'clean up
    
    Set RS = Nothing
    Set oConn = Nothing
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function DeleteActLogItem"
End Function

Private Sub mnuSelectAllActLog_Click()
    On Error GoTo EH
    Dim itmX As ListItem
    
    For Each itmX In lstvActLog.ListItems
        itmX.Selected = True
    Next
    
    Set itmX = Nothing
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub mnuSelectAllActLog_Click"
End Sub

Public Function PrintActLog(psIDAssignments As String) As Boolean
    On Error GoTo EH
    Dim MyActLog As Object
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
                                    "_arActivityLog", _
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
            goUtil.utShellExecute , "OPEN", sPDFFilePath, , , vbNormalFocus, True, True, True, "Activity Log"
            DoEvents
            Sleep 1000
            goUtil.utDeleteFile sPDFFilePath
            oCarList.CLEANUP
            Set oCarList = Nothing
        End If
    Else
    
        Set MyActLog = oCarList.GetARReport(sReportName, lrptVersion, sParams)
    
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
            MyActLog.Run False 'True
            .objARvReport = MyActLog
            .sRptTitle = "Activity Log"
            .HidePrintButton = False
            .ShowReportOnForm moForm, vbModeless
            Unload .objARvReport
            Set .objARvReport = Nothing
            Unload MyActLog
            Set MyActLog = Nothing
            oCarList.CLEANUP
            Set oCarList = Nothing
        End With
    End If
    
    'Clear the Local ref to this report object only
    'The actual cleanup of this active report object will occur within gARV
    Set oConn = Nothing
    PrintActLog = True
    
    Exit Function
EH:
    Screen.MousePointer = vbDefault
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function PrintActLog"
End Function

Public Function AddActivityLogInfoItem(pudtActLogInfo As GuiActLogInfoItem) As Boolean
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim adoRS As ADODB.Recordset
    Dim sSQL As String
    Dim sID As String
    
    Screen.MousePointer = MousePointerConstants.vbHourglass
    
    'Be Sure that there is not already an item in the info table for this
    'Assignment
    sSQL = "SELECT [AssignmentsID] "
    sSQL = sSQL & "FROM RTActivityLogInfo "
    sSQL = sSQL & "WHERE [AssignmentsID] = " & msAssignmentsID & " "
    
    Set oConn = New ADODB.Connection
    Set adoRS = New ADODB.Recordset
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    'Use Disconnected Record Set on asUseClient Cusor ONLY !
    adoRS.CursorLocation = adUseClient
    adoRS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set adoRS.ActiveConnection = Nothing
    
    If adoRS.RecordCount > 0 Then
        GoTo CLEAN_UP
    End If
    
    With pudtActLogInfo
        .AssignmentsID = msAssignmentsID
        .IDAssignments = msAssignmentsID 'not set here
    End With
    
    sSQL = "INSERT INTO RTActivityLogInfo "
    sSQL = sSQL & "( "
    sSQL = sSQL & "[AssignmentsID], "
    sSQL = sSQL & "[IDAssignments], "
    sSQL = sSQL & "[AL01_sPresentDurringInspection], "
    sSQL = sSQL & "[AL02_sExplainedEstimate], "
    sSQL = sSQL & "[AL03_sExplainedRCV], "
    sSQL = sSQL & "[AL03_sExplainedRCVNA], "
    sSQL = sSQL & "[AL04_sConfirmMortgageeIsCorrect], "
    sSQL = sSQL & "[AL04_sConfirmMortgageeIsCorrectNA], "
    sSQL = sSQL & "[AL05_sExplainedMortgageeChecks], "
    sSQL = sSQL & "[AL05_sExplainedMortgageeChecksNA], "
    sSQL = sSQL & "[AL06_sConfirmedCoverage], "
    sSQL = sSQL & "[AL07_sPriorLoss], "
    sSQL = sSQL & "[AL07_sPriorLossNA], "
    sSQL = sSQL & "[AL08_sSalvage], "
    sSQL = sSQL & "[AL09_sSubrogation], "
    sSQL = sSQL & "[IsDeleted], "
    sSQL = sSQL & "[DownLoadMe], "
    sSQL = sSQL & "[UpLoadMe], "
    sSQL = sSQL & "[AdminComments], "
    sSQL = sSQL & "[DateLastUpdated], "
    sSQL = sSQL & "[UpdateByUserID] "
    sSQL = sSQL & ") "
    sSQL = sSQL & "SELECT "
    sSQL = sSQL & pudtActLogInfo.AssignmentsID & " As [AssignmentsID], "
    sSQL = sSQL & pudtActLogInfo.IDAssignments & " As [IDAssignments], "
    sSQL = sSQL & pudtActLogInfo.AL01_sPresentDurringInspection & " As [AL01_sPresentDurringInspection], "
    sSQL = sSQL & pudtActLogInfo.AL02_sExplainedEstimate & " As [AL02_sExplainedEstimate], "
    sSQL = sSQL & pudtActLogInfo.AL03_sExplainedRCV & " As [AL03_sExplainedRCV], "
    sSQL = sSQL & pudtActLogInfo.AL03_sExplainedRCVNA & " As [AL03_sExplainedRCVNA], "
    sSQL = sSQL & pudtActLogInfo.AL04_sConfirmMortgageeIsCorrect & " As [AL04_sConfirmMortgageeIsCorrect], "
    sSQL = sSQL & pudtActLogInfo.AL04_sConfirmMortgageeIsCorrectNA & " As [AL04_sConfirmMortgageeIsCorrectNA], "
    sSQL = sSQL & pudtActLogInfo.AL05_sExplainedMortgageeChecks & " As [AL05_sExplainedMortgageeChecks], "
    sSQL = sSQL & pudtActLogInfo.AL05_sExplainedMortgageeChecksNA & " As [AL05_sExplainedMortgageeChecksNA], "
    sSQL = sSQL & pudtActLogInfo.AL06_sConfirmedCoverage & " As [AL06_sConfirmedCoverage], "
    sSQL = sSQL & pudtActLogInfo.AL07_sPriorLoss & " As [AL07_sPriorLoss], "
    sSQL = sSQL & pudtActLogInfo.AL07_sPriorLossNA & " As [AL07_sPriorLossNA], "
    sSQL = sSQL & pudtActLogInfo.AL08_sSalvage & " As [AL08_sSalvage], "
    sSQL = sSQL & pudtActLogInfo.AL09_sSubrogation & " As [AL09_sSubrogation], "
    sSQL = sSQL & pudtActLogInfo.IsDeleted & " As [IsDeleted], "
    sSQL = sSQL & pudtActLogInfo.DownLoadMe & " As [DownLoadMe], "
    sSQL = sSQL & pudtActLogInfo.UpLoadMe & " As [UpLoadMe], "
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtActLogInfo.AdminComments) & "'" & " As [AdminComments], "
    sSQL = sSQL & "#" & pudtActLogInfo.DateLastUpdated & "#" & " As [DateLastUpdated], "
    sSQL = sSQL & pudtActLogInfo.UpdateByUserID & " As [UpdateByUserID] "
   
    oConn.Execute sSQL
    
    Screen.MousePointer = MousePointerConstants.vbDefault
    AddActivityLogInfoItem = True
CLEAN_UP:

    'Clean up
    Set oConn = Nothing
    Set adoRS = Nothing
    Exit Function
EH:
    Screen.MousePointer = MousePointerConstants.vbDefault
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function AddActivityLogInfoItem"
End Function

Public Function EditActivityLogInfoItem(pudtActLogInfo As GuiActLogInfoItem) As Boolean
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    
    'Check for Nullstring ID
    If pudtActLogInfo.AssignmentsID = vbNullString Then
        pudtActLogInfo.AssignmentsID = "null"
    End If
    If pudtActLogInfo.IDAssignments = vbNullString Then
        pudtActLogInfo.IDAssignments = "null"
    End If
    
    sSQL = "UPDATE RTActivityLogInfo Set "
    sSQL = sSQL & "[AssignmentsID] = " & pudtActLogInfo.AssignmentsID & ", "
    sSQL = sSQL & "[IDAssignments] = " & pudtActLogInfo.IDAssignments & ", "
    sSQL = sSQL & "[AL01_sPresentDurringInspection] = " & pudtActLogInfo.AL01_sPresentDurringInspection & ", "
    sSQL = sSQL & "[AL02_sExplainedEstimate] = " & pudtActLogInfo.AL02_sExplainedEstimate & ", "
    sSQL = sSQL & "[AL03_sExplainedRCV] = " & pudtActLogInfo.AL03_sExplainedRCV & ", "
    sSQL = sSQL & "[AL03_sExplainedRCVNA] = " & pudtActLogInfo.AL03_sExplainedRCVNA & ", "
    sSQL = sSQL & "[AL04_sConfirmMortgageeIsCorrect] = " & pudtActLogInfo.AL04_sConfirmMortgageeIsCorrect & ", "
    sSQL = sSQL & "[AL04_sConfirmMortgageeIsCorrectNA] = " & pudtActLogInfo.AL04_sConfirmMortgageeIsCorrectNA & ", "
    sSQL = sSQL & "[AL05_sExplainedMortgageeChecks] = " & pudtActLogInfo.AL05_sExplainedMortgageeChecks & ", "
    sSQL = sSQL & "[AL05_sExplainedMortgageeChecksNA] = " & pudtActLogInfo.AL05_sExplainedMortgageeChecksNA & ", "
    sSQL = sSQL & "[AL06_sConfirmedCoverage] = " & pudtActLogInfo.AL06_sConfirmedCoverage & ", "
    sSQL = sSQL & "[AL07_sPriorLoss] = " & pudtActLogInfo.AL07_sPriorLoss & ", "
    sSQL = sSQL & "[AL07_sPriorLossNA] = " & pudtActLogInfo.AL07_sPriorLossNA & ", "
    sSQL = sSQL & "[AL08_sSalvage] = " & pudtActLogInfo.AL08_sSalvage & ", "
    sSQL = sSQL & "[AL09_sSubrogation] = " & pudtActLogInfo.AL09_sSubrogation & ", "
    sSQL = sSQL & "[IsDeleted] = " & pudtActLogInfo.IsDeleted & ", "
    sSQL = sSQL & "[DownLoadMe] = " & pudtActLogInfo.DownLoadMe & ", "
    sSQL = sSQL & "[UpLoadMe] = " & pudtActLogInfo.UpLoadMe & ", "
    sSQL = sSQL & "[AdminComments] = '" & goUtil.utCleanSQLString(pudtActLogInfo.AdminComments) & "', "
    sSQL = sSQL & "[DateLastUpdated] = #" & pudtActLogInfo.DateLastUpdated & "#, "
    sSQL = sSQL & "[UpdateByUserID]  = " & pudtActLogInfo.UpdateByUserID & " "
    sSQL = sSQL & "WHERE IDAssignments = " & pudtActLogInfo.IDAssignments & " "

    Set oConn = New ADODB.Connection
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    oConn.Execute sSQL
    
    EditActivityLogInfoItem = True
    'Clean up
    Set oConn = Nothing
    Exit Function
    
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function EditActivityLogInfoItem"
End Function

Public Function AddActLogItem(pudtActLog As GuiActLogItem) As Boolean
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    Dim sID As String
    
    Screen.MousePointer = MousePointerConstants.vbHourglass
    
    sID = goUtil.GetAccessDBUID("ID", "RTActivityLog")
    
    With pudtActLog
        .RTActivityLogID = sID
        .AssignmentsID = msAssignmentsID
        'Use Current Billing Item if Available
        .BillingCountID = mfrmClaim.GetCurrentBillingCountID(False)
        .ID = sID
        .IDAssignments = msAssignmentsID 'not set here
        .IDBillingCount = mfrmClaim.GetCurrentBillingCountID(True)
    End With
    
    sSQL = "INSERT INTO RTActivityLog "
    sSQL = sSQL & "( "
    sSQL = sSQL & "[RTActivityLogID], "
    sSQL = sSQL & "[AssignmentsID], "
    sSQL = sSQL & "[BillingCountID], "
    sSQL = sSQL & "[ID], "
    sSQL = sSQL & "[IDAssignments], "
    sSQL = sSQL & "[IDBillingCount], "
    sSQL = sSQL & "[ServiceTime], "
    sSQL = sSQL & "[ActDate], "
    sSQL = sSQL & "[ActText], "
    sSQL = sSQL & "[ActTime], "
    sSQL = sSQL & "[PageBreakAfter], "
    sSQL = sSQL & "[BlankPageAfter], "
    sSQL = sSQL & "[BlankRowsAfter], "
    sSQL = sSQL & "[IsMgrEntry], "
    sSQL = sSQL & "[IsDeleted], "
    sSQL = sSQL & "[DownLoadMe], "
    sSQL = sSQL & "[UpLoadMe], "
    sSQL = sSQL & "[AdminComments], "
    sSQL = sSQL & "[DateLastUpdated], "
    sSQL = sSQL & "[UpdateByUserID] "
    sSQL = sSQL & ") "
    sSQL = sSQL & "SELECT "
    sSQL = sSQL & pudtActLog.RTActivityLogID & " As [RTActivityLogID], "
    sSQL = sSQL & pudtActLog.AssignmentsID & " As [AssignmentsID], "
    sSQL = sSQL & pudtActLog.BillingCountID & " As [BillingCountID] , "
    sSQL = sSQL & pudtActLog.ID & " As [ID], "
    sSQL = sSQL & pudtActLog.IDAssignments & " As [IDAssignments], "
    sSQL = sSQL & pudtActLog.IDBillingCount & " As [IDBillingCount], "
    sSQL = sSQL & pudtActLog.ServiceTime & " As [ServiceTime], "
    sSQL = sSQL & "#" & pudtActLog.ActDate & "#" & " As [ActDate], "
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtActLog.ActText) & "'" & " As [ActText], "
    sSQL = sSQL & "#" & pudtActLog.ActTime & "#" & " As [ActTime], "
    sSQL = sSQL & pudtActLog.PageBreakAfter & " As [PageBreakAfter], "
    sSQL = sSQL & pudtActLog.BlankPageAfter & " As [BlankPageAfter], "
    sSQL = sSQL & pudtActLog.BlankRowsAfter & " As [BlankRowsAfter], "
    sSQL = sSQL & pudtActLog.IsMgrEntry & " As [IsMgrEntry], "
    sSQL = sSQL & pudtActLog.IsDeleted & " As [IsDeleted], "
    sSQL = sSQL & pudtActLog.DownLoadMe & " As [DownLoadMe], "
    sSQL = sSQL & pudtActLog.UpLoadMe & " As [UpLoadMe], "
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtActLog.AdminComments) & "'" & " As [AdminComments], "
    sSQL = sSQL & "#" & pudtActLog.DateLastUpdated & "#" & " As [DateLastUpdated], "
    sSQL = sSQL & pudtActLog.UpdateByUserID & " As [UpdateByUserID] "
   
    Set oConn = New ADODB.Connection
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    oConn.Execute sSQL
    
    Screen.MousePointer = MousePointerConstants.vbDefault
    AddActLogItem = True
    
    'Clean up
    Set oConn = Nothing
    Exit Function
EH:
    Screen.MousePointer = vbDefault
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function AddActLogItem"
End Function

Public Function EditActLogItem(pudtActLog As GuiActLogItem) As Boolean
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    
    With pudtActLog
        If .RTActivityLogID = vbNullString Or .RTActivityLogID = "0" Then
            .RTActivityLogID = "Null"
        End If
        .AssignmentsID = msAssignmentsID
        'Use Current Billing Item if Available
        .BillingCountID = mfrmClaim.GetCurrentBillingCountID(False)
        If .ID = vbNullString Or .ID = "0" Then
            .ID = "Null"
        End If
        .IDAssignments = msAssignmentsID 'not set here
        .IDBillingCount = mfrmClaim.GetCurrentBillingCountID(True)
    End With
    
    sSQL = "UPDATE RTActivityLog Set "
    sSQL = sSQL & "[RTActivityLogID] = " & pudtActLog.RTActivityLogID & ", "
    sSQL = sSQL & "[AssignmentsID] = " & pudtActLog.AssignmentsID & ", "
    sSQL = sSQL & "[BillingCountID] = " & pudtActLog.BillingCountID & ", "
    sSQL = sSQL & "[ID] = " & pudtActLog.ID & ", "
    sSQL = sSQL & "[IDAssignments] = " & pudtActLog.IDAssignments & ", "
    sSQL = sSQL & "[IDBillingCount] = " & pudtActLog.IDBillingCount & ", "
    sSQL = sSQL & "[ServiceTime] = " & pudtActLog.ServiceTime & ", "
    sSQL = sSQL & "[ActDate] = #" & pudtActLog.ActDate & "#, "
    sSQL = sSQL & "[ActText] = '" & goUtil.utCleanSQLString(pudtActLog.ActText) & "', "
    sSQL = sSQL & "[ActTime] = #" & pudtActLog.ActTime & "#, "
    sSQL = sSQL & "[PageBreakAfter] = " & pudtActLog.PageBreakAfter & ", "
    sSQL = sSQL & "[BlankPageAfter] = " & pudtActLog.BlankPageAfter & ", "
    sSQL = sSQL & "[BlankRowsAfter] = " & pudtActLog.BlankRowsAfter & ", "
    sSQL = sSQL & "[IsMgrEntry] = " & pudtActLog.IsMgrEntry & ", "
    sSQL = sSQL & "[IsDeleted] = " & pudtActLog.IsDeleted & ", "
    sSQL = sSQL & "[DownLoadMe] = " & pudtActLog.DownLoadMe & ", "
    sSQL = sSQL & "[UpLoadMe] = " & pudtActLog.UpLoadMe & ", "
    sSQL = sSQL & "[AdminComments] = '" & goUtil.utCleanSQLString(pudtActLog.AdminComments) & "', "
    sSQL = sSQL & "[DateLastUpdated] = #" & pudtActLog.DateLastUpdated & "#, "
    sSQL = sSQL & "[UpdateByUserID]  = " & pudtActLog.UpdateByUserID & " "
    sSQL = sSQL & "WHERE IDAssignments = " & pudtActLog.IDAssignments & " "
    sSQL = sSQL & "AND ID = " & pudtActLog.ID & " "
    
    Set oConn = New ADODB.Connection
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    oConn.Execute sSQL
    
    EditActLogItem = True
    'Clean up
    Set oConn = Nothing
    Exit Function
    
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function EditActLogItem"
End Function

Private Sub optAL01_sPresentDurringInspection_Click(Index As Integer)
    On Error GoTo EH
    If Not mbLoadingMe Then
        cmdSave.Enabled = True
    End If
    mfrmClaim.HighlightOpt optAL01_sPresentDurringInspection(GuiOptValues.No)
    mfrmClaim.HighlightOpt optAL01_sPresentDurringInspection(GuiOptValues.YES)
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub optAL01_sPresentDurringInspection_Click"
End Sub

Private Sub optAL02_sExplainedEstimate_Click(Index As Integer)
    On Error GoTo EH
    If Not mbLoadingMe Then
        cmdSave.Enabled = True
    End If
    cmdSave.Enabled = True
    mfrmClaim.HighlightOpt optAL02_sExplainedEstimate(GuiOptValues.No)
    mfrmClaim.HighlightOpt optAL02_sExplainedEstimate(GuiOptValues.YES)
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub optAL02_sExplainedEstimate_Click"
End Sub

Private Sub optAL03_sExplainedRCV_Click(Index As Integer)
    On Error GoTo EH
    If Not mbLoadingMe Then
        cmdSave.Enabled = True
    End If
    mfrmClaim.HighlightOpt optAL03_sExplainedRCV(GuiOptValues.No)
    mfrmClaim.HighlightOpt optAL03_sExplainedRCV(GuiOptValues.YES)
    mfrmClaim.HighlightOpt optAL03_sExplainedRCV(GuiOptValues.NA)
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub optAL03_sExplainedRCV_Click"
End Sub

Private Sub optAL04_sConfirmMortgageeIsCorrect_Click(Index As Integer)
    On Error GoTo EH
    If Not mbLoadingMe Then
        cmdSave.Enabled = True
    End If
    mfrmClaim.HighlightOpt optAL04_sConfirmMortgageeIsCorrect(GuiOptValues.No)
    mfrmClaim.HighlightOpt optAL04_sConfirmMortgageeIsCorrect(GuiOptValues.YES)
    mfrmClaim.HighlightOpt optAL04_sConfirmMortgageeIsCorrect(GuiOptValues.NA)
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub optAL04_sConfirmMortgageeIsCorrect_Click"
End Sub

Private Sub optAL05_sExplainedMortgageeChecks_Click(Index As Integer)
    On Error GoTo EH
    If Not mbLoadingMe Then
        cmdSave.Enabled = True
    End If
    mfrmClaim.HighlightOpt optAL05_sExplainedMortgageeChecks(GuiOptValues.No)
    mfrmClaim.HighlightOpt optAL05_sExplainedMortgageeChecks(GuiOptValues.YES)
    mfrmClaim.HighlightOpt optAL05_sExplainedMortgageeChecks(GuiOptValues.NA)
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub optAL05_sExplainedMortgageeChecks_Click"
End Sub

Private Sub optAL06_sConfirmedCoverage_Click(Index As Integer)
    On Error GoTo EH
    If Not mbLoadingMe Then
        cmdSave.Enabled = True
    End If
    mfrmClaim.HighlightOpt optAL06_sConfirmedCoverage(GuiOptValues.No)
    mfrmClaim.HighlightOpt optAL06_sConfirmedCoverage(GuiOptValues.YES)
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub optAL06_sConfirmedCoverage_Click"
End Sub

Private Sub optAL07_sPriorLoss_Click(Index As Integer)
    On Error GoTo EH
    If Not mbLoadingMe Then
        cmdSave.Enabled = True
    End If
    mfrmClaim.HighlightOpt optAL07_sPriorLoss(GuiOptValues.No)
    mfrmClaim.HighlightOpt optAL07_sPriorLoss(GuiOptValues.YES)
    mfrmClaim.HighlightOpt optAL07_sPriorLoss(GuiOptValues.NA)
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub optAL07_sPriorLoss_Click"
End Sub

Private Sub optAL08_sSalvage_Click(Index As Integer)
    On Error GoTo EH
    If Not mbLoadingMe Then
        cmdSave.Enabled = True
    End If
    mfrmClaim.HighlightOpt optAL08_sSalvage(GuiOptValues.No)
    mfrmClaim.HighlightOpt optAL08_sSalvage(GuiOptValues.YES)
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub optAL08_sSalvage_Click"
End Sub

Private Sub optAL09_sSubrogation_Click(Index As Integer)
    On Error GoTo EH
    If Not mbLoadingMe Then
        cmdSave.Enabled = True
    End If
    mfrmClaim.HighlightOpt optAL09_sSubrogation(GuiOptValues.No)
    mfrmClaim.HighlightOpt optAL09_sSubrogation(GuiOptValues.YES)
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub optAL09_sSubrogation_Click"
End Sub
