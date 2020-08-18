VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form AddPhoto 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Photo"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11595
   Icon            =   "AddPhoto.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   11595
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLoadAll 
      Caption         =   "&Load All"
      Height          =   855
      Left            =   9000
      TabIndex        =   28
      Top             =   600
      Width           =   1095
   End
   Begin VB.Frame framSortOrder 
      Height          =   5535
      Left            =   7320
      TabIndex        =   22
      Top             =   0
      Width           =   4215
      Begin VB.OptionButton optView 
         Caption         =   "Settings"
         Height          =   255
         Index           =   1
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   120
         Width           =   1215
      End
      Begin VB.OptionButton optView 
         Caption         =   "Photos"
         Height          =   255
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   120
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.Frame framCommands 
         Height          =   5055
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   3975
         Begin VB.FileListBox FilePhoto 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2610
            Left            =   120
            Pattern         =   "*.JPG;*.BMP"
            TabIndex        =   36
            Top             =   2370
            Width           =   3735
         End
         Begin VB.CommandButton cmdImportPhotoPath 
            Height          =   330
            Left            =   3480
            Picture         =   "AddPhoto.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   33
            ToolTipText     =   "Browse"
            Top             =   1410
            Width           =   375
         End
         Begin VB.TextBox txtImportPhotoPath 
            Height          =   360
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   32
            Tag             =   "Directory"
            Top             =   1395
            Width           =   3720
         End
         Begin VB.CheckBox chk35mm 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   2760
            Picture         =   "AddPhoto.frx":0486
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox txtSortOrder 
            DataField       =   "OutBuildingsFee"
            DataSource      =   "Claims"
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            MaxLength       =   3
            TabIndex        =   26
            Tag             =   "Numeric"
            Top             =   240
            Width           =   375
         End
         Begin VB.CommandButton cmdMenu 
            Caption         =   "&Delete"
            Height          =   855
            Index           =   2
            Left            =   360
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   240
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.CommandButton cmdMenu 
            Caption         =   "&Actual Size"
            Height          =   855
            Index           =   3
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.CheckBox chkDelOrigPhoto 
            Caption         =   "Delete Original Photo After Attach:"
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
            TabIndex        =   34
            Top             =   1800
            Width           =   3615
         End
         Begin VB.Label Label1 
            Caption         =   "Select Photo to Import from the Import Path"
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
            Left            =   120
            TabIndex        =   35
            Top             =   2160
            Width           =   3615
         End
         Begin VB.Label Label1 
            Caption         =   "Import Path: (Browse to Digital Camera)"
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
            Left            =   120
            TabIndex        =   31
            Top             =   1200
            Width           =   3720
         End
      End
      Begin VB.Frame framPhotoQuality 
         Caption         =   "Photo Quality"
         Height          =   1335
         Left            =   120
         TabIndex        =   37
         Top             =   360
         Visible         =   0   'False
         Width           =   3975
         Begin VB.CheckBox chkAlwaysUseDefaultQuality 
            Alignment       =   1  'Right Justify
            Caption         =   "Always Use Default Quality"
            Height          =   255
            Left            =   240
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   960
            Width           =   2415
         End
         Begin VB.CommandButton cmdDefaultPhotoQuality 
            Caption         =   "Default"
            Height          =   255
            Left            =   2760
            TabIndex        =   43
            Top             =   960
            Width           =   975
         End
         Begin MSComctlLib.Slider sldJPGQuality 
            Height          =   525
            Left            =   120
            TabIndex        =   41
            ToolTipText     =   "Save Quality "
            Top             =   435
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   926
            _Version        =   393216
            Min             =   40
            Max             =   100
            SelStart        =   50
            TickStyle       =   1
            TickFrequency   =   5
            Value           =   50
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H008080FF&
            Caption         =   "Maximum"
            Height          =   255
            Index           =   7
            Left            =   2520
            TabIndex        =   40
            Top             =   270
            Width           =   1335
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H0080FFFF&
            Caption         =   "Medium"
            Height          =   255
            Index           =   6
            Left            =   1440
            TabIndex        =   39
            Top             =   270
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H0080FF80&
            Caption         =   "Minimum"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   38
            Top             =   270
            Width           =   1335
         End
      End
      Begin VB.Frame framResize 
         Caption         =   "Resize % of Original (applies to photo upload only)"
         Height          =   1335
         Left            =   120
         TabIndex        =   44
         Top             =   1680
         Visible         =   0   'False
         Width           =   3975
         Begin VB.CheckBox chkAlwaysUseDefaultResize 
            Alignment       =   1  'Right Justify
            Caption         =   "Always Use Default Resize %"
            Height          =   255
            Left            =   240
            TabIndex        =   49
            TabStop         =   0   'False
            Top             =   960
            Width           =   2415
         End
         Begin VB.CommandButton cmdDefaultPhotoResize 
            Caption         =   "Default"
            Height          =   255
            Left            =   2760
            TabIndex        =   50
            Top             =   960
            Width           =   975
         End
         Begin MSComctlLib.Slider sldResize 
            Height          =   525
            Left            =   120
            TabIndex        =   48
            ToolTipText     =   "Save Resized"
            Top             =   435
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   926
            _Version        =   393216
            Min             =   10
            Max             =   100
            SelStart        =   50
            TickStyle       =   1
            TickFrequency   =   5
            Value           =   50
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "10% - 40%"
            Height          =   225
            Index           =   9
            Left            =   240
            TabIndex        =   45
            Top             =   210
            Width           =   1190
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "40% - 70%"
            Height          =   220
            Index           =   8
            Left            =   1420
            TabIndex        =   46
            Top             =   210
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "70% - 100%"
            Height          =   225
            Index           =   3
            Left            =   2520
            TabIndex        =   47
            Top             =   210
            Width           =   1215
         End
      End
      Begin VB.Frame framPhotoProperties 
         Caption         =   "Photo Properties"
         Height          =   2415
         Left            =   120
         TabIndex        =   51
         Top             =   3000
         Visible         =   0   'False
         Width           =   3975
         Begin VB.CheckBox chkHighlightSettings 
            Alignment       =   1  'Right Justify
            Caption         =   "Highlight Photo Properties"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   53
            TabStop         =   0   'False
            Top             =   465
            Width           =   3735
         End
         Begin VB.CheckBox chkViewSettings 
            Alignment       =   1  'Right Justify
            Caption         =   "Show Photo Properties"
            Height          =   255
            Left            =   120
            TabIndex        =   52
            TabStop         =   0   'False
            Top             =   240
            Width           =   3735
         End
      End
   End
   Begin VB.PictureBox picBox 
      BackColor       =   &H00000000&
      Height          =   5355
      Left            =   60
      MouseIcon       =   "AddPhoto.frx":34E0
      MousePointer    =   99  'Custom
      ScaleHeight     =   5295
      ScaleWidth      =   7080
      TabIndex        =   0
      ToolTipText     =   "Double Click to View Actual Size of Photo."
      Top             =   120
      Width           =   7140
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Upload Accessed:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   19
         Left            =   1800
         TabIndex        =   20
         Tag             =   "ViewSettings"
         Top             =   4800
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label lblUploadAccessedDateTime 
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   3960
         TabIndex        =   21
         Tag             =   "ViewSettings"
         Top             =   4800
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Upload Modified:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   18
         Left            =   1800
         TabIndex        =   18
         Tag             =   "ViewSettings"
         Top             =   4440
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label lblUploadModifiedDateTime 
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   3960
         TabIndex        =   19
         Tag             =   "ViewSettings"
         Top             =   4440
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Upload Created:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   17
         Left            =   1800
         TabIndex        =   16
         Tag             =   "ViewSettings"
         Top             =   4080
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label lblUploadCreatedDateTime 
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   3960
         TabIndex        =   17
         Tag             =   "ViewSettings"
         Top             =   4080
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Label lblCreatedDateTime 
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   3960
         TabIndex        =   5
         Tag             =   "ViewSettings"
         Top             =   1440
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Source Created:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   16
         Left            =   1800
         TabIndex        =   4
         Tag             =   "ViewSettings"
         Top             =   1440
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label lblModifiedDateTime 
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   3960
         TabIndex        =   7
         Tag             =   "ViewSettings"
         Top             =   1800
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Source Modified:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   15
         Left            =   1800
         TabIndex        =   6
         Tag             =   "ViewSettings"
         Top             =   1800
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label lblAccessedDateTime 
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   3960
         TabIndex        =   9
         Tag             =   "ViewSettings"
         Top             =   2160
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Source Accessed:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   14
         Left            =   1800
         TabIndex        =   8
         Tag             =   "ViewSettings"
         Top             =   2160
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Image Image3 
         Height          =   5250
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   7050
      End
      Begin VB.Label lblHandWCurrentUpload 
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   3960
         TabIndex        =   15
         Tag             =   "ViewSettings"
         Top             =   3720
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Upload Size:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   13
         Left            =   1800
         TabIndex        =   14
         Tag             =   "ViewSettings"
         Top             =   3720
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Resize (approximate):"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   12
         Left            =   1800
         TabIndex        =   12
         Tag             =   "ViewSettings"
         Top             =   3120
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Source Size:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   11
         Left            =   1800
         TabIndex        =   2
         Tag             =   "ViewSettings"
         Top             =   1080
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Optimal Size:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   10
         Left            =   1800
         TabIndex        =   10
         Tag             =   "ViewSettings"
         Top             =   2760
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label lblHandWOptimal 
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   3960
         TabIndex        =   11
         Tag             =   "ViewSettings"
         Top             =   2760
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Label lblHandWResized 
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   3960
         TabIndex        =   13
         Tag             =   "ViewSettings"
         Top             =   3120
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Label lblHandWOriginal 
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   3960
         TabIndex        =   3
         Tag             =   "ViewSettings"
         Top             =   1080
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Image Image2 
         Height          =   5250
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   7050
      End
      Begin VB.Label lblFileName 
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Tag             =   "ViewSettings"
         Top             =   45
         Visible         =   0   'False
         Width           =   6855
      End
      Begin VB.Image Image1 
         Height          =   5250
         Left            =   0
         MouseIcon       =   "AddPhoto.frx":37EA
         MousePointer    =   99  'Custom
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to View Actual Size of Photo."
         Top             =   0
         Width           =   7050
      End
   End
   Begin VB.Frame framSelection 
      BorderStyle     =   0  'None
      Height          =   3480
      Left            =   120
      TabIndex        =   62
      Top             =   3960
      Width           =   11655
      Begin VB.CommandButton cmdMenu 
         Caption         =   "Sa&ve"
         Default         =   -1  'True
         Height          =   855
         Index           =   0
         Left            =   9360
         MaskColor       =   &H00000000&
         Picture         =   "AddPhoto.frx":3AF4
         Style           =   1  'Graphical
         TabIndex        =   60
         ToolTipText     =   "Exit"
         Top             =   2565
         Width           =   975
      End
      Begin VB.CommandButton cmdMenu 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   855
         Index           =   1
         Left            =   10440
         MaskColor       =   &H00000000&
         Picture         =   "AddPhoto.frx":3C3E
         Style           =   1  'Graphical
         TabIndex        =   61
         ToolTipText     =   "Exit"
         Top             =   2565
         Width           =   975
      End
      Begin VB.CommandButton cmdAssignedDate 
         Height          =   375
         Left            =   6720
         Picture         =   "AddPhoto.frx":3F48
         Style           =   1  'Graphical
         TabIndex        =   58
         Tag             =   "Date"
         Top             =   1995
         Width           =   375
      End
      Begin VB.TextBox txtPhotoDate 
         Height          =   360
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   57
         Tag             =   "Date"
         Top             =   1995
         Width           =   1800
      End
      Begin VB.CommandButton cmdSpelling 
         Caption         =   "Spellin&g"
         Height          =   855
         Left            =   8280
         MaskColor       =   &H00000000&
         Picture         =   "AddPhoto.frx":438A
         Style           =   1  'Graphical
         TabIndex        =   59
         ToolTipText     =   "Exit"
         Top             =   2565
         Width           =   975
      End
      Begin VB.TextBox txtDescription 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1365
         Left            =   -15
         MaxLength       =   254
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   56
         Top             =   1995
         Width           =   5175
      End
      Begin VB.Label lblWarning 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1080
         TabIndex        =   54
         Top             =   1515
         Width           =   6135
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Description:"
         Height          =   255
         Index           =   1
         Left            =   -240
         TabIndex        =   55
         Top             =   1755
         Width           =   1095
      End
   End
End
Attribute VB_Name = "AddPhoto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmClaim As frmClaim
Private mfrmPhotos As frmPhotos
Private moGUI As V2ECKeyBoard.clsCarGUI
Private msAssignmentsID As String
Private msIDRTPhotoReport As String
Private msIBNUM As String

Private mbAdding As Boolean
Private msPhotoID As String
Private mbLoading As Boolean
Private mlMaxSort As Long
Private mlPhotoCount As Long
Private mbFirstSave As Boolean
'DeRes Stuff
Private msTimeStamp As String
Private moDeRes As V2ECKeyBoard.clsLists
'Editing a Deresed photo
'need to delete it upon saveing over it
Private msEditDeResPath As String
Private mbQualityWarning As Boolean
Private mbResizeMaxWarning As Boolean
Private mbResizeMinWarning As Boolean
Private mlOriginal_H As Long
Private mlOriginal_W As Long
Private mbInitResize As Boolean
Private mbInitINI As Boolean
Private mbZoom100 As Boolean
Private mbSavedPhoto As Boolean
Private mbExitWithoutRefresh As Boolean
Private mbShowActualSize As Boolean
Private mbLoadingAll As Boolean
Private mbSavingPhoto As Boolean
Private mitmX As MSComctlLib.ListItem
Private moListView As MSComctlLib.ListView
Private moFI As V2ECKeyBoard.clsFileVersion
Private mFI As V2ECKeyBoard.FILE_INFORMATION
Private mFA As V2ECKeyBoard.FILE_ATTRIBUTES
Private Const ZOOM_WARNING As String = "ZOOMING PHOTO... TO UNZOOM CLICK THE PHOTO AGAIN!"
Private Const Optimal_H As Long = 5250
Private Const Optimal_W As Long = 7050
Private Const DefaultPhotoQuality As Long = 80
Private Const SE_ERR_NOASSOC = 31
'Private Const MAX_PHOTOS_ALLOWED As Long = 20
'The default for PhotoResize is really temporary
'as this value will be figured out by ResizeToOptimal
Private Const DefaultPhotoResize As Long = 75
'Option View
Private Const OPT_VIEW_PHOTOS As Long = 0
Private Const OPT_VIEW_SETTINGS As Long = 1


Private Declare Function GetDesktopWindow Lib "user32" () As Long

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

Public Property Let IDRTPhotoReport(psIDRTPhotoReport As String)
    msIDRTPhotoReport = psIDRTPhotoReport
End Property
Public Property Get IDRTPhotoReport() As String
    IDRTPhotoReport = msIDRTPhotoReport
End Property

Public Property Let IBNUM(psIBNUM As String)
    msIBNUM = psIBNUM
End Property
Public Property Get IBNUM() As String
    IBNUM = msIBNUM
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

Public Property Let MyPhotos(poForm As Object)
    Set mfrmPhotos = poForm
End Property
Public Property Set MyPhotos(poForm As Object)
    Set mfrmPhotos = poForm
End Property
Public Property Get MyPhotos() As Object
    Set MyPhotos = mfrmPhotos
End Property

Public Property Let EditDeResPath(psPath As String)
    msEditDeResPath = psPath
End Property
Public Property Get EditDeResPath() As String
    EditDeResPath = msEditDeResPath
End Property

Public Property Let MaxSort(plSort As Long)
    mlMaxSort = plSort
End Property
Public Property Get MaxSort() As Long
    MaxSort = mlMaxSort
End Property

Public Property Let PhotoCount(plCount As Long)
    mlPhotoCount = plCount
End Property
Public Property Get PhotoCount() As Long
    MaxSort = mlPhotoCount
End Property

Public Property Let Loading(pbFlag As Boolean)
    mbLoading = pbFlag
End Property

Public Property Let PhotoID(psActId As String)
    msPhotoID = psActId
End Property

Public Property Let Adding(pbFlag As Boolean)
    mbAdding = pbFlag
End Property

Private Property Get msClassName() As String
    msClassName = Me.Name
End Property

Private Sub chk35mm_Click()
    On Error GoTo EH
    
    If chk35mm.Value = vbChecked Then
        FilePhoto.Enabled = False
        cmdLoadAll.Enabled = False
        lblFileName.Caption = goUtil.gsInstallDir & "\Icons\35mm.bmp"
        If goUtil.utFileExists(lblFileName.Caption) Then
            Image1 = LoadPicture(lblFileName.Caption)
            SetOriginalPhoto
        End If
    Else
        FilePhoto.Enabled = True
        cmdLoadAll.Enabled = True
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub chk35mm_Click"
End Sub

Private Sub chkAlwaysUseDefaultQuality_Click()
    On Error GoTo EH
    SaveSetting App.EXEName, "GENERAL", "PHOTO_AlwaysUseDefaultQuality", chkAlwaysUseDefaultQuality.Value
    If chkAlwaysUseDefaultQuality.Value = vbChecked Then
        sldJPGQuality.Enabled = False
        cmdDefaultPhotoQuality_Click
    Else
        MsgBox "It is highly recommended that you leave this option checked!", vbExclamation, "Override Default Settings"
        sldJPGQuality.Enabled = True
    End If
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub chkAlwaysUseDefaultQuality_Click"
End Sub

Private Sub chkAlwaysUseDefaultResize_Click()
    On Error GoTo EH
    SaveSetting App.EXEName, "GENERAL", "PHOTO_AlwaysUseDefaultResize", chkAlwaysUseDefaultResize.Value
    If chkAlwaysUseDefaultResize.Value = vbChecked Then
        sldResize.Enabled = False
        cmdDefaultPhotoResize_Click
    Else
        MsgBox "It is highly recommended that you leave this option checked!", vbExclamation, "Override Default Settings"
        sldResize.Enabled = True
    End If
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub chkAlwaysUseDefaultResize_Click"
End Sub

Private Sub chkDelOrigPhoto_Click()
    On Error GoTo EH
    If mbLoading Then
        Exit Sub
    End If
    Dim bDelOrigPhoto As Boolean
    
    If chkDelOrigPhoto.Value = vbChecked Then
        bDelOrigPhoto = True
    Else
        bDelOrigPhoto = False
    End If
    
    SaveSetting App.EXEName, "GENERAL", "DELETE_ORIG_PHOTO", bDelOrigPhoto
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub chkDelOrigPhoto_Click"
End Sub

Private Sub chkHighlightSettings_Click()
    On Error GoTo EH
    Dim lBackStyle As Long
    Dim lForeColor As Long
    Dim oControl As Control
    SaveSetting App.EXEName, "GENERAL", "PHOTO_HighlightSettings", chkHighlightSettings.Value
    If chkHighlightSettings.Value = vbChecked Then
        lBackStyle = 1
        lForeColor = &H80000017
    Else
        lBackStyle = 0
        lForeColor = &HFFFFFF
    End If
    
    For Each oControl In Me.Controls
        If TypeOf oControl Is Label Then
            If oControl.Tag = "ViewSettings" Then
                oControl.BackStyle = lBackStyle
                oControl.ForeColor = lForeColor
            End If
        End If
    Next
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub chkHighlightSettings_Click"
End Sub

Private Sub chkViewSettings_Click()
    On Error GoTo EH
    Dim bViewSettings As Boolean
    Dim oControl As Control
    SaveSetting App.EXEName, "GENERAL", "PHOTO_ViewSettings", chkViewSettings.Value
    If chkViewSettings.Value = vbChecked Then
        bViewSettings = True
        chkHighlightSettings.Enabled = True
    Else
        bViewSettings = False
        chkHighlightSettings.Enabled = False
    End If
    
    For Each oControl In Me.Controls
        If TypeOf oControl Is Label Then
            If oControl.Tag = "ViewSettings" Then
                oControl.Visible = bViewSettings
            End If
        End If
    Next
    chkHighlightSettings_Click
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub chkViewSettings_Click"
End Sub

Private Sub cmdAssignedDate_Click()
    On Error GoTo EH
    mfrmClaim.ShowCalendar txtPhotoDate
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdAssignedDate_Click"
End Sub

Private Sub cmdDefaultPhotoQuality_Click()
    On Error GoTo EH
    
    sldJPGQuality.Value = DefaultPhotoQuality
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdDefaultPhotoQuality_Click"
End Sub

Private Sub cmdDefaultPhotoResize_Click()
    On Error GoTo EH
    ResizeToOptimal
    DisplayResize
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdDefaultPhotoResize_Click"
End Sub

Private Sub cmdImportPhotoPath_Click()
    On Error GoTo EH
    Dim sPath As String
    Dim sSelFile As String
    Dim lCount As Long
    sPath = txtImportPhotoPath.Text
    sPath = goUtil.utGetPath(App.EXEName, "ImportPhotoPath", "Browse to the photo(s) you want to import", "CLICK OPEN TO SAVE PATH", sPath, Me.hWnd, , sSelFile)
    If goUtil.utFileExists(sPath, True) Then
        If StrComp(sPath, goUtil.PhotoReposPath, vbTextCompare) = 0 Then
            MsgBox "Can't use this directory for attaching Photos!", vbExclamation + vbOKOnly, "INVALID DIRECTORY!"
            txtImportPhotoPath.Text = vbNullString
            Exit Sub
        End If
        SaveSetting App.EXEName, "GENERAL", "PHOTO_PATH", sPath
        txtImportPhotoPath.Text = sPath
        FilePhoto.Path = sPath
        If sSelFile <> vbNullString Then
            For lCount = 0 To FilePhoto.ListCount - 1
                If FilePhoto.List(lCount) = sSelFile Then
                    FilePhoto.ListIndex = lCount
                    Exit For
                End If
            Next
        End If
    End If
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdImportPhotoPath_Click"
End Sub

Private Sub cmdMenu_Click(Index As Integer)
    On Error GoTo EH
    Dim oViewImg As V2ECKeyBoard.clsLists
    Dim lRet As Long
    Dim sDir As String
    Dim sFile As String
    
  Select Case Index
    Case 0 ' save
        Screen.MousePointer = vbHourglass
        
        'Validate a couple of control here
        goUtil.utValidate , txtDescription
        goUtil.utValidate , txtPhotoDate
        goUtil.utValidate , txtSortOrder
        
        If chk35mm.Value = vbUnchecked Then
            '195  6.17.2002 Editing Photo Description
            'Added Or Not mbAdding
            If FilePhoto.ListCount > 0 Or Not mbAdding Then
                If mbFirstSave Then
                    mlMaxSort = mlMaxSort + 1
                    mlPhotoCount = mlPhotoCount + 1
                End If
                If mbAdding Then
                    If FilePhoto.ListIndex < 0 Then
                        FilePhoto.ListIndex = 0
                    End If
                End If
                cmdMenu(0).Enabled = False
                cmdMenu(1).Enabled = False
                SavePhoto
                'Need to Set the FilePhoto List because the original photo(s)
                'Could be removed as they are being added
                If chkDelOrigPhoto.Value = vbChecked Then
                    SetFilePhoto
                End If
                cmdMenu(0).Enabled = True
                cmdMenu(1).Enabled = True
                If Not goUtil.utFormExists(Forms, "AddPhoto") Then
                    Exit Sub
                End If
            End If
        Else
            cmdMenu(0).Enabled = False
            SavePhoto
            cmdMenu(0).Enabled = True
            chk35mm.Value = vbUnchecked
        End If
        
        If mbAdding Then
            If txtDescription.Visible And txtDescription.Enabled Then
                txtDescription.SetFocus
            End If
        End If
        Screen.MousePointer = vbDefault
    Case 1 ' Quit
        If Not mbSavedPhoto Then
            mbExitWithoutRefresh = True
        End If
        goUtil.utFormWinRegPos goUtil.gsMainAppEXEName, Me, True, , , , True
        Me.Visible = False
    Case 2 ' Delete
        If mfrmPhotos.DeletePhotoItem(msPhotoID) Then
            If moListView Is Nothing Then
                Set moListView = mfrmPhotos.lstvPhotos
            End If
            mfrmPhotos.LoadMe
        End If
        goUtil.utFormWinRegPos goUtil.gsMainAppEXEName, Me, True, , , , True
        Me.Visible = False
    Case 3 'View Actual Size
        If mbShowActualSize Then
            Exit Sub
        End If
        mbShowActualSize = True
        cmdMenu(3).Enabled = False
        'Don't show the actual file, since don't want them
        'to edit and save that photo outside of the Application!!
        sFile = lblFileName.Caption
        If goUtil.utFileExists(sFile) Then
            goUtil.utCopyFile sFile, App.Path & "\TempPhoto.jpg"
            sFile = App.Path & "\TempPhoto.jpg"
            lRet = goUtil.utShellExecute(GetDesktopWindow, "OPEN", sFile, vbNullString, App.Path, vbNormalFocus)
                'Check to see if the Associated Application opened the file
                'If not that means there is no application associated with the file
                'in that case open Explorer to give the user a chance to open the file
            If lRet = SE_ERR_NOASSOC Then
                sDir = goUtil.utGetSystemDir
                lRet = goUtil.utShellExecute(GetDesktopWindow, vbNullString, "RUNDLL32.EXE", "shell32.dll,OpenAs_RunDLL " & sFile, sDir, vbNormalFocus)
            End If
        Else
            MsgBox "File Not Found!", vbExclamation + vbOKOnly, "File Not Found"
        End If
  End Select
    Exit Sub
EH:
    Screen.MousePointer = vbDefault
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdMenu_Click"
End Sub


Private Sub cmdLoadAll_Click()
    On Error GoTo EH
    Dim lCount As Long
    Dim bSkipUnload As Boolean
    
    If mbLoadingAll Then
        If cmdLoadAll.Caption = "&CANCEL" Then
            mbLoadingAll = False
        End If
        cmdLoadAll.Enabled = False
        Exit Sub
    End If
    
    If FilePhoto.ListCount > 0 Then
        If FilePhoto.ListIndex <= 0 Then
            FilePhoto.ListIndex = 0
        Else
            If MsgBox("Do you want to start Loading from the Top ?", vbYesNo, "Start From The Top?") = vbYes Then
                FilePhoto.ListIndex = 0
            Else
                lCount = FilePhoto.ListIndex
            End If
        End If
        
        cmdLoadAll.Caption = "&CANCEL"
        cmdLoadAll.Cancel = True
        mbLoadingAll = True
        mlMaxSort = mlMaxSort + 1
        mlPhotoCount = mlPhotoCount + 1
        framSortOrder.Enabled = False
        framSelection.Enabled = False
        Do
            DoEvents
            Sleep 10
            If Not Me.Visible Then
                Exit Sub
            End If
            txtSortOrder.Text = vbNullString
            SavePhoto
            lCount = lCount + 1
            If lCount > FilePhoto.ListCount - 1 Or Not mbLoadingAll Then
                If Not mbLoadingAll Then
                    bSkipUnload = True
                    mlMaxSort = mlMaxSort - 1
                    mlPhotoCount = mlPhotoCount + 1
                End If
                Exit Do
            End If
        Loop
        mbLoadingAll = False
        framSortOrder.Enabled = True
        framSelection.Enabled = True
        If Not bSkipUnload Then
            goUtil.utFormWinRegPos goUtil.gsMainAppEXEName, Me, True, , , , True
            Me.Visible = False
        Else
            cmdLoadAll.Enabled = True
            cmdLoadAll.Caption = "&Load All"
            cmdMenu(1).Cancel = True
        End If
    End If
    
    'Need to Set the FilePhoto List because the original photo(s)
    'Could be removed as they are being added
    If chkDelOrigPhoto.Value = vbChecked Then
        SetFilePhoto
    End If
    
  Exit Sub
EH:
    mbLoadingAll = False
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdLoadAll_Click"
End Sub

Private Sub cmdSpelling_Click()
    On Error GoTo EH
    
    goUtil.utLoadSP
    goUtil.goSP.CheckSP txtDescription
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdSpelling_Click"
End Sub


Private Sub FilePhoto_Click()
    On Error GoTo EH
    Dim sHandW As String
    
    lblFileName = Replace(txtImportPhotoPath.Text & "\" & FilePhoto.FileName, "\\", "\")
    
    If goUtil.utFileExists(lblFileName.Caption) Then
        Image1 = LoadPicture(lblFileName.Caption)
        SetOriginalPhoto
        mbShowActualSize = False
        cmdMenu(3).Enabled = True
    End If
      
   Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub FilePhoto_Click"
End Sub

Private Sub ResizeToOptimal()
    On Error GoTo EH
    mbInitResize = True
    Dim lNewSizePercent As Long
    'Use the Optimal and original settings to figure out the Resize percentage
    
    'Equation to figure out the New Percentage
    If mlOriginal_H > 0 Then
        lNewSizePercent = Optimal_H / mlOriginal_H * 100
    End If

    sldResize.Value = lNewSizePercent
    mbInitResize = False
    Exit Sub
EH:
    mbInitResize = False
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub ResizeToOptimal"
End Sub

Private Sub FilePhoto_DblClick()
  lblFileName = Replace(txtImportPhotoPath.Text & "\" & FilePhoto.FileName, "\\", "\")
  On Error Resume Next
    If goUtil.utFileExists(lblFileName.Caption) Then
        Image1 = LoadPicture(lblFileName.Caption)
    End If
  On Error GoTo 0
  txtDescription = vbNullString
  txtDescription.SetFocus
End Sub

Private Sub FilePhoto_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = KeyCodeConstants.vbKeyReturn Then
    Call FilePhoto_DblClick
  End If
End Sub

Private Sub Form_Load()
    On Error GoTo EH
    Dim bDelOrigPhoto As Boolean
    
    mbLoading = True
    
    'Set the Form Icon
    Me.Icon = mfrmClaim.optClaim(GuiClaimOptions.opt05_Photos).Picture
    
    goUtil.utFormWinRegPos goUtil.gsMainAppEXEName, Me, , , , , True
    
    'Set Check box for Deleting original PHOTO after Attach
    bDelOrigPhoto = CBool(GetSetting(App.EXEName, "GENERAL", "DELETE_ORIG_PHOTO", False))
    If bDelOrigPhoto Then
        chkDelOrigPhoto.Value = vbChecked
    Else
        chkDelOrigPhoto.Value = vbUnchecked
    End If
    
    Image1.Stretch = True
    mbFirstSave = True
    SetFilePhoto
    
    '10.21.2002 Set jpg quality slider bar value
    sldJPGQuality.Value = GetSetting(App.EXEName, "GENERAL", "PHOTO_QUALITY", DefaultPhotoQuality)
    '2.24.2004 Set the Resize slider bar value
    sldResize.Value = GetSetting(App.EXEName, "GENERAL", "PHOTO_RESIZE", DefaultPhotoResize)
    chkAlwaysUseDefaultQuality.Value = GetSetting(App.EXEName, "GENERAL", "PHOTO_AlwaysUseDefaultQuality", vbChecked)
    chkAlwaysUseDefaultResize.Value = GetSetting(App.EXEName, "GENERAL", "PHOTO_AlwaysUseDefaultResize", vbChecked)
    chkHighlightSettings.Value = GetSetting(App.EXEName, "GENERAL", "PHOTO_HighlightSettings", vbUnchecked)
    chkViewSettings.Value = GetSetting(App.EXEName, "GENERAL", "PHOTO_ViewSettings", vbUnchecked)
    mbZoom100 = True
    mbInitINI = True
    
    mbLoading = False
  Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_Load"
End Sub

Public Function SetFilePhoto() As Boolean
    On Error GoTo EH
    Dim sPhotoPath As String
    sPhotoPath = GetSetting(App.EXEName, "GENERAL", "PHOTO_PATH", vbNullString)
    If goUtil.utFileExists(sPhotoPath, True) Then
        FilePhoto.Path = sPhotoPath
        FilePhoto.Refresh
        lblFileName = sPhotoPath
        txtImportPhotoPath.Text = sPhotoPath
    End If
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function SetFilePhoto"
End Function


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo EH
    
    Select Case UnloadMode
        Case vbFormControlMenu
            If mbLoadingAll Or mbSavingPhoto Then
                Cancel = True
                If mbLoadingAll Then
                    If cmdLoadAll.Caption = "&CANCEL" Then
                        mbLoadingAll = False
                    End If
                    cmdLoadAll.Enabled = False
                End If
            Else
                goUtil.utFormWinRegPos goUtil.gsMainAppEXEName, Me, True, , , , True
                Me.Visible = False
                Cancel = True
            End If
        Case Else
            If mbLoadingAll Or mbSavingPhoto Then
                Cancel = True
            End If
    End Select
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_QueryUnload"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo EH
    Dim bSelectItmx As Boolean
    Screen.MousePointer = vbHourglass
    If Not mbExitWithoutRefresh Then
        If StrComp(msEditDeResPath, lblFileName.Caption, vbTextCompare) <> 0 Then
            Sleep 1000
        End If
        mfrmPhotos.RefreshPhotos
        'Now Select the Photo That was Last Edited
        If Not moListView Is Nothing Then
            For Each mitmX In moListView.ListItems
                If mitmX.SubItems(GuiPhotoListView.ID - 1) = msPhotoID Then
                    bSelectItmx = True
                    Exit For
                End If
            Next
            If bSelectItmx Then
                mitmX.Selected = True
                mitmX.EnsureVisible
            End If
        End If
    End If
    CLEANUP
    Screen.MousePointer = vbDefault
    Exit Sub
EH:
    Screen.MousePointer = vbDefault
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_Unload"
End Sub




Private Sub Image1_DblClick()
    On Error GoTo EH
    
    cmdMenu_Click 3
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Image1_DblClick"
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mbZoom100 = Not mbZoom100
    DisplayResize
End Sub


Private Sub optView_Click(Index As Integer)
    On Error GoTo EH
    
    Select Case Index
        Case OPT_VIEW_PHOTOS
            framCommands.Visible = True
            framPhotoQuality.Visible = False
            framResize.Visible = False
            cmdLoadAll.Visible = True
            framPhotoProperties.Visible = False
        Case OPT_VIEW_SETTINGS
            framCommands.Visible = False
            framPhotoQuality.Visible = True
            framResize.Visible = True
            cmdLoadAll.Visible = False
            framPhotoProperties.Visible = True
    End Select
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub optView_Click"
End Sub

Private Sub picBox_DblClick()
    On Error GoTo EH
    
    cmdMenu_Click 3
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub picBox_DblClick"
End Sub

Private Sub picBox_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mbZoom100 = Not mbZoom100
    DisplayResize
End Sub

Private Sub sldJPGQuality_Change()
    On Error GoTo EH
    Dim sWarning As String
    
    SaveSetting App.EXEName, "GENERAL", "PHOTO_QUALITY", sldJPGQuality.Value
    mbInitINI = True
    If sldJPGQuality.Value >= 81 Then
        If Me.Visible Then
            sWarning = "Maximum Photo Quality = Longer Upload Times!"
            If Not mbQualityWarning Then
                MsgBox sWarning, vbExclamation, "Warning"
                mbQualityWarning = True
            End If
            lblWarning.Caption = sWarning
            lblWarning.Refresh
        End If
    End If
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub sldJPGQuality_Change"
End Sub

Private Sub sldResize_Change()
    On Error GoTo EH
    Dim sWarning As String
    If mbInitResize Then
        Exit Sub
    End If
    SaveSetting App.EXEName, "GENERAL", "PHOTO_RESIZE", sldResize.Value
    If sldResize.Value >= 81 Then
        If Me.Visible Then
            sWarning = "Maximum Resize Percentage = Longer Upload Times!"
            If Not mbResizeMaxWarning Then
'                MsgBox sWarning, vbExclamation, "Warning"
                mbResizeMaxWarning = True
            End If
            lblWarning.Caption = sWarning
            lblWarning.Refresh
        End If
    End If
    If sldResize.Value <= 59 Then
        If Me.Visible Then
            sWarning = "Minimum Resize Percentage can result in poor photo quality!"
            If Not mbResizeMinWarning Then
'                MsgBox sWarning, vbExclamation, "Warning"
                mbResizeMinWarning = True
            End If
            lblWarning.Caption = sWarning
            lblWarning.Refresh
        End If
    End If
    
    'Display the Resized Image
    DisplayResize
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub sldResize_Change"
End Sub

Private Sub DisplayResize()
    On Error GoTo EH
    Dim sMess As String
    
    Image1.Stretch = True
    If sldResize.Value = 100 Or mbZoom100 Then
        Image1.Width = Optimal_W
        Image1.left = -10
        Image1.Height = Optimal_H
        Image1.top = -10
        
        If Not mbZoom100 Then
            'Update the Resized Dims of the actual photo
            lblHandWResized.Caption = "(Width = " & ConvertTwipsToPixels(Image2.Width) & " Height = " & ConvertTwipsToPixels(Image2.Height) & ")"
        Else
            lblHandWResized.Caption = "ZOOMIMG PHOTO"
            sMess = ZOOM_WARNING
            lblWarning.Caption = sMess
        End If
        lblHandWResized.Refresh
    Else
        Image1.Width = (Optimal_W * (sldResize.Value / 100))
        Image1.left = -10 + (Optimal_W - (Optimal_W * (sldResize.Value / 100))) / 2
        Image1.Height = (Optimal_H * (sldResize.Value / 100))
        Image1.top = -10 + (Optimal_H - (Optimal_H * (sldResize.Value / 100))) / 2
        lblHandWResized.Caption = "(Width = " & ConvertTwipsToPixels(Round(mlOriginal_W * (sldResize.Value / 100), 0)) & " Height = " & ConvertTwipsToPixels(Round(mlOriginal_H * (sldResize.Value / 100), 0)) & ")"
        lblHandWResized.Refresh
        If lblWarning.Caption = ZOOM_WARNING Then
            lblWarning.Caption = vbNullString
        End If
    End If
    lblWarning.Refresh
    Image1.Refresh
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub DisplayResize"
End Sub

Private Sub txtDescription_Change()
    On Error GoTo EH
    
    
    mfrmClaim.RemoveVBCRLF txtDescription
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub txtDescription_Change"
End Sub

Private Sub txtDescription_GotFocus()
    goUtil.utSelText txtDescription
End Sub

Private Sub txtDescription_LostFocus()
    goUtil.utValidate , txtDescription
End Sub

Private Sub txtImportPhotoPath_GotFocus()
    goUtil.utSelText txtImportPhotoPath
End Sub

Private Sub txtImportPhotoPath_LostFocus()
    goUtil.utValidate , txtImportPhotoPath
End Sub

Private Sub txtPhotoDate_GotFocus()
    goUtil.utSelText txtPhotoDate
End Sub

Private Sub txtPhotoDate_LostFocus()
    goUtil.utValidate , txtPhotoDate
End Sub

Private Sub txtSortOrder_GotFocus()
    goUtil.utSelText txtSortOrder
End Sub

Private Function SavePhoto() As Boolean
    On Error GoTo EH
    Dim udtPhoto As GuiPhotoItem
    Dim udtDeResPhoto As V2ECKeyBoard.udtDeResPhoto
    Dim sTimeStamp As String
    Dim sBuildPath As String
    Dim sPhotoName As String
    Dim sTempPhotoPath As String
    Dim sMess As String
    'Parameters in udtDeResPhoto.colParameters
    Dim plReSizePercent As Long
    Dim plPhotoQuality As Long
    Dim plOptimalPixelWidth As Long
    Dim plOptimalPixelHeight As Long
    Dim sFlagText As String
    Dim bUploadPhoto As Boolean
    Dim lIDRTPhotoReport As Long
    
    mbSavingPhoto = True
    
    
    'Need to see if the Maximum Number of Photos Allowed has been reached
    If mlPhotoCount > mfrmPhotos.MAX_PHOTOS_ALLOWED Then
        'BGS Making GUI Changes to Photos 1.10.2005 Per Rob Petrovics and Elizabeth Warner-Simpson Request
        If mfrmPhotos.cmdAddMultiReport.Visible Then
            'If the Add multi report is visible then manually adding photo reports
            sMess = "A Maximum of " & mfrmPhotos.MAX_PHOTOS_ALLOWED & " photos are allowed to be entered per Photo Report!" & vbCrLf
            sMess = sMess & "Please add another Photo Report if you need to add more photos!" & vbCrLf
            MsgBox sMess, vbExclamation + vbOKOnly, "MAX OF " & mfrmPhotos.MAX_PHOTOS_ALLOWED & " ALLOWED!"
            If mbLoadingAll Then
                If cmdLoadAll.Caption = "&CANCEL" Then
                    mbLoadingAll = False
                    cmdLoadAll.Caption = "&Load All"
                    cmdMenu(1).Cancel = True
                End If
            End If
            mbSavingPhoto = False
            cmdMenu_Click 1
            Exit Function
        Else
            'Need to Add the Next Photo Report
            If Not mfrmPhotos.AddNextPhotoReport(lIDRTPhotoReport) Then
                If mbLoadingAll Then
                    If cmdLoadAll.Caption = "&CANCEL" Then
                        mbLoadingAll = False
                        cmdLoadAll.Caption = "&Load All"
                        cmdMenu(1).Cancel = True
                    End If
                End If
                mbSavingPhoto = False
                cmdMenu_Click 1
            Else
                IDRTPhotoReport = lIDRTPhotoReport
            End If
        End If
    End If
    
    lblFileName.Caption = UCase(lblFileName.Caption)
    
    'Save the current photo to TempFile to be processed
    
    'Set up the Deres UDT
    udtDeResPhoto.sPhotoPath = lblFileName.Caption
    
    'Only do the time stamp if adding or editing a photo that
    'changes the source photo
    If StrComp(msEditDeResPath, lblFileName.Caption, vbTextCompare) <> 0 Then
        'Set the time stamp used to build Photo Paths
        sTimeStamp = Format(Now, "YYMMDDHHMMSS")
        DoEvents
        Sleep 1000
        'Set the DeRes Stuff here
        If moDeRes Is Nothing Then
            Set moDeRes = New V2ECKeyBoard.clsLists
        End If
        
        'Set the Build path for the de res photo
        sPhotoName = msIBNUM & "_" & sTimeStamp & "_"
        sBuildPath = goUtil.PhotoReposPath & sPhotoName
        udtDeResPhoto.sBuildPath = sBuildPath
        'Add parametres to it
        Set udtDeResPhoto.colParameters = New Collection
        udtDeResPhoto.colParameters.Add mbInitINI, "pbInitINI"
        mbInitINI = False
        udtDeResPhoto.colParameters.Add sldJPGQuality.Value, "plPhotoQuality"
        udtDeResPhoto.colParameters.Add sldResize.Value, "plReSizePercent"
        udtDeResPhoto.colParameters.Add goUtil.ConvertTwipsToPixels(Optimal_W), "plOptimalPixelWidth"
        udtDeResPhoto.colParameters.Add goUtil.ConvertTwipsToPixels(Optimal_H), "plOptimalPixelHeight"
    End If
    
    
    If mbAdding Then
        With udtPhoto
            .RTPhotoLogID = "null" ' not set here
            .RTPhotoReportID = "null" ' not set here
            .AssignmentsID = "null"  ' not set here
            .BillingCountID = "null"  'not set here
            .ID = "null"  'not set here
            .IDRTPhotoReport = "null"   'not set here
            .IDAssignments = "null"  'not set here
            .IDBillingCount = "null"  'not set here
            .PhotoDate = txtPhotoDate.Text
            If Val(txtSortOrder.Text) > 0 Then
                .SortOrder = txtSortOrder.Text
            Else
                .SortOrder = mlMaxSort
            End If
            .Description = txtDescription.Text
            .PhotoName = sPhotoName & "1.jpg"
            .Photo = vbNullString 'Not set here
            'Set Flag for Upload (Highres = 0, Normal = 1, Thumb = 2)
            .DownloadPhoto = "False"
            .UpLoadPhoto = "True"
            .PhotoThumb = vbNullString 'Not set here
            .DownloadPhotoThumb = "False"
            .UpLoadPhotoThumb = "True"
            .PhotoHighRes = vbNullString 'Not set here
            .DownloadPhotoHighRes = "False"
            .UploadPhotoHighRes = "False" ' This option is not yet supported 7.22.2004
            .IsDeleted = "False"
            .DownLoadMe = "False"
            .UpLoadMe = "True"
            .AdminComments = vbNullString
            .DateLastUpdated = Format(Now(), "MM/DD/YYYY HH:MM:SS")
            .UpdateByUserID = goUtil.gsCurUsersID
        End With
        
        'Make the De Res Photo
        If moDeRes.DeResPhoto(udtDeResPhoto) Then
            If mfrmPhotos.AddPhotoItem(udtPhoto) Then
                'If sucessful save of photo and the
                'Delete Original Photo After Attach is checked...
                If chkDelOrigPhoto.Value = vbChecked Then
                    If chk35mm.Value = vbUnchecked Then
                        goUtil.utDeleteFile udtDeResPhoto.sPhotoPath
                    End If
                End If
                mbFirstSave = False
                mlMaxSort = mfrmPhotos.GetMaxSort(msAssignmentsID, msIDRTPhotoReport) + 1
                mlPhotoCount = mfrmPhotos.GetPhotoCount(msAssignmentsID, msIDRTPhotoReport) + 1
                If FilePhoto.Enabled And FilePhoto.Visible Then
                    FilePhoto.SetFocus
                End If
                If FilePhoto.ListIndex + 1 < FilePhoto.ListCount Then
                    FilePhoto.ListIndex = FilePhoto.ListIndex + 1
                End If
                On Error GoTo EH
            End If
        End If
        
    Else 'Edit
        With udtPhoto
            'If editing need to Update the Selected ListItem as well
            Set moListView = mfrmPhotos.lstvPhotos
            Set mitmX = moListView.SelectedItem
            .RTPhotoLogID = mitmX.SubItems(GuiPhotoListView.RTPhotoLogID - 1)
            .RTPhotoReportID = mitmX.SubItems(GuiPhotoListView.RTPhotoReportID - 1)
            .AssignmentsID = mitmX.SubItems(GuiPhotoListView.AssignmentsID - 1)
            .BillingCountID = mitmX.SubItems(GuiPhotoListView.BillingCountID - 1)
            .ID = mitmX.SubItems(GuiPhotoListView.ID - 1)
            .IDRTPhotoReport = mitmX.SubItems(GuiPhotoListView.IDRTPhotoReport - 1)
            .IDAssignments = mitmX.SubItems(GuiPhotoListView.IDAssignments - 1)
            .IDBillingCount = mitmX.SubItems(GuiPhotoListView.IDBillingCount - 1)
            .PhotoDate = txtPhotoDate.Text
            If Val(txtSortOrder.Text) > 0 Then
                .SortOrder = txtSortOrder.Text
            Else
                .SortOrder = mlMaxSort
            End If
            .Description = txtDescription.Text
            
            If sPhotoName = vbNullString Then
                sPhotoName = mitmX.SubItems(GuiPhotoListView.PhotoName - 1)
            Else
                sPhotoName = sPhotoName & "1.jpg"
            End If
        
            .PhotoName = sPhotoName
            'Set Flag for Upload (Highres = 0, Normal = 1, Thumb = 2)
            If StrComp(msEditDeResPath, lblFileName.Caption, vbTextCompare) = 0 Then
                bUploadPhoto = False
            Else
                bUploadPhoto = True
            End If
            .Photo = mitmX.SubItems(GuiPhotoListView.Photo - 1)
            .DownloadPhoto = goUtil.GetFlagFromText(mitmX.SubItems(GuiPhotoListView.DownloadPhoto - 1))
            If goUtil.GetFlagFromText(mitmX.SubItems(GuiPhotoListView.UpLoadPhoto - 1)) Then
                .UpLoadPhoto = True
            Else
                .UpLoadPhoto = bUploadPhoto
            End If
            .PhotoThumb = mitmX.SubItems(GuiPhotoListView.PhotoThumb - 1)
            .DownloadPhotoThumb = goUtil.GetFlagFromText(mitmX.SubItems(GuiPhotoListView.DownloadPhotoThumb - 1))
            If goUtil.GetFlagFromText(mitmX.SubItems(GuiPhotoListView.UpLoadPhotoThumb - 1)) Then
                .UpLoadPhotoThumb = True
            Else
                .UpLoadPhotoThumb = bUploadPhoto
            End If
            ' This option is not yet supported 7.22.2004
            .PhotoHighRes = mitmX.SubItems(GuiPhotoListView.PhotoHighRes - 1)
            .DownloadPhotoHighRes = goUtil.GetFlagFromText(mitmX.SubItems(GuiPhotoListView.DownloadPhotoHighRes - 1))
            If goUtil.GetFlagFromText(mitmX.SubItems(GuiPhotoListView.UploadPhotoHighRes - 1)) Then
                .UploadPhotoHighRes = True
            Else
                .UploadPhotoHighRes = bUploadPhoto
            End If
            .IsDeleted = goUtil.GetFlagFromText(mitmX.SubItems(GuiPhotoListView.IsDeleted - 1))
            .DownLoadMe = goUtil.GetFlagFromText(mitmX.SubItems(GuiPhotoListView.DownLoadMe - 1))
            .UpLoadMe = True
            .AdminComments = mitmX.SubItems(GuiPhotoListView.AdminComments - 1)
            .DateLastUpdated = Format(Now(), "MM/DD/YYYY HH:MM:SS")
            .UpdateByUserID = goUtil.gsCurUsersID
            
        End With
        'Only Deres and Edit to existing photo if a new photo is selected
        If StrComp(msEditDeResPath, lblFileName.Caption, vbTextCompare) <> 0 Then
            If moDeRes.DeResPhoto(udtDeResPhoto) Then
                If mfrmPhotos.EditPhotoItem(udtPhoto) Then
                    'If sucessful save of photo and the
                    'Delete Original Photo After Attach is checked...
                    If chkDelOrigPhoto.Value = vbChecked Then
                        If chk35mm.Value = vbUnchecked Then
                            goUtil.utDeleteFile udtDeResPhoto.sPhotoPath
                        End If
                    End If
                    
                    'We need to get rid of the old Pic
                    If goUtil.utFileExists(msEditDeResPath) Then
                        SetAttr msEditDeResPath, vbNormal
                        goUtil.utDeleteFile msEditDeResPath
                    End If
                    'Remove Highres
                    msEditDeResPath = Replace(msEditDeResPath, "_1.jpg", "_0.jpg")
                    If goUtil.utFileExists(msEditDeResPath) Then
                        SetAttr msEditDeResPath, vbNormal
                        goUtil.utDeleteFile msEditDeResPath
                    End If
                    'Remove Thumbnails
                    msEditDeResPath = Replace(msEditDeResPath, "_0.jpg", "_2.jpg")
                    If goUtil.utFileExists(msEditDeResPath) Then
                        SetAttr msEditDeResPath, vbNormal
                        goUtil.utDeleteFile msEditDeResPath
                    End If
                    'cleanup
                    Set udtDeResPhoto.colParameters = Nothing
                    mbSavingPhoto = False
                    mbSavedPhoto = True
                    goUtil.utFormWinRegPos goUtil.gsMainAppEXEName, Me, True, , , , True
                    Me.Visible = False
                End If
            End If
        Else
            'If the User Did not select a new source photo Then do not Deres the
            'Already IMported Photo.  This will cause the photo to degrade
            'Just Update the Photo Recorset
            mfrmPhotos.EditPhotoItem udtPhoto
            'cleanup
            Set udtDeResPhoto.colParameters = Nothing
            mbSavingPhoto = False
            mbSavedPhoto = True
            goUtil.utFormWinRegPos goUtil.gsMainAppEXEName, Me, True, , , , True
            Me.Visible = False
        End If
    End If
    
    mbSavingPhoto = False
    'cleanup
    Set udtDeResPhoto.colParameters = Nothing
    mbSavedPhoto = True
    Exit Function
EH:
    mbSavingPhoto = False
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Function SavePhoto"
End Function

Public Sub SetOriginalPhoto()
    On Error GoTo EH
    Dim sUploadPhotoPath  As String
    Dim sUploadPhotoSizeKB As String
    Dim sOrigPhotoSizeKB As String
    
    Image2 = Image1.Picture
    lblHandWOptimal.Caption = "(Width = " & ConvertTwipsToPixels(Optimal_W) & " Height = " & ConvertTwipsToPixels(Optimal_H) & ")"
    mlOriginal_H = Image2.Height
    mlOriginal_W = Image2.Width
    GetFileSettings lblFileName.Caption, mFI, mFA
    lblCreatedDateTime.Caption = mFI.dtCreationDate
    lblModifiedDateTime.Caption = mFI.dtLastModifyTime
    lblAccessedDateTime.Caption = mFI.dtLastAccessTime
    sOrigPhotoSizeKB = Round(mFI.nFileSize / 1024, 0)
    lblHandWOriginal.Caption = "(Width = " & ConvertTwipsToPixels(mlOriginal_W) & " Height = " & ConvertTwipsToPixels(mlOriginal_H) & ") " & sOrigPhotoSizeKB & " KB"
    
    'IF editing an existing photo look for the upload photo
    If msEditDeResPath <> vbNullString Then
        sUploadPhotoPath = msEditDeResPath
        sUploadPhotoPath = Replace(sUploadPhotoPath, "_0.jpg", "_1.jpg", , , vbTextCompare)
        If goUtil.utFileExists(sUploadPhotoPath) Then
            On Error Resume Next
            Image3.Picture = LoadPicture(sUploadPhotoPath)
            If Err.Number > 0 Then
                Err.Clear
                Image3.Picture = mfrmPhotos.imgPhotoStatus.ListImages(GuiPhotoStatusList.IsDeleted).Picture
            End If
            On Error GoTo EH
            GetFileSettings sUploadPhotoPath, mFI, mFA
            lblUploadCreatedDateTime.Caption = mFI.dtCreationDate
            lblUploadModifiedDateTime.Caption = mFI.dtLastModifyTime
            lblUploadAccessedDateTime.Caption = mFI.dtLastAccessTime
            sUploadPhotoSizeKB = Round(mFI.nFileSize / 1024, 0)
            lblHandWCurrentUpload.Caption = "(Width = " & ConvertTwipsToPixels(Image3.Width) & " Height = " & ConvertTwipsToPixels(Image3.Height) & ") " & sUploadPhotoSizeKB & " KB"
        End If
    End If
    
    'Look for the UploadPhoto
    'Set the Resize as close as possible to the Optimal Size
    ResizeToOptimal
    DisplayResize
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub SetOriginalPhoto"
End Sub

Private Sub GetFileSettings(sFilePath, pFI As V2ECKeyBoard.FILE_INFORMATION, pFA As V2ECKeyBoard.FILE_ATTRIBUTES)
    On Error GoTo EH
    
    Set moFI = New V2ECKeyBoard.clsFileVersion
    pFI = moFI.GetFileInformation(sFilePath)
    pFA = mFI.faFileAttributes
    Set moFI = Nothing
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub GetFileSettings"
End Sub

Public Function CLEANUP() As Boolean
    On Error GoTo EH
    Set mfrmClaim = Nothing
    Set mfrmPhotos = Nothing
    Set moGUI = Nothing
    Set moDeRes = Nothing
    Set mitmX = Nothing
    Set moListView = Nothing
    Set moFI = Nothing
    
    CLEANUP = True
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function CLEANUP"
End Function

Private Sub txtSortOrder_LostFocus()
    goUtil.utValidate , txtSortOrder
End Sub


