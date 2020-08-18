VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLossReports 
   AutoRedraw      =   -1  'True
   Caption         =   "Loss Reports"
   ClientHeight    =   6615
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
   Icon            =   "frmLossReports.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6615
   ScaleWidth      =   9480
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame framReports 
      Appearance      =   0  'Flat
      Caption         =   "Repor&ts"
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
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9250
      Begin VB.Timer Timer_Resize 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   8640
         Top             =   840
      End
      Begin MSComctlLib.ImageList imgLossReports 
         Left            =   8520
         Top             =   240
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483633
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   16777215
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   36
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLossReports.frx":0442
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLossReports.frx":0894
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLossReports.frx":0CE6
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLossReports.frx":1138
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLossReports.frx":158A
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLossReports.frx":19DC
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLossReports.frx":1E2E
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLossReports.frx":2280
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLossReports.frx":26D2
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLossReports.frx":2B24
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLossReports.frx":2F76
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLossReports.frx":33C8
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLossReports.frx":381A
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLossReports.frx":3C6C
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLossReports.frx":3F86
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLossReports.frx":43D8
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLossReports.frx":482A
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLossReports.frx":4C7C
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLossReports.frx":50CE
               Key             =   ""
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLossReports.frx":5520
               Key             =   ""
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLossReports.frx":5972
               Key             =   ""
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLossReports.frx":5DC4
               Key             =   ""
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLossReports.frx":6216
               Key             =   ""
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLossReports.frx":6668
               Key             =   ""
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLossReports.frx":6ABA
               Key             =   ""
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLossReports.frx":6F0C
               Key             =   ""
            EndProperty
            BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLossReports.frx":735E
               Key             =   ""
            EndProperty
            BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLossReports.frx":77B0
               Key             =   ""
            EndProperty
            BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLossReports.frx":7C02
               Key             =   ""
            EndProperty
            BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLossReports.frx":8054
               Key             =   ""
            EndProperty
            BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLossReports.frx":84A6
               Key             =   ""
            EndProperty
            BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLossReports.frx":88F8
               Key             =   ""
            EndProperty
            BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLossReports.frx":8D4A
               Key             =   ""
            EndProperty
            BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLossReports.frx":919C
               Key             =   ""
            EndProperty
            BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLossReports.frx":95EE
               Key             =   ""
            EndProperty
            BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLossReports.frx":9A40
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lvwLossReports 
         Height          =   2175
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   3836
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "imgLossReports"
         SmallIcons      =   "imgLossReports"
         ColHdrIcons     =   "imgLossReports"
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
   Begin VB.Frame framPrnOptions 
      Appearance      =   0  'Flat
      Caption         =   "Printer"
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
      Height          =   3975
      Left            =   4080
      TabIndex        =   32
      Top             =   2520
      Width           =   5280
      Begin VB.CheckBox chkViewMess 
         Caption         =   "Vie&w Messages"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   39
         ToolTipText     =   "View misc. messages."
         Top             =   2160
         Width           =   1815
      End
      Begin VB.TextBox txtDaysAgo 
         Enabled         =   0   'False
         Height          =   360
         Left            =   3540
         MaxLength       =   3
         TabIndex        =   42
         Text            =   "-1"
         Top             =   3030
         Width           =   375
      End
      Begin VB.TextBox txtPDFPath 
         Enabled         =   0   'False
         Height          =   360
         Left            =   120
         TabIndex        =   43
         Top             =   3480
         Width           =   3420
      End
      Begin VB.CommandButton cmdPDFPath 
         Enabled         =   0   'False
         Height          =   330
         Left            =   3540
         Picture         =   "frmLossReports.frx":9B9A
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Browse"
         Top             =   3495
         Width           =   375
      End
      Begin VB.CheckBox chkPrintToPDF 
         Caption         =   "Print to PDF F&ile Path"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   2880
         Width           =   3885
      End
      Begin VB.CheckBox ChkAppAdjDateStamp 
         Caption         =   "Append date to File Na&me"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   3120
         Value           =   1  'Checked
         Width           =   3405
      End
      Begin VB.OptionButton optPrnFormat 
         Caption         =   "F&ormatted"
         Height          =   855
         Index           =   1
         Left            =   1470
         Picture         =   "frmLossReports.frx":A014
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   480
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton optPrnFormat 
         Caption         =   "&Raw Data"
         Height          =   855
         Index           =   0
         Left            =   120
         Picture         =   "frmLossReports.frx":A456
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton cmdPrintList 
         Caption         =   "Sho&w Item List"
         Height          =   855
         Left            =   2820
         MaskColor       =   &H00000000&
         Picture         =   "frmLossReports.frx":A898
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Exit"
         Top             =   480
         Width           =   1095
      End
      Begin VB.CheckBox chkPreviewScreen 
         Caption         =   "Previe&w to Screen"
         Height          =   255
         Left            =   2040
         TabIndex        =   38
         Top             =   1920
         Width           =   2000
      End
      Begin VB.CheckBox chkChain 
         Caption         =   "Chai&n Reports"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         ToolTipText     =   "Keep Together"
         Top             =   1920
         Width           =   1815
      End
      Begin VB.ComboBox cboSelectPrinter 
         Height          =   360
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   1455
         Width           =   3795
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Height          =   375
         Left            =   4040
         TabIndex        =   49
         Top             =   3000
         Width           =   1100
      End
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "&Exit"
         Height          =   375
         Left            =   4040
         TabIndex        =   50
         Top             =   3480
         Width           =   1100
      End
      Begin VB.CheckBox chkViewGrid 
         Caption         =   "&Grid OFF"
         Height          =   375
         Left            =   4040
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   1920
         Width           =   1100
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Find"
         Height          =   375
         Left            =   4040
         TabIndex        =   47
         Top             =   960
         Width           =   1100
      End
      Begin VB.CommandButton cmdSelectAll 
         Caption         =   "&Select All"
         Height          =   375
         Left            =   4040
         TabIndex        =   46
         Top             =   480
         Width           =   1100
      End
      Begin VB.TextBox txtDummy 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3525
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   3480
         Width           =   390
      End
   End
   Begin VB.Frame framIncludeDocs 
      Appearance      =   0  'Flat
      Caption         =   "Append Documents"
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
      Height          =   3975
      Left            =   120
      TabIndex        =   2
      Tag             =   "AppDoc"
      Top             =   2520
      Width           =   3840
      Begin VB.CheckBox chkDoc 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   13
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   31
         Tag             =   "AppDoc"
         Top             =   3600
         Width           =   3360
      End
      Begin VB.CheckBox chkDoc 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   12
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   29
         Tag             =   "AppDoc"
         Top             =   3360
         Width           =   3360
      End
      Begin VB.CheckBox chkDoc 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   11
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   27
         Tag             =   "AppDoc"
         Top             =   3120
         Width           =   3360
      End
      Begin VB.CheckBox chkDoc 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   10
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   25
         Tag             =   "AppDoc"
         Top             =   2880
         Width           =   3360
      End
      Begin VB.CheckBox chkDoc 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   9
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   23
         Tag             =   "AppDoc"
         Top             =   2640
         Width           =   3360
      End
      Begin VB.CheckBox chkDoc 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   8
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   21
         Tag             =   "AppDoc"
         Top             =   2400
         Width           =   3360
      End
      Begin VB.CheckBox chkDoc 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   7
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   19
         Tag             =   "AppDoc"
         Top             =   2160
         Width           =   3360
      End
      Begin VB.CheckBox chkDoc 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   6
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   17
         Tag             =   "AppDoc"
         Top             =   1920
         Width           =   3360
      End
      Begin VB.CheckBox chkDoc 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   5
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   15
         Tag             =   "AppDoc"
         Top             =   1680
         Width           =   3360
      End
      Begin VB.CheckBox chkDoc 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   13
         Tag             =   "AppDoc"
         Top             =   1440
         Width           =   3360
      End
      Begin VB.CheckBox chkDoc 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   11
         Tag             =   "AppDoc"
         Top             =   1200
         Width           =   3360
      End
      Begin VB.CheckBox chkDoc 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   9
         Tag             =   "AppDoc"
         Top             =   960
         Width           =   3360
      End
      Begin VB.CheckBox chkDoc 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   7
         Tag             =   "AppDoc"
         Top             =   720
         Width           =   3360
      End
      Begin VB.CheckBox chkDoc 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   5
         Tag             =   "AppDoc"
         Top             =   480
         Width           =   3360
      End
      Begin VB.CheckBox chkAllDoc 
         Caption         =   "&Append All Documents"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   3615
      End
      Begin VB.Label lblAppNum 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "14"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   13
         Left            =   105
         TabIndex        =   30
         Tag             =   "AppDoc"
         Top             =   3600
         Width           =   225
      End
      Begin VB.Label lblAppNum 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "13"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   12
         Left            =   105
         TabIndex        =   28
         Tag             =   "AppDoc"
         Top             =   3360
         Width           =   225
      End
      Begin VB.Label lblAppNum 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   11
         Left            =   105
         TabIndex        =   26
         Tag             =   "AppDoc"
         Top             =   3120
         Width           =   225
      End
      Begin VB.Label lblAppNum 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "11"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   10
         Left            =   105
         TabIndex        =   24
         Tag             =   "AppDoc"
         Top             =   2880
         Width           =   225
      End
      Begin VB.Label lblAppNum 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   9
         Left            =   105
         TabIndex        =   22
         Tag             =   "AppDoc"
         Top             =   2640
         Width           =   225
      End
      Begin VB.Label lblAppNum 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   8
         Left            =   105
         TabIndex        =   20
         Tag             =   "AppDoc"
         Top             =   2400
         Width           =   225
      End
      Begin VB.Label lblAppNum 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   7
         Left            =   105
         TabIndex        =   18
         Tag             =   "AppDoc"
         Top             =   2160
         Width           =   225
      End
      Begin VB.Label lblAppNum 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   6
         Left            =   105
         TabIndex        =   16
         Tag             =   "AppDoc"
         Top             =   1920
         Width           =   225
      End
      Begin VB.Label lblAppNum 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   5
         Left            =   105
         TabIndex        =   14
         Tag             =   "AppDoc"
         Top             =   1680
         Width           =   225
      End
      Begin VB.Label lblAppNum 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   4
         Left            =   105
         TabIndex        =   12
         Tag             =   "AppDoc"
         Top             =   1440
         Width           =   225
      End
      Begin VB.Label lblAppNum 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   105
         TabIndex        =   10
         Tag             =   "AppDoc"
         Top             =   1215
         Width           =   225
      End
      Begin VB.Label lblAppNum 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   105
         TabIndex        =   8
         Tag             =   "AppDoc"
         Top             =   975
         Width           =   225
      End
      Begin VB.Label lblAppNum 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   105
         TabIndex        =   6
         Tag             =   "AppDoc"
         Top             =   735
         Width           =   225
      End
      Begin VB.Label lblAppNum 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   105
         TabIndex        =   4
         Tag             =   "AppDoc"
         Top             =   495
         Width           =   225
      End
   End
End
Attribute VB_Name = "frmLossReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Tag for Append Documents controls
Private Const APP_DOC As String = "AppDoc"
Private Const CHK_DOC As String = "chkDoc"

'Control Size in Ref to Form Diff Constants
Private Const FORM_W As Long = 9600
Private Const FORM_H As Long = 7020
Private Const framReports_H As Long = 4485
Private Const framReports_W As Long = 350
Private Const lvwLossReports_H As Long = 4845
Private Const lvwLossReports_W As Long = 585

'framIncludeDocs
Private Const framIncludeDocs_T As Long = 4500
Private Const framIncludeDocs_W As Long = 5760
Private Const chkDoc_W  As Long = 6240
'framPrnOptions
Private Const framPrnOptions_T As Long = 4500
Private Const framPrnOptions_W As Long = 4320
Private Const framPrnOptions_L As Long = 5520




Private mbResize As Boolean    'Flag the Resize event to be accomplished in Timer
Private msPDFOutPath As String 'Will Use Default path in Regsetting if this is nullstring
Private mbLoading As Boolean 'True if still in the form load code
Private mbUnLoadMe As Boolean 'True if ready to unload by code
Private mcolLossReports As Collection 'Will contain a collection of clsLoss@@@@? objects depending on prn format they are
Private moLRs As V2ECKeyBoard.clsLossReports
Private mbPrinting As Boolean 'True if currently printing
Private moProgForm As V2ECKeyBoard.clsProgForm 'V2ECKeyBoard.frmProgress
Private mlColumn As Long
Private mbColSearch As Boolean 'False until a Column Header is clicked for first time.

Private Property Get msClassName() As String
    msClassName = Me.Name
End Property

Public Property Let ProgForm(poProgForm As V2ECKeyBoard.clsProgForm)
    Set moProgForm = poProgForm
End Property
Public Property Set ProgForm(poProgForm As V2ECKeyBoard.clsProgForm)
    Set moProgForm = poProgForm
End Property

Public Property Let LossReportsCol(pcolLossReports As Collection)
    Set mcolLossReports = pcolLossReports
End Property
Public Property Set LossReportsCol(pcolLossReports As Collection)
    Set mcolLossReports = pcolLossReports
End Property
Public Property Get LossReportsCol() As Collection
    Set LossReportsCol = mcolLossReports
End Property

Public Property Let LRs(poLRs As clsLossReports)
    Set moLRs = poLRs
End Property
Public Property Set LRs(poLRs As clsLossReports)
    Set moLRs = poLRs
End Property
Public Property Get LRs() As clsLossReports
    Set LRs = moLRs
End Property

Public Property Let UnLoadMe(pbFlag As Boolean)
    mbUnLoadMe = pbFlag
End Property
Public Property Get UnLoadMe() As Boolean
    UnLoadMe = mbUnLoadMe
End Property

Public Property Let PDFOutPath(psPath As String)
    msPDFOutPath = psPath
End Property
Public Property Get PDFOutPath() As String
    PDFOutPath = msPDFOutPath
End Property

Private Sub cboSelectPrinter_Click()
    On Error GoTo EH
    
    If Not mbLoading Then
        SaveSetting App.EXEName, "Printer", "LossReport", cboSelectPrinter.Text
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cboSelectPrinter_Click"
End Sub

Private Sub chkAllDoc_Click()
    On Error GoTo EH
    Dim oChk As Object
    
    'Check Printing flag first
    If mbPrinting Then
        Exit Sub 'Bail
    End If
    
    'Check or un check all the App doc check boxes
    'if they are visible and enabled.
    For Each oChk In Me.Controls
        If TypeOf oChk Is CheckBox Then
            If oChk.Tag = APP_DOC And oChk.Name = CHK_DOC Then
                If oChk.Visible And oChk.Enabled Then
                    If chkAllDoc.Value = vbChecked Then
                        oChk.Value = vbChecked
                    Else
                        oChk.Value = vbUnchecked
                    End If
                End If
                
            End If
        End If
    Next
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub chkAllDoc_Click"
End Sub

Private Sub ChkAppAdjDateStamp_Click()
    On Error GoTo EH
    
    If ChkAppAdjDateStamp.Value = vbChecked Then
        txtDaysAgo.Visible = True
    Else
        txtDaysAgo.Visible = False
    End If
    AppendDate
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub ChkAppAdjDateStamp_Click"
End Sub

Private Sub AppendDate()
    On Error GoTo EH
    Dim dDays As Double
    Dim sTemp As String

    sTemp = txtPDFPath.Text
    sTemp = left(sTemp, InStrRev(sTemp, "\", , vbBinaryCompare))
    
    If ChkAppAdjDateStamp.Value = vbChecked Then
        dDays = CDbl(txtDaysAgo.Text)
        sTemp = sTemp & Format(DateAdd("d", dDays, Now()), "MMDDYY") & "_"
    End If
    
    txtPDFPath.Text = sTemp
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub AppendDate"
End Sub

Private Sub ChkAppAdjDateStamp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetDaysAgo
End Sub

Private Sub chkChain_Click()
    On Error GoTo EH
    'Modify Heights acoording to Appending Documents Flag
    If moLRs.AppDocFlag Then
        If chkChain.Value = vbUnchecked Then
            moLRs.LoadAppDoc
            EnableAppDoc
            chkAllDoc.Enabled = True
            framIncludeDocs.Enabled = True
        Else
            GoTo HIDETHEM
        End If
    Else
HIDETHEM:
        EnableAppDoc
        chkAllDoc.Enabled = False
        chkAllDoc.Value = vbUnchecked
        framIncludeDocs.Enabled = False
    End If
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub chkChain_Click"
End Sub

Private Sub chkDoc_Click(Index As Integer)
    On Error GoTo EH
    
    If chkDoc(Index).Value = vbChecked Then
        chkDoc(Index).FontBold = True
    Else
        chkDoc(Index).FontBold = False
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub chkDoc_Click"
End Sub

Private Sub chkPrintToPDF_Click()
    On Error GoTo EH
    
    If chkPrintToPDF.Value = vbChecked Then
        If Not goUtil.utFileExists(txtPDFPath.Text, True) Then
            cmdPrint.Enabled = False
        Else
            cmdPrint.Enabled = True
        End If
        txtPDFPath.Enabled = True
        optPrnFormat(PrintFormat.Translated).Value = True
        optPrnFormat(PrintFormat.RawText).Enabled = False
        optPrnFormat(PrintFormat.Translated).Enabled = False
        chkPreviewScreen.Value = vbUnchecked
        ChkAppAdjDateStamp.Enabled = True
        txtDaysAgo.Enabled = True
        chkPreviewScreen.Enabled = False
        cmdPDFPath.Enabled = True
        cmdPDFPath.SetFocus
    Else
        cmdPrint.Enabled = True
        cmdPDFPath.Enabled = False
        txtPDFPath.Enabled = False
        optPrnFormat(PrintFormat.RawText).Enabled = True
        optPrnFormat(PrintFormat.Translated).Enabled = True
        chkPreviewScreen.Enabled = True
        ChkAppAdjDateStamp.Enabled = False
        txtDaysAgo.Enabled = False
    End If
    
    AppendDate
    EnablePrint
    
    If ChkAppAdjDateStamp.Enabled Then
        ChkAppAdjDateStamp.SetFocus
    End If

    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub chkPrintToPDF_Click"
End Sub

Private Sub chkViewGrid_Click()
    On Error GoTo EH
    Dim bGridOn As Boolean
    
    If chkViewGrid.Value = vbChecked Then
        chkViewGrid.Caption = "&Grid ON"
        bGridOn = True
    Else
        chkViewGrid.Caption = "&Grid OFF"
        bGridOn = False
    End If
    
    SaveSetting App.EXEName, "GENERAL", "GRID_ON", bGridOn
    lvwLossReports.Gridlines = bGridOn
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub chkViewGrid_Click"
End Sub

Private Sub chkViewMess_Click()
    On Error GoTo EH
    
    Populate
    If chkViewMess.Value = vbChecked Then
        cmdPrint.Enabled = False
    Else
        cmdPrint.Enabled = True
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub chkViewMess_Click"
End Sub

'Private Sub cmdDelete_Click()
'    DeleteItems
'End Sub

Private Sub cmdExit_Click()
    On Error GoTo EH
    'Check Printing flag first
    If mbPrinting Then
        Exit Sub 'Bail
    End If
    Me.Visible = False
    UnLoadMe = True
    If Not goUtil.gfrmECTray Is Nothing Then
        Unload Me
    End If
    If Not moLRs Is Nothing Then
        moLRs.FireCleanUpLossReports
    End If
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdExit_Click"
End Sub

Private Sub cmdFind_Click()
    FindText
End Sub

Private Sub cmdPDFPath_Click()
    On Error GoTo EH
    txtPDFPath.Text = goUtil.utGetPath(App.EXEName, "PDFOutPath", "BROWSE TO PDF OUTPUT PATH", "CLICK OPEN TO SAVE PATH", txtPDFPath.Text, Me.hWnd)
    AppendDate
    EnablePrint
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdPDFPath_Click"
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo EH
    Dim itmX As listItem
    Dim oLR As V2ECKeyBoard.clsCarLR ' used to ref a Loss Report object in the Collection
    Dim sXlateReportName As String 'Translation Report Name
    Dim PFormat As PrintFormat
    Dim bPreView As Boolean
    Dim bPDF As Boolean
    Dim bOpenPrn As Boolean
    Dim sPrnDeviceName As String ' will be populated by moLRs.OpenPrn
    'Used for Chaining
    Dim MyChainType As ChainType
    Dim lSelected As Long
    Dim bChaining As Boolean
    Dim sPDFPath As String
    Dim dDays As Double
    Dim sFile As String
    Dim sFilePath As String
    Dim bProgBar As Boolean
    Dim dicMaxAllowed As scripting.Dictionary
    Dim lMaxAllowedCount As Long
    Dim sTemp As String
    Dim sMess As String
    
    'Check Printing flag first
    If mbPrinting Then
        Exit Sub 'Bail
    End If
    Me.MousePointer = vbHourglass
    mbPrinting = True
    
    If Not IsNumeric(txtDaysAgo.Text) Then
        txtDaysAgo.Text = 0
    End If
    
    'Set the print format
    If optPrnFormat(PrintFormat.RawText).Value Then
        PFormat = RawText
    ElseIf optPrnFormat(PrintFormat.Translated).Value Then
        PFormat = Translated
    End If
    
    'Set the preview flag
    If chkPreviewScreen.Value = vbChecked Then
        bPreView = True
    End If
    
    'Set the print to PDF flag
    If chkPrintToPDF.Value = vbChecked Then
        bPDF = True
    End If
    
    'Get Count of selected to figure ChainType
    For Each itmX In lvwLossReports.ListItems
        If itmX.Selected Then
            If chkChain.Value = vbChecked Or chkPreviewScreen.Value = vbChecked Then
                Set oLR = mcolLossReports(itmX.SubItems(LossReports.RKey - 1))
                sTemp = goUtil.GetItemFromDictionary(dicMaxAllowed, oLR.ClassName)
                If sTemp = vbNullString Then
                    lMaxAllowedCount = oLR.MaxAllowedInChain - 1
                Else
                    lMaxAllowedCount = CLng(sTemp) - 1
                End If
                goUtil.RemoveItemFromDictionary dicMaxAllowed, oLR.ClassName
                goUtil.AddItemToDictionary dicMaxAllowed, CStr(lMaxAllowedCount), oLR.ClassName
                If lMaxAllowedCount < 0 Then
                    sMess = "A maximum of " & oLR.MaxAllowedInChain & " " & oLR.ClassName & vbCrLf & "reports are allowed to be chained or Print Previewed."
                End If
            End If
            lSelected = lSelected + 1
            Set oLR = Nothing
        End If
    Next
    'Clean up the dictionary
    Set dicMaxAllowed = Nothing
    
    'Set the chain type
    If lSelected > 1 And chkChain.Value = vbChecked Then
        MyChainType = FirstInChain
        bChaining = True
    Else
        MyChainType = NotChain
    End If
    If sMess <> vbNullString Then
        MsgBox "You selected " & lSelected & vbCrLf & vbCrLf & sMess, vbExclamation, "Exceeded Maximum Allowed In Chain or Print Preview"
        GoTo CLEAN_UP
    Else
        If MsgBox("You selected " & lSelected & " Loss Reports." & vbCrLf & vbCrLf & _
               "Allow for approximately " & lSelected * 2 & " seconds to process." & vbCrLf & vbCrLf & _
               "Thank You!", vbInformation + vbOKCancel) = vbCancel Then
            GoTo CLEAN_UP
        End If
    End If
    
    'ProgBar
    If Not moProgForm Is Nothing And chkViewMess.Value = vbUnchecked And lSelected > 1 Then
        If chkPreviewScreen.Value = vbChecked And chkChain.Value = vbUnchecked Then
            bProgBar = False
        Else
            If Not moProgForm Is Nothing Then
                If Not moProgForm.Object Is Nothing Then
                    moProgForm.ShowForm True
                    moProgForm.SetFocus
                    bProgBar = True
                    moProgForm.PBarFile.Max = lSelected
                    moProgForm.PBarFile.Value = 0
                End If
            End If
        End If
    End If
    
    For Each itmX In lvwLossReports.ListItems
        If itmX.Selected Then
            If Not bOpenPrn And optPrnFormat(PrintFormat.RawText).Value Then
                bOpenPrn = moLRs.OpenPrn(sPrnDeviceName) 'open printer
            End If
            Set oLR = mcolLossReports.Item(itmX.SubItems(LossReports.RKey - 1))
            
            If Not bOpenPrn And sPrnDeviceName = vbNullString Then
                'If we are using Formated print then we still need to get
                'the Print device name to be passed to Active Report Printer Object
                moLRs.OpenPrn sPrnDeviceName, True
            End If
            If bProgBar Then
                If Not moProgForm Is Nothing Then
                    If Not moProgForm.Object Is Nothing Then
                        moProgForm.lblFileText = oLR.PrnKey
                        moProgForm.RefreshMe
                    End If
                End If
            End If
            oLR.PrintMe sPrnDeviceName, PFormat, bPreView, bPDF, Me, Me.hWnd, MyChainType
            'Set ChainType
            If bChaining Then
                MyChainType = NextLink
            End If
            If bProgBar Then
                moProgForm.PBarFile.Value = moProgForm.PBarFile.Value + 1
                'See if the user Cancled the process
                DoEvents
                Sleep 10
                If moProgForm.CancelMe Then
                    moProgForm.CancelMe = False
                    GoTo CLEAN_UP
                End If
            End If
            Set oLR = Nothing ' clean local memory
        End If
        
    Next
        
    'If we have chained reports we need to either print preview them in one single
    'report. Or Send them to PDF as one single report
    If bChaining Then
        'If we are previewing then we need to use ARV object
        If bPreView Then
            If Not moLRs.ChainReport Is Nothing Then
                If goUtil.gARV Is Nothing Then
                    Set goUtil.gARV = CreateObject("V2ARViewer.clsARViewer")
                End If
                If bProgBar Then
                    moProgForm.lblFileText = "Loading preview.  Please wait!"
                    moProgForm.RefreshMe
                    With goUtil.gARV
                        .SetUtilObject goUtil
                        moLRs.ChainReport.Run True 'TRUE = Run Asynch, False or blank =Run before it gets to ARV object
                        .objARvReport = moLRs.ChainReport
                        .sRptTitle = moLRs.ChainReportName
                        .ShowReport vbModeless
                    End With
                    moProgForm.ShowForm False
                Else
                    With goUtil.gARV
                        .SetUtilObject goUtil
                        moLRs.ChainReport.Run True 'TRUE = Run Asynch, False or blank =Run before it gets to ARV object
                        .objARvReport = moLRs.ChainReport
                        .sRptTitle = moLRs.ChainReportName
                        .ShowReport vbModeless
                    End With
                End If
            Else
                moProgForm.ShowForm False
            End If
        Else 'If we not preview then check to see if this is going to PDF
            If bPDF Then
                'Set PDF File
                sPDFPath = Trim(txtPDFPath.Text)
                sFile = moLRs.ChainAdjuster
                sPDFPath = sPDFPath & sFile
                If bProgBar Then
                    moProgForm.lblFileText = "Printing to PDF File:  " & sPDFPath & " Please Wait!"
                    moProgForm.RefreshMe
                    Me.Refresh
                    moLRs.ExportFile moLRs.ChainReport, sPDFPath, ExportType.ARPdf 'Export to PDF
                    moProgForm.ShowForm False
                Else
                    moLRs.ExportFile moLRs.ChainReport, sPDFPath, ExportType.ARPdf 'Export to PDF
                End If
            Else
                If bProgBar Then
                    moProgForm.lblFileText = "Printing.  Please Wait! "
                    moProgForm.RefreshMe
                    Me.Refresh
                    moLRs.ChainReport.PrintReport False 'Print without showing print dialog
                    moProgForm.ShowForm False
                Else
                    moLRs.ChainReport.PrintReport False 'Print without showing print dialog
                End If
            End If
        End If
    Else
        If bProgBar Then
            moProgForm.ShowForm False
        End If
    End If
    
CLEAN_UP:
    Set oLR = Nothing
    Set itmX = Nothing
    Set moLRs.ChainReport = Nothing
    If bOpenPrn Then
        moLRs.ClosePrn 'close printer
    End If
    mbPrinting = False
    Me.MousePointer = vbDefault
    If Me.Visible Then
        Me.SetFocus
    End If
    Exit Sub
EH:
    mbPrinting = False
    Me.MousePointer = vbDefault
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdPrint_Click"
End Sub

Private Sub cmdPrintList_Click()
    On Error GoTo EH

    goUtil.utPrintListView goUtil.gsAppEXEName, lvwLossReports, "Loss Report Items", ddOPortrait, vbModeless, 0, True
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdPrintList_Click"
End Sub

Private Sub cmdSelectAll_Click()
    On Error GoTo EH
    Dim itmX As listItem
    Dim lCount As Long
    'Check Printing flag first
    If mbPrinting Then
        Exit Sub 'Bail
    End If
    
    For Each itmX In lvwLossReports.ListItems
        itmX.Selected = True
        For lCount = 1 To itmX.ListSubItems.Count
            itmX.ListSubItems(lCount).ReportIcon = Empty
        Next
    Next
    
    lvwLossReports.SetFocus
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdSelectAll_Click"
End Sub

Private Sub Form_Load()
    On Error GoTo EH
    Dim prn As Printer
    Dim sLastSelectedPrinter As String
    
    'init flags
    mbLoading = True
    mbUnLoadMe = False
    
    Screen.MousePointer = vbHourglass
    
    'Get Form Posn
     goUtil.utFormWinRegPos goUtil.gsMainAppEXEName, Me
    
    'Modify Heights acoording to Appending Documents Flag
    If moLRs.AppDocFlag Then
        moLRs.LoadAppDoc
        EnableAppDoc
        chkAllDoc.Enabled = True
        framIncludeDocs.Enabled = True
    Else
        EnableAppDoc
        chkAllDoc.Enabled = False
        framIncludeDocs.Enabled = False
    End If
    
    'Fill the Printers Combo Box
    For Each prn In Printers
        cboSelectPrinter.AddItem prn.DeviceName & " on " & prn.Port
    Next prn
    
    sLastSelectedPrinter = GetSetting(App.EXEName, "Printer", "LossReport", vbNullString)
    'Select Deafult printer from list
    SelectDefaultPrinter cboSelectPrinter, sLastSelectedPrinter
   
    'Intit the PDF out path
    If Not goUtil.utFileExists(msPDFOutPath, True) Then
        txtPDFPath.Text = GetSetting(App.EXEName, "Dir", "PDFOutPath", vbNullString)
    Else
        txtPDFPath.Text = msPDFOutPath
    End If
    
    'Init the Loss Reports List View
    LoadHeader
    Populate
    SetDaysAgo
   
    
    Screen.MousePointer = vbDefault
    mbLoading = False
    
   Exit Sub
EH:
    Screen.MousePointer = vbDefault
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_Load"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo EH
    'Check Printing flag first
    If mbPrinting Then
        Cancel = True
        Exit Sub 'Bail
    End If
    If UnloadMode = vbFormControlMenu Then
        If goUtil.gfrmECTray Is Nothing Then
            Cancel = True
        End If
    End If
    
    goUtil.utFormWinRegPos goUtil.gsMainAppEXEName, Me, True
    Me.Visible = False
    UnLoadMe = True
    If Not moLRs Is Nothing Then
        moLRs.FireCleanUpLossReports
    End If
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_QueryUnload"
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
    framReports.Visible = pbVisible
    framIncludeDocs.Visible = pbVisible
    framPrnOptions.Visible = pbVisible
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub VisibleFrames"
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



Private Sub lvwLossReports_Click()
    On Error GoTo EH
    Dim itmX As listItem
    Dim itmXSel As listItem
    'This will find all Like Dates since it is the first Item
    Set itmX = lvwLossReports.SelectedItem
    
    'if the item selected is an HTML help globe then
    'unselect all other items and execute print event
    If Not itmX Is Nothing Then
        If itmX.SmallIcon = LRPic.lrHTMLHelp Then
            
            chkChain.Value = vbUnchecked
            For Each itmXSel In lvwLossReports.ListItems
                If itmXSel.Selected Then
                    itmXSel.Selected = False
                End If
            Next
            
            itmX.Selected = True
            cmdPrint_Click
        End If
    End If
    
    Set itmX = Nothing
    Set itmXSel = Nothing
    Exit Sub
EH:
    Me.MousePointer = vbDefault
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lvwLossReports_Click"
End Sub

Private Sub lvwLossReports_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    On Error GoTo EH
    
    Dim ColHead As ColumnHeader
    
    mlColumn = ColumnHeader.SubItemIndex
    
    For Each ColHead In lvwLossReports.ColumnHeaders
        If ColHead.Index = ColumnHeader.Index Then
            ColHead.Icon = LRPic.SearchColumn
        Else
            ColHead.Icon = Empty
        End If
        If Not mbColSearch And ColHead.Index < LossReports.RSort Then
            ColHead.Width = ColHead.Width + 100
        End If
    Next
    
    mbColSearch = True
    
    If lvwLossReports.SortOrder = lvwAscending Then
        lvwLossReports.SortOrder = lvwDescending
    Else
        lvwLossReports.SortOrder = lvwAscending
    End If
    
    Select Case ColumnHeader.Index
        Case LossReports.DateAsgn
            lvwLossReports.SortKey = ColumnHeader.Index
        Case Else
            lvwLossReports.SortKey = ColumnHeader.Index - 1
    End Select
     
    
    lvwLossReports.Sorted = True
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lvwLossReports_ColumnClick"
End Sub

Private Sub lvwLossReports_DblClick()
    On Error GoTo EH
    Dim itmX As listItem
    'This will find all Like Dates since it is the first Item
    Set itmX = lvwLossReports.SelectedItem
    If itmX Is Nothing Then
        Exit Sub
    End If
    
    Me.MousePointer = vbHourglass
    If mlColumn > 0 Then
        FindText itmX.ListSubItems(mlColumn).Text
    Else
        FindText itmX.Text
    End If
    Me.MousePointer = vbDefault
    Exit Sub
EH:
    Me.MousePointer = vbDefault
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lvwLossReports_DblClick"
End Sub

Private Sub lvwLossReports_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo EH
    Dim itmX As listItem
    Dim iPreView As Integer
    Dim iChain As Integer
    Select Case KeyCode
        Case vbKeyDelete
'            DeleteItems
        Case vbKeyF, vbKeySpace
            FindText
        Case vbKeyReturn
            Set itmX = lvwLossReports.SelectedItem
            If Not itmX Is Nothing Then
                If itmX.SmallIcon = LRPic.lrHTMLHelp Then
                    lvwLossReports_Click
                Else
                    iPreView = chkPreviewScreen.Value
                    iChain = chkChain.Value
                    chkPreviewScreen.Value = vbChecked
                    chkChain.Value = vbChecked
                    cmdPrint_Click
                    chkPreviewScreen.Value = iPreView
                    chkChain.Value = iChain
                End If
            End If
    End Select
    
    Set itmX = Nothing
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lvwLossReports_KeyDown"
End Sub

Private Sub optPrnFormat_Click(Index As Integer)
    'Check Printing flag first
    If mbPrinting Then
        Exit Sub 'Bail
    End If
    
    If optPrnFormat(PrintFormat.RawText).Value = True Then
        chkPreviewScreen.Value = vbUnchecked
        chkPreviewScreen.Enabled = False
        chkChain.Enabled = False
        chkChain.Value = vbUnchecked
    Else
        chkPreviewScreen.Enabled = True
        chkChain.Enabled = True
    End If
End Sub

Private Sub Timer_Resize_Timer()
    On Error GoTo EH
    Dim lH As Long
    Dim lW As Long
    Dim lCount As Long
    
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
    
    'framReports
    framReports.Height = lH - framReports_H
    framReports.Width = lW - framReports_W
    lvwLossReports.Height = lH - lvwLossReports_H
    lvwLossReports.Width = lW - lvwLossReports_W
    
    
    'Tops framIncludeDocs / framPrnOptions
    framPrnOptions.top = lH - framPrnOptions_T
    framIncludeDocs.top = lH - framIncludeDocs_T
    
    'width framIncludeDocs
    framIncludeDocs.Width = lW - framIncludeDocs_W
    For lCount = chkDoc.LBound To chkDoc.UBound
        chkDoc(lCount).Width = lW - chkDoc_W
    Next
    
    'Left framPrnOptions
    framPrnOptions.left = lW - framPrnOptions_L
    
    VisibleFrames True
    
    mbResize = False
    Exit Sub
EH:
    Err.Clear
    Resume Next
End Sub
Private Sub txtDaysAgo_Change()
    SetDaysAgo
End Sub

Private Sub txtDaysAgo_GotFocus()
    goUtil.utSelText txtDaysAgo
    SetDaysAgo
End Sub

Private Sub txtDaysAgo_LostFocus()
    If Not IsNumeric(txtDaysAgo.Text) Then
        txtDaysAgo.Text = 0
    End If
    SetDaysAgo
End Sub

Private Sub SetDaysAgo()
    On Error GoTo EH
    Dim sDate As String
    sDate = Format(DateAdd("d", CDbl(txtDaysAgo.Text), Now()), "MM/DD/YYYY")
    txtDaysAgo.ToolTipText = sDate
    ChkAppAdjDateStamp.ToolTipText = sDate
    AppendDate
    Exit Sub
EH:
    txtDaysAgo.ToolTipText = vbNullString
    ChkAppAdjDateStamp.ToolTipText = vbNullString
End Sub

Private Sub txtDaysAgo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetDaysAgo
End Sub

Private Sub txtPDFPath_Change()
    EnablePrint
End Sub

Private Sub txtPDFPath_GotFocus()
    goUtil.utSelText txtPDFPath
End Sub


Private Sub EnablePrint()
    On Error GoTo EH
    Dim sPath As String
    
    sPath = txtPDFPath.Text
    
    sPath = left(sPath, InStrRev(sPath, "\", , vbBinaryCompare))
    
    If Not goUtil.utFileExists(sPath, True) Then
        If chkPrintToPDF.Value = vbChecked Then
            cmdPrint.Enabled = False
        End If
    Else
        cmdPrint.Enabled = True
        SaveSetting App.EXEName, "Dir", "PDFOutPath", sPath
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub EnablePrint"
End Sub


Private Sub LoadHeader()
    On Error GoTo EH
    Dim bGridOn As Boolean
    'set the columnheaders
    With lvwLossReports
              
        .ColumnHeaders.Add , "DateAsgn", "Date Asgn"
        .ColumnHeaders.Add , "DateAsgnSort", "Sort Date Asgn"
        .ColumnHeaders.Add , "AssignmentType", "Type"
        .ColumnHeaders.Add , "Status", "Status"
        .ColumnHeaders.Add , "CatName", "Cat Name"
        .ColumnHeaders.Add , "CatCode", "Cat Code"
        .ColumnHeaders.Add , "Ajuster", "Adjuster"
        .ColumnHeaders.Add , "ACID", "ACID"
        .ColumnHeaders.Add , "CLIENTNUM", "CLIENTNUM"
        .ColumnHeaders.Add , "IBNUM", "IBNUM"
        .ColumnHeaders.Add , "InsuredName", "Insured Name"
        .ColumnHeaders.Add , "HPhone", "H Phone"
        .ColumnHeaders.Add , "WPhone", "W Phone"
        .ColumnHeaders.Add , "RFormat", "Format"
        .ColumnHeaders.Add , "RSort", "Sort"
        .ColumnHeaders.Add , "RKey", "Key"
        .Sorted = True
        .SortKey = LossReports.RSort - 1
        .SortOrder = lvwAscending
        
        'DateAsgn
        .ColumnHeaders.Item(LossReports.DateAsgn).Width = 1200
        .ColumnHeaders.Item(LossReports.DateAsgn).Alignment = lvwColumnLeft
        'DateAsgnSort
        .ColumnHeaders.Item(LossReports.DateAsgnSort).Width = 0  'Hidden Sort Column
        .ColumnHeaders.Item(LossReports.DateAsgnSort).Alignment = lvwColumnLeft
        'Assignment Type
        .ColumnHeaders.Item(LossReports.AssignmentType).Width = 1200
        .ColumnHeaders.Item(LossReports.AssignmentType).Alignment = lvwColumnLeft
         'Status
        .ColumnHeaders.Item(LossReports.Status).Width = 1500
        .ColumnHeaders.Item(LossReports.Status).Alignment = lvwColumnLeft
        'Cat Name
        .ColumnHeaders.Item(LossReports.CatName).Width = 1500
        .ColumnHeaders.Item(LossReports.CatName).Alignment = lvwColumnLeft
        'Cat Code
        .ColumnHeaders.Item(LossReports.CatCode).Width = 1500
        .ColumnHeaders.Item(LossReports.CatCode).Alignment = lvwColumnLeft
        'Adjuster
        .ColumnHeaders.Item(LossReports.Adjuster).Width = 1500
        .ColumnHeaders.Item(LossReports.Adjuster).Alignment = lvwColumnLeft
        'ACID
        .ColumnHeaders.Item(LossReports.ACID).Width = 1500
        .ColumnHeaders.Item(LossReports.ACID).Alignment = lvwColumnLeft
        'CLIENTNUM
        .ColumnHeaders.Item(LossReports.CLIENTNUM).Width = 1500
        .ColumnHeaders.Item(LossReports.CLIENTNUM).Alignment = lvwColumnLeft
        'IBNUM
        .ColumnHeaders.Item(LossReports.IBNUM).Width = 1500
        .ColumnHeaders.Item(LossReports.IBNUM).Alignment = lvwColumnLeft
        'InsuredName
        .ColumnHeaders.Item(LossReports.InsuredName).Width = 5000
        .ColumnHeaders.Item(LossReports.InsuredName).Alignment = lvwColumnLeft
        'HPhone
        .ColumnHeaders.Item(LossReports.HPhone).Width = 1500
        .ColumnHeaders.Item(LossReports.HPhone).Alignment = lvwColumnLeft
        'WPhone
        .ColumnHeaders.Item(LossReports.WPhone).Width = 1500
        .ColumnHeaders.Item(LossReports.WPhone).Alignment = lvwColumnLeft
        'RFormat
        .ColumnHeaders.Item(LossReports.RFormat).Width = 5000
        .ColumnHeaders.Item(LossReports.RFormat).Alignment = lvwColumnLeft
        'Hidden RSort (Sort will be combination of DateAsgn & InsuredName
        .ColumnHeaders.Item(LossReports.RSort).Width = 0
        .ColumnHeaders.Item(LossReports.RSort).Alignment = lvwColumnLeft
        'Hidden RKey (Key will be File Path)
        .ColumnHeaders.Item(LossReports.RKey).Width = 0
        .ColumnHeaders.Item(LossReports.RKey).Alignment = lvwColumnLeft
    End With
    
    bGridOn = CBool(GetSetting(App.EXEName, "GENERAL", "GRID_ON", False))
    If bGridOn Then
        chkViewGrid.Value = vbChecked
    Else
        chkViewGrid.Value = vbUnchecked
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub LoadHeader"
End Sub

Public Sub Populate()
    On Error GoTo EH
    Dim itmX As listItem
    Dim oLR As V2ECKeyBoard.clsCarLR
    Dim sPrnFileName As String
    Dim sCreateDate As String
    
    'Get collection of Loss Report Objects
    'Add set our memeber collection here too :)
    If mcolLossReports Is Nothing Then
        Set mcolLossReports = moLRs.GetLossReportCol(moProgForm)
    End If
    If goUtil Is Nothing Then
        Exit Sub
    End If
    'Clear any existing items
    lvwLossReports.ListItems.Clear
    If chkViewMess.Value = vbUnchecked Then
        chkViewMess.Enabled = False
    End If
    
    'Add the loss reports to the listview
    'There may be loss reports of derived from varying formats
    If Not mcolLossReports Is Nothing Then
        For Each oLR In mcolLossReports
            'enable the Check box if appropriate
            If Not chkViewMess.Enabled Then
                If TypeOf oLR Is V2ECKeyBoard.clsLossUnknown And oLR.OleType = "HTML" Then
                    chkViewMess.Enabled = True
                End If
            End If
            'If View Messages is checked then only allow the Unknown class that is HTML to Add itself
            'Otherwise Allow all but the unknown class to load.
            If chkViewMess.Value = vbChecked Then
                If (Not TypeOf oLR Is V2ECKeyBoard.clsLossUnknown) Or oLR.OleType <> "HTML" Then
                    GoTo SKIP_LR
                End If
            Else
                If TypeOf oLR Is V2ECKeyBoard.clsLossUnknown And oLR.OleType = "HTML" Then
                    GoTo SKIP_LR
                End If
            End If
            If Not oLR.AdditmX(itmX, lvwLossReports) Then
                'If the Loss report object fails to add itself
                'Then put this error in the list
                sCreateDate = Format(moLRs.GetCreateDate(oLR.PrnKey), "MM/DD/YY")
                sPrnFileName = oLR.PrnKey
                sPrnFileName = Mid(sPrnFileName, InStrRev(sPrnFileName, "\") + 1)
                'Use File create date since we don't know Assigned date
                Set itmX = lvwLossReports.ListItems.Add(, , sCreateDate & " - Error", , LRPic.lrError)
                
                itmX.SubItems(LossReports.DateAsgnSort - 1) = Format(sCreateDate, "YYYY/MM/DD")
                'Use File name instead of ClaimNo
                itmX.SubItems(LossReports.Adjuster - 1) = sPrnFileName & " - Error"
                
                'CLIENTNUM
                itmX.SubItems(LossReports.CLIENTNUM - 1) = sPrnFileName & " - Error"
                
                'CLIENTNUM
                itmX.SubItems(LossReports.IBNUM - 1) = sPrnFileName & " - Error"
                
                'Use Error instead of Insured Name
                itmX.SubItems(LossReports.InsuredName - 1) = "Error"
                
                'Use Error instead of Home Phone
                itmX.SubItems(LossReports.HPhone - 1) = "Error"
                
                'Use Error instead of Work Phone
                itmX.SubItems(LossReports.WPhone - 1) = "Error"
                
                'Use Class name for format
                itmX.SubItems(LossReports.RFormat - 1) = oLR.ClassName & " - Error"
                
                'Sort by File creation Date and File Name instead of by Assigned Date and Insured Name
                'Format the Date so that it will sort by year first
                itmX.SubItems(LossReports.RSort - 1) = Format(sCreateDate, "YY/MM/DD") & sPrnFileName
                
                'Rememeber that the File path is used as the Key.
                'This allows for quick reference to this particular Report
                'when it is selcted from the listview
                itmX.SubItems(LossReports.RKey - 1) = oLR.PrnKey
                
                itmX.Selected = False
            End If
SKIP_LR:
        Next
    End If

CLEANUP:
    'Cleanup

    Set itmX = Nothing
    Exit Sub
EH:
    Set itmX = Nothing
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub Populate"
End Sub


Public Function CLEANUP() As Boolean
    On Error GoTo EH
    Dim oLR As V2ECKeyBoard.clsCarLR
    
    'Since this cleanup could take a while lets fire
    'an event to let APPS know
    moLRs.FireMemoryCleanUpAlert
    
    'Clean up each Loss Report object within Loss reports Collection
    'This should free up allocated memory  better than if we just set the
    'Collection To nothing
    If Not mcolLossReports Is Nothing Then
        For Each oLR In mcolLossReports
            If Not oLR Is Nothing Then
                oLR.CLEANUP
                Set oLR = Nothing
            End If
        Next
        Set mcolLossReports = Nothing
    End If
    moLRs.FireMemoryCleanUpFinished
    'ref to clsLossReports
    If Not moLRs Is Nothing Then
        Set moLRs = Nothing
    End If
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function CleanUp"
    Resume Next
End Function

Private Function EnableAppDoc() As Boolean
    On Error GoTo EH
    Dim oAD As Object
    Dim colAppDocs As Collection
    Dim Doc As udtAppDoc
    
    Set colAppDocs = moLRs.AppDocsCol
    'If we are doing app docs then need to make all controls
    'tagged with appdoc visible and enabled (only enable checkbox controls with captions)
    For Each oAD In Me.Controls
        If TypeOf oAD Is Frame Or TypeOf oAD Is CheckBox Or TypeOf oAD Is Label Then
            If oAD.Tag = APP_DOC Then
                If moLRs.AppDocFlag Then
                    oAD.Visible = True
                    If TypeOf oAD Is CheckBox Then
                        'Look in the AppDoc Collection
                        If colAppDocs.Count > 0 Then
                            If oAD.Index <= colAppDocs.Count - 1 Then
                                Doc = colAppDocs(oAD.Index + 1)
                                oAD.Caption = Doc.DocName
                                oAD.Value = IIf(Doc.Selected, vbChecked, vbUnchecked)
                                oAD.Visible = True
                                If chkChain.Value = vbChecked Then
                                    oAD.Enabled = False
                                Else
                                    oAD.Enabled = True
                                End If
                            End If
                        End If
                        'Enable is it has a Doc name in the caption
                        If oAD.Caption = vbNullString Then
                            oAD.Visible = False
                        End If
                    ElseIf TypeOf oAD Is Label Then
                        If chkChain.Value = vbChecked Then
                            oAD.BackColor = &H8000000F 'Gray Button Face
                        Else
                            oAD.BackColor = &H80000005 'White windows background
                        End If
                    End If
                Else
                    If Not TypeOf oAD Is Frame Then
                        oAD.Visible = False
                    End If
                    oAD.Enabled = False
                End If
            End If
        End If
    Next
    
    EnableAppDoc = True
    
    'CleanUp
    Set colAppDocs = Nothing
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Function EnableAppDoc"
End Function

'Private Sub DeleteItems()
'    On Error GoTo EH
'    Dim itmX As ListItem
'    Dim sFile As String
'    Dim lRet As Long
'    Dim bDel As Boolean 'True if user has answered the Question if they want to delete
'
'    'Check Printing flag first
'    If mbPrinting Then
'        Exit Sub 'Bail
'    End If
'
'    For Each itmX In lvwLossReports.ListItems
'        If itmX.Selected = True Then
'            If Not bDel Then
'                bDel = True
'                lRet = MsgBox("Are you sure you want to delete the selected item(s)?", vbOKCancel + vbQuestion, "Delete Loss Report")
'                If lRet = vbCancel Then
'                    Exit Sub
'                End If
'                Screen.MousePointer = vbHourglass
'            End If
'            'First delete the File
'            sFile = itmX.SubItems(LossReports.RKey - 1)
'            If goUtil.utFileExists(sFile) Then
'                SetAttr sFile, vbNormal
'                Kill sFile
'            End If
'
'            'then remove from mcolLossReports
'            mcolLossReports.Remove sFile
'        End If
'    Next
'
'    'Now populate the listview only if the collection was modified
'    If bDel Then
'        Populate
'    End If
'
'    lvwLossReports.SetFocus
'
'    Screen.MousePointer = vbDefault
'    Exit Sub
'EH:
'    Screen.MousePointer = vbDefault
'    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub DeleteItems"
'End Sub

Private Sub FindText(Optional psFindText As String = vbNullString)
    On Error GoTo EH
    Dim itmX As listItem
    Dim sFind As String
    Dim sTemp As String
    Dim lCount As Long
    Dim bFound As Boolean
    Dim itmXLastFound As listItem
    
    'First unselect current selections
    For Each itmX In lvwLossReports.ListItems
        If itmX.Selected Then
            itmX.Selected = False
        End If
        For lCount = 1 To itmX.ListSubItems.Count
            itmX.ListSubItems(lCount).ReportIcon = Empty
        Next
    Next
    
    'Go through entire list and search each Item and select if the string exists in
    'any of the columns
    If psFindText = vbNullString Then
        sFind = InputBox("Enter search text", "Find", , Me.left, Me.top)
    Else
        sFind = psFindText
    End If
    
    If Trim(sFind) <> vbNullString Then
        For Each itmX In lvwLossReports.ListItems
            sTemp = itmX.Text
            If InStr(1, sTemp, sFind, vbTextCompare) > 0 Then
                bFound = True
                itmX.Selected = True
                Set itmXLastFound = itmX
            End If
            For lCount = 1 To itmX.ListSubItems.Count
                sTemp = itmX.ListSubItems(lCount).Text
                If InStr(1, sTemp, sFind, vbTextCompare) > 0 And lCount < LossReports.RSort Then
                    bFound = True
                    itmX.Selected = True
                    itmX.ListSubItems(lCount).ReportIcon = LRPic.Found
                    Set itmXLastFound = itmX
                End If
            Next
        Next
        
        If Not bFound Then
            MsgBox """" & sFind & """ Not found.", vbInformation + vbOKOnly, "Find"
        Else
            itmXLastFound.EnsureVisible
            lvwLossReports.SetFocus
        End If
    End If
    
    
    'cleanup
    Set itmX = Nothing
    Set itmXLastFound = Nothing
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub FindText"
    Set itmX = Nothing
End Sub
