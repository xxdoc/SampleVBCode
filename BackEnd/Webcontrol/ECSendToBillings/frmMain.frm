VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Send to Billings (Record selection for uploaded claims)"
   ClientHeight    =   6165
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   9750
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   9750
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkShowRebills 
      Caption         =   "Rebills"
      Height          =   375
      Left            =   7800
      TabIndex        =   16
      Top             =   1560
      Width           =   1935
   End
   Begin VB.CheckBox chkHideSent 
      Caption         =   "Hide Sent"
      Height          =   375
      Left            =   7800
      TabIndex        =   15
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CheckBox chkHideSend 
      Caption         =   "Hide Send"
      Height          =   375
      Left            =   7800
      TabIndex        =   14
      Top             =   840
      Width           =   1935
   End
   Begin VB.CheckBox chkHideSnailMail 
      Caption         =   "Hide Snail Mail"
      Height          =   375
      Left            =   7800
      TabIndex        =   13
      Top             =   480
      Width           =   1935
   End
   Begin VB.CheckBox chkHideSkipped 
      Caption         =   "Hide Skipped"
      Height          =   375
      Left            =   7800
      TabIndex        =   12
      Top             =   120
      Width           =   1935
   End
   Begin VB.ComboBox cboByAdjuster 
      Height          =   390
      ItemData        =   "frmMain.frx":08CA
      Left            =   1320
      List            =   "frmMain.frx":08CC
      TabIndex        =   11
      Top             =   3480
      Width           =   6375
   End
   Begin VB.ComboBox cboByCatSite 
      Height          =   390
      ItemData        =   "frmMain.frx":08CE
      Left            =   1320
      List            =   "frmMain.frx":08D0
      TabIndex        =   9
      Top             =   3000
      Width           =   6375
   End
   Begin VB.Frame framSplash 
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2200
      Left            =   1800
      TabIndex        =   34
      Top             =   360
      Visible         =   0   'False
      Width           =   4750
      Begin VB.Image imgLogo 
         Height          =   2100
         Left            =   45
         Picture         =   "frmMain.frx":08D2
         Top             =   45
         Width           =   4650
      End
   End
   Begin VB.Timer Timer_Status 
      Interval        =   500
      Left            =   3240
      Top             =   0
   End
   Begin VB.Frame framEndDate 
      Caption         =   "End Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   3960
      TabIndex        =   4
      Top             =   0
      Width           =   3735
      Begin VB.ComboBox cboEndDate 
         Height          =   390
         Left            =   1590
         TabIndex        =   7
         Top             =   2400
         Width           =   1815
      End
      Begin MSACAL.Calendar calEndDate 
         Height          =   2145
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   3435
         _Version        =   524288
         _ExtentX        =   6059
         _ExtentY        =   3784
         _StockProps     =   1
         BackColor       =   -2147483633
         Year            =   2003
         Month           =   1
         Day             =   23
         DayLength       =   1
         MonthLength     =   0
         DayFontColor    =   0
         FirstDay        =   1
         GridCellEffect  =   1
         GridFontColor   =   10485760
         GridLinesColor  =   -2147483632
         ShowDateSelectors=   -1  'True
         ShowDays        =   -1  'True
         ShowHorizontalGrid=   -1  'True
         ShowTitle       =   0   'False
         ShowVerticalGrid=   -1  'True
         TitleFontColor  =   10485760
         ValueIsNull     =   -1  'True
         BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblEndDate 
         Alignment       =   1  'Right Justify
         Caption         =   "End Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A00000&
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   2400
         Width           =   1215
      End
   End
   Begin VB.Frame framFromDate 
      Caption         =   "Start Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3735
      Begin VB.ComboBox cboStartDate 
         Height          =   390
         ItemData        =   "frmMain.frx":C8BC
         Left            =   1590
         List            =   "frmMain.frx":C8BE
         TabIndex        =   3
         Top             =   2400
         Width           =   1815
      End
      Begin MSACAL.Calendar calStartDate 
         Height          =   2145
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3435
         _Version        =   524288
         _ExtentX        =   6059
         _ExtentY        =   3784
         _StockProps     =   1
         BackColor       =   -2147483633
         Year            =   2003
         Month           =   1
         Day             =   23
         DayLength       =   1
         MonthLength     =   0
         DayFontColor    =   0
         FirstDay        =   1
         GridCellEffect  =   1
         GridFontColor   =   10485760
         GridLinesColor  =   -2147483632
         ShowDateSelectors=   -1  'True
         ShowDays        =   -1  'True
         ShowHorizontalGrid=   -1  'True
         ShowTitle       =   0   'False
         ShowVerticalGrid=   -1  'True
         TitleFontColor  =   10485760
         ValueIsNull     =   -1  'True
         BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Start Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A00000&
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   2415
         Width           =   1215
      End
   End
   Begin VB.Frame framCommands 
      Height          =   2295
      Left            =   120
      TabIndex        =   17
      Top             =   3840
      Width           =   9495
      Begin VB.Frame framSelectDSN 
         Height          =   1935
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Visible         =   0   'False
         Width           =   9255
         Begin VB.ListBox lstDSN 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1020
            Left            =   2400
            TabIndex        =   30
            Top             =   750
            Width           =   6735
         End
         Begin VB.CheckBox chkSQLSVR 
            Caption         =   "SQL SERVER ?"
            Height          =   270
            Left            =   2400
            TabIndex        =   28
            Top             =   360
            Width           =   2775
         End
         Begin VB.Frame framPassword 
            Caption         =   "Password"
            Height          =   735
            Left            =   120
            TabIndex        =   27
            Top             =   1035
            Width           =   2175
            Begin VB.TextBox txtPassWord 
               Height          =   360
               IMEMode         =   3  'DISABLE
               Left            =   120
               PasswordChar    =   "*"
               TabIndex        =   29
               Top             =   240
               Width           =   1935
            End
         End
         Begin VB.Frame framUserID 
            Caption         =   "User ID"
            Height          =   735
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   2175
            Begin VB.TextBox txtUserID 
               Height          =   390
               Left            =   120
               TabIndex        =   26
               Top             =   240
               Width           =   1935
            End
         End
      End
      Begin VB.CommandButton cmdLoad 
         Caption         =   "&Load"
         Enabled         =   0   'False
         Height          =   855
         Left            =   7320
         Picture         =   "frmMain.frx":C8C0
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Send All"
         Top             =   1320
         Width           =   975
      End
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   855
         Left            =   8400
         MaskColor       =   &H00000000&
         Picture         =   "frmMain.frx":CBCA
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Exit"
         Top             =   1320
         Width           =   975
      End
      Begin VB.Frame framProgress 
         Caption         =   "Progress"
         Height          =   975
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   9255
         Begin MSComctlLib.ProgressBar progBar 
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   600
            Width           =   9015
            _ExtentX        =   15901
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   0
         End
         Begin VB.Label lblTime 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6840
            TabIndex        =   33
            Top             =   240
            Width           =   2295
         End
         Begin VB.Label lblWarning 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   360
            Width           =   9015
         End
      End
      Begin VB.Frame framClaimsBillingDBName 
         Caption         =   "Claims Billing Data Base (DSN)"
         Height          =   855
         Left            =   120
         TabIndex        =   21
         Top             =   1320
         Width           =   7095
         Begin VB.CommandButton cmdClaimsBillingDBName 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   6600
            Picture         =   "frmMain.frx":CED4
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Browse"
            Top             =   375
            Width           =   375
         End
         Begin VB.TextBox txtClaimsBillingDBName 
            Height          =   360
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   22
            Top             =   360
            Width           =   6855
         End
      End
   End
   Begin VB.Image imgV2SendToBillings 
      BorderStyle     =   1  'Fixed Single
      Height          =   960
      Left            =   8227
      Picture         =   "frmMain.frx":D01E
      Stretch         =   -1  'True
      ToolTipText     =   "VS 2.0 Send To Billings"
      Top             =   2640
      Width           =   1080
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Adjuster"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A00000&
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Cat Site"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A00000&
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuPrintSetUp 
         Caption         =   "&Printer"
         Begin VB.Menu mnuSetPrinterManually 
            Caption         =   "Setup &Printer (Windows Default)"
         End
         Begin VB.Menu mnuUseWinDefaultPrinter 
            Caption         =   "&Use Windows Default Printer"
         End
      End
      Begin VB.Menu BarPrinter 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private moDSNText As Object
Private mbLoading As Boolean
Private mcolDates As Collection
Private mcolCatSites As Collection
Private mcolAdjusters As Collection


'Timed Flags
Private mbUnloadFlag As Boolean
Private mbCleanUpFlag As Boolean
Private mbCheckPrinter As Boolean
Private mbChangeHide As Boolean

Public Property Let FlagCheckPrinter(pbFlag As Boolean)
    mbCheckPrinter = pbFlag
End Property
Public Property Get FlagCheckPrinter() As Boolean
    FlagCheckPrinter = mbCheckPrinter
End Property

Public Property Let UnloadFlag(pbFlag As Boolean)
    mbUnloadFlag = pbFlag
End Property
Public Property Get UnloadFlag() As Boolean
    UnloadFlag = mbUnloadFlag
End Property

Public Property Let colDates(pcolObject As Collection)
    Set mcolDates = pcolObject
End Property
Public Property Set colDates(pcolObject As Collection)
    Set mcolDates = pcolObject
End Property
Public Property Get colDates() As Collection
    Set colDates = mcolDates
End Property

Public Property Let colCatSites(pcolObject As Collection)
    Set mcolCatSites = pcolObject
End Property
Public Property Set colCatSites(pcolObject As Collection)
    Set mcolCatSites = pcolObject
End Property
Public Property Get colCatSites() As Collection
    Set colCatSites = mcolCatSites
End Property

Public Property Let colAdjusters(pcolObject As Collection)
    Set mcolAdjusters = pcolObject
End Property
Public Property Set colAdjusters(pcolObject As Collection)
    Set mcolAdjusters = pcolObject
End Property
Public Property Get colAdjusters() As Collection
    Set colAdjusters = mcolAdjusters
End Property


Private Sub calEndDate_AfterUpdate()
    On Error GoTo EH
    
    cboEndDate.Text = calEndDate.Value
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, goUtil.gsAppEXEName, Me.Name, "Private Sub calEndDate_AfterUpdate"
End Sub

Private Sub calEndDate_DblClick()
    On Error GoTo EH
    
    cboEndDate.Text = calEndDate.Value
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, goUtil.gsAppEXEName, Me.Name, "Private Sub calEndDate_DblClick"
End Sub

Private Sub CalStartDate_AfterUpdate()
    On Error GoTo EH
    
    cboStartDate.Text = calStartDate.Value
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, goUtil.gsAppEXEName, Me.Name, "Private Sub CalStartDate_AfterUpdate"
End Sub

Private Sub LoadDSNNames(poDSNText As Object, Optional psInit As String)
    On Error GoTo EH
    Dim oReg As V2ECKeyBoard.clsRegSetting
    Dim vDSN As Variant
    Dim vInit As Variant
    Dim lCount As Long
    Dim bShowDSN As Boolean
    Dim lSpace As Long
    
    'Remember what DBNAME we are working on
    Set moDSNText = poDSNText
    
    lstDSN.Clear
    'Load any initial DSNs here
    If Not IsEmpty(psInit) Then
        vInit = Split(psInit, ",")
        For lCount = 0 To UBound(vInit)
            bShowDSN = True
            lstDSN.AddItem vInit(lCount)
        Next
    End If
    
    Set oReg = New V2ECKeyBoard.clsRegSetting
    'Enumerate all the DSN names in the Registry
    vDSN = oReg.EnumValues(HKEY_LOCAL_MACHINE, "SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources")
    'Add them to the List
    For lCount = 0 To UBound(vDSN, 1)
        If vDSN(lCount, 0) <> vbNullString Then
            bShowDSN = True
            lSpace = 20 - Len(vDSN(lCount, 0))
            If lSpace < 0 Then
                lSpace = 0
            End If
            lstDSN.AddItem vDSN(lCount, 0) & Chr(160) & String(lSpace, Chr(32)) & "[" & vDSN(lCount, 1) & "]"
        End If
    Next
    
    
    If bShowDSN Then
        framSelectDSN.Caption = "Select DSN (" & moDSNText.Container.Caption & ")"
        framSelectDSN.Visible = True
        framSelectDSN.ZOrder
        txtUserID.SetFocus
    End If
    
    'CLeanup
    Set oReg = Nothing
    Exit Sub
EH:
    goUtil.utErrorLog Err, goUtil.gsAppEXEName, Me.Name, "Private Sub GetDSNName"
End Sub

Private Sub calStartDate_DblClick()
     On Error GoTo EH
    
    cboStartDate.Text = calStartDate.Value
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, goUtil.gsAppEXEName, Me.Name, "Private Sub calStartDate_DblClick"
End Sub

Private Sub calStartDate_GotFocus()
    calStartDate.BackColor = &H8000000E
End Sub

Private Sub calStartDate_LostFocus()
    calStartDate.BackColor = &H8000000F
End Sub

Private Sub calEndDate_GotFocus()
    calEndDate.BackColor = &H8000000E
End Sub

Private Sub calEndDate_LostFocus()
    calEndDate.BackColor = &H8000000F
End Sub

Private Sub cboByAdjuster_Change()
    EnableLoadClaims
End Sub

Private Sub cboByAdjuster_Click()
    EnableLoadClaims
End Sub

Private Sub cboByCatSite_Change()
    EnableLoadClaims
End Sub

Private Sub cboByCatSite_Click()
    EnableLoadClaims
End Sub

Private Sub cboEndDate_Click()
    EndDateEdit
End Sub

Private Sub cboStartDate_Change()
    StartDateEdit
End Sub

Private Sub cboEndDate_Change()
    EndDateEdit
End Sub

Private Sub cboStartDate_Click()
    StartDateEdit
End Sub

Private Sub StartDateEdit()
    On Error GoTo EH
    
    If IsDate(cboStartDate.Text) Then
        cboStartDate.BackColor = &H80000005
        If Format(cboStartDate.Text, "MM/DD/YYYY") <> Format(calStartDate.Value) Then
            calStartDate.Value = cboStartDate.Text
        End If
    Else
        cboStartDate.BackColor = &HC0C0FF
    End If
    
    EnableLoadClaims
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, goUtil.gsAppEXEName, Me.Name, "Private Sub StartDateEdit"
End Sub

Private Sub EndDateEdit()
    On Error GoTo EH
    
    If IsDate(cboEndDate.Text) Then
        cboEndDate.BackColor = &H80000005
        If Format(cboEndDate.Text, "MM/DD/YYYY") <> Format(calEndDate.Value) Then
            calEndDate.Value = cboEndDate.Text
        End If
    Else
        cboEndDate.BackColor = &HC0C0FF
    End If
    
    EnableLoadClaims
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, goUtil.gsAppEXEName, Me.Name, "Private Sub EndDateEdit"
End Sub

Private Sub chkHideSend_Click()
    On Error GoTo EH
    Dim bHideSendOn As Boolean
    
    If chkHideSend.Value = vbChecked Then
        chkHideSent.Value = vbUnchecked
        bHideSendOn = True
    Else
        
        bHideSendOn = False
    End If
    
    SaveSetting goUtil.gsAppEXEName, "GENERAL", "HIDESEND_ON", bHideSendOn
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, goUtil.gsAppEXEName, Me.Name, "Private Sub Private Sub chkHideSend_Click"
End Sub

Private Sub chkHideSent_Click()
    On Error GoTo EH
    Dim bHideSentOn As Boolean
    
    If chkHideSent.Value = vbChecked Then
        chkHideSend.Value = vbUnchecked
        bHideSentOn = True
    Else
        bHideSentOn = False
    End If
    
    SaveSetting goUtil.gsAppEXEName, "GENERAL", "HIDESENT_ON", bHideSentOn
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, goUtil.gsAppEXEName, Me.Name, "Private Sub Private Sub chkHideSent_Click"
End Sub

Private Sub chkHideSkipped_Click()
    On Error GoTo EH
    Dim bHideSkippedOn As Boolean
    
    If chkHideSkipped.Value = vbChecked Then
        bHideSkippedOn = True
    Else
        bHideSkippedOn = False
    End If
    
    SaveSetting goUtil.gsAppEXEName, "GENERAL", "HIDESKIPPED_ON", bHideSkippedOn
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, goUtil.gsAppEXEName, Me.Name, "Private Sub Private Sub chkHideSkipped_Click"
End Sub

Private Sub chkHideSnailMail_Click()
    On Error GoTo EH
    Dim bHideSnailMailOn As Boolean
    
    If chkHideSnailMail.Value = vbChecked Then
        bHideSnailMailOn = True
    Else
        bHideSnailMailOn = False
    End If
    
    SaveSetting goUtil.gsAppEXEName, "GENERAL", "HIDESNAILMAIL_ON", bHideSnailMailOn
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, goUtil.gsAppEXEName, Me.Name, "Private Sub Private Sub chkHideSnailMail_Click"
End Sub

Private Sub chkShowRebills_Click()
    On Error GoTo EH
    Dim bRebillsOn As Boolean
    
    If chkShowRebills.Value = vbChecked Then
        chkShowRebills.Caption = "&Rebills ON"
        bRebillsOn = True
    Else
        chkShowRebills.Caption = "&Rebills OFF"
        bRebillsOn = False
    End If
    
    SaveSetting goUtil.gsAppEXEName, "GENERAL", "REBILLS_ON", bRebillsOn
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, goUtil.gsAppEXEName, Me.Name, "Private Sub chkShowGrid_Click"
End Sub

Private Sub chkSQLSVR_Click()
    On Error GoTo EH
    
    SaveSetting goUtil.gsAppEXEName, "DBConn", "USE_SQLSERVER", chkSQLSVR.Value
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, goUtil.gsAppEXEName, Me.Name, "Private Sub chkSQLSVR_Click"
End Sub

Private Sub cmdClaimsBillingDBName_Click()
    On Error GoTo EH
    Dim sRet As String
    Dim xPos As Long
    Dim yPos As Long
    xPos = frmMain.left
    xPos = xPos + framCommands.left
    xPos = xPos + framSelectDSN.left
    yPos = frmMain.top
    yPos = yPos + framCommands.top
    yPos = yPos + framSelectDSN.top
    
    sRet = InputBox("Enter Password", "Password", , xPos, yPos)
    If sRet = Format(Now(), "DDYYMM") Then
        LoadDSNNames txtClaimsBillingDBName
    Else
        If sRet <> vbNullString Then
            Err.Raise -999, , "Invalid Password!"
        End If
    End If
    Exit Sub
EH:
    goUtil.utErrorLog Err, goUtil.gsAppEXEName, Me.Name, "Private Sub cmdClaimsBillingDBName_Click"
End Sub

Private Sub cmdExit_Click()
    On Error GoTo EH
    'Check to see if just exiting the DSN selector
    If framSelectDSN.Visible Then
        framSelectDSN.Visible = False
        cboStartDate.SetFocus
        Exit Sub
    End If
    
    If MsgBox("Are you sure you want to Exit?", vbQuestion + vbYesNo, "Are You Sure?") = vbYes Then
        mbUnloadFlag = True
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, goUtil.gsAppEXEName, Me.Name, "Private Sub cmdExit_Click"
End Sub

Private Sub cmdLoad_Click()
    On Error GoTo EH
    Dim sStartDate As String
    Dim sEndDate As String
    Dim sCatSite As String
    Dim sAdjuster As String
    
    sStartDate = cboStartDate.Text
    sEndDate = cboEndDate.Text
    sCatSite = cboByCatSite.Text
    sAdjuster = cboByAdjuster.Text
    cmdLoad.Enabled = False
    framFromDate.Enabled = False
    framEndDate.Enabled = False
    cboStartDate.Enabled = False
    cboEndDate.Enabled = False
    cboByCatSite.Enabled = False
    cboByAdjuster.Enabled = False
    goECBill.progBar = frmMain.progBar
    
    If goECBill.SendToBillings(sStartDate, sEndDate, sCatSite, sAdjuster) Then
        If UnloadFlag Then
            Exit Sub
        End If
        EnableLoadClaims
        cboStartDate.Enabled = True
        cboEndDate.Enabled = True
        cboByCatSite.Enabled = True
        cboByAdjuster.Enabled = True
        framFromDate.Enabled = True
        framEndDate.Enabled = True
        Me.WindowState = vbNormal
    End If
    
    Exit Sub
EH:
    If Not goUtil Is Nothing Then
        goUtil.utErrorLog Err, goUtil.gsAppEXEName, Me.Name, "Private Sub cmdLoad_Click"
    Else
        Err.Clear
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo EH
    mbLoading = True
    goUtil.utFormWinRegPos goUtil.gsAppEXEName, Me
    
    LoadDefaultSettings
    'Enable Printer Timer It monitors For User changing Default Windows Printer
    InitDefaultPrintMenu
    
    mbLoading = False
    Exit Sub
EH:
    mbLoading = False
    goUtil.utErrorLog Err, goUtil.gsAppEXEName, Me.Name, "Private Sub Form_Load"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo EH
    
    Me.WindowState = vbNormal
    If UnloadMode = vbFormControlMenu Then
        If MsgBox("Are you sure you want to Exit?", vbQuestion + vbYesNo, "Are You Sure?") = vbYes Then
            mbUnloadFlag = True
        End If
        Cancel = True
        Exit Sub
    End If
    
    mbUnloadFlag = True
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, goUtil.gsAppEXEName, Me.Name, "Private Sub Form_QueryUnload"
End Sub

Private Sub Form_Resize()
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo EH
    
    goUtil.utFormWinRegPos goUtil.gsAppEXEName, Me, True
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, goUtil.gsAppEXEName, Me.Name, "Private Sub Form_Unload"
End Sub

Private Sub Image1_Click()

End Sub

Private Sub lstDSN_Click()
    On Error GoTo EH
    
    lstDSN.ToolTipText = lstDSN.Text
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, goUtil.gsAppEXEName, Me.Name, "Private Sub lstDSN_Click"
End Sub

Private Sub lstDSN_DblClick()
    On Error GoTo EH
    
    CHangeDSN
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, goUtil.gsAppEXEName, Me.Name, "Private Sub lstDSN_DblClick"
End Sub

Private Sub lstDSN_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo EH
    
    If KeyCode = vbKeyReturn Then
        CHangeDSN
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, goUtil.gsAppEXEName, Me.Name, "Private Sub lstDSN_KeyDown"
End Sub

Private Sub CHangeDSN()
    On Error GoTo EH
    Dim sDSN As String
    
    sDSN = lstDSN.Text
    If InStr(1, sDSN, Chr(160), vbBinaryCompare) > 0 Then
        sDSN = left(sDSN, InStrRev(sDSN, Chr(160)) - 1)
    End If
    framSelectDSN.Visible = False
    cboStartDate.SetFocus
    moDSNText.Text = sDSN
    Exit Sub
EH:
    goUtil.utErrorLog Err, goUtil.gsAppEXEName, Me.Name, "Private Sub CHangeDSN"
End Sub

Private Sub mnuExit_Click()
    If MsgBox("Are you sure you want to Exit?", vbQuestion + vbYesNo, "Are You Sure?") = vbYes Then
        mbUnloadFlag = True
    End If
End Sub

Private Sub mnuSetPrinterManually_Click()
    On Error GoTo EH
    ShowPrinter Me.hWnd, False
    mnuUseWinDefaultPrinter.Checked = True
    FlagCheckPrinter = True
    Exit Sub
EH:
    goUtil.utErrorLog Err, goUtil.gsAppEXEName, Me.Name, "Private Sub mnuSetPrinterManually_Click"
End Sub

Private Sub mnuUseWinDefaultPrinter_Click()
    On Error GoTo EH
    
    If mnuUseWinDefaultPrinter.Checked Then
        mnuUseWinDefaultPrinter.Checked = False
        SaveSetting goUtil.gsAppEXEName, "PRINTER", "INIT_WITH_WIN_DEFAULT_PRINTER", False
        FlagCheckPrinter = False
    Else
        mnuUseWinDefaultPrinter.Checked = True
        SaveSetting goUtil.gsAppEXEName, "PRINTER", "INIT_WITH_WIN_DEFAULT_PRINTER", True
        FlagCheckPrinter = True
    End If
        
    Exit Sub
EH:
    goUtil.utErrorLog Err, goUtil.gsAppEXEName, Me.Name, "Private Sub mnuUseWinDefaultPrinter_Click"
End Sub

Private Sub Timer_Status_Timer()
    On Error GoTo EH
    
    'Check for Timed Flags
    
    'Cleanup, Ie shutting down.
    If mbCleanUpFlag Then
        Timer_Status.Enabled = False
        CleanUpAndExit
        Exit Sub
    End If
    
    'Check for The unload flag
    If mbUnloadFlag And Not mbCleanUpFlag Then
        'Set the Shutting Down Flag
        mbCleanUpFlag = True
        lblWarning.ForeColor = &H80000012
        lblWarning.Caption = "Exiting... Please Wait."
    End If
    
    'Printer
    If FlagCheckPrinter Then
        CheckPrinter
    End If
    
    
    lblTime.Caption = Now()
    
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, goUtil.gsAppEXEName, Me.Name, "Private Sub Timer_Status_Timer"
End Sub

Private Sub CheckPrinter(Optional pbIgnoreDefault As Boolean)
    On Error GoTo EH
    Dim sDefaultPrinter As String
    Dim nret As Integer
    Dim sRet As String
    Dim lPos As Long
    Dim sTemp As String
    
    If pbIgnoreDefault Then
        sDefaultPrinter = GetSetting(goUtil.gsAppEXEName, "PRINTER", "PRINTER_NAME", vbNullString)
    Else
        'Get Default Printer Name
        sDefaultPrinter = Space(255)
        nret = GetProfileString("Windows", ByVal "device", "", sDefaultPrinter, Len(sDefaultPrinter))
        'Trim it
        If nret Then
            sDefaultPrinter = left(sDefaultPrinter, InStr(sDefaultPrinter, ",") - 1)
        End If
    End If
    
    sTemp = mnuUseWinDefaultPrinter.Caption
    lPos = InStr(1, sTemp, "(", vbTextCompare)
    If lPos = 0 Then
        sTemp = sTemp & " (" & sDefaultPrinter & ")"
    Else
        sTemp = left(sTemp, lPos)
        sTemp = sTemp & sDefaultPrinter & ")"
    End If
    mnuUseWinDefaultPrinter.Caption = sTemp
    
    sRet = GetSetting(goUtil.gsAppEXEName, "PRINTER", "PRINTER_NAME", vbNullString)
    
    If StrComp(sRet, sDefaultPrinter, vbTextCompare) = 0 Then
        Exit Sub
    Else
        goUtil.utSaveDefaultPrinterSettings goUtil.gsAppEXEName, sDefaultPrinter
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, goUtil.gsAppEXEName, Me.Name, "Private Sub TimerPrinter_Timer"
End Sub

Private Sub txtClaimsBillingDBName_Change()
    On Error GoTo EH
    
    SaveSetting goUtil.gsAppEXEName, "DBConn", "Claims", txtClaimsBillingDBName.Text
    If Not mbLoading Then
        ShowSplash
        goECBill.PopulateLookUp
        Me.PopulateLookupCbo
        HideSplash
        Me.EnableLoadClaims
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, goUtil.gsAppEXEName, Me.Name, "Private Sub txtClaimsBillingDBName_Change"
    Me.EnableLoadClaims
End Sub

Private Sub LoadDefaultSettings()
    On Error GoTo EH
    Dim sUserID As String
    Dim sPassWord As String
    Dim bHideSkippedOn As Boolean
    Dim bHideSnailMailOn As Boolean
    Dim bHideSendOn As Boolean
    Dim bHideSentOn As Boolean
    Dim bRebillsOn As Boolean
    
    calStartDate.Value = Now
    calEndDate.Value = Now
    
    txtClaimsBillingDBName.Text = GetSetting(goUtil.gsAppEXEName, "DBConn", "Claims", vbNullString)
    sUserID = goUtil.utGetECSCryptSetting(goUtil.gsAppEXEName, "DBConn", "USERID")
    If sUserID = Chr(32) Then
        sUserID = vbNullString
    End If
    txtUserID.Text = sUserID
    sPassWord = goUtil.utGetECSCryptSetting(goUtil.gsAppEXEName, "DBConn", "PASSWORD")
    If sPassWord = Chr(32) Then
        sPassWord = vbNullString
    End If
    txtPassWord.Text = sPassWord
    chkSQLSVR.Value = GetSetting(goUtil.gsAppEXEName, "DBConn", "USE_SQLSERVER", vbUnchecked)
    
    bHideSkippedOn = CBool(GetSetting(goUtil.gsAppEXEName, "GENERAL", "HIDESKIPPED_ON", False))
    If bHideSkippedOn Then
        chkHideSkipped.Value = vbChecked
    Else
        chkHideSkipped.Value = vbUnchecked
    End If
    
    bHideSnailMailOn = CBool(GetSetting(goUtil.gsAppEXEName, "GENERAL", "HIDESNAILMAIL_ON", False))
    If bHideSnailMailOn Then
        chkHideSnailMail.Value = vbChecked
    Else
        chkHideSnailMail.Value = vbUnchecked
    End If
    
    bHideSendOn = CBool(GetSetting(goUtil.gsAppEXEName, "GENERAL", "HIDESEND_ON", False))
    If bHideSendOn Then
        chkHideSend.Value = vbChecked
    Else
        chkHideSend.Value = vbUnchecked
    End If
    
    bHideSentOn = CBool(GetSetting(goUtil.gsAppEXEName, "GENERAL", "HIDESENT_ON", False))
    If bHideSentOn Then
        chkHideSent.Value = vbChecked
    Else
        chkHideSent.Value = vbUnchecked
    End If
    
    bRebillsOn = CBool(GetSetting(goUtil.gsAppEXEName, "GENERAL", "REBILLS_ON", False))
    If bRebillsOn Then
        chkShowRebills.Value = vbChecked
    Else
        chkShowRebills.Value = vbUnchecked
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, goUtil.gsAppEXEName, Me.Name, "Private Sub LoadDefaultSettings"
End Sub

Public Sub EnableLoadClaims()
    On Error GoTo EH
    Dim bEnable As Boolean
    Dim sMess As String
    Dim lDays As Long
    Dim sStartDate As String
    Dim sEndDate As String
    Dim lMaxDays As Long
    
    lMaxDays = GetSetting(goUtil.gsAppEXEName, "SETTINGS", "MAX_DAYS", 31)
    
    sStartDate = cboStartDate.Text
    sEndDate = cboEndDate.Text
    sMess = vbNullString
    lblWarning.ForeColor = &HFF&
    'Check for Error Loading Claims DSN
    If framSplash.Visible Then
        sMess = "ERROR !"
        bEnable = False
    Else
        If sStartDate <> vbNullString And sEndDate <> vbNullString Then
            If cboStartDate.BackColor = &H80000005 And cboEndDate.BackColor = &H80000005 Then
                If IsDate(sStartDate) And IsDate(sEndDate) Then
                    If CDate(sEndDate) >= CDate(sStartDate) Then
                        'Limit the number of days to span to a default
                        lDays = DateDiff("d", sStartDate, sEndDate) + 1
                        If lDays <= lMaxDays Then
                            lblWarning.ForeColor = &H80000012
                            sMess = "Current selection spans " & lDays & " days."
                            bEnable = True
                        Else
                            sMess = "Maximum date span is " & lMaxDays & " days. "
                            sMess = sMess & "Current selection spans " & lDays & " days."
                        End If
                    Else
                        sMess = "End Date must same day or a day after the Start Date."
                    End If
                End If
            Else
                sMess = "Invalid date format."
            End If
        End If
    End If
    
    'If dates pass then check for the BYcatsite and ByAdjusters
    If bEnable Then
        If cboByCatSite.ListIndex = -1 Then
            bEnable = False
            sMess = sMess & " Select Cat Site! "
        End If
        If cboByAdjuster.ListIndex = -1 Then
            bEnable = False
            sMess = sMess & " Select Adjuster! "
        End If
    End If
    
    lblWarning.Caption = sMess
    cmdLoad.Enabled = bEnable
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, goUtil.gsAppEXEName, Me.Name, "Public Sub EnableLoadClaims"
End Sub

Public Sub PopulateLookupCbo()
    On Error GoTo EH
    Dim vString As Variant
    
    'Dates
    If Not mcolDates Is Nothing Then
        cboStartDate.Clear
        cboEndDate.Clear
        For Each vString In mcolDates
            cboStartDate.AddItem vString
            cboEndDate.AddItem vString
        Next
    End If
    
    'Cat Sites
    If Not mcolCatSites Is Nothing Then
        cboByCatSite.Clear
        cboByCatSite.AddItem "(--By All Cat Sites--)"
        For Each vString In mcolCatSites
            cboByCatSite.AddItem vString
        Next
    End If
    
    'Adjusters
    If Not mcolAdjusters Is Nothing Then
        cboByAdjuster.Clear
        cboByAdjuster.AddItem "(--By All Adjusters--)"
        For Each vString In mcolAdjusters
            cboByAdjuster.AddItem vString
        Next
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, goUtil.gsAppEXEName, Me.Name, "Public Sub PopulateLookupCbo"
End Sub

Private Sub InitDefaultPrintMenu()
    On Error GoTo EH
    Dim sUseDefault As String
    
    sUseDefault = GetSetting(goUtil.gsAppEXEName, "PRINTER", "INIT_WITH_WIN_DEFAULT_PRINTER", vbNullString)
    
    If sUseDefault = vbNullString Then
        SaveSetting goUtil.gsAppEXEName, "PRINTER", "INIT_WITH_WIN_DEFAULT_PRINTER", True
        mnuUseWinDefaultPrinter.Checked = True
        FlagCheckPrinter = True
    ElseIf CBool(sUseDefault) Then
        mnuUseWinDefaultPrinter.Checked = True
        FlagCheckPrinter = True
    Else
        mnuUseWinDefaultPrinter.Checked = False
        FlagCheckPrinter = False
        CheckPrinter True
    End If

    
    Exit Sub
EH:
    goUtil.utErrorLog Err, goUtil.gsAppEXEName, Me.Name, "Private Sub InitDefaultPrintMenu"
End Sub

Private Sub txtPassWord_GotFocus()
    goUtil.utSelText txtPassWord
End Sub

Private Sub txtUserID_Change()
    On Error GoTo EH
    Dim sUserID As String
    
    sUserID = txtUserID.Text
    If sUserID = vbNullString Then
        sUserID = Chr(32)
    End If
    
    goUtil.utSaveECSCryptSetting goUtil.gsAppEXEName, "DBConn", "USERID", sUserID
    Exit Sub
EH:
    goUtil.utErrorLog Err, goUtil.gsAppEXEName, Me.Name, "Private Sub txtUserID_Change"
End Sub

Private Sub txtPassWord_Change()
    On Error GoTo EH
    Dim sPassWord As String
    
    sPassWord = txtPassWord.Text
    If sPassWord = vbNullString Then
        sPassWord = Chr(32)
    End If
    
    goUtil.utSaveECSCryptSetting goUtil.gsAppEXEName, "DBConn", "PASSWORD", sPassWord
    Exit Sub
EH:
    goUtil.utErrorLog Err, goUtil.gsAppEXEName, Me.Name, "Private Sub txtPassWord_Change"
End Sub

Private Sub txtUserID_GotFocus()
    goUtil.utSelText txtUserID
End Sub
