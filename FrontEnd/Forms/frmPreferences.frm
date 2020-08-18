VERSION 5.00
Begin VB.Form frmPreferences 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Preferences"
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10740
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPreferences.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   10740
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame framEasyClaimNavigator 
      Appearance      =   0  'Flat
      Caption         =   "Easy Claim Navigator"
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
      Height          =   7335
      Left            =   6240
      TabIndex        =   40
      Top             =   0
      Width           =   4335
      Begin VB.Frame framReportOptions 
         Caption         =   "Report Options"
         Height          =   855
         Left            =   120
         TabIndex        =   47
         Top             =   1800
         Width           =   4095
         Begin VB.OptionButton optPrintPreviewAdobe 
            Caption         =   "Print Preview (Use Adobe Reader)"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   240
            Width           =   3855
         End
         Begin VB.OptionButton optPrintPreviewActive 
            Caption         =   "Print Preview (Use Active Report Viewer)"
            Height          =   255
            Left            =   120
            TabIndex        =   49
            ToolTipText     =   "Some Operating systems may encounter problems using Active Report Viewer.  If you do experience problems, please use Adobe Reader."
            Top             =   480
            Width           =   3855
         End
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&Ok"
         Height          =   375
         Left            =   1200
         TabIndex        =   55
         Top             =   6840
         Width           =   1455
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   2760
         TabIndex        =   56
         Top             =   6840
         Width           =   1455
      End
      Begin VB.CommandButton cmdWallPaperImage 
         Enabled         =   0   'False
         Height          =   350
         Left            =   3840
         Picture         =   "frmPreferences.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Browse"
         Top             =   1380
         Width           =   375
      End
      Begin VB.TextBox txtWallPaperImage 
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   45
         Text            =   "(Blank Wall Paper)"
         Top             =   1365
         Width           =   3720
      End
      Begin VB.CheckBox chkAlwaysOnTop 
         Alignment       =   1  'Right Justify
         Caption         =   "Always On Top"
         Height          =   255
         Left            =   2280
         TabIndex        =   43
         ToolTipText     =   "Show Navigator Always On Top"
         Top             =   720
         Width           =   1935
      End
      Begin VB.CheckBox chkUseWallPaper 
         Alignment       =   1  'Right Justify
         Caption         =   "Use Wall Paper"
         Height          =   255
         Left            =   2280
         TabIndex        =   44
         Top             =   1005
         Width           =   1935
      End
      Begin VB.OptionButton optAlignRight 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "RIGHT"
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
         Height          =   855
         Left            =   1185
         Picture         =   "frmPreferences.frx":08BC
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton optAlignLeft 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "LEFT"
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
         Height          =   855
         Left            =   120
         Picture         =   "frmPreferences.frx":0BC6
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   360
         Width           =   975
      End
      Begin VB.Frame FramPrintListOpt 
         Caption         =   "Show Item List Report Format Options"
         Height          =   855
         Left            =   120
         TabIndex        =   50
         Top             =   2640
         Width           =   4095
         Begin VB.OptionButton optShowListXLS 
            Caption         =   "Use Excel Format (.xls)"
            Height          =   255
            Left            =   720
            TabIndex        =   51
            ToolTipText     =   "Some Operating systems may encounter problems using Active Report Viewer.  If you do experience problems, please use Adobe Reader."
            Top             =   240
            Width           =   3255
         End
         Begin VB.OptionButton optShowListTXT 
            Caption         =   "Use Text Format (.txt)"
            Height          =   255
            Left            =   720
            TabIndex        =   52
            Top             =   480
            Width           =   3255
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   120
            Picture         =   "frmPreferences.frx":0ED0
            Top             =   240
            Width           =   480
         End
      End
      Begin VB.Frame FramErrorHandler 
         Caption         =   "Error Handler"
         Height          =   615
         Left            =   120
         TabIndex        =   53
         Top             =   3480
         Width           =   4095
         Begin VB.CheckBox chkUseSilentError 
            Caption         =   "Use Silent Error Handling"
            Height          =   255
            Left            =   120
            TabIndex        =   54
            ToolTipText     =   "Show Navigator Always On Top"
            Top             =   240
            Width           =   3855
         End
      End
   End
   Begin VB.Frame framADJ 
      Appearance      =   0  'Flat
      Caption         =   "Adjuster Information"
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
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      Begin VB.TextBox txtTeamLeaderSup 
         Height          =   375
         Left            =   2400
         MaxLength       =   50
         TabIndex        =   26
         Top             =   4320
         Width           =   3465
      End
      Begin VB.TextBox txtADJOtherPostCode 
         Height          =   375
         Left            =   2400
         MaxLength       =   20
         TabIndex        =   24
         Top             =   3960
         Width           =   3465
      End
      Begin VB.TextBox txtADJZip4 
         Height          =   375
         Left            =   4905
         MaxLength       =   4
         TabIndex        =   22
         Top             =   3600
         Width           =   960
      End
      Begin VB.TextBox txtADJZip 
         Height          =   375
         Left            =   2400
         MaxLength       =   5
         TabIndex        =   20
         Top             =   3600
         Width           =   960
      End
      Begin VB.TextBox txtADJState 
         Height          =   375
         Left            =   2400
         MaxLength       =   50
         TabIndex        =   18
         Top             =   3240
         Width           =   960
      End
      Begin VB.TextBox txtADJCity 
         Height          =   375
         Left            =   2400
         MaxLength       =   50
         TabIndex        =   16
         Top             =   2880
         Width           =   3465
      End
      Begin VB.TextBox txtADJAddress 
         Height          =   375
         Left            =   2400
         MaxLength       =   200
         TabIndex        =   14
         Top             =   2520
         Width           =   3465
      End
      Begin VB.TextBox txtADJEmergencyPhone 
         Height          =   375
         Left            =   2400
         MaxLength       =   50
         TabIndex        =   12
         Top             =   2160
         Width           =   3465
      End
      Begin VB.TextBox txtADJContactPhone 
         Height          =   375
         Left            =   2400
         MaxLength       =   50
         TabIndex        =   10
         Top             =   1800
         Width           =   3465
      End
      Begin VB.TextBox txtADJEmail 
         Height          =   375
         Left            =   2400
         MaxLength       =   50
         TabIndex        =   8
         Top             =   1440
         Width           =   3465
      End
      Begin VB.TextBox txtSSN 
         Height          =   375
         Left            =   2400
         MaxLength       =   9
         TabIndex        =   6
         Top             =   1080
         Width           =   3465
      End
      Begin VB.TextBox txtLName 
         Height          =   375
         Left            =   2400
         MaxLength       =   50
         TabIndex        =   4
         Top             =   720
         Width           =   3465
      End
      Begin VB.TextBox txtFName 
         Height          =   375
         Left            =   2400
         MaxLength       =   50
         TabIndex        =   2
         Top             =   360
         Width           =   3465
      End
      Begin VB.Label lblGlobalPref 
         Caption         =   "Other Post Code"
         Height          =   255
         Index           =   16
         Left            =   120
         TabIndex        =   23
         Top             =   4080
         Width           =   4815
      End
      Begin VB.Label lblGlobalPref 
         Caption         =   "Zip4"
         Height          =   255
         Index           =   15
         Left            =   4440
         TabIndex        =   21
         Top             =   3720
         Width           =   1455
      End
      Begin VB.Label lblGlobalPref 
         Caption         =   "Zip"
         Height          =   255
         Index           =   14
         Left            =   120
         TabIndex        =   19
         Top             =   3720
         Width           =   3375
      End
      Begin VB.Label lblGlobalPref 
         Caption         =   "State"
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   17
         Top             =   3360
         Width           =   3255
      End
      Begin VB.Label lblGlobalPref 
         Caption         =   "City"
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   15
         Top             =   3000
         Width           =   4815
      End
      Begin VB.Label lblGlobalPref 
         Caption         =   "Address"
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   13
         Top             =   2640
         Width           =   4815
      End
      Begin VB.Label lblGlobalPref 
         Caption         =   "Emergency Phone"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   11
         Top             =   2280
         Width           =   4815
      End
      Begin VB.Label lblGlobalPref 
         Caption         =   "Team Leader/Supervisor"
         Height          =   255
         Index           =   13
         Left            =   120
         TabIndex        =   25
         Top             =   4440
         Width           =   4815
      End
      Begin VB.Label lblGlobalPref 
         Caption         =   "Contact Phone"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   9
         Top             =   1920
         Width           =   4815
      End
      Begin VB.Label lblGlobalPref 
         Caption         =   "E-Mail"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   1560
         Width           =   4815
      End
      Begin VB.Label lblGlobalPref 
         Caption         =   "Social Security Number"
         Height          =   255
         Index           =   4
         Left            =   90
         TabIndex        =   5
         Top             =   1200
         Width           =   3885
      End
      Begin VB.Label lblGlobalPref 
         Caption         =   "Last Name"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   4815
      End
      Begin VB.Label lblGlobalPref 
         Caption         =   "First Name"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   4815
      End
   End
   Begin VB.Frame framWebControlID 
      Appearance      =   0  'Flat
      Caption         =   "Internet Security"
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
      Height          =   2415
      Left            =   120
      TabIndex        =   28
      Top             =   4920
      Width           =   6015
      Begin VB.CheckBox chkUseSSL 
         Caption         =   "Use SSL (HTTPS)"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         ToolTipText     =   "Show Navigator Always On Top"
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox txtLogonHintAnswer 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2400
         MaxLength       =   50
         PasswordChar    =   "*"
         TabIndex        =   39
         Top             =   1800
         Width           =   3465
      End
      Begin VB.TextBox txtLogonHintQuestion 
         Height          =   375
         Left            =   2400
         MaxLength       =   50
         TabIndex        =   37
         Top             =   1440
         Width           =   3465
      End
      Begin VB.CommandButton cmdEnterPass 
         DownPicture     =   "frmPreferences.frx":1312
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   5475
         Picture         =   "frmPreferences.frx":145C
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   1100
         Width           =   375
      End
      Begin VB.TextBox txtPass 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2400
         Locked          =   -1  'True
         MaxLength       =   50
         PasswordChar    =   "*"
         TabIndex        =   34
         Top             =   1080
         Width           =   3465
      End
      Begin VB.TextBox txtUserName 
         Height          =   375
         Left            =   2400
         MaxLength       =   15
         TabIndex        =   32
         Top             =   720
         Width           =   3465
      End
      Begin VB.TextBox txtWebHost 
         Height          =   375
         Left            =   2400
         MaxLength       =   50
         TabIndex        =   30
         Top             =   360
         Width           =   3465
      End
      Begin VB.Label lblGlobalPref 
         Caption         =   "Logon Hint Answer"
         Height          =   255
         Index           =   17
         Left            =   120
         TabIndex        =   38
         Top             =   1920
         Width           =   4695
      End
      Begin VB.Label lblGlobalPref 
         Caption         =   "Logon Hint Question ?"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   36
         Top             =   1560
         Width           =   4695
      End
      Begin VB.Label lblGlobalPref 
         Caption         =   "Password"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   33
         Top             =   1200
         Width           =   4845
      End
      Begin VB.Label lblGlobalPref 
         Caption         =   "User Name"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   31
         Top             =   840
         Width           =   4695
      End
      Begin VB.Label lblGlobalPref 
         Caption         =   "Web Host"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   29
         Top             =   480
         Width           =   4815
      End
   End
End
Attribute VB_Name = "frmPreferences"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mbLoading As Boolean

Private Property Get msClassName() As String
    msClassName = Me.Name
End Property

Private Sub chkUseSilentError_Click()
    On Error GoTo EH
    If mbLoading Then
        Exit Sub
    End If
    Dim bUseSilentError As Boolean
    
    If chkUseSilentError.Value = vbChecked Then
        bUseSilentError = True
    Else
        bUseSilentError = False
    End If
    SaveSetting "ECS", "Dir", "SILENT_ERROR", CStr(bUseSilentError)
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub chkUseSSL_Click"
End Sub

Private Sub chkUseSSL_Click()
    On Error GoTo EH
    If mbLoading Then
        Exit Sub
    End If
    Dim bUseSSL As Boolean
    
    If chkUseSSL.Value = vbChecked Then
        bUseSSL = True
    Else
        bUseSSL = False
    End If
    SaveSetting "ECS", "WEB_SECURITY", "USE_SSL", bUseSSL
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub chkUseSSL_Click"
End Sub

Private Sub chkAlwaysOnTop_Click()
    On Error GoTo EH
    If mbLoading Then
        Exit Sub
    End If
    Dim bAlwaysOnTop As Boolean
    
    If chkAlwaysOnTop.Value = vbChecked Then
        bAlwaysOnTop = True
    Else
        bAlwaysOnTop = False
    End If
    
    SaveSetting App.EXEName, "GENERAL", "ECTRAY_ALWAYS_ON_TOP", bAlwaysOnTop
    goUtil.utAlwaysOnTop gfrmECTray, bAlwaysOnTop
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub chkAlwaysOnTop_Click"
End Sub

Private Sub chkUseWallPaper_Click()
    On Error GoTo EH
    If mbLoading Then
        Exit Sub
    End If
    Dim bUseWallPaper As Boolean
    
    If chkUseWallPaper = vbChecked Then
        bUseWallPaper = True
        txtWallPaperImage.Enabled = True
        cmdWallPaperImage.Enabled = True
        gfrmECTray.LoadWallPaperImage
    Else
        txtWallPaperImage.Enabled = False
        cmdWallPaperImage.Enabled = False
        bUseWallPaper = False
    End If
    
    SaveSetting App.EXEName, "GENERAL", "USE_WALL_PAPER", bUseWallPaper
    
    gfrmECTray.ShowWallPaper
    
    If bUseWallPaper Then
        gfrmECTray.POSWallPic
        gfrmECTray.ShowAll
        AppActivate Me.Caption, True
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub chkUseWallPaper_Click"
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdWallPaperImage_Click()
    On Error GoTo EH
    Dim sPath As String
    Dim sSelFile As String
    Dim lFileSize As Long
    Dim sMyFilter As String
    
    sMyFilter = sMyFilter & "All File Types" & " (*." & "*" & ")" & SD & "*." & "*" & SD
    sMyFilter = sMyFilter & "Bit Map Files" & " (*." & "bmp" & ")" & SD & "*." & "bmp" & SD
    sMyFilter = sMyFilter & "Graphic Interchange Format" & " (*." & "gif" & ")" & SD & "*." & "gif" & SD
    sMyFilter = sMyFilter & "JPEG File Interchange Format" & " (*." & "jpg" & ")" & SD & "*." & "jpg" & SD
   
    
    sPath = goUtil.utGetPath(App.EXEName, "WallPaperImagePath", "Browse to the Wall Paper Image you want to use", "CLICK OPEN TO SAVE", sPath, Me.hWnd, sMyFilter, sSelFile)
    'Since we are using a windows form here need to check
    'for easy Claim shut down
    If gfrmECTray.FlagShutDownEasyClaim Then
        Exit Sub
    End If
    If goUtil.utFileExists(sSelFile) Then
        On Error Resume Next
        'CHeck for File SIze first
        lFileSize = FileLen(sSelFile) / 1000
        If lFileSize > 210 Then
            Err.Raise -999
        End If
        If Err.Number = 0 Then
            LoadPicture sSelFile
        End If
        If Err.Number <> 0 Or InStr(1, sSelFile, "\") = 0 Then
            Err.Clear
            If lFileSize > 210 Then
                Err.Raise -999, , lFileSize & "KB Image too big! Maximum allowed is 200KB."
            End If
            'Try the Path and selfile
            If Err.Number = 0 Then
                LoadPicture sPath & "\" & sSelFile
            End If
            If Err.Number <> 0 Then
                MsgBox Err.Description, vbExclamation, "Please Try Again"
                Err.Clear
                On Error GoTo EH
                GoTo BLANK
            Else
                sSelFile = sPath & "\" & sSelFile
                On Error GoTo EH
            End If
            
        Else
            On Error GoTo EH
        End If
        SaveSetting App.EXEName, "DIR", "WALL_PAPER_IMAGE", sSelFile
        If InStr(1, sSelFile, "\", vbBinaryCompare) > 0 Then
            sSelFile = Mid(sSelFile, InStrRev(sSelFile, "\") + 1)
        End If
        txtWallPaperImage.Text = sSelFile
    Else
BLANK:
        txtWallPaperImage.Text = "(Blank Wall Paper)"
        SaveSetting App.EXEName, "DIR", "WALL_PAPER_IMAGE", txtWallPaperImage.Text
    End If
    
    gfrmECTray.LoadWallPaperImage
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdWallPaperImage_Click"
End Sub

Private Sub optAlignLeft_Click()
    On Error GoTo EH
    If mbLoading Then
        Exit Sub
    End If
    SaveSetting App.EXEName, "GENERAL", "NAV_SCREEN_POS", "LEFT"
    gfrmECTray.PosECTRAY
    gfrmECTray.ShowMe , False
    Me.SetFocus
    optAlignLeft.SetFocus
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub optAlignLeft_Click"
End Sub

Private Sub optAlignRight_Click()
    On Error GoTo EH
    If mbLoading Then
        Exit Sub
    End If
    SaveSetting App.EXEName, "GENERAL", "NAV_SCREEN_POS", "RIGHT"
    gfrmECTray.PosECTRAY
    gfrmECTray.ShowMe , False
    Me.SetFocus
    optAlignRight.SetFocus
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub optAlignRight_Click"
End Sub

Private Sub optPrintPreviewActive_Click()
    On Error GoTo EH
    If mbLoading Then
        Exit Sub
    End If
    SaveSetting App.EXEName, "GENERAL", "PRINT_PREVIEW", "USE_ACTIVE"
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub optPrintPreviewActive_Click"
End Sub

Private Sub optPrintPreviewAdobe_Click()
    On Error GoTo EH
    If mbLoading Then
        Exit Sub
    End If
    SaveSetting App.EXEName, "GENERAL", "PRINT_PREVIEW", "USE_ADOBE"
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub optPrintPreviewAdobe_Click"
End Sub

Private Sub optShowListTXT_Click()
    On Error GoTo EH
    If mbLoading Then
        Exit Sub
    End If
    SaveSetting App.EXEName, "GENERAL", "SHOWLIST_FORMAT", ".txt"
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub optShowListTXT_Click"
End Sub

Private Sub optShowListXLS_Click()
    On Error GoTo EH
    If mbLoading Then
        Exit Sub
    End If
    SaveSetting App.EXEName, "GENERAL", "SHOWLIST_FORMAT", ".xls"
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub optShowListXLS_Click"
End Sub

Private Sub txtADJAddress_GotFocus()
    goUtil.utSelText txtADJAddress
End Sub

Private Sub txtADJAddress_LostFocus()
    txtADJAddress.Text = UCase(Trim(txtADJAddress.Text))
End Sub

Private Sub txtADJCity_GotFocus()
    goUtil.utSelText txtADJCity
End Sub

Private Sub txtADJCity_LostFocus()
    txtADJCity.Text = UCase(Trim(txtADJCity.Text))
End Sub

Private Sub txtADJContactPhone_GotFocus()
    goUtil.utSelText txtADJContactPhone
End Sub

Private Sub txtADJContactPhone_LostFocus()
    txtADJContactPhone.Text = Trim(txtADJContactPhone.Text)
End Sub

Private Sub txtADJEmail_GotFocus()
    goUtil.utSelText txtADJEmail
End Sub

Private Sub txtADJEmail_LostFocus()
    txtADJEmail.Text = Trim(txtADJEmail.Text)
End Sub

Private Sub txtADJEmergencyPhone_GotFocus()
    goUtil.utSelText txtADJEmergencyPhone
End Sub

Private Sub txtADJEmergencyPhone_LostFocus()
    txtADJEmergencyPhone.Text = Trim(txtADJEmergencyPhone.Text)
End Sub

Private Sub txtADJOtherPostCode_GotFocus()
    goUtil.utSelText txtADJOtherPostCode
End Sub

Private Sub txtADJOtherPostCode_LostFocus()
    txtADJOtherPostCode.Text = UCase(Trim(txtADJOtherPostCode.Text))
End Sub

Private Sub txtADJState_GotFocus()
    goUtil.utSelText txtADJState
End Sub

Private Sub txtADJState_LostFocus()
    txtADJState.Text = UCase(Trim(txtADJState.Text))
End Sub

Private Sub txtADJZip_GotFocus()
    goUtil.utSelText txtADJZip
End Sub

Private Sub txtADJZip_LostFocus()
    txtADJZip.Text = Trim(txtADJZip.Text)
End Sub

Private Sub txtADJZip4_GotFocus()
    goUtil.utSelText txtADJZip4
End Sub

Private Sub txtADJZip4_LostFocus()
    txtADJZip4.Text = Trim(txtADJZip4.Text)
End Sub

Private Sub txtLogonHintAnswer_GotFocus()
    goUtil.utSelText txtLogonHintAnswer
End Sub

Private Sub txtLogonHintQuestion_GotFocus()
    goUtil.utSelText txtLogonHintQuestion
End Sub

Private Sub txtPass_Click()
    cmdEnterPass_Click
End Sub

Private Sub txtUserName_GotFocus()
    goUtil.utSelText txtUserName
End Sub

Private Sub txtUserName_LostFocus()
    txtUserName.Text = UCase(Trim(txtUserName.Text))
End Sub

Private Sub txtFName_GotFocus()
    goUtil.utSelText txtFName
End Sub

Private Sub txtFName_LostFocus()
    txtFName.Text = UCase(Trim(txtFName.Text))
End Sub

Private Sub txtLName_GotFocus()
    goUtil.utSelText txtLName
End Sub

Private Sub txtLName_LostFocus()
    txtLName.Text = UCase(Trim(txtLName.Text))
End Sub

Private Sub txtPass_GotFocus()
    goUtil.utSelText txtPass
End Sub

Private Sub txtSSN_GotFocus()
    goUtil.utSelText txtSSN
End Sub

Private Sub txtTeamLeaderSup_GotFocus()
    goUtil.utSelText txtTeamLeaderSup
End Sub

Private Sub txtTeamLeaderSup_LostFocus()
    txtTeamLeaderSup.Text = UCase(Trim(txtTeamLeaderSup.Text))
End Sub

Private Sub txtWallPaperImage_GotFocus()
    goUtil.utSelText txtWallPaperImage
End Sub

Private Sub txtWebHost_GotFocus()
    goUtil.utSelText txtWebHost
End Sub

Private Sub cmdEnterPass_Click()
    On Error GoTo EH
    Dim sMess As String
    If GetSetting("ECS", "WEB_SECURITY", "RESET_PASSWORD", False) Then
        sMess = "Previous Password changes must be uploaded first." & vbCrLf
        sMess = sMess & "You may then change your password again."
        MsgBox sMess, vbInformation, "Can't Change Password"
    Else
        EnterPass
    End If

    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdEnterPass_Click"
End Sub

Public Sub EnterPass()
    goUtil.utEnterUserPass "WEB_SECURITY", txtSSN.Text, txtPass, Me
End Sub

'Private Sub cmdSyncUserName_Click()
'    SynchronizeUserName txtUserName.Text
'End Sub

Private Sub cmdOK_Click()
    On Error GoTo EH
    Dim bCancel As Boolean
    Dim sSSN As String
    
    'Trim again incase used ALT+OK
    'Adjuster Information
    txtFName.Text = UCase(Trim(txtFName.Text))
    txtLName.Text = UCase(Trim(txtLName.Text))
    txtADJEmail.Text = Trim(txtADJEmail.Text)
    txtADJContactPhone.Text = Trim(txtADJContactPhone.Text)
    txtADJEmergencyPhone.Text = Trim(txtADJEmergencyPhone)
    txtADJAddress.Text = UCase(Trim(txtADJAddress.Text))
    txtADJCity.Text = UCase(Trim(txtADJCity.Text))
    txtADJState.Text = UCase(Trim(txtADJState.Text))
    txtADJZip.Text = Trim(txtADJZip.Text)
    txtADJZip4.Text = Trim(txtADJZip4.Text)
    txtADJOtherPostCode.Text = UCase(Trim(txtADJOtherPostCode.Text))
    txtTeamLeaderSup.Text = UCase(Trim(txtTeamLeaderSup.Text))
     
     'Internet Security
     txtWebHost.Text = UCase(Trim(txtWebHost.Text))
     txtUserName.Text = UCase(Trim(txtUserName.Text))
     'Password set by cmdEnterPass_Click
     
     
    
    If txtFName.Text = vbNullString Then
        bCancel = True
    Else
         SaveSetting App.EXEName, "GENERAL", "ADJUSTOR_FIRST_NAME", txtFName.Text
    End If
    
    If txtLName.Text = vbNullString Then
        bCancel = True
    Else
         SaveSetting App.EXEName, "GENERAL", "ADJUSTOR_LAST_NAME", txtLName.Text
    End If

    If txtADJEmail.Text = vbNullString Then
        bCancel = True
    Else
          SaveSetting App.EXEName, "GENERAL", "ADJ_EMAIL", txtADJEmail.Text
    End If
    
    If txtADJContactPhone.Text = vbNullString Then
        bCancel = True
    Else
          SaveSetting App.EXEName, "GENERAL", "ADJ_CONTACT_PHONE", txtADJContactPhone.Text
    End If
    
    If txtADJEmergencyPhone.Text = vbNullString Then
        bCancel = True
    Else
          SaveSetting App.EXEName, "GENERAL", "ADJ_EMERGENCY_PHONE", txtADJEmergencyPhone.Text
    End If
    
    If txtADJAddress.Text = vbNullString Then
        bCancel = True
    Else
        SaveSetting App.EXEName, "GENERAL", "ADJ_ADDRESS", txtADJAddress.Text
    End If
    
    If txtADJCity.Text = vbNullString Then
        bCancel = True
    Else
        SaveSetting App.EXEName, "GENERAL", "ADJ_CITY", txtADJCity.Text
    End If
    
    If txtADJState.Text = vbNullString Then
        bCancel = True
    Else
        SaveSetting App.EXEName, "GENERAL", "ADJ_STATE", txtADJState.Text
    End If
    
    If txtADJZip.Text = vbNullString Then
         bCancel = True
    ElseIf Not IsNumeric(txtADJZip.Text) Then
        bCancel = True
    Else
        SaveSetting App.EXEName, "GENERAL", "ADJ_ZIP", txtADJZip.Text
    End If
    
    If txtADJZip4.Text <> vbNullString Then
        If Not IsNumeric(txtADJZip4.Text) Then
            bCancel = True
        Else
            SaveSetting App.EXEName, "GENERAL", "ADJ_ZIP4", txtADJZip4.Text
        End If
    End If
    
    SaveSetting App.EXEName, "GENERAL", "ADJ_OTHER_POSTCODE", txtADJOtherPostCode.Text
    SaveSetting App.EXEName, "GENERAL", "TEAM_LEADER", txtTeamLeaderSup.Text
    
    If txtWebHost.Text = vbNullString Then
        bCancel = True
    Else
          SaveSetting "ECS", "WEB_SECURITY", "WEB_HOST", txtWebHost.Text
    End If
    
    If txtUserName.Text = vbNullString Then
        bCancel = True
    Else
         goUtil.utSaveECSCryptSetting "ECS", "WEB_SECURITY", "USER_NAME", txtUserName.Text
    End If
    
    If txtPass.Text = vbNullString Then
        bCancel = True
    Else
        'Password set by cmdEnterPass_Click
    End If
    
    If txtLogonHintQuestion.Text = vbNullString Then
        bCancel = True
    Else
         goUtil.utSaveECSCryptSetting "ECS", "WEB_SECURITY", "LOGON_HINT_Q", txtLogonHintQuestion.Text
    End If
    
    If txtLogonHintAnswer.Text = vbNullString Then
        bCancel = True
    Else
         goUtil.utSaveECSCryptSetting "ECS", "WEB_SECURITY", "LOGON_HINT_A", txtLogonHintAnswer.Text
    End If
    
    SaveSetting App.EXEName, "GENERAL", "TEAM_LEADER", txtTeamLeaderSup.Text
    
    'BGS 11.20.2001 make sure that they have entered the SSN correctly
    'the ssn should be numeric and 9 digits in length
    'BGS 3.1.2002 153  EasyClaim will not accept a SSN beginning with 0
    'Need better validation so its there now
    If Not IsNumeric(txtSSN.Text) And Not bCancel Then
        txtSSN.Text = Val(txtSSN.Text)
    Else
        sSSN = Val(txtSSN.Text)
        If Len(sSSN) = 1 Then
            txtSSN.Text = "Error"
        Else
            sSSN = txtSSN.Text
            If InStr(1, sSSN, "+") > 0 Or InStr(1, sSSN, "-") > 0 Or InStr(1, sSSN, "e", vbTextCompare) > 0 Then
                txtSSN.Text = "Error"
            End If
        End If
    End If

    If Len(txtSSN.Text) = txtSSN.MaxLength Then
        If txtSSN.Enabled Then
            goUtil.utSaveECSCryptSetting "ECS", "WEB_SECURITY", "SSN", Trim(txtSSN.Text)
        Else
            If txtSSN.Enabled Then
                bCancel = True
            End If
        End If
    Else
        MsgBox "Please enter your SSN # without ""-"" and no spaces." & vbCrLf & "This number will only be used for billing purposes and will be kept private!", vbExclamation + vbOKOnly, "SSN"
        bCancel = True
    End If
    
EXIT_HERE:
    If Not bCancel Then
        Unload Me
    Else
        MsgBox "ALL fields must be filled out correctly!", vbExclamation + vbOKOnly, Me.Caption
    End If
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdOK_Click"
End Sub

Private Sub Form_Load()
    On Error GoTo EH
    mbLoading = True
    goUtil.utFormWinRegPos goUtil.gsMainAppEXEName, Me

    LoadPref
    
    mbLoading = False
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_Load"
End Sub


Public Sub LoadPref()
    On Error GoTo EH
    
    'Adjuster Information
    txtFName.Text = GetSetting(App.EXEName, "GENERAL", "ADJUSTOR_FIRST_NAME", vbNullString)
    txtLName.Text = GetSetting(App.EXEName, "GENERAL", "ADJUSTOR_LAST_NAME", vbNullString)
    txtSSN.Text = goUtil.utGetECSCryptSetting("ECS", "WEB_SECURITY", "SSN")
    txtADJEmail.Text = GetSetting(App.EXEName, "GENERAL", "ADJ_EMAIL", vbNullString)
    txtADJContactPhone.Text = GetSetting(App.EXEName, "GENERAL", "ADJ_CONTACT_PHONE", vbNullString)
    txtADJEmergencyPhone.Text = GetSetting(App.EXEName, "GENERAL", "ADJ_EMERGENCY_PHONE", vbNullString)
    txtADJAddress.Text = GetSetting(App.EXEName, "GENERAL", "ADJ_ADDRESS", vbNullString)
    txtADJCity.Text = GetSetting(App.EXEName, "GENERAL", "ADJ_CITY", vbNullString)
    txtADJState.Text = GetSetting(App.EXEName, "GENERAL", "ADJ_STATE", vbNullString)
    txtADJZip.Text = GetSetting(App.EXEName, "GENERAL", "ADJ_ZIP", vbNullString)
    txtADJZip4.Text = GetSetting(App.EXEName, "GENERAL", "ADJ_ZIP4", vbNullString)
    txtADJOtherPostCode.Text = GetSetting(App.EXEName, "GENERAL", "ADJ_OTHER_POSTCODE", vbNullString)
    txtTeamLeaderSup.Text = GetSetting(App.EXEName, "GENERAL", "TEAM_LEADER", vbNullString)
    
    'Internet Security
    txtWebHost.Text = GetSetting("ECS", "WEB_SECURITY", "WEB_HOST", "WWW.EBERLS.NET")
    txtUserName.Text = goUtil.utGetECSCryptSetting("ECS", "WEB_SECURITY", "USER_NAME", vbNullString)
    txtPass.Text = goUtil.utGetECSCryptSetting("ECS", "WEB_SECURITY", "PASSWORD")
    txtLogonHintQuestion.Text = goUtil.utGetECSCryptSetting("ECS", "WEB_SECURITY", "LOGON_HINT_Q", "Mother's Maiden Name?")
    txtLogonHintAnswer.Text = goUtil.utGetECSCryptSetting("ECS", "WEB_SECURITY", "LOGON_HINT_A", vbNullString)
    

    'BGS 10.10.2002 Put trail "..." on Lables
    goUtil.utSuffixLabels lblGlobalPref, 50
    
    
    'Load NavPref
    LoadNavPref
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub LoadPref"
End Sub

Public Sub LoadNavPref()
    On Error GoTo EH
    Dim sWallPicImage As String
    Dim sNavScreenPos As String
    Dim sPrintPreview As String
    Dim sShowListFormat As String
    Dim bUseSSL As Boolean
    Dim bUseWallPaper As Boolean
    Dim bAlwaysOnTop As Boolean
    Dim bUseSilentError As Boolean
    
    'Check For SSL Option
    bUseSSL = CBool(GetSetting("ECS", "WEB_SECURITY", "USE_SSL", True))
    If bUseSSL Then
        chkUseSSL.Value = vbChecked
    Else
        chkUseSSL.Value = vbUnchecked
    End If
    
    'CHeck For Wall Paper
    bUseWallPaper = CBool(GetSetting(App.EXEName, "GENERAL", "USE_WALL_PAPER", True))
    If bUseWallPaper Then
        chkUseWallPaper.Value = vbChecked
    Else
        chkUseWallPaper.Value = vbUnchecked
    End If
    
    bAlwaysOnTop = CBool(GetSetting(App.EXEName, "GENERAL", "ECTRAY_ALWAYS_ON_TOP", False))
    If bAlwaysOnTop Then
        chkAlwaysOnTop.Value = vbChecked
    Else
        chkAlwaysOnTop.Value = vbUnchecked
    End If
    
    sNavScreenPos = GetSetting(App.EXEName, "GENERAL", "NAV_SCREEN_POS", "RIGHT")
    
    Select Case UCase(sNavScreenPos)
        Case "RIGHT"
            optAlignRight.Value = True
            optAlignLeft.Value = False
        Case "LEFT"
            optAlignRight.Value = False
            optAlignLeft.Value = True
    End Select
    
    'Set the options for print Preview
    sPrintPreview = GetSetting(App.EXEName, "GENERAL", "PRINT_PREVIEW", "USE_ADOBE")
    
    Select Case UCase(sPrintPreview)
        Case "USE_ADOBE"
            optPrintPreviewAdobe.Value = True
            optPrintPreviewActive.Value = False
        Case "USE_ACTIVE"
            optPrintPreviewAdobe.Value = False
            optPrintPreviewActive.Value = True
    End Select
    
    'Set the options for Show List Format
    sShowListFormat = GetSetting(App.EXEName, "GENERAL", "SHOWLIST_FORMAT", ".xls")
    Select Case UCase(sShowListFormat)
        Case ".XLS"
            optShowListXLS.Value = True
            optShowListTXT.Value = False
        Case ".TXT"
            optShowListTXT.Value = True
            optShowListXLS.Value = False
    End Select
    
    'Check For Silent Error Option
    bUseSilentError = CBool(GetSetting("ECS", "Dir", "SILENT_ERROR", True))
    If bUseSilentError Then
        chkUseSilentError.Value = vbChecked
    Else
        chkUseSilentError.Value = vbUnchecked
    End If
    
    
    sWallPicImage = GetSetting(App.EXEName, "DIR", "WALL_PAPER_IMAGE", "(Blank Wall Paper)")
    If InStr(1, sWallPicImage, "\", vbBinaryCompare) > 0 Then
        sWallPicImage = Mid(sWallPicImage, InStrRev(sWallPicImage, "\") + 1)
    End If
    txtWallPaperImage.Text = sWallPicImage
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub LoadNavPref"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo EH
    goUtil.utFormWinRegPos goUtil.gsMainAppEXEName, Me, True
    gfrmECTray.ShowMe False
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_Unload"
End Sub

'Private Function SynchronizeUserName(psUserName As String) As Boolean
'    On Error GoTo EH
'    Dim sSQL As String
'    Dim bdoSync As Boolean
'    Dim sMess As String
'    Dim lRecAff As Long
'
'    If Trim(psUserName) = vbNullString Then
'        MsgBox "Claim Rep ID (UserName) must be filled out first!", vbExclamation, lblGlobalPref(8).Caption
'        Exit Function
'    End If
'
'    sMess = "Are you sure you want to proceed ? " & vbCrLf & vbCrLf
'    sMess = sMess & "All UserName for Current CAT " & Mid(goUtil.gsCurCatDir, InStrRev(goUtil.gsCurCatDir, "\") + 1) & vbCrLf
'    sMess = sMess & "will be replaced with UserName--> " & psUserName
'
'    If MsgBox(sMess, vbQuestion + vbYesNo, lblGlobalPref(8).Caption) = vbNo Then
'        Exit Function
'    End If
'
'    'If they answer yes to the warning then synchronize the UserName
'
'    '1. First update the UserName...
'    sSQL = "UPDATE Assignments SET "
'    sSQL = sSQL & "Assignments.ClaimRepIDNO = '" & goUtil.utCleanSQLString(psUserName) & "' "
'    sSQL = sSQL & "WHERE Assignments.ClaimRepIDNO <> '" & goUtil.utCleanSQLString(psUserName) & "' "
'
'    goUtil.gCurDB.Execute sSQL
'    lRecAff = goUtil.gCurDB.RecordsAffected
'
'    '2. Then
'    '10.23.2002 added INNER join To cheks to only mark the ones that Have
'    'a check entry to be Real time Uploaded.
'    sSQL = "UPDATE Assignments INNER JOIN Checks "
'    sSQL = sSQL & "ON Assignments.ClaimNo = Checks.ClaimNo "
'    sSQL = sSQL & "SET "
'    sSQL = sSQL & "Assignments.RTUpLoadMe = True "
'
'    goUtil.gCurDB.Execute sSQL
'
'    MsgBox lRecAff & " Claim(s) were affected.", vbInformation, lblGlobalPref(8).Caption
'
'    Exit Function
'EH:
'    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Function SynchronizeUserName"
'End Function

Private Sub txtWebHost_LostFocus()
    txtWebHost.Text = UCase(Trim(txtWebHost.Text))
End Sub
