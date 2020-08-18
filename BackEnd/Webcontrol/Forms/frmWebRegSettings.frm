VERSION 5.00
Object = "{B0AA617A-8DB4-4E9E-BBC1-CA4E3B6280AA}#2.0#0"; "ecsTimeOCX.ocx"
Begin VB.Form frmWebRegSettings 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Webcontrol Service Settings"
   ClientHeight    =   9255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10455
   Icon            =   "frmWebRegSettings.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   9255
   ScaleWidth      =   10455
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FramSMTP 
      Appearance      =   0  'Flat
      Caption         =   "SMTP (Send Email Account Settings)"
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
      Height          =   2655
      Left            =   6000
      TabIndex        =   26
      Top             =   3720
      Width           =   4335
      Begin VB.Frame FramSMTPHost 
         Caption         =   "SMTP Host"
         Height          =   735
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   4095
         Begin VB.TextBox txtSMTPHost 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   120
            TabIndex        =   28
            Top             =   240
            Width           =   3855
         End
      End
      Begin VB.Frame FramSMTPUserName 
         Caption         =   "User Name"
         Height          =   735
         Left            =   120
         TabIndex        =   29
         Top             =   960
         Width           =   4095
         Begin VB.TextBox txtSMTPUserName 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   120
            TabIndex        =   30
            Top             =   240
            Width           =   3855
         End
      End
      Begin VB.Frame FramSMTPPassword 
         Caption         =   "Password"
         Height          =   735
         Left            =   120
         TabIndex        =   31
         Top             =   1680
         Width           =   4095
         Begin VB.TextBox txtSMTPPassword 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            IMEMode         =   3  'DISABLE
            Left            =   120
            PasswordChar    =   "*"
            TabIndex        =   32
            Top             =   240
            Width           =   3855
         End
      End
   End
   Begin VB.Frame framUSER 
      Appearance      =   0  'Flat
      Caption         =   "Data Base Login"
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
      Height          =   2175
      Left            =   6000
      TabIndex        =   33
      Top             =   6480
      Width           =   4335
      Begin VB.Frame framPassword 
         Caption         =   "Password"
         Height          =   735
         Left            =   120
         TabIndex        =   36
         Top             =   1080
         Width           =   4095
         Begin VB.TextBox txtPassWord 
            Height          =   360
            IMEMode         =   3  'DISABLE
            Left            =   120
            PasswordChar    =   "*"
            TabIndex        =   37
            Top             =   240
            Width           =   3855
         End
      End
      Begin VB.Frame framUserID 
         Caption         =   "User ID"
         Height          =   735
         Left            =   120
         TabIndex        =   34
         Top             =   240
         Width           =   4095
         Begin VB.TextBox txtUserID 
            Height          =   360
            Left            =   120
            TabIndex        =   35
            Top             =   240
            Width           =   3855
         End
         Begin VB.CommandButton Command1 
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
            Left            =   1440
            Picture         =   "frmWebRegSettings.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   39
            Top             =   255
            Width           =   375
         End
      End
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8400
      TabIndex        =   38
      Top             =   8760
      Width           =   1920
   End
   Begin VB.Frame framWEBCONTROL 
      Appearance      =   0  'Flat
      Caption         =   "WebControl"
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
      Height          =   4935
      Left            =   120
      TabIndex        =   15
      Top             =   3720
      Width           =   5775
      Begin VB.Frame framSelectDSN 
         Caption         =   "Select DSN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   120
         TabIndex        =   24
         Top             =   2160
         Visible         =   0   'False
         Width           =   5535
         Begin VB.ListBox lstDSN 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2160
            ItemData        =   "frmWebRegSettings.frx":058C
            Left            =   120
            List            =   "frmWebRegSettings.frx":058E
            TabIndex        =   25
            Top             =   240
            Width           =   5295
         End
      End
      Begin VB.Frame framDBName 
         Caption         =   "Production Data Base (DSN)"
         Height          =   735
         Left            =   120
         TabIndex        =   21
         Top             =   1560
         Width           =   5535
         Begin VB.CommandButton cmdDBName 
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
            Left            =   5040
            Picture         =   "frmWebRegSettings.frx":0590
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Browse DSN"
            Top             =   255
            Width           =   375
         End
         Begin VB.TextBox txtDBName 
            Height          =   360
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   22
            Top             =   240
            Width           =   5295
         End
      End
      Begin VB.Frame framWebControlServer 
         Caption         =   "WebControl Server Share"
         Height          =   735
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   5535
         Begin VB.CommandButton cmdWebControlServer 
            Height          =   330
            Left            =   5040
            Picture         =   "frmWebRegSettings.frx":06DA
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Browse"
            Top             =   255
            Width           =   375
         End
         Begin VB.TextBox txtWebControlServer 
            Height          =   360
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   17
            Top             =   240
            Width           =   5295
         End
      End
      Begin VB.Frame framLossReportPrinter 
         Caption         =   "Loss Report Default Printer"
         Height          =   735
         Left            =   120
         TabIndex        =   19
         Top             =   960
         Width           =   5535
         Begin VB.ComboBox cboSelectPrinter 
            Height          =   315
            ItemData        =   "frmWebRegSettings.frx":0B54
            Left            =   120
            List            =   "frmWebRegSettings.frx":0B56
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   240
            Width           =   5295
         End
      End
   End
   Begin VB.Frame framAutoImport 
      Appearance      =   0  'Flat
      Caption         =   "AutoImport"
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
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10215
      Begin VB.Frame framFTPSite 
         Caption         =   "FTP Site (Directory)"
         Height          =   735
         Left            =   120
         TabIndex        =   12
         Top             =   2520
         Width           =   9975
         Begin VB.CommandButton cmdFTPSitePath 
            Height          =   330
            Left            =   9480
            Picture         =   "frmWebRegSettings.frx":0B58
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Browse"
            Top             =   255
            Width           =   375
         End
         Begin VB.TextBox txtFTPSitePath 
            Height          =   360
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   13
            Top             =   240
            Width           =   9735
         End
      End
      Begin VB.Frame framFLS 
         Caption         =   "FLS (File Life-Span Dat Directory)"
         Height          =   735
         Left            =   120
         TabIndex        =   6
         Top             =   980
         Width           =   9975
         Begin VB.CommandButton cmdFLSPath 
            Height          =   330
            Left            =   9480
            Picture         =   "frmWebRegSettings.frx":0FD2
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Browse"
            Top             =   255
            Width           =   375
         End
         Begin VB.TextBox txtFLSPath 
            Height          =   360
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   240
            Width           =   9735
         End
      End
      Begin VB.Frame framWebSite 
         Caption         =   "Web Site (Directory)"
         Height          =   735
         Left            =   120
         TabIndex        =   9
         Top             =   1695
         Width           =   9975
         Begin VB.CommandButton cmdWebSitePath 
            Height          =   330
            Left            =   9480
            Picture         =   "frmWebRegSettings.frx":144C
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Browse"
            Top             =   255
            Width           =   375
         End
         Begin VB.TextBox txtWebSitePath 
            Height          =   360
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   10
            Top             =   240
            Width           =   9735
         End
      End
      Begin VB.Frame framInterval 
         Caption         =   "Update Interval"
         Height          =   735
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1935
         Begin VB.CommandButton cmdInterval 
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
            Left            =   1440
            Picture         =   "frmWebRegSettings.frx":18C6
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   255
            Width           =   375
         End
         Begin VB.TextBox txtInterval 
            Height          =   360
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   2
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Frame framFLSDelTime 
         Caption         =   "FLS (Delete FLS Files Time)"
         Height          =   735
         Left            =   6600
         TabIndex        =   4
         Top             =   240
         Width           =   3495
         Begin ecsTimeOCX.ecsTime ecsTimeFLSDel 
            Height          =   375
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Text            =   "11:43 AM"
            Appearance      =   1
            Object.TabStop         =   -1  'True
         End
      End
   End
End
Attribute VB_Name = "frmWebRegSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbLoading As Boolean
Private mbResetProcess As Boolean
Private moDSNText As Object

Private Sub cboSelectPrinter_Click()
    On Error GoTo EH
    If Not mbLoading Then
        SaveSetting "V2WebControl", "Printer", "LossReport", cboSelectPrinter.Text
    End If
    Exit Sub
EH:
    ShowError Err, "Private Sub cboSelectPrinter_Click", Me
End Sub

'Private Sub cmdClaimsBillingDBName_Click()
'    LoadDSNNames txtClaimsBillingDBName
'End Sub

Private Sub cmdDBName_Click()
    On Error GoTo EH
    If framSelectDSN.Visible Then
        framSelectDSN.Visible = False
    Else
        LoadDSNNames txtDBName, "ACCESS_2000"
    End If
    Exit Sub
EH:
    ShowError Err, "Private Sub cmdFLSPath_Click", Me
End Sub

Private Sub cmdExit_Click()
    On Error GoTo EH
    If mbResetProcess Then
        ResetProcess
        mbResetProcess = False
    End If
    Unload Me
    Exit Sub
EH:
    ShowError Err, "Private Sub cmdExit_Click", Me
End Sub

Private Sub cmdFLSPath_Click()
    On Error GoTo EH
    txtFLSPath.Text = GetPath("FLSdatPath", "Browse to File Life-Span Dat Directory", "CLICK OPEN TO SAVE PATH", txtFLSPath.Text, Me.hWnd)
    Exit Sub
EH:
    ShowError Err, "Private Sub cmdFLSPath_Click", Me
End Sub

Private Sub cmdFTPSitePath_Click()
    On Error GoTo EH
    txtFTPSitePath.Text = GetPath("FTPSitePath", "Browse to FTP Site Directory", "CLICK OPEN TO SAVE PATH", txtFTPSitePath.Text, Me.hWnd)
    Exit Sub
EH:
    ShowError Err, "Private Sub cmdFTPSitePath_Click", Me
End Sub

Private Sub cmdInterval_Click()
    On Error GoTo EH
    txtInterval.Text = GetInterval
    Exit Sub
EH:
    ShowError Err, "Private Sub cmdInterval_Click", Me
End Sub

'Private Sub cmdProdUploadDBName_Click()
'    LoadDSNNames txtProdUploadDBName
'End Sub

Private Sub cmdWebControlServer_Click()
    On Error GoTo EH
    txtWebControlServer.Text = GetPath("V2WebControl_SERVER_SHARE", "Browse to the directory above V2WebControl Folder", "CLICK OPEN TO SAVE PATH", txtWebControlServer.Text, Me.hWnd)
    Exit Sub
EH:
    ShowError Err, "Private Sub cmdWebControlServer_Click", Me
End Sub

Private Sub cmdWebSitePath_Click()
    On Error GoTo EH
    txtWebSitePath.Text = GetPath("WebSitePath", "Browse to Web Site Directory", "CLICK OPEN TO SAVE PATH", txtWebSitePath.Text, Me.hWnd)
    Exit Sub
EH:
    ShowError Err, "Private Sub cmdWebSitePath_Click", Me
End Sub

'Private Sub cmkDisableClaimsBilling_Click()
'    On Error GoTo EH
'    Dim bUpdate As Boolean
'
'    If cmkDisableClaimsBilling.Value = vbChecked Then
'        bUpdate = True
'    Else
'        bUpdate = False
'    End If
'
'    SaveSetting "V2WebControl", "DBConn", "DisableClaimsUpdate", bUpdate
'
'    Exit Sub
'EH:
'    ShowError Err, "Private Sub cmkDisableClaimsBilling_Click", Me
'End Sub

Private Sub ecsTimeFLSDel_time24HR(ps24HR As String)
    On Error GoTo EH
    
    SaveSetting "V2AutoImport", "Msg", "FLSDelTime", ps24HR
    
    Exit Sub
EH:
    ShowError Err, "Private Sub ecsTimeFLSDel_time24HR", Me
End Sub


Private Sub Form_Load()
    On Error GoTo EH
    mbLoading = True
    FormWinRegPos Me
    
    LoadDefaultSettings
    mbResetProcess = False
    
    mbLoading = False
    Exit Sub
EH:
    ShowError Err, "Private Sub Form_Load", Me
End Sub

Private Sub LoadDefaultSettings()
    On Error GoTo EH
    '1. V2AutoImport
    txtInterval.Text = GetInterval(True)
    ecsTimeFLSDel.ecsTime = GetSetting("V2WebControl", "Msg", "FLSDelTime", "12:15 AM")
    txtFLSPath.Text = GetSetting("V2WebControl", "Dir", "FLSdatPath", vbNullString)
    txtWebSitePath.Text = GetSetting("V2WebControl", "Dir", "WebSitePath", vbNullString)
    txtFTPSitePath.Text = GetSetting("V2WebControl", "Dir", "FTPSitePath", vbNullString)
    '2 V2WebControl
    txtWebControlServer.Text = GetSetting("V2WebControl", "Dir", "V2WebControl_SERVER_SHARE", vbNullString)
    LoadPrinters
    txtDBName.Text = GetSetting("V2WebControl", "DSN", "NAME", "ACCESS_2000")
'    txtProdUploadDBName.Text = GetSetting("V2WebControl", "DBConn", "Approach", vbNullString)
'    txtClaimsBillingDBName.Text = GetSetting("V2WebControl", "DBConn", "Claims", vbNullString)
    txtUserID.Text = GetECSCryptSetting("V2WebControl", "DBConn", "USERID")
    txtPassWord.Text = GetECSCryptSetting("V2WebControl", "DBConn", "PASSWORD")
    'SMTP
    txtSMTPHost.Text = GetECSCryptSetting("V2WebControl", "SMTP", "Host")
    txtSMTPUserName.Text = GetECSCryptSetting("V2WebControl", "SMTP", "UserName")
    txtSMTPPassword.Text = GetECSCryptSetting("V2WebControl", "SMTP", "Password")
    
'    If CBool(GetSetting("V2WebControl", "DBConn", "DisableClaimsUpdate", True)) Then
'        cmkDisableClaimsBilling.Value = vbChecked
'    Else
'        cmkDisableClaimsBilling.Value = vbUnchecked
'    End If
    
    Exit Sub
EH:
    ShowError Err, "Private Sub LoadDefaultSettings", Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    FormWinRegPos Me, True
    'Clean up
    Unload moDSNText
    Set moDSNText = Nothing
    
End Sub

Public Function GetInterval(Optional pbDefaultOnly As Boolean) As String
    On Error GoTo EH
    Dim sDefault As String
    Dim sMess As String
    Dim sRet As String
    
    sDefault = GetSetting("V2WebControl", "Msg", "UpdateHourly", 30000)
    If Val(sDefault) >= 1000 Then
        sDefault = Val(sDefault) / 1000
    End If
    
    If pbDefaultOnly Then
        GetInterval = sDefault
        Exit Function
    Else
        sMess = "Please enter 10 to 60 seconds." & vbCrLf & vbCrLf
        sMess = sMess & "Or..." & vbCrLf & vbCrLf
        sMess = sMess & "Type the word 'Hourly' for an Hourly Interval."
        sRet = InputBox(sMess, "Update Interval", sDefault, Me.left, Me.top)
        
        If sRet = vbNullString Then
            sRet = sDefault
        End If
        GetInterval = sRet
        
        If InStr(1, sRet, "Hourly", vbTextCompare) > 0 Then
            sRet = "Hourly"
        Else
            sRet = Val(sRet) * 1000
        End If
    End If
    
    SaveSetting "V2WebControl", "Msg", "UpdateHourly", sRet
    
    Exit Function
EH:
    ShowError Err, "Public Function GetInterval", Me
End Function


Private Sub lstDSN_Click()
    On Error GoTo EH
    
    lstDSN.ToolTipText = lstDSN.Text
    
    Exit Sub
EH:
    ShowError Err, "Private Sub lstDSN_Click", Me
End Sub

Private Sub lstDSN_DblClick()
    On Error GoTo EH
    Dim sDSN As String
    
    sDSN = lstDSN.Text
    If InStr(1, sDSN, Chr(160), vbBinaryCompare) > 0 Then
        sDSN = left(sDSN, InStrRev(sDSN, Chr(160)) - 1)
    End If
    moDSNText.Text = sDSN
    framSelectDSN.Visible = False
    framSelectDSN.Refresh
    ResetProcess
    mbResetProcess = False
    Exit Sub
EH:
    Screen.MousePointer = vbNormal
    ShowError Err, "Private Sub lstDSN_DblClick", Me
End Sub

Private Sub Text1_Change()

End Sub

'Private Sub txtClaimsBillingDBName_Change()
'    On Error GoTo EH
'
'    SaveSetting "V2WebControl", "DBConn", "Claims", txtClaimsBillingDBName.Text
'
'    Exit Sub
'EH:
'    ShowError Err, "Private Sub txtClaimsBillingDBName_Change", Me
'End Sub

'Private Sub txtClaimsBillingDBName_GotFocus()
'    SelText txtClaimsBillingDBName
'End Sub

Private Sub txtDBName_Change()
    On Error GoTo EH
    
    SaveSetting "V2WebControl", "DSN", "NAME", txtDBName.Text
    
    Exit Sub
EH:
    ShowError Err, "Private Sub txtDBName_Change", Me
End Sub

Private Sub txtDBName_GotFocus()
    SelText txtDBName
End Sub

Private Sub txtFLSPath_Change()
    On Error GoTo EH
    SaveSetting "V2WebControl", "Dir", "FLSdatPath", txtFLSPath.Text
    Exit Sub
EH:
    ShowError Err, "Private Sub txtFLSPath_Change", Me
End Sub

Private Sub txtFLSPath_GotFocus()
    SelText txtFLSPath
End Sub

Private Sub txtFTPSitePath_Change()
    On Error GoTo EH
    SaveSetting "V2WebControl", "Dir", "FTPSitePath", txtFTPSitePath.Text
    Exit Sub
EH:
    ShowError Err, "Private Sub txtFTPSitePath_Change", Me
End Sub

Private Sub txtFTPSitePath_GotFocus()
    SelText txtFTPSitePath
End Sub

Private Sub txtInterval_GotFocus()
    SelText txtInterval
End Sub

Private Sub LoadPrinters()
    On Error GoTo EH
    Dim prn As Printer
    Dim sLastSelectedPrinter As String
    
    'Fill the Printers Combo Box
    cboSelectPrinter.Clear
    For Each prn In Printers
        cboSelectPrinter.AddItem prn.DeviceName & " on " & prn.Port
    Next prn
    
    sLastSelectedPrinter = GetSetting("V2WebControl", "Printer", "LossReport", vbNullString)
    'Select Deafult printer from list
    SelectDefaultPrinter cboSelectPrinter, sLastSelectedPrinter
    
    Exit Sub
EH:
    ShowError Err, "Private Sub LoadPrinters", Me
End Sub

Private Sub txtPassWord_Change()
    On Error GoTo EH
    If txtPassWord.Text <> vbNullString Then
        SaveECSCryptSetting "V2WebControl", "DBConn", "PASSWORD", txtPassWord.Text
    End If
    mbResetProcess = True
    Exit Sub
EH:
    ShowError Err, "Private Sub txtPassWord_Change", Me
End Sub

Private Sub txtPassWord_GotFocus()
    SelText txtPassWord
End Sub

Private Sub txtSMTPHost_Change()
    On Error GoTo EH
    If txtSMTPHost.Text <> vbNullString Then
        SaveECSCryptSetting "V2WebControl", "SMTP", "Host", txtSMTPHost.Text
    End If
    Exit Sub
EH:
    ShowError Err, "Private Sub txtSMTPHost_Change", Me
End Sub

Private Sub txtSMTPHost_GotFocus()
    SelText txtSMTPHost
End Sub

Private Sub txtSMTPPassword_Change()
    On Error GoTo EH
    If txtSMTPPassword.Text <> vbNullString Then
        SaveECSCryptSetting "V2WebControl", "SMTP", "Password", txtSMTPPassword.Text
    End If
    Exit Sub
EH:
    ShowError Err, "Private Sub txtSMTPPassword_Change", Me
End Sub

Private Sub txtSMTPPassword_GotFocus()
    SelText txtSMTPPassword
End Sub

Private Sub txtSMTPUserName_Change()
    On Error GoTo EH
    If txtSMTPUserName.Text <> vbNullString Then
        SaveECSCryptSetting "V2WebControl", "SMTP", "UserName", txtSMTPUserName.Text
    End If
    Exit Sub
EH:
    ShowError Err, "Private Sub txtSMTPUserName_Change", Me
End Sub

Private Sub txtSMTPUserName_GotFocus()
    SelText txtSMTPUserName
End Sub

Private Sub txtUserID_Change()
    On Error GoTo EH
    If txtUserID.Text <> vbNullString Then
        SaveECSCryptSetting "V2WebControl", "DBConn", "USERID", txtUserID.Text
    End If
    mbResetProcess = True
    Exit Sub
EH:
    ShowError Err, "Private Sub txtUserID_Change", Me
End Sub

Private Sub txtUserID_GotFocus()
    SelText txtUserID
End Sub

Private Sub txtWebControlServer_Change()
    On Error GoTo EH
    SaveSetting "V2WebControl", "Dir", "V2WebControl_SERVER_SHARE", txtWebControlServer.Text
    Exit Sub
EH:
    ShowError Err, "Private Sub txtWebControlServer_Change", Me
End Sub

Private Sub txtWebControlServer_GotFocus()
    SelText txtWebControlServer
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
        If DynamicArraySet(vInit) Then
            For lCount = 0 To UBound(vInit)
                bShowDSN = True
                lstDSN.AddItem vInit(lCount)
            Next
        End If
    End If
    
    Set oReg = New V2ECKeyBoard.clsRegSetting
    'Enumerate all the DSN names in the Registry
    vDSN = oReg.EnumValues(HKEY_LOCAL_MACHINE, "SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources")
    'Add them to the List
    If DynamicArraySet(vDSN) Then
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
    End If
    
    If bShowDSN Then
        framSelectDSN.Caption = "Select DSN (" & moDSNText.Container.Caption & ")"
        framSelectDSN.Visible = True
        framSelectDSN.ZOrder
    End If
    
    'CLeanup
    Set oReg = Nothing
    Exit Sub
EH:
    ShowError Err, "Private Sub GetDSNName", Me
End Sub

Private Sub ResetProcess()
    On Error GoTo EH
    Screen.MousePointer = vbHourglass
    frmProcessData.cmdViewLR.Enabled = False
    frmProcessData.cboLossFormat.Clear
    frmProcessData.LoadDataPaths
    frmProcessData.LoadCompanies
    Screen.MousePointer = vbNormal
    Exit Sub
EH:
    Screen.MousePointer = vbNormal
    ShowError Err, "Private Sub ResetProcess", Me
End Sub
    
Private Sub txtWebSitePath_Change()
    On Error GoTo EH
    SaveSetting "V2WebControl", "Dir", "WebSitePath", txtWebSitePath.Text
    Exit Sub
EH:
    ShowError Err, "Private Sub txtWebSitePath_Change", Me
End Sub

Private Sub txtWebSitePath_GotFocus()
    SelText txtWebSitePath
End Sub
