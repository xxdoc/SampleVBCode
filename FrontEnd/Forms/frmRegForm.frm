VERSION 5.00
Begin VB.Form frmRegForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registration"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7755
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRegForm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   7755
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicMain 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5850
      Left            =   -105
      ScaleHeight     =   5790
      ScaleWidth      =   7815
      TabIndex        =   0
      Top             =   0
      Width           =   7870
      Begin VB.PictureBox picSpinner 
         Appearance      =   0  'Flat
         BackColor       =   &H00A70C0C&
         BorderStyle     =   0  'None
         DrawStyle       =   5  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   3840
         ScaleHeight     =   735
         ScaleWidth      =   615
         TabIndex        =   8
         Top             =   1600
         Width           =   615
      End
      Begin VB.CommandButton cmdOK 
         Cancel          =   -1  'True
         Caption         =   "&Ok"
         Default         =   -1  'True
         Height          =   650
         Left            =   5760
         TabIndex        =   4
         Top             =   4422
         Width           =   1455
      End
      Begin VB.CommandButton cmdViewLic 
         Caption         =   "&View License"
         Height          =   650
         Left            =   5760
         TabIndex        =   3
         Top             =   3702
         Width           =   1455
      End
      Begin VB.TextBox txtCustomerCode 
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
         Left            =   4365
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   3870
         Width           =   1335
      End
      Begin VB.TextBox txtTempCode 
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
         Left            =   4365
         TabIndex        =   1
         Top             =   4590
         Width           =   1335
      End
      Begin VB.Timer TimerUnloadSplash 
         Enabled         =   0   'False
         Interval        =   2000
         Left            =   6960
         Top             =   360
      End
      Begin VB.Label lblMess 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   1335
         Left            =   525
         TabIndex        =   7
         Top             =   3720
         Width           =   3735
      End
      Begin VB.Label lblTempCode 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Enter TempCode"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   4365
         TabIndex        =   5
         Top             =   4380
         Width           =   1335
      End
      Begin VB.Label lblCustomerCode 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   4365
         TabIndex        =   6
         Top             =   3675
         Width           =   1335
      End
      Begin VB.Image imgECSLogo 
         Height          =   5750
         Left            =   0
         Picture         =   "frmRegForm.frx":0442
         Stretch         =   -1  'True
         Top             =   0
         Width           =   7800
      End
   End
End
Attribute VB_Name = "frmRegForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MESS_W As Long = 5175
Private Const LIC_MESS_W As Long = 3735

Private Property Get msClassName() As String
    msClassName = Me.Name
End Property

Private Sub cmdOK_Click()
    On Error GoTo EH
    Dim sTodayDateCode As String
    Dim sTempCode As String
    Dim sMess As String
    
    TimerUnloadSplash.Enabled = False
    If goUtil.gsAppEXEName = goUtil.gsMainAppExeName Then
        gfrmECTray.Timer_SpinMe.Enabled = False
        Set gfrmECTray.Spinner = Nothing
    End If
    
    If txtTempCode.Visible Then
        sTodayDateCode = Format(DateAdd("d", CDbl(Format(Now(), "D")) + CDbl(Format(Now(), "M")), Now()), "MM/DD/YYYY")
        sTodayDateCode = CLng(Format(sTodayDateCode, "YYDM")) & Format(Now(), "DDYYMM")
        sTempCode = Trim(txtTempCode.Text)
        If StrComp(sTodayDateCode, sTempCode, vbBinaryCompare) = 0 Then
            goUtil.gbValidLic = True
            goUtil.utSaveLic 10 'Send in 10 days for initial amount
        End If
    End If
    Me.Visible = False
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdOK_Click"
End Sub

Private Sub cmdViewLic_Click()
    On Error GoTo EH
    Dim dProc As Double
    dProc = Shell("notepad.exe " & goUtil.gsInstallDir & "\License.txt", vbNormalFocus)
    AppActivate dProc, True
    Exit Sub
EH:
    MsgBox "License.txt not found!", vbExclamation, "License Text Not Found"
    Err.Clear
End Sub

Private Sub Form_Load()
    On Error GoTo EH
    Dim sMess As String
    Dim sDaysLeft As String
    Dim lDaysLeft As Long
    
    'If there is an error then they messed with Security settings
    On Error Resume Next
    sDaysLeft = goUtil.utGetECSCryptSetting("ECS", "WEB_SECURITY", "LIC")
    If Err.Number <> 0 Then
        sDaysLeft = "0"
        Err.Clear
    End If
    'Check to see if this is the very first time running on this machine.
    'If it is they will need to contact support to get Temp code for 10 free days.
    'They will then have 10 free days to connect to eberls to get the official lic days.
    'After the 10 days is up The only thing that can be accessed is Communications
    'which will be needed to reset the Lic after contacting support to Reset Lic on the server.
    If sDaysLeft = vbNullString Then
        goUtil.gbInitLic = True 'Running for first time
        lblMess.Width = LIC_MESS_W
        lblMess.Alignment = vbLeftJustify
        lblCustomerCode.Visible = True
        txtCustomerCode.Visible = True
        lblTempCode.Visible = True
        txtTempCode.Visible = True
        txtCustomerCode.Text = Format(Now(), "YYDD00MM")
        sMess = "This application must be licensed!" & vbCrLf
        sMess = sMess & "Please contact your Manager or Support at" & vbCrLf & "www.eberls.com" & vbCrLf
        sMess = sMess & "By entering the Temp Code, you agree to the terms listed in the license."
    Else
LIC:
        lblMess.Width = MESS_W
        lblMess.Alignment = vbCenter
        lblCustomerCode.Visible = False
        txtCustomerCode.Visible = False
        lblTempCode.Visible = False
        txtTempCode.Visible = False
        If Val(sDaysLeft) > 0 Then
            goUtil.gbValidLic = True
'            If Val(sDaysLeft) <= 10 Then
'                MsgBox "Your license will soon expire with " & sDaysLeft & " days of use." & vbCrLf & vbCrLf & "Please contact support to renew your license.", vbInformation, "License Notification"
'            End If
            lDaysLeft = CLng(sDaysLeft)
            goUtil.utSaveLic lDaysLeft
        Else
            goUtil.gbValidLic = False
            MsgBox "Your license has expired!" & vbCrLf & vbCrLf & "Please contact support to renew your license.", vbExclamation, "License Expired"
        End If
        sMess = vbCrLf
        sMess = sMess & "Copyright 2001 - " & Format(Now(), "YYYY") & " Eberls Claim Service, Inc." & vbCrLf
        sMess = sMess & "All Rights Reserved" & vbCrLf & vbCrLf
        sMess = sMess & "License days remaining... " & vbCrLf & lDaysLeft
        TimerUnloadSplash.Enabled = True
    End If
    'Set the mess text
    lblMess.Caption = sMess
    'Start the spinner
    If Not gfrmECTray Is Nothing Then
        Set gfrmECTray.Spinner = Me.picSpinner
        gfrmECTray.Timer_SpinMe.Enabled = True
    End If
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_Load"
End Sub

Private Sub Form_Paint()
'    goUtil.utAlwaysOnTop Me, True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo EH
    Dim sMess As String
    If UnloadMode = vbFormControlMenu Then
        ' Just hide form if user presses Closes by X
        TimerUnloadSplash.Enabled = False
        gfrmECTray.Timer_SpinMe.Enabled = False
        Set gfrmECTray.Spinner = Nothing
        Me.Visible = False
        Cancel = True
    End If
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_QueryUnload"
End Sub

Private Sub TimerUnloadSplash_Timer()
    On Error GoTo EH
    TimerUnloadSplash.Enabled = False
    If goUtil.gsAppEXEName = goUtil.gsMainAppExeName Then
        gfrmECTray.Timer_SpinMe.Enabled = False
        Set gfrmECTray.Spinner = Nothing
    End If
    Me.Visible = False
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub TimerUnloadSplash_Timer"
End Sub

Private Sub txtCustomerCode_GotFocus()
    goUtil.utSelText txtCustomerCode
End Sub

Private Sub txtTempCode_GotFocus()
    goUtil.utSelText txtTempCode
End Sub
