VERSION 5.00
Begin VB.Form frmSupport 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Support"
   ClientHeight    =   7380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10335
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSupport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7380
   ScaleWidth      =   10335
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   8760
      TabIndex        =   10
      Top             =   6960
      Width           =   1455
   End
   Begin VB.Frame framEasyClaim 
      Appearance      =   0  'Flat
      Caption         =   "Easy Claim Support"
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
      Height          =   6855
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   10095
      Begin VB.CommandButton cmdReinstallSP 
         Caption         =   "&Reinstall / Repair"
         Height          =   375
         Left            =   7200
         TabIndex        =   5
         Top             =   600
         Width           =   2775
      End
      Begin VB.TextBox txtWinOSVersion 
         Height          =   1815
         Left            =   2400
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   4920
         Width           =   7575
      End
      Begin VB.TextBox txtApplication 
         Height          =   3855
         Left            =   2400
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   1080
         Width           =   7575
      End
      Begin VB.TextBox txtEmail 
         Height          =   375
         Left            =   2400
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   4
         Top             =   720
         Width           =   2535
      End
      Begin VB.TextBox txtPhone 
         Height          =   375
         Left            =   2400
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   2
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label lblSupport 
         Caption         =   "Windows OS Version"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   8
         Top             =   5700
         Width           =   4995
      End
      Begin VB.Label lblSupport 
         Caption         =   "Application Version"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   2880
         Width           =   5295
      End
      Begin VB.Label lblSupport 
         Caption         =   "Contact E-Mail"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   3735
      End
      Begin VB.Label lblSupport 
         Caption         =   "Contact Phone"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   3735
      End
   End
End
Attribute VB_Name = "frmSupport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Property Get msClassName() As String
    msClassName = Me.Name
End Property

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub cmdReinstallSP_Click()
    On Error GoTo EH
    Dim sMess As String
    Dim lRet As Long
    Dim bFTPExists As Boolean
    Dim sECSPLexePath As String
    Dim sMyCommandStr As String
    Dim sECFTPListPath As String
    Dim sECFTPListData As String  'used to build the data for ECFTP.lst File
    Dim sSPFile As String
    
    
    'Build List of Installed Regesettings, Documents, and Applications
    'Execute them.  This will require the System to Reboot.
    
    sMess = "Reinstallation will require ALL " & App.EXEName & " Components " & vbCrLf
    sMess = sMess & "to shut down.  Click ""OK"" if you do not need to save your work."
    
    If MsgBox(sMess, vbOKCancel + vbExclamation, "Reinstall / Repair") = vbCancel Then
        Exit Sub
    End If
    
     'If the FTP Application is currently connected and download or Uploading
     'Data then can't run this at all !
    If goUtil.utFTPConnected Then
        sMess = "FTP connection is currently active." & vbCrLf & vbCrLf
        sMess = sMess & "You must wait until the current connection is finished " & vbCrLf
        sMess = sMess & "before running this utility."
        MsgBox sMess, vbExclamation + vbOKOnly, "FTP Connection Detected"
        Exit Sub
    End If
    
    'If the FTP Application is running then need to close it before
    'Running the Compact Repair.
    
    bFTPExists = goUtil.utFTPExists
    If bFTPExists Then
        sMess = "Closing FTP Application!" & vbCrLf & vbCrLf
        MsgBox sMess, vbInformation + vbOKOnly, "FTP SHUTDOWN"
        goUtil.utShutDownFTP
        Sleep 2000 'Wait couple seconds
    End If
    
    'Run the Reinstall
    
    'Build List of Installed Regesettings, Documents, and Applications
    sECFTPListPath = goUtil.gsInstallDir & "\" & goUtil.utGetTickCount & "_ECFTP.lst"
    sECSPLexePath = goUtil.gsInstallDir & "\ECSPL.exe"
    sMyCommandStr = " " & sECFTPListPath
    
    'Regsetting
    sSPFile = Dir(goUtil.SPPath & "RegSetting\SP\*.exe", vbNormal)
    Do Until sSPFile = vbNullString
        If sECFTPListData <> vbNullString Then
            sECFTPListData = sECFTPListData & vbCrLf
        End If
        sECFTPListData = sECFTPListData & goUtil.SPPath & "RegSetting\SP\" & sSPFile
        sSPFile = Dir
    Loop
    'Document
    sSPFile = Dir(goUtil.SPPath & "Document\SP\*.exe", vbNormal)
    Do Until sSPFile = vbNullString
        If sECFTPListData <> vbNullString Then
            sECFTPListData = sECFTPListData & vbCrLf
        End If
        sECFTPListData = sECFTPListData & goUtil.SPPath & "Document\SP\" & sSPFile
         sSPFile = Dir
    Loop
    'Application
    sSPFile = Dir(goUtil.SPPath & "Application\SP\*.exe", vbNormal)
    Do Until sSPFile = vbNullString
        'Do not include the ECSPL.exe Software Package
        If InStr(1, sSPFile, "ECSPL.exe", vbTextCompare) = 0 Then
            If sECFTPListData <> vbNullString Then
                sECFTPListData = sECFTPListData & vbCrLf
            End If
            sECFTPListData = sECFTPListData & goUtil.SPPath & "Application\SP\" & sSPFile
        End If
        sSPFile = Dir
    Loop
    'DataBase
     sSPFile = Dir(goUtil.SPPath & "DataBase\SP\*.exe", vbNormal)
    Do Until sSPFile = vbNullString
        If sECFTPListData <> vbNullString Then
            sECFTPListData = sECFTPListData & vbCrLf
        End If
        sECFTPListData = sECFTPListData & goUtil.SPPath & "DataBase\SP\" & sSPFile
         sSPFile = Dir
    Loop
    
    If sECFTPListData = vbNullString Then
        MsgBox "No Software Found!", vbExclamation + vbOKOnly, "Software Not Found!"
    Else
        goUtil.utSaveFileData sECFTPListPath, sECFTPListData
        Shell sECSPLexePath & sMyCommandStr, vbNormalFocus
        DoEvents
        Sleep 2000
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_Load"
End Sub

Private Sub Form_Load()
    On Error GoTo EH
    goUtil.utFormWinRegPos goUtil.gsMainAppExeName, Me
    
    'BGS 10.10.2002 Put trail "..." on Lables
    goUtil.utSuffixLabels lblSupport, 50
    
    PopulateForm
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_Load"
End Sub

Private Sub PopulateForm()
    On Error GoTo EH
    
    'Use Regsetting incase we ever change these :)
    txtPhone.Text = GetSetting(App.EXEName, "GENERAL", "SUPPORT_PHONE", "303 988 6286")
    txtEmail.Text = GetSetting(App.EXEName, "GENERAL", "SUPPORT_EMAIL", "support@eberls.com")
    txtApplication.Text = goUtil.utGetAppVSInfo(App.EXEName, goUtil.gsInstallDir)
    txtWinOSVersion.Text = goUtil.utGetWinOSVersion
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub PopulateForm"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo EH
    goUtil.utFormWinRegPos goUtil.gsMainAppExeName, Me, True
    Me.Visible = False
    If Not gfrmECTray Is Nothing Then
        gfrmECTray.ShowMe False
    End If
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_Unload"
End Sub

Private Sub txtApplication_GotFocus()
    goUtil.utSelText txtApplication
End Sub

Private Sub txtEmail_GotFocus()
    goUtil.utSelText txtEmail
End Sub

Private Sub txtPhone_GotFocus()
    goUtil.utSelText txtPhone
End Sub

Private Sub txtWinOSVersion_GotFocus()
    goUtil.utSelText txtWinOSVersion
End Sub
