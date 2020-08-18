VERSION 5.00
Begin VB.Form frmLogOn 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Log On"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5310
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogOn.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   5310
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   3720
      TabIndex        =   6
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Frame framLogOnInfo 
      Appearance      =   0  'Flat
      Caption         =   "Log On Information"
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
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5055
      Begin VB.CommandButton cmdEnterPass 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   4200
         Picture         =   "frmLogOn.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox txtPass 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1560
         MaxLength       =   50
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   720
         Width           =   2500
      End
      Begin VB.TextBox txtUserName 
         Height          =   375
         Left            =   1560
         MaxLength       =   15
         TabIndex        =   2
         Top             =   360
         Width           =   2500
      End
      Begin VB.Label lblLogOn 
         Caption         =   "User Name"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   4815
      End
      Begin VB.Label lblLogOn 
         Caption         =   "Password"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   4845
      End
   End
End
Attribute VB_Name = "frmLogOn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Property Get msClassName() As String
    msClassName = Me.Name
End Property

Private Sub cmdEnterPass_Click()
    On Error GoTo EH
    Dim sRet As String
    Dim sQuestion As String
    Dim sAnswer As String
    Dim sUserName As String
    Dim sPassword As String
    
    sQuestion = goUtil.utGetECSCryptSetting("ECS", "WEB_SECURITY", "LOGON_HINT_Q")
    sAnswer = goUtil.utGetECSCryptSetting("ECS", "WEB_SECURITY", "LOGON_HINT_A")
    sUserName = goUtil.utGetECSCryptSetting("ECS", "WEB_SECURITY", "USER_NAME", vbNullString)
    sPassword = goUtil.utGetECSCryptSetting("ECS", "WEB_SECURITY", "PASSWORD", vbNullString)
    If sQuestion <> vbNullString And sAnswer <> vbNullString Then
        sRet = InputBox(sQuestion, "Logon Hint Question", vbNullString)
        If sRet <> vbNullString Then
            If StrComp(Trim(sAnswer), Trim(sRet), vbTextCompare) = 0 Then
                MsgBox "User Name: " & sUserName & vbCrLf & vbCrLf & "Password: " & sPassword, vbOKOnly, "Logon Info"
            Else
                MsgBox "Invalid Answer!", vbExclamation + vbOKOnly, "Invalid"
            End If
        End If
    Else
        MsgBox "Hint information not found!", vbCritical + vbOKOnly, "Error"
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdEnterPass_Click"
End Sub

Private Sub cmdOK_Click()
    On Error GoTo EH
    If txtUserName.Text = vbNullString Or txtPass.Text = vbNullString Then
        txtUserName.SetFocus
        Exit Sub
    End If
    Me.Hide
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdOK_Click"
End Sub

Private Sub Form_Load()
    On Error GoTo EH
    'BGS 10.10.2002 Put trail "..." on Lables
    goUtil.utSuffixLabels lblLogOn, 50
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_Load"
End Sub

Private Sub txtPass_GotFocus()
    goUtil.utSelText txtPass
End Sub

Private Sub txtUserName_GotFocus()
    goUtil.utSelText txtUserName
End Sub
