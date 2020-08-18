VERSION 5.00
Begin VB.Form frmWallPaper 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5640
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7035
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   7035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picWallPaper 
      BackColor       =   &H8000000C&
      DrawStyle       =   4  'Dash-Dot-Dot
      FillStyle       =   7  'Diagonal Cross
      Height          =   5520
      Left            =   60
      ScaleHeight     =   5460
      ScaleWidth      =   6855
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   60
      Width           =   6920
      Begin VB.Image imgWallPic 
         Height          =   3015
         Left            =   240
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   2895
      End
   End
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5775
      Left            =   1440
      ScaleHeight     =   5775
      ScaleWidth      =   5655
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   -120
      Width           =   5655
   End
End
Attribute VB_Name = "frmWallPaper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private msLastPicPath As String

Private Property Get msClassName() As String
    msClassName = Me.Name
End Property

Public Property Get LastPicPath() As String
    LastPicPath = msLastPicPath
End Property
Public Property Let LastPicPath(psPath As String)
    msLastPicPath = psPath
End Property

Private Sub Form_Load()
    On Error GoTo EH
    
    Me.Visible = False
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_Load"
End Sub

Private Sub Form_Paint()
    On Error GoTo EH
    Dim sNavScreenPos As String
    
    sNavScreenPos = GetSetting(App.EXEName, "GENERAL", "NAV_SCREEN_POS", "RIGHT")
    
    If Not gfrmECTray Is Nothing Then
        If gfrmECTray.WindowState = vbMinimized Then
            gfrmECTray.HideAll True
            Exit Sub
        End If
    End If
    Select Case UCase(sNavScreenPos)
        Case "RIGHT"
            picBack.left = 0
        Case "LEFT"
            picBack.left = 1500
    End Select
    
    Exit Sub
EH:
    Err.Clear
    Resume Next
End Sub

Private Sub Form_Resize()
    On Error GoTo EH
     Dim sNavScreenPos As String
    
    sNavScreenPos = GetSetting(App.EXEName, "GENERAL", "NAV_SCREEN_POS", "RIGHT")
    
    
    picWallPaper.Width = Me.Width - 115
    picWallPaper.Height = Me.Height - 120
    'Need to cover the Edges of the Form except those covered by ECtray
    'So ECtray will be the only form to Fire Wall paper Paint event.
    picBack.Width = Me.Width - 1500
    picBack.top = 0
    picBack.Height = Me.Height
    Select Case UCase(sNavScreenPos)
        Case "RIGHT"
            picBack.left = 0
        Case "LEFT"
            picBack.left = 1500
    End Select
    
    Exit Sub
EH:
    If Err.Number > 0 Then
        Err.Clear
        Resume Next
    End If
End Sub

Private Sub imgWallPic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If goUtil.gbValidLic Then
        gfrmECTray.ShowMe , False
    End If
End Sub


Private Sub picBack_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If goUtil.gbValidLic Then
        gfrmECTray.ShowMe , False
    End If
End Sub

Private Sub picWallPaper_GotFocus()
    On Error Resume Next
    If gfrmECTray.Visible Then
        gfrmECTray.SetFocus
    End If
End Sub


Private Sub picWallPaper_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If goUtil.gbValidLic Then
        gfrmECTray.ShowMe , False
    End If
End Sub
