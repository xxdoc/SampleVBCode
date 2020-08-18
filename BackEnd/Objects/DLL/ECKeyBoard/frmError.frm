VERSION 5.00
Begin VB.Form frmError 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ERROR"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4125
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmError.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   4125
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2640
      Width           =   615
   End
   Begin VB.TextBox txtError 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   2415
      Left            =   140
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "frmError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbLoaded As Boolean

Public Property Get Loaded() As Boolean
    Loaded = mbLoaded
End Property

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    mbLoaded = True
    txtError.Text = vbNullString
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mbLoaded = False
End Sub

Private Sub txtError_DblClick()
    SelText txtError
End Sub


