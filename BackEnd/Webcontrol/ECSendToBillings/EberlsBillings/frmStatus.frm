VERSION 5.00
Begin VB.Form frmStatus 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Send Progress"
   ClientHeight    =   1770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3735
   Icon            =   "frmStatus.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   3735
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame framMess 
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      Begin VB.Label lblMess 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmkOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
End
Attribute VB_Name = "frmStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbOK As Boolean
Private mbCancel As Boolean

Public Property Get OKMe() As Boolean
    OKMe = mbOK
End Property

Public Property Get CancelMe() As Boolean
    CancelMe = mbCancel
End Property

Private Sub cmdCancel_Click()
    mbCancel = True
    mbOK = False
    Me.Visible = False
End Sub

Private Sub cmkOk_Click()
    mbOK = True
    mbCancel = False
    Me.Visible = False
End Sub

Private Sub Form_Activate()
    goUtil.utAlwaysOnTop Me, True
End Sub

Public Sub ResetME()
    mbOK = False
    mbCancel = False
End Sub

