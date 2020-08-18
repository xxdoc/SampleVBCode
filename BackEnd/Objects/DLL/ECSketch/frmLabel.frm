VERSION 5.00
Begin VB.Form frmLabel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add A Label To Your Sketch"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "frmLabel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton cmdAccept 
      Caption         =   "Accept"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox txtLabel 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   4335
   End
   Begin VB.Label lblHeader 
      Caption         =   "(10 Lines Maximum)"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbCancel As Boolean

Public Property Get Cancel() As Boolean
    Cancel = mbCancel
End Property


Private Sub cmdAccept_Click()
    Me.visible = False
End Sub

Private Sub cmdCancel_Click()
    mbCancel = True
    Me.visible = False
End Sub

Private Sub Form_Activate()
    txtLabel.Text = ""
End Sub
Private Sub txtLabel_Change()
    If txtLabel.Text <> "" Then
        cmdAccept.Enabled = True
    Else
        cmdAccept.Enabled = False
    End If
End Sub

