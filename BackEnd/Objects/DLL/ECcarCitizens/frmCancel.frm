VERSION 5.00
Begin VB.Form frmCancel 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Send to Xactimate"
   ClientHeight    =   1725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4620
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCancel.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   4620
   Begin VB.Frame framCancel 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4335
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&CANCEL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1440
         TabIndex        =   2
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label lblMess 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4095
      End
   End
End
Attribute VB_Name = "frmCancel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents moMyUtil As V2ECKeyBoard.clsUtil
Attribute moMyUtil.VB_VarHelpID = -1

Public Property Let Util(poUtil As V2ECKeyBoard.clsUtil)
    Set moMyUtil = poUtil
End Property
Public Property Set Util(poUtil As V2ECKeyBoard.clsUtil)
    Set moMyUtil = poUtil
End Property
Public Property Get Util() As V2ECKeyBoard.clsUtil
    Set Util = moMyUtil
End Property

Private Sub cmdCancel_Click()
    On Error GoTo EH
    
    goUtil.goXact.CancelSendToXactimate = True
    Screen.MousePointer = vbDefault
    cmdCancel.Enabled = False
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, Me.Name, "Private Sub cmdCancel_Click"
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    goUtil.utAlwaysOnTop Me, True
End Sub

Private Sub Form_Load()
    On Error GoTo EH
    
    Me.top = 100
    Me.left = (Screen.Width / 2) - (Me.Width / 2)
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, Me.Name, "Private Sub Form_Load"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo EH
    
    goUtil.goXact.CancelSendToXactimate = True
    Screen.MousePointer = vbDefault
    Set moMyUtil = Nothing
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, Me.Name, "Private Sub Form_Unload"
End Sub

Private Sub moMyUtil_XactimateProgress(ByVal sMessage As String)
    On Error GoTo EH
    lblMess.Caption = sMessage
    lblMess.Refresh
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, Me.Name, "Private Sub moMyUtil_XactimateProgress"
End Sub

