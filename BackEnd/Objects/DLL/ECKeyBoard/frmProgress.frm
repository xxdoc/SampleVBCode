VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProgress 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Progress "
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8655
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
   ScaleHeight     =   4110
   ScaleWidth      =   8655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6960
      MaskColor       =   &H80000014&
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Frame framMain 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8415
      Begin VB.TextBox txtDummy 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   7440
         TabIndex        =   11
         Top             =   3600
         Width           =   255
      End
      Begin VB.Frame framRecord 
         Appearance      =   0  'Flat
         Caption         =   "Record Progress"
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
         Height          =   975
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   8175
         Begin MSComctlLib.ProgressBar PBarRecord 
            Height          =   375
            Left            =   120
            TabIndex        =   6
            Top             =   480
            Width           =   7935
            _ExtentX        =   13996
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   0
         End
         Begin VB.Label lblField 
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Tag             =   "Refresh"
            Top             =   240
            Width           =   7905
         End
      End
      Begin VB.Frame framTable 
         Appearance      =   0  'Flat
         Caption         =   "Table Progress"
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
         Height          =   975
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   8175
         Begin MSComctlLib.ProgressBar PBarTable 
            Height          =   375
            Left            =   120
            TabIndex        =   3
            Top             =   480
            Width           =   7935
            _ExtentX        =   13996
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   0
         End
         Begin VB.Label lblTable 
            Height          =   255
            Left            =   120
            TabIndex        =   2
            Tag             =   "Refresh"
            Top             =   240
            Width           =   7935
         End
      End
      Begin VB.Frame framFile 
         Appearance      =   0  'Flat
         Caption         =   "File Progress"
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
         Height          =   975
         Left            =   120
         TabIndex        =   7
         Top             =   2400
         Width           =   8175
         Begin MSComctlLib.ProgressBar PBarFile 
            Height          =   375
            Left            =   120
            TabIndex        =   9
            Top             =   480
            Width           =   7935
            _ExtentX        =   13996
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   0
         End
         Begin VB.Label lblFile 
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Tag             =   "Refresh"
            Top             =   240
            Width           =   7905
         End
      End
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbCancel As Boolean

Private Property Get msClassName() As String
    msClassName = Me.Name
End Property

Public Property Let CancelMe(pbFlag As Boolean)
    mbCancel = pbFlag
End Property
Public Property Get CancelMe() As Boolean
    CancelMe = mbCancel
End Property

Public Sub RefreshMe()
    On Error GoTo EH
    Dim MyControl As Control
    
    For Each MyControl In Me.Controls
        If MyControl.Tag = "Refresh" Then
            MyControl.Refresh
        End If
    Next
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub RefreshMe"
End Sub

Private Sub cmdCancel_Click()
    On Error GoTo EH
    'Hide the Form but Ignore modality
    goUtil.utShowFormIgnoreModality Me, False
    mbCancel = True
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdCancel_Click"
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    goUtil.utAlwaysOnTop Me, True
    
End Sub

Private Sub Form_Paint()
    On Error Resume Next
    goUtil.utAlwaysOnTop Me, True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo EH
    If UnloadMode = vbFormControlMenu Then
        ' Just hide form if user presses Closes by X
        mbCancel = True
        goUtil.utShowFormIgnoreModality Me, False
        Cancel = True
    End If
     Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_QueryUnload"
End Sub



