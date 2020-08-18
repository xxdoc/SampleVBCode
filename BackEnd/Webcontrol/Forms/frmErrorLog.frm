VERSION 5.00
Begin VB.Form frmErrorLog 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Error Log"
   ClientHeight    =   7860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9615
   Icon            =   "frmErrorLog.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   9615
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame framErrors 
      Caption         =   "Error Events"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9375
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   6960
         TabIndex        =   2
         Top             =   240
         Width           =   2280
      End
      Begin VB.CommandButton cmdDelAll 
         Caption         =   "Delete &All"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   6960
         TabIndex        =   3
         Top             =   780
         Width           =   2280
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "&Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   8160
         TabIndex        =   5
         Top             =   1320
         Width           =   1080
      End
      Begin VB.CheckBox chkSelAll 
         Caption         =   "Select All &Text"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   6960
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1320
         Width           =   1095
      End
      Begin VB.ListBox lstErrors 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1530
         ItemData        =   "frmErrorLog.frx":0742
         Left            =   120
         List            =   "frmErrorLog.frx":0744
         TabIndex        =   1
         Top             =   240
         Width           =   6735
      End
   End
   Begin VB.TextBox txtErrors 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5775
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   1920
      Width           =   9375
   End
End
Attribute VB_Name = "frmErrorLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDelAll_Click()
    On Error GoTo EH
    Dim lCount As Long
    Dim sFile As String
    
    If MsgBox("Are you sure?", vbQuestion + vbYesNo, "Delete All") = vbYes Then
        If lstErrors.ListCount > 0 Then
            For lCount = 0 To lstErrors.ListCount - 1
                sFile = lstErrors.List(lCount)
                If FileExists(sFile) Then
                    SetAttr sFile, vbNormal
                    Kill sFile
                End If
            Next
        End If
        
        lstErrors.Clear
        txtErrors.Text = vbNullString
    End If
    
    Exit Sub
EH:
    ShowError Err, "Private Sub cmdDelAll_Click", Me
End Sub

Private Sub cmdDelete_Click()
    On Error GoTo EH
    Dim lCount As Long
    Dim sFile As String
    
    If MsgBox("Are you sure?", vbQuestion + vbYesNo, "Delete Selected Item") = vbYes Then
        If lstErrors.ListCount > 0 Then
            For lCount = 0 To lstErrors.ListCount - 1
                If lstErrors.Selected(lCount) Then
                    sFile = lstErrors.List(lCount)
                    If FileExists(sFile) Then
                        SetAttr sFile, vbNormal
                        Kill sFile
                        lstErrors.RemoveItem lCount
                        txtErrors.Text = vbNullString
                        Exit For
                    End If
                End If
            Next
        End If
    End If
    
    Exit Sub
EH:
    ShowError Err, "Private Sub cmdDelete_Click", Me
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdView_Click()
    ViewLog
End Sub

Private Sub Form_Load()
    On Error GoTo EH
    Dim colLogs As Collection
    Dim sFile As String
    Dim vLogFile As Variant
    sFile = App.Path
    
    sFile = Dir(sFile & "\*_Error.log", vbNormal)
    Do Until sFile = vbNullString
        If colLogs Is Nothing Then
            Set colLogs = New Collection
        End If
        colLogs.Add App.Path & "\" & sFile
        sFile = Dir
    Loop
    
    If Not colLogs Is Nothing Then
        For Each vLogFile In colLogs
            sFile = vLogFile
            lstErrors.AddItem sFile
        Next
        
    End If
    
    
    Exit Sub
EH:
    ShowError Err, "Private Sub Form_Load", Me
End Sub

Private Sub lstErrors_Click()
    ViewLog
End Sub

Private Sub txtErrors_GotFocus()
    If chkSelAll.Value = vbChecked Then
        SelText txtErrors
    End If
End Sub

Private Sub ViewLog()
    On Error GoTo EH
    Dim sFile As String
    
    sFile = lstErrors.Text
    
    If FileExists(sFile) Then
        txtErrors.Text = GetFileData(sFile)
    End If
    
    Exit Sub
EH:
    ShowError Err, "Private Sub ViewLog", Me
End Sub
