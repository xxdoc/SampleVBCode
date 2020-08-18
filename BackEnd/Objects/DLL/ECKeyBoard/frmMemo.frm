VERSION 5.00
Begin VB.Form frmMemo 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Memo"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8925
   Icon            =   "frmMemo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   8925
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame framADJ_FTP 
      Caption         =   "Adjuster FTP Directory"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   4440
      Width           =   8655
      Begin VB.TextBox txtMemoTitle 
         Height          =   360
         Left            =   4800
         MaxLength       =   20
         TabIndex        =   6
         Top             =   480
         Width           =   2055
      End
      Begin VB.CommandButton cmdPDFPath 
         Height          =   330
         Left            =   4320
         Picture         =   "frmMemo.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Browse"
         Top             =   495
         Width           =   375
      End
      Begin VB.TextBox txtADJFTPPath 
         Height          =   360
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   4575
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "&Send"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6960
         TabIndex        =   7
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label lblMemoTitle 
         Caption         =   "Memo Title"
         Height          =   255
         Left            =   4800
         TabIndex        =   5
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lblADJFTPPath 
         Caption         =   "ADJ_FTP Path"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   4455
      End
   End
   Begin VB.TextBox txtMess 
      Height          =   4215
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Width           =   8655
   End
End
Attribute VB_Name = "frmMemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mcolCrids As Collection

Private Property Get msClassName() As String
    msClassName = Me.Name
End Property

Public Sub AddCrid(psCRID As String)
    On Error GoTo EH
    If mcolCrids Is Nothing Then
        Set mcolCrids = New Collection
    End If
    mcolCrids.Add psCRID, psCRID
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub AddCrid"
End Sub

Private Sub cmdPDFPath_Click()
    On Error GoTo EH
    txtADJFTPPath.Text = goUtil.utGetPath(App.EXEName, "ADJFTPPath", "BROWSE TO ADJ_FTP PATH", "CLICK OPEN TO SAVE PATH", txtADJFTPPath.Text, Me.hWnd)
    EnableSend
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdPDFPath_Click"
End Sub

Private Sub cmdSend_Click()
    On Error GoTo EH
    Dim sCrid As String
    Dim vCrid As Variant
    Dim sFile As String
    Dim lCount As Long
    
    EnableSend
    If Not cmdSend.Enabled Then
        Exit Sub
    End If
    
    Me.MousePointer = vbHourglass
    Me.Refresh
    For Each vCrid In mcolCrids
        sCrid = vCrid
        sFile = sCrid & "_" & txtMemoTitle.Text & "_" & Format(Now, "YYMMDD") & ".txt"
        sFile = txtADJFTPPath.Text & "\" & sFile
        If goUtil.utFileExists(sFile) Then
            SetAttr sFile, vbNormal
            Kill sFile
        End If
        If Trim(txtMess) = vbNullString Then
            Me.MousePointer = vbDefault
            Me.Refresh
            MsgBox "Nothing to send!", vbInformation + vbOKOnly, "Exiting"
            Me.Visible = False
            Exit Sub
        End If
        goUtil.utSaveFileData sFile, txtMess.Text
        lCount = lCount + 1
    Next
    Me.MousePointer = vbDefault
    Me.Refresh
    MsgBox "Sent MEMO: '" & txtMemoTitle.Text & "' to " & lCount & " recipients.", vbInformation + vbOKOnly, "Memo Sent"
    
    Me.Visible = False
    Exit Sub
EH:
    Me.MousePointer = vbDefault
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdSend_Click"
End Sub

Private Sub Form_Load()
    On Error GoTo EH
     goUtil.utFormWinRegPos goUtil.gsMainAppEXEName, Me
    txtADJFTPPath.Text = GetSetting(App.EXEName, "Dir", "ADJFTPPath", vbNullString)
    EnableSend
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_Load"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Me.Visible = False
        Cancel = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo EH
     goUtil.utFormWinRegPos goUtil.gsMainAppEXEName, Me, True
    Set mcolCrids = Nothing
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_Unload"
End Sub

Private Sub EnableSend()
    On Error GoTo EH
    
    If Not goUtil.utFileExists(txtADJFTPPath.Text, True) Then
        cmdSend.Enabled = False
    Else
        cmdSend.Enabled = True
        SaveSetting App.EXEName, "Dir", "ADJFTPPath", txtADJFTPPath.Text
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub EnableSend"
End Sub

Private Sub txtADJFTPPath_Change()
    goUtil.utCleanFileFolderName txtADJFTPPath, True
    EnableSend
End Sub

Private Sub txtADJFTPPath_GotFocus()
    goUtil.utSelText txtADJFTPPath
End Sub

Private Sub txtMemoTitle_Change()
    goUtil.utCleanFileFolderName txtMemoTitle
    EnableSend
End Sub

Private Sub txtMemoTitle_GotFocus()
    goUtil.utSelText txtMemoTitle
End Sub

Private Sub txtMess_GotFocus()
    goUtil.utSelText txtMess
End Sub
