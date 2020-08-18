VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmHelpViewer 
   AutoRedraw      =   -1  'True
   Caption         =   "Browsing"
   ClientHeight    =   4425
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9135
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmHelpViewer.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   9135
   Begin SHDocVwCtl.WebBrowser WBView 
      Height          =   4335
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   9045
      ExtentX         =   15954
      ExtentY         =   7646
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   900
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
      End
      Begin VB.Menu PrintBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuRefresh 
         Caption         =   "&Refresh"
      End
      Begin VB.Menu RefreshBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTextSize 
         Caption         =   "&Text Size"
         Begin VB.Menu mnuSize 
            Caption         =   "&Smallest"
            Index           =   0
         End
         Begin VB.Menu mnuSize 
            Caption         =   "S&maller"
            Index           =   1
         End
         Begin VB.Menu mnuSize 
            Caption         =   "Me&dium"
            Index           =   2
         End
         Begin VB.Menu mnuSize 
            Caption         =   "&Larger"
            Index           =   3
         End
         Begin VB.Menu mnuSize 
            Caption         =   "Lar&gest"
            Index           =   4
         End
      End
   End
End
Attribute VB_Name = "frmHelpViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'SPacing Values Relative to Form Width and Height
Private Const W_WB As Long = 215
Private Const H_WB As Long = 750

Private msHelpFile As String
Private msHelpCaption As String

Private Property Get msClassName() As String
    msClassName = Me.Name
End Property

Public Property Let Helpfile(psFile As String)
    msHelpFile = psFile
End Property
Public Property Get Helpfile() As String
    Helpfile = msHelpFile
End Property

Public Property Let HelpCaption(psCaption As String)
    msHelpCaption = psCaption
End Property
Public Property Get HelpCaption() As String
    HelpCaption = msHelpCaption
End Property

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    On Error GoTo EH
    Dim lErrorCount As Long
    
    WBView.Refresh
    
    Exit Sub
EH:
    If Err.Number = WEB_REFRESH_ERROR And lErrorCount < 20 Then
        lErrorCount = lErrorCount + 1
        DoEvents
        Sleep 200
        Resume
    Else
        goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_Activate"
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo EH
 
    Me.Caption = "Browsing (" & msHelpCaption & ")"
 
    'Get Form Posn from registry
     goUtil.utFormWinRegPos goUtil.gsMainAppEXEName, Me, , , , False

    WBView.Navigate msHelpFile

    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_Load"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo EH
    If UnloadMode = vbFormControlMenu Then
        Me.Visible = False
        Cancel = True
    End If
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_QueryUnload"
End Sub

Private Sub Form_Resize()
    On Error GoTo EH
    'Size Web view
    With WBView
        .Width = Me.Width - W_WB
        If (Me.Height - H_WB) > 0 Then
            .Height = Me.Height - H_WB
        End If
    End With
        
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_Resize"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo EH
    
     goUtil.utFormWinRegPos goUtil.gsMainAppEXEName, Me, True, , , False
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_Unload"
End Sub


Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuPrint_Click()
    On Error GoTo EH
    Me.MousePointer = vbHourglass
    WBView.ExecWB OLECMDID_PRINT, vbNull, vbNull, vbNull
    Me.MousePointer = vbDefault
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub mnuPrint_Click"
End Sub

Private Sub mnuRefresh_Click()
    On Error GoTo EH
    WBView.ExecWB OLECMDID_REFRESH, vbNull, vbNull, vbNull
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub mnuRefresh_Click"
End Sub


Private Sub mnuSize_Click(Index As Integer)
    On Error GoTo EH
    Dim iCount As Integer
    ZOOM Index
    For iCount = 0 To 4
        If iCount = Index Then
            mnuSize(iCount).Checked = True
        Else
            mnuSize(iCount).Checked = False
        End If
    Next
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub mnuSize_Click"
End Sub


Private Sub ZOOM(Index As Integer)
    On Error GoTo EH
    
    WBView.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(Index), vbNull
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub ZOOM"
End Sub

