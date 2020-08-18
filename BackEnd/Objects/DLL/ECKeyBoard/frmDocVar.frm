VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmDocVar 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000C&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Document Variables"
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6915
   FillColor       =   &H80000003&
   Icon            =   "frmDocVar.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   6915
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer timAppCHeck 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5160
      Top             =   1920
   End
   Begin VB.PictureBox PicTop 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   6855
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   0
      Width           =   6855
      Begin VB.Frame framName 
         Caption         =   "Selected Report"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   680
         Left            =   120
         TabIndex        =   19
         Top             =   0
         Width           =   6615
         Begin VB.Label lblDate 
            Height          =   375
            Left            =   4320
            TabIndex        =   26
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label lblName 
            Height          =   375
            Left            =   555
            TabIndex        =   20
            Top             =   240
            Width           =   3615
         End
         Begin VB.Image imgSelected 
            Height          =   360
            Left            =   120
            Stretch         =   -1  'True
            Top             =   247
            Width           =   360
         End
      End
      Begin VB.Line Line1 
         X1              =   140
         X2              =   6760
         Y1              =   720
         Y2              =   720
      End
   End
   Begin MSComctlLib.ImageList imgVarDoc 
      Left            =   5640
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocVar.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocVar.frx":077E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocVar.frx":0AAA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picBottom 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3120
      Left            =   0
      ScaleHeight     =   3120
      ScaleWidth      =   6915
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   4680
      Width           =   6915
      Begin VB.CommandButton cmdWordXL 
         Caption         =   "&Reload"
         Enabled         =   0   'False
         Height          =   615
         Left            =   6000
         Picture         =   "frmDocVar.frx":0DCE
         Style           =   1  'Graphical
         TabIndex        =   27
         TabStop         =   0   'False
         Tag             =   "Enable"
         ToolTipText     =   "Reload Word and Excel Applications"
         Top             =   135
         Width           =   855
      End
      Begin VB.Frame framReports 
         Caption         =   "Reports"
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
         Height          =   2895
         Left            =   120
         TabIndex        =   5
         Tag             =   "Enable"
         Top             =   8
         Width           =   5775
         Begin VB.CommandButton cmdSelSaved 
            Caption         =   "&Select All"
            Enabled         =   0   'False
            Height          =   255
            Left            =   3640
            TabIndex        =   10
            Tag             =   "Enable"
            Top             =   2520
            Width           =   975
         End
         Begin VB.CommandButton cmdLoadAvail 
            Caption         =   "&Load"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Tag             =   "Enable"
            Top             =   2520
            Width           =   855
         End
         Begin VB.CommandButton cmdDelSaved 
            Caption         =   "&Delete"
            Enabled         =   0   'False
            Height          =   255
            Left            =   4770
            TabIndex        =   11
            Tag             =   "Enable"
            Top             =   2520
            Width           =   855
         End
         Begin VB.CommandButton cmdLoadSaved 
            Caption         =   "L&oad"
            Enabled         =   0   'False
            Height          =   255
            Left            =   2640
            TabIndex        =   9
            Tag             =   "Enable"
            Top             =   2520
            Width           =   855
         End
         Begin MSComctlLib.ListView lvwAvail 
            Height          =   1935
            Left            =   120
            TabIndex        =   6
            Tag             =   "Enable"
            Top             =   480
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   3413
            View            =   3
            Arrange         =   2
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            _Version        =   393217
            SmallIcons      =   "imgVarDoc"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            Enabled         =   0   'False
            NumItems        =   0
         End
         Begin MSComctlLib.ListView lvwSaved 
            Height          =   1935
            Left            =   2640
            TabIndex        =   8
            Tag             =   "Enable"
            Top             =   480
            Width           =   2985
            _ExtentX        =   5265
            _ExtentY        =   3413
            View            =   3
            Arrange         =   2
            LabelEdit       =   1
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            _Version        =   393217
            SmallIcons      =   "imgVarDoc"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            Enabled         =   0   'False
            NumItems        =   0
         End
         Begin VB.Label lblSaved 
            Caption         =   "Printed Reports"
            Height          =   255
            Left            =   2640
            TabIndex        =   22
            Top             =   240
            Width           =   3015
         End
         Begin VB.Label lblAvail 
            Caption         =   "Available Reports"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   2295
         End
      End
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "&Exit"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6000
         TabIndex        =   13
         Tag             =   "Enable"
         Top             =   2528
         Width           =   855
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6000
         TabIndex        =   12
         Tag             =   "Enable"
         Top             =   2055
         Width           =   855
      End
      Begin VB.Line Line2 
         X1              =   140
         X2              =   6740
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.VScrollBar VSDocVar 
      CausesValidation=   0   'False
      Height          =   3975
      LargeChange     =   5
      Left            =   6495
      Min             =   1
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   720
      Value           =   1
      Width           =   275
   End
   Begin VB.Frame framDocVar 
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   6015
      Begin VB.CommandButton cmdDel 
         Caption         =   "Delete"
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   3
         Tag             =   "Variable"
         Top             =   375
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdQuery 
         Caption         =   "Query"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Tag             =   "Variable"
         Top             =   375
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtValue 
         Height          =   285
         HideSelection   =   0   'False
         Index           =   0
         Left            =   3480
         TabIndex        =   4
         Tag             =   "Variable"
         Top             =   360
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label lblHCommands 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Commands"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblHValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Value"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   255
         Left            =   3480
         TabIndex        =   24
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label lblVarName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   1560
         TabIndex        =   15
         Tag             =   "Variable"
         Top             =   375
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label lblHVarName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Variable Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   255
         Left            =   1560
         TabIndex        =   23
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.PictureBox picRight 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4800
      Left            =   6760
      ScaleHeight     =   4800
      ScaleWidth      =   195
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   0
      Width           =   200
      Begin VB.Line Line4 
         X1              =   0
         X2              =   0
         Y1              =   600
         Y2              =   4800
      End
   End
   Begin VB.PictureBox picLeft 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4800
      Left            =   0
      ScaleHeight     =   4800
      ScaleWidth      =   135
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   0
      Width           =   135
      Begin VB.Line Line3 
         X1              =   120
         X2              =   120
         Y1              =   720
         Y2              =   4680
      End
   End
   Begin VB.Line lvShadow 
      BorderColor     =   &H80000007&
      BorderWidth     =   15
      DrawMode        =   5  'Not Copy Pen
      Visible         =   0   'False
      X1              =   6240
      X2              =   6240
      Y1              =   840
      Y2              =   1560
   End
   Begin VB.Line lhShadow 
      BorderColor     =   &H80000007&
      BorderWidth     =   15
      DrawMode        =   5  'Not Copy Pen
      Visible         =   0   'False
      X1              =   360
      X2              =   6240
      Y1              =   1560
      Y2              =   1560
   End
End
Attribute VB_Name = "frmDocVar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const DELIM As String = ","
Private Const VAR_DELIM As String = "@"

Private Const FRAM_DOCVAR_INIT_H As Long = 855
Private Const FRAM_DOCVAR_INIT_T As Long = 720
Private Const LV_SHADOW_INIT_Y1 As Long = 840
Private Const LV_SHADOW_INIT_Y2 As Long = 1560
Private Const LH_SHADOW_INIT_Y1 As Long = 1560
Private Const LH_SHADOW_INIT_Y2 As Long = 1560
Private Const MAX_HEIGHT As Long = 3240
Private Const BUFFER_SPACE As Long = 2880
Private Const NEXT_ROW As Long = 360

Private mlVshadowY1 As Long
Private mlVshadowY2 As Long
Private mlHShadowY1 As Long
Private mlHShadowY2 As Long
Private mframDocVarTOP As Long
Private mbClicked As Boolean
Private moWordXL As clsWordXL
Private marySavedDocVar() As udtSavedDocVar
Private mbSaveCurDoc As Boolean

Private Property Get msClassName() As String
    msClassName = Me.Name
End Property

Public Property Let WordXL(poWordXL As clsWordXL)
    Set moWordXL = poWordXL
End Property
Public Property Set WordXL(poWordXL As clsWordXL)
    Set moWordXL = poWordXL
End Property

Private Sub cmdDel_Click(Index As Integer)
    DeleteRow Index
End Sub

Private Sub cmdDel_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    mbClicked = True
End Sub

Private Sub cmdDelSaved_Click()
    DeleteSaved
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdLoadAvail_Click()
    LoadAvail
End Sub

Private Sub cmdLoadSaved_Click()
    LoadSaved
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo EH
    Dim iType As Pic
    
    If lblName.Caption <> vbNullString Then
        If InStr(1, lblName.Caption, ".xls", vbTextCompare) > 0 Then
            iType = Pic.XL
        ElseIf InStr(1, lblName.Caption, ".doc", vbTextCompare) > 0 Then
            iType = Pic.Word
        End If
        Screen.MousePointer = vbHourglass
        If moWordXL.PrintIt(Me, iType, lblName.Caption, lblDate.Caption, lblVarName, txtValue, lvwSaved) Then
            'yipee
        End If
        ClearVariables
        Screen.MousePointer = vbDefault
    End If
    Exit Sub
EH:
    Screen.MousePointer = vbDefault
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdPrint_Click"
End Sub

Private Sub cmdQuery_Click(Index As Integer)
    txtValue(Index).Text = moWordXL.QueryVariable(lblVarName(Index).Caption)
    txtValue(Index).SetFocus
End Sub

Private Sub cmdQuery_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    mbClicked = True
End Sub

Private Sub cmdSelSaved_Click()
    On Error GoTo EH
    Dim itmX As listItem
    
    For Each itmX In lvwSaved.ListItems
        itmX.Selected = True
    Next
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdSelSaved_Click"
End Sub

Private Sub cmdWordXL_Click()
    On Error GoTo EH
    
    timAppCHeck.Enabled = False
    ClearVariables
    Screen.MousePointer = vbHourglass
    moWordXL.CLEANUP
    If moWordXL.LoadWordXLAPP(Me) Then
        timAppCHeck.Enabled = True
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
EH:
    Screen.MousePointer = vbDefault
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdWordXL_Click"
End Sub

Private Sub Form_Activate()
    On Error GoTo EH
    Dim MyControl As Control
    Me.Refresh
    
    If moWordXL.LoadWordXLAPP(Me) Then
        For Each MyControl In Me.Controls
            If InStr(1, MyControl.Tag, "Enable", vbTextCompare) > 0 Then
                MyControl.Enabled = True
            End If
        Next
        'BGS 1.4.2002 Keep track if the WOrd and XL apps have closed
        timAppCHeck.Enabled = True
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_Activate"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo EH
    Select Case KeyCode
        Case vbKeyDown
            If VSDocVar.Value + VSDocVar.SmallChange <= VSDocVar.Max Then
                VSDocVar.Value = VSDocVar.Value + VSDocVar.SmallChange
            Else
                VSDocVar.Value = VSDocVar.Max
            End If
        Case vbKeyUp
            If VSDocVar.Value - VSDocVar.SmallChange >= VSDocVar.Min Then
                VSDocVar.Value = VSDocVar.Value - VSDocVar.SmallChange
            Else
                VSDocVar.Value = VSDocVar.Min
            End If
        Case vbKeyPageDown
            If VSDocVar.Value + VSDocVar.LargeChange <= VSDocVar.Max Then
                VSDocVar.Value = VSDocVar.Value + VSDocVar.LargeChange
            Else
                VSDocVar.Value = VSDocVar.Max
            End If
        Case vbKeyPageUp
            If VSDocVar.Value - VSDocVar.LargeChange >= VSDocVar.Min Then
                VSDocVar.Value = VSDocVar.Value - VSDocVar.LargeChange
            Else
                VSDocVar.Value = VSDocVar.Min
            End If
    End Select
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_KeyDown"
End Sub

Private Sub Form_Load()
    goUtil.utFormWinRegPos goUtil.gsMainAppEXEName, Me
    
    LoadHeader
    PopulatelvwReports
    mbSaveCurDoc = False
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.BackColor = &H80000001
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.BackColor = &H8000000C
End Sub

Private Sub Form_Resize()
    If Me.WindowState <> vbMinimized Then
        AlwaysOnTop Me, True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo EH
    goUtil.utFormWinRegPos goUtil.gsMainAppEXEName, Me, True
    'bgs 1.4.2002 store the printed documents
    moWordXL.StoreSaved lvwSaved
    'clean up
    timAppCHeck.Enabled = False
    Set moWordXL = Nothing
    
    If Not goUtil.gfrmECTray Is Nothing Then
        goUtil.gfrmECTray.ShowMe False
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_Unload"
End Sub

Private Sub lvwAvail_DblClick()
    LoadAvail
End Sub

Private Sub lvwAvail_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo EH
    
    Select Case KeyCode
        Case vbKeyReturn
            LoadAvail
    End Select
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lvwAvail_KeyDown"
End Sub

Private Sub lvwSaved_DblClick()
    LoadSaved
End Sub

Private Sub lvwSaved_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo EH
    
    Select Case KeyCode
        Case vbKeyReturn
            LoadSaved
        Case vbKeyDelete
            DeleteSaved
    End Select
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lvwSaved_KeyDown"
End Sub

Private Sub timAppCHeck_Timer()
    On Error GoTo EH
    
    If GetCaption(moWordXL.hWndWord) = vbNullString Then
        timAppCHeck.Enabled = False
        moWordXL.CLEANUP
        Unload Me
        Exit Sub
    End If
    
    If GetCaption(moWordXL.hWndXL) = vbNullString Then
        timAppCHeck.Enabled = False
        moWordXL.CLEANUP
        Unload Me
        Exit Sub
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub timAppCHeck_Timer"
End Sub

Private Sub txtValue_GotFocus(Index As Integer)
    On Error GoTo EH
    Dim lTop As Long
    
    
    VSDocVar.Value = Index
    SelText txtValue(Index)
    txtValue(Index).SetFocus
    If Not mbClicked Then
        If txtValue(Index).top - txtValue(0).top > MAX_HEIGHT Then
            lTop = txtValue(Index).top - BUFFER_SPACE
            framDocVar.top = mframDocVarTOP - lTop
            lvShadow.Y1 = mlVshadowY1 - lTop
            lvShadow.Y2 = mlVshadowY2 - lTop
            lhShadow.Y1 = mlHShadowY1 - lTop
            lhShadow.Y2 = mlHShadowY2 - lTop
        Else
            framDocVar.top = FRAM_DOCVAR_INIT_T
            lvShadow.Y1 = mlVshadowY1
            lvShadow.Y2 = mlVshadowY2
            lhShadow.Y1 = mlHShadowY1
            lhShadow.Y2 = mlHShadowY2
        End If
        Me.Refresh
    Else
        mbClicked = False
    End If
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub txtValue_GotFocus"
End Sub

Private Sub MoveMe(Index As Integer)
    On Error GoTo EH
    Dim lTop As Long
    
    If txtValue(Index).top - txtValue(0).top > MAX_HEIGHT Then
        lTop = txtValue(Index).top - BUFFER_SPACE
        framDocVar.top = mframDocVarTOP - lTop
        lvShadow.Y1 = mlVshadowY1 - lTop
        lvShadow.Y2 = mlVshadowY2 - lTop
        lhShadow.Y1 = mlHShadowY1 - lTop
        lhShadow.Y2 = mlHShadowY2 - lTop
    Else
        framDocVar.top = FRAM_DOCVAR_INIT_T
        lvShadow.Y1 = mlVshadowY1
        lvShadow.Y2 = mlVshadowY2
        lhShadow.Y1 = mlHShadowY1
        lhShadow.Y2 = mlHShadowY2
    End If
    Me.Refresh
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub MoveMe"
End Sub
Private Sub txtValue_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo EH
    Select Case KeyCode
        Case vbKeyReturn
            If VSDocVar.Value + VSDocVar.SmallChange <= VSDocVar.Max Then
                VSDocVar.Value = VSDocVar.Value + VSDocVar.SmallChange
            Else
                VSDocVar.Value = VSDocVar.Max
            End If
    End Select
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub txtValue_KeyDown"
End Sub

Private Sub txtValue_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    mbClicked = True
End Sub

Private Sub VSDocVar_Change()
    On Error GoTo EH
    Dim lCount As Long
    Dim Index As Integer
    
    VSDocVar.SetFocus
    
    Index = VSDocVar.Value
    If Index <= txtValue.UBound Then
        If txtValue(Index).Visible Then
            MoveMe Index
            For lCount = txtValue.LBound To txtValue.UBound
                txtValue(lCount).SelLength = 0
            Next
            SelText txtValue(Index)
        End If
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub VSDocVar_Change"
End Sub

Public Function DeleteRow(piIndex As Integer) As Boolean
    On Error GoTo EH
    Dim lCount As Long
    
    cmdQuery(piIndex).Visible = False
    cmdQuery(piIndex).TabStop = False
    cmdDel(piIndex).Visible = False
    cmdDel(piIndex).TabStop = False
    lblVarName(piIndex).Enabled = False
    lblVarName(piIndex).BackColor = &HC0C0FF
    txtValue(piIndex).Locked = True
    txtValue(piIndex).BackColor = &H80000013
    txtValue(piIndex).TabStop = False
    txtValue(piIndex).SetFocus
    DeleteRow = True
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function DeleteRow"
End Function

Private Sub LoadHeader()
    On Error GoTo EH
    'set the columnheaders
    With lvwAvail
        .Sorted = True
        .ColumnHeaders.Add , "Name", "Name"
        
        '"Avail WOrd XL Forms"
        .ColumnHeaders.Item(AvailDocs.Name).Width = 5000
        .ColumnHeaders.Item(AvailDocs.Name).Alignment = lvwColumnLeft
       
    End With
    
    With lvwSaved
        .Sorted = True
        .ColumnHeaders.Add , "Name", "Name"
        .SortOrder = lvwAscending
        .ColumnHeaders.Add , "Date", "Date"
        .SortOrder = lvwAscending
        .ColumnHeaders.Add , "Variables", "Variables"
        
        '"Saved WOrd XL Forms"
        .ColumnHeaders.Item(SavedDocs.Name).Width = 2000
        .ColumnHeaders.Item(SavedDocs.Date).Width = 3000
        .ColumnHeaders.Item(SavedDocs.Variables).Width = 0 'Hidden
    End With
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub LoadHeader"
End Sub

Public Sub PopulatelvwReports()
    On Error GoTo EH
    'Source Variables
    Dim itmX As listItem
    Dim varyAvail As Variant
    Dim varySaved As Variant
    Dim sAvail As String
    Dim SavedDocVar As udtSavedDocVar
    Dim iCount As Integer
    Dim iPic As Integer
    
    lvwAvail.ListItems.Clear
    lvwSaved.ListItems.Clear
    
    'BGS 1.2.2002 load the Avail reports
    If moWordXL.GetAvail(varyAvail) Then
        If IsArray(varyAvail) Then
            For iCount = LBound(varyAvail, 1) To UBound(varyAvail, 1)
                sAvail = varyAvail(iCount)
                If InStr(1, sAvail, ".xls", vbTextCompare) Then
                    iPic = Pic.XL
                ElseIf InStr(1, sAvail, ".doc", vbTextCompare) Then
                    iPic = Pic.Word
                End If
                lvwAvail.ListItems.Add , , sAvail, , iPic
            Next
        End If
    End If
    
    'BGS 1.2.2002 load the Saved Reports
     If moWordXL.GetSaved(varySaved) Then
        If IsArray(varySaved) Then
            marySavedDocVar = varySaved
            For iCount = LBound(marySavedDocVar, 1) To UBound(marySavedDocVar, 1)
                SavedDocVar = marySavedDocVar(iCount)
                If InStr(1, SavedDocVar.a01Name, ".xls", vbTextCompare) Then
                    iPic = Pic.XL
                ElseIf InStr(1, SavedDocVar.a01Name, ".doc", vbTextCompare) Then
                    iPic = Pic.Word
                End If
                Set itmX = lvwSaved.ListItems.Add(, , SavedDocVar.a01Name, , iPic)
                itmX.SubItems(SavedDocs.Date - 1) = SavedDocVar.a02Date
                itmX.SubItems(SavedDocs.Variables - 1) = SavedDocVar.a03Variables
            Next
        End If
        
     End If
    'cleanup
    Set itmX = Nothing
    Exit Sub
EH:
    Set itmX = Nothing
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub PopulatelvwReports"
End Sub

Public Function LoadVariables() As Boolean
    On Error GoTo EH
    Dim myValue As TextBox
    Dim MyName As Label
    Dim MyCmdQuery As CommandButton
    Dim MyCmdDel As CommandButton
    Dim lTopVarName As Long
    Dim lLeftVarName As Long
    Dim lTopValue As Long
    Dim lLeftValue As Long
    Dim lTabIndex As Long
    Dim lTopQuery As Long
    Dim lLeftQuery As Long
    Dim lTopDel As Long
    Dim lLeftDel As Long
    Dim lCount As Long
    Dim aryDocVariables() As QVariable
    
    aryDocVariables = moWordXL.aryDocVariables
    
    Screen.MousePointer = vbHourglass
    framDocVar.Visible = False
    lvShadow.Visible = False
    lhShadow.Visible = False
    
    lTopQuery = cmdQuery(0).top
    lLeftQuery = cmdQuery(0).left
    lTopDel = cmdDel(0).top
    lLeftDel = cmdDel(0).left
    lTopVarName = lblVarName(0).top
    lLeftVarName = lblVarName(0).left
    lTopValue = txtValue(0).top
    lLeftValue = txtValue(0).left
    lTabIndex = txtValue(0).TabIndex
    
    
    For lCount = LBound(aryDocVariables, 1) To UBound(aryDocVariables, 1)
        lTopQuery = lTopQuery + NEXT_ROW
        lTopDel = lTopDel + NEXT_ROW
        lTopVarName = lTopVarName + NEXT_ROW
        lTopValue = lTopValue + NEXT_ROW
        framDocVar.Height = framDocVar.Height + NEXT_ROW
        lvShadow.Y2 = lvShadow.Y2 + NEXT_ROW
        lhShadow.Y1 = lhShadow.Y1 + NEXT_ROW
        lhShadow.Y2 = lhShadow.Y2 + NEXT_ROW
        lTabIndex = lTabIndex + 1
        'Vert Scroll bar
        VSDocVar.Max = lCount
        
        'BGS Load the Next row Controls
        Load cmdQuery(lCount)
        Load cmdDel(lCount)
        Load lblVarName(lCount)
        Load txtValue(lCount)
        
        txtValue(lCount).Text = aryDocVariables(lCount).Value
        lblVarName(lCount).Caption = aryDocVariables(lCount).Name
        
        
        Set MyCmdQuery = cmdQuery(lCount)
        Set MyCmdDel = cmdDel(lCount)
        Set myValue = txtValue(lCount)
        Set MyName = lblVarName(lCount)
        
        MyCmdQuery.top = lTopQuery
        MyCmdQuery.left = lLeftQuery
        MyCmdQuery.TabIndex = lTabIndex
        MyCmdQuery.Visible = True
        
        lTabIndex = lTabIndex + 1
        
        
        MyCmdDel.top = lTopDel
        MyCmdDel.left = lLeftDel
        MyCmdDel.TabIndex = lTabIndex
        MyCmdDel.Visible = True
        
        lTabIndex = lTabIndex + 1
        
        
        MyName.top = lTopVarName
        MyName.left = lLeftVarName
        MyName.Visible = True
        
        myValue.top = lTopValue
        myValue.left = lLeftValue
        myValue.TabIndex = lTabIndex
        myValue.Visible = True
        
    Next
    
    'BGS remember the initial tops
    mlVshadowY1 = lvShadow.Y1
    mlVshadowY2 = lvShadow.Y2
    mlHShadowY1 = lhShadow.Y1
    mlHShadowY2 = lhShadow.Y2
    mframDocVarTOP = framDocVar.top
    
    framDocVar.Visible = True
    lvShadow.Visible = True
    lhShadow.Visible = True
    
    Me.Refresh
    Screen.MousePointer = vbDefault
    
    LoadVariables = True
    Exit Function
EH:
    Screen.MousePointer = vbDefault
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function LoadVariables"
End Function

Public Function ClearVariables() As Boolean
    On Error GoTo EH
    Dim MyControl As Control
    framDocVar.Visible = False
    lvShadow.Visible = False
    lhShadow.Visible = False
    Screen.MousePointer = vbHourglass
    For Each MyControl In Me.Controls
        If InStr(1, MyControl.Tag, "Variable", vbTextCompare) > 0 Then
            If MyControl.Index > 0 Then
                Unload MyControl
                Set MyControl = Nothing
            End If
        End If
    Next
    
    framDocVar.top = FRAM_DOCVAR_INIT_T
    framDocVar.Height = FRAM_DOCVAR_INIT_H
    lvShadow.Y1 = LV_SHADOW_INIT_Y1
    lvShadow.Y2 = LV_SHADOW_INIT_Y2
    lhShadow.Y1 = LH_SHADOW_INIT_Y1
    lhShadow.Y2 = LH_SHADOW_INIT_Y2
    
    'BGS clear the Header pane too
    imgSelected.Picture = Nothing
    lblName.Caption = vbNullString
    lblDate.Caption = vbNullString
    
    Screen.MousePointer = vbDefault
    ClearVariables = True
    Exit Function
EH:
    Screen.MousePointer = vbDefault
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function ClearVariables"
End Function

Private Sub VSDocVar_GotFocus()
    mbClicked = False
End Sub

Private Function LoadAvail() As Boolean
    On Error GoTo EH
    Dim itmX As listItem
    For Each itmX In lvwAvail.ListItems
        If itmX.Selected Then
            ClearVariables
            
            imgSelected.Picture = imgVarDoc.ListImages.Item(Pic.HourGlass).Picture
            imgSelected.Refresh
            lblName.Caption = itmX.Text
            lblName.Refresh
            lblDate.Caption = "Loading, please wait..."
            lblDate.Refresh
            
            If moWordXL.LoadaryDocVariables(itmX.SmallIcon, itmX.Text) Then
                LoadVariables
                lblDate.Caption = "Print ?"
                mbSaveCurDoc = True
                lblDate.Refresh
                imgSelected.Picture = imgVarDoc.ListImages.Item(itmX.SmallIcon).Picture
                imgSelected.Refresh
            Else
                imgSelected.Picture = imgVarDoc.ListImages.Item(itmX.SmallIcon).Picture
                imgSelected.Refresh
                lblDate.Caption = "No variables found."
                mbSaveCurDoc = False
            End If
            LoadAvail = True
            Exit For
        End If
    Next
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Function LoadAvail"
End Function

Private Function LoadSaved() As Boolean
    On Error GoTo EH
    Dim itmX As listItem
    For Each itmX In lvwSaved.ListItems
        If itmX.Selected Then
            ClearVariables
            
            imgSelected.Picture = imgVarDoc.ListImages.Item(Pic.HourGlass).Picture
            imgSelected.Refresh
            lblName.Caption = itmX.Text
            lblName.Refresh
            lblDate.Caption = "Loading, please wait..."
            lblDate.Refresh
            
            
            
            If moWordXL.LoadaryDocVariables(itmX.SmallIcon, itmX.Text, True, itmX.SubItems(SavedDocs.Variables - 1)) Then
                LoadVariables
                lblDate.Caption = itmX.SubItems(SavedDocs.Date - 1)
                mbSaveCurDoc = True
                lblDate.Refresh
                imgSelected.Picture = imgVarDoc.ListImages.Item(itmX.SmallIcon).Picture
                imgSelected.Refresh
            Else
                imgSelected.Picture = imgVarDoc.ListImages.Item(itmX.SmallIcon).Picture
                imgSelected.Refresh
                lblDate.Caption = "No variables found."
                mbSaveCurDoc = False
            End If
            LoadSaved = True
            Exit For
        End If
        
    Next
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Function LoadSaved"
End Function

Private Function DeleteSaved() As Boolean
    On Error GoTo EH
    Dim itmX As listItem
DELETE_ITEMS:
    For Each itmX In lvwSaved.ListItems
        If itmX.Selected Then
            DeleteSaved = True
            lvwSaved.ListItems.Remove itmX.Index
            Set itmX = Nothing
            GoTo DELETE_ITEMS
        End If
    Next
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Function DeleteSaved"
End Function


