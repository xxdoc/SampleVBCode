VERSION 5.00
Object = "{B0AA617A-8DB4-4E9E-BBC1-CA4E3B6280AA}#2.0#0"; "ecsTimeOCX.ocx"
Begin VB.Form AddActLog 
   AutoRedraw      =   -1  'True
   Caption         =   "Activity Log"
   ClientHeight    =   6915
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8415
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "AddActLog.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   8415
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame framCommands 
      Height          =   1215
      Left            =   4980
      TabIndex        =   10
      Top             =   0
      Width           =   3375
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   855
         Left            =   2280
         MaskColor       =   &H00000000&
         Picture         =   "AddActLog.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Exit"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Sa&ve"
         Enabled         =   0   'False
         Height          =   855
         Left            =   1200
         MaskColor       =   &H00000000&
         Picture         =   "AddActLog.frx":0316
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Exit"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdSpelling 
         Caption         =   "Spellin&g"
         Height          =   855
         Left            =   120
         MaskColor       =   &H00000000&
         Picture         =   "AddActLog.frx":0460
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Exit"
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame framActText 
      Caption         =   "Activity Text ( - 32000 character limit -): "
      Height          =   4935
      Left            =   60
      TabIndex        =   8
      Top             =   1920
      Width           =   8295
      Begin VB.TextBox txtActText 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4560
         Left            =   120
         MaxLength       =   32000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   240
         Width           =   8055
      End
   End
   Begin VB.Frame framDatesTime 
      Caption         =   "Date && Time"
      Height          =   1815
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   3495
      Begin VB.TextBox txtServiceTime 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1800
         MaxLength       =   5
         TabIndex        =   7
         Tag             =   "HoursInDecimal"
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox txtActDate 
         Height          =   360
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   2
         Tag             =   "Date"
         Top             =   240
         Width           =   1080
      End
      Begin VB.CommandButton cmdActDate 
         Height          =   375
         Left            =   2985
         Picture         =   "AddActLog.frx":08AA
         Style           =   1  'Graphical
         TabIndex        =   3
         Tag             =   "Date"
         Top             =   240
         Width           =   375
      End
      Begin ecsTimeOCX.ecsTime timeActTime 
         Height          =   255
         Left            =   1785
         TabIndex        =   5
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Text            =   "02:13 PM"
         Appearance      =   1
         Object.TabStop         =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Service Fee Time: (Number of Hours) "
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   3255
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Time:"
         Height          =   255
         Index           =   4
         Left            =   960
         TabIndex        =   4
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Date:"
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   1
         Top             =   300
         Width           =   735
      End
   End
End
Attribute VB_Name = "AddActLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmClaim As frmClaim
Private mfrmActivityLog As frmActivityLog
Private moGUI As V2ECKeyBoard.clsCarGUI
Private msAssignmentsID As String

Private mbAdding As Boolean
Private msActLogID As String
Private mbLoading As Boolean

Public Property Let MyGUI(poGUI As V2ECKeyBoard.clsCarGUI)
    Set moGUI = poGUI
End Property
Public Property Set MyGUI(poGUI As V2ECKeyBoard.clsCarGUI)
    Set moGUI = poGUI
End Property
Public Property Get MyGUI() As V2ECKeyBoard.clsCarGUI
    Set MyGUI = moGUI
End Property

Public Property Let AssignmentsID(psAssignmentsID As String)
    msAssignmentsID = psAssignmentsID
End Property
Public Property Get AssignmentsID() As String
    AssignmentsID = msAssignmentsID
End Property

Public Property Let MyfrmClaim(poForm As Object)
    Set mfrmClaim = poForm
End Property
Public Property Set MyfrmClaim(poForm As Object)
    Set mfrmClaim = poForm
End Property
Public Property Get MyfrmClaim() As Object
    Set MyfrmClaim = mfrmClaim
End Property

Public Property Let MyActivityLog(poForm As Object)
    Set mfrmActivityLog = poForm
End Property
Public Property Set MyActivityLog(poForm As Object)
    Set mfrmActivityLog = poForm
End Property
Public Property Get MyActivityLog() As Object
    Set MyActivityLog = mfrmActivityLog
End Property

Public Property Let ActLogID(psActId As String)
    msActLogID = psActId
End Property

Public Property Let Adding(pbFlag As Boolean)
    mbAdding = pbFlag
End Property

Public Property Let Loading(pbFlag As Boolean)
    mbLoading = pbFlag
End Property
Public Property Get Loading() As Boolean
    Loading = mbLoading
End Property

Private Property Get msClassName() As String
    msClassName = Me.Name
End Property




Private Sub cmdActDate_Click()
    On Error GoTo EH
    
    mfrmClaim.ShowCalendar txtActDate
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdActDate_Click"
End Sub



Private Sub cmdExit_Click()
    On Error GoTo EH
    
    goUtil.utFormWinRegPos goUtil.gsMainAppEXEName, Me, True, , , , True
    Me.Visible = False
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdExit_Click"
End Sub

Private Sub cmdSave_Click()
    On Error GoTo EH
    
    SaveMe
    
    goUtil.utFormWinRegPos goUtil.gsMainAppEXEName, Me, True, , , , True
    Me.Visible = False
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdSave_Click"
End Sub

Private Sub cmdSpelling_Click()
    On Error GoTo EH
    
    goUtil.utLoadSP
    goUtil.goSP.CheckSP txtActText
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdSpelling_Click"
End Sub

Private Sub Form_Load()
    On Error GoTo EH
    'Set the Form Icon
    Me.Icon = mfrmClaim.optClaim(GuiClaimOptions.opt02_ActivityLog).Picture
    
    goUtil.utFormWinRegPos goUtil.gsMainAppEXEName, Me, , , , , True
    
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_Load"
End Sub

Public Function CLEANUP() As Boolean
    On Error GoTo EH

    Set mfrmClaim = Nothing
    Set mfrmActivityLog = Nothing
    Set moGUI = Nothing
    
    CLEANUP = True
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function CLEANUP"
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo EH
    
     Select Case UnloadMode
        Case vbFormControlMenu
            goUtil.utFormWinRegPos goUtil.gsMainAppEXEName, Me, True, , , , True
            Me.Visible = False
            Cancel = True
    End Select
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_QueryUnload"
End Sub

Private Sub timeActTime_time24HR(ps24HR As String)
    cmdSave.Enabled = True
End Sub

Private Sub txtActDate_Change()
    cmdSave.Enabled = True
End Sub

Private Sub txtActDate_GotFocus()
    goUtil.utSelText txtActDate
End Sub

Private Sub txtActDate_LostFocus()
    goUtil.utValidate , txtActDate
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    'RePos Controls
    'Width and Lefts
    framCommands.left = Me.Width - 3555
    framActText.Width = Me.Width - 240
    txtActText.Width = Me.Width - 480
    
    
    'Heights and Tops
    framActText.Height = Me.Height - 2385
    txtActText.Height = Me.Height - 2760
    
End Sub

Private Sub txtActText_Change()
    On Error GoTo EH
    Dim lPos As Long
    
    cmdSave.Enabled = True
    
'    If InStr(1, txtActText.Text, vbCrLf, vbBinaryCompare) > 0 Then
'        lPos = txtActText.SelStart
'        txtActText.Text = Replace(txtActText.Text, vbCrLf, vbNullString)
'        txtActText.SelStart = lPos
'    End If
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub txtActText_Change"
End Sub


Public Function SaveMe() As Boolean
    On Error GoTo EH
    Dim udtActLog As GuiActLogItem
    Dim oListView As ListView
    Dim itmX As ListItem
    
    'Validate some stuff first
    goUtil.utValidate Me
    'ADD
    If mbAdding Then
        With udtActLog
            .RTActivityLogID = "null"  ' not set here
            .AssignmentsID = "null"   ' not set here
            .BillingCountID = "null"   'not set here
            .ID = "null"   'not set here
            .IDAssignments = "null"   'not set here
            .IDBillingCount = "null"   'not set here
            .ActDate = txtActDate.Text
            .ActTime = txtActDate.Text & " " & timeActTime.time24HR
            .ServiceTime = txtServiceTime.Text
            .ActText = txtActText.Text
            .PageBreakAfter = "False"
            .BlankPageAfter = "False"
            .BlankRowsAfter = "0"
            .IsMgrEntry = "False"
            .IsDeleted = "False"
            .DownLoadMe = "False"
            .UpLoadMe = "True"
            .AdminComments = vbNullString
            .DateLastUpdated = Format(Now(), "MM/DD/YYYY HH:MM:SS")
            .UpdateByUserID = goUtil.gsCurUsersID
        End With
        
        'Add this entry
        mfrmActivityLog.AddActLogItem udtActLog
        mfrmActivityLog.RefreshActLog
        'Select the one just Added'
        Set oListView = mfrmActivityLog.lstvActLog
        For Each itmX In oListView.ListItems
            If itmX.SubItems(GuiActLogListView.DateLastUpdated - 1) = udtActLog.DateLastUpdated Then
                itmX.Selected = True
                itmX.EnsureVisible
                Exit For
            End If
        Next
    End If

    'EDIT
    If Not mbAdding Then

        With udtActLog
            'If editing need to Update the Selected ListItem as well
            Set oListView = mfrmActivityLog.lstvActLog
            Set itmX = oListView.SelectedItem
            .RTActivityLogID = itmX.SubItems(GuiActLogListView.RTActivityLogID - 1)
            .AssignmentsID = itmX.SubItems(GuiActLogListView.AssignmentsID - 1)
            .BillingCountID = itmX.SubItems(GuiActLogListView.BillingCountID - 1)
            .ID = itmX.SubItems(GuiActLogListView.ID - 1)
            .IDAssignments = itmX.SubItems(GuiActLogListView.IDAssignments - 1)
            .IDBillingCount = itmX.SubItems(GuiActLogListView.IDBillingCount - 1)
            itmX.Text = txtActDate.Text
            itmX.SubItems(GuiActLogListView.ActDateSort - 1) = Format(txtActDate.Text & " " & timeActTime.Text, "YYYY/MM/DD HH:MM")
            .ActDate = txtActDate.Text
            itmX.SubItems(GuiActLogListView.ActTime - 1) = timeActTime.time24HR
            itmX.SubItems(GuiActLogListView.ActTimeSort - 1) = Format(timeActTime.time24HR, "HH:MM")
            .ActTime = txtActDate.Text & " " & timeActTime.time24HR
            itmX.SubItems(GuiActLogListView.ServiceTime - 1) = txtServiceTime.Text
            .ServiceTime = txtServiceTime.Text
            itmX.SubItems(GuiActLogListView.ActText - 1) = txtActText.Text
            .ActText = txtActText.Text
            .PageBreakAfter = itmX.SubItems(GuiActLogListView.PageBreakAfter - 1)
            .BlankPageAfter = itmX.SubItems(GuiActLogListView.BlankPageAfter - 1)
            .BlankRowsAfter = itmX.SubItems(GuiActLogListView.BlankRowsAfter - 1)
            .IsMgrEntry = itmX.SubItems(GuiActLogListView.IsMgrEntry - 1)
            .IsDeleted = goUtil.GetFlagFromText(itmX.SubItems(GuiActLogListView.IsDeleted - 1))
            .DownLoadMe = goUtil.GetFlagFromText(itmX.SubItems(GuiActLogListView.DownLoadMe - 1))
            itmX.ListSubItems(GuiActLogListView.UpLoadMe - 1).ReportIcon = GuiActLogStatusList.UpLoadMe
            .UpLoadMe = True
            .AdminComments = itmX.SubItems(GuiActLogListView.AdminComments - 1)
            .DateLastUpdated = Format(Now(), "MM/DD/YYYY HH:MM:SS")
            itmX.SubItems(GuiActLogListView.DateLastUpdated - 1) = .DateLastUpdated
            itmX.SubItems(GuiActLogListView.DateLastUpdatedSort - 1) = Format(.DateLastUpdated, "YYYY/MM/DD HH:MM:SS")
            itmX.SubItems(GuiActLogListView.UpdateByUserID - 1) = goUtil.gsCurUsersID
            .UpdateByUserID = itmX.SubItems(GuiActLogListView.UpdateByUserID - 1)
        End With
        
        'Edit this entry
        mfrmActivityLog.EditActLogItem udtActLog
        oListView.SortKey = GuiActLogListView.ActDate
        oListView.Sorted = True
        
        'now be sure its visible
        Set itmX = oListView.SelectedItem
        itmX.EnsureVisible
        
    End If

    SaveMe = True
    
    'cleanup
    Set oListView = Nothing
    Set itmX = Nothing

    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function SaveMe"
End Function

Private Sub txtActText_GotFocus()
    goUtil.utSelText txtActText
End Sub

Private Sub txtActText_LostFocus()
    goUtil.utValidate , txtActText
End Sub

Private Sub txtServiceTime_Change()
    cmdSave.Enabled = True
End Sub


Private Sub txtServiceTime_GotFocus()
    goUtil.utSelText txtServiceTime
End Sub

Private Sub txtServiceTime_LostFocus()
    goUtil.utValidate , txtServiceTime
End Sub
