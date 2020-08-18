VERSION 5.00
Begin VB.Form EditAttach 
   AutoRedraw      =   -1  'True
   Caption         =   "Attachment"
   ClientHeight    =   3840
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8415
   Icon            =   "EditAttach.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   8415
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame framCommands 
      Height          =   1215
      Left            =   4920
      TabIndex        =   9
      Top             =   0
      Width           =   3375
      Begin VB.CommandButton cmdSpelling 
         Caption         =   "Spellin&g"
         Height          =   855
         Left            =   120
         MaskColor       =   &H00000000&
         Picture         =   "EditAttach.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Exit"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Sa&ve"
         Default         =   -1  'True
         Height          =   855
         Left            =   1200
         MaskColor       =   &H00000000&
         Picture         =   "EditAttach.frx":0456
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Exit"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   855
         Left            =   2280
         MaskColor       =   &H00000000&
         Picture         =   "EditAttach.frx":05A0
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Exit"
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame framDateName 
      Caption         =   "Attachment Date && Name"
      Height          =   1215
      Left            =   60
      TabIndex        =   1
      Top             =   0
      Width           =   4695
      Begin VB.TextBox txtAttachName 
         Height          =   360
         Left            =   720
         MaxLength       =   50
         TabIndex        =   6
         Tag             =   "FileOrFolderName"
         Top             =   720
         Width           =   3840
      End
      Begin VB.CommandButton cmdAttachDate 
         Height          =   375
         Left            =   1905
         Picture         =   "EditAttach.frx":08AA
         Style           =   1  'Graphical
         TabIndex        =   4
         Tag             =   "Date"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtAttachDate 
         Height          =   360
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   3
         Tag             =   "Date"
         Top             =   240
         Width           =   1080
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Name:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   780
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Date:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   300
         Width           =   495
      End
   End
   Begin VB.Frame framDescription 
      Caption         =   "Description ( - 254 character limit -): "
      Height          =   2415
      Left            =   60
      TabIndex        =   7
      Top             =   1320
      Width           =   8295
      Begin VB.ListBox lstDescription 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1980
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Visible         =   0   'False
         Width           =   8055
      End
      Begin VB.TextBox txtDescription 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2040
         Left            =   120
         MaxLength       =   254
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   240
         Width           =   8055
      End
   End
End
Attribute VB_Name = "EditAttach"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmClaim As frmClaim
Private mfrmAttachments As frmAttachments
Private moGUI As V2ECKeyBoard.clsCarGUI
Private msAssignmentsID As String
Private mbLoadingDesc As Boolean

Private msAttachID As String
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

Public Property Let MyAttachments(poForm As Object)
    Set mfrmAttachments = poForm
End Property
Public Property Set MyAttachments(poForm As Object)
    Set mfrmAttachments = poForm
End Property
Public Property Get MyAttachments() As Object
    Set MyAttachments = mfrmAttachments
End Property

Public Property Let AttachID(psID As String)
    msAttachID = psID
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


Private Sub cmdAttachDate_Click()
    On Error GoTo EH
    
    mfrmClaim.ShowCalendar txtAttachDate
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdAttachDate_Click"
End Sub



Private Sub cmdExit_Click()
    On Error GoTo EH
    Dim sLossFormat As String
    
    sLossFormat = goUtil.IsNullIsVbNullString(mfrmClaim.adoRSAssignments.Fields("LRFormat"))
    If StrComp(sLossFormat, "V2ECcarFarmers.clsLossXML01", vbTextCompare) = 0 Then
        If Trim(txtDescription.Text) = vbNullString Then
            MsgBox "You must select a description for this item!", vbExclamation, "Description Required"
            Exit Sub
        Else
            If cmdSave.Enabled Then
                SaveMe
            End If
        End If
    End If
    goUtil.utFormWinRegPos goUtil.gsMainAppEXEName, Me, True, , , , True
    Me.Visible = False
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdExit_Click"
End Sub

Private Sub cmdSave_Click()
    On Error GoTo EH
    Dim sLossFormat As String
    
    sLossFormat = goUtil.IsNullIsVbNullString(mfrmClaim.adoRSAssignments.Fields("LRFormat"))
    If StrComp(sLossFormat, "V2ECcarFarmers.clsLossXML01", vbTextCompare) = 0 Then
        If Trim(txtDescription.Text) = vbNullString Then
            MsgBox "You must select a description for this item!", vbExclamation, "Description Required"
            Exit Sub
        Else
            If cmdSave.Enabled Then
                SaveMe
            End If
        End If
    Else
        SaveMe
    End If
    
    goUtil.utFormWinRegPos goUtil.gsMainAppEXEName, Me, True, , , , True
    Me.Visible = False
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdSave_Click"
End Sub

Private Sub cmdSpelling_Click()
    On Error GoTo EH
    
    goUtil.utLoadSP
    goUtil.goSP.CheckSP txtDescription
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdSpelling_Click"
End Sub

Private Sub Form_Load()
    On Error GoTo EH
    'Set the Form Icon
    Me.Icon = mfrmClaim.optClaim(GuiClaimOptions.opt07_Attachments).Picture
    
    goUtil.utFormWinRegPos goUtil.gsMainAppEXEName, Me, , , , , True
    
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_Load"
End Sub

Public Sub LoadDescriptionList()
    On Error GoTo EH
    Dim sDescList As String
    Dim saryList() As String
    Dim lCount As Long
    Dim sSelItem As String
    Dim lSelIndex As Long
    Dim sLossFormat As String
    
    sLossFormat = goUtil.IsNullIsVbNullString(mfrmClaim.adoRSAssignments.Fields("LRFormat"))
    If StrComp(sLossFormat, "V2ECcarFarmers.clsLossXML01", vbTextCompare) = 0 Then
        cmdSpelling.Enabled = False
        lstDescription.Visible = True
        txtDescription.Visible = False
        sDescList = GetSetting(goUtil.gsMainAppEXEName, "V2ECcarFarmers.clsLossXML01", "XML01_DOCTYPES", vbNullString)
    Else
        Exit Sub
    End If
    
    mbLoadingDesc = True
    
    sSelItem = Trim(txtDescription.Text)
    lSelIndex = -1
    
    If sDescList <> vbNullString Then
        saryList() = Split(sDescList, "|")
        'BUbble sort this
        goUtil.utBubbleSort saryList
        For lCount = LBound(saryList, 1) To UBound(saryList, 1)
            lstDescription.AddItem saryList(lCount)
            If sSelItem <> vbNullString Then
                If StrComp(lstDescription.List(lstDescription.NewIndex), sSelItem, vbTextCompare) = 0 Then
                    lSelIndex = lstDescription.NewIndex
                End If
            End If
        Next
    Else
        lstDescription.Visible = False
        txtDescription.Visible = True
    End If
    
    lstDescription.ListIndex = lSelIndex
    
    If lstDescription.ListIndex = -1 Then
        txtDescription = vbNullString
    End If
    
    mbLoadingDesc = False
   
    Exit Sub
EH:
    mbLoadingDesc = False
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub LoadDescriptionList"
End Sub

Public Function CLEANUP() As Boolean
    On Error GoTo EH

    Set mfrmClaim = Nothing
    Set mfrmAttachments = Nothing
    Set moGUI = Nothing
    
    CLEANUP = True
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function CLEANUP"
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo EH
    Dim sLossFormat As String
    
    sLossFormat = goUtil.IsNullIsVbNullString(mfrmClaim.adoRSAssignments.Fields("LRFormat"))
    
     Select Case UnloadMode
        Case vbFormControlMenu
            If StrComp(sLossFormat, "V2ECcarFarmers.clsLossXML01", vbTextCompare) = 0 Then
                If Trim(txtDescription.Text) = vbNullString Then
                    MsgBox "You must select a description for this item!", vbExclamation, "Description Required"
                    Cancel = True
                    Exit Sub
                Else
                    If cmdSave.Enabled Then
                        SaveMe
                    End If
                End If
            End If
            goUtil.utFormWinRegPos goUtil.gsMainAppEXEName, Me, True, , , , True
            Me.Visible = False
            Cancel = True
    End Select
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_QueryUnload"
End Sub


Private Sub lstDescription_Click()
    On Error GoTo EH
    
    If mbLoadingDesc Then
        Exit Sub
    End If
    
    If lstDescription.ListIndex > -1 Then
        txtDescription.Text = lstDescription.Text
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lstDescription_Click"
End Sub

Private Sub lstDescription_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case KeyCodeConstants.vbKeyReturn
            If lstDescription.Enabled And lstDescription.Visible Then
                lstDescription_Click
            End If
    End Select
End Sub


Private Sub lstDescription_DblClick()
    On Error GoTo EH
    
    If lstDescription.ListIndex > -1 Then
        txtDescription.Text = lstDescription.Text
        SaveMe
        goUtil.utFormWinRegPos goUtil.gsMainAppEXEName, Me, True, , , , True
        Me.Visible = False
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lstDescription_DblClick"
End Sub

Private Sub txtAttachDate_Change()
    cmdSave.Enabled = True
End Sub

Private Sub txtAttachDate_GotFocus()
    goUtil.utSelText txtAttachDate
End Sub

Private Sub txtAttachDate_LostFocus()
    goUtil.utValidate , txtAttachDate
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    'RePos Controls
    'Width and Lefts
    framCommands.left = Me.Width - 3555
    framDescription.Width = Me.Width - 240
    txtDescription.Width = Me.Width - 480
    lstDescription.Width = Me.Width - 480
    
    
    'Heights and Tops
    framDescription.Height = Me.Height - 1830
    txtDescription.Height = Me.Height - 2205
    lstDescription.Height = Me.Height - 2205
End Sub

Private Sub txtDescription_Change()
    On Error GoTo EH
    cmdSave.Enabled = True
    Dim lPos As Long
    If InStr(1, txtDescription.Text, vbCrLf, vbBinaryCompare) > 0 Then
        lPos = txtDescription.SelStart
        txtDescription.Text = Replace(txtDescription.Text, vbCrLf, vbNullString)
        txtDescription.SelStart = lPos
    End If
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub txtDescription_Change"
End Sub


Public Function SaveMe() As Boolean
    On Error GoTo EH
    Dim udtAttach As GuiAttachItem
    Dim oListView As ListView
    Dim itmX As ListItem
    
   ' Validate some stuff first
    goUtil.utValidate Me

    With udtAttach
        Set oListView = mfrmAttachments.lstvAttachments
        Set itmX = oListView.SelectedItem
        .RTAttachmentsID = itmX.SubItems(GuiAttachListView.RTAttachmentsID - 1)
        .AssignmentsID = itmX.SubItems(GuiAttachListView.AssignmentsID - 1)
        .ID = itmX.SubItems(GuiAttachListView.ID - 1)
        .IDAssignments = itmX.SubItems(GuiAttachListView.IDAssignments - 1)
        itmX.Text = txtAttachDate.Text
        .AttachDate = itmX.Text
        .SortOrder = itmX.SubItems(GuiAttachListView.SortOrder - 1)
        'Set the text to the corrected text
        itmX.SubItems(GuiAttachListView.Description - 1) = txtDescription.Text
        .Description = txtDescription.Text
        itmX.SubItems(GuiAttachListView.AttachName - 1) = Replace(txtAttachName.Text, "_", Chr(32), , , vbBinaryCompare)
        .AttachName = itmX.SubItems(GuiAttachListView.AttachName - 1)
        .Attachment = itmX.SubItems(GuiAttachListView.Attachment - 1)
        .DownloadAttachment = goUtil.GetFlagFromText(itmX.SubItems(GuiAttachListView.DownloadAttachment - 1))
        .UploadAttachment = goUtil.GetFlagFromText(itmX.SubItems(GuiAttachListView.UploadAttachment - 1))
        .IsDeleted = goUtil.GetFlagFromText(itmX.SubItems(GuiAttachListView.IsDeleted - 1))
        .DownLoadMe = goUtil.GetFlagFromText(itmX.SubItems(GuiAttachListView.DownLoadMe - 1))
        itmX.ListSubItems(GuiAttachListView.UpLoadMe - 1).ReportIcon = GuiAttachStatusList.UpLoadMe
        .UpLoadMe = "True"
        .AdminComments = itmX.SubItems(GuiAttachListView.AdminComments - 1)
        .DateLastUpdated = Format(Now(), "MM/DD/YYYY HH:MM:SS")
        itmX.SubItems(GuiAttachListView.DateLastUpdated - 1) = .DateLastUpdated
        itmX.SubItems(GuiAttachListView.DateLastUpdatedSort - 1) = Format(.DateLastUpdated, "YYYY/MM/DD HH:MM:SS")
        .UpdateByUserID = goUtil.gsCurUsersID
    End With

    'Edit this entry
    mfrmAttachments.EditAttachmentItem udtAttach
    oListView.SortKey = GuiAttachListView.SortOrder
    oListView.Sorted = True

    'now be sure its visible
    Set itmX = oListView.SelectedItem
    itmX.EnsureVisible
    itmX.Selected = False
    SaveMe = True
    
    'cleanup
    Set oListView = Nothing
    Set itmX = Nothing

    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function SaveMe"
End Function

Private Sub txtDescription_GotFocus()
    goUtil.utSelText txtDescription
End Sub

Private Sub txtDescription_LostFocus()
    goUtil.utValidate , txtDescription
End Sub

Private Sub txtAttachName_Change()
    cmdSave.Enabled = True
End Sub

Private Sub txtAttachName_GotFocus()
    goUtil.utSelText txtAttachName
End Sub

Private Sub txtAttachName_LostFocus()
    goUtil.utValidate , txtAttachName
End Sub

