VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmClaimsList 
   AutoRedraw      =   -1  'True
   Caption         =   "Claims List View"
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8280
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmClaimsList.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6975
   ScaleWidth      =   8280
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtToolTip 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1320
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   6000
      Width           =   3615
   End
   Begin VB.CommandButton cmdRequestApproval 
      Caption         =   "Request Appr&oval"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5040
      MaskColor       =   &H00000000&
      Picture         =   "frmClaimsList.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Exit"
      Top             =   6000
      Width           =   975
   End
   Begin VB.TextBox txtParamValue 
      Height          =   375
      Left            =   360
      MaxLength       =   300
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   5160
      Visible         =   0   'False
      Width           =   5655
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   6000
      Width           =   1100
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   855
      Left            =   7200
      MaskColor       =   &H00000000&
      Picture         =   "frmClaimsList.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Exit"
      Top             =   6000
      Width           =   975
   End
   Begin VB.CommandButton cmdPrintList 
      Caption         =   "Sho&w Item List"
      Height          =   855
      Left            =   6120
      MaskColor       =   &H00000000&
      Picture         =   "frmClaimsList.frx":0B8E
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Exit"
      Top             =   6000
      Width           =   975
   End
   Begin VB.Frame framAssignments 
      Appearance      =   0  'Flat
      Caption         =   "Assignments"
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
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8055
      Begin VB.CheckBox chkViewGrid 
         Caption         =   "&Grid OFF"
         Height          =   375
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   480
         Width           =   1100
      End
      Begin VB.CommandButton cmdFindNext 
         Caption         =   "Find &Next"
         Height          =   375
         Left            =   3360
         TabIndex        =   5
         Top             =   480
         Width           =   1100
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Find"
         Height          =   375
         Left            =   2280
         TabIndex        =   4
         Top             =   480
         Width           =   1100
      End
      Begin VB.CommandButton cmdSelectAll 
         Caption         =   "&Select All"
         Height          =   375
         Left            =   1200
         TabIndex        =   3
         Top             =   480
         Width           =   1100
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   1100
      End
      Begin VB.CheckBox chkAllowItemUpdates 
         Caption         =   "Enable Item updates from this screen"
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   200
         Width           =   5535
      End
      Begin VB.ComboBox cboSelStatus 
         Height          =   360
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   840
         Width           =   7815
      End
      Begin VB.Timer Timer_UnloadClaim 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   7320
         Top             =   1800
      End
      Begin MSComctlLib.ImageList imgAssignmentsList 
         Left            =   7320
         Top             =   1320
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483633
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   16777215
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   8
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClaimsList.frx":0FD0
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClaimsList.frx":1422
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClaimsList.frx":14C8
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClaimsList.frx":18D0
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClaimsList.frx":1CBC
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClaimsList.frx":1DAA
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClaimsList.frx":1E82
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClaimsList.frx":2191
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lvwAssignments 
         Height          =   4455
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   7858
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "imgAssignmentsList"
         SmallIcons      =   "imgAssignmentsList"
         ColHdrIcons     =   "imgAssignmentsList"
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.CheckBox chkHideDeleted 
         Caption         =   "Sho&w Deleted"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   480
         Visible         =   0   'False
         Width           =   1100
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete / Undelete"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6120
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   480
         Visible         =   0   'False
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmClaimsList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private msFindText As String
Private mlLastFindIndex As Long
Private mbFormLoading As Boolean
Private mitmXSelected As ListItem 'Currently selected Claim Item
Private msAssignmentsID As String
Private madoRSAssignments As ADODB.Recordset
Private mArv As V2ARViewer.clsARViewer
Private moForm As Form
Private mbUnloadingClaim As Boolean 'Set to true just before the Timer_UnloadClaim is enabled
Private mlSelStatusID As Long
Private mMyfrmClaim As frmClaim  ' Claim Form
Private moGUI As V2ECKeyBoard.clsCarGUI

' In General Declarations
Private Const LVM_SUBITEMASSGN As Long = (LVM_FIRST + 57)
Private Const LVHT_ONITEMICON As Long = &H2
Private Const LVHT_ONITEMLABEL As Long = &H4
Private Const LVHT_ONITEMSTATEICON As Long = &H8
Private Const LVHT_ONITEM As Long = (LVHT_ONITEMICON Or _
                                    LVHT_ONITEMLABEL Or _
                                    LVHT_ONITEMSTATEICON)
Private Type POINTAPI
  X As Long
  Y As Long
End Type

Private Type LVASSGNINFO
   pt As POINTAPI
   flags As Long
   iItem As Long
   iSubItem  As Long
End Type

Dim mlX As Single, mlY As Single

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long

'Assignemnts RS
Public Property Let adoRSAssignments(padoRS As ADODB.Recordset)
    Set madoRSAssignments = padoRS
End Property
Public Property Set adoRSAssignments(padoRS As ADODB.Recordset)
    Set madoRSAssignments = padoRS
End Property
Public Property Get adoRSAssignments() As ADODB.Recordset
    Set adoRSAssignments = madoRSAssignments
End Property

Private Sub chkAllowItemUpdates_Click()
    On Error GoTo EH
    Dim sMess As String
    
    If chkAllowItemUpdates.Value = vbChecked Then
        sMess = "Are you sure you want to " & chkAllowItemUpdates.Caption & "?"
        If MsgBox(sMess, vbQuestion + vbYesNo, chkAllowItemUpdates.Caption) = vbNo Then
            chkAllowItemUpdates.Value = vbUnchecked
RESET_HEADERICONS:
            With lvwAssignments
                'LossDate
                .ColumnHeaders.Item(GuiAssignments.LossDate).Icon = Empty
                'AssignedDate
                .ColumnHeaders.Item(GuiAssignments.AssignedDate).Icon = Empty
                'ReceivedDate
                .ColumnHeaders.Item(GuiAssignments.ReceivedDate).Icon = Empty
                'ContactDate
                .ColumnHeaders.Item(GuiAssignments.ContactDate).Icon = Empty
                'InspectedDate
                .ColumnHeaders.Item(GuiAssignments.InspectedDate).Icon = Empty
                '10.28.2005 BGS  Per Rob Petrovics Request... Close Dates for All Profiles
                'will no longer be updateable by the adjuster, only a manager or admin may
                'update the closed date.
'                'CloseDate
'                .ColumnHeaders.Item(GuiAssignments.CloseDate).Icon = Empty
            End With
        Else
            With lvwAssignments
                'LossDate
                .ColumnHeaders.Item(GuiAssignments.LossDate).Icon = GuiAssignmentsPic.CalandarPic
                'AssignedDate
                .ColumnHeaders.Item(GuiAssignments.AssignedDate).Icon = GuiAssignmentsPic.CalandarPic
                'ReceivedDate
                .ColumnHeaders.Item(GuiAssignments.ReceivedDate).Icon = GuiAssignmentsPic.CalandarPic
                'ContactDate
                .ColumnHeaders.Item(GuiAssignments.ContactDate).Icon = GuiAssignmentsPic.CalandarPic
                'InspectedDate
                .ColumnHeaders.Item(GuiAssignments.InspectedDate).Icon = GuiAssignmentsPic.CalandarPic
                 '10.28.2005 BGS  Per Rob Petrovics Request... Close Dates for All Profiles
                'will no longer be updateable by the adjuster, only a manager or admin may
                'update the closed date.
'                'CloseDate
'                .ColumnHeaders.Item(GuiAssignments.CloseDate).Icon = GuiAssignmentsPic.CalandarPic
            End With
        End If
    Else
        GoTo RESET_HEADERICONS
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub chkAllowItemUpdates_Click"
End Sub

Private Sub lvwAssignments_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   mlX = X
   mlY = Y
End Sub


Public Property Let UnloadingClaim(pbFlag As Boolean)
    mbUnloadingClaim = pbFlag
End Property
Public Property Get UnloadingClaim() As Boolean
    UnloadingClaim = mbUnloadingClaim
End Property

Public Property Let MyGUI(poGUI As V2ECKeyBoard.clsCarGUI)
    Set moGUI = poGUI
End Property
Public Property Set MyGUI(poGUI As V2ECKeyBoard.clsCarGUI)
    Set moGUI = poGUI
End Property
Public Property Get MyGUI() As V2ECKeyBoard.clsCarGUI
    Set MyGUI = moGUI
End Property

Public Property Get MyfrmClaim() As Object
    Set MyfrmClaim = mMyfrmClaim
End Property

Public Property Let itmXSelected(pitmX As ListItem)
    On Error GoTo EH
    Set mitmXSelected = pitmX
    PopulateFrmCaptionAssignmentInfo Me, "Claims List View"
    If Not mitmXSelected Is Nothing Then
        cmdEdit.Enabled = True
        cmdRequestApproval.Enabled = True
        msAssignmentsID = mitmXSelected.ListSubItems(GuiAssignments.RKey - 1)
    Else
        cmdEdit.Enabled = False
        cmdRequestApproval.Enabled = False
        msAssignmentsID = "0"
    End If
    Exit Property
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Property Let itmXSelected"
End Property
Public Property Set itmXSelected(pitmX As ListItem)
    On Error GoTo EH
    Set mitmXSelected = pitmX
    PopulateFrmCaptionAssignmentInfo Me, "Claims List View"
    If Not mitmXSelected Is Nothing Then
        cmdEdit.Enabled = True
        cmdRequestApproval.Enabled = True
        msAssignmentsID = mitmXSelected.ListSubItems(GuiAssignments.RKey - 1)
    Else
        cmdEdit.Enabled = False
        cmdRequestApproval.Enabled = False
        msAssignmentsID = "0"
    End If
    Exit Property
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Property Set itmXSelected"
End Property
Public Property Get itmXSelected() As ListItem
    Set itmXSelected = mitmXSelected
End Property

Private Property Get msClassName() As String
    msClassName = Me.Name
End Property

Private Sub cboSelStatus_Click()
    On Error GoTo EH
    Dim lNewStatusID As Long
    
    If mbFormLoading Then
        Exit Sub
    End If
    
    
    lNewStatusID = cboSelStatus.ItemData(cboSelStatus.ListIndex)
    
    If lNewStatusID <> mlSelStatusID Then
        Set mitmXSelected = Nothing
        mlSelStatusID = lNewStatusID
        RefreshMe
    End If

    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cboSelStatus_Click"
End Sub

Private Sub chkHideDeleted_Click()
    On Error GoTo EH
    Dim bHideDeleted As Boolean
    
    If chkHideDeleted.Value = vbChecked Then
        chkHideDeleted.Caption = "Hide &Deleted"
        bHideDeleted = True
    Else
        chkHideDeleted.Caption = "Show &Deleted"
        bHideDeleted = False
    End If
    
    SaveSetting App.EXEName, "GENERAL", "HIDE_DELETED", bHideDeleted
    If Not mbFormLoading Then
        Populatelvw
    End If
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub chkHideDeleted_Click"
End Sub

Private Sub chkViewGrid_Click()
    On Error GoTo EH
    Dim bGridOn As Boolean
    
    If chkViewGrid.Value = vbChecked Then
        chkViewGrid.Caption = "&Grid ON"
        bGridOn = True
    Else
        chkViewGrid.Caption = "&Grid OFF"
        bGridOn = False
    End If
    
    SaveSetting App.EXEName, "GENERAL", "GRID_ON", bGridOn
    lvwAssignments.GridLines = bGridOn
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub chkViewGrid_Click"
End Sub

Private Sub cmdDelete_Click()
    On Error GoTo EH
    Dim itmX As ListItem
    Dim sMess As String
    Dim oConn As New ADODB.Connection
    Dim sSQL As String
    Dim sAssIDIN As String
    Dim bSelected As Boolean 'if at least one record is selected
    Dim lRecordsAffected As Long
    Dim sDate As String
    
    sMess = "Are you sure you want to delete / undelete the selected record(s) ?"
    If MsgBox(sMess, vbQuestion + vbYesNo, "Delete Records") = vbNo Then
        Exit Sub
    End If
    
    'If deleting or undeleting need to reset the currently selceted Item
    Set itmXSelected = Nothing
    
    'Set the data source
    Set oConn = New ADODB.Connection
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    'Loop through the Listview and Mark as deleted onything selected
    
    sMess = vbNullString
    For Each itmX In lvwAssignments.ListItems
        If itmX.Selected Then
            bSelected = True
            'If the Currently selected Claim Item
            'is selected for deletion then set it to nothing
            
            'check for Locked items
            If goUtil.GetFlagFromText(itmX.SubItems(GuiAssignments.Islocked - 1)) Then
                sMess = sMess & itmX.SubItems(GuiAssignments.IBNUM - 1) & " "
                sMess = sMess & "Is Locked!" & vbCrLf
                GoTo NEXT_ITEM
            End If
            'Check for Reassigned Items
            sDate = itmX.SubItems(GuiAssignments.DateReassigned - 1)
            If IsDate(sDate) Then
                If CDate(sDate) <> NULL_DATE And CDate(sDate) <> CDate("1/1/1900") Then
                    sMess = sMess & itmX.SubItems(GuiAssignments.IBNUM - 1) & " "
                    sMess = sMess & "Has Been Reassigned!" & vbCrLf
                    GoTo NEXT_ITEM
                End If
            End If

            If sAssIDIN = vbNullString Then
                sAssIDIN = itmX.SubItems(GuiAssignments.RKey - 1)
            Else
                sAssIDIN = sAssIDIN & ", "
                sAssIDIN = sAssIDIN & itmX.SubItems(GuiAssignments.RKey - 1)
            End If
            
        End If
NEXT_ITEM:
    Next
    
    If bSelected Then
        If sAssIDIN <> vbNullString Then
            sSQL = "UPDATE Assignments SET "
            sSQL = sSQL & "IsDeleted = IIF(IsDeleted = True, False,True), "
            sSQL = sSQL & "UpLoadMe = True, "
            sSQL = sSQL & "DateLastUpdated = #" & Now() & "#, "
            sSQL = sSQL & "UpdateByUserID = " & goUtil.gsCurUsersID & " "
            sSQL = sSQL & "WHERE    AssignmentsID IN (" & sAssIDIN & ") "
            
            oConn.Execute sSQL, lRecordsAffected
        End If
        
        sMess = "Deleted / Undeleted " & lRecordsAffected & " Record(s)." & vbCrLf & vbCrLf & sMess
        MsgBox sMess, vbExclamation + vbOKOnly, "Delete / Undelete Record(s)"
        cmdRefresh_Click
    End If
    
    'cleanup
    Set oConn = Nothing
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdDelete_Click"
End Sub

Private Sub cmdEdit_Click()
    ShowClaim
End Sub

Private Sub cmdExit_Click()
    On Error GoTo EH
    Dim sMess As String
    
    cmdExit.Enabled = False
    CLEANUP
    Unload Me
    Exit Sub
    
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdExit_Click"
End Sub

Private Sub cmdFind_Click()
    On Error GoTo EH
    If lvwAssignments.ListItems.Count > 0 Then
        mlLastFindIndex = 0
        goUtil.utFindListItem Me, lvwAssignments, msFindText, mlLastFindIndex
    End If
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdFind_Click"
End Sub


Private Sub cmdFindNext_Click()
    On Error GoTo EH
    
    If msFindText <> vbNullString And lvwAssignments.ListItems.Count > 0 Then
        goUtil.utFindListItem Me, lvwAssignments, msFindText, mlLastFindIndex
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdFindNext_Click"
End Sub

Private Sub cmdPrintList_Click()
    On Error GoTo EH
    goUtil.utPrintListView App.EXEName, lvwAssignments, "Claims List"
     Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdPrintList_Click"
End Sub

Private Sub cmdRefresh_Click()
    On Error GoTo EH
    cmdRefresh.Enabled = False
    RefreshMe
    cmdRefresh.Enabled = True
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdRefresh_Click"
End Sub

Public Function RefreshMe() As Boolean
    On Error GoTo EH
    Dim sFindText As String
    Dim lLastFindIndex As Long
    
    Screen.MousePointer = MousePointerConstants.vbHourglass
    
    If Not mitmXSelected Is Nothing Then
        lLastFindIndex = 1
        sFindText = mitmXSelected.SubItems(GuiAssignments.IBNUM - 1)
    End If
    
    Populatelvw
    
    If sFindText <> vbNullString Then
        If cboSelStatus.ItemData(cboSelStatus.ListIndex) = -1 Then
            goUtil.utFindListItem Me, lvwAssignments, sFindText, lLastFindIndex
        End If
        Set itmXSelected = lvwAssignments.SelectedItem
    End If
    
    Screen.MousePointer = MousePointerConstants.vbDefault
    
    RefreshMe = True
    Exit Function
EH:
    Screen.MousePointer = MousePointerConstants.vbDefault
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function RefreshMe"
End Function

Private Sub cmdSelectAll_Click()
    On Error GoTo EH
    Dim itmX As ListItem
    For Each itmX In lvwAssignments.ListItems
        itmX.Selected = True
    Next
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdSelectAll_Click"
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    If Not mMyfrmClaim Is Nothing Then
        Me.Visible = False
    End If
End Sub


Private Sub Form_Load()
    On Error GoTo EH
    mbFormLoading = True
    Me.Visible = False
    mlSelStatusID = -1
    LoadHeader
    Populatelvw
    PopulateStatusList
    mbFormLoading = False
    Exit Sub
EH:
    mbFormLoading = False
   goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_Load"
End Sub

Private Sub Form_Paint()
    On Error Resume Next
    Screen.MousePointer = MousePointerConstants.vbDefault
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo EH
    Dim sMess As String
    
    Select Case UnloadMode
        Case vbFormControlMenu
            sMess = "Are you sure you want to Exit Claims List View ?"
'            If MsgBox(sMess, vbQuestion + vbOKCancel, "Exit Claims List View") <> vbOK Then
'                Cancel = True
'            Else
                CLEANUP
'            End If
    End Select
    
    
    Exit Sub
EH:
   goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_QueryUnload"
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState <> vbMinimized Then
        Me.Visible = False
    End If
    If Not mMyfrmClaim Is Nothing Then
        Me.Visible = False
    End If
    framAssignments.Width = Me.Width - 345
    framAssignments.Height = Me.Height - 1605
    cboSelStatus.Width = Me.Width - 615
    lvwAssignments.Width = Me.Width - 615
    lvwAssignments.Height = Me.Height - 2925
    cmdRefresh.top = Me.Height - 1380
    txtToolTip.top = Me.Height - 1380
    txtToolTip.Width = Me.Width - 4785
'    cmdEdit.top = Me.Height - 1380
'    chkAllowItemUpdates.top = Me.Height - 1380
    cmdExit.top = Me.Height - 1380
    cmdExit.left = Me.Width - 1230
    cmdPrintList.top = Me.Height - 1380
    cmdPrintList.left = Me.Width - 2310
    cmdRequestApproval.top = Me.Height - 1380
    cmdRequestApproval.left = Me.Width - 3360
    If mMyfrmClaim Is Nothing Then
        If Me.WindowState = vbMaximized Then
            If Not goUtil Is Nothing Then
                If Not goUtil.gfrmECTray Is Nothing Then
                    goUtil.gfrmECTray.Visible = False
                End If
            End If
            
        Else
            If Not goUtil Is Nothing Then
                If Not goUtil.gfrmECTray Is Nothing Then
                    goUtil.gfrmECTray.Visible = True
                End If
            End If
        End If
        If Me.WindowState <> vbMinimized Then
            Me.Visible = True
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo EH
    goUtil.utFormWinRegPos goUtil.gsMainAppEXEName, Me, True, , , , True
    If Not goUtil.gfrmECTray Is Nothing Then
        goUtil.gfrmECTray.ShowMe False
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_Unload"
End Sub

Private Sub LoadHeader()
    On Error GoTo EH
    Dim bGridOn As Boolean
    Dim bHideDeleted As Boolean
    Dim bHideUploadFlags As Boolean
    
    bHideDeleted = CBool(GetSetting(App.EXEName, "GENERAL", "HIDE_DELETED", True))
    bHideUploadFlags = CBool(GetSetting(App.EXEName, "GENERAL", "HIDE_UPLOAD_FLAGS", True))
    
    'set the columnheaders
    With lvwAssignments
        .ColumnHeaders.Add , "CLIENTNUM", "Claim Number"
        .ColumnHeaders.Add , "Insured", "Insured"
        .ColumnHeaders.Add , "LossDate", "Loss Date"
        .ColumnHeaders.Add , "LossDateSort", "Sort Loss Date"
        .ColumnHeaders.Add , "AssignedDate", "Assigned Date"
        .ColumnHeaders.Add , "AssignedDateSort", "Sort Assigned Date"
        .ColumnHeaders.Add , "ReceivedDate", "Received Date"
        .ColumnHeaders.Add , "ReceivedDateSort", "Sort Received Date"
        .ColumnHeaders.Add , "ContactDate", "Contact Date"
        .ColumnHeaders.Add , "ContactDateSort", "Sort Contact Date"
        .ColumnHeaders.Add , "InspectedDate", "Inspected Date"
        .ColumnHeaders.Add , "InspectedDateSort", "Sort Inspected Date"
        .ColumnHeaders.Add , "Status", "Status"
        .ColumnHeaders.Add , "CloseDate", "Close Date"
        .ColumnHeaders.Add , "CloseDateSort", "Sort Close Date"
        .ColumnHeaders.Add , "AssignmentType", "Type"
        .ColumnHeaders.Add , "IsLocked", "Is Locked"
        .ColumnHeaders.Add , "IsDeleted", "Is Deleted"
        .ColumnHeaders.Add , "UpLoadMe", "UpLoad Me"
        .ColumnHeaders.Add , "SP", "SP"
        .ColumnHeaders.Add , "CatName", "Cat Name"
        .ColumnHeaders.Add , "CatCode", "Cat Code"
        .ColumnHeaders.Add , "IBNUM", "IBNUM"
        .ColumnHeaders.Add , "MAStreet", "Mailing Street"
        .ColumnHeaders.Add , "MAStreeSort", "Sort Mailing Street"
        .ColumnHeaders.Add , "MACity", "City"
        .ColumnHeaders.Add , "MAState", "ST"
        .ColumnHeaders.Add , "MAZIP", "Zip"
        .ColumnHeaders.Add , "MAZIP4", "-0000"
        .ColumnHeaders.Add , "PAStreet", "Property Street"
        .ColumnHeaders.Add , "PAStreetSort", "Sort Property Street"
        .ColumnHeaders.Add , "PACity", "City"
        .ColumnHeaders.Add , "PAState", "ST"
        .ColumnHeaders.Add , "PAZIP", "Zip"
        .ColumnHeaders.Add , "PAZIP4", "-0000"
        .ColumnHeaders.Add , "Adjuster", "Adjuster"
        .ColumnHeaders.Add , "ACID", "ACID"
        .ColumnHeaders.Add , "DateReassigned", "Reassigned Date"
        .ColumnHeaders.Add , "DateReassignedSort", "Sort Reassigned Date"
        .ColumnHeaders.Add , "DateLastUpdated", "Date Last Updated"
        .ColumnHeaders.Add , "DateLastUpdatedSort", "Date Last Updated"
        .ColumnHeaders.Add , "AdminComments", "Admin Comments"
        .ColumnHeaders.Add , "RKey", "Key"
        .Sorted = False
        .SortOrder = lvwAscending
        
        'Assignment Type
        .ColumnHeaders.Item(GuiAssignments.AssignmentType).Width = 1200
        .ColumnHeaders.Item(GuiAssignments.AssignmentType).Alignment = lvwColumnLeft
        'Status
        .ColumnHeaders.Item(GuiAssignments.Status).Width = 1500
        .ColumnHeaders.Item(GuiAssignments.Status).Alignment = lvwColumnLeft
        'IsLocked
        .ColumnHeaders.Item(GuiAssignments.Islocked).Width = 400
        .ColumnHeaders.Item(GuiAssignments.Islocked).Alignment = lvwColumnCenter
        .ColumnHeaders.Item(GuiAssignments.Islocked).Icon = GuiAssignmentsPic.Islocked
        'IsDeleted
        .ColumnHeaders.Item(GuiAssignments.IsDeleted).Width = 400
        .ColumnHeaders.Item(GuiAssignments.IsDeleted).Alignment = lvwColumnCenter
        .ColumnHeaders.Item(GuiAssignments.IsDeleted).Icon = GuiAssignmentsPic.IsDeleted
        'UpLoadMe
        .ColumnHeaders.Item(GuiAssignments.UpLoadMe).Width = 400
        .ColumnHeaders.Item(GuiAssignments.UpLoadMe).Alignment = lvwColumnCenter
        .ColumnHeaders.Item(GuiAssignments.UpLoadMe).Icon = GuiAssignmentsPic.UpLoadMe
        'SP
        .ColumnHeaders.Item(GuiAssignments.SP).Width = 500
        .ColumnHeaders.Item(GuiAssignments.SP).Alignment = lvwColumnCenter
        'Cat Name
        .ColumnHeaders.Item(GuiAssignments.CatName).Width = 1500
        .ColumnHeaders.Item(GuiAssignments.CatName).Alignment = lvwColumnLeft
        'Cat Code
        .ColumnHeaders.Item(GuiAssignments.CatCode).Width = 1000
        .ColumnHeaders.Item(GuiAssignments.CatCode).Alignment = lvwColumnLeft
        'IBNUM
        .ColumnHeaders.Item(GuiAssignments.IBNUM).Width = 1500
        .ColumnHeaders.Item(GuiAssignments.IBNUM).Alignment = lvwColumnLeft
        'CLIENTNUM
        .ColumnHeaders.Item(GuiAssignments.CLIENTNUM).Width = 1700
        .ColumnHeaders.Item(GuiAssignments.CLIENTNUM).Alignment = lvwColumnLeft
        'Insured
        .ColumnHeaders.Item(GuiAssignments.Insured).Width = 3000
        .ColumnHeaders.Item(GuiAssignments.Insured).Alignment = lvwColumnLeft
        'MAStreet
        .ColumnHeaders.Item(GuiAssignments.MAStreet).Width = 2500
        .ColumnHeaders.Item(GuiAssignments.MAStreet).Alignment = lvwColumnLeft
        'MAStreetSort
        .ColumnHeaders.Item(GuiAssignments.MAStreetSort).Width = 0  'Hidden Sort Column
        .ColumnHeaders.Item(GuiAssignments.MAStreetSort).Alignment = lvwColumnLeft
        'MACity
        .ColumnHeaders.Item(GuiAssignments.MACity).Width = 1050
        .ColumnHeaders.Item(GuiAssignments.MACity).Alignment = lvwColumnLeft
        'MAState
        .ColumnHeaders.Item(GuiAssignments.MAState).Width = 550
        .ColumnHeaders.Item(GuiAssignments.MAState).Alignment = lvwColumnLeft
        'MAZIP
        .ColumnHeaders.Item(GuiAssignments.MAZIP).Width = 700
        .ColumnHeaders.Item(GuiAssignments.MAZIP).Alignment = lvwColumnLeft
        'MAZIP4
        .ColumnHeaders.Item(GuiAssignments.MAZIP4).Width = 700
        .ColumnHeaders.Item(GuiAssignments.MAZIP4).Alignment = lvwColumnLeft
        'PAStreet
        .ColumnHeaders.Item(GuiAssignments.PAStreet).Width = 2500
        .ColumnHeaders.Item(GuiAssignments.PAStreet).Alignment = lvwColumnLeft
        'PAStreetSort
        .ColumnHeaders.Item(GuiAssignments.PAStreetSort).Width = 0  'Hidden Sort Column
        .ColumnHeaders.Item(GuiAssignments.PAStreetSort).Alignment = lvwColumnLeft
        'PACity
        .ColumnHeaders.Item(GuiAssignments.PACity).Width = 1050
        .ColumnHeaders.Item(GuiAssignments.PACity).Alignment = lvwColumnLeft
        'PAState
        .ColumnHeaders.Item(GuiAssignments.PAState).Width = 550
        .ColumnHeaders.Item(GuiAssignments.PAState).Alignment = lvwColumnLeft
        'PAZIP
        .ColumnHeaders.Item(GuiAssignments.PAZIP).Width = 700
        .ColumnHeaders.Item(GuiAssignments.PAZIP).Alignment = lvwColumnLeft
        'PAZIP4
        .ColumnHeaders.Item(GuiAssignments.PAZIP4).Width = 700
        .ColumnHeaders.Item(GuiAssignments.PAZIP4).Alignment = lvwColumnLeft
        'Adjuster
        .ColumnHeaders.Item(GuiAssignments.Adjuster).Width = 1500
        .ColumnHeaders.Item(GuiAssignments.Adjuster).Alignment = lvwColumnLeft
        'ACID
        .ColumnHeaders.Item(GuiAssignments.ACID).Width = 1500
        .ColumnHeaders.Item(GuiAssignments.ACID).Alignment = lvwColumnLeft
        'LossDate
        .ColumnHeaders.Item(GuiAssignments.LossDate).Width = 1700
        .ColumnHeaders.Item(GuiAssignments.LossDate).Alignment = lvwColumnLeft
'        .ColumnHeaders.Item(GuiAssignments.LossDate).Icon = GuiAssignmentsPic.CalandarPic
        'LossDateSort
        .ColumnHeaders.Item(GuiAssignments.LossDateSort).Width = 0 'Hidden Sort Column
        .ColumnHeaders.Item(GuiAssignments.LossDateSort).Alignment = lvwColumnLeft
        'AssignedDate
        .ColumnHeaders.Item(GuiAssignments.AssignedDate).Width = 1700
        .ColumnHeaders.Item(GuiAssignments.AssignedDate).Alignment = lvwColumnLeft
'        .ColumnHeaders.Item(GuiAssignments.AssignedDate).Icon = GuiAssignmentsPic.CalandarPic
        'AssignedDateSort
        .ColumnHeaders.Item(GuiAssignments.AssignedDateSort).Width = 0 'Hidden Sort Column
        .ColumnHeaders.Item(GuiAssignments.AssignedDateSort).Alignment = lvwColumnLeft
        'ReceivedDate
        .ColumnHeaders.Item(GuiAssignments.ReceivedDate).Width = 1500
        .ColumnHeaders.Item(GuiAssignments.ReceivedDate).Alignment = lvwColumnLeft
'        .ColumnHeaders.Item(GuiAssignments.ReceivedDate).Icon = GuiAssignmentsPic.CalandarPic
        'ReceivedDateSort
        .ColumnHeaders.Item(GuiAssignments.ReceivedDateSort).Width = 0 'Hidden Sort Column
        .ColumnHeaders.Item(GuiAssignments.ReceivedDateSort).Alignment = lvwColumnLeft
        'ContactDate
        .ColumnHeaders.Item(GuiAssignments.ContactDate).Width = 1700
        .ColumnHeaders.Item(GuiAssignments.ContactDate).Alignment = lvwColumnLeft
'        .ColumnHeaders.Item(GuiAssignments.ContactDate).Icon = GuiAssignmentsPic.CalandarPic
        'ContactDateSort
        .ColumnHeaders.Item(GuiAssignments.ContactDateSort).Width = 0 'Hidden Sort Column
        .ColumnHeaders.Item(GuiAssignments.ContactDateSort).Alignment = lvwColumnLeft
        'InspectedDate
        .ColumnHeaders.Item(GuiAssignments.InspectedDate).Width = 1700
        .ColumnHeaders.Item(GuiAssignments.InspectedDate).Alignment = lvwColumnLeft
'        .ColumnHeaders.Item(GuiAssignments.InspectedDate).Icon = GuiAssignmentsPic.CalandarPic
        'InspectedDateSort
        .ColumnHeaders.Item(GuiAssignments.InspectedDateSort).Width = 0 'Hidden Sort Column
        .ColumnHeaders.Item(GuiAssignments.InspectedDateSort).Alignment = lvwColumnLeft
        'CloseDate
        .ColumnHeaders.Item(GuiAssignments.CloseDate).Width = 1700
        .ColumnHeaders.Item(GuiAssignments.CloseDate).Alignment = lvwColumnLeft
'        .ColumnHeaders.Item(GuiAssignments.CloseDate).Icon = GuiAssignmentsPic.CalandarPic
        'CloseDateSort
        .ColumnHeaders.Item(GuiAssignments.CloseDateSort).Width = 0 'Hidden Sort Column
        .ColumnHeaders.Item(GuiAssignments.CloseDateSort).Alignment = lvwColumnLeft
        'DateReassigned
        .ColumnHeaders.Item(GuiAssignments.DateReassigned).Width = 1200
        .ColumnHeaders.Item(GuiAssignments.DateReassigned).Alignment = lvwColumnLeft
        'DateReassignedSort
        .ColumnHeaders.Item(GuiAssignments.DateReassignedSort).Width = 0  'Hidden Sort Column
        .ColumnHeaders.Item(GuiAssignments.DateReassignedSort).Alignment = lvwColumnLeft
        'DateLastUpdated
        .ColumnHeaders.Item(GuiAssignments.DateLastUpdated).Width = 2200
        .ColumnHeaders.Item(GuiAssignments.DateLastUpdated).Alignment = lvwColumnLeft
        'DateLastUpdatedSort
        .ColumnHeaders.Item(GuiAssignments.DateLastUpdatedSort).Width = 0  'Hidden Sort Column
        .ColumnHeaders.Item(GuiAssignments.DateLastUpdatedSort).Alignment = lvwColumnLeft
        'AdminComments
        .ColumnHeaders.Item(GuiAssignments.AdminComments).Width = 10000
        .ColumnHeaders.Item(GuiAssignments.AdminComments).Alignment = lvwColumnLeft
        'Hidden RKey (Key will be AssignmentsID)
        .ColumnHeaders.Item(GuiAssignments.RKey).Width = 0
        .ColumnHeaders.Item(GuiAssignments.RKey).Alignment = lvwColumnLeft
    End With
    
    bGridOn = CBool(GetSetting(App.EXEName, "GENERAL", "GRID_ON", False))
    If bGridOn Then
        chkViewGrid.Value = vbChecked
    Else
        chkViewGrid.Value = vbUnchecked
    End If
    
    If bHideDeleted Then
        chkHideDeleted.Value = vbChecked
    Else
        chkHideDeleted.Value = vbUnchecked
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub LoadHeader"
End Sub

Private Sub Populatelvw()
    On Error GoTo EH
    Dim oConn As New ADODB.Connection
    Dim adoRS As ADODB.Recordset
    Dim sSQL As String
    Dim itmX As ListItem
    Dim iMyIcon As Long
    Dim sStatus As String
    Dim iMyStatus As V2ECKeyBoard.AssgnStatus
    Dim lCLOSED As Long
    Dim lDELETED As Long
    Dim lINTERIM As Long
    Dim lNEW As Long
    Dim lPENDING As Long
    Dim lREASSIGNED As Long
    Dim lREOPEN As Long
    Dim lUpLoadMe As Long
    Dim lStatusID As Long
    Dim sStatusTag As String
    Dim sStatusCaption As String
    Dim sStatusToolTip As String
    Dim lStatusAssignmentCount As Long
    Dim sFlagText As String
    Dim lSubCount As Long
    Dim sTemp As String
    Dim bDoToolTip As Boolean
    Dim lRED As Long
    Dim lBlue As Long
    Dim lGreen As Long
    Dim lLightyellow As Long
    Dim lBlack As Long
   
    'Clear Any Existing Items
    lvwAssignments.ListItems.Clear
   
    Set oConn = New ADODB.Connection
    Set adoRS = New ADODB.Recordset
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    'Need to Get Companies, not client Companies
    'that the DB User Name has access to...
    sSQL = "SELECT A.*, "
    sSQL = sSQL & "( "
    sSQL = sSQL & "SELECT UserName "
    sSQL = sSQL & "FROM USERS "
    sSQL = sSQL & "WHERE UsersID =  " & goUtil.gsCurUsersID & " "
    sSQL = sSQL & ") As AdjusterSpecUserName, "
    sSQL = sSQL & "( "
    sSQL = sSQL & "SELECT  ACID "
    sSQL = sSQL & "FROM ClientCoAdjusterSpec "
    sSQL = sSQL & "Where ClientCoAdjusterSpecID = A.[AdjusterSPecID] "
    sSQL = sSQL & ") As AdjusterSpecACID, "
    sSQL = sSQL & "( "
    sSQL = sSQL & "SELECT  ACIDDescription "
    sSQL = sSQL & "FROM ClientCoAdjusterSpec "
    sSQL = sSQL & "Where ClientCoAdjusterSpecID = A.[AdjusterSPecID] "
    sSQL = sSQL & ") As AdjusterSpecAcidDescription, "
    sSQL = sSQL & "( "
    sSQL = sSQL & "SELECT  Type "
    sSQL = sSQL & "FROM    AssignmentType "
    sSQL = sSQL & "WHERE   AssignmentTypeID = A.[AssignmentTypeID] "
    sSQL = sSQL & ") As AssignmentTypeType, "
    sSQL = sSQL & "( "
    sSQL = sSQL & "SELECT  CatCode "
    sSQL = sSQL & "FROM    ClientCompanyCatSpec "
    sSQL = sSQL & "WHERE   ClientCompanyCatSpecID = A.[ClientCompanyCatSpecID] "
    sSQL = sSQL & ") As ClientCompanyCatSpecCatCode, "
    sSQL = sSQL & "( "
    sSQL = sSQL & "SELECT Name "
    sSQL = sSQL & "FROM CAT "
    sSQL = sSQL & "WHERE CATID = " & goUtil.gsCurCat & " "
    sSQL = sSQL & ") As CatName, "
    sSQL = sSQL & "S.StatusAlias As Status, "
    '9.23.2005 Also Include the Count Of Rejected PackageItems + Count of Items marked to Send
    sSQL = sSQL & "(SELECT  Count(PKGI.[IsClientCoReject])  As CountClientReject "
    sSQL = sSQL & "FROM PackageItem PKGI "
    sSQL = sSQL & "WHERE PKGI.[AssignmentsID] = A.[AssignmentsID] "
    sSQL = sSQL & "AND PKGI.[IsDeleted] = 0 AND PKGI.[IsClientCoReject] = -1) As PKGIClientReject, "
    
    sSQL = sSQL & "(SELECT  Count(PKGI.[IsClientCoDelete])  As CountClientDelete "
    sSQL = sSQL & "FROM PackageItem PKGI "
    sSQL = sSQL & "WHERE PKGI.[AssignmentsID] = A.[AssignmentsID] "
    sSQL = sSQL & "AND PKGI.[IsDeleted] = 0 AND PKGI.[IsClientCoDelete] = -1) As PKGIClientDelete, "

    sSQL = sSQL & "(SELECT  Count(PKGI.[IsClientCoApprove])  As CountClientApprove "
    sSQL = sSQL & "FROM PackageItem PKGI "
    sSQL = sSQL & "WHERE PKGI.[AssignmentsID] = A.[AssignmentsID] "
    sSQL = sSQL & "AND PKGI.[IsDeleted] = 0 AND PKGI.[IsClientCoApprove] = -1) As PKGIClientApprove, "

    sSQL = sSQL & "(SELECT  Count(PKGI.[SendMe]) As CountSendMe "
    sSQL = sSQL & "FROM PackageItem PKGI "
    sSQL = sSQL & "WHERE PKGI.[AssignmentsID] = A.[AssignmentsID] "
    sSQL = sSQL & "AND PKGI.[IsDeleted] = 0 AND PKGI.[SendMe] = -1) As PKGISendMe, "
    sSQL = sSQL & "(SELECT TOP 1 IIF(IsNull(PKG.[AdminComments]),'',PKG.[AdminComments]) As retAdminComments "
    sSQL = sSQL & "FROM Package PKG "
    sSQL = sSQL & "WHERE PKG.[AssignmentsID] = A.[AssignmentsID]) As PKGAdminComments, "
    sSQL = sSQL & "CCCS.CatCode "
    sSQL = sSQL & "FROM (Assignments A "
    sSQL = sSQL & "INNER JOIN STATUS S ON A.StatusID = S.StatusID) "
    sSQL = sSQL & "INNER JOIN CLIENTCOMPANYCATSPEC CCCS ON (A.ClientCompanyCatSpecID = CCCS.ClientCompanyCatSpecID) "
    sSQL = sSQL & "WHERE A.ClientCompanyCatSpecID IN "
    sSQL = sSQL & "( "
    sSQL = sSQL & "SELECT  ClientCompanyCatSpecID "
    sSQL = sSQL & "FROM ClientCompanyCatSpec "
    sSQL = sSQL & "WHERE ClientCompanyID = " & goUtil.gsCurCar & " "
    sSQL = sSQL & "AND     CATID = " & goUtil.gsCurCat & " "
    sSQL = sSQL & ") "
    sSQL = sSQL & "AND A.AdjusterSpecID IN "
    sSQL = sSQL & "( "
    sSQL = sSQL & "SELECT  ClientCoAdjusterSpecID "
    sSQL = sSQL & "FROM ClientCoAdjusterSpec "
    sSQL = sSQL & "Where ClientCompanyID = " & goUtil.gsCurCar & " "
    sSQL = sSQL & "AND UsersID = " & goUtil.gsCurUsersID & " "
    sSQL = sSQL & ") "
    If CBool(GetSetting(App.EXEName, "GENERAL", "HIDE_DELETED", True)) Then
        sSQL = sSQL & "AND A.IsDeleted = False "
    End If
    If mlSelStatusID > -1 Then
        If mlSelStatusID = iAssignmentsStatus_REASSIGNED Then
            sSQL = sSQL & "AND A.REASSIGNED = True "
        ElseIf mlSelStatusID = iAssignmentsStatus_DELETED Then
            sSQL = sSQL & "AND A.IsDeleted = True "
        Else
            sSQL = sSQL & "AND A.StatusID = " & mlSelStatusID & " "
            sSQL = sSQL & "AND A.REASSIGNED = False "
        End If
    End If
    sSQL = sSQL & "ORDER BY A.CLIENTNUM "
    
    'Use Disconnected Record Set on asUseClient Cusor ONLY !
    adoRS.CursorLocation = adUseClient
    adoRS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set adoRS.ActiveConnection = Nothing
    
    If Not adoRS.EOF Then
        framAssignments.Caption = "Assignments Record Count (" & adoRS.RecordCount & ")"
        adoRS.MoveFirst
        Do Until adoRS.EOF
            
            If Not IsNull(adoRS!AssignmentsID) And Not IsNull(adoRS!CLIENTNUM) Then
                Set itmX = lvwAssignments.ListItems.Add(, """" & CStr(adoRS!AssignmentsID) & """", adoRS!CLIENTNUM)
            Else
                Exit Sub
            End If
            
            'Add Totals for diff status as go along
            If Not IsNull(adoRS!statusid) Then
                iMyStatus = adoRS!statusid
                'Check for separate Reassigned Flag
                If Not IsNull(adoRS!REASSIGNED) Then
                    If CBool(adoRS!REASSIGNED) Then
                        iMyStatus = iAssignmentsStatus_REASSIGNED
                    End If
                End If
                'Check for spearate Deleted flag
                If Not IsNull(adoRS!IsDeleted) Then
                    If CBool(adoRS!IsDeleted) Then
                        iMyStatus = iAssignmentsStatus_DELETED
                    End If
                End If
                Select Case iMyStatus
                    Case V2ECKeyBoard.AssgnStatus.iAssignmentsStatus_CLOSED
                        lCLOSED = lCLOSED + 1
                    Case V2ECKeyBoard.AssgnStatus.iAssignmentsStatus_DELETED
                        lDELETED = lDELETED + 1
                    Case V2ECKeyBoard.AssgnStatus.iAssignmentsStatus_INTERIM
                        lINTERIM = lINTERIM + 1
                    Case V2ECKeyBoard.AssgnStatus.iAssignmentsStatus_NEW
                        lNEW = lNEW + 1
                    Case V2ECKeyBoard.AssgnStatus.iAssignmentsStatus_PENDING
                        lPENDING = lPENDING + 1
                    Case V2ECKeyBoard.AssgnStatus.iAssignmentsStatus_REASSIGNED
                        lREASSIGNED = lREASSIGNED + 1
                    Case V2ECKeyBoard.AssgnStatus.iAssignmentsStatus_REOPEN
                        lREOPEN = lREOPEN + 1
                End Select
            End If
            'Also Keep track of how many Assignments need to be uploaded
            If Not IsNull(adoRS!UpLoadMe) Then
                If CBool(adoRS!UpLoadMe) Then
                    lUpLoadMe = lUpLoadMe + 1
                End If
            End If
            
            'Assignment Type
            
            If Not IsNull(adoRS!AssignmentTypeID) Then
                Select Case adoRS!AssignmentTypeID
                    Case AssgnType.iAssignmentType_Auto
                        iMyIcon = GuiAssignmentsPic.AssignmentType_Auto
                    Case AssgnType.iAssignmentType_Property
                        iMyIcon = GuiAssignmentsPic.AssignmentType_Property
                    Case Else
                        iMyIcon = Empty
                End Select
                itmX.SubItems(GuiAssignments.AssignmentType - 1) = adoRS!AssignmentTypeType
                itmX.ListSubItems(GuiAssignments.AssignmentType - 1).ReportIcon = iMyIcon
            Else
                itmX.SubItems(GuiAssignments.AssignmentType - 1) = vbNullString
                itmX.ListSubItems(GuiAssignments.AssignmentType - 1).ReportIcon = Empty
            End If
            
            'Status
            sStatus = vbNullString
            If Not IsNull(adoRS!Status) Then
                If Not IsNull(adoRS!REASSIGNED) Then
                    If CBool(adoRS!REASSIGNED) Then
                        sStatus = "REASSIGNED (" & adoRS!Status & ")"
                    End If
                End If
                If sStatus = vbNullString Then
                    sStatus = adoRS!Status
                End If
                itmX.SubItems(GuiAssignments.Status - 1) = sStatus
            Else
                itmX.SubItems(GuiAssignments.Status - 1) = vbNullString
            End If
            
            
            
            'IsLocked
            If Not IsNull(adoRS!Islocked) Then
                If CBool(adoRS!Islocked) Then
                    iMyIcon = GuiAssignmentsPic.Islocked
                Else
                    iMyIcon = Empty
                End If
                
                sFlagText = goUtil.GetFlagText(adoRS!Islocked)

                itmX.SubItems(GuiAssignments.Islocked - 1) = sFlagText
                itmX.ListSubItems(GuiAssignments.Islocked - 1).ReportIcon = iMyIcon
            Else
                itmX.SubItems(GuiAssignments.Islocked - 1) = vbNullString
            End If
            'IsDeleted
            If Not IsNull(adoRS!IsDeleted) Then
                If CBool(adoRS!IsDeleted) Then
                    iMyIcon = GuiAssignmentsPic.IsDeleted
                Else
                    iMyIcon = Empty
                End If
                
                sFlagText = goUtil.GetFlagText(adoRS!IsDeleted)
                
                itmX.SubItems(GuiAssignments.IsDeleted - 1) = sFlagText
                itmX.ListSubItems(GuiAssignments.IsDeleted - 1).ReportIcon = iMyIcon
            Else
                itmX.SubItems(GuiAssignments.IsDeleted - 1) = vbNullString
            End If
            'UpLoadMe
            If Not IsNull(adoRS!UpLoadMe) Then
               If CBool(adoRS!UpLoadMe) Then
                    iMyIcon = GuiAssignmentsPic.UpLoadMe
                Else
                    iMyIcon = Empty
                End If
                
                sFlagText = goUtil.GetFlagText(adoRS!UpLoadMe)
               
                itmX.SubItems(GuiAssignments.UpLoadMe - 1) = sFlagText
                itmX.ListSubItems(GuiAssignments.UpLoadMe - 1).ReportIcon = iMyIcon
            Else
                itmX.SubItems(GuiAssignments.UpLoadMe - 1) = vbNullString
            End If
            'SP
            If Not IsNull(adoRS!SPVersion) Then
                itmX.SubItems(GuiAssignments.SP - 1) = adoRS!SPVersion
            Else
                itmX.SubItems(GuiAssignments.SP - 1) = vbNullString
            End If
            
            'Cat Name
            If Not IsNull(adoRS!CatName) Then
                itmX.SubItems(GuiAssignments.CatName - 1) = adoRS!CatName
            Else
                itmX.SubItems(GuiAssignments.CatName - 1) = vbNullString
            End If
            
            'Cat Code
            If Not IsNull(adoRS!CatCode) Then
                itmX.SubItems(GuiAssignments.CatCode - 1) = adoRS!CatCode
            Else
                itmX.SubItems(GuiAssignments.CatCode - 1) = vbNullString
            End If
            'IBNUM
            If Not IsNull(adoRS!IBNUM) Then
                itmX.SubItems(GuiAssignments.IBNUM - 1) = adoRS!IBNUM
            Else
                itmX.SubItems(GuiAssignments.IBNUM - 1) = vbNullString
            End If
            'Insured
            If Not IsNull(adoRS!Insured) Then
                itmX.SubItems(GuiAssignments.Insured - 1) = adoRS!Insured
            Else
                itmX.SubItems(GuiAssignments.Insured - 1) = vbNullString
            End If
            'MAStreet
            If Not IsNull(adoRS!MAStreet) Then
                itmX.SubItems(GuiAssignments.MAStreet - 1) = adoRS!MAStreet
            Else
                itmX.SubItems(GuiAssignments.MAStreet - 1) = vbNullString
            End If
            'MAStreetSort
            If Not IsNull(adoRS!MAStreet) Then
                itmX.SubItems(GuiAssignments.MAStreetSort - 1) = goUtil.utNumInTextSortFormat(adoRS!MAStreet)
            Else
                itmX.SubItems(GuiAssignments.MAStreetSort - 1) = vbNullString
            End If
            'MACity
            If Not IsNull(adoRS!MACity) Then
                itmX.SubItems(GuiAssignments.MACity - 1) = adoRS!MACity
            Else
                itmX.SubItems(GuiAssignments.MACity - 1) = vbNullString
            End If
            'MAState
            If Not IsNull(adoRS!MAState) Then
                itmX.SubItems(GuiAssignments.MAState - 1) = adoRS!MAState
            Else
                itmX.SubItems(GuiAssignments.MAState - 1) = vbNullString
            End If
            'MAZIP
            If Not IsNull(adoRS!MAZIP) Then
                itmX.SubItems(GuiAssignments.MAZIP - 1) = adoRS!MAZIP
            Else
                itmX.SubItems(GuiAssignments.MAZIP - 1) = vbNullString
            End If
            'MAZIP4
            If Not IsNull(adoRS!MAZIP4) Then
                itmX.SubItems(GuiAssignments.MAZIP4 - 1) = adoRS!MAZIP4
            Else
                itmX.SubItems(GuiAssignments.MAZIP4 - 1) = vbNullString
            End If
            'PAStreet
            If Not IsNull(adoRS!PAStreet) Then
                itmX.SubItems(GuiAssignments.PAStreet - 1) = adoRS!PAStreet
            Else
                itmX.SubItems(GuiAssignments.PAStreet - 1) = vbNullString
            End If
            'PAStreetSort
            If Not IsNull(adoRS!PAStreet) Then
                itmX.SubItems(GuiAssignments.PAStreetSort - 1) = goUtil.utNumInTextSortFormat(adoRS!PAStreet)
            Else
                itmX.SubItems(GuiAssignments.PAStreetSort - 1) = vbNullString
            End If
            'PACity
            If Not IsNull(adoRS!PACity) Then
                itmX.SubItems(GuiAssignments.PACity - 1) = adoRS!PACity
            Else
                itmX.SubItems(GuiAssignments.PACity - 1) = vbNullString
            End If
            'PAState
            If Not IsNull(adoRS!PAState) Then
                itmX.SubItems(GuiAssignments.PAState - 1) = adoRS!PAState
            Else
                itmX.SubItems(GuiAssignments.PAState - 1) = vbNullString
            End If
            'PAZIP
            If Not IsNull(adoRS!PAZIP) Then
                itmX.SubItems(GuiAssignments.PAZIP - 1) = adoRS!PAZIP
            Else
                itmX.SubItems(GuiAssignments.PAZIP - 1) = vbNullString
            End If
            'PAZIP4
            If Not IsNull(adoRS!PAZIP4) Then
                itmX.SubItems(GuiAssignments.PAZIP4 - 1) = adoRS!PAZIP4
            Else
                itmX.SubItems(GuiAssignments.PAZIP4 - 1) = vbNullString
            End If
            'Adjuster
            If Not IsNull(adoRS!AdjusterSpecUserName) Then
                itmX.SubItems(GuiAssignments.Adjuster - 1) = adoRS!AdjusterSpecUserName
            Else
                itmX.SubItems(GuiAssignments.Adjuster - 1) = vbNullString
            End If
            'ACID
            If Not IsNull(adoRS!AdjusterSpecACID) Then
                itmX.SubItems(GuiAssignments.ACID - 1) = adoRS!AdjusterSpecACID
            Else
                itmX.SubItems(GuiAssignments.ACID - 1) = vbNullString
            End If
            'LossDate
            If Not IsNull(adoRS!LossDate) Then
                If IsDate(adoRS!LossDate) Then
                    itmX.SubItems(GuiAssignments.LossDate - 1) = Format(adoRS!LossDate, "MM/DD/YYYY")
                    itmX.SubItems(GuiAssignments.LossDateSort - 1) = Format(adoRS!LossDate, "YYYY/MM/DD")
                Else
                    itmX.SubItems(GuiAssignments.LossDate - 1) = vbNullString
                    itmX.SubItems(GuiAssignments.LossDateSort - 1) = vbNullString
                End If
            Else
                itmX.SubItems(GuiAssignments.LossDate - 1) = vbNullString
                itmX.SubItems(GuiAssignments.LossDateSort - 1) = vbNullString
            End If
'            itmX.ListSubItems(GuiAssignments.LossDate - 1).ReportIcon = GuiAssignmentsPic.CalandarPic
            
            'AssignedDate
            If Not IsNull(adoRS!AssignedDate) Then
                If IsDate(adoRS!AssignedDate) Then
                    itmX.SubItems(GuiAssignments.AssignedDate - 1) = Format(adoRS!AssignedDate, "MM/DD/YYYY")
                    itmX.SubItems(GuiAssignments.AssignedDateSort - 1) = Format(adoRS!AssignedDate, "YYYY/MM/DD")
                Else
                    itmX.SubItems(GuiAssignments.AssignedDate - 1) = vbNullString
                    itmX.SubItems(GuiAssignments.AssignedDateSort - 1) = vbNullString
                End If
            Else
                itmX.SubItems(GuiAssignments.AssignedDate - 1) = vbNullString
            End If
'            itmX.ListSubItems(GuiAssignments.AssignedDate - 1).ReportIcon = GuiAssignmentsPic.CalandarPic
            
            'ReceivedDate
            If Not IsNull(adoRS!ReceivedDate) Then
                If IsDate(adoRS!ReceivedDate) Then
                    itmX.SubItems(GuiAssignments.ReceivedDate - 1) = Format(adoRS!ReceivedDate, "MM/DD/YYYY")
                    itmX.SubItems(GuiAssignments.ReceivedDateSort - 1) = Format(adoRS!ReceivedDate, "YYYY/MM/DD")
                Else
                    itmX.SubItems(GuiAssignments.ReceivedDate - 1) = vbNullString
                    itmX.SubItems(GuiAssignments.ReceivedDateSort - 1) = vbNullString
                End If
            Else
                itmX.SubItems(GuiAssignments.ReceivedDate - 1) = vbNullString
                itmX.SubItems(GuiAssignments.ReceivedDateSort - 1) = vbNullString
            End If
'            itmX.ListSubItems(GuiAssignments.ReceivedDate - 1).ReportIcon = GuiAssignmentsPic.CalandarPic
            
            
            'ContactDate
            If Not IsNull(adoRS!ContactDate) Then
                If IsDate(adoRS!ContactDate) Then
                    itmX.SubItems(GuiAssignments.ContactDate - 1) = Format(adoRS!ContactDate, "MM/DD/YYYY")
                    itmX.SubItems(GuiAssignments.ContactDateSort - 1) = Format(adoRS!ContactDate, "YYYY/MM/DD")
                Else
                    itmX.SubItems(GuiAssignments.ContactDate - 1) = vbNullString
                    itmX.SubItems(GuiAssignments.ContactDateSort - 1) = vbNullString
                End If
            Else
                itmX.SubItems(GuiAssignments.ContactDate - 1) = vbNullString
                itmX.SubItems(GuiAssignments.ContactDateSort - 1) = vbNullString
            End If
'            itmX.ListSubItems(GuiAssignments.ContactDate - 1).ReportIcon = GuiAssignmentsPic.CalandarPic
            
            'InspectedDate
            If Not IsNull(adoRS!InspectedDate) Then
                If IsDate(adoRS!InspectedDate) Then
                    itmX.SubItems(GuiAssignments.InspectedDate - 1) = Format(adoRS!InspectedDate, "MM/DD/YYYY")
                    itmX.SubItems(GuiAssignments.InspectedDateSort - 1) = Format(adoRS!InspectedDate, "YYYY/MM/DD")
                Else
                    itmX.SubItems(GuiAssignments.InspectedDate - 1) = vbNullString
                    itmX.SubItems(GuiAssignments.InspectedDateSort - 1) = vbNullString
                End If
            Else
                itmX.SubItems(GuiAssignments.InspectedDate - 1) = vbNullString
                itmX.SubItems(GuiAssignments.InspectedDateSort - 1) = vbNullString
            End If
'            itmX.ListSubItems(GuiAssignments.InspectedDate - 1).ReportIcon = GuiAssignmentsPic.CalandarPic
            
            'CloseDate
            If Not IsNull(adoRS!CloseDate) Then
                If IsDate(adoRS!CloseDate) Then
                    itmX.SubItems(GuiAssignments.CloseDate - 1) = Format(adoRS!CloseDate, "MM/DD/YYYY")
                    itmX.SubItems(GuiAssignments.CloseDateSort - 1) = Format(adoRS!CloseDate, "YYYY/MM/DD")
                Else
                    itmX.SubItems(GuiAssignments.CloseDate - 1) = vbNullString
                    itmX.SubItems(GuiAssignments.CloseDateSort - 1) = vbNullString
                End If
            Else
                itmX.SubItems(GuiAssignments.CloseDate - 1) = vbNullString
                itmX.SubItems(GuiAssignments.CloseDateSort - 1) = vbNullString
            End If
'            itmX.ListSubItems(GuiAssignments.CloseDate - 1).ReportIcon = GuiAssignmentsPic.CalandarPic
            
            'DateReassigned
            If Not IsNull(adoRS!DateReassigned) Then
                If IsDate(adoRS!DateReassigned) Then
                    itmX.SubItems(GuiAssignments.DateReassigned - 1) = Format(adoRS!DateReassigned, "MM/DD/YYYY")
                    itmX.SubItems(GuiAssignments.DateReassignedSort - 1) = Format(adoRS!DateReassigned, "YYYY/MM/DD")
                Else
                    itmX.SubItems(GuiAssignments.DateReassigned - 1) = vbNullString
                    itmX.SubItems(GuiAssignments.DateReassignedSort - 1) = vbNullString
                End If
            Else
                itmX.SubItems(GuiAssignments.DateReassigned - 1) = vbNullString
                itmX.SubItems(GuiAssignments.DateReassignedSort - 1) = vbNullString
            End If
            'DateLastUpdated
            If Not IsNull(adoRS!DateLastUpdated) Then
                If IsDate(adoRS!DateLastUpdated) Then
                    itmX.SubItems(GuiAssignments.DateLastUpdated - 1) = Format(adoRS!DateLastUpdated, "MM/DD/YYYY HH:MM:SS")
                    itmX.SubItems(GuiAssignments.DateLastUpdatedSort - 1) = Format(adoRS!DateLastUpdated, "YYYY/MM/DD HH:MM:SS")
                Else
                    itmX.SubItems(GuiAssignments.DateLastUpdated - 1) = vbNullString
                    itmX.SubItems(GuiAssignments.DateLastUpdatedSort - 1) = vbNullString
                End If
            Else
                itmX.SubItems(GuiAssignments.DateLastUpdated - 1) = vbNullString
                itmX.SubItems(GuiAssignments.DateLastUpdatedSort - 1) = vbNullString
            End If
            'AdminComments
            If Not IsNull(adoRS!AdminComments) Then
                itmX.SubItems(GuiAssignments.AdminComments - 1) = adoRS!AdminComments
            Else
                itmX.SubItems(GuiAssignments.AdminComments - 1) = vbNullString
            End If
            
            'Hidden RKey (Key will be AssignmentsID)
            itmX.SubItems(GuiAssignments.RKey - 1) = adoRS!AssignmentsID
            'Set the colors for Tool tips
            lRED = &HFF&
            lBlue = &HFF8080
            lGreen = &H80FF80
            lBlack = &H0&
            lLightyellow = &HECFFFF
            bDoToolTip = False
            If InStr(1, sStatus, "reject", vbTextCompare) > 0 Then
                bDoToolTip = True
                itmX.ForeColor = lRED
                itmX.ToolTipText = "REJECTED - NEEDS YOUR ATTENTION!"
            Else
                'Also Account for Client reject counts of package items
                'PKGIClientReject
                If adoRS.Fields("PKGIClientReject").Value > 0 Then
                    bDoToolTip = True
                    itmX.ForeColor = lRED
                    sTemp = CStr(adoRS.Fields("PKGIClientReject").Value)
                    itmX.ToolTipText = "CLIENT IS REJECTING " & sTemp & " DOCUMENTS - NEEDS YOUR ATTENTION!"
                'PKGISendMe
                ElseIf adoRS.Fields("PKGISendMe").Value > 0 Then
                    bDoToolTip = True
                    itmX.ForeColor = lBlue
                    sTemp = CStr(adoRS.Fields("PKGISendMe").Value)
                    itmX.ToolTipText = sTemp & " DOCUMENTS HAVE NOT BEEN SENT/DELIVERED TO THE CLIENT - NEEDS YOUR ATTENTION!"
                'PKGIClientApprove
                ElseIf adoRS.Fields("PKGIClientApprove").Value > 0 Then
                    bDoToolTip = True
                    itmX.ForeColor = lGreen
                    sTemp = CStr(adoRS.Fields("PKGIClientApprove").Value)
                    itmX.ToolTipText = sTemp & " documents have been approved by the client - for your information."
                'PKGIClientDelete
                ElseIf adoRS.Fields("PKGIClientDelete").Value > 0 Then
                    bDoToolTip = True
                    itmX.ForeColor = lBlack
                    sTemp = CStr(adoRS.Fields("PKGIClientDelete").Value)
                    itmX.ToolTipText = "Client deleted " & sTemp & " documents - for your information. "
                Else
                End If
            End If
            If bDoToolTip Then
                itmX.ToolTipText = itmX.ToolTipText & vbCrLf & adoRS.Fields("PKGAdminComments").Value
                itmX.Bold = True
                lvwAssignments.BackColor = lLightyellow
                For lSubCount = 1 To itmX.ListSubItems.Count
                    itmX.ListSubItems(lSubCount).Bold = itmX.Bold
                    itmX.ListSubItems(lSubCount).ForeColor = itmX.ForeColor
                    itmX.ListSubItems(lSubCount).ToolTipText = itmX.ToolTipText
                Next
            Else
                itmX.Bold = False
                itmX.ForeColor = lBlack
                itmX.ToolTipText = vbNullString
            End If
            itmX.Selected = False
            
            'Need to add init log entry if there are no entried yet
            InsertInitialActivityLogEntry adoRS!AssignmentsID, oConn
            
            adoRS.MoveNext
        Loop
    Else
        framAssignments.Caption = "Assignments Record Count (0)"
    End If
    
    'close the current ado rs
    adoRS.Close
    
    'Need to get list of Status and populate the Totals for all the
    'displayed assignments Status
    sSQL = "SELECT  StatusID, "
    sSQL = sSQL & "Status, "
    sSQL = sSQL & "Description, "
    sSQL = sSQL & "AdminComments "
    sSQL = sSQL & "FROM     Status "
    sSQL = sSQL & "ORDER BY StatusID "
    
    'Use Disconnected Record Set on asUseClient Cusor ONLY !
    adoRS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set adoRS.ActiveConnection = Nothing
    
     If Not adoRS.EOF Then
        adoRS.MoveFirst
        Do Until adoRS.EOF
            'Only have available lables for status 1 to 7
            If Not IsNull(adoRS!statusid) Then
                lStatusID = adoRS!statusid
                lStatusAssignmentCount = 0
                'Set the Tag to whatever the Status text is
                sStatusTag = adoRS!Status
                Select Case lStatusID
                    Case V2ECKeyBoard.AssgnStatus.iAssignmentsStatus_CLOSED
                        lStatusAssignmentCount = lCLOSED
                    Case V2ECKeyBoard.AssgnStatus.iAssignmentsStatus_DELETED
                        lStatusAssignmentCount = lDELETED
                    Case V2ECKeyBoard.AssgnStatus.iAssignmentsStatus_INTERIM
                        lStatusAssignmentCount = lINTERIM
                    Case V2ECKeyBoard.AssgnStatus.iAssignmentsStatus_NEW
                        lStatusAssignmentCount = lNEW
                    Case V2ECKeyBoard.AssgnStatus.iAssignmentsStatus_PENDING
                        lStatusAssignmentCount = lPENDING
                    Case V2ECKeyBoard.AssgnStatus.iAssignmentsStatus_REASSIGNED
                        lStatusAssignmentCount = lREASSIGNED
                    Case V2ECKeyBoard.AssgnStatus.iAssignmentsStatus_REOPEN
                        lStatusAssignmentCount = lREOPEN
                End Select
                If Not IsNull(adoRS!Status) Then
                    sStatusCaption = adoRS!Status & " (" & lStatusAssignmentCount & ")"
                Else
                    sStatusCaption = vbNullString
                End If
                'Build the tool tip for the Status Label
                If Not IsNull(adoRS!Description) Then
                    sStatusToolTip = adoRS!Description
                Else
                    sStatusToolTip = vbNullString
                End If
                'Put Admin Comments on the tool tip too
                If Not IsNull(adoRS!AdminComments) Then
                    If Trim(adoRS!AdminComments) <> vbNullString Then
                        sStatusToolTip = sStatusToolTip & " (" & adoRS!AdminComments & ") "
                    End If
                End If
                
            End If
            adoRS.MoveNext
        Loop
    End If
    
    'Cleanup
    Set adoRS = Nothing
    Set oConn = Nothing
    Set itmX = Nothing
    
    
    Exit Sub
EH:
    Set adoRS = Nothing
    Set oConn = Nothing
    Set itmX = Nothing
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Populatelvw"
End Sub

Public Function SelectAssignments(Optional piMyAssgnStatus As V2ECKeyBoard.AssgnStatus, Optional psFlagName As String)
    On Error GoTo EH
    Dim itmX As ListItem
    Dim sStatus As String
    
    For Each itmX In lvwAssignments.ListItems
        If piMyAssgnStatus > 0 Then
            sStatus = itmX.SubItems(GuiAssignments.Status - 1)
            If InStr(1, sStatus, psFlagName, vbTextCompare) > 0 Then
                itmX.Selected = True
            Else
                itmX.Selected = False
            End If
        Else
            If Not psFlagName = vbNullString Then
                Select Case UCase(psFlagName)
                    Case UCase("Deleted")
                        If goUtil.GetFlagFromText(itmX.SubItems(GuiAssignments.IsDeleted - 1)) Then
                            itmX.Selected = True
                        Else
                            itmX.Selected = False
                        End If
                    Case UCase("UpLoadMe")
                        If goUtil.GetFlagFromText(itmX.SubItems(GuiAssignments.UpLoadMe - 1)) Then
                            itmX.Selected = True
                        Else
                            itmX.Selected = False
                        End If
                End Select
            End If
        End If
    Next
    
    'cleanup
    Set itmX = Nothing
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function SelectAssignments"
End Function


Private Sub lvwAssignments_Click()
    On Error GoTo EH
    Dim LVAInfo As LVASSGNINFO
    Dim bisDateClick As Boolean
    Dim sCloseDate As String
    Dim IsLockedFlagText As String
    Dim sMess As String
    
    'Set the selected claim
    itmXSelected = lvwAssignments.SelectedItem
    If chkAllowItemUpdates.Value = vbUnchecked Then
        txtToolTip.Text = itmXSelected.ToolTipText
        Exit Sub
    End If
    
    'need to Check for the Clicked Item and Get the Subitem index
    'to see if it happend to be an updateable item
    With LVAInfo
       .pt.X = (mlX \ Screen.TwipsPerPixelX)
       .pt.Y = (mlY \ Screen.TwipsPerPixelY)
       .flags = LVHT_ONITEM
    End With
    
    Call SendMessage(lvwAssignments.hWnd, LVM_SUBITEMASSGN, 0, LVAInfo)

    If (LVAInfo.iItem < 0) Then
        Exit Sub
    End If
    
    'check for List updateable items
    Select Case LVAInfo.iSubItem + 1
        Case GuiAssignments.LossDate
            txtParamValue.Tag = "LossDate"
            bisDateClick = True
        Case GuiAssignments.AssignedDate
            txtParamValue.Tag = "AssignedDate"
            bisDateClick = True
        Case GuiAssignments.ReceivedDate
            txtParamValue.Tag = "ReceivedDate"
            bisDateClick = True
        Case GuiAssignments.ContactDate
            txtParamValue.Tag = "ContactDate"
            bisDateClick = True
        Case GuiAssignments.InspectedDate
            txtParamValue.Tag = "InspectedDate"
            bisDateClick = True
        Case GuiAssignments.CloseDate
            '10.28.2005 BGS  Per Rob Petrovics Request... Close Dates for All Profiles
            'will no longer be updateable by the adjuster, only a manager or admin may
            'update the closed date.
'            txtParamValue.Tag = "CloseDate"
'            bisDateClick = True
            
    End Select
    
    'If the Closed Date is Set Then Need to Give message
    'indicating the the assignment is closed first !
    sCloseDate = itmXSelected.SubItems(GuiAssignments.CloseDate - 1)
    If IsDate(sCloseDate) And (LVAInfo.iSubItem + 1) <> GuiAssignments.CloseDate Then
        sMess = "Can't update this item from this screen once the close date is set!"
        MsgBox sMess, vbInformation + vbOKOnly, "Item is Closed!"
        Exit Sub
    End If
    
    'if the Is Locked Flag is Set for this Record Can not Allow Chnages to Dates
    If Not AllowThisStatus(sMess) Then
        sMess = Replace(sMess, "[CHECK_FOR_UNDO_APPROVALRequest]", vbNullString, , , vbTextCompare)
        MsgBox sMess, vbInformation + vbOKOnly, "Operation Not Allowed!"
        Exit Sub
    End If
    
    If bisDateClick Then
        txtParamValue.Text = itmXSelected.SubItems(LVAInfo.iSubItem)
        txtParamValue.MaxLength = 20
        MyGUI.ShowCalendar txtParamValue
        UpdateEditDate LVAInfo.iSubItem
    End If
    
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lvwAssignments_Click"
End Sub

Private Sub lvwAssignments_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    On Error GoTo EH
    Dim sMess As String
    If lvwAssignments.SortOrder = lvwAscending Then
        lvwAssignments.SortOrder = lvwDescending
        sMess = "Descending"
    Else
        lvwAssignments.SortOrder = lvwAscending
        sMess = "Ascending"
    End If
    
    'Set the Tool Tip
    lvwAssignments.ToolTipText = "Sort By " & ColumnHeader.Text & " " & sMess
    
    'Need to see if the Column clicked was a NON text sort
    'Like Date or number.  If a non textual column was clicked
    'need to use the next column as the sort since this next column
    'will be hidden and contain a sort friendly format.
    'Sort Key is Base 0 where CoulmnHeader is not
    Select Case ColumnHeader.Index
        Case GuiAssignments.MAStreet
            lvwAssignments.SortKey = ColumnHeader.Index
        Case GuiAssignments.PAStreet
            lvwAssignments.SortKey = ColumnHeader.Index
        Case GuiAssignments.LossDate
            lvwAssignments.SortKey = ColumnHeader.Index
        Case GuiAssignments.AssignedDate
            lvwAssignments.SortKey = ColumnHeader.Index
        Case GuiAssignments.ReceivedDate
            lvwAssignments.SortKey = ColumnHeader.Index
        Case GuiAssignments.ContactDate
            lvwAssignments.SortKey = ColumnHeader.Index
        Case GuiAssignments.InspectedDate
            lvwAssignments.SortKey = ColumnHeader.Index
        Case GuiAssignments.CloseDate
            lvwAssignments.SortKey = ColumnHeader.Index
        Case GuiAssignments.DateReassigned
            lvwAssignments.SortKey = ColumnHeader.Index
        Case GuiAssignments.DateLastUpdated
            lvwAssignments.SortKey = ColumnHeader.Index
        Case Else
            lvwAssignments.SortKey = ColumnHeader.Index - 1
    End Select
    
    lvwAssignments.Sorted = True
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lvwAssignments_ColumnClick"
End Sub

Private Sub lvwAssignments_DblClick()
    On Error GoTo EH
    
    'Set the selected claim
    
    itmXSelected = lvwAssignments.SelectedItem
    If Not lvwAssignments.SelectedItem Is Nothing Then
        ShowClaim
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lvwAssignments_DblClick"
End Sub

Public Function PopulateFrmCaptionAssignmentInfo(poForm As Object, _
                                                Optional psCaptionPrefix As String, _
                                                Optional psCaptionSuffix As String) As Boolean
    On Error GoTo EH
    Dim myForm As Form
    Dim sIBNUM As String
    Dim sCLIENTNUM As String
    Dim sInsured As String
    Dim sAssignedDate As String
    Dim sCaption As String
    
    'Need to show the current selected assignment in the caption
    If mitmXSelected Is Nothing Then
        sCaption = psCaptionPrefix
        If psCaptionSuffix <> vbNullString Then
            sCaption = sCaption & " - " & psCaptionSuffix
        End If
        If Not TypeOf poForm Is Form Then
            Exit Function
        End If
    
        Set myForm = poForm
        
        myForm.Caption = sCaption
        PopulateFrmCaptionAssignmentInfo = True
        Set myForm = Nothing
        Exit Function
    End If
    
    If Not TypeOf poForm Is Form Then
        Exit Function
    End If
    
    Set myForm = poForm
    
    sIBNUM = mitmXSelected.SubItems(GuiAssignments.IBNUM - 1)
    sCLIENTNUM = mitmXSelected.Text
    sInsured = mitmXSelected.SubItems(GuiAssignments.Insured - 1)
    sAssignedDate = mitmXSelected.SubItems(GuiAssignments.AssignedDate - 1)
    sCaption = psCaptionPrefix & " ("
    sCaption = sCaption & "IB: " & sIBNUM
    sCaption = sCaption & " - CLAIM: " & sCLIENTNUM
    sCaption = sCaption & " - ASSIGNED: " & sAssignedDate
    sCaption = sCaption & " - INSURED: " & sInsured
    If psCaptionSuffix <> vbNullString Then
        sCaption = sCaption & " - " & psCaptionSuffix
    End If
    sCaption = sCaption & ")"
    
    myForm.Caption = sCaption
    
    PopulateFrmCaptionAssignmentInfo = True
    
    'cleanup
    Set myForm = Nothing
    
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function PopulateFrmCaptionAssignmentInfo"
End Function

Public Function GetClaimItemAsString(pClaimItem As GuiAssignments) As String
    On Error GoTo EH
    Dim sValue As String
    
    If mitmXSelected Is Nothing Then
        Exit Function
    End If
    If pClaimItem = 1 Then
        sValue = mitmXSelected.Text
    Else
        sValue = mitmXSelected.SubItems(pClaimItem - 1)
    End If
    
    GetClaimItemAsString = sValue
    
    Exit Function
EH:
   goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function GetClaimItemAsString"
End Function

Private Sub lvwAssignments_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo EH
    
    'Set the selected claim
    If Not lvwAssignments.SelectedItem Is Nothing Then
        itmXSelected = lvwAssignments.SelectedItem
        txtToolTip.Text = itmXSelected.ToolTipText
        Select Case KeyCode
            Case KeyCodeConstants.vbKeyReturn
                ShowClaim
        End Select
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lvwAssignments_KeyDown"
End Sub

Private Sub lvwAssignments_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo EH
    
    'Set the selected claim
    If Not lvwAssignments.SelectedItem Is Nothing Then
        itmXSelected = lvwAssignments.SelectedItem
        txtToolTip.Text = itmXSelected.ToolTipText
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lvwAssignments_KeyUp"
End Sub

Public Function ShowClaim() As Boolean
    On Error GoTo EH
    Dim sMess As String
    
    If mitmXSelected Is Nothing Then
        Exit Function
    End If
    
    If AllowThisStatus(sMess) Then
        Set mMyfrmClaim = New frmClaim
        Set mMyfrmClaim.MyClaimsList = Me
        Set mMyfrmClaim.MyGUI = Me.MyGUI
        mMyfrmClaim.AssignmentsID = mitmXSelected.SubItems(GuiAssignments.RKey - 1)
        Load mMyfrmClaim
        PopulateFrmCaptionAssignmentInfo mMyfrmClaim, mMyfrmClaim.Caption
        mMyfrmClaim.Show vbModeless
        
        Me.Visible = False
        
        ShowClaim = True
    Else
        sMess = Replace(sMess, "[CHECK_FOR_UNDO_APPROVALRequest]", vbNullString, , , vbTextCompare)
        MsgBox sMess, vbExclamation + vbOKOnly, "Can't View Claim"
    End If

    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function ShowClaim"
End Function

Public Function UnloadClaim() As Boolean
    On Error GoTo EH
    
    If Not mMyfrmClaim Is Nothing Then
        mMyfrmClaim.CLEANUP
        Unload mMyfrmClaim
        Set mMyfrmClaim = Nothing
    End If
    
    Me.Visible = True
    
    UnloadClaim = True
    mbUnloadingClaim = False
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function UnloadClaim"
    mbUnloadingClaim = False
End Function



Private Sub Timer_UnloadClaim_Timer()
    On Error GoTo EH
    Timer_UnloadClaim.Enabled = False
    UnloadClaim
    
    Exit Sub
EH:
    Timer_UnloadClaim.Enabled = False
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Timer_UnloadClaim_Timer"
End Sub


Public Function CLEANUP() As Boolean
    On Error GoTo EH
    
    If Not mMyfrmClaim Is Nothing Then
        mMyfrmClaim.CLEANUP
        Unload mMyfrmClaim
        Set mMyfrmClaim = Nothing
    End If
    
    Set madoRSAssignments = Nothing
    
    Set moGUI = Nothing
    
    If Not mArv Is Nothing Then
        mArv.CLEANUP
        Set mArv = Nothing
    End If
    
    If Not moForm Is Nothing Then
        Unload moForm
        Set moForm = Nothing
    End If
    
    CLEANUP = True
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function CLEANUP"
End Function

Public Function InsertInitialActivityLogEntry(psAssignmentsID, Optional poConn As ADODB.Connection) As Boolean
    On Error GoTo EH
    Dim sAssignmentsID As String
    Dim sSQL As String
    Dim RS As ADODB.Recordset
    Dim oConn As ADODB.Connection
    Dim sID As String
    Dim sActText As String
    Dim sPolicyDesc As String
    Dim sDeductible As String
    Dim sPolicyLimits As String
    
    sAssignmentsID = psAssignmentsID
    Set oConn = poConn
    
    'Check to see if there are any activitylog entries for this assignments id yet.
    '(include deleted records to tell if anything has ever been added to actlog)
    'if there are none then add an initial entry
    
    sSQL = "SELECT [RTActivityLogID] FROM RTActivityLog RTAL "
    sSQL = sSQL & "WHERE RTAL.AssignmentsID = " & sAssignmentsID & " "
    
    'use the passed in connection unless it is nothing
    If oConn Is Nothing Then
        Set oConn = New ADODB.Connection
        goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    End If
    Set RS = New ADODB.Recordset
    
    
    'Use Disconnected Record Set on asUseClient Cusor ONLY !
    RS.CursorLocation = adUseClient
    RS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set RS.ActiveConnection = Nothing
    
    'If the record count is 0 then need to add an entry
    If RS.RecordCount > 0 Then
        'otherwise cleanup and bail
        GoTo CLEAN_UP
    End If
    
    Set RS = Nothing
    
    'Need to Build the Activity Text for the initial Log Entry
    sSQL = "SELECT "
    sSQL = sSQL & "( "
    sSQL = sSQL & "SELECT  A.[PolicyDescription] "
    sSQL = sSQL & "FROM    Assignments A "
    sSQL = sSQL & "Where A.[AssignmentsID] = PL.[AssignmentsID] "
    sSQL = sSQL & ") As [PolicyDescription], "
    sSQL = sSQL & "( "
    sSQL = sSQL & "SELECT  A2.[Deductible] "
    sSQL = sSQL & "FROM    Assignments A2 "
    sSQL = sSQL & "Where A2.[AssignmentsID] = PL.[AssignmentsID] "
    sSQL = sSQL & ") As [Deductible], "
    sSQL = sSQL & "( "
    sSQL = sSQL & "SELECT  CT.[Description] "
    sSQL = sSQL & "FROM    ClassType CT "
    sSQL = sSQL & "Where CT.[ClassTypeID] = PL.[ClassTypeID] "
    sSQL = sSQL & ") As [LimitDescription], "
    sSQL = sSQL & "PL.[LimitAmount] "
    sSQL = sSQL & "FROM PolicyLimits PL "
    sSQL = sSQL & "WHERE PL.[AssignmentsID] = "
    sSQL = sSQL & "( "
    sSQL = sSQL & "SELECT  [AssignmentsID] "
    sSQL = sSQL & "From Assignments "
    sSQL = sSQL & "Where [AssignmentsID] = " & sAssignmentsID & " "
    sSQL = sSQL & ") "
    sSQL = sSQL & "And PL.[LimitAmount] > 0 "
    sSQL = sSQL & "And PL.[IsDeleted] = 0 "
    
    Set RS = New ADODB.Recordset
    'Use Disconnected Record Set on asUseClient Cusor ONLY !
    RS.CursorLocation = adUseClient
    RS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set RS.ActiveConnection = Nothing
    
    'If the record count is 0 Bail because there is nothing to talk about
    If RS.RecordCount = 0 Then
        'otherwise cleanup and bail
        GoTo CLEAN_UP
    Else
        RS.MoveFirst
    End If
    
    'loop through the RS and build the ActText
    Do Until RS.EOF
        'set the Vars
        sPolicyDesc = goUtil.IsNullIsVbNullString(RS.Fields("PolicyDescription")) & " "
        sDeductible = Format(goUtil.IsNullIsVbNullString(RS.Fields("Deductible")), "###,###,###") & " "
        sPolicyLimits = sPolicyLimits & goUtil.IsNullIsVbNullString(RS.Fields("LimitDescription"))
        sPolicyLimits = sPolicyLimits & " = "
        sPolicyLimits = sPolicyLimits & Format(goUtil.IsNullIsVbNullString(RS.Fields("LimitAmount")), "###,###,###") & " "
        RS.MoveNext
    Loop
    
    sActText = "Claim received. Policy Description: " & sPolicyDesc & " "
    sActText = sActText & "Policy Limits: " & sPolicyLimits & " "
    sActText = sActText & "Deductible: " & sDeductible & " "
    
    'need to Get Unique Id for Act Log
    sID = goUtil.GetAccessDBUID("ID", "RTActivityLog")
    
    sSQL = "INSERT INTO RTActivityLog "
    sSQL = sSQL & "( "
    sSQL = sSQL & "[RTActivityLogID], "
    sSQL = sSQL & "[AssignmentsID], "
    sSQL = sSQL & "[BillingCountID], "
    sSQL = sSQL & "[ID], "
    sSQL = sSQL & "[IDAssignments], "
    sSQL = sSQL & "[IDBillingCount], "
    sSQL = sSQL & "[ServiceTime], "
    sSQL = sSQL & "[ActDate], "
    sSQL = sSQL & "[ActText], "
    sSQL = sSQL & "[ActTime], "
    sSQL = sSQL & "[PageBreakAfter], "
    sSQL = sSQL & "[BlankPageAfter], "
    sSQL = sSQL & "[BlankRowsAfter], "
    sSQL = sSQL & "[IsMgrEntry], "
    sSQL = sSQL & "[IsDeleted], "
    sSQL = sSQL & "[DownLoadMe], "
    sSQL = sSQL & "[UpLoadMe], "
    sSQL = sSQL & "[AdminComments], "
    sSQL = sSQL & "[DateLastUpdated], "
    sSQL = sSQL & "[UpdateByUserID] "
    sSQL = sSQL & ") "
    sSQL = sSQL & "SELECT "
    sSQL = sSQL & sID & " As [RTActivityLogID], "
    sSQL = sSQL & sAssignmentsID & " As [AssignmentsID], "
    sSQL = sSQL & "null" & " As [BillingCountID] , "
    sSQL = sSQL & sID & " As [ID], "
    sSQL = sSQL & sAssignmentsID & " As [IDAssignments], "
    sSQL = sSQL & "null" & " As [IDBillingCount], "
    sSQL = sSQL & "0.00" & " As [ServiceTime], "
    sSQL = sSQL & "#" & Format(Now(), "MM/DD/YYYY") & "#" & " As [ActDate], "
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(sActText) & "'" & " As [ActText], "
    sSQL = sSQL & "#" & Now() & "#" & " As [ActTime], "
    sSQL = sSQL & "False" & " As [PageBreakAfter], "
    sSQL = sSQL & "False" & " As [BlankPageAfter], "
    sSQL = sSQL & "0" & " As [BlankRowsAfter], "
    sSQL = sSQL & "False" & " As [IsMgrEntry], "
    sSQL = sSQL & "False" & " As [IsDeleted], "
    sSQL = sSQL & "False" & " As [DownLoadMe], "
    sSQL = sSQL & "True" & " As [UpLoadMe], "
    sSQL = sSQL & "''" & " As [AdminComments], "
    sSQL = sSQL & "#" & Now() & "#" & " As [DateLastUpdated], "
    sSQL = sSQL & goUtil.gsCurUsersID & " As [UpdateByUserID] "
    
    oConn.Execute sSQL
    InsertInitialActivityLogEntry = True
CLEAN_UP:
    
    'cleanup
    Set RS = Nothing
    Set oConn = Nothing
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function InsertInitialActivityLogEntry"
End Function
 
 
Public Function PopulateStatusList() As Boolean
    On Error GoTo EH
    Dim oConn As New ADODB.Connection
    Dim RS As ADODB.Recordset
    Dim sSQL As String
    
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    Set RS = New ADODB.Recordset
    
    sSQL = "SELECT  [StatusID], "
    sSQL = sSQL & "[StatusAlias], "
    sSQL = sSQL & "[Status], "
    sSQL = sSQL & "[Description], "
    sSQL = sSQL & "[AdminComments], "
    sSQL = sSQL & "[IsDeleted], "
    sSQL = sSQL & "[DateLastUpdated], "
    sSQL = sSQL & "[UpdateByUserID] "
    sSQL = sSQL & "FROM Status "
    sSQL = sSQL & "WHERE [IsDeleted] = 0 "
    sSQL = sSQL & "ORDER BY [StatusAlias] "
    
    RS.CursorLocation = adUseClient
    RS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set RS.ActiveConnection = Nothing
    
    cboSelStatus.AddItem "ALL STATUSES - (CLICK HERE TO FILTER BY STATUS)"
    cboSelStatus.ItemData(cboSelStatus.NewIndex) = -1
    
    Do Until RS.EOF
        cboSelStatus.AddItem goUtil.IsNullIsVbNullString(RS.Fields("StatusAlias")) & " - (" & goUtil.IsNullIsVbNullString(RS.Fields("Description")) & ")"
        cboSelStatus.ItemData(cboSelStatus.NewIndex) = goUtil.IsNullIsVbNullString(RS.Fields("StatusID"))
        RS.MoveNext
    Loop
    
    cboSelStatus.ListIndex = 0
    
    Exit Function
EH:
  goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function PopulateStatusList"
End Function

Public Function UpdateEditDate(pMySubitemIndex As Long) As Boolean
    On Error GoTo EH
    Dim itmX As ListItem
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    Dim lRecordsAffected As Long
    Dim sUpdateColName As String
    Dim sUpdateColValue As String
    Dim sFlagText As String
    Dim sDateLastUpdated As String
    Dim sAssignmentsID As String
    
    Set itmX = itmXSelected
    
    If Not ValidDateWhenCloseDate(txtParamValue, itmX) Then
        Exit Function
    End If
    
    'If the Param value is not a valid Date then abort any updates!
    If Not IsDate(txtParamValue.Text) Then
        Set itmX = Nothing
        Exit Function
    ElseIf txtParamValue.Text = itmX.ListSubItems(pMySubitemIndex).Text Then
        'Exit this function if the date being changed is the same
        'As the one already in the DB
        Set itmX = Nothing
        Exit Function
    End If
    
    Screen.MousePointer = MousePointerConstants.vbHourglass
    
    sUpdateColName = txtParamValue.Tag
    sUpdateColValue = txtParamValue.Text
    sFlagText = goUtil.GetFlagText(True)
    sDateLastUpdated = Now()
    sAssignmentsID = itmX.ListSubItems(GuiAssignments.RKey - 1)
    
    sSQL = "UPDATE Assignments SET "
    sSQL = sSQL & "[" & sUpdateColName & "] = #" & sUpdateColValue & "#, "
    If StrComp(sUpdateColName, "CloseDate", vbTextCompare) = 0 Then
        sSQL = sSQL & "[StatusID] = " & iAssignmentsStatus_CLOSED & ", "
    End If
    sSQL = sSQL & "[UpLoadMe] = True, "
    sSQL = sSQL & "[DateLastUpdated] = #" & sDateLastUpdated & "#, "
    sSQL = sSQL & "[UpdateByUserID] = " & goUtil.gsCurUsersID & " "
    sSQL = sSQL & "WHERE [ID] = " & sAssignmentsID & " "

    'first Update the DB
    Set oConn = New ADODB.Connection
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    oConn.Execute sSQL, lRecordsAffected
    Sleep 100

    'if the Affected Records > 0 then update the List view Gui
    If lRecordsAffected = 1 Then
        itmX.ListSubItems(pMySubitemIndex).Text = sUpdateColValue
        If StrComp(sUpdateColName, "Closedate", vbTextCompare) = 0 Then
            itmX.ListSubItems(GuiAssignments.Status - 1).Text = "CLOSED"
        End If
        itmX.ListSubItems(GuiAssignments.UpLoadMe - 1).Text = sFlagText
        itmX.ListSubItems(GuiAssignments.UpLoadMe - 1).ReportIcon = GuiAssignmentsPic.UpLoadMe
        itmX.ListSubItems(GuiAssignments.DateLastUpdated - 1).Text = Format(sDateLastUpdated, "MM/DD/YYYY HH:MM:SS")
    End If

    Screen.MousePointer = MousePointerConstants.vbDefault
    
    UpdateEditDate = True

    'cleanup
     Set itmX = Nothing
     Set oConn = Nothing
     
    Exit Function
EH:
    Screen.MousePointer = MousePointerConstants.vbDefault
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function UpdateEditDate"
End Function

Public Function ValidDateWhenCloseDate(poDateBox As Object, poitmX As ListItem) As Boolean
    On Error GoTo EH
    Dim oDateBox As TextBox
    Dim oOtherDate As TextBox
    Dim oControl As Control
    Dim sDateName As String
    Dim sMess As String
    Dim itmX As ListItem
    Dim sLossDate As String
    Dim sAssignedDate As String
    Dim sReceivedDate As String
    Dim sContactDate As String
    Dim sInspectedDate As String
    Dim sCloseDate As String
    Dim bOtherDatesNotFilledOut As Boolean

    If Not TypeOf poDateBox Is TextBox Then
        Exit Function
    ElseIf Not TypeOf poitmX Is ListItem Then
        Exit Function
    Else
        Set oDateBox = poDateBox
        Set itmX = poitmX
    End If
    
    'set all the dates for the currently selected item
    sLossDate = itmX.ListSubItems(GuiAssignments.LossDate - 1).Text
    sAssignedDate = itmX.ListSubItems(GuiAssignments.AssignedDate - 1).Text
    sReceivedDate = itmX.ListSubItems(GuiAssignments.ReceivedDate - 1).Text
    sContactDate = itmX.ListSubItems(GuiAssignments.ContactDate - 1).Text
    sInspectedDate = itmX.ListSubItems(GuiAssignments.InspectedDate - 1).Text
    sCloseDate = itmX.ListSubItems(GuiAssignments.CloseDate - 1).Text
    
    If IsDate(sCloseDate) Then
        If Not IsDate(oDateBox.Text) Then
            sMess = "Dates can not be blank if the Close Date is Set!" & vbCrLf & vbCrLf
        End If
    End If
    
    'make sure that the Dates Jive with each other
    Select Case UCase(oDateBox.Tag)
        Case UCase("LossDate")
            sDateName = "Loss Date"
            If Not IsDate(oDateBox.Text) Then
                GoTo MESSAGE_HERE
            End If
            'Loss date Can't be > than any other Dates!
            If IsDate(sAssignedDate) Then
                If CDate(oDateBox.Text) > CDate(sAssignedDate) Then
                    sMess = sMess & sDateName & " can not be later than " & sAssignedDate & vbCrLf
                End If
            End If
            If IsDate(sReceivedDate) Then
                If CDate(oDateBox.Text) > CDate(sReceivedDate) Then
                    sMess = sMess & sDateName & " can not be later than " & sReceivedDate & vbCrLf
                End If
            End If
            If IsDate(sContactDate) Then
                If CDate(oDateBox.Text) > CDate(sContactDate) Then
                    sMess = sMess & sDateName & " can not be later than " & sContactDate & vbCrLf
                End If
            End If
            If IsDate(sInspectedDate) Then
                If CDate(oDateBox.Text) > CDate(sInspectedDate) Then
                    sMess = sMess & sDateName & " can not be later than " & sInspectedDate & vbCrLf
                End If
            End If
            If IsDate(sCloseDate) Then
                If CDate(oDateBox.Text) > CDate(sCloseDate) Then
                    sMess = sMess & sDateName & " can not be later than " & sCloseDate & vbCrLf
                End If
            End If
        Case UCase("AssignedDate")
            sDateName = "Assigned Date"
            If Not IsDate(oDateBox.Text) Then
                GoTo MESSAGE_HERE
            End If
            'Assigned Date Can't be > than any other Dates! excpet for Loss Date.
            If IsDate(sReceivedDate) Then
                If CDate(oDateBox.Text) > CDate(sReceivedDate) Then
                    sMess = sMess & sDateName & " can not be later than " & sReceivedDate & vbCrLf
                End If
            End If
            If IsDate(sContactDate) Then
                If CDate(oDateBox.Text) > CDate(sContactDate) Then
                    sMess = sMess & sDateName & " can not be later than " & sContactDate & vbCrLf
                End If
            End If
            If IsDate(sInspectedDate) Then
                If CDate(oDateBox.Text) > CDate(sInspectedDate) Then
                    sMess = sMess & sDateName & " can not be later than " & sInspectedDate & vbCrLf
                End If
            End If
            If IsDate(sCloseDate) Then
                If CDate(oDateBox.Text) > CDate(sCloseDate) Then
                    sMess = sMess & sDateName & " can not be later than " & sCloseDate & vbCrLf
                End If
            End If
            'Assigned Date Can't be < than Loss Date.
            If IsDate(sLossDate) Then
                If CDate(oDateBox.Text) < CDate(sLossDate) Then
                    sMess = sMess & sDateName & " can not be earlier than " & sLossDate & vbCrLf
                End If
            End If
        Case UCase("ReceivedDate")
            sDateName = "Received Date"
            If Not IsDate(oDateBox.Text) Then
                GoTo MESSAGE_HERE
            End If
             'Received Date Can't be > than any other Dates! excpet for Loss Date And Assigned Date.
            If IsDate(sContactDate) Then
                If CDate(oDateBox.Text) > CDate(sContactDate) Then
                    sMess = sMess & sDateName & " can not be later than " & sContactDate & vbCrLf
                End If
            End If
            If IsDate(sInspectedDate) Then
                If CDate(oDateBox.Text) > CDate(sInspectedDate) Then
                    sMess = sMess & sDateName & " can not be later than " & sInspectedDate & vbCrLf
                End If
            End If
            If IsDate(sCloseDate) Then
                If CDate(oDateBox.Text) > CDate(sCloseDate) Then
                    sMess = sMess & sDateName & " can not be later than " & sCloseDate & vbCrLf
                End If
            End If
            'Received Date Can't be < than Loss Date or Assigned Date
            If IsDate(sLossDate) Then
                If CDate(oDateBox.Text) < CDate(sLossDate) Then
                    sMess = sMess & sDateName & " can not be earlier than " & sLossDate & vbCrLf
                End If
            End If
            If IsDate(sAssignedDate) Then
                If CDate(oDateBox.Text) < CDate(sAssignedDate) Then
                    sMess = sMess & sDateName & " can not be earlier than " & sAssignedDate & vbCrLf
                End If
            End If
        Case UCase("ContactDate")
            sDateName = "Contact Date"
            If Not IsDate(oDateBox.Text) Then
                GoTo MESSAGE_HERE
            End If
            'Contact Date  Can't be < than any other Dates! excpet for txtInspected Date And txtClose Date.
            If IsDate(sLossDate) Then
                If CDate(oDateBox.Text) < CDate(sLossDate) Then
                    sMess = sMess & sDateName & " can not be earlier than " & sLossDate & vbCrLf
                End If
            End If
            If IsDate(sAssignedDate) Then
                If CDate(oDateBox.Text) < CDate(sAssignedDate) Then
                    sMess = sMess & sDateName & " can not be earlier than " & sAssignedDate & vbCrLf
                End If
            End If
            If IsDate(sReceivedDate) Then
                If CDate(oDateBox.Text) < CDate(sReceivedDate) Then
                    sMess = sMess & sDateName & " can not be earlier than " & sReceivedDate & vbCrLf
                End If
            End If
            'Contact Date  Can't be > than txtInspected Date Or txtClose Date.
            If IsDate(sInspectedDate) Then
                If CDate(oDateBox.Text) > CDate(sInspectedDate) Then
                    sMess = sMess & sDateName & " can not be later than " & sInspectedDate & vbCrLf
                End If
            End If
            If IsDate(sCloseDate) Then
                If CDate(oDateBox.Text) > CDate(sCloseDate) Then
                    sMess = sMess & sDateName & " can not be later than " & sCloseDate & vbCrLf
                End If
            End If
        Case UCase("InspectedDate")
            sDateName = "Inspected Date"
            If Not IsDate(oDateBox.Text) Then
                GoTo MESSAGE_HERE
            End If
            'txtInspected Date  Can't be < than any other Dates! excpet for txtClose Date.
            If IsDate(sLossDate) Then
                If CDate(oDateBox.Text) < CDate(sLossDate) Then
                    sMess = sMess & sDateName & " can not be earlier than " & sLossDate & vbCrLf
                End If
            End If
            If IsDate(sAssignedDate) Then
                If CDate(oDateBox.Text) < CDate(sAssignedDate) Then
                    sMess = sMess & sDateName & " can not be earlier than " & sAssignedDate & vbCrLf
                End If
            End If
            If IsDate(sReceivedDate) Then
                If CDate(oDateBox.Text) < CDate(sReceivedDate) Then
                    sMess = sMess & sDateName & " can not be earlier than " & sReceivedDate & vbCrLf
                End If
            End If
            If IsDate(sContactDate) Then
                If CDate(oDateBox.Text) < CDate(sContactDate) Then
                    sMess = sMess & sDateName & " can not be earlier than " & sContactDate & vbCrLf
                End If
            End If
            'Inspected Date  Can't be > than txtClose Date.
            If IsDate(sCloseDate) Then
                If CDate(oDateBox.Text) > CDate(sCloseDate) Then
                    sMess = sMess & sDateName & " can not be later than " & sCloseDate & vbCrLf
                End If
            End If
        Case UCase("CloseDate")
            sDateName = "Close Date"
            If Not IsDate(oDateBox.Text) Then
                GoTo MESSAGE_HERE
            End If

            'first Check to see if all the other dates have been filled out
            If Not IsDate(sLossDate) Then
                 bOtherDatesNotFilledOut = True
            ElseIf Not IsDate(sAssignedDate) Then
                 bOtherDatesNotFilledOut = True
            ElseIf Not IsDate(sReceivedDate) Then
                 bOtherDatesNotFilledOut = True
            ElseIf Not IsDate(sContactDate) Then
                 bOtherDatesNotFilledOut = True
            ElseIf Not IsDate(sInspectedDate) Then
                 bOtherDatesNotFilledOut = True
            End If
            
            If bOtherDatesNotFilledOut And IsDate(oDateBox.Text) Then
                sMess = "You must fill out all other dates before the Close Date."
                GoTo MESSAGE_HERE
            End If

            'Close Date date Can't be < than any other Dates!
            If IsDate(sLossDate) Then
                If CDate(oDateBox.Text) < CDate(sLossDate) Then
                    sMess = sMess & sDateName & " can not be earlier than " & sLossDate & vbCrLf
                End If
            End If
            If IsDate(sAssignedDate) Then
                If CDate(oDateBox.Text) < CDate(sAssignedDate) Then
                    sMess = sMess & sDateName & " can not be earlier than " & sAssignedDate & vbCrLf
                End If
            End If
            If IsDate(sReceivedDate) Then
                If CDate(oDateBox.Text) < CDate(sReceivedDate) Then
                    sMess = sMess & sDateName & " can not be earlier than " & sReceivedDate & vbCrLf
                End If
            End If
            If IsDate(sContactDate) Then
                If CDate(oDateBox.Text) < CDate(sContactDate) Then
                    sMess = sMess & sDateName & " can not be earlier than " & sContactDate & vbCrLf
                End If
            End If
            If IsDate(sInspectedDate) Then
                If CDate(oDateBox.Text) < CDate(sInspectedDate) Then
                    sMess = sMess & sDateName & " can not be earlier than " & sInspectedDate & vbCrLf
                End If
            End If
            
    End Select

MESSAGE_HERE:

    If sMess <> vbNullString Then
        MsgBox sMess, vbExclamation + vbOKOnly, "Invalid Date Entry"
        oDateBox.Text = vbNullString
    Else
        ValidDateWhenCloseDate = True
    End If
    'cleanup
    Set oDateBox = Nothing
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub CheckMyDateWhenCloseDateSet"
End Function

Private Sub cmdRequestApproval_Click()
    On Error GoTo EH
    
    If mitmXSelected Is Nothing Then
        Exit Sub
    End If
    
    If RequestClaimApproval(cmdRequestApproval, lvwAssignments) Then
        RefreshMe
    End If
    
    Exit Sub
EH:
   goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdRequestApproval_Click"
End Sub


Public Function RequestClaimApproval(pocmdRequestApproval As Object, _
                                     Optional polvwAssignments As Object, _
                                     Optional polvwPackageItem As Object) As Boolean
    On Error GoTo EH
    Dim sSQL As String
    Dim oConn As ADODB.Connection
    Dim lRecordsAffected As Long
    Dim bCancelPreviousRequest As Boolean
    Dim sStatusID As String
    Dim sAssignmentsID As String
    Dim RS As ADODB.Recordset
    Dim sMess As String
    Dim lRet As Long
    Dim sRet As String
    Dim sPass As String
    Dim ocmdRequestApproval As CommandButton
    Dim oListView As ListView
    Dim sTickCount As String
    Dim sErrorFileName As String
    Dim sErrorFullPath As String
    Dim sAdjName As String
    
    Screen.MousePointer = MousePointerConstants.vbHourglass

    Set ocmdRequestApproval = pocmdRequestApproval
    If Not polvwAssignments Is Nothing Then
        Set oListView = polvwAssignments
    End If
    If Not polvwPackageItem Is Nothing Then
        Set oListView = polvwPackageItem
    End If
    
    If oListView Is Nothing Then
        GoTo CLEAN_UP
    End If
    
    ocmdRequestApproval.Enabled = False
    
    sMess = "Are you sure you want to flag this claim for Approval?" & vbCrLf & vbCrLf
    sMess = sMess & "Once this claim is flagged for approval," & vbCrLf
    sMess = sMess & "you will no longer be able to make changes" & vbCrLf
    sMess = sMess & "to the entire claim until the following occurs:" & vbCrLf & vbCrLf
    sMess = sMess & "1. If the claim fails document integrity" & vbCrLf
    sMess = sMess & "(missing files, photos etc.)... " & vbCrLf
    sMess = sMess & "The claim will be rejected and you will have" & vbCrLf
    sMess = sMess & "access to make corrections." & vbCrLf & vbCrLf
    sMess = sMess & "2. If the claim passes document integrity check... " & vbCrLf
    sMess = sMess & "The claim must then pass manager approval." & vbCrLf
    sMess = sMess & "If the manager rejects the claim, you will have" & vbCrLf
    sMess = sMess & "access to make corrections. " & vbCrLf
    sMess = sMess & "If the manager approves the claim, it will be" & vbCrLf
    sMess = sMess & "scheduled for delivery to the client." & vbCrLf & vbCrLf & vbCrLf
    sMess = sMess & "3. Once the claim has been delivered..." & vbCrLf & vbCrLf
    sMess = sMess & "A.  You will have access to the claim again." & vbCrLf
    sMess = sMess & "However, you will not be able to send," & vbCrLf
    sMess = sMess & "previously delivered documents unless those documents" & vbCrLf
    sMess = sMess & "are rejected at some point by the client." & vbCrLf & vbCrLf
    sMess = sMess & "B.  You will be able to add new documents and make" & vbCrLf
    sMess = sMess & "subsequent approval requests for those new documents." & vbCrLf
    

    If oListView.ListItems.Count = 0 Then
        sMess = "No items to approve!"
        MsgBox sMess, vbExclamation, "Nothing to approve"
        ocmdRequestApproval.Enabled = True
        GoTo CLEAN_UP
    End If
    
    'Need Get the Package Items and Verify
    'Verify that the Status is not already in the Approval stage
    If Not AllowThisStatus(sMess) Then
        If InStr(1, sMess, "[CHECK_FOR_UNDO_APPROVALRequest]", vbTextCompare) > 0 Then
            'Strip out the UNDO flag
            sMess = Replace(sMess, "[CHECK_FOR_UNDO_APPROVALRequest]", vbNullString, , , vbTextCompare)
            'If this item is still marked for Upload then allow the User to
            'Undo the previous request for Approval!
            If goUtil.GetFlagFromText(mitmXSelected.SubItems(GuiAssignments.UpLoadMe - 1)) Then
                sMess = sMess & vbCrLf
                sMess = sMess & "This request has not been sent yet." & vbCrLf
                sMess = sMess & "Do you want to cancel this request?"
                bCancelPreviousRequest = True
            End If
        End If
        If Not bCancelPreviousRequest Then
            MsgBox sMess, vbExclamation, "Nothing to approve"
            ocmdRequestApproval.Enabled = True
            GoTo CLEAN_UP
        End If
    End If
    
    If bCancelPreviousRequest Then
        sRet = InputBox(sMess, "Enter Password to proceed!", "Enter your Easy Claim password!")
        If sRet = vbNullString Then
            ocmdRequestApproval.Enabled = True
            GoTo CLEAN_UP
        End If
        sPass = goUtil.utGetECSCryptSetting("ECS", "WEB_SECURITY", "PASSWORD")
    End If

    If StrComp(sPass, sRet, vbTextCompare) <> 0 Then
        sMess = "Invalid Password!"
        MsgBox sMess, vbExclamation, "Invalid Password"
        ocmdRequestApproval.Enabled = True
        GoTo CLEAN_UP
    End If
    sAssignmentsID = mitmXSelected.ListSubItems(GuiAssignments.RKey - 1)
    'Need to Verify the Integrity of the package Items locally.
    If Not bCancelPreviousRequest Then
        If Not VerifyIntegrity(sAssignmentsID, sMess) Then
'                        MsgBox sMess, vbExclamation, "File Integrity Failed!"
            sTickCount = goUtil.utGetTickCount
            sAdjName = GetSetting(goUtil.gsAppEXEName, "GENERAL", "ADJUSTOR_FIRST_NAME", msClassName)
            sAdjName = sAdjName & "_"
            sAdjName = sAdjName & GetSetting(goUtil.gsAppEXEName, "GENERAL", "ADJUSTOR_LAST_NAME", vbNullString)
            sErrorFileName = sAdjName & "_" & sTickCount & "_" & sAssignmentsID & "_FileIntegrity_ErrorReport.txt"
            sErrorFullPath = GetSetting("ECS", "Dir", "ERRORLOG_DIR", App.Path)
            sErrorFullPath = sErrorFullPath & "\ErrorLog\" & sErrorFileName
            goUtil.utSaveFileData sErrorFullPath, sMess
            lRet = goUtil.utShellExecute(GetDesktopWindow, "OPEN", sErrorFullPath, vbNullString, App.Path, vbNormalFocus, True, True, True)
            ocmdRequestApproval.Enabled = True
            GoTo CLEAN_UP
        End If
    End If
    'Get the Appropriate StatusID
    If bCancelPreviousRequest Then
        sStatusID = GetStatusID("NEW")
    Else
        sStatusID = GetStatusID("APPROVALRequest")
    End If
    
    If Not SetRequestClaimApproval(sStatusID, sAssignmentsID) Then
        ocmdRequestApproval.Enabled = True
    End If
    
    RequestClaimApproval = True
    
CLEAN_UP:
    Screen.MousePointer = MousePointerConstants.vbDefault
    Set RS = Nothing
    Set oConn = Nothing
    Set ocmdRequestApproval = Nothing
    Set oListView = Nothing
    
    Exit Function
EH:
    Screen.MousePointer = MousePointerConstants.vbDefault
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function RequestClaimApproval"
End Function


Public Function VerifyIntegrity(psAssignmentsID As String, psMess As String) As Boolean
    On Error GoTo EH
    Dim sSQL As String
    Dim oConn As ADODB.Connection
    Dim RS As ADODB.Recordset
    Dim sMess As String
    Dim sReportFormat As String
    Dim sJPEGFileName As String
    Dim sPDFFileName As String
    Dim sPhotoReposPath As String
    Dim sAttachReposPath As String
    Dim FI As V2ECKeyBoard.FILE_INFORMATION
    Dim FA As V2ECKeyBoard.FILE_ATTRIBUTES
    Dim bPrintActiveReport As Boolean
    Dim bVerifyPhoto As Boolean
    Dim sTickCount As String
    Dim sTempFileDir As String
    Dim WddxFilePath As String
    Dim WddxFileName As String
    Dim sWddxData As String
    Dim oWddxSer As WDDXDeserializer
    Dim oWddxStruct As WDDXStruct
    Dim oWddxPhotoRS As WDDXRecordset
    Dim oWddxDataRS As WDDXRecordset
    Dim sPropNames As String
    Dim sTemp As String
    'Package Doc Sequence
    Dim lDocBase As Long
    Dim lThisDocNo As Long
    Dim lShouldBeDocNo As Long
    Dim sDocNo As String
    Dim sDocFail As String
    'Photo Sequence
    Dim lPhotoBase As Long
    Dim lThisPhotoNo As Long
    Dim lShouldBePhotoNo As Long
    Dim sPhotoNo As String
    Dim sPhotoFail As String
    Dim lRSPos As Long
    Dim lDocRsPos As Long
    
    
    'Be sure Temp File Dir Exists
    If Not goUtil.CreateTempDir(sTempFileDir) Then
        sTempFileDir = App.Path
    Else
        sTempFileDir = "C:\Temp\"
    End If
    
    sTickCount = goUtil.utGetTickCount
    WddxFilePath = sTempFileDir
    
    VerifyIntegrity = True
    
    'Set the Attach and Photo Repository dirs
    sPhotoReposPath = goUtil.PhotoReposPath
    sAttachReposPath = goUtil.AttachReposPath
    
    'Need to get the Package itmes for this Assingnment
    sSQL = "SELECT "
    sSQL = sSQL & "PI.[PackageItemID], "
    sSQL = sSQL & "PI.[PackageID], "
    sSQL = sSQL & "PI.[AssignmentsID], "
    sSQL = sSQL & "PI.[ID], "
    sSQL = sSQL & "PI.[IDPackage], "
    sSQL = sSQL & "PI.[IDAssignments], "
    sSQL = sSQL & "PI.[ReportFormat], "
    sSQL = sSQL & "PI.[RTAttachmentsID], "
    sSQL = sSQL & "PI.[IDRTAttachments], "
    sSQL = sSQL & "PI.[Number], "
    sSQL = sSQL & "PI.[AttachmentName], "
    sSQL = sSQL & "PI.[SortOrder], "
    sSQL = sSQL & "PI.[Name], "
    sSQL = sSQL & "PI.[Description], "
    sSQL = sSQL & "PI.[IsCoApprove], "
    sSQL = sSQL & "PI.[CoApproveDate], "
    sSQL = sSQL & "PI.[CoApproveDesc], "
    sSQL = sSQL & "PI.[IsClientCoReject], "
    sSQL = sSQL & "PI.[ClientCoRejectDate], "
    sSQL = sSQL & "PI.[ClientCoRejectDesc], "
    sSQL = sSQL & "PI.[IsClientCoDelete], "
    sSQL = sSQL & "PI.[ClientCoDeleteDate], "
    sSQL = sSQL & "PI.[ClientCoDeleteDesc], "
    sSQL = sSQL & "PI.[IsClientCoApprove], "
    sSQL = sSQL & "PI.[ClientCoApproveDate], "
    sSQL = sSQL & "PI.[ClientCoApproveDesc], "
    sSQL = sSQL & "PI.[PackageItemGUID], "
    sSQL = sSQL & "PI.[SendMe], "
    sSQL = sSQL & "PI.[SentDate], "
    sSQL = sSQL & "PI.[IsDeleted], "
    sSQL = sSQL & "PI.[DownLoadMe], "
    sSQL = sSQL & "PI.[UpLoadMe], "
    sSQL = sSQL & "PI.[AdminComments], "
    sSQL = sSQL & "PI.[DateLastUpdated], "
    sSQL = sSQL & "PI.[UpdateByUserID]"
    sSQL = sSQL & "FROM PackageItem PI "
    sSQL = sSQL & "WHERE    [AssignmentsID] = " & psAssignmentsID & " "
    sSQL = sSQL & "AND      [IsDeleted] = False "
    sSQL = sSQL & "ORDER BY [SortOrder] "
    
    Set oConn = New ADODB.Connection
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    Set RS = New ADODB.Recordset
    
    
    RS.CursorLocation = adUseClient
    RS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set RS.ActiveConnection = Nothing
    
    If RS.RecordCount = 0 Then
        sMess = "No Records Found!"
        VerifyIntegrity = False
        GoTo CLEAN_UP
    Else
        sMess = String(36, "-") & "File Integrity Failure Report " & String(36, "-") & vbCrLf
        sMess = sMess & Now() & vbCrLf
        RS.MoveFirst
        Do Until RS.EOF
            lDocRsPos = lDocRsPos + 1
            'First Check the Sequence for this doc see if it is not saved with correct sort number
            sDocNo = goUtil.IsNullIsVbNullString(RS.Fields("SortOrder"))
            'Make sure the photo Number matches the Sequence.
            'if not then user needs to save the current sort!
            lDocBase = 1
            lThisDocNo = CLng(sDocNo)
            lShouldBeDocNo = lDocBase + (lDocRsPos - 1)
            If lThisDocNo <> lShouldBeDocNo Then
                VerifyIntegrity = False
                sDocFail = vbCrLf & "Invalid Document Sequence / Sort order!" & vbCrLf
                sDocFail = sDocFail & "Please make sure Documents are in order, then save sort order."
                sMess = sMess & BuildFailureReason(RS, sDocFail)
                GoTo CLEAN_UP
            End If
            '.1. If the Item has an External file Associated (PhotoReport, DiagramReport, or is an actual PDF Attachment)
            '   Does the file exist where it is suppose to.  IE did the adjuster manipulate the file
            '   Outside of Easy Claim.  Or was there some problem on the adjusters box that molested the
            'file ?  Is the File zero Length.
            sReportFormat = goUtil.IsNullIsVbNullString(RS.Fields("ReportFormat"))
            If InStr(1, sReportFormat, "_arRptPhotos", vbTextCompare) > 0 Then
                'Need to create Wddx Packet that conatins the list of photos
                'Build the WddxFileName
                WddxFileName = "Temp_" & sTickCount & "_Photos.xml"
                bPrintActiveReport = PrintActiveReport(Nothing, , vbNullString, False, WddxFilePath, WddxFileName, True, True, sReportFormat)
                'associated with this report and verify they exist and file len > 0
                'As well need to verify that the sort order is correct.
                If Not bPrintActiveReport Then
                    VerifyIntegrity = False
                    sMess = sMess & BuildFailureReason(RS, "System Error while Verifying Integrity!")
                    GoTo NEXT_RS
                Else
                    bVerifyPhoto = True
                    sPhotoFail = vbNullString
                    sWddxData = goUtil.utGetFileData(WddxFilePath & WddxFileName)
                    Set oWddxSer = New WDDXDeserializer
                    Set oWddxStruct = oWddxSer.deserialize(sWddxData)
                    sPropNames = Join(oWddxStruct.getPropNames, "|")
                    If InStr(1, sPropNames, "PhotosRS", vbTextCompare) = 0 Then
                        VerifyIntegrity = False
                        sPhotoFail = "Empty Photo Report! Please remove from package."
                        sMess = sMess & BuildFailureReason(RS, sPhotoFail)
                        GoTo NEXT_RS
                    ElseIf InStr(1, sPropNames, "DataRS", vbTextCompare) = 0 Then
                        VerifyIntegrity = False
                        sPhotoFail = "Error Reading Data Recordset!"
                        sMess = sMess & BuildFailureReason(RS, sPhotoFail)
                        GoTo NEXT_RS
                    End If
                    Set oWddxPhotoRS = oWddxStruct.getProp("PhotosRS")
                    Set oWddxDataRS = oWddxStruct.getProp("DataRS")
                    'Loop through each Photo and determine if every thing is Kosher
                    For lRSPos = 1 To oWddxPhotoRS.getRowCount
                        sJPEGFileName = oWddxPhotoRS.getField(lRSPos, "imgPhotoPath")
                        sPhotoNo = oWddxPhotoRS.getField(lRSPos, "fPhotoNo")
                        'Make sure the photo Number matches the Sequence.
                        'if not then user needs to save the current sort!
                        sTemp = oWddxDataRS.getField(1, "f_Description")
                        sTemp = left(sTemp, 3)
                        lPhotoBase = CLng(sTemp)
                        lThisPhotoNo = CLng(sPhotoNo)
                        lShouldBePhotoNo = lPhotoBase + (lRSPos - 1)
                        If lThisPhotoNo <> lShouldBePhotoNo Then
                            VerifyIntegrity = False
                            sPhotoFail = vbCrLf & "Invalid Photo Sequence / Sort order!" & vbCrLf
                            sPhotoFail = sPhotoFail & "Please make sure photos are in order, then save sort order."
                            sMess = sMess & BuildFailureReason(RS, sPhotoFail)
                            GoTo NEXT_RS
                        End If
                        
                        If Not goUtil.utFileExists(sJPEGFileName) Then
                            bVerifyPhoto = False
                            sJPEGFileName = vbCrLf & "Photo No: [" & sPhotoNo & "] File Not Found!: " & sJPEGFileName
                            sPhotoFail = sPhotoFail & sJPEGFileName
                            GoTo NEXT_PHOTO
                        Else
                            'Need to ensure that the pdf file exists and the file len > 0
                            GetFileSettings sJPEGFileName, FI, FA
                            If FI.nFileSize = 0 Then
                                bVerifyPhoto = False
                                sJPEGFileName = vbCrLf & "Photo No: [" & sPhotoNo & "] Invalid File Size!: " & sJPEGFileName
                                sPhotoFail = sPhotoFail & sJPEGFileName
                                GoTo NEXT_PHOTO
                            End If
                        End If
NEXT_PHOTO:
                    Next
                    If Not bVerifyPhoto Then
                        VerifyIntegrity = False
                        sMess = sMess & BuildFailureReason(RS, sPhotoFail)
                        GoTo NEXT_RS
                    End If
                End If
                
            ElseIf InStr(1, sReportFormat, ".pdf|", vbTextCompare) > 0 Then
                'PDF ATTACHMENT
                'FRE27745_050608152131_1.pdf|
                'Need to ensure that the pdf file exists and the file len > 0
                sPDFFileName = Trim(Right(sReportFormat, 200))
                sPDFFileName = left(sPDFFileName, Len(sPDFFileName) - 1)
                sPDFFileName = sAttachReposPath & sPDFFileName
                If Not goUtil.utFileExists(sPDFFileName) Then
                    VerifyIntegrity = False
                    sPDFFileName = vbCrLf & "File Not Found!: " & sPDFFileName
                    sMess = sMess & BuildFailureReason(RS, sPDFFileName)
                    GoTo NEXT_RS
                Else
                    'Need to ensure that the pdf file exists and the file len > 0
                    GetFileSettings sPDFFileName, FI, FA
                    If FI.nFileSize = 0 Then
                        VerifyIntegrity = False
                        sPDFFileName = vbCrLf & "Invalid File Size!: " & sPDFFileName
                        sMess = sMess & BuildFailureReason(RS, sPDFFileName)
                        GoTo NEXT_RS
                    End If
                End If
            ElseIf InStr(1, sReportFormat, "_arWorkSheetDiag|", vbTextCompare) > 0 Then
                'DIAGRAM WORKSHEET
                'ECrptFarmers_arWorkSheetDiag|clsLists|1|1|FRE27745_050608155648.jpg
                'need to verify the jpg photo exists and file len > 0
                sJPEGFileName = Trim(Right(sReportFormat, 200))
                sJPEGFileName = Mid(sJPEGFileName, InStrRev(sJPEGFileName, "|", , vbBinaryCompare) + 1)
                sJPEGFileName = sPhotoReposPath & sJPEGFileName
                If Not goUtil.utFileExists(sJPEGFileName) Then
                    VerifyIntegrity = False
                    sJPEGFileName = vbCrLf & "File Not Found!: " & sJPEGFileName
                    sMess = sMess & BuildFailureReason(RS, sJPEGFileName)
                    GoTo NEXT_RS
                Else
                    'Need to ensure that the pdf file exists and the file len > 0
                    GetFileSettings sJPEGFileName, FI, FA
                    If FI.nFileSize = 0 Then
                        VerifyIntegrity = False
                        sJPEGFileName = vbCrLf & "Invalid File Size!: " & sJPEGFileName
                        sMess = sMess & BuildFailureReason(RS, sJPEGFileName)
                        GoTo NEXT_RS
                    End If
                End If
            End If
            'Check each item for the following...
            '2. Is there an Item still flagged for Upload!
            '   If there is, then this will be cause to fail integrity becuase it
            '   Does not yet Exist on the Web Server
            If CBool(RS.Fields("UpLoadMe")) Then
                VerifyIntegrity = False
                sMess = sMess & BuildFailureReason(RS, "RECORD NEEDS TO BE UPLOADED!")
                GoTo NEXT_RS
            End If
NEXT_RS:
            RS.MoveNext
        Loop
    End If
    
CLEAN_UP:
    psMess = sMess
    Set RS = Nothing
    Set oConn = Nothing
    Set oWddxSer = Nothing
    Set oWddxStruct = Nothing
    Set oWddxPhotoRS = Nothing
    Set oWddxDataRS = Nothing
    Exit Function
EH:
    VerifyIntegrity = False
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function VerifyIntegrity"
    psMess = "System Error while Verifying Integrity!"
End Function

Private Function BuildFailureReason(pRS As ADODB.Recordset, sFailureReason As String) As String
    On Error GoTo EH
    Dim sMess As String
    
    sMess = String(102, "-") & vbCrLf
    sMess = sMess & "[Sort]" & String(4, vbTab) & goUtil.IsNullIsVbNullString(pRS.Fields("SortOrder")) & vbCrLf
    sMess = sMess & "[Name]" & String(4, vbTab) & goUtil.IsNullIsVbNullString(pRS.Fields("Name")) & vbCrLf
    sMess = sMess & "[Description]" & String(3, vbTab) & goUtil.IsNullIsVbNullString(pRS.Fields("Description")) & vbCrLf
    sMess = sMess & "[AttachmentName]" & String(2, vbTab) & goUtil.IsNullIsVbNullString(pRS.Fields("AttachmentName")) & vbCrLf
    sMess = sMess & "[Failure Reason]" & String(2, vbTab) & sFailureReason & vbCrLf
    sMess = sMess & String(102, "-") & vbCrLf & vbCrLf
    
    BuildFailureReason = sMess
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Function BuildFailureReason"
End Function

Public Function AllowThisStatus(psMess As String) As Boolean
    On Error GoTo EH
    Dim sSQL As String
    Dim sDate As String
    Dim sMess As String
    Dim sStatusAlias As String
    Dim sGetStatusAlias As String
    Dim oConn As ADODB.Connection
    Dim RS As ADODB.Recordset
    
    'Don't allow certain claims to be viewed
    'check for Deleted item
    If goUtil.GetFlagFromText(mitmXSelected.SubItems(GuiAssignments.IsDeleted - 1)) Then
        sMess = sMess & mitmXSelected.SubItems(GuiAssignments.IBNUM - 1) & " "
        sMess = sMess & "Is Deleted!" & vbCrLf
        GoTo CLEAN_UP
    End If
    'check for Locked item
    If goUtil.GetFlagFromText(mitmXSelected.SubItems(GuiAssignments.Islocked - 1)) Then
        sMess = sMess & mitmXSelected.SubItems(GuiAssignments.IBNUM - 1) & " "
        sMess = sMess & "Is Locked!" & vbCrLf
        GoTo CLEAN_UP
    End If
    'Check for Reassigned Item
    sDate = mitmXSelected.SubItems(GuiAssignments.DateReassigned - 1)
    If IsDate(sDate) Then
        If CDate(sDate) <> NULL_DATE And CDate(sDate) <> CDate("1/1/1900") Then
            sMess = sMess & mitmXSelected.SubItems(GuiAssignments.IBNUM - 1) & " "
            sMess = sMess & "Has Been Reassigned!" & vbCrLf
            GoTo CLEAN_UP
        End If
    End If
    'Check for certain status
    sStatusAlias = mitmXSelected.SubItems(GuiAssignments.Status - 1)
    Set oConn = New ADODB.Connection
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name

    'Check for Approval in progress status
    GetStatusID "APPROVALRequest", True, sGetStatusAlias
    
    If StrComp(sGetStatusAlias, sStatusAlias, vbTextCompare) = 0 Then
        sMess = sMess & mitmXSelected.SubItems(GuiAssignments.Status - 1) & " "
        sMess = sMess & "In Progress!" & vbCrLf
        sMess = sMess & "[CHECK_FOR_UNDO_APPROVALRequest]"
        GoTo CLEAN_UP
    End If

    'Check for (PENDINGDelivery) Status
    GetStatusID "PENDINGDelivery", True, sGetStatusAlias
    
    If StrComp(sGetStatusAlias, sStatusAlias, vbTextCompare) = 0 Then
        sMess = sMess & mitmXSelected.SubItems(GuiAssignments.Status - 1) & " "
        sMess = sMess & "In Progress!" & vbCrLf
        GoTo CLEAN_UP
    End If
    
    If sMess = vbNullString Then
       AllowThisStatus = True
    Else
CLEAN_UP:
        psMess = sMess
    End If
    
    Set oConn = Nothing
    Set RS = Nothing
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function AllowThisStatus"
End Function

Public Function SetRequestClaimApproval(psStatusID As String, _
                                        psAssignmentsID As String) As Boolean
    On Error GoTo EH
    Dim sSQL As String
    Dim oConn As ADODB.Connection
    Dim RS As ADODB.Recordset
    Dim sDateLastUpdated As String
    Dim lRecordsAffected As Long
    
    
    sDateLastUpdated = Now()
    
    Set oConn = New ADODB.Connection
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    sSQL = "UPDATE ASSIGNMENTS SET  "
    sSQL = sSQL & "[StatusID] = " & psStatusID & ", "
    sSQL = sSQL & "[UpLoadMe] = True, "
    sSQL = sSQL & "[DateLastUpdated] = #" & sDateLastUpdated & "#, "
    sSQL = sSQL & "[UpdateByUserID] = " & goUtil.gsCurUsersID & " "
    sSQL = sSQL & "WHERE [ID] = " & psAssignmentsID & " "
    sSQL = sSQL & "AND [AssignmentsID] = " & psAssignmentsID & " "
    
    oConn.Execute sSQL, lRecordsAffected
    
    SetRequestClaimApproval = CBool(lRecordsAffected)
    
    'cleanup
    Set oConn = Nothing
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function SetRequestClaimApproval"
End Function

Public Function GetStatusID(psStatusName As String, _
                            Optional pbCheckAlias As Boolean, _
                            Optional psStatusAlias As String) As String
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim RS As ADODB.Recordset
    Dim sSQL As String
    
    Set oConn = New ADODB.Connection
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    Set RS = New ADODB.Recordset
    
    sSQL = "SELECT  [StatusID], "
    sSQL = sSQL & "[Status], "
    sSQL = sSQL & "[StatusAlias] "
    sSQL = sSQL & "FROM Status "
    sSQL = sSQL & "WHERE [IsDeleted] = 0 "
    sSQL = sSQL & "AND [Status] = '" & goUtil.utCleanSQLString(psStatusName) & "' "
    
    RS.CursorLocation = adUseClient
    RS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set RS.ActiveConnection = Nothing
    
    If RS.RecordCount = 1 Then
        If pbCheckAlias Then
            psStatusAlias = RS.Fields("StatusAlias").Value
        End If
        GetStatusID = RS.Fields("StatusID").Value
    End If
    
CLEAN_UP:
    Set RS = Nothing
    Set oConn = Nothing
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function GetStatusID"
End Function

Public Function PrintActiveReport(poReportItem As Object, _
                                Optional piMode As VBRUN.FormShowConstants = vbModeless, _
                                Optional psCopyName As String = vbNullString, _
                                Optional pbPrintPreview As Boolean = True, _
                                Optional psSaveToFilePath As String, _
                                Optional psSaveToFileName As String, _
                                Optional pbExportXML As Boolean, _
                                Optional pbExportXMLOnly As Boolean, _
                                Optional psReportFormat As String) As Boolean
    On Error GoTo EH
    Dim sParams As String
    Dim sReportName As String
    Dim sReportTitle As String
    Dim srptProjectName As String
    Dim srptClassName As String
    Dim lrptVersion As Long
    Dim sData As String
    Dim saryData() As String
    Dim ocboReport As Object
    Dim itmXReport As ListItem
    Dim MyActReport As ActiveReport
    Dim oCarList As V2ECKeyBoard.clsCarLists
    'If using Adobe PDF Viewer
    Dim sPDFFilePath As String
    Dim sPrintPreview As String
    Dim bUseAdobeReader As Boolean
    'Some Reports need extra Params passed to them
    'Payments
    Dim sRTChecksID As String
    Dim sCheckNum As String
    'Internal Billing
    Dim sIBID As String
    Dim sSupplement As String
    'Photo Reports (Multi Report)
    Dim sPhotoReportNumber As String
    'Worksheet Diagram (Multi Report)
    Dim sDiagramNumber As String
    Dim sNumber As String
    'Loss Report
    Dim oLR As V2ECKeyBoard.clsLossReports
    Dim MyAssignmentsRS As ADODB.Recordset
    Dim sLRFormat As String
    Dim sLossReport As String
    Dim sLRData As String
    'Export to XML FileName
    Dim sXMLFilePath As String
    Dim sXMLFileName As String
    
    sPrintPreview = GetSetting(goUtil.gsMainAppEXEName, "GENERAL", "PRINT_PREVIEW", "USE_ADOBE")
    
    Select Case UCase(sPrintPreview)
        Case "USE_ADOBE"
            bUseAdobeReader = True
    End Select
    
    'if saving to file path always use adobe reader
    If psSaveToFilePath <> vbNullString Then
        bUseAdobeReader = True
    End If
    
    If psReportFormat <> vbNullString Then
        sData = psReportFormat
    ElseIf TypeOf poReportItem Is ListBox Or TypeOf poReportItem Is ComboBox Then
        Set ocboReport = poReportItem
        sData = ocboReport.Text
    ElseIf TypeOf poReportItem Is ListItem Then
        Set itmXReport = poReportItem
        sData = itmXReport.ListSubItems(GuiPackageItemListView.ReportFormat - 1)
    Else
        Exit Function
    End If
    
    If sData <> vbNullString Then
        sReportTitle = Trim(left(sData, 200))
        goUtil.utCleanFileFolderName sReportTitle, False
        sData = Mid(sData, InStr(1, sData, String(200, " "), vbBinaryCompare))
        sData = Trim(sData)
        saryData() = Split(sData, "|", , vbBinaryCompare)
        If UBound(saryData, 1) <= 1 Then
            'Check for Loss Report
            sLRFormat = saryData(0)
            If StrComp(sLRFormat, "LRFormat", vbTextCompare) = 0 Then
                Me.SetadoRSAssignments msAssignmentsID
                Set MyAssignmentsRS = Me.adoRSAssignments
                sLRFormat = goUtil.IsNullIsVbNullString(MyAssignmentsRS.Fields("LRFormat"))
                sLossReport = goUtil.IsNullIsVbNullString(MyAssignmentsRS.Fields("LossReport"))
                
                If InStr(1, sLRFormat, "OLEType_pdf", vbTextCompare) > 0 Then
                    sPDFFilePath = goUtil.AttachReposPath & sLossReport
                    If psSaveToFilePath <> vbNullString Then
                        'Do not do Loss Report if Export to xml only is true
                        If pbExportXML And pbExportXMLOnly Then
                            PrintActiveReport = False
                            Set ocboReport = Nothing
                            Set itmXReport = Nothing
                            Set oLR = Nothing
                            Set MyAssignmentsRS = Nothing
                            MsgBox "Loss Reports can not be part of and XML ONLY Export!", vbExclamation
                            Exit Function
                        End If
                        goUtil.utCopyFile sPDFFilePath, psSaveToFilePath & psSaveToFileName
                    Else
                        If pbPrintPreview Then
                            goUtil.utShellExecute , "OPEN", sPDFFilePath, , , vbNormalFocus, True, True, True, sReportTitle
                        Else
                            goUtil.utShellExecute , "PRINT", sPDFFilePath, , , vbNormalFocus, True, True, True, sReportTitle
                        End If
                        DoEvents
                        Sleep 100
                    End If
                Else
                    sPDFFilePath = goUtil.gsInstallDir & "\TempLossReport" & goUtil.utGetTickCount & ".pdf"
                    Set oLR = New V2ECKeyBoard.clsLossReports
                    If StrComp(sLRFormat, "TEXT", vbTextCompare) <> 0 Then
                        sLRData = sLRFormat & vbCrLf & sLossReport
                    Else
                        sLRData = sLossReport
                    End If
                    oLR.CreateExport sLRData, sPDFFilePath, ARPdf
                    If psSaveToFilePath <> vbNullString Then
                        'Do not do Loss Report if Export to xml only is true
                        If pbExportXML And pbExportXMLOnly Then
                            PrintActiveReport = False
                            Set ocboReport = Nothing
                            Set itmXReport = Nothing
                            Set oLR = Nothing
                            Set MyAssignmentsRS = Nothing
                            MsgBox "Loss Reports can not be part of and XML ONLY Export!", vbExclamation
                            Exit Function
                        End If
                        goUtil.utCopyFile sPDFFilePath, psSaveToFilePath & psSaveToFileName
                    Else
                        If pbPrintPreview Then
                            goUtil.utShellExecute , "OPEN", sPDFFilePath, , , vbNormalFocus, True, True, True, sReportTitle
                        Else
                            goUtil.utShellExecute , "PRINT", sPDFFilePath, , , vbNormalFocus, True, True, True, sReportTitle
                        End If
                        DoEvents
                        Sleep 100
                    End If
                    goUtil.utDeleteFile sPDFFilePath
                End If
                PrintActiveReport = True
                GoTo CLEAN_UP
            End If
            
            sPDFFilePath = saryData(0)
            If InStr(1, sPDFFilePath, ".pdf", vbTextCompare) > 0 Then
                sPDFFilePath = goUtil.AttachReposPath & sPDFFilePath
                'Check for Pdf Attachment file
                If psSaveToFilePath <> vbNullString Then
                    'Do not do Attachments if Export to xml only is true
                    If pbExportXML And pbExportXMLOnly Then
                        PrintActiveReport = False
                        Set ocboReport = Nothing
                        Set itmXReport = Nothing
                        Set oLR = Nothing
                        Set MyAssignmentsRS = Nothing
                        MsgBox "Attachments can not be part of and XML ONLY Export!", vbExclamation
                        Exit Function
                    End If
                    goUtil.utCopyFile sPDFFilePath, psSaveToFilePath & psSaveToFileName
                Else
                    If pbPrintPreview Then
                        goUtil.utShellExecute , "OPEN", sPDFFilePath, , , vbNormalFocus, True, True, True, sReportTitle
                    Else
                        goUtil.utShellExecute , "PRINT", sPDFFilePath, , , vbNormalFocus, True, True, True, sReportTitle
                    End If
                    DoEvents
                    Sleep 100
                End If
            End If
            PrintActiveReport = True
            GoTo CLEAN_UP
        End If
        srptProjectName = saryData(0)
        srptClassName = saryData(1)
        lrptVersion = saryData(2)
        'Check For Multi Reports Here
        If psReportFormat <> vbNullString Then
            If UBound(saryData, 1) >= 3 Then
                sNumber = saryData(3)
            End If
            
            'If this is coming from the Package Screen need to populate the Number for certain reports
            If InStr(1, srptProjectName, "_arRptPhotos", vbTextCompare) > 0 Then
                sPhotoReportNumber = sNumber
            ElseIf InStr(1, srptProjectName, "_arWorkSheetDiag", vbTextCompare) > 0 Then
                sDiagramNumber = sNumber
            ElseIf InStr(1, srptProjectName, "_arRptAddlChk", vbTextCompare) > 0 Then
                sCheckNum = sNumber
            ElseIf InStr(1, srptProjectName, "_arRptIB", vbTextCompare) > 0 Then
                sSupplement = sNumber
            End If
        ElseIf TypeOf poReportItem Is ListItem Then
            If UBound(saryData, 1) >= 3 Then
                sNumber = saryData(3)
            End If
            
            'If this is coming from the Package Screen need to populate the Number for certain reports
            If InStr(1, srptProjectName, "_arRptPhotos", vbTextCompare) > 0 Then
                sPhotoReportNumber = sNumber
            ElseIf InStr(1, srptProjectName, "_arWorkSheetDiag", vbTextCompare) > 0 Then
                sDiagramNumber = sNumber
            ElseIf InStr(1, srptProjectName, "_arRptAddlChk", vbTextCompare) > 0 Then
                sCheckNum = sNumber
            ElseIf InStr(1, srptProjectName, "_arRptIB", vbTextCompare) > 0 Then
                sSupplement = sNumber
            End If
        ElseIf InStr(1, srptProjectName, "_arRptPhotos", vbTextCompare) > 0 Then
            'Photo Reports (Multi Report)
            sPhotoReportNumber = saryData(3)
        ElseIf InStr(1, srptProjectName, "_arWorkSheetDiag", vbTextCompare) > 0 Then
            'Worksheet Diagram (Multi Report)
            sDiagramNumber = saryData(3)
        End If
    Else
        Exit Function
    End If
    
    'Build Params List to be passed in to Create Report Object
    'This Object will have list of Report Parameters it requires
    
    sParams = vbNullString
    sParams = sParams & "psAssignmentsID=" & msAssignmentsID & "|"
    'If using Adobe PDF Viewer
    If bUseAdobeReader Then
        sPDFFilePath = goUtil.gsInstallDir & "\TempActiveReport" & goUtil.utGetTickCount & ".pdf"
        sParams = sParams & "psXportPath=" & sPDFFilePath & "|"
        sParams = sParams & "pPDFJPEGQuality=" & "50" & "|"
        sParams = sParams & "pXportType=" & ExportType.ARPdf & "|"
    Else
        sParams = sParams & "pbGetObjectOnly=" & "True" & "|"
    End If
    
    'Certain Reports Need to have some more Params Passed in
    If InStr(1, srptProjectName, "_arRptAddlChk", vbTextCompare) > 0 Then
        'Need to Get the ChecksID and Check Number
        If Not ocboReport Is Nothing Then
            sRTChecksID = CStr(ocboReport.ItemData(ocboReport.ListIndex))
            If Not GetPaymentsParams(sRTChecksID, sCheckNum) Then
                GoTo CLEAN_UP
            End If
        ElseIf Not itmXReport Is Nothing Then
            'the schecknum was already set above
            If Not GetPaymentsParams(sRTChecksID, sCheckNum) Then
                GoTo CLEAN_UP
            End If
        End If
        sParams = sParams & "pRTChecksID=" & sRTChecksID & "|"
        sParams = sParams & "psCheckNum=" & sCheckNum & "|"
    ElseIf InStr(1, srptProjectName, "_arRptIB", vbTextCompare) > 0 Then
        'If the IBID and Supplement Parameters already exist then use them
        'Otherwise have to do Data Call to get em.
        If InStr(1, sData, "pIBID=", vbTextCompare) > 0 And InStr(1, sData, "pSupplement=", vbTextCompare) > 0 Then
            sParams = sParams & saryData(3) & "|"
            sParams = sParams & saryData(4) & "|"
            'Check for Report Title As Well
            If InStr(1, sData, "psReportTitle=", vbTextCompare) > 0 Then
                sReportTitle = Mid(saryData(5), InStr(1, saryData(5), "=", vbTextCompare) + 1)
            End If
        Else
            If Not ocboReport Is Nothing Then
                sIBID = CStr(ocboReport.ItemData(ocboReport.ListIndex))
                If Not GetIBParams(sIBID, sSupplement) Then
                    GoTo CLEAN_UP
                End If
            ElseIf Not itmXReport Is Nothing Then
                'The supplement was already set above
                If Not GetIBParams(sIBID, sSupplement) Then
                    GoTo CLEAN_UP
                End If
            End If
            sParams = sParams & "pIBID=" & sIBID & "|"
            sParams = sParams & "pSupplement=" & sSupplement & "|"
        End If
        sParams = sParams & "pCopyName=" & psCopyName & "|"
    ElseIf InStr(1, srptProjectName, "_arRptPhotos", vbTextCompare) > 0 Then
        'Photo Reports (Multi Report)
        sParams = sParams & "pNumber=" & sPhotoReportNumber & "|"
    ElseIf InStr(1, srptProjectName, "_arWorkSheetDiag", vbTextCompare) > 0 Then
        'Worksheet Diagram (Multi Report)
        sParams = sParams & "pNumber=" & sDiagramNumber & "|"
    End If
    
    sReportName = srptProjectName & "." & srptClassName

    If StrComp(psCopyName, "(-ALL COPIES-)", vbTextCompare) = 0 Then
        'Do a recursive call until All Copies are printed
        'Client company Copy
        If Not PrintActiveReport(poReportItem, , goUtil.gsCurCarDBName & " Copy", pbPrintPreview, psSaveToFilePath, psSaveToFileName) Then
            GoTo CLEAN_UP
        End If
        'Company Copy
        If Not PrintActiveReport(poReportItem, , GetSetting(goUtil.gsAppEXEName, "GENERAL", "CURRENT_COMPANY_NAME", "Company") & " Copy", pbPrintPreview, psSaveToFilePath, psSaveToFileName) Then
            GoTo CLEAN_UP
        End If
        'Remit Copy
        If Not PrintActiveReport(poReportItem, , "Remit Copy", pbPrintPreview, psSaveToFilePath, psSaveToFileName) Then
            GoTo CLEAN_UP
        End If
        'Adjuster Copy
        If Not PrintActiveReport(poReportItem, , "Adjuster Copy", pbPrintPreview, psSaveToFilePath, psSaveToFileName) Then
            GoTo CLEAN_UP
        End If
    Else
        Set oCarList = CreateObject(goUtil.goCurCarList.ClassName)
        If bUseAdobeReader Then
            'Add Export XML Parameters here
            If pbExportXML Then
                sParams = sParams & "pbExportXML=True|"
                If pbExportXMLOnly Then
                    sParams = sParams & "pbExportXMLOnly=True|"
                End If
            End If
            Set MyActReport = oCarList.GetARReport(sReportName, lrptVersion, sParams)
            If goUtil.utFileExists(sPDFFilePath) Or (pbExportXML And pbExportXMLOnly) Then
                If psSaveToFilePath <> vbNullString Then
                    If pbExportXML Then
                        If Not pbExportXMLOnly Then
                            If psCopyName <> vbNullString Then
                                goUtil.utCopyFile sPDFFilePath, psSaveToFilePath & Replace(psSaveToFileName, ".pdf", "_" & psCopyName & ".pdf", , 1, vbTextCompare)
                            Else
                                goUtil.utCopyFile sPDFFilePath, psSaveToFilePath & psSaveToFileName
                            End If
                        End If
                    Else
                        If psCopyName <> vbNullString Then
                            goUtil.utCopyFile sPDFFilePath, psSaveToFilePath & Replace(psSaveToFileName, ".pdf", "_" & psCopyName & ".pdf", , 1, vbTextCompare)
                        Else
                            goUtil.utCopyFile sPDFFilePath, psSaveToFilePath & psSaveToFileName
                        End If
                    End If
                    
                    If pbExportXML Then
                        'change the pdffile path the XML
                        sXMLFilePath = sPDFFilePath
                        sXMLFilePath = left(sXMLFilePath, InStrRev(sXMLFilePath, ".", , vbBinaryCompare))
                        sXMLFilePath = sXMLFilePath & "xml"
                        'Change the pdf to XML file path
                        sXMLFileName = psSaveToFileName
                        sXMLFileName = left(sXMLFileName, InStrRev(sXMLFileName, ".", , vbBinaryCompare))
                        sXMLFileName = sXMLFileName & "xml"
                       If psCopyName <> vbNullString Then
                            goUtil.utCopyFile sXMLFilePath, psSaveToFilePath & Replace(sXMLFileName, ".xml", "_" & psCopyName & ".xml", , 1, vbTextCompare)
                        Else
                            goUtil.utCopyFile sXMLFilePath, psSaveToFilePath & sXMLFileName
                        End If
                    End If
                Else
                    If pbPrintPreview Then
                        goUtil.utShellExecute , "OPEN", sPDFFilePath, , , vbNormalFocus, True, True, True, sReportTitle
                    Else
                        goUtil.utShellExecute , "PRINT", sPDFFilePath, , , vbNormalFocus, True, True, True, sReportTitle
                    End If
                    
                End If
                ' DoEvents
                Sleep 100
                goUtil.utDeleteFile sPDFFilePath
                goUtil.utDeleteFile sXMLFilePath
                If Not MyActReport Is Nothing Then
                    Unload MyActReport
                    Set MyActReport = Nothing
                End If
'                oCarList.CLEANUP
                Set oCarList = Nothing
            End If
        Else
            'Using Active Report Viewer
            Set MyActReport = oCarList.GetARReport(sReportName, lrptVersion, sParams)
        
            If pbPrintPreview Then
                If mArv Is Nothing Then
                    Set mArv = New V2ARViewer.clsARViewer
                    mArv.SetUtilObject goUtil
                End If
'                If Not moForm Is Nothing Then
'                    If StrComp(psCopyName, "(-ALL COPIES-)", vbTextCompare) <> 0 Then
'                        Unload moForm
'                        Set moForm = Nothing
'                    End If
'                End If
                With mArv
                    'Pass in true to have Active reports process on separate thread.
                    'This will allow the viewer to load while the report is processing
                    'false will force the report to run on single thread
                    MyActReport.Run False 'True
                    .objARvReport = MyActReport
                    .sRptTitle = sReportTitle
                    .HidePrintButton = False
                    .ShowReportOnForm moForm, piMode
        
                    Unload .objARvReport
                    Set .objARvReport = Nothing
                End With
            Else
                MyActReport.PrintReport False
            End If
            Unload MyActReport
            Set MyActReport = Nothing
'            oCarList.CLEANUP
            Set oCarList = Nothing
        End If
    End If
    PrintActiveReport = True
CLEAN_UP:
    'Cleanup
    Set ocboReport = Nothing
    Set itmXReport = Nothing
    Set oLR = Nothing
    Set MyAssignmentsRS = Nothing
    
    'Clear the Local ref to this report object only
    'The actual cleanup of this active report object will occur within gARV
    PrintActiveReport = True
    Exit Function
EH:
    Screen.MousePointer = vbDefault
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function PrintActiveReport"
End Function

Public Function SetadoRSAssignments(psIDAssignments As String) As Boolean
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    
    'rese the typeof loss rs
    If Not madoRSAssignments Is Nothing Then
        Set madoRSAssignments = Nothing
    End If
    
    Set oConn = New ADODB.Connection
    Set madoRSAssignments = New ADODB.Recordset
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    sSQL = "SELECT * "
    sSQL = sSQL & "FROM Assignments "
    sSQL = sSQL & "WHERE ID = " & psIDAssignments & " "
    
    madoRSAssignments.CursorLocation = adUseClient
    madoRSAssignments.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set madoRSAssignments.ActiveConnection = Nothing
    
    SetadoRSAssignments = True
    
    Set oConn = Nothing
    Exit Function
EH:
    Set oConn = Nothing
    SetadoRSAssignments = False
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function SetadoRSAssignments"
End Function

Public Function GetPaymentsParams(psRTChecksID As String, psCheckNum As String) As Boolean
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim RS As ADODB.Recordset
    Dim sSQL As String
    Dim sRTChecksID As String
    Dim sCheckNum As String
    
    Set oConn = New ADODB.Connection
    Set RS = New ADODB.Recordset
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    sSQL = "SELECT RTC.[RTChecksID], "
    sSQL = sSQL & "RTC.[CheckNum] "
    sSQL = sSQL & "FROM RTChecks RTC "
    sSQL = sSQL & "WHERE RTC.[AssignmentsID] = " & msAssignmentsID & " "
    If psRTChecksID = vbNullString Then
        sSQL = sSQL & "AND RTC.[CheckNum] = " & psCheckNum & " "
    Else
        sSQL = sSQL & "AND RTC.[RTChecksID] = " & psRTChecksID & " "
    End If
    
    RS.CursorLocation = adUseClient
    RS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set RS.ActiveConnection = Nothing
    
    If Not RS.EOF Then
        sRTChecksID = RS.Fields("RTChecksID").Value
        sCheckNum = RS.Fields("CheckNum").Value
    End If
    
    
    psRTChecksID = sRTChecksID
    psCheckNum = sCheckNum
    GetPaymentsParams = True
    
    'Cleanup
    Set oConn = Nothing
    Set RS = Nothing
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function GetPaymentsParams"
End Function

Public Function GetIBParams(psIBID As String, psSupplement As String) As Boolean
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim RS As ADODB.Recordset
    Dim sSQL As String
    Dim sIBID As String
    Dim sSupplement As String
    
    
    Set oConn = New ADODB.Connection
    Set RS = New ADODB.Recordset
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    sSQL = "SELECT IB.[IBID], "
    sSQL = sSQL & "BC.[Supplement] "
    sSQL = sSQL & "FROM IB "
    sSQL = sSQL & "INNER JOIN BillingCount BC ON IB.BillingCountID = BC.BillingCountID "
    sSQL = sSQL & "WHERE IB.[AssignmentsID] = " & msAssignmentsID & " "
    If psIBID = vbNullString Then
        sSQL = sSQL & "AND IB.[IB14a_sSupplement] = " & psSupplement & " "
    Else
        sSQL = sSQL & "AND IB.[IBID] = " & psIBID & " "
    End If
    
    RS.CursorLocation = adUseClient
    RS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set RS.ActiveConnection = Nothing
    
    If Not RS.EOF Then
        sIBID = RS.Fields("IBID").Value
        sSupplement = RS.Fields("Supplement").Value
    End If
    
    
    psIBID = sIBID
    psSupplement = sSupplement
    GetIBParams = True
    
    'Cleanup
    Set oConn = Nothing
    Set RS = Nothing
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function GetIBParams"
End Function

Public Sub GetFileSettings(sFilePath, pFI As V2ECKeyBoard.FILE_INFORMATION, pFA As V2ECKeyBoard.FILE_ATTRIBUTES)
    On Error GoTo EH
    Dim oFI As V2ECKeyBoard.clsFileVersion
    
    Set oFI = New V2ECKeyBoard.clsFileVersion
    pFI = oFI.GetFileInformation(sFilePath)
    pFA = pFI.faFileAttributes
    Set oFI = Nothing
    
    'CleanUp
    Set oFI = Nothing
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub GetFileSettings"
End Sub

Private Sub txtToolTip_DblClick()
    On Error GoTo EH
    Dim lRet As Long
    Dim sData As String
    Dim sTickCount As String
    Dim sFileName As String
    
    
    sTickCount = goUtil.utGetTickCount
    sData = txtToolTip.Text
    sFileName = goUtil.AttachReposPath & "\Comments_" & sTickCount & ".txt"
    
    goUtil.utSaveFileData sFileName, sData
    
    lRet = goUtil.utShellExecute(GetDesktopWindow, "OPEN", sFileName, vbNullString, App.Path, vbNormalFocus, False, False, True)
    
    Sleep 1000
    
    goUtil.utDeleteFile sFileName
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub txtToolTip_DblClick"
End Sub

Private Sub txtToolTip_GotFocus()
    goUtil.utSelText txtToolTip
End Sub

