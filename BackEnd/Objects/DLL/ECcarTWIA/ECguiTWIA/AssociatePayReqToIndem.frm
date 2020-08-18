VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form AssociatePayReqToIndem 
   Caption         =   "Associate Payment Request to selected Indemnity Items:"
   ClientHeight    =   6555
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11820
   Icon            =   "AssociatePayReqToIndem.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   11820
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer_Exit 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   7800
      Top             =   5520
   End
   Begin VB.Timer Timer_GetAmount 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   7200
      Top             =   5520
   End
   Begin VB.Frame framCommands 
      Height          =   1215
      Left            =   10515
      TabIndex        =   5
      Top             =   5280
      Width           =   1215
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   855
         Left            =   120
         MaskColor       =   &H00000000&
         Picture         =   "AssociatePayReqToIndem.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Exit"
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame framAssocRTChecksID 
      Caption         =   "Associate Payment Request to selected Indemnity Items:"
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   5280
      Width           =   6735
      Begin VB.CommandButton cmdAssocRTChecksID 
         Caption         =   "Associate &Pay Request"
         Enabled         =   0   'False
         Height          =   735
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   330
         Width           =   1335
      End
      Begin VB.ComboBox cboAssocRTChecksID 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   360
         Width           =   5055
      End
   End
   Begin VB.Frame framIndemnity 
      Height          =   5175
      Left            =   80
      TabIndex        =   0
      Top             =   0
      Width           =   11655
      Begin MSComctlLib.ImageList imgIndemStatus 
         Left            =   10440
         Top             =   360
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483633
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   16777215
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "AssociatePayReqToIndem.frx":0316
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "AssociatePayReqToIndem.frx":0470
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "AssociatePayReqToIndem.frx":085C
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lvwIndemnity 
         Height          =   4815
         Left            =   120
         TabIndex        =   1
         Tag             =   "Enable"
         ToolTipText     =   "Right Click for Menu"
         Top             =   240
         Width           =   11400
         _ExtentX        =   20108
         _ExtentY        =   8493
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         OLEDragMode     =   1
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "imgIndemStatus"
         ColHdrIcons     =   "imgIndemStatus"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
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
         OLEDragMode     =   1
         NumItems        =   0
      End
   End
End
Attribute VB_Name = "AssociatePayReqToIndem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mfrmClaim As frmClaim
Private mfrmIndemnity As frmIndemnity
Private moGUI As V2ECKeyBoard.clsCarGUI
Private msAssignmentsID As String
Private mbUnloadMe As Boolean
Private mbLoading As Boolean
Private mbLoadingMe As Boolean
Private msPayReqID As String
Private msSelClassOfLossID As String
Private mlSelRTChecksID As Long
Private mbAssociateFromIndem As Boolean
Private mbEditGetAmount As Boolean

Public Property Let EditGetAmount(pbFlag As Boolean)
    mbEditGetAmount = pbFlag
End Property
Public Property Get EditGetAmount() As Boolean
    EditGetAmount = mbEditGetAmount
End Property

Public Property Let AssociateFromIndem(pbFlag As Boolean)
    mbAssociateFromIndem = pbFlag
End Property
Public Property Get AssociateFromIndem() As Boolean
    AssociateFromIndem = mbAssociateFromIndem
End Property

Public Property Let SelClassOfLossID(psClassOfLossID As String)
    msSelClassOfLossID = psClassOfLossID
End Property
Public Property Get SelClassOfLossID() As String
    SelClassOfLossID = msSelClassOfLossID
End Property

Public Property Let SelRTChecksID(plRTChecksID As Long)
    mlSelRTChecksID = plRTChecksID
End Property
Public Property Get SelRTChecksID() As Long
    SelRTChecksID = mlSelRTChecksID
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

Public Property Let MyIndemnity(poForm As Object)
    Set mfrmIndemnity = poForm
End Property
Public Property Set MyIndemnity(poForm As Object)
    Set mfrmIndemnity = poForm
End Property
Public Property Get MyIndemnity() As Object
    Set MyIndemnity = mfrmIndemnity
End Property

Public Property Let PayReqID(psID As String)
    msPayReqID = psID
End Property

Public Property Let Loading(pbFlag As Boolean)
    mbLoading = pbFlag
End Property
Public Property Get Loading() As Boolean
    Loading = mbLoading
End Property

Public Property Let UnloadMe(pbFlag As Boolean)
    mbUnloadMe = pbFlag
End Property
Public Property Get UnloadMe() As Boolean
    UnloadMe = mbUnloadMe
End Property

Private Property Get msClassName() As String
    msClassName = Me.Name
End Property


Private Sub cboAssocRTChecksID_Click()
    On Error GoTo EH
    
    If cboAssocRTChecksID.ListIndex > -1 Then
        cmdAssocRTChecksID.Enabled = True
    Else
        cmdAssocRTChecksID.Enabled = False
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cboAssocRTChecksID_Click"
End Sub

Private Sub cmdAssocRTChecksID_Click()
    On Error GoTo EH
    Dim sMess As String
    Dim sPayReqDesc As String
    Dim sIndemDesc As String
    Dim sTitle As String
    Dim bItemSelected As Boolean
    Dim itmX As MSComctlLib.ListItem
    Dim sIndemID As String
    Dim vIndemID As Variant
    Dim colIndemID As Collection
    Dim sFlagText As String
    
    
    If lvwIndemnity.ListItems.Count > 0 Then
        
        If cboAssocRTChecksID.ListIndex = -1 Then
            sMess = "You must select a Payment Request from the Drop down List!"
            MsgBox sMess, vbExclamation + vbOKOnly, "Select a Payment Request."
            cmdAssocRTChecksID.Enabled = False
            Exit Sub
        End If
        
        For Each itmX In lvwIndemnity.ListItems
            If itmX.Selected Then
                bItemSelected = True
                Exit For
            End If
        Next
        
        'See if there is a selected item
        If Not bItemSelected Then
            sMess = "You must select at least one Indemnity item from the View!"
            MsgBox sMess, vbExclamation + vbOKOnly, "Select at least one Indemnity item."
            Exit Sub
        End If
    
        
        sPayReqDesc = cboAssocRTChecksID.Text
        
        lvwIndemnity.Visible = False
        sMess = vbNullString
        Set colIndemID = New Collection
        For Each itmX In lvwIndemnity.ListItems
            If itmX.Selected Then
                'Do not allow association to indemnities marked as Previous Payment
                sFlagText = itmX.SubItems(GuiIndemListView.IsPreviousPayment - 1)
                If Not goUtil.GetFlagFromText(sFlagText) Then
                    colIndemID.Add itmX.SubItems(GuiIndemListView.ID - 1), itmX.SubItems(GuiIndemListView.ID - 1)
                Else
                    sIndemDesc = itmX.SubItems(GuiIndemListView.ClassOfLoss - 1)
                    sIndemDesc = sIndemDesc & " (" & itmX.SubItems(GuiIndemListView.ClassOfLossCode - 1) & ") - "
                    sIndemDesc = sIndemDesc & itmX.SubItems(GuiIndemListView.TypeOfLoss - 1)
                    sIndemDesc = sIndemDesc & " (" & itmX.SubItems(GuiIndemListView.TypeOfLossCode - 1) & ") - "
                    sIndemDesc = sIndemDesc & " [" & itmX.SubItems(GuiIndemListView.ReplacementCost - 1) & "]"
                    sMess = sMess & "Could not update " & sPayReqDesc & " for... " & vbCrLf
                    sMess = sMess & sIndemDesc & vbCrLf
                    sMess = sMess & "This item is marked as Previous Payment."
                End If
            End If
        Next
        
        If sMess <> vbNullString Then
            sTitle = "Could not update some items"
            MsgBox sMess, vbInformation + vbOKOnly, sTitle
        End If
        
        For Each vIndemID In colIndemID
            sIndemID = vIndemID
            If Not mfrmIndemnity.AssocIndemItemToRTChecksID(sIndemID, cboAssocRTChecksID) Then
                Exit Sub
            End If
        Next
    End If
    
    RefreshIndemnity
    mfrmIndemnity.RefreshIndemnity
    
    lvwIndemnity.Visible = True
    cmdAssocRTChecksID.Enabled = True
    
    Set itmX = Nothing
    Set colIndemID = Nothing
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdAssocRTChecksID_Click"
End Sub

Private Sub cmdExit_Click()
    On Error GoTo EH
    mbUnloadMe = True

'    goUtil.utFormWinRegPos goUtil.gsMainAppEXEName, Me, True, , , , True

    Me.Visible = False
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdExit_Click"
End Sub

Private Sub Form_Load()
    On Error GoTo EH
    Dim bHideDeleted As Boolean
    
    mbLoading = True
    'Set the Form Icon
    Me.Icon = mfrmClaim.optClaim(GuiClaimOptions.opt03_Indemnity).Picture
    

    'Show the form but hide it from view
    Me.left = -50000

'    goUtil.utFormWinRegPos goUtil.gsMainAppEXEName, Me, , , , , True

    
    mfrmIndemnity.LoadHeaderlvwIndemnity lvwIndemnity
    
    LoadMe
    
    If mbAssociateFromIndem Then
        Timer_GetAmount.Enabled = True
    ElseIf mbEditGetAmount Then
        Timer_Exit.Enabled = True
    Else
        Timer_Exit.Enabled = True
    End If
    
    mbLoading = False
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_Load"
End Sub

Public Function LoadMe() As Boolean
    On Error GoTo EH
    Dim itmX As ListItem
    Dim lCount As Long
    mbLoadingMe = True
    
    RefreshIndemnity
    
    'Select default items for association
    For Each itmX In lvwIndemnity.ListItems
        If itmX.SubItems(GuiIndemListView.ClassOfLossID - 1) = msSelClassOfLossID Then
            'Only Select Already Selected Items and Items that have yet to be selected
            If itmX.SubItems(GuiIndemListView.RTChecksID - 1) = CStr(mlSelRTChecksID) Or itmX.SubItems(GuiIndemListView.RTChecksID - 1) = "0" Then
                itmX.Selected = True
            End If
        End If
    Next
    
    For lCount = 0 To cboAssocRTChecksID.ListCount - 1
        If cboAssocRTChecksID.ItemData(lCount) = mlSelRTChecksID Then
            cboAssocRTChecksID.ListIndex = lCount
            Exit For
        End If
    Next
    
    LoadMe = True
    mbLoadingMe = False
    Exit Function
EH:
    mbLoadingMe = False
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function LoadMe"
End Function

Public Sub RefreshIndemnity()
    On Error GoTo EH
    
    'populate the Totals on the Indem screen
    If Not mfrmClaim.SetadoRSRTIndemnity(msAssignmentsID) Then
        Exit Sub
    End If
    mfrmIndemnity.PopulatelvwIndemnity lvwIndemnity
    
    'Load Payment Req RS (RTChecks)
    mfrmClaim.SetadoRSPayment msAssignmentsID, True
    cboAssocRTChecksID.Clear
    cboAssocRTChecksID.AddItem "(--Disassociate Payment--)"
    '0 indicates Null ID since ID must be >= 1 or <= -1
    '>=1    : WEB Server Synched
    '<=-1   : Client has yet to synch Data ID with Web Server.
    cboAssocRTChecksID.ItemData(cboAssocRTChecksID.NewIndex) = 0
    mfrmClaim.PopulateLookUp mfrmClaim.adoRSPayment, _
                        Nothing, _
                        cboAssocRTChecksID, _
                        "ID", _
                        vbNullString, _
                        "CheckNum", _
                        "PayReqDescription"
                        
    Exit Sub
EH:
   goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub RefreshIndemnity"
End Sub

Public Sub ReSizeMe()
    On Error Resume Next
    'RePos Controls
    'Width and Lefts
    framIndemnity.Width = Me.Width - 285
    lvwIndemnity.Width = Me.Width - 540
    
    
    'framCommands
    framCommands.left = Me.Width - 1425
    
    
    'Heights and Tops
    framIndemnity.Height = Me.Height - 1785
    lvwIndemnity.Height = Me.Height - 2145
    framAssocRTChecksID.top = Me.Height - 1680
    
    'framCommands
    framCommands.top = Me.Height - 1680
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo EH
    
     Select Case UnloadMode
        Case vbFormControlMenu
            
'            goUtil.utFormWinRegPos goUtil.gsMainAppEXEName, Me, True, , , , True

            mbUnloadMe = True
            Me.Visible = False
            Cancel = True
    End Select
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_QueryUnload"
End Sub

Private Sub Form_Resize()
    If Not mbLoading Then
        ReSizeMe
    End If
End Sub

Public Function CLEANUP() As Boolean
    On Error GoTo EH

    Set mfrmClaim = Nothing
    Set mfrmIndemnity = Nothing
    Set moGUI = Nothing
    
    CLEANUP = True
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function CLEANUP"
End Function

Private Sub lvwIndemnity_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    On Error GoTo EH
    Dim sMess As String
    If lvwIndemnity.SortOrder = lvwAscending Then
        lvwIndemnity.SortOrder = lvwDescending
        sMess = "Descending"
    Else
        lvwIndemnity.SortOrder = lvwAscending
        sMess = "Ascending"
    End If
    
    'Set the Tool Tip
    lvwIndemnity.ToolTipText = "Sort By " & ColumnHeader.Text & " " & sMess
    
    'Need to see if the Column clicked was a NON text sort
    'Like Date or number.  If a non textual column was clicked
    'need to use the next column as the sort since this next column
    'will be hidden and contain a sort friendly format.
    'Sort Key is Base 0 where CoulmnHeader is not
    Select Case ColumnHeader.Index
        Case GuiIndemListView.DateLastUpdated
            lvwIndemnity.SortKey = ColumnHeader.Index
        Case GuiIndemListView.ACVClaim
            lvwIndemnity.SortKey = ColumnHeader.Index
        Case GuiIndemListView.ACVLessExcessLimits
            lvwIndemnity.SortKey = ColumnHeader.Index
        Case GuiIndemListView.ExcessLimits
            lvwIndemnity.SortKey = ColumnHeader.Index
        Case GuiIndemListView.NonRecoverableDep
            lvwIndemnity.SortKey = ColumnHeader.Index
        Case GuiIndemListView.RecoverableDep
            lvwIndemnity.SortKey = ColumnHeader.Index
        Case GuiIndemListView.ReplacementCost
            lvwIndemnity.SortKey = ColumnHeader.Index
        Case GuiIndemListView.SpecialLimits
            lvwIndemnity.SortKey = ColumnHeader.Index
        Case Else
            lvwIndemnity.SortKey = ColumnHeader.Index - 1
    End Select
    
    lvwIndemnity.Sorted = True
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lvwIndemnity_ColumnClick"
End Sub

Private Sub Timer_Exit_Timer()
    On Error GoTo EH
    
    Timer_Exit.Enabled = False
    
    'Exit
    cmdExit_Click
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Timer_Exit_Timer"
End Sub

Private Sub Timer_GetAmount_Timer()
    On Error GoTo EH
    Dim vRTIndemnityID As Variant
    Dim sRTIndemnityID As String
    Dim colSelIndemID As Collection
    Dim itmX As ListItem
    
    Timer_GetAmount.Enabled = False
    
    'Need to set ref to indemnity listview
    Set colSelIndemID = mfrmIndemnity.colSelIndemID
    
    'Deselected any already selected items here
    For Each itmX In lvwIndemnity.ListItems
        itmX.Selected = False
    Next
    
    'Select the previously selected items on the main Indemnity Screen
    For Each vRTIndemnityID In colSelIndemID
        sRTIndemnityID = vRTIndemnityID
        SelectItemX sRTIndemnityID
    Next
    
    'Associate the items to the payment request
    cmdAssocRTChecksID_Click
    
    'Exit
    cmdExit_Click
    
    'Cleanup
    
    Set itmX = Nothing
    Set colSelIndemID = Nothing
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Timer_GetAmount_Timer"
End Sub

Private Function SelectItemX(psRTIndemnityID As String) As Boolean
    On Error GoTo EH
    Dim itmX As ListItem
    Dim sRTIndemnityID As String
    Dim sThisID As String
    
    sRTIndemnityID = psRTIndemnityID
    
    For Each itmX In lvwIndemnity.ListItems
        sThisID = itmX.SubItems(GuiIndemListView.RTIndemnityID - 1)
        If StrComp(sRTIndemnityID, sThisID, vbTextCompare) = 0 Then
            itmX.Selected = True
            Exit For
        End If
    Next
    
    Set itmX = Nothing

    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Function SelectItemX"
End Function
