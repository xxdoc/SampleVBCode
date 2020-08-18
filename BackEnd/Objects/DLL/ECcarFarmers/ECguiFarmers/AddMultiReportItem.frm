VERSION 5.00
Begin VB.Form AddMultiReportItem 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Multi Report Item"
   ClientHeight    =   1290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7410
   Icon            =   "AddMultiReportItem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   7410
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer_SaveMe 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4800
      Top             =   120
   End
   Begin VB.Frame framMultiReport 
      Caption         =   "Report Name && Description"
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.TextBox txtDescription 
         Height          =   360
         Left            =   720
         MaxLength       =   20
         TabIndex        =   6
         Tag             =   "FileOrFolderName"
         Top             =   720
         Width           =   2640
      End
      Begin VB.TextBox txtNumber 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   360
         Left            =   4080
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   4
         Top             =   240
         Width           =   480
      End
      Begin VB.TextBox txtName 
         Enabled         =   0   'False
         Height          =   360
         Left            =   720
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   2
         Tag             =   "FileOrFolderName"
         Top             =   240
         Width           =   2640
      End
      Begin VB.Label Label1 
         Caption         =   "Max 20 chars"
         Height          =   255
         Index           =   3
         Left            =   3480
         TabIndex        =   7
         Top             =   825
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Desc:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   780
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "#"
         Height          =   255
         Index           =   0
         Left            =   3480
         TabIndex        =   3
         Top             =   300
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Name:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Width           =   495
      End
   End
   Begin VB.Frame framCommands 
      Height          =   1215
      Left            =   4980
      TabIndex        =   8
      Top             =   0
      Width           =   2295
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   855
         Left            =   1200
         MaskColor       =   &H00000000&
         Picture         =   "AddMultiReportItem.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Exit"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Sa&ve"
         Default         =   -1  'True
         Enabled         =   0   'False
         Height          =   855
         Left            =   120
         MaskColor       =   &H00000000&
         Picture         =   "AddMultiReportItem.frx":0316
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Exit"
         Top             =   240
         Width           =   975
      End
   End
End
Attribute VB_Name = "AddMultiReportItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmClaim As frmClaim

Private msTableName As String
Private msAssignmentsID As String
Private mbLoading As Boolean

Public Property Let TableName(psName As String)
    msTableName = psName
End Property
Public Property Get TableName() As String
    TableName = msTableName
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

Public Property Let Loading(pbFlag As Boolean)
    mbLoading = pbFlag
End Property
Public Property Get Loading() As Boolean
    Loading = mbLoading
End Property

Private Property Get msClassName() As String
    msClassName = Me.Name
End Property

Private Sub cmdExit_Click()
    On Error GoTo EH

    Me.Visible = False

    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdExit_Click"
End Sub

Private Sub cmdSave_Click()
    On Error GoTo EH

    SaveMe
    
    Me.Visible = False

    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdSave_Click"
End Sub

Private Sub Form_Load()
    On Error GoTo EH
    'Set the Form Icon
    Me.Icon = mfrmClaim.optClaim(GuiClaimOptions.opt06_Reports).Picture
    
    PopulateMultiRptItems
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_Load"
End Sub

Public Function CLEANUP() As Boolean
    On Error GoTo EH

    Set mfrmClaim = Nothing

    CLEANUP = True
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function CLEANUP"
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo EH

     Select Case UnloadMode
        Case vbFormControlMenu
            Me.Visible = False
            Cancel = True
    End Select

    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_QueryUnload"
End Sub

Public Function SaveMe() As Boolean
    On Error GoTo EH
    Dim udtAttach As GuiAttachItem
    Dim oListView As ListView
    Dim itmX As ListItem
    Dim sName As String
    Dim sDesc As String
    Dim lNumber As Long
    
    
    sName = txtName.Text
    sDesc = txtDescription.Text
    If IsNumeric(txtNumber.Text) Then
        lNumber = CLng(txtNumber.Text)
    End If

   ' Validate some stuff first
    goUtil.utValidate Me

    
    'ADD this entry
    If Not mfrmClaim Is Nothing Then
        mfrmClaim.AddMultiReport msTableName, sName, sDesc, lNumber
    Else
        SaveMe = False
    End If

    SaveMe = True

    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function SaveMe"
End Function

Private Sub Timer_SaveMe_Timer()
    On Error GoTo EH
    Timer_SaveMe.Enabled = False
    If cmdSave.Enabled Then
        cmdSave_Click
    End If
    Exit Sub
EH:
    'do nothing
End Sub

Private Sub txtDescription_Change()
    On Error GoTo EH
    cmdSave.Enabled = True
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub txtDescription_Change"
End Sub

Private Sub txtDescription_GotFocus()
    goUtil.utSelText txtDescription
End Sub

Private Sub txtDescription_LostFocus()
    goUtil.utValidate , txtDescription
End Sub

Private Sub txtName_GotFocus()
    goUtil.utSelText txtName
End Sub

Private Sub txtName_LostFocus()
    goUtil.utValidate , txtName
End Sub

Private Sub txtNumber_GotFocus()
    goUtil.utSelText txtNumber
End Sub

Private Sub txtNumber_LostFocus()
    goUtil.utValidate , txtNumber
End Sub

Private Sub PopulateMultiRptItems()
    On Error GoTo EH
    Dim RS As ADODB.Recordset
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    Dim sID As String
    Dim sName As String
    Dim sDescription As String
    Dim sNumber As String
    Dim sClaimNum As String
    
    Set oConn = New ADODB.Connection
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    'If adding new then
    sID = goUtil.GetAccessDBUID("ID", msTableName)
    
    'Need to get the Max Sort
    sSQL = "SELECT   MAX([Number]) + 1 As [Number] "
    sSQL = sSQL & "FROM     " & msTableName & " "
    sSQL = sSQL & "WHERE    [IDAssignments] = " & msAssignmentsID & " "
    
    Set RS = New ADODB.Recordset
    'Use Disconnected Record Set on asUseClient Cusor ONLY !
    RS.CursorLocation = adUseClient
    RS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set RS.ActiveConnection = Nothing
    
    If RS.RecordCount = 1 Then
        sNumber = goUtil.IsNullIsVbNullString(RS.Fields("Number"))
        If sNumber = vbNullString Or sNumber = 0 Then
            sNumber = "1"
        End If
    Else
        sNumber = "1"
    End If
    
    If StrComp(msTableName, "RTPhotoReport", vbTextCompare) = 0 Then
        sName = "PhotoReport" & Format(sNumber, "000")
        sDescription = vbNullString
    ElseIf StrComp(msTableName, "RTWSDiagram", vbTextCompare) = 0 Then
        sName = "WSDiagram" & Format(sNumber, "000")
        sDescription = vbNullString
    ElseIf StrComp(msTableName, "Package", vbTextCompare) = 0 Then
        sClaimNum = mfrmClaim.MyClaimsList.GetClaimItemAsString(GuiAssignments.CLIENTNUM)
        sName = "Package" & Format(sNumber, "000")
        If sNumber = "1" Then
            sDescription = "Claim File"
        Else
            sDescription = vbNullString
        End If
    End If
    
    txtName.Text = sName
    txtNumber.Text = sNumber
    txtDescription.Text = sDescription
    'cleanup
    Set RS = Nothing
    Set oConn = Nothing
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub PopulateMultiRptItems"
End Sub
