VERSION 5.00
Begin VB.Form frmMiscellaneous 
   AutoRedraw      =   -1  'True
   Caption         =   "Miscellaneous"
   ClientHeight    =   6555
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11820
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6555
   ScaleWidth      =   11820
   StartUpPosition =   3  'Windows Default
   Tag             =   "Miscellaneous"
   Begin VB.Frame framCommands 
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3375
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   855
         Left            =   2280
         MaskColor       =   &H00000000&
         Picture         =   "frmMiscellaneous.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
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
         Picture         =   "frmMiscellaneous.frx":0282
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Exit"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdSpelling 
         Caption         =   "Spellin&g"
         Enabled         =   0   'False
         Height          =   855
         Left            =   120
         MaskColor       =   &H00000000&
         Picture         =   "frmMiscellaneous.frx":03CC
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Exit"
         Top             =   240
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmMiscellaneous"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mfrmClaim As frmClaim
Private mbUnloadMe As Boolean
Private msAssignmentsID As String
Private moGUI As V2ECKeyBoard.clsCarGUI
Private moCurrentTextBox As TextBox
Private mbLoading As Boolean 'Loading Form
Private mbLoadingMe As Boolean 'Just loading data

Public Property Let CurrentTextBox(poTextBox As Object)
    On Error GoTo EH
    Set moCurrentTextBox = poTextBox
    If moCurrentTextBox Is Nothing Then
        cmdSpelling.Enabled = False
    Else
        cmdSpelling.Enabled = True
    End If
    Exit Property
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Property Let CurrentTextBox"
End Property
Public Property Set CurrentTextBox(poTextBox As Object)
    On Error GoTo EH
    Set moCurrentTextBox = poTextBox
    If moCurrentTextBox Is Nothing Then
        cmdSpelling.Enabled = False
    Else
        cmdSpelling.Enabled = True
    End If
    Exit Property
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Property Set CurrentTextBox"
End Property
Public Property Get CurrentTextBox() As Object
    Set CurrentTextBox = moCurrentTextBox
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

Public Property Let UnloadMe(pbFlag As Boolean)
    mbUnloadMe = pbFlag
End Property
Public Property Get UnloadMe() As Boolean
    UnloadMe = mbUnloadMe
End Property

Public Property Let MyfrmClaim(pofrmClaim As Object)
    Set mfrmClaim = pofrmClaim
End Property
Public Property Set MyfrmClaim(pofrmClaim As Object)
    Set mfrmClaim = pofrmClaim
End Property
Public Property Get MyfrmClaim() As Object
    Set MyfrmClaim = mfrmClaim
End Property

Private Property Get msClassName() As String
    msClassName = Me.Name
End Property

Private Sub cmdExit_Click()
    On Error GoTo EH
    
    mbUnloadMe = True
    Me.Visible = False
    mfrmClaim.Timer_UnloadForm.Enabled = True
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdExit_Click"
End Sub

Private Sub cmdSave_Click()
    On Error GoTo EH
    If SaveMe Then
        mfrmClaim.RefreshMe
        cmdSave.Enabled = False
    End If
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdSave_Click"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo EH
    
    Select Case KeyCode
        Case KeyCodeConstants.vbKeyF11 'maximize window / Normalize window
            If Me.WindowState = VBRUN.FormWindowStateConstants.vbMaximized Then
                Me.WindowState = VBRUN.FormWindowStateConstants.vbNormal
            Else
                Me.WindowState = VBRUN.FormWindowStateConstants.vbMaximized
            End If
            
        Case Else
            If Not mfrmClaim Is Nothing Then
                mfrmClaim.Form_KeyDown KeyCode, Shift
            End If
    End Select
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_KeyDown"
End Sub

Private Sub Form_Load()
    On Error GoTo EH
    mbLoading = True
    'Set the Form Icon
    Me.Icon = mfrmClaim.optClaim(GuiClaimOptions.opt08_Miscellaneous).Picture
'    LoadMe
    CheckStatus
    
'    ShowFrame
    cmdSave.Enabled = False
    mbLoading = False
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_Load"
End Sub

Public Function LoadMe() As Boolean
    On Error GoTo EH
    Dim oConn As New ADODB.Connection
    Dim adoRS As ADODB.Recordset
    Dim RS As ADODB.Recordset
    Dim sSQL As String
    Dim sMess As String
    Dim sTemp As String
    
    mbLoadingMe = True
    
    Set oConn = New ADODB.Connection
    Set adoRS = New ADODB.Recordset
    
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
'    sSQL = "SELECT A.*, "
'    sSQL = sSQL & "( "
'    sSQL = sSQL & "SELECT UserName "
'    sSQL = sSQL & "FROM USERS "
'    sSQL = sSQL & "WHERE UsersID =  " & goUtil.gsCurUsersID & " "
'    sSQL = sSQL & ") As AdjusterSpecUserName, "
'    sSQL = sSQL & "( "
'    sSQL = sSQL & "SELECT  ACID "
'    sSQL = sSQL & "FROM ClientCoAdjusterSpec "
'    sSQL = sSQL & "Where ClientCoAdjusterSpecID = A.[AdjusterSPecID] "
'    sSQL = sSQL & ") As AdjusterSpecACID, "
'    sSQL = sSQL & "( "
'    sSQL = sSQL & "SELECT  ACIDDescription "
'    sSQL = sSQL & "FROM ClientCoAdjusterSpec "
'    sSQL = sSQL & "Where ClientCoAdjusterSpecID = A.[AdjusterSPecID] "
'    sSQL = sSQL & ") As AdjusterSpecAcidDescription, "
'    sSQL = sSQL & "( "
'    sSQL = sSQL & "SELECT  Type "
'    sSQL = sSQL & "FROM    AssignmentType "
'    sSQL = sSQL & "WHERE   AssignmentTypeID = A.[AssignmentTypeID] "
'    sSQL = sSQL & ") As AssignmentTypeType, "
'    sSQL = sSQL & "( "
'    sSQL = sSQL & "SELECT  CatCode "
'    sSQL = sSQL & "FROM    ClientCompanyCatSpec "
'    sSQL = sSQL & "WHERE   ClientCompanyCatSpecID = A.[ClientCompanyCatSpecID] "
'    sSQL = sSQL & ") As ClientCompanyCatSpecCatCode, "
'    sSQL = sSQL & "( "
'    sSQL = sSQL & "SELECT Name "
'    sSQL = sSQL & "FROM CAT "
'    sSQL = sSQL & "WHERE CATID = " & goUtil.gsCurCat & " "
'    sSQL = sSQL & ") As CatName, "
'    sSQL = sSQL & "S.Status As Status, "
'    sSQL = sSQL & "CCCS.CatCode "
'    sSQL = sSQL & "FROM (Assignments A "
'    sSQL = sSQL & "INNER JOIN STATUS S ON A.StatusID = S.StatusID) "
'    sSQL = sSQL & "INNER JOIN CLIENTCOMPANYCATSPEC CCCS ON (A.ClientCompanyCatSpecID = CCCS.ClientCompanyCatSpecID) "
'    sSQL = sSQL & "WHERE A.ClientCompanyCatSpecID IN "
'    sSQL = sSQL & "( "
'    sSQL = sSQL & "SELECT  ClientCompanyCatSpecID "
'    sSQL = sSQL & "FROM ClientCompanyCatSpec "
'    sSQL = sSQL & "WHERE ClientCompanyID = " & goUtil.gsCurCar & " "
'    sSQL = sSQL & "AND     CATID = " & goUtil.gsCurCat & " "
'    sSQL = sSQL & ") "
'    sSQL = sSQL & "AND A.AdjusterSpecID IN "
'    sSQL = sSQL & "( "
'    sSQL = sSQL & "SELECT  ClientCoAdjusterSpecID "
'    sSQL = sSQL & "FROM ClientCoAdjusterSpec "
'    sSQL = sSQL & "Where ClientCompanyID = " & goUtil.gsCurCar & " "
'    sSQL = sSQL & "AND UsersID = " & goUtil.gsCurUsersID & " "
'    sSQL = sSQL & ") "
'    sSQL = sSQL & "AND A.ID = " & msAssignmentsID & " "
'
'    'Use Disconnected Record Set on asUseClient Cusor ONLY !
'    adoRS.CursorLocation = adUseClient
'    adoRS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
'    Set adoRS.ActiveConnection = Nothing
'
'    If adoRS.RecordCount > 1 Then
'        adoRS.MoveFirst
'        sMess = "Database Error.  Duplicate Record ID found!" & vbCrLf & vbCrLf & "ID = " & adoRS!ID & vbCrLf & vbCrLf & "AssignmentsID = " & adoRS!AssignmentsID
'        Err.Raise -999, , sMess
'    ElseIf adoRS.RecordCount = 0 Then
'        sMess = "Database Error.  Record ID Not found!" & vbCrLf & vbCrLf & "AssignmentsID = " & msAssignmentsID
'        Err.Raise -999, , sMess
'    End If
'
'    adoRS.MoveFirst
'
'    'Populate the Available Type Of Loss Info
'    If Not MyGUI.adoRSTypeOfLoss Is Nothing Then
'        PopulateLookUp MyGUI.adoRSTypeOfLoss, _
'                        adoRS, _
'                        cboTypeOfLoss, _
'                        "TypeOfLossID", _
'                        "TypeOfLossID", _
'                        "TypeOfLoss", _
'                        "Code"
'    End If
'
'    'Populate the Available Assignment type
'    If Not MyGUI.adoRSAssignmentType Is Nothing Then
'        PopulateLookUp MyGUI.adoRSAssignmentType, _
'                        adoRS, _
'                        cboAssignmentType, _
'                        "AssignmentTypeID", _
'                        "AssignmentTypeID", _
'                        "Type", _
'                        "Description"
'    End If
'
'    'Populate the Available ACID (Adjuster Client Identification)
'    If Not MyGUI.adoRSACID Is Nothing Then
'        PopulateLookUp MyGUI.adoRSACID, _
'                        adoRS, _
'                        cboACID, _
'                        "ClientCoAdjusterSpecID", _
'                        "AdjusterSpecID", _
'                        "ACID", _
'                        "ACIDDescription"
'    End If
'
'    'Populate the Available ACID Display (Adjuster Client Identification)
'    'This will show on billing information
'    If Not MyGUI.adoRSACID Is Nothing Then
'        PopulateLookUp MyGUI.adoRSACID, _
'                        adoRS, _
'                        cboACIDDisplay, _
'                        "ClientCoAdjusterSpecID", _
'                        "AdjusterSpecIDDisplay", _
'                        "ACID", _
'                        "ACIDDescription"
'    End If
'
'    'Populate the Available Cat Code
'    If Not MyGUI.adoRSCatCode Is Nothing Then
'        PopulateLookUp MyGUI.adoRSCatCode, _
'                        adoRS, _
'                        cboCatCode, _
'                        "ClientCompanyCatSpecID", _
'                        "ClientCompanyCatSpecID", _
'                        "CatCode", _
'                        "Comments"
'    End If
'
'    'Specifics
'    txtIBNUM.Text = goUtil.IsNullIsVbNullString(adoRS.Fields("IBNUM"))
'    txtCLIENTNUM.Text = goUtil.IsNullIsVbNullString(adoRS.Fields("CLIENTNUM"))
'    txtPolicyNo.Text = goUtil.IsNullIsVbNullString(adoRS.Fields("PolicyNo"))
'    txtPolicyDescription.Text = goUtil.IsNullIsVbNullString(adoRS.Fields("PolicyDescription"))
'    txtMortgageeName.Text = goUtil.IsNullIsVbNullString(adoRS.Fields("MortgageeName"))
'    txtAgentNo.Text = goUtil.IsNullIsVbNullString(adoRS.Fields("AgentNo"))
'    txtReportedBy.Text = goUtil.IsNullIsVbNullString(adoRS.Fields("ReportedBy"))
'    txtReportedByPhone.Text = goUtil.IsNullIsVbNullString(adoRS.Fields("ReportedByPhone"))
'
'    'Dates
'    txtLossDate.Text = goUtil.IsNullIsVbNullString(adoRS.Fields("LossDate"))
'    txtLossDate.Text = Format(txtLossDate.Text, "MM/DD/YYYY")
'    txtAssignedDate.Text = goUtil.IsNullIsVbNullString(adoRS.Fields("AssignedDate"))
'    txtAssignedDate.Text = Format(txtAssignedDate.Text, "MM/DD/YYYY")
'    txtReceivedDate.Text = goUtil.IsNullIsVbNullString(adoRS.Fields("ReceivedDate"))
'    txtReceivedDate.Text = Format(txtReceivedDate.Text, "MM/DD/YYYY")
'    txtContactDate.Text = goUtil.IsNullIsVbNullString(adoRS.Fields("ContactDate"))
'    txtContactDate.Text = Format(txtContactDate.Text, "MM/DD/YYYY")
'    txtInspectedDate.Text = goUtil.IsNullIsVbNullString(adoRS.Fields("InspectedDate"))
'    txtInspectedDate.Text = Format(txtInspectedDate.Text, "MM/DD/YYYY")
'    txtCloseDate.Text = goUtil.IsNullIsVbNullString(adoRS.Fields("CloseDate"))
'    txtCloseDate.Text = Format(txtCloseDate.Text, "MM/DD/YYYY")
'
'    'Insured Info
'    txtInsured.Text = goUtil.IsNullIsVbNullString(adoRS.Fields("Insured"))
'    txtHomePhone.Text = goUtil.IsNullIsVbNullString(adoRS.Fields("HomePhone"))
'    txtBusinessPhone.Text = goUtil.IsNullIsVbNullString(adoRS.Fields("BusinessPhone"))
'    'Property Address
'    txtPAStreet.Text = goUtil.IsNullIsVbNullString(adoRS.Fields("PAStreet"))
'    txtPACity.Text = goUtil.IsNullIsVbNullString(adoRS.Fields("PACity"))
'    'Populate the Available Sates for Property address
'    If Not MyGUI.adoRSState Is Nothing Then
'        PopulateLookUp MyGUI.adoRSState, _
'                        adoRS, _
'                        cboPAState, _
'                        "StateID", _
'                        "", _
'                        "Code", _
'                        "Name", _
'                        True, _
'                        "PAState"
'    End If
'    txtPAZIP.Text = Format(goUtil.IsNullIsVbNullString(adoRS.Fields("PAZIP")), "00000")
'    txtPAZIP4.Text = Format(goUtil.IsNullIsVbNullString(adoRS.Fields("PAZIP4")), "0000")
'    txtPAOtherPostCode.Text = goUtil.IsNullIsVbNullString(adoRS.Fields("PAOtherPostCode"))
'
'    'Mailing Address
'    txtMAStreet.Text = goUtil.IsNullIsVbNullString(adoRS.Fields("MAStreet"))
'    txtMACity.Text = goUtil.IsNullIsVbNullString(adoRS.Fields("MACity"))
'    'Populate the Available Sates for Mailling address
'    If Not MyGUI.adoRSState Is Nothing Then
'        PopulateLookUp MyGUI.adoRSState, _
'                        adoRS, _
'                        cboMAState, _
'                        "StateID", _
'                        "", _
'                        "Code", _
'                        "Name", _
'                        True, _
'                        "MAState"
'    End If
'    txtMAZIP.Text = Format(goUtil.IsNullIsVbNullString(adoRS.Fields("MAZIP")), "00000")
'    txtMAZIP4.Text = Format(goUtil.IsNullIsVbNullString(adoRS.Fields("MAZIP4")), "0000")
'    txtMAOtherPostCode.Text = goUtil.IsNullIsVbNullString(adoRS.Fields("MAOtherPostCode"))
'
'    'Policy limits
'    txtDeductible.Text = goUtil.IsNullIsVbNullString(adoRS.Fields("Deductible"))
'    If mbLoading Then
'        LoadHeaderPLClassTypeID
'    End If
'    LoadPolicyLimitsStuff
'
'    'Loss Report
'    LoadLossReportStuff
'    sTemp = goUtil.IsNullIsVbNullString(adoRS.Fields("LRFormat"))
'    If sTemp <> vbNullString Then
'        If StrComp(sTemp, "TEXT", vbTextCompare) = 0 Then
'            txtLossReport.Text = goUtil.IsNullIsVbNullString(adoRS.Fields("LossReport"))
'        Else
'            txtLossReport.Text = vbNullString
'        End If
'    Else
'        txtLossReport.Text = vbNullString
'    End If
'
'    'Admin Comments
'    txtAdminComments.Text = goUtil.IsNullIsVbNullString(adoRS.Fields("AdminComments"))
    
    'cleanup
    Set RS = Nothing
    Set adoRS = Nothing
    Set oConn = Nothing
    
    LoadMe = True
    mbLoadingMe = False
    Exit Function
EH:
    mbLoadingMe = False
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function LoadMe"
    Set adoRS = Nothing
    Set oConn = Nothing
End Function

Public Function SaveMe() As Boolean
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim MyadoRSAssignments As ADODB.Recordset
    Dim iCurrentStatus As V2ECKeyBoard.AssgnStatus
    Dim sSQL As String
    'If Close date is Set then be sure all the other dates are set tooooo
    Dim bCloseDateIsSet As Boolean
    ' Vars
    
    'validate all the fields on this form
    goUtil.utValidate Me
    
'    'Check for Drop Down Items Not Selected that should be
'    If cboAssignmentType.ListIndex = -1 Then
'        sMess = sMess & "Assignment Type not selected !" & vbCrLf
'    End If
'    If cboCatCode.ListIndex = -1 Then
'        sMess = sMess & "Cat Code not selected !" & vbCrLf
'    End If
'    If cboACID.ListIndex = -1 Then
'        sMess = sMess & "ACID not selected !" & vbCrLf
'    End If
'    If cboACIDDisplay.ListIndex = -1 Then
'        sMess = sMess & "ACID Display not selected !" & vbCrLf
'    End If
'    If cboMAState.ListIndex = -1 Then
'        sMess = sMess & "Mailing State not selected !" & vbCrLf
'    End If
'    If cboPAState.ListIndex = -1 Then
'        sMess = sMess & "Property State not selected !" & vbCrLf
'    End If
'    If cboTypeOfLoss.ListIndex = -1 Then
'        sMess = sMess & "Type Of Loss not selected !" & vbCrLf
'    End If
'
'    'DATES !!!!
'    'Close Date
'    If IsDate(txtCloseDate.Text) Then
'        'Check for Close date but no other dates filled out
'        bCloseDateIsSet = True
'        sCloseDate = "#" & Format(txtCloseDate.Text, "MM/DD/YYYY") & "#"
'    Else
'        sCloseDate = "null"
'    End If
'
'    'Loss Date
'    If IsDate(txtLossDate.Text) Then
'        sLossDate = "#" & Format(txtLossDate.Text, "MM/DD/YYYY") & "#"
'    Else
'        'Check for Close date but no other dates filled out
'        If bCloseDateIsSet Then
'            sMess = sMess & "Loss Date is not set!" & vbCrLf
'        End If
'        sLossDate = "null"
'    End If
'    'Assigned Date
'    If IsDate(txtAssignedDate.Text) Then
'        sAssignedDate = "#" & Format(txtAssignedDate.Text, "MM/DD/YYYY") & "#"
'    Else
'        If bCloseDateIsSet Then
'            sMess = sMess & "Assigned Date is not set!" & vbCrLf
'        End If
'        sAssignedDate = "null"
'    End If
'    'Received Date
'    If IsDate(txtReceivedDate.Text) Then
'        sReceivedDate = "#" & Format(txtReceivedDate.Text, "MM/DD/YYYY") & "#"
'    Else
'        If bCloseDateIsSet Then
'            sMess = sMess & "Received Date is not set!" & vbCrLf
'        End If
'        sReceivedDate = "null"
'    End If
'    'Contact Date
'    If IsDate(txtContactDate.Text) Then
'        sContactDate = "#" & Format(txtContactDate.Text, "MM/DD/YYYY") & "#"
'    Else
'        If bCloseDateIsSet Then
'            sMess = sMess & "Contact Date is not set!" & vbCrLf
'        End If
'        sContactDate = "null"
'    End If
'    'Inspected Date
'    If IsDate(txtInspectedDate.Text) Then
'        sInspectedDate = "#" & Format(txtInspectedDate.Text, "MM/DD/YYYY") & "#"
'    Else
'        If bCloseDateIsSet Then
'            sMess = sMess & "Inspected Date is not set!" & vbCrLf
'        End If
'        sInspectedDate = "null"
'    End If
'
'    If sMess <> vbNullString Then
'        sMess = "Could not save " & Me.Caption & vbCrLf & vbCrLf & sMess
'        MsgBox sMess, vbExclamation + vbOKOnly, "Could Not Save Claim Information."
'        Exit Function
'    End If
'
'    'Use this to check new values to be inserted
'    'against the current values in this recordset
'    Set MyadoRSAssignments = mfrmClaim.adoRSAssignments
'
'    'set the Assignemtn vars
'    sAssignmentsID = msAssignmentsID
'
'    sID = msAssignmentsID
'
'    sAssignmentTypeID = cboAssignmentType.ItemData(cboAssignmentType.ListIndex)
'
'    sClientCompanyCatSpecID = cboCatCode.ItemData(cboCatCode.ListIndex)
'
'    sAdjusterSpecID = cboACID.ItemData(cboACID.ListIndex)
'
'    sAdjusterSpecIDDisplay = cboACIDDisplay.ItemData(cboACIDDisplay.ListIndex)
'
'    sSPVersion = "[SPVersion]"
'    'IBNUM
'    sIBNUM = "'" & goUtil.utCleanSQLString(UCase(txtIBNUM.Text)) & "'"
'    'CLIENTNUM
'    sCLIENTNUM = "'" & goUtil.utCleanSQLString(UCase(txtCLIENTNUM.Text)) & "'"
'    'Policy Number
'    sPolicyNo = "'" & goUtil.utCleanSQLString(UCase(txtPolicyNo.Text)) & "'"
'    'Policty Description
'    sPolicyDescription = "'" & goUtil.utCleanSQLString(UCase(txtPolicyDescription.Text)) & "'"
'    'Insured
'    sInsured = "'" & goUtil.utCleanSQLString(UCase(txtInsured.Text)) & "'"
'
'    'Mailing Address
'    'Street
'    sMAStreet = UCase(txtMAStreet.Text)
'    'City
'    sMACity = UCase(txtMACity.Text)
'    'State
'    sMAState = left(UCase(cboMAState.Text), 2)
'    'Zip
'    sMAZIP = txtMAZIP.Text
'    'Zip4
'    sMAZIP4 = txtMAZIP4.Text
'    'Other Post Code
'    sMAOtherPostCode = UCase(txtMAOtherPostCode.Text)
'    'Build entire Address
'    If sMAZIP = "00000" & sMAZIP4 = "0000" Then
'        sMailingAddress = "'" & goUtil.utCleanSQLString(goUtil.utJoinToAddress(sMAStreet, sMACity, sMAState, sMAOtherPostCode)) & "'"
'    Else
'        sMailingAddress = "'" & goUtil.utCleanSQLString(goUtil.utJoinToAddress(sMAStreet, sMACity, sMAState, Format(sMAZIP, "00000") & "-" & Format(sMAZIP4, "0000"))) & "'"
'    End If
'    'Street
'    sMAStreet = "'" & goUtil.utCleanSQLString(UCase(txtMAStreet.Text)) & "'"
'    'City
'    sMACity = "'" & goUtil.utCleanSQLString(UCase(txtMACity.Text)) & "'"
'    'State
'    sMAState = "'" & goUtil.utCleanSQLString(left(UCase(cboMAState.Text), 2)) & "'"
'    'Zip
'    sMAZIP = txtMAZIP.Text
'    'Zip4
'    sMAZIP4 = txtMAZIP4.Text
'    'Other Post Code
'    sMAOtherPostCode = "'" & goUtil.utCleanSQLString(UCase(txtMAOtherPostCode.Text)) & "'"
'    'End Mailing Address
'
'    'Property Address
'    'Street
'    sPAStreet = UCase(txtPAStreet.Text)
'    'City
'    sPACity = UCase(txtPACity.Text)
'    'State
'    sPAState = left(UCase(cboPAState.Text), 2)
'    'Zip
'    sPAZIP = txtPAZIP.Text
'    'Zip4
'    sPAZIP4 = txtPAZIP4.Text
'    'other PostCode
'    sPAOtherPostCode = UCase(txtPAOtherPostCode.Text)
'    'Build entire Address
'    If sPAZIP = "00000" & sPAZIP4 = "0000" Then
'        sPropertyAddress = "'" & goUtil.utCleanSQLString(goUtil.utJoinToAddress(sPAStreet, sPACity, sPAState, sPAOtherPostCode)) & "'"
'    Else
'        sPropertyAddress = "'" & goUtil.utCleanSQLString(goUtil.utJoinToAddress(sPAStreet, sPACity, sPAState, Format(sPAZIP, "00000") & "-" & Format(sPAZIP4, "0000"))) & "'"
'    End If
'    'Street
'    sPAStreet = "'" & goUtil.utCleanSQLString(UCase(txtPAStreet.Text)) & "'"
'    'City
'    sPACity = "'" & goUtil.utCleanSQLString(UCase(txtPACity.Text)) & "'"
'    'State
'    sPAState = "'" & goUtil.utCleanSQLString(left(UCase(cboPAState.Text), 2)) & "'"
'    'Zip
'    sPAZIP = txtPAZIP.Text
'    'Zip4
'    sPAZIP4 = txtPAZIP4.Text
'    'Other Post Code
'    sPAOtherPostCode = "'" & goUtil.utCleanSQLString(UCase(txtPAOtherPostCode.Text)) & "'"
'    'End Property Address
'
'    'Home Phone
'    sHomePhone = "'" & goUtil.utCleanSQLString(UCase(txtHomePhone.Text)) & "'"
'    'Business Phone
'    sBusinessPhone = "'" & goUtil.utCleanSQLString(UCase(txtBusinessPhone.Text)) & "'"
'    'Mortgage name
'    sMortgageeName = "'" & goUtil.utCleanSQLString(UCase(txtMortgageeName.Text)) & "'"
'    'Agent No
'    sAgentNo = "'" & goUtil.utCleanSQLString(UCase(txtAgentNo.Text)) & "'"
'    'Reported By
'    sReportedBy = "'" & goUtil.utCleanSQLString(UCase(txtReportedBy.Text)) & "'"
'    'Reported by Phone
'    sReportedByPhone = "'" & goUtil.utCleanSQLString(UCase(txtReportedByPhone.Text)) & "'"
'    'Deductible
'    sDeductible = txtDeductible.Text
'
'    sAppDedClassTypeIDOrder = "[AppDedClassTypeIDOrder]"
'
'    'if the Loss report was changed to TEXT then need to update these vars
'    'otherwise they remain the same!
'    '(Attaching a PDF Loss Report already updates Assignments table See --> Private Sub cmdAttachPDFLossReport_Click)
'    If StrComp(cboAssignmentLossReportFormat.Text, "TEXT", vbTextCompare) = 0 Then
'        sLRFormat = "'TEXT'"
'        sLossReport = "'" & goUtil.utCleanSQLString(txtLossReport.Text) & "'"
'
'        sDownLoadLossReport = "[DownLoadLossReport]"
'
'        sUpLoadLossReport = "True"
'    Else
'        sLRFormat = "[LRFormat]"
'
'        sLossReport = "[LossReport]"
'
'        sDownLoadLossReport = "[DownLoadLossReport]"
'
'        sUpLoadLossReport = "[UpLoadLossReport]"
'    End If
'    'Type Of Loss
'    sTypeOfLossID = cboTypeOfLoss.ItemData(cboTypeOfLoss.ListIndex)
'
'    sXactTypeOfLoss = "[XactTypeOfLoss]"
'
'    sSentToXact = "[SentToXact]"
'
'    sReassigned = "[Reassigned]"
'
'    sDateReassigned = "[DateReassigned]"
'
'
'    'STATUS ID !
'    'Check for Closed Date
'    'Change Status ID
'    If IsDate(txtCloseDate.Text) Then
'        sStatusID = CStr(V2ECKeyBoard.AssgnStatus.iAssignmentsStatus_CLOSED)
'    Else
'        'Check to see if the Current Status is Closed if it Is Need to
'        'Change the Status to NEW
'        iCurrentStatus = MyadoRSAssignments.Fields("StatusID").Value
'        Select Case iCurrentStatus
'            Case V2ECKeyBoard.AssgnStatus.iAssignmentsStatus_CLOSED
'                sStatusID = V2ECKeyBoard.AssgnStatus.iAssignmentsStatus_NEW
'            Case Else
'                sStatusID = "[StatusID]"
'        End Select
'
'    End If
'
'    sRAAdjusterSpecID = "[RAAdjusterSpecID]"
'
'    sIsLocked = "[IsLocked]"
'
'    sIsDeleted = "[IsDeleted]"
'
'    sDownLoadMe = "[DownLoadMe]"
'
'    sUpLoadMe = "True"
'
'    sDownLoadAll = "[DownLoadAll]"
'
'    sUpLoadAll = "[UpLoadAll]"
'    'Admin Comments
'    sAdminComments = "'" & goUtil.utCleanSQLString(txtAdminComments.Text) & "'"
'
'    sMiscDelimSettings = "[MiscDelimSettings]"
'
'    sDateLastUpdated = "#" & Format(Now(), "MM/DD/YYYY") & "#"
'
'    sUID = goUtil.gsCurUsersID
'
'
'    sSQL = "Update Assignments Set "
'    sSQL = sSQL & "[AssignmentsID] = " & sAssignmentsID & ", "
'    sSQL = sSQL & "[ID] = " & sID & ", "
'    sSQL = sSQL & "[AssignmentTypeID] = " & sAssignmentTypeID & ", "
'    sSQL = sSQL & "[ClientCompanyCatSpecID] = " & sClientCompanyCatSpecID & ", "
'    sSQL = sSQL & "[AdjusterSpecID] = " & sAdjusterSpecID & ", "
'    sSQL = sSQL & "[AdjusterSpecIDDisplay] =" & sAdjusterSpecIDDisplay & ", "
'    sSQL = sSQL & "[SPVersion] = " & sSPVersion & ", "
'    sSQL = sSQL & "[IBNUM] = " & sIBNUM & ", "
'    sSQL = sSQL & "[CLIENTNUM] = " & sCLIENTNUM & ", "
'    sSQL = sSQL & "[PolicyNo] = " & sPolicyNo & ", "
'    sSQL = sSQL & "[PolicyDescription] = " & sPolicyDescription & ", "
'    sSQL = sSQL & "[Insured] = " & sInsured & ", "
'    sSQL = sSQL & "[MailingAddress] = " & sMailingAddress & ", "
'    sSQL = sSQL & "[MAStreet] = " & sMAStreet & ", "
'    sSQL = sSQL & "[MACity] = " & sMACity & ", "
'    sSQL = sSQL & "[MAState] = " & sMAState & ", "
'    sSQL = sSQL & "[MAZIP] = " & sMAZIP & ", "
'    sSQL = sSQL & "[MAZIP4] = " & sMAZIP4 & ", "
'    sSQL = sSQL & "[MAOtherPostCode] = " & sMAOtherPostCode & ", "
'    sSQL = sSQL & "[HomePhone]  = " & sHomePhone & ", "
'    sSQL = sSQL & "[BusinessPhone] = " & sBusinessPhone & ", "
'    sSQL = sSQL & "[PropertyAddress] = " & sPropertyAddress & ", "
'    sSQL = sSQL & "[PAStreet]  = " & sPAStreet & ", "
'    sSQL = sSQL & "[PACity]  = " & sPACity & ", "
'    sSQL = sSQL & "[PAState] = " & sPAState & ", "
'    sSQL = sSQL & "[PAZIP]  = " & sPAZIP & ", "
'    sSQL = sSQL & "[PAZIP4] = " & sPAZIP4 & ", "
'    sSQL = sSQL & "[PAOtherPostCode]  = " & sPAOtherPostCode & ", "
'    sSQL = sSQL & "[MortgageeName]  = " & sMortgageeName & ", "
'    sSQL = sSQL & "[AgentNo]  = " & sAgentNo & ", "
'    sSQL = sSQL & "[ReportedBy] = " & sReportedBy & ", "
'    sSQL = sSQL & "[ReportedByPhone] = " & sReportedByPhone & ", "
'    sSQL = sSQL & "[Deductible]  = " & sDeductible & ", "
'    sSQL = sSQL & "[AppDedClassTypeIDOrder] = " & sAppDedClassTypeIDOrder & ", "
'    sSQL = sSQL & "[LRFormat]  = " & sLRFormat & ", "
'    sSQL = sSQL & "[LossReport] = " & sLossReport & ", "
'    sSQL = sSQL & "[DownLoadLossReport] = " & sDownLoadLossReport & ", "
'    sSQL = sSQL & "[UpLoadLossReport] = " & sUpLoadLossReport & ", "
'    sSQL = sSQL & "[StatusID]  = " & sStatusID & ", "
'    sSQL = sSQL & "[TypeOfLossID] = " & sTypeOfLossID & ", "
'    sSQL = sSQL & "[XactTypeOfLoss] = " & sXactTypeOfLoss & ", "
'    sSQL = sSQL & "[SentToXact] = " & sSentToXact & ", "
'    sSQL = sSQL & "[LossDate] = " & sLossDate & ", "
'    sSQL = sSQL & "[AssignedDate] = " & sAssignedDate & ", "
'    sSQL = sSQL & "[ReceivedDate] = " & sReceivedDate & ", "
'    sSQL = sSQL & "[ContactDate] = " & sContactDate & ", "
'    sSQL = sSQL & "[InspectedDate] = " & sInspectedDate & ", "
'    sSQL = sSQL & "[CloseDate]  = " & sCloseDate & ", "
'    sSQL = sSQL & "[Reassigned]  = " & sReassigned & ", "
'    sSQL = sSQL & "[DateReassigned] = " & sDateReassigned & ", "
'    sSQL = sSQL & "[RAAdjusterSpecID] = " & sRAAdjusterSpecID & ", "
'    sSQL = sSQL & "[IsLocked] = " & sIsLocked & ", "
'    sSQL = sSQL & "[IsDeleted] = " & sIsDeleted & ", "
'    sSQL = sSQL & "[DownLoadMe] = " & sDownLoadMe & ", "
'    sSQL = sSQL & "[UpLoadMe] = " & sUpLoadMe & ", "
'    sSQL = sSQL & "[DownLoadAll] = " & sDownLoadAll & ", "
'    sSQL = sSQL & "[UpLoadAll] = " & sUpLoadAll & ", "
'    sSQL = sSQL & "[AdminComments] = " & sAdminComments & ", "
'    sSQL = sSQL & "[MiscDelimSettings] = " & sMiscDelimSettings & ", "
'    sSQL = sSQL & "[DateLastUpdated] = " & sDateLastUpdated & ", "
'    sSQL = sSQL & "[UpdateByUserID] = " & sUID & " "
'    sSQL = sSQL & "WHERE AssignmentsID = " & sAssignmentsID & " "
'
'
'    Set oConn = New ADODB.Connection
'    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
'
'    oConn.Execute sSQL
    
    cmdSave.Enabled = False
    SaveMe = True
    
    'cleanup
    Set oConn = Nothing
'    Set MyadoRSAssignments = Nothing
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function SaveMe"
End Function

Public Function CheckStatus() As Boolean
    On Error GoTo EH
'    Dim lTabsPos As Long
'    Dim oFrame As Control
'    Dim MyFrame As Frame
'    Dim oControl As Control
'    Dim MyTextBox As TextBox
'    Dim MycmdButton As CommandButton
'    Dim sFrameName As String
'
'     'If this claim is closed only certain things can be edited
'    If mfrmClaim.MyStatus = iAssignmentsStatus_CLOSED Then
'        For lTabsPos = 1 To TSClaimInfo.Tabs.Count
'            Select Case UCase(TSClaimInfo.Tabs(lTabsPos).Tag)
'                Case UCase(framSpecifics.Name), _
'                        UCase(framInsuredInfo.Name), _
'                        UCase(framPolicyLimits.Name)
'                    sFrameName = TSClaimInfo.Tabs(lTabsPos).Tag
'                    For Each oFrame In Me.Controls
'                        If TypeOf oFrame Is Frame Then
'                            Set MyFrame = oFrame
'                            If StrComp(MyFrame.Name, sFrameName, vbTextCompare) = 0 Then
'                                MyFrame.Enabled = False
'                                Exit For
'                            End If
'                        End If
'                    Next
'                Case UCase(framDates.Name)
'                    'Need to disable all dates except the closedate
'                    For Each oControl In Me.Controls
'                        If TypeOf oControl Is TextBox Then
'                            Set MyTextBox = oControl
'                            If StrComp(MyTextBox.Tag, "Date", vbTextCompare) = 0 Then
'                                If StrComp(MyTextBox.Name, txtCloseDate.Name, vbTextCompare) <> 0 Then
'                                    MyTextBox.Enabled = False
'                                End If
'                            End If
'                        ElseIf TypeOf oControl Is CommandButton Then
'                            Set MycmdButton = oControl
'                            If StrComp(MycmdButton.Tag, "Date", vbTextCompare) = 0 Then
'                                If StrComp(MycmdButton.Name, cmdCloseDate.Name, vbTextCompare) <> 0 Then
'                                    MycmdButton.Enabled = False
'                                End If
'                            End If
'                        End If
'                    Next
'                Case UCase(framLossReport.Name)
'                    'Need to disable all control except the closedate
'                    For Each oControl In Me.Controls
'                        If TypeOf oControl Is CommandButton Then
'                            Set MycmdButton = oControl
'                            If StrComp(MycmdButton.Name, cmdViewPDFLossReport.Name, vbTextCompare) = 0 Then
'                                MycmdButton.Enabled = True
'                            End If
'                        Else
'                            If (Not TypeOf oControl Is ImageList) And (Not TypeOf oControl Is TabStrip) Then
'                                If oControl.Container.Name = framLossReport.Name Then
'                                    oControl.Enabled = False
'                                End If
'                            End If
'                        End If
'                    Next
'            End Select
'        Next
'    End If
    
    CheckStatus = True
    
    'cleanup
'    Set oFrame = Nothing
'    Set MyFrame = Nothing
'    Set oControl = Nothing
'    Set MyTextBox = Nothing
'    Set MycmdButton = Nothing
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function CheckStatus"
End Function


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo EH
    
    Select Case UnloadMode
        Case vbFormControlMenu
            Cancel = True
            mbUnloadMe = True
            Me.Visible = False
            mfrmClaim.Timer_UnloadForm.Enabled = True
        Case Else
            CLEANUP
    End Select
    
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_QueryUnload"
End Sub

Private Sub Form_Resize()
    ReSizeMe
End Sub

Public Sub ReSizeMe()
    On Error Resume Next
    Dim sNavScreenPos As String
    
    If Me.WindowState <> vbMinimized Then
        Me.Visible = False
    End If
    
    If Me.WindowState <> vbMinimized And Me.WindowState <> vbMaximized Then
        Me.Width = Screen.Width - 10
        Me.Height = Screen.Height - (10 + mfrmClaim.Height + goUtil.utGetTaskbarHeight)
        Me.top = mfrmClaim.top + mfrmClaim.Height
        Me.left = 10
    End If
    
    If Me.WindowState <> vbMinimized Then
        Me.Visible = True
    End If
End Sub

Public Function CLEANUP() As Boolean
    On Error GoTo EH
    If cmdSave.Enabled Then
        SaveMe
        If Not mfrmClaim Is Nothing Then
            mfrmClaim.RefreshMe
        End If
    End If
    Set mfrmClaim = Nothing
    Set MyGUI = Nothing
'    Set madoRSPolicyLimits = Nothing
'    Set mitmXPLSelected = Nothing
    Set moCurrentTextBox = Nothing
    
    CLEANUP = True
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function CLEANUP"
End Function




