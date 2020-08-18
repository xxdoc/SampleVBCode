VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCarGUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Carrier GUI Polymorphism
'Carrier GUI InterFace Definiton
'Use Implements V2ECKeyBoard.clsCarGUI in ALL clsGUI For All Carrier GUI Objects

'Implements V2ECKeyBoard.clsCarGUI

'Private mfrmClaimsList As frmClaimsList
'Private madoRSTypeOfLoss As ADODB.Recordset
'Private madoRSAssignmentType As ADODB.Recordset
'Private madoRSACID As ADODB.Recordset
'Private madoRSCatCode As ADODB.Recordset
'Private madoRSState As ADODB.Recordset
'Private madoClassType As ADODB.Recordset
'Private madoFeeSchedule As ADODB.Recordset
'Private madoFeeScheduleLevels As ADODB.Recordset
'Private madoFeeScheduleFeeTypes As ADODB.Recordset

Public Property Let adoFeeScheduleFeeTypes(padoRS As ADODB.Recordset)
'    Set madoFeeScheduleFeeTypes = padoRS
End Property
Public Property Set adoFeeScheduleFeeTypes(padoRS As ADODB.Recordset)
'    Set madoFeeScheduleFeeTypes = padoRS
End Property
Public Property Get adoFeeScheduleFeeTypes() As ADODB.Recordset
'    Set adoFeeScheduleFeeTypes = madoFeeScheduleFeeTypes
End Property

Public Property Let adoFeeScheduleLevels(padoRS As ADODB.Recordset)
'    Set madoFeeScheduleLevels = padoRS
End Property
Public Property Set adoFeeScheduleLevels(padoRS As ADODB.Recordset)
'    Set madoFeeScheduleLevels = padoRS
End Property
Public Property Get adoFeeScheduleLevels() As ADODB.Recordset
'    Set adoFeeScheduleLevels = madoFeeScheduleLevels
End Property

Public Property Let adoFeeSchedule(padoRS As ADODB.Recordset)
'    Set madoFeeSchedule = padoRS
End Property
Public Property Set adoFeeSchedule(padoRS As ADODB.Recordset)
'    Set madoFeeSchedule = padoRS
End Property
Public Property Get adoFeeSchedule() As ADODB.Recordset
'    Set adoFeeSchedule = madoFeeSchedule
End Property

Public Property Let adoClassType(padoRS As ADODB.Recordset)
'    Set madoClassType = padoRS
End Property
Public Property Set adoClassType(padoRS As ADODB.Recordset)
'    Set madoClassType = padoRS
End Property
Public Property Get adoClassType() As ADODB.Recordset
'    Set adoClassType = madoClassType
End Property

Public Property Let adoRSState(padoRS As ADODB.Recordset)
'    Set madoRSState = padoRS
End Property
Public Property Set adoRSState(padoRS As ADODB.Recordset)
'    Set madoRSState = padoRS
End Property
Public Property Get adoRSState() As ADODB.Recordset
'    Set adoRSState = madoRSState
End Property

Public Property Let adoRSCatCode(padoRS As ADODB.Recordset)
'    Set madoRSCatCode = padoRS
End Property
Public Property Set adoRSCatCode(padoRS As ADODB.Recordset)
'    Set madoRSCatCode = padoRS
End Property
Public Property Get adoRSCatCode() As ADODB.Recordset
'    Set adoRSCatCode = madoRSCatCode
End Property

Public Property Let adoRSACID(padoRS As ADODB.Recordset)
'    Set madoRSACID = padoRS
End Property
Public Property Set adoRSACID(padoRS As ADODB.Recordset)
'    Set madoRSACID = padoRS
End Property
Public Property Get adoRSACID() As ADODB.Recordset
'    Set adoRSACID = madoRSACID
End Property

Public Property Let adoRSAssignmentType(padoRS As ADODB.Recordset)
'    Set madoRSAssignmentType = padoRS
End Property
Public Property Set adoRSAssignmentType(padoRS As ADODB.Recordset)
'    Set madoRSAssignmentType = padoRS
End Property
Public Property Get adoRSAssignmentType() As ADODB.Recordset
'    Set adoRSAssignmentType = madoRSAssignmentType
End Property

Public Property Let adoRSTypeOfLoss(padoRS As ADODB.Recordset)
'    Set madoRSTypeOfLoss = padoRS
End Property
Public Property Set adoRSTypeOfLoss(padoRS As ADODB.Recordset)
'    Set madoRSTypeOfLoss = padoRS
End Property
Public Property Get adoRSTypeOfLoss() As ADODB.Recordset
'    Set adoRSTypeOfLoss = madoRSTypeOfLoss
End Property

Public Property Get MyClaimsList() As Object
'    Set MyClaimsList = mfrmClaimsList
End Property

Public Property Get ClassName() As String
'    ClassName = App.EXEName & ".clsGUI"
End Property

Public Property Get goForms() As Object
'    Set clsCarGUI_goForms = Forms
End Property

Public Sub ShowForm(psFormName As String, Optional pMode As VBRUN.FormShowConstants = vbModal, Optional oMnuItem As Object)
'    On Error GoTo EH
'    Dim lErrNum As Long
'    Dim sErrDesc As String
'
'    If Not mfrmClaimsList Is Nothing Then
'        mfrmClaimsList.Show
'        If mfrmClaimsList.WindowState <> vbMaximized Then
'            mfrmClaimsList.WindowState = vbNormal
'        End If
'        Exit Sub
'    End If
'
'    Set mfrmClaimsList = New frmClaimsList
'
'    Load mfrmClaimsList
'    mfrmClaimsList.Show piMode
'
'    Exit Sub
'EH:
'    lErrNum = Err.Number
'    sErrDesc = Err.Description
'    Err.Raise lerrnum, , App.EXEName & vbCrLf & ClassName & vbCrLf & "Public Sub ShowClaimsList" & vbCrLf & "Error # " & lerrnum & vbCrLf & serrdesc & vbCrLf
End Sub

Public Sub ShowAllForms(oECTray As Object, psNavPos As String, Optional psSkipForms As String)
'    goUtil.utShowAllForms Forms, oECTray, psNavPos, psSkipForms
End Sub

Public Sub HideAllForms(Optional psSkipForms As String)
'    goUtil.utHideAllForms Forms, psSkipForms
End Sub

Public Function SetGlobalObjects(pcolObjects As Collection) As Boolean
'    On Error Resume Next
'
'    Set goUtil = pcolObjects("goUtil")
    
End Function

Public Function SetUtilObject(oUtilObject As Object) As Boolean
'    On Error Resume Next
'    Set goUtil = oUtilObject
'    If Err.Number = 0 Then
'        SetUtilObject = True
'    Else
'        SetUtilObject = False
'    End If
End Function

Public Function SetadoRSTypeOfLoss() As Boolean
'    On Error GoTo EH
'    Dim oConn As ADODB.Connection
'    Dim sSQL As String
'    Dim lErrNum As Long
'    Dim sErrDesc As String
'
'    'rese the typeof loss rs
'    If Not adoRSTypeOfLoss Is Nothing Then
'        Set adoRSTypeOfLoss = Nothing
'    End If
'
'    Set oConn = New ADODB.Connection
'    Set adoRSTypeOfLoss = New ADODB.Recordset
'    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
'
'    sSQL = "SELECT RetTypeOfLoss.* "
'    sSQL = sSQL & "FROM( "
'    sSQL = sSQL & "SELECT "
'    sSQL = sSQL & "[TypeOfLossID], "
'    sSQL = sSQL & "(SELECT    NAME "
'    sSQL = sSQL & "FROM       Company "
'    sSQL = sSQL & "WHERE CompanyID = T.[ClientCompanyID]) As ClientCompanyName, "
'    sSQL = sSQL & "[ClientCompanyID], "
'    sSQL = sSQL & "[TypeOfLoss], "
'    sSQL = sSQL & "[Description], "
'    sSQL = sSQL & "[Code], "
'    sSQL = sSQL & "[IsDeleted] , "
'    sSQL = sSQL & "[DateLastUpdated] , "
'    sSQL = sSQL & "(SELECT    USERNAME "
'    sSQL = sSQL & "FROM       USERS "
'    sSQL = sSQL & "WHERE  USERSID = T.[UpdateByUserID]) As [UpdateByUserName],  "
'    sSQL = sSQL & "[UpdateByUserID]  "
'    sSQL = sSQL & "FROM TypeOfLoss T "
'    sSQL = sSQL & ") RetTypeOfLoss "
'    sSQL = sSQL & "WHERE ClientCompanyID = " & goUtil.gsCurCar & " "
'    sSQL = sSQL & "AND IsDeleted = False "
'
'
'    adoRSTypeOfLoss.CursorLocation = adUseClient
'    adoRSTypeOfLoss.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
'    Set adoRSTypeOfLoss.ActiveConnection = Nothing
'
'    SetadoRSTypeOfLoss = True
'    Set oConn = Nothing
'    Exit Function
'EH:
'    lErrNum = Err.Number
'    sErrDesc = Err.Description
'    Set oConn = Nothing
'    SetadoRSTypeOfLoss = False
'    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarGUI_ClassName & vbCrLf & "Public Function SetadoRSTypeOfLoss"
End Function

Public Function SetadoRSAssignmentType() As Boolean
'    On Error GoTo EH
'    Dim oConn As ADODB.Connection
'    Dim sSQL As String
'    Dim lErrNum As Long
'    Dim sErrDesc As String
'
'    'rese the typeof loss rs
'    If Not adoRSAssignmentType Is Nothing Then
'        Set adoRSAssignmentType = Nothing
'    End If
'
'    Set oConn = New ADODB.Connection
'    Set adoRSAssignmentType = New ADODB.Recordset
'    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
'
'    sSQL = "SELECT RetAssignmentType.* "
'    sSQL = sSQL & "FROM( "
'    sSQL = sSQL & "SELECT "
'    sSQL = sSQL & "[AssignmentTypeID], "
'    sSQL = sSQL & "[Type], "
'    sSQL = sSQL & "[Description], "
'    sSQL = sSQL & "[AdminComments], "
'    sSQL = sSQL & "[IsDeleted] , "
'    sSQL = sSQL & "[DateLastUpdated] , "
'    sSQL = sSQL & "(SELECT    USERNAME "
'    sSQL = sSQL & "FROM       USERS "
'    sSQL = sSQL & "WHERE  USERSID = S.[UpdateByUserID]) As [UpdateByUserName],  "
'    sSQL = sSQL & "[UpdateByUserID]  "
'    sSQL = sSQL & "FROM AssignmentType S "
'    sSQL = sSQL & ") RetAssignmentType "
'    sSQL = sSQL & "WHERE IsDeleted = False "
'
'
'    adoRSAssignmentType.CursorLocation = adUseClient
'    adoRSAssignmentType.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
'    Set adoRSAssignmentType.ActiveConnection = Nothing
'
'    SetadoRSAssignmentType = True
'    Set oConn = Nothing
'    Exit Function
'EH:
'    lErrNum = Err.Number
'    sErrDesc = Err.Description
'    Set oConn = Nothing
'    SetadoRSAssignmentType = False
'    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarGUI_ClassName & vbCrLf & "Public Function SetadoRSAssignmentType"
End Function

Public Function SetadoRSACID() As Boolean
'    On Error GoTo EH
'    Dim oConn As ADODB.Connection
'    Dim sSQL As String
'    Dim lErrNum As Long
'    Dim sErrDesc As String
'
'    'rese the typeof loss rs
'    If Not adoRSACID Is Nothing Then
'        Set adoRSACID = Nothing
'    End If
'
'    Set oConn = New ADODB.Connection
'    Set adoRSACID = New ADODB.Recordset
'    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
'
'    sSQL = "SELECT RetClientCoAdjusterSpec.* "
'    sSQL = sSQL & "FROM( "
'    sSQL = sSQL & "SELECT "
'    sSQL = sSQL & "ClientCoAdjusterSpecID, "
'    sSQL = sSQL & "(SELECT    NAME "
'    sSQL = sSQL & "FROM   Company "
'    sSQL = sSQL & "WHERE CompanyID = C.[ClientCompanyID]) As ClientCompanyName, "
'    sSQL = sSQL & "(SELECT    IsClientOF "
'    sSQL = sSQL & "FROM Company "
'    sSQL = sSQL & "WHERE CompanyID = C.[ClientCompanyID]) As IsClientOF, "
'    sSQL = sSQL & "[ClientCompanyID], "
'    sSQL = sSQL & "(SELECT    USERNAME "
'    sSQL = sSQL & "FROM   USERS "
'    sSQL = sSQL & "WHERE  USERSID = C.[UsersID]) As [UserName],  "
'    sSQL = sSQL & "(SELECT    LastName + ', ' + FirstName "
'    sSQL = sSQL & "FROM   USERS "
'    sSQL = sSQL & "WHERE  USERSID = C.[UsersID]) As [LFName],  "
'    sSQL = sSQL & "[UsersID], "
'    sSQL = sSQL & "[ACID], "
'    sSQL = sSQL & "[ACIDDescription], "
'    sSQL = sSQL & "[Comments], "
'    sSQL = sSQL & "[ClientCompanyCatSpecID], "
'    sSQL = sSQL & "[ZipCode], "
'    sSQL = sSQL & "[ZipCodeCount], "
'    sSQL = sSQL & "[ActiveDate] , "
'    sSQL = sSQL & "[InactiveDate] , "
'    sSQL = sSQL & "[DateLastUpdated] , "
'    sSQL = sSQL & "(SELECT    USERNAME "
'    sSQL = sSQL & "FROM   USERS "
'    sSQL = sSQL & "WHERE  USERSID = C.[UpdateByUserID]) As [UpdateByUserName],  "
'    sSQL = sSQL & "[UpdateByUserID] "
'    sSQL = sSQL & "FROM ClientCoAdjusterSpec C WHERE [ZipCode] Is Null "
'    sSQL = sSQL & ") RetClientCoAdjusterSpec "
'    sSQL = sSQL & "WHERE ClientCompanyID = " & goUtil.gsCurCar & " "
'    sSQL = sSQL & "AND UsersID = " & goUtil.gsCurUsersID & " "
'
'    adoRSACID.CursorLocation = adUseClient
'    adoRSACID.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
'    Set adoRSACID.ActiveConnection = Nothing
'
'    SetadoRSACID = True
'
'    Set oConn = Nothing
'    Exit Function
'EH:
'    lErrNum = Err.Number
'    sErrDesc = Err.Description
'    Set oConn = Nothing
'    SetadoRSACID = False
'    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarGUI_ClassName & vbCrLf & "Public Function SetadoRSACID"
End Function

Public Function SetadoRSCatCode() As Boolean
'    On Error GoTo EH
'    Dim oConn As ADODB.Connection
'    Dim sSQL As String
'    Dim lErrNum As Long
'    Dim sErrDesc As String
'
'    'rese the typeof loss rs
'    If Not adoRSCatCode Is Nothing Then
'        Set adoRSCatCode = Nothing
'    End If
'
'    Set oConn = New ADODB.Connection
'    Set adoRSCatCode = New ADODB.Recordset
'    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
'
'    sSQL = "SELECT RetClientCompanyCatSpec.* "
'    sSQL = sSQL & "FROM( "
'    sSQL = sSQL & "SELECT "
'    sSQL = sSQL & "ClientCompanyCatSpecID, "
'    sSQL = sSQL & "(SELECT    NAME "
'    sSQL = sSQL & "FROM   Company "
'    sSQL = sSQL & "WHERE CompanyID = C.[ClientCompanyID]) As ClientCompanyName, "
'    sSQL = sSQL & "(SELECT    IsClientOF "
'    sSQL = sSQL & "FROM Company "
'    sSQL = sSQL & "WHERE CompanyID = C.[ClientCompanyID]) As IsClientOF, "
'    sSQL = sSQL & "[ClientCompanyID], "
'    sSQL = sSQL & "(SELECT    NAME "
'    sSQL = sSQL & "FROM   Cat "
'    sSQL = sSQL & "WHERE CatID = C.[CATID]) As CatName, "
'    sSQL = sSQL & "[CATID], "
'    sSQL = sSQL & "[CatCode], "
'    sSQL = sSQL & "[Comments], "
'    sSQL = sSQL & "[ActiveDate] , "
'    sSQL = sSQL & "[InactiveDate] , "
'    sSQL = sSQL & "[DateLastUpdated] , "
'    sSQL = sSQL & "(SELECT    USERNAME "
'    sSQL = sSQL & "FROM   USERS "
'    sSQL = sSQL & "WHERE  USERSID = C.[UpdateByUserID]) As [UpdateByUserName],  "
'    sSQL = sSQL & "[UpdateByUserID] "
'    sSQL = sSQL & "FROM ClientCompanyCatSpec C "
'    sSQL = sSQL & ") RetClientCompanyCatSpec "
'    sSQL = sSQL & "WHERE [InactiveDate] Is Null "
'    sSQL = sSQL & "AND ClientCompanyID = " & goUtil.gsCurCar & " "
'    sSQL = sSQL & "AND CATID = " & goUtil.gsCurCat & " "
'
'    adoRSCatCode.CursorLocation = adUseClient
'    adoRSCatCode.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
'    Set adoRSCatCode.ActiveConnection = Nothing
'
'    SetadoRSCatCode = True
'
'    Set oConn = Nothing
'    Exit Function
'EH:
'    lErrNum = Err.Number
'    sErrDesc = Err.Description
'    Set oConn = Nothing
'    SetadoRSCatCode = False
'    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarGUI_ClassName & vbCrLf & "Public Function SetadoRSCatCode"
End Function

Public Function SetadoRSState() As Boolean
'    On Error GoTo EH
'    Dim oConn As ADODB.Connection
'    Dim sSQL As String
'    Dim lErrNum As Long
'    Dim sErrDesc As String
'
'    'rese the typeof loss rs
'    If Not adoRSState Is Nothing Then
'        Set adoRSState = Nothing
'    End If
'
'    Set oConn = New ADODB.Connection
'    Set adoRSState = New ADODB.Recordset
'    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
'
'    sSQL = "SELECT RetState.* "
'    sSQL = sSQL & "FROM( "
'    sSQL = sSQL & "SELECT "
'    sSQL = sSQL & "[StateID], "
'    sSQL = sSQL & "[Name], "
'    sSQL = sSQL & "[Code], "
'    sSQL = sSQL & "[Comments], "
'    sSQL = sSQL & "[IsDeleted] , "
'    sSQL = sSQL & "[DateLastUpdated] , "
'    sSQL = sSQL & "(SELECT    USERNAME "
'    sSQL = sSQL & "FROM       USERS "
'    sSQL = sSQL & "WHERE  USERSID = S.[UpdateByUserID]) As [UpdateByUserName],  "
'    sSQL = sSQL & "[UpdateByUserID]  "
'    sSQL = sSQL & "FROM State S "
'    sSQL = sSQL & ") RetState "
'    sSQL = sSQL & "WHERE IsDeleted = 0 "
'
'    adoRSState.CursorLocation = adUseClient
'    adoRSState.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
'    Set adoRSState.ActiveConnection = Nothing
'
'    SetadoRSState = True
'
'    Set oConn = Nothing
'    Exit Function
'EH:
'    lErrNum = Err.Number
'    sErrDesc = Err.Description
'    Set oConn = Nothing
'    SetadoRSState = False
'    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarGUI_ClassName & vbCrLf & "Public Function SetadoRSState"
End Function

Public Function SetadoClassType() As Boolean
'    On Error GoTo EH
'    Dim oConn As ADODB.Connection
'    Dim sSQL As String
'    Dim lErrNum As Long
'    Dim sErrDesc As String
'
'    'rese the typeof loss rs
'    If Not adoClassType Is Nothing Then
'        Set adoClassType = Nothing
'    End If
'
'    Set oConn = New ADODB.Connection
'    Set adoClassType = New ADODB.Recordset
'    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
'
'    sSQL = "SELECT RetClassType.* "
'    sSQL = sSQL & "FROM( "
'    sSQL = sSQL & "SELECT "
'    sSQL = sSQL & "[ClassTypeID], "
'    sSQL = sSQL & "[Class], "
'    sSQL = sSQL & "[Description], "
'    sSQL = sSQL & "[IsDeleted] , "
'    sSQL = sSQL & "[DateLastUpdated] , "
'    sSQL = sSQL & "(SELECT    USERNAME "
'    sSQL = sSQL & "FROM       USERS "
'    sSQL = sSQL & "WHERE  USERSID = CT.[UpdateByUserID]) As [UpdateByUserName],  "
'    sSQL = sSQL & "[UpdateByUserID]  "
'    sSQL = sSQL & "FROM ClassType CT "
'    sSQL = sSQL & ") RetClassType "
'    sSQL = sSQL & "WHERE IsDeleted = 0 "
'
'    adoClassType.CursorLocation = adUseClient
'    adoClassType.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
'    Set adoClassType.ActiveConnection = Nothing
'
'    SetadoClassType = True
'
'    Set oConn = Nothing
'    Exit Function
'EH:
'    lErrNum = Err.Number
'    sErrDesc = Err.Description
'    Set oConn = Nothing
'    SetadoClassType = False
'    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarGUI_ClassName & vbCrLf & "Public Function SetadoClassType"
End Function

Public Function SetadoRSFeeScheduleFeeTypes(Optional psFeeScheduleID As String) As Boolean
'    On Error GoTo EH
'    Dim oConn As ADODB.Connection
'    Dim sSQL As String
'    Dim lErrNum As Long
'    Dim sErrDesc As String
'
'    'rese the typeof loss rs
'    If Not adoClassType Is Nothing Then
'        Set adoClassType = Nothing
'    End If
'
'    Set oConn = New ADODB.Connection
'    Set adoClassType = New ADODB.Recordset
'    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
'
'    sSQL = "SELECT RetClassType.* "
'    sSQL = sSQL & "FROM( "
'    sSQL = sSQL & "SELECT "
'    sSQL = sSQL & "[ClassTypeID], "
'    sSQL = sSQL & "[Class], "
'    sSQL = sSQL & "[Description], "
'    sSQL = sSQL & "[IsDeleted] , "
'    sSQL = sSQL & "[DateLastUpdated] , "
'    sSQL = sSQL & "(SELECT    USERNAME "
'    sSQL = sSQL & "FROM       USERS "
'    sSQL = sSQL & "WHERE  USERSID = CT.[UpdateByUserID]) As [UpdateByUserName],  "
'    sSQL = sSQL & "[UpdateByUserID]  "
'    sSQL = sSQL & "FROM ClassType CT "
'    sSQL = sSQL & ") RetClassType "
'    sSQL = sSQL & "WHERE IsDeleted = 0 "
'
'    adoClassType.CursorLocation = adUseClient
'    adoClassType.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
'    Set adoClassType.ActiveConnection = Nothing
'
'    SetadoClassType = True
'
'    Set oConn = Nothing
'    Exit Function
'EH:
'    lErrNum = Err.Number
'    sErrDesc = Err.Description
'    Set oConn = Nothing
'    SetadoClassType = False
'    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarGUI_ClassName & vbCrLf & "Public Function SetadoClassType"
End Function

Public Function SetadoRSFeeScheduleLevels(Optional psFeeScheduleID As String) As Boolean
'    On Error GoTo EH
'    Dim oConn As ADODB.Connection
'    Dim sSQL As String
'    Dim lErrNum As Long
'    Dim sErrDesc As String
'
'    'rese the typeof loss rs
'    If Not adoClassType Is Nothing Then
'        Set adoClassType = Nothing
'    End If
'
'    Set oConn = New ADODB.Connection
'    Set adoClassType = New ADODB.Recordset
'    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
'
'    sSQL = "SELECT RetClassType.* "
'    sSQL = sSQL & "FROM( "
'    sSQL = sSQL & "SELECT "
'    sSQL = sSQL & "[ClassTypeID], "
'    sSQL = sSQL & "[Class], "
'    sSQL = sSQL & "[Description], "
'    sSQL = sSQL & "[IsDeleted] , "
'    sSQL = sSQL & "[DateLastUpdated] , "
'    sSQL = sSQL & "(SELECT    USERNAME "
'    sSQL = sSQL & "FROM       USERS "
'    sSQL = sSQL & "WHERE  USERSID = CT.[UpdateByUserID]) As [UpdateByUserName],  "
'    sSQL = sSQL & "[UpdateByUserID]  "
'    sSQL = sSQL & "FROM ClassType CT "
'    sSQL = sSQL & ") RetClassType "
'    sSQL = sSQL & "WHERE IsDeleted = 0 "
'
'    adoClassType.CursorLocation = adUseClient
'    adoClassType.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
'    Set adoClassType.ActiveConnection = Nothing
'
'    SetadoClassType = True
'
'    Set oConn = Nothing
'    Exit Function
'EH:
'    lErrNum = Err.Number
'    sErrDesc = Err.Description
'    Set oConn = Nothing
'    SetadoClassType = False
'    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarGUI_ClassName & vbCrLf & "Public Function SetadoClassType"
End Function

Public Function SetadoRSFeeSchedule(Optional psFeeScheduleID As String) As Boolean
'    On Error GoTo EH
'    Dim oConn As ADODB.Connection
'    Dim sSQL As String
'    Dim lErrNum As Long
'    Dim sErrDesc As String
'
'    'rese the typeof loss rs
'    If Not adoClassType Is Nothing Then
'        Set adoClassType = Nothing
'    End If
'
'    Set oConn = New ADODB.Connection
'    Set adoClassType = New ADODB.Recordset
'    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
'
'    sSQL = "SELECT RetClassType.* "
'    sSQL = sSQL & "FROM( "
'    sSQL = sSQL & "SELECT "
'    sSQL = sSQL & "[ClassTypeID], "
'    sSQL = sSQL & "[Class], "
'    sSQL = sSQL & "[Description], "
'    sSQL = sSQL & "[IsDeleted] , "
'    sSQL = sSQL & "[DateLastUpdated] , "
'    sSQL = sSQL & "(SELECT    USERNAME "
'    sSQL = sSQL & "FROM       USERS "
'    sSQL = sSQL & "WHERE  USERSID = CT.[UpdateByUserID]) As [UpdateByUserName],  "
'    sSQL = sSQL & "[UpdateByUserID]  "
'    sSQL = sSQL & "FROM ClassType CT "
'    sSQL = sSQL & ") RetClassType "
'    sSQL = sSQL & "WHERE IsDeleted = 0 "
'
'    adoClassType.CursorLocation = adUseClient
'    adoClassType.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
'    Set adoClassType.ActiveConnection = Nothing
'
'    SetadoClassType = True
'
'    Set oConn = Nothing
'    Exit Function
'EH:
'    lErrNum = Err.Number
'    sErrDesc = Err.Description
'    Set oConn = Nothing
'    SetadoClassType = False
'    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarGUI_ClassName & vbCrLf & "Public Function SetadoClassType"
End Function

Public Function CLEANUP() As Boolean
'    On Error GoTo EH
'    Dim lErrNum As Long
'    Dim sErrDesc As String
'
'    If Not mfrmClaimsList Is Nothing Then
'        Unload mfrmClaimsList
'        Set mfrmClaimsList = Nothing
'    End If
'
'    CLEANUP = True
'    Exit Function
'EH:
'    lErrNum = Err.Number
'    sErrDesc = Err.Description
'    Err.Raise lerrnum, , serrdesc & vbCrLf & ClassName & vbCrLf & "Public Function CLEANUP"
End Function

Public Function ShowCalendar(poTextBox As Object) As Boolean
'    On Error GoTo EH
'    Dim oCalendar As V2ECKeyBoard.clsCalendar
'    Dim oTextBox As TextBox
'    Dim lErrNum As Long
'    Dim sErrDesc As String
'
'    Set oCalendar = New V2ECKeyBoard.clsCalendar
'
'    If TypeOf poTextBox Is TextBox Then
'        Set oTextBox = poTextBox
'        If Not IsDate(oTextBox.Text) Then
'            oTextBox.Text = vbNullString
'            oCalendar.CurrentDate = Now()
'        Else
'            oCalendar.CurrentDate = oTextBox.Text
'        End If
'        oCalendar.ShowCalendar
'        If oCalendar.CurrentDate = NULL_DATE Then
'            If IsDate(oTextBox.Text) Then
'                If CDate(oTextBox.Text) = NULL_DATE Then
'                    oTextBox.Text = vbNullString
'                End If
'            Else
'                oTextBox.Text = vbNullString
'            End If
'        Else
'            oTextBox.Text = Format(oCalendar.CurrentDate, "MM/DD/YYYY")
'        End If
'    End If
'
'    ShowCalendar = True
'
'    'cleanup
'    Set oCalendar = Nothing
'
'    Exit Function
'EH:
'    lErrNum = Err.Number
'    sErrDesc = Err.Description
'    ShowCalendar = False
'    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarGUI_ClassName & vbCrLf & "Public Function ShowCalendar"
End Function



