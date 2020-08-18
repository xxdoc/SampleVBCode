VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRptDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Implements V2ECKeyBoard.clsRptDB
''ACCESS
'Private mCurDB As DAO.Database
'Private RS As DAO.Recordset
''SQL SERVER
'Private mConn As ADODB.Connection
'Private mRS As ADODB.Recordset
'
'Private mARV As V2ARViewer.clsARViewer
'Private moLRs As V2ECKeyBoard.clsLossReports
'Private msAssignmentsID As String   'Use String for SQL Server DB ID
'Private msID As String              'Use String for USer DB ID
'Private msClientClaimNo As String
'Private msIBNo As String
'Private mlPrintCopies As Long
'Private mbPreview As Boolean
'Private mbShowPrintButton As Boolean
'Private mbGetObjectOnly as Boolean
'Private msXportPath As String
'Private mXportType As ExportType
'Private mlPDFJPEGQuality as Long
'Private moARObject As Object

Public Type MiscReportParam
    MiscReportParamID As String
    AssignmentsID As String
    ID As String
    IDAssignments As String
    Number As String
    ProjectName As String
    ClassName As String
    ParamName As String
    ParamCaption As String
    ParamDataType As VbVarType
    ParamValue As String
    SortMe As String
    IsDeleted As String
    DownLoadMe As String
    UpLoadMe As String
    AdminComments As String
    DateLastUpdated As String
    UpdateByUserID As String
End Type


Public Type rptMiscReportParam
    MiscReportParamID As String
    AssignmentsID As String
    ID As String
    IDAssignments As String
    ProjectName As String
    ClassName As String
    ParamName As String
    ParamCaption As String
    ParamDataType As VbVarType
    ParamValue As String
    SortMe As String
    IsDeleted As String
    DownLoadMe As String
    UpLoadMe As String
    AdminComments As String
    DateLastUpdated As String
    UpdateByUserID As String
End Type

Public Type rptMiscDelimParam
    ProjectName As String
    ClassName As String
    ParamName As String
    ParamCaption As String
    ParamDataType As VbVarType
    ParamValue As String
End Type

Public Property Get ClassName() As String
'    clsRptDB_ClassName = App.EXEName & ".clsRptDB"
End Property

Public Property Get MiscDelimParamsCol() As Collection
'   Set clsRptDB_MiscDelimParamsCol = GetMiscDelimParamsCol()
End Property

Public Property Let AssignmentsID(psAssignmentsID As String)
'    msAssignmentsID = psAssignmentsID
End Property
Public Property Get AssignmentsID() As String
'    clsRptDB_AssignmentsID = msAssignmentsID
End Property

Public Property Get ID() As String
'    clsRptDB_ID = msID
End Property
Public Property Let ID(psID As String)
'    msID = psID
End Property

Public Property Let ClientClaimNo(psClientClaimNo As String)
'    msClientClaimNo = psClientClaimNo
End Property
Public Property Get ClientClaimNo() As String
'    clsRptDB_ClientClaimNo = msClientClaimNo
End Property

Public Property Let IBNo(psIBNo As String)
'    msIBNo = psIBNo
End Property
Public Property Get IBNo() As String
'    clsRptDB_IBNo = msIBNo
End Property

Public Property Let sXportPath(psXportPath As String)
'    msXportPath = psXportPath
End Property
Public Property Get sXportPath() As String
'    clsRptDB_sXportPath = msXportPath
End Property

Public Property Let XportType(pXportType As V2ECKeyBoard.ExportType)
'    mXportType = pXportType
End Property
Public Property Get XportType() As V2ECKeyBoard.ExportType
'    clsRptDB_XportType = mXportType
End Property

Public Property Let PDFJPEGQuality(plPDFJPEGQuality As Long)
'    mlPDFJPEGQuality = plPDFJPEGQuality
End Property
Public Property Get PDFJPEGQuality() As Long
'    clsRptDB_PDFJPEGQuality = mlPDFJPEGQuality
End Property

Public Property Set CurDB(poDB As dao.Database)
'    Set mCurDB = poDB
End Property
Public Property Let CurDB(poDB As dao.Database)
'    Set mCurDB = poDB
End Property
Public Property Get CurDB() As dao.Database
'    Set clsRptDB_CurDB = mCurDB
End Property

Public Property Get ARV() As Object
'    Set clsRptDB_ARV = mARV
End Property
Public Property Let ARV(poARV As Object)
'    Set mARV = poARV
End Property
Public Property Set ARV(poARV As Object)
'    Set mARV = poARV
End Property

Public Property Let bPreView(pbFlag As Boolean)
'    mbPreview = pbFlag
End Property
Public Property Get bPreView() As Boolean
'    clsRptDB_bPreview = mbPreview
End Property

Public Property Let bShowPrintButton(pbFlag As Boolean)
'    mbShowPrintButton = pbFlag
End Property

Public Property Get bShowPrintButton() As Boolean
'    clsRptDB_bShowPrintButton = mbShowPrintButton
End Property

Public Property Let PrintCopies(plPrintCopies As Long)
'    mlPrintCopies = plPrintCopies
End Property
Public Property Get PrintCopies() As Long
'    clsRptDB_PrintCopies = mlPrintCopies
End Property

Public Property Let bGetObjectOnly(pbFlag As Boolean)
'    mbGetObjectOnly = pbFlag
End Property
Public Property Get bGetObjectOnly() As Boolean
'   clsRptDB_bGetObjectOnly = mbGetObjectOnly
End Property

Public Property Let ARObject(poARObject As Object)
'    SET moARObject = poARObject
End Property
Public Property Set ARObject(poARObject As Object)
'    SET moARObject = poARObject
End Property
Public Property Get ARObject() As Object
'   SET clsRptDB_ARObject = moARObject
End Property

Public Function OpenConnection() As Boolean
'    On Error GoTo EH
'    Dim sUserID As String
'    Dim sPassword As String
'    Dim sProdDSN As String
'    Dim lErrNum As Long
'    Dim sErrDesc As String
'
'    clsRptDB_OpenConnection = True
'    'ADO Object
'    If mConn Is Nothing Then
'        sUserID = goUtil.utGetECSCryptSetting("V2WebControl", "DBConn", "USERID")
'        sPassword = goUtil.utGetECSCryptSetting("V2WebControl", "DBConn", "PASSWORD")
'        sProdDSN = GetSetting("V2WebControl", "DSN", "NAME", vbNullString)
'        Set mConn = New ADODB.Connection
'        mConn.Open sProdDSN, sUserID, sPassword
'    End If
'
'    Exit Function
'EH:
'    lErrNum = Err.Number
'    sErrDesc = Err.Description
'    clsRptDB_OpenConnection = False
'    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsRptDB_ClassName & vbCrLf & "Private Function clsRptDB_OpenConnection"
End Function

Public Function SetUtilObject(oUtilObject As Object) As Boolean
'    On Error Resume Next
'    Set goUtil = oUtilObject
'    If Err.Number = 0 Then
'        clsRptDB_SetUtilObject = True
'    Else
'        clsRptDB_SetUtilObject = False
'    End If
End Function

Public Function CLEANUP() As Boolean
'    On Error Resume Next
'    If Not RS Is Nothing Then
'        RS.Close
'        Set RS = Nothing
'    End If
'    If Not mRS Is Nothing Then
'        mRS.Close
'        Set mRS = Nothing
'    End If
'    If Not mConn Is Nothing Then
'        mConn.Close
'        Set mConn = Nothing
'    End If
'    'Only set this ref to nothing (Don't want to close Adjuster's current DB session)
'    If Not mCurDB Is Nothing Then
'        Set mCurDB = Nothing
'    End If
'    'Only set this ref to nothing (Don't want to close Adjuster's current gARV)
'    If Not mARV Is Nothing Then
'        Set mARV = Nothing
'    End If
End Function

Public Function PrintARReport() As Boolean

End Function


