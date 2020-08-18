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
Implements V2ECKeyBoard.clsRptDB

'ACCESS
Private mCurDB As DAO.Database
Private RS As DAO.Recordset
'SQL SERVER
Private mConn As ADODB.Connection
Private mRS As ADODB.Recordset

Private mARV As V2ARViewer.clsARViewer
Private moLRs As V2ECKeyBoard.clsLossReports
Private msAssignmentsID As String   'Use String for SQL Server DB ID
Private msID As String              'Use String for USer DB ID
Private msClientClaimNo As String
Private msIBNo As String
Private mlPrintCopies As Long
Private mbPreview As Boolean
Private mbShowPrintButton As Boolean
Private mbGetObjectOnly As Boolean
Private msXportPath As String
Private mXportType As ExportType
Private mlPDFJPEGQuality As Long
Private moARObject As Object

'Report Specific version for Latest version
Private marRpt As ECrpt_arUsersInfo.arUsersInfo

'    [UsersID] [int] IDENTITY (1, 1) NOT NULL ,
'    [UserName] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
'    [PassWord] [varchar] (300) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
'    [FirstName] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
'    [LastName] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
'    [SSN] [int] NULL ,
'    [Email] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
'    [ContactPhone] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
'    [EmergencyPhone] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
'    [Address] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
'    [City] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
'    [State] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
'    [Zip] [int] NULL ,
'    [ZIP4] [int] NULL ,
'    [OtherPostCode] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
'    [Active] [bit] NOT NULL ,
'    [ActiveDate] [datetime] NOT NULL ,
'    [InactiveDate] [datetime] NULL ,
'    [SecurityLevel] [int] NOT NULL ,
'    [Comments] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
'    [DateLastUpdated] [datetime] NOT NULL ,
'    [UpdateByUserID] Not [Int]


Public Type udtUsers01
    UsersID As Long
    SecurityLevel As Long
    f_UserName As String
    f_PassWord As String
    f_FirstName As String
    f_LastName As String
    f_Email As String
    f_ContactPhone As String
    f_DateLastUpdated As String
End Type

Private Type udtarRpt
    f_PrintDate As Date
    f_gndTtlCatName As String
    f_gndTtlSiteLocation As String
    coludtUsers01 As Collection
End Type

Private msClientCompanyID As String
Private msCATID As String
Private msUSERSID As String
Private mbHideDetails As Boolean
Private mbExportXML As Boolean
Private mbExportXMLOnly As Boolean

Public Property Let HideDetails(pbFlag As Boolean)
    mbHideDetails = pbFlag
End Property
Public Property Get HideDetails() As Boolean
    HideDetails = mbHideDetails
End Property

Public Property Let UsersID(psUSERSID As String)
    msUSERSID = psUSERSID
End Property
Public Property Get UsersID() As String
    UsersID = msUSERSID
End Property

Public Property Let CATID(psCATID As String)
    msCATID = psCATID
End Property
Public Property Get CATID() As String
    CATID = msCATID
End Property

Public Property Let ClientCompanyID(psClientCompanyID As String)
    msClientCompanyID = psClientCompanyID
End Property
Public Property Get ClientCompanyID() As String
    ClientCompanyID = msClientCompanyID
End Property


Private Property Get clsRptDB_MiscDelimParamsCol() As Collection
    Dim bCleanup As Boolean
    'Add Report Parameters that are not updated by DB Fields.
    If marRpt Is Nothing Then
        bCleanup = True
        Set marRpt = New ECrpt_arUsersInfo.arUsersInfo
    End If
    Set clsRptDB_MiscDelimParamsCol = GetMiscDelimParamsCol()
    If bCleanup Then
        Set mConn = Nothing
        Set mRS = Nothing
        Unload marRpt
        Set marRpt = Nothing
    End If
End Property

Private Property Set clsRptDB_ARObject(poARObject As Object)
    Set moARObject = poARObject
End Property
Private Property Let clsRptDB_ARObject(poARObject As Object)
     Set moARObject = poARObject
End Property
Private Property Get clsRptDB_ARObject() As Object
    Set clsRptDB_ARObject = moARObject
End Property

Private Property Let clsRptDB_bGetObjectOnly(pbFlag As Boolean)
    mbGetObjectOnly = pbFlag
End Property
Private Property Get clsRptDB_bGetObjectOnly() As Boolean
    clsRptDB_bGetObjectOnly = mbGetObjectOnly
End Property

Private Property Get clsRptDB_ClassName() As String
    clsRptDB_ClassName = App.EXEName & ".clsRptDB"
End Property

Private Property Get clsRptDB_AssignmentsID() As String
    clsRptDB_AssignmentsID = msAssignmentsID
End Property
Private Property Let clsRptDB_AssignmentsID(psAssignmentsID As String)
    msAssignmentsID = psAssignmentsID
End Property

Private Property Get clsRptDB_ID() As String
    clsRptDB_ID = msID
End Property
Private Property Let clsRptDB_ID(psID As String)
    msID = psID
End Property

Private Property Let clsRptDB_ClientClaimNo(psClientClaimNo As String)
    msClientClaimNo = psClientClaimNo
End Property
Private Property Get clsRptDB_ClientClaimNo() As String
    clsRptDB_ClientClaimNo = msClientClaimNo
End Property

Private Property Let clsRptDB_IBNo(psIBNo As String)
    msIBNo = psIBNo
End Property
Private Property Get clsRptDB_IBNo() As String
    clsRptDB_IBNo = msIBNo
End Property

Private Property Let clsRptDB_sXportPath(psXportPath As String)
    msXportPath = psXportPath
End Property
Private Property Get clsRptDB_sXportPath() As String
    clsRptDB_sXportPath = msXportPath
End Property

Private Property Let clsRptDB_XportType(pXportType As V2ECKeyBoard.ExportType)
    mXportType = pXportType
End Property
Private Property Get clsRptDB_XportType() As V2ECKeyBoard.ExportType
    clsRptDB_XportType = mXportType
End Property

Private Property Let clsRptDB_PDFJPEGQuality(plPDFJPEGQuality As Long)
    mlPDFJPEGQuality = plPDFJPEGQuality
End Property
Private Property Get clsRptDB_PDFJPEGQuality() As Long
    clsRptDB_PDFJPEGQuality = mlPDFJPEGQuality
End Property

Private Property Set clsRptDB_CurDB(poDB As DAO.Database)
    Set mCurDB = poDB
End Property
Private Property Let clsRptDB_CurDB(poDB As DAO.Database)
    Set mCurDB = poDB
End Property
Private Property Get clsRptDB_CurDB() As DAO.Database
    Set clsRptDB_CurDB = mCurDB
End Property

Private Property Get clsRptDB_ARV() As Object
    Set clsRptDB_ARV = mARV
End Property
Private Property Let clsRptDB_ARV(poARV As Object)
    Set mARV = poARV
End Property
Private Property Set clsRptDB_ARV(poARV As Object)
    Set mARV = poARV
End Property

Private Property Let clsRptDB_bPreview(pbFlag As Boolean)
    mbPreview = pbFlag
End Property
Private Property Get clsRptDB_bPreview() As Boolean
    clsRptDB_bPreview = mbPreview
End Property

Private Property Let clsRptDB_bShowPrintButton(pbFlag As Boolean)
    mbShowPrintButton = pbFlag
End Property

Private Property Get clsRptDB_bShowPrintButton() As Boolean
    clsRptDB_bShowPrintButton = mbShowPrintButton
End Property

Private Property Let clsRptDB_PrintCopies(plPrintCopies As Long)
    mlPrintCopies = plPrintCopies
End Property
Private Property Get clsRptDB_PrintCopies() As Long
    clsRptDB_PrintCopies = mlPrintCopies
End Property

Public Property Let ExportXML(pbFlag As Boolean)
    mbExportXML = pbFlag
End Property
Public Property Get ExportXML() As Boolean
    ExportXML = mbExportXML
End Property

Public Property Let ExportXMLOnly(pbFlag As Boolean)
    mbExportXMLOnly = pbFlag
End Property
Public Property Get ExportXMLOnly() As Boolean
    ExportXMLOnly = mbExportXMLOnly
End Property

Private Function clsRptDB_OpenConnection() As Boolean
    On Error GoTo EH
    Dim sUserID As String
    Dim sPassword As String
    Dim sProdDSN As String
    Dim lErrNum As Long
    Dim sErrDesc As String

    clsRptDB_OpenConnection = True
    'ADO Object
    If mConn Is Nothing Then
        sUserID = goUtil.utGetECSCryptSetting("V2WebControl", "DBConn", "USERID")
        sPassword = goUtil.utGetECSCryptSetting("V2WebControl", "DBConn", "PASSWORD")
        sProdDSN = GetSetting("V2WebControl", "DSN", "NAME", vbNullString)
        Set mConn = New ADODB.Connection
        mConn.Open sProdDSN, sUserID, sPassword
    End If

    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    clsRptDB_OpenConnection = False
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsRptDB_ClassName & vbCrLf & "Private Function clsRptDB_OpenConnection"
End Function

Private Function clsRptDB_SetUtilObject(oUtilObject As Object) As Boolean
    On Error Resume Next
    Set goUtil = oUtilObject
    If Err.Number = 0 Then
        clsRptDB_SetUtilObject = True
    Else
        clsRptDB_SetUtilObject = False
    End If
End Function

Private Function clsRptDB_CLEANUP() As Boolean
    On Error Resume Next
    If Not RS Is Nothing Then
        RS.Close
        Set RS = Nothing
    End If
    If Not mRS Is Nothing Then
        mRS.Close
        Set mRS = Nothing
    End If
    If Not mConn Is Nothing Then
        mConn.Close
        Set mConn = Nothing
    End If
    'Only set this ref to nothing (Don't want to close Adjuster's current DB session)
    If Not mCurDB Is Nothing Then
        Set mCurDB = Nothing
    End If
    'Only set this ref to nothing (Don't want to close Adjuster's current gARV)
    If Not mARV Is Nothing Then
        Set mARV = Nothing
    End If
End Function

Private Function clsRptDB_PrintARReport() As Boolean
    On Error GoTo EH
    Dim udtarRpt As udtarRpt
    Dim colParams As Collection
    Dim vParams As Variant
    Dim MyParams As MiscReportParam
    Dim sXML As String
    Dim sXMLXportPath As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    Set marRpt = New ECrpt_arUsersInfo.arUsersInfo
    
    'First Set any parameters
    Set colParams = GetMiscDelimParamsCol()
    If Not colParams Is Nothing Then
        For Each vParams In colParams
            '2/8/2004 MiscReportParam , and MiscReportParam01 to MiscReportParam30
            If VarType(vParams) <> vbString Then
                MyParams = vParams
                marRpt.SetProperty MyParams.ParamName, MyParams.ParamValue, MyParams.ParamDataType
            End If
            '2/8/2004 MiscReportParam , and MiscReportParam01 to MiscReportParam30
        Next
    End If
    
    udtarRpt = PopulateudtarRpt
    
    With udtarRpt
        marRpt.SetProperty "f_PrintDate", .f_PrintDate, vbDate
        marRpt.SetProperty "f_gndTtlCatName", .f_gndTtlCatName, vbString
        marRpt.SetProperty "f_gndTtlSiteLocation", .f_gndTtlSiteLocation, vbString
        marRpt.SetProperty "coludtUsers01", .coludtUsers01, vbObject
    End With
    
        'If only getting the Object then do not print it
    With marRpt
        If Not goUtil.gMainDB Is Nothing Then
            '190  Printing on HP1218  5.15.2002
            'Need to use WIN default printer setting if it is selected under the menu
            On Error Resume Next
            .Printer.DeviceName = GetSetting(goUtil.gsAppEXEName, "PRINTER", "PRINTER_NAME", vbNullString)
            If Err.Number > 0 Then
                Err.Clear
            End If
            .Printer.Port = GetSetting(goUtil.gsAppEXEName, "PRINTER", "PRINTER_PORT", vbNullString)
            If Err.Number > 0 Then
                Err.Clear
            End If
        End If
        On Error GoTo EH
        .Printer.Orientation = ddOPortrait
        If mlPrintCopies > 0 Then
            .Printer.Copies = mlPrintCopies
        End If
    End With
    
    If mbGetObjectOnly Then
        Set moARObject = marRpt
        clsRptDB_PrintARReport = True
        Exit Function
    End If
    
    'Web DB
    Set moLRs = New V2ECKeyBoard.clsLossReports
    'Check for XML Export Only Flag
    If mbExportXML And mbExportXMLOnly Then
        GoTo EXPORT_XML_ONLY
    End If
    If moLRs.ExportFile(marRpt, msXportPath, mXportType, mlPDFJPEGQuality) Then
EXPORT_XML_ONLY:
        'Check for XML Export Flag
        If mbExportXML Then
            sXML = marRpt.GetXMLExport()
            If sXML <> vbNullString Then
                sXMLXportPath = msXportPath
                sXMLXportPath = Left(sXMLXportPath, InStrRev(sXMLXportPath, ".", , vbBinaryCompare))
                sXMLXportPath = sXMLXportPath & "xml"
                goUtil.utSaveFileData sXMLXportPath, sXML
            End If
        End If
        'End check for XML Export Flag
        clsRptDB_PrintARReport = True
    End If
    moLRs.CleanUp
    Set moLRs = Nothing
    Set colParams = Nothing
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Screen.MousePointer = vbDefault
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsRptDB_ClassName & vbCrLf & "Public Function PrintARReport"
End Function

Private Function PopulateudtarRpt() As udtarRpt
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim sSQL As String
    Dim MyUsers As udtUsers01
    
    'Check to see if this is Access DB.
    'If it is need to open connection to it
    If Not goUtil.gMainDB Is Nothing Then
        If mConn Is Nothing Then
            Set mConn = New ADODB.Connection
            goUtil.utOpenDatabaseADOConn mConn, goUtil.gMainDB.Name
        End If
    Else
        'WEB Server Connection
         clsRptDB_OpenConnection
    End If
    
    'B.  get the UsersInfo Report Info
    sSQL = "SELECT Top 1 "
    
'    UsersID As String
'    UserName As String
'    PassWord As String
'    FirstName As String
'    LastName As String
'    SSN As String
'    Email As String
'    ContactPhone As String
'    EmergencyPhone As String
'    ActiveDate As String
'    DateLastUpdated As String
    
    sSQL = sSQL & "U.[UsersID], "
    sSQL = sSQL & "U.[SecurityLevel], "
    sSQL = sSQL & "U.[UserName], "
    sSQL = sSQL & "U.[PassWord], "
    sSQL = sSQL & "U.[FirstName], "
    sSQL = sSQL & "U.[LastName], "
    sSQL = sSQL & "U.[Email], "
    sSQL = sSQL & "U.[ContactPhone], "
    sSQL = sSQL & "U.[DateLastUpdated], "
    sSQL = sSQL & "( "
    sSQL = sSQL & "SELECT   Name + ' - ' + Description "
    sSQL = sSQL & "FROM     CAT "
    sSQL = sSQL & "WHERE CATID = " & msCATID & " "
    sSQL = sSQL & ") As gndTtlCatName, "
    sSQL = sSQL & "( "
    sSQL = sSQL & "SELECT   SACity + ', ' + SAState "
    sSQL = sSQL & "FROM     ClientCompanyCat "
    sSQL = sSQL & "WHERE    ClientCompanyID = " & msClientCompanyID & " "
    sSQL = sSQL & "AND      CATID = " & msCATID & " "
    sSQL = sSQL & ") As gndTtlSiteLocation "
    sSQL = sSQL & "FROM Users U "
    If msUSERSID = "-1" Then
        sSQL = sSQL & "WHERE U.UsersID IN "
                        sSQL = sSQL & "( "
                        sSQL = sSQL & "SELECT UsersID FROM ClientCompanyUsersCAT "
                        sSQL = sSQL & "WHERE ClientCompanyID = " & msClientCompanyID & " "
                        sSQL = sSQL & "AND CATID = " & msCATID & " "
                        sSQL = sSQL & "AND ACTIVE = 1 "
                        sSQL = sSQL & ") "
    Else
        sSQL = sSQL & "WHERE U.UsersID = " & msUSERSID & " "
    End If
    
    Set mRS = New ADODB.Recordset
    mRS.CursorLocation = adUseClient
    mRS.Open sSQL, mConn, adOpenStatic, adLockReadOnly
    Set mRS.ActiveConnection = Nothing
    
    With PopulateudtarRpt
        .f_PrintDate = NULL_DATE
        .f_gndTtlCatName = vbNullString
        .f_gndTtlSiteLocation = vbNullString
        
        On Error Resume Next
        .f_PrintDate = Now()
        .f_gndTtlCatName = mRS!gndTtlCatName
        .f_gndTtlSiteLocation = mRS!gndTtlSiteLocation
        On Error GoTo EH
    End With
    
    sSQL = "SELECT "
    sSQL = sSQL & "U.[UsersID], "
    sSQL = sSQL & "U.[SecurityLevel], "
    sSQL = sSQL & "U.[UserName], "
    sSQL = sSQL & "U.[PassWord], "
    sSQL = sSQL & "U.[FirstName], "
    sSQL = sSQL & "U.[LastName], "
    sSQL = sSQL & "U.[SSN], "
    sSQL = sSQL & "U.[Email], "
    sSQL = sSQL & "U.[ContactPhone], "
    sSQL = sSQL & "U.[EmergencyPhone], "
    sSQL = sSQL & "U.[DateLastUpdated] "
    sSQL = sSQL & "FROM Users U "
    If msUSERSID = "-1" Then
        sSQL = sSQL & "WHERE U.UsersID IN "
                        sSQL = sSQL & "( "
                        sSQL = sSQL & "SELECT UsersID FROM ClientCompanyUsersCAT "
                        sSQL = sSQL & "WHERE ClientCompanyID = " & msClientCompanyID & " "
                        sSQL = sSQL & "AND CATID = " & msCATID & " "
                        sSQL = sSQL & "AND ACTIVE = 1 "
                        sSQL = sSQL & ") "
    Else
        sSQL = sSQL & "WHERE U.UsersID = " & msUSERSID & " "
    End If
    sSQL = sSQL & "Order By U.[LastName] "
    
    Set mRS = New ADODB.Recordset
    mRS.CursorLocation = adUseClient
    mRS.Open sSQL, mConn, adOpenStatic, adLockReadOnly
    Set mRS.ActiveConnection = Nothing
    
    Set PopulateudtarRpt.coludtUsers01 = New Collection
    
    If mbHideDetails Then
        mRS.Close
        Set mRS = Nothing
        Exit Function
    End If
    
    If Not mRS.EOF Then
        mRS.MoveFirst
        'First need to build the Photo Path. The RS will only contain the
        'Name of the photo. depending upon what DB working with the actual
        'Photo path will differ.
        Do Until mRS.EOF
            'SQL Server the Photo path will be
            With MyUsers
                .f_ContactPhone = vbNullString
                .f_DateLastUpdated = vbNullString
                .f_Email = vbNullString
                .f_FirstName = vbNullString
                .f_LastName = vbNullString
                .f_PassWord = vbNullString
                .f_UserName = vbNullString
                .SecurityLevel = 0
                .UsersID = 0
            End With
            With MyUsers
                On Error Resume Next
                .f_ContactPhone = mRS!ContactPhone
                .f_DateLastUpdated = mRS!DateLastUpdated
                .f_Email = mRS!Email
                .f_FirstName = mRS!FirstName
                .f_LastName = mRS!LastName
                .f_PassWord = mRS!Password
                .f_UserName = mRS!UserName
                .SecurityLevel = mRS!SecurityLevel
                .UsersID = mRS!UsersID
                On Error GoTo EH
            End With
            PopulateudtarRpt.coludtUsers01.Add MyUsers, CStr(MyUsers.UsersID)
            mRS.MoveNext
        Loop
    End If
    mRS.Close
    Set mRS = Nothing
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Set RS = Nothing
    Set mRS = Nothing
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsRptDB_ClassName & vbCrLf & "Private Function PopulateUDTRCAFarmers"
End Function

Private Function GetMiscDelimParamsCol() As Collection
    Set GetMiscDelimParamsCol = Nothing
End Function






