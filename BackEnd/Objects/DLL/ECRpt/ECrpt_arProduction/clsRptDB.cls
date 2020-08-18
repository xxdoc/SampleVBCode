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
Private marRpt As ECrpt_arProduction.arProduction

Public Type udtClaims01
    AssignmentsID As Long
    f_CLIENTNUM As String
    f_Insured As String
    f_ReceivedDate As Date
    f_ContactDate As Date
    f_InspectedDate As Date
    f_CloseDate As Date
    f_DaysInspToClose As Long
End Type

Private Type udtarRpt
    f_PrintDate As Date
    f_ADJName As String
    f_gndTtlNoOfClaims As Long
    f_gndTtlNoOfContacted As Long
    f_gndTtlNoOfInspected As Long
    f_gndTtlNoOfClosed As Long
    f_gndTtlAvgDaysInspToClose As Long
    f_gndTtlClosedDivByInspected As Double
    f_gndTtlPercentContact As Double
    f_gndTtlInspectedDivByReceived As Double
    f_gndTtlClosedDivByReceived As Double
    f_gndTtlCatName As String
    f_gndTtlSiteLocation As String
    coludtClaims01 As Collection
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

Public Property Let USERSID(psUSERSID As String)
    msUSERSID = psUSERSID
End Property
Public Property Get USERSID() As String
    USERSID = msUSERSID
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
        Set marRpt = New ECrpt_arProduction.arProduction
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
    
    Set marRpt = New ECrpt_arProduction.arProduction
    
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
        marRpt.SetProperty "f_ADJName", .f_ADJName, vbString
        marRpt.SetProperty "f_gndTtlNoOfClaims", .f_gndTtlNoOfClaims, vbLong
        marRpt.SetProperty "f_gndTtlNoOfContacted", .f_gndTtlNoOfContacted, vbLong
        marRpt.SetProperty "f_gndTtlNoOfInspected", .f_gndTtlNoOfInspected, vbLong
        marRpt.SetProperty "f_gndTtlNoOfClosed", .f_gndTtlNoOfClosed, vbLong
        marRpt.SetProperty "f_gndTtlAvgDaysInspToClose", .f_gndTtlAvgDaysInspToClose, vbLong
        marRpt.SetProperty "f_gndTtlClosedDivByInspected", .f_gndTtlClosedDivByInspected, vbDouble
        marRpt.SetProperty "f_gndTtlPercentContact", .f_gndTtlPercentContact, vbDouble
        marRpt.SetProperty "f_gndTtlInspectedDivByReceived", .f_gndTtlInspectedDivByReceived, vbDouble
        marRpt.SetProperty "f_gndTtlClosedDivByReceived", .f_gndTtlClosedDivByReceived, vbDouble
        marRpt.SetProperty "f_gndTtlCatName", .f_gndTtlCatName, vbString
        marRpt.SetProperty "f_gndTtlSiteLocation", .f_gndTtlSiteLocation, vbString
        marRpt.SetProperty "coludtClaims01", .coludtClaims01, vbObject
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
    Dim MyClaims As udtClaims01
    
    'gnd Total Vars
    Dim lgndTtlNoOfClaims As Long
    Dim lgndTtlNoOfContacted As Long
    Dim lgndTtlNoOfInspected As Long
    Dim lgndTtlNoOfClosed As Long
    Dim lgndTtlDaysInspToClose As Long
    Dim dgndTtlAvgDaysInspToClose As Double
    Dim dgndTtlClosedDivByInspected As Double
    Dim dgndTtlPercentContact As Double
    Dim dgndTtlInspectedDivByReceived As Double
    Dim dgndTtlClosedDivByReceived As Double
    
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
    
    'B.  get the Production Report Info
    sSQL = "SELECT "
    sSQL = sSQL & "( "
    sSQL = sSQL & "SELECT   LastName + ', ' + FirstName "
    sSQL = sSQL & "FROM     USERS "
    sSQL = sSQL & "WHERE    USERSID = " & msUSERSID & " "
    sSQL = sSQL & ") As ADJName, "
    sSQL = sSQL & "COUNT(A.AssignmentsID) As gndTtlNoOfClaims, "
    sSQL = sSQL & "COUNT(A.ContactDate) As gndTtlNoOfContacted, "
    sSQL = sSQL & "COUNT(A.InspectedDate) As gndTtlNoOfInspected, "
    sSQL = sSQL & "COUNT(A.CloseDate) As gndTtlNoOfClosed, "
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
    sSQL = sSQL & "FROM Assignments A "
    sSQL = sSQL & "WHERE A.ClientCompanyCatSpecID IN "
                                        sSQL = sSQL & "( "
                                        sSQL = sSQL & "SELECT   ClientCompanyCatSpecID "
                                        sSQL = sSQL & "FROM     ClientCompanyCatSpec "
                                        sSQL = sSQL & "WHERE    ClientCompanyID = " & msClientCompanyID & " "
                                        sSQL = sSQL & "AND      CATID = " & msCATID & " "
                                        sSQL = sSQL & ") "
    If msUSERSID <> "-1" Then
        sSQL = sSQL & "AND A.AdjusterSpecID IN "
                                            sSQL = sSQL & "( "
                                            sSQL = sSQL & "SELECT   ClientCoAdjusterSpecID "
                                            sSQL = sSQL & "FROM     ClientCoAdjusterSpec "
                                            sSQL = sSQL & "WHERE    USERSID = " & msUSERSID & " "
                                            sSQL = sSQL & ") "
    End If
    sSQL = sSQL & "AND A.IsDeleted = 0 "
    sSQL = sSQL & "AND (A.Reassigned = 0 Or A.Reassigned Is Null) "
    sSQL = sSQL & "AND A.CLIENTNUM Not Like '%TEST%' "
    
    Set mRS = New ADODB.Recordset
    mRS.CursorLocation = adUseClient
    mRS.Open sSQL, mConn, adOpenStatic, adLockReadOnly
    Set mRS.ActiveConnection = Nothing
    
    With PopulateudtarRpt
            .f_gndTtlAvgDaysInspToClose = 0
            .f_PrintDate = NULL_DATE
            .f_ADJName = vbNullString
            .f_gndTtlNoOfClaims = 0
            .f_gndTtlNoOfContacted = 0
            .f_gndTtlNoOfInspected = 0
            .f_gndTtlNoOfClosed = 0
            .f_gndTtlClosedDivByInspected = 0
            .f_gndTtlPercentContact = 0
            .f_gndTtlInspectedDivByReceived = 0
            .f_gndTtlClosedDivByReceived = 0
            .f_gndTtlCatName = vbNullString
            .f_gndTtlSiteLocation = vbNullString
            On Error Resume Next
            lgndTtlNoOfClaims = mRS!gndTtlNoOfClaims
            lgndTtlNoOfContacted = mRS!gndTtlNoOfContacted
            lgndTtlNoOfInspected = mRS!gndTtlNoOfInspected
            lgndTtlNoOfClosed = mRS!gndTtlNoOfClosed
            On Error GoTo EH
            'Calc the percentages
            'check for div by zero!
            
            '1.gndTtlClosedDivByInspected
            If lgndTtlNoOfInspected > 0 Then
                dgndTtlClosedDivByInspected = (lgndTtlNoOfClosed / lgndTtlNoOfInspected)
            Else
                dgndTtlClosedDivByInspected = 0
            End If
            
            '2. gndTtlPercentContact
            If lgndTtlNoOfClaims > 0 Then
                dgndTtlPercentContact = (lgndTtlNoOfContacted / lgndTtlNoOfClaims)
            Else
                dgndTtlPercentContact = 0
            End If
                            
            '3.gndTtlInspectedDivByReceived use Claims
            If lgndTtlNoOfClaims > 0 Then
                dgndTtlInspectedDivByReceived = (lgndTtlNoOfInspected / lgndTtlNoOfClaims)
            Else
                dgndTtlInspectedDivByReceived = 0
            End If
            
            '4. gndTtlClosedDivByReceived use claims
             If lgndTtlNoOfClaims > 0 Then
                dgndTtlClosedDivByReceived = (lgndTtlNoOfClosed / lgndTtlNoOfClaims)
            Else
                dgndTtlClosedDivByReceived = 0
            End If
            
            On Error Resume Next
            .f_PrintDate = Now()
            .f_ADJName = mRS!ADJName
            .f_gndTtlNoOfClaims = lgndTtlNoOfClaims
            .f_gndTtlNoOfContacted = lgndTtlNoOfContacted
            .f_gndTtlNoOfInspected = lgndTtlNoOfInspected
            .f_gndTtlNoOfClosed = lgndTtlNoOfClosed
            .f_gndTtlClosedDivByInspected = dgndTtlClosedDivByInspected
            .f_gndTtlPercentContact = dgndTtlPercentContact
            .f_gndTtlInspectedDivByReceived = dgndTtlInspectedDivByReceived
            .f_gndTtlClosedDivByReceived = dgndTtlClosedDivByReceived
            .f_gndTtlCatName = mRS!gndTtlCatName
            .f_gndTtlSiteLocation = mRS!gndTtlSiteLocation
            On Error GoTo EH
        End With
    
    sSQL = "SELECT "
    sSQL = sSQL & "A.AssignmentsID, "
    sSQL = sSQL & "A.CLIENTNUM, "
    sSQL = sSQL & "A.Insured, "
    sSQL = sSQL & "A.ReceivedDate, "
    sSQL = sSQL & "A.ContactDate, "
    sSQL = sSQL & "A.InspectedDate, "
    sSQL = sSQL & "A.CloseDate, "
    If goUtil.gMainDB Is Nothing Then
        '------------------------------------BEGIN SQL SERVER DB-------------------------
        '
        '--------------------------------------------------------------------------------
         sSQL = sSQL & "DATEDIFF(Day, A.InspectedDate,A.CloseDate) As DaysInspToClose "
        '------------------------------------END SQL SERVER DB---------------------------
        '
        '--------------------------------------------------------------------------------
    Else
        '------------------------------------BEGIN Access DB----------------------------
        '
        '-------------------------------------------------------------------------------
         sSQL = sSQL & "DATEDIFF('d', A.InspectedDate,A.CloseDate) As DaysInspToClose "
        '------------------------------------END Access DB------------------------------
        '
        '-------------------------------------------------------------------------------
    End If
   
    sSQL = sSQL & "FROM Assignments A "
    sSQL = sSQL & "WHERE A.ClientCompanyCatSpecID IN "
                                        sSQL = sSQL & "( "
                                        sSQL = sSQL & "SELECT   ClientCompanyCatSpecID "
                                        sSQL = sSQL & "FROM     ClientCompanyCatSpec "
                                        sSQL = sSQL & "WHERE    ClientCompanyID = " & msClientCompanyID & " "
                                        sSQL = sSQL & "AND      CATID = " & msCATID & " "
                                        sSQL = sSQL & ") "
    If msUSERSID <> "-1" Then
        sSQL = sSQL & "AND A.AdjusterSpecID IN "
                                            sSQL = sSQL & "( "
                                            sSQL = sSQL & "SELECT   ClientCoAdjusterSpecID "
                                            sSQL = sSQL & "FROM     ClientCoAdjusterSpec "
                                            sSQL = sSQL & "WHERE    USERSID = " & msUSERSID & " "
                                            sSQL = sSQL & ") "
    End If
    sSQL = sSQL & "AND A.IsDeleted = 0 "
    sSQL = sSQL & "AND (A.Reassigned = 0 Or A.Reassigned Is Null) "
    sSQL = sSQL & "AND A.CLIENTNUM Not Like '%TEST%' "
    sSQL = sSQL & "ORDER BY A.CLIENTNUM "
    
    Set mRS = New ADODB.Recordset
    mRS.CursorLocation = adUseClient
    mRS.Open sSQL, mConn, adOpenStatic, adLockReadOnly
    Set mRS.ActiveConnection = Nothing
    
    Set PopulateudtarRpt.coludtClaims01 = New Collection
    
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
            With MyClaims
                .AssignmentsID = 0
                .f_CLIENTNUM = vbNullString
                .f_Insured = vbNullString
                .f_ReceivedDate = NULL_DATE
                .f_ContactDate = NULL_DATE
                .f_InspectedDate = NULL_DATE
                .f_CloseDate = NULL_DATE
                .f_DaysInspToClose = 0
            End With
            With MyClaims
                On Error Resume Next
                If Not IsNull(mRS!InspectedDate) And Not IsNull(mRS!CloseDate) Then
                    If mRS!InspectedDate <> NULL_DATE And mRS!CloseDate <> NULL_DATE Then
                        lgndTtlDaysInspToClose = lgndTtlDaysInspToClose + mRS!DaysInspToClose
                    End If
                End If
                .AssignmentsID = mRS!AssignmentsID
                .f_CLIENTNUM = mRS!CLIENTNUM
                .f_Insured = mRS!Insured
                .f_ReceivedDate = mRS!ReceivedDate
                .f_ContactDate = mRS!ContactDate
                .f_InspectedDate = mRS!InspectedDate
                .f_CloseDate = mRS!CloseDate
                If Not IsNull(mRS!InspectedDate) And Not IsNull(mRS!CloseDate) Then
                    If mRS!InspectedDate <> NULL_DATE And mRS!CloseDate <> NULL_DATE Then
                        .f_DaysInspToClose = mRS!DaysInspToClose
                    End If
                End If
                On Error GoTo EH
            End With
            PopulateudtarRpt.coludtClaims01.Add MyClaims, CStr(MyClaims.AssignmentsID)
            mRS.MoveNext
        Loop
        'Figure out the average days to close
        '1.gndTtlAvgDaysInspToClose
        
        If lgndTtlNoOfClaims > 0 Then
            dgndTtlAvgDaysInspToClose = (lgndTtlDaysInspToClose / lgndTtlNoOfClaims)
        Else
            dgndTtlAvgDaysInspToClose = 0
        End If
        PopulateudtarRpt.f_gndTtlAvgDaysInspToClose = dgndTtlAvgDaysInspToClose
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
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim sSQL As String
    Dim oPreviousParamsCol As Collection
    Dim oParamsCol As Collection
    Dim vParams As Variant
    Dim MyParams As V2ECKeyBoard.MiscReportParam
    Dim RS As ADODB.Recordset
    
    'Active Rport Field Object
    Dim oField As Object
    Dim sTag As String
    Dim saryTag() As String
    Dim sParamCaption As String
    Dim sParamName As String
    Dim sSortMe As String
    Dim sDataType As String
    Dim iDataType As VBA.VbVarType
    Dim sDefaultValue As String
    
    '2/8/2004 MiscReportParam , and MiscReportParam01 to MiscReportParam30
    'This is the only place the name of the Misc Param table should be set
    'Once set this report MUST always use the same MiscParam table.
    'If a parameter table must change then need to create another version
    'of this report and change the name there!!!
    Dim sMiscReportParamName As String
    sMiscReportParamName = "MiscReportParam"
    '2/8/2004 MiscReportParam , and MiscReportParam01 to MiscReportParam30
    
    If msAssignmentsID = vbNullString Or StrComp(msAssignmentsID, "ALL", vbTextCompare) = 0 Then
        Exit Function
    End If
    
   'See if we are not using using Adjuster DB need
    'to open connection to WEB DB
    If goUtil.gMainDB Is Nothing Then
        clsRptDB_OpenConnection
    Else
        'Open Connection to Access DB
        Set mConn = New ADODB.Connection
        goUtil.utOpenDatabaseADOConn mConn, goUtil.gMainDB.Name
    End If
    
    'SQL Server
    sSQL = "SELECT  "
    sSQL = sSQL & "[MiscReportParamID], "
    sSQL = sSQL & "[AssignmentsID], "
    sSQL = sSQL & "[ID], "
    sSQL = sSQL & "[IDAssignments], "
    sSQL = sSQL & "[Number], "
    sSQL = sSQL & "[ProjectName], "
    sSQL = sSQL & "[ClassName], "
    sSQL = sSQL & "[ParamName], "
    sSQL = sSQL & "[ParamCaption], "
    sSQL = sSQL & "[ParamDataType], "
    sSQL = sSQL & "[ParamValue], "
    sSQL = sSQL & "[SortMe], "
    sSQL = sSQL & "[IsDeleted], "
    sSQL = sSQL & "[DownLoadMe], "
    sSQL = sSQL & "[UpLoadMe], "
    sSQL = sSQL & "[AdminComments], "
    sSQL = sSQL & "[DateLastUpdated], "
    sSQL = sSQL & "[UpdateByUserID] "
    '2/8/2004 MiscReportParam , and MiscReportParam01 to MiscReportParam30
    sSQL = sSQL & "FROM    " & sMiscReportParamName & " "
    '2/8/2004 MiscReportParam , and MiscReportParam01 to MiscReportParam30
    sSQL = sSQL & "WHERE   [AssignmentsID] = " & msAssignmentsID & " "
    sSQL = sSQL & "AND ClassName = '" & goUtil.utCleanSQLString(clsRptDB_ClassName) & "' "
    sSQL = sSQL & "AND IsDeleted = 0 "
    Set RS = New ADODB.Recordset
    RS.CursorLocation = adUseClient
    RS.Open sSQL, mConn, adOpenStatic, adLockReadOnly
    Set RS.ActiveConnection = Nothing
    
    If RS.RecordCount > 0 Then
        RS.MoveFirst
        Do Until RS.EOF
            With MyParams
                .MiscReportParamID = goUtil.IsNullIsVbNullString(RS.Fields("MiscReportParamID"))
                .AssignmentsID = goUtil.IsNullIsVbNullString(RS.Fields("AssignmentsID"))
                .ID = goUtil.IsNullIsVbNullString(RS.Fields("ID"))
                .IDAssignments = goUtil.IsNullIsVbNullString(RS.Fields("IDAssignments"))
                .Number = goUtil.IsNullIsVbNullString(RS.Fields("Number"))
                .ProjectName = goUtil.IsNullIsVbNullString(RS.Fields("ProjectName"))
                .ClassName = goUtil.IsNullIsVbNullString(RS.Fields("ClassName"))
                .ParamName = goUtil.IsNullIsVbNullString(RS.Fields("ParamName"))
                .ParamCaption = goUtil.IsNullIsVbNullString(RS.Fields("ParamCaption"))
                .ParamDataType = goUtil.IsNullIsVbNullString(RS.Fields("ParamDataType"))
                .ParamValue = goUtil.IsNullIsVbNullString(RS.Fields("ParamValue"))
                .SortMe = goUtil.IsNullIsVbNullString(RS.Fields("SortMe"))
                .IsDeleted = goUtil.IsNullIsVbNullString(RS.Fields("IsDeleted"))
                .DownLoadMe = goUtil.IsNullIsVbNullString(RS.Fields("DownLoadMe"))
                .UpLoadMe = goUtil.IsNullIsVbNullString(RS.Fields("UpLoadMe"))
                .AdminComments = goUtil.IsNullIsVbNullString(RS.Fields("AdminComments"))
                .DateLastUpdated = goUtil.IsNullIsVbNullString(RS.Fields("DateLastUpdated"))
                .UpdateByUserID = goUtil.IsNullIsVbNullString(RS.Fields("UpdateByUserID"))
            End With
            If oPreviousParamsCol Is Nothing Then
                Set oPreviousParamsCol = New Collection
            End If
            On Error Resume Next
            oPreviousParamsCol.Add MyParams, MyParams.ParamName
            If Err.Number > 0 Then
                Err.Clear
            End If
            On Error GoTo EH
            RS.MoveNext
        Loop
    End If
    
    
    Set oParamsCol = New Collection
        
    'For this Particular Report since there are so Many Params
    'Need to Loop Through Field Objects with Params in the Tag Property.
    For Each oField In marRpt.Detail.Controls
        If TypeOf oField Is DDActiveReports.Field Then
            sTag = oField.Tag
            If Left(sTag, 1) = "p" Then
                saryTag() = Split(sTag, "_")
                sParamName = oField.Name
                sSortMe = saryTag(0)
                sParamCaption = saryTag(1)
                sDataType = saryTag(2)
                sDefaultValue = saryTag(3)
                iDataType = goUtil.GetvbVarTypeFromText(sDataType)
                With MyParams
                    .MiscReportParamID = vbNullString
                    .AssignmentsID = msAssignmentsID
                    .ID = vbNullString
                    .IDAssignments = msAssignmentsID
                    .Number = vbNullString
                    .ClassName = clsRptDB_ClassName
                    .ParamCaption = sParamCaption
                    .ParamDataType = iDataType
                    .ParamName = sParamName
                    .ParamValue = sDefaultValue
                    .SortMe = sSortMe
                    .ProjectName = App.EXEName
                    .IsDeleted = 0
                    .DownLoadMe = 0
                    .UpLoadMe = "True"
                    .AdminComments = vbNullString
                    .DateLastUpdated = Now()
                    .UpdateByUserID = goUtil.gsCurUsersID
                End With
                oParamsCol.Add MyParams, MyParams.ParamName
            End If
        End If
    Next
    
    If Not oPreviousParamsCol Is Nothing And Not oParamsCol Is Nothing Then
        For Each vParams In oPreviousParamsCol
            MyParams = vParams
            RemoveParam MyParams.ParamName, oParamsCol
            oParamsCol.Add MyParams, MyParams.ParamName
        Next
    End If
    
    '2/8/2004 MiscReportParam , and MiscReportParam01 to MiscReportParam30
    'Add the name used to the collection, so the calle knows what TableName...
    'MiscReportParam or MiscReportParam01 to MiscReportParam30 to use if needed.
    If Not oParamsCol Is Nothing Then
        oParamsCol.Add sMiscReportParamName, "sMiscReportParamName"
    End If
    '2/8/2004 MiscReportParam , and MiscReportParam01 to MiscReportParam30
    
    Set GetMiscDelimParamsCol = oParamsCol
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Set GetMiscDelimParamsCol = Nothing
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsRptDB_ClassName & vbCrLf & "Private Function GetMiscDelimParamsCol"
End Function






