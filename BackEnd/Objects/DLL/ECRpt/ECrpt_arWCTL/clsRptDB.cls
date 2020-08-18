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
Private marRpt As ECrpt_arWCTL.arWCTL

Public Type udtWCTL01
    AssignmentsID As Long
    f_ClaimNumber As String
    f_Status As String
    f_Mon As Double
    f_Tue As Double
    f_Wed As Double
    f_Thu As Double
    f_Fri As Double
    f_Sat As Double
    f_Sun As Double
    f_ClaimTtl As Double
End Type

Private Type udtarRpt
    f_PrintDate As String
    f_ADJName As String
    f_FromDate As Date
    f_ToDate As Date
    f_NewClaimsThisWeek As Long
    f_ClosedClaimsThisWeek As Long
    f_TtlOpenInventroy As Long
    f_HeadMon As Double
    f_HeadTue As Double
    f_HeadWed As Double
    f_HeadThu As Double
    f_HeadFri As Double
    f_HeadSat As Double
    f_HeadSun As Double
    f_HeadWeekTtl As Double
    f_FootMon As Double
    f_FootTue As Double
    f_FootWed As Double
    f_FootThu As Double
    f_FootFri As Double
    f_FootSat As Double
    f_FootSun As Double
    f_FootWeekTtl As Double
    f_gndTtlCatName As String
    f_gndTtlSiteLocation As String
    coludtWCTL01 As Collection
End Type

Private msClientCompanyID As String
Private msCATID As String
Private msUSERSID As String
Private mdtFromSundayDate As Date
Private mbHideDetails As Boolean
Private mbExportXML As Boolean
Private mbExportXMLOnly As Boolean

Public Property Let HideDetails(pbFlag As Boolean)
    mbHideDetails = pbFlag
End Property
Public Property Get HideDetails() As Boolean
    HideDetails = mbHideDetails
End Property

Public Property Let FromSundayDate(pdtFromSundayDate As Date)
    'Need to Be sure this date is a Sunday.
    'If it is Not then find the Sunday date previous to the date passed in!
    If Weekday(pdtFromSundayDate) = VbDayOfWeek.vbSunday Then
        mdtFromSundayDate = pdtFromSundayDate
    Else
        mdtFromSundayDate = DateAdd("d", (Weekday(pdtFromSundayDate) * -1) + 1, pdtFromSundayDate)
    End If
End Property
Public Property Get FromSundayDate() As Date
    FromSundayDate = mdtFromSundayDate
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
        Set marRpt = New ECrpt_arWCTL.arWCTL
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
    
    Set marRpt = New ECrpt_arWCTL.arWCTL
    
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
        marRpt.SetProperty "f_PrintDate", .f_PrintDate, vbString
        marRpt.SetProperty "f_ADJName", .f_ADJName, vbString
        marRpt.SetProperty "f_FromDate", .f_FromDate, vbDate
        marRpt.SetProperty "f_ToDate", .f_ToDate, vbDate
        marRpt.SetProperty "f_NewClaimsThisWeek", .f_NewClaimsThisWeek, vbLong
        marRpt.SetProperty "f_ClosedClaimsThisWeek", .f_ClosedClaimsThisWeek, vbLong
        marRpt.SetProperty "f_TtlOpenInventroy", .f_TtlOpenInventroy, vbLong
        marRpt.SetProperty "f_HeadMon", .f_HeadMon, vbDouble
        marRpt.SetProperty "f_HeadTue", .f_HeadTue, vbDouble
        marRpt.SetProperty "f_HeadWed", .f_HeadWed, vbDouble
        marRpt.SetProperty "f_HeadThu", .f_HeadThu, vbDouble
        marRpt.SetProperty "f_HeadFri", .f_HeadFri, vbDouble
        marRpt.SetProperty "f_HeadSat", .f_HeadSat, vbDouble
        marRpt.SetProperty "f_HeadSun", .f_HeadSun, vbDouble
        marRpt.SetProperty "f_HeadWeekTtl", .f_HeadWeekTtl, vbDouble
        marRpt.SetProperty "f_FootMon", .f_FootMon, vbDouble
        marRpt.SetProperty "f_FootTue", .f_FootTue, vbDouble
        marRpt.SetProperty "f_FootWed", .f_FootWed, vbDouble
        marRpt.SetProperty "f_FootThu", .f_FootThu, vbDouble
        marRpt.SetProperty "f_FootFri", .f_FootFri, vbDouble
        marRpt.SetProperty "f_FootSat", .f_FootSat, vbDouble
        marRpt.SetProperty "f_FootSun", .f_FootSun, vbDouble
        marRpt.SetProperty "f_FootWeekTtl", .f_FootWeekTtl, vbDouble
        marRpt.SetProperty "f_gndTtlCatName", .f_gndTtlCatName, vbString
        marRpt.SetProperty "f_gndTtlSiteLocation", .f_gndTtlSiteLocation, vbString
        marRpt.SetProperty "coludtWCTL01", .coludtWCTL01, vbObject
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
    Dim sSQL As String
    Dim sSQLWhere As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim MyWCTL As udtWCTL01
    
    Dim dtDay As Date
    Dim lAssignmentsID As Long
    Dim sStatus As String
    Dim sCLIENTNUM As String
    Dim bLastRecord As Boolean
    'Weekday Vars
    Dim dMON As Double
    Dim dTUE As Double
    Dim dWED As Double
    Dim dTHU As Double
    Dim dFRI As Double
    Dim dSAT As Double
    Dim dSUN As Double
    Dim dTtlWeek As Double
    'TTL WEEKDAY Vars
    Dim dMONgndTtl As Double
    Dim dTUEgndTtl As Double
    Dim dWEDgndTtl As Double
    Dim dTHUgndTtl As Double
    Dim dFRIgndTtl As Double
    Dim dSATgndTtl As Double
    Dim dSUNgndTtl As Double
    Dim dTtlWeekgndTtl As Double
    
    
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
    
    'Build the Main Where Clause
    sSQL = sSQL & "WHERE ClientCompanyCatSpecID IN "
                                        sSQL = sSQL & "( "
                                        sSQL = sSQL & "SELECT   ClientCompanyCatSpecID "
                                        sSQL = sSQL & "FROM     ClientCompanyCatSpec "
                                        sSQL = sSQL & "WHERE    ClientCompanyID = " & msClientCompanyID & " "
                                        sSQL = sSQL & "AND      CATID = " & msCATID & " "
                                        sSQL = sSQL & ") "
    sSQL = sSQL & "AND AdjusterSpecID IN "
                                        sSQL = sSQL & "( "
                                        sSQL = sSQL & "SELECT   ClientCoAdjusterSpecID "
                                        sSQL = sSQL & "FROM     ClientCoAdjusterSpec "
                                        sSQL = sSQL & "WHERE    USERSID = " & msUSERSID & " "
                                        sSQL = sSQL & ") "
    sSQL = sSQL & "AND IsDeleted = 0 "
    sSQL = sSQL & "AND (Reassigned = 0 Or Reassigned Is Null) "
    sSQL = sSQL & "AND CLIENTNUM Not Like '%TEST%' "
    sSQLWhere = sSQL


    'First get the WCTL Header Info (Does Not include Header/Footer totals YET!)
    sSQL = "SELECT "
    sSQL = sSQL & "( "
    sSQL = sSQL & "U.LastName + ', ' + U.FirstName "
    sSQL = sSQL & ") As ADJName, "
    'NewClaimsThisWeek
    sSQL = sSQL & "( "
        sSQL = sSQL & "SELECT   COUNT(AssignmentsID) As CountOFAssignmentsID "
        sSQL = sSQL & "FROM     Assignments "
        sSQL = sSQL & sSQLWhere
        sSQL = sSQL & "AND StatusID = " & AssgnStatus.iAssignmentsStatus_NEW & " "
        If goUtil.gMainDB Is Nothing Then
            '------------------------------------BEGIN SQL SERVER DB-------------------------
            '
            '--------------------------------------------------------------------------------
            sSQL = sSQL & "AND AssignedDate >= Cast('" & mdtFromSundayDate & "' As DateTime) "
            sSQL = sSQL & "AND AssignedDate <= Cast('" & DateAdd("d", 6, mdtFromSundayDate) & "' As DateTime) "
            '------------------------------------END SQL SERVER DB---------------------------
            '
            '--------------------------------------------------------------------------------
        Else
            '------------------------------------BEGIN Access DB----------------------------
            '
            '-------------------------------------------------------------------------------
            sSQL = sSQL & "AND AssignedDate >= #" & mdtFromSundayDate & "# "
            sSQL = sSQL & "AND AssignedDate <= #" & DateAdd("d", 6, mdtFromSundayDate) & "# "
            '------------------------------------END Access DB------------------------------
            '
            '-------------------------------------------------------------------------------
        End If
    sSQL = sSQL & ") As NewClaimsThisWeek, "
    'ClosedClaimsThisWeek
    sSQL = sSQL & "( "
        sSQL = sSQL & "SELECT   COUNT(AssignmentsID) As CountOFAssignmentsID "
        sSQL = sSQL & "FROM     Assignments "
        sSQL = sSQL & sSQLWhere
        sSQL = sSQL & "AND StatusID = " & AssgnStatus.iAssignmentsStatus_CLOSED & " "
        If goUtil.gMainDB Is Nothing Then
            '------------------------------------BEGIN SQL SERVER DB-------------------------
            '
            '--------------------------------------------------------------------------------
            sSQL = sSQL & "AND CloseDate >= Cast('" & mdtFromSundayDate & "' As DateTime) "
            sSQL = sSQL & "AND CloseDate <= Cast('" & DateAdd("d", 6, mdtFromSundayDate) & "' As DateTime) "
            '------------------------------------END SQL SERVER DB---------------------------
            '
            '--------------------------------------------------------------------------------
        Else
            '------------------------------------BEGIN Access DB----------------------------
            '
            '-------------------------------------------------------------------------------
            sSQL = sSQL & "AND CloseDate >= #" & mdtFromSundayDate & "# "
            sSQL = sSQL & "AND CloseDate <= #" & DateAdd("d", 6, mdtFromSundayDate) & "# "
            '------------------------------------END Access DB------------------------------
            '
            '-------------------------------------------------------------------------------
        End If
    sSQL = sSQL & ") As ClosedClaimsThisWeek, "
    'TtlOpenInventroy
    sSQL = sSQL & "( "
        sSQL = sSQL & "SELECT   COUNT(AssignmentsID) As CountOFAssignmentsID "
        sSQL = sSQL & "FROM     Assignments "
        sSQL = sSQL & sSQLWhere
        sSQL = sSQL & "AND StatusID <> " & AssgnStatus.iAssignmentsStatus_CLOSED & " "
    sSQL = sSQL & ") As TtlOpenInventroy, "
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
    sSQL = sSQL & "FROM     USERS U "
    sSQL = sSQL & "WHERE    U.USERSID = " & msUSERSID & " "
    
    
    Set mRS = New ADODB.Recordset
    mRS.CursorLocation = adUseClient
    mRS.Open sSQL, mConn, adOpenStatic, adLockReadOnly
    Set mRS.ActiveConnection = Nothing
    With PopulateudtarRpt
        .f_PrintDate = vbNullString
        .f_ADJName = vbNullString
        .f_FromDate = NULL_DATE
        .f_ToDate = NULL_DATE
        .f_NewClaimsThisWeek = 0
        .f_ClosedClaimsThisWeek = 0
        .f_TtlOpenInventroy = 0
        .f_gndTtlCatName = vbNullString
        .f_gndTtlSiteLocation = vbNullString
        On Error Resume Next
        .f_PrintDate = Now()
        .f_ADJName = mRS!ADJName
        .f_FromDate = mdtFromSundayDate
        .f_ToDate = DateAdd("d", 6, mdtFromSundayDate)
        .f_NewClaimsThisWeek = mRS!NewClaimsThisWeek
        .f_ClosedClaimsThisWeek = mRS!ClosedClaimsThisWeek
        .f_TtlOpenInventroy = mRS!TtlOpenInventroy
        .f_gndTtlCatName = mRS!gndTtlCatName
        .f_gndTtlSiteLocation = mRS!gndTtlSiteLocation
        On Error GoTo EH
   End With
   
    'Next Get List of Act Log entries and parse out the Days set them to vars
    sSQL = "SELECT "
    sSQL = sSQL & "A.AssignmentsID, "
    sSQL = sSQL & "A.CLIENTNUM, "
    sSQL = sSQL & "( "
    sSQL = sSQL & "SELECT  StatusAlias As Status "
    sSQL = sSQL & "FROM     Status "
    sSQL = sSQL & "WHERE    StatusID = A.[StatusID] "
    sSQL = sSQL & ") As Status, "
    sSQL = sSQL & "RTAL.ServiceTime, "
    sSQL = sSQL & "RTAL.ActDate "
    sSQL = sSQL & "FROM Assignments A "
    sSQL = sSQL & "INNER JOIN RTActivityLog RTAL ON A.AssignmentsID = RTAL.AssignmentsID "
    sSQL = sSQL & "WHERE A.ClientCompanyCatSpecID IN "
                                        sSQL = sSQL & "( "
                                        sSQL = sSQL & "SELECT   ClientCompanyCatSpecID "
                                        sSQL = sSQL & "FROM     ClientCompanyCatSpec "
                                        sSQL = sSQL & "WHERE    ClientCompanyID = " & msClientCompanyID & " "
                                        sSQL = sSQL & "AND      CATID = " & msCATID & " "
                                        sSQL = sSQL & ") "
    sSQL = sSQL & "AND A.AdjusterSpecID IN "
                                        sSQL = sSQL & "( "
                                        sSQL = sSQL & "SELECT   ClientCoAdjusterSpecID "
                                        sSQL = sSQL & "FROM     ClientCoAdjusterSpec "
                                        sSQL = sSQL & "WHERE    USERSID = " & msUSERSID & " "
                                        sSQL = sSQL & ") "
    sSQL = sSQL & "AND RTAL.IsMgrEntry = 0 "
    sSQL = sSQL & "AND RTAL.IsDeleted = 0 "
    sSQL = sSQL & "AND A.IsDeleted = 0 "
    If goUtil.gMainDB Is Nothing Then
            '------------------------------------BEGIN SQL SERVER DB-------------------------
            '
            '--------------------------------------------------------------------------------
            sSQL = sSQL & "AND RTAL.ActDate >= Cast('" & mdtFromSundayDate & "' As DateTime) "
            sSQL = sSQL & "AND RTAL.ActDate <= Cast('" & DateAdd("d", 6, mdtFromSundayDate) & "' As DateTime) "
            '------------------------------------END SQL SERVER DB---------------------------
            '
            '--------------------------------------------------------------------------------
        Else
            '------------------------------------BEGIN Access DB----------------------------
            '
            '-------------------------------------------------------------------------------
            sSQL = sSQL & "AND RTAL.ActDate >= #" & mdtFromSundayDate & "# "
            sSQL = sSQL & "AND RTAL.ActDate <= #" & DateAdd("d", 6, mdtFromSundayDate) & "# "
            '------------------------------------END Access DB------------------------------
            '
            '-------------------------------------------------------------------------------
        End If
    sSQL = sSQL & "ORDER BY A.CLIENTNUM "
    
    Set mRS = New ADODB.Recordset
    mRS.CursorLocation = adUseClient
    mRS.Open sSQL, mConn, adOpenStatic, adLockReadOnly
    Set mRS.ActiveConnection = Nothing
    
    Set PopulateudtarRpt.coludtWCTL01 = New Collection
    
    If Not mRS.EOF Then
        mRS.MoveFirst
        lAssignmentsID = 0
        Do Until mRS.EOF And bLastRecord
            bLastRecord = mRS.EOF
            If bLastRecord Then
                GoTo ENTER_LAST_RECORD
            End If
            If mRS!AssignmentsID <> lAssignmentsID And lAssignmentsID <> 0 Then
ENTER_LAST_RECORD:
                'Add WCTL Item to collection
                With MyWCTL
                    .AssignmentsID = lAssignmentsID
                    .f_ClaimNumber = sCLIENTNUM
                    .f_Status = sStatus
                    .f_Mon = dMON
                    .f_Tue = dTUE
                    .f_Wed = dWED
                    .f_Thu = dTHU
                    .f_Fri = dFRI
                    .f_Sat = dSAT
                    .f_Sun = dSUN
                    .f_ClaimTtl = dTtlWeek
                End With
                
                If Not mbHideDetails Then
                    PopulateudtarRpt.coludtWCTL01.Add MyWCTL, CStr(MyWCTL.AssignmentsID)
                End If
                
                'Update Grand Totals
                dMONgndTtl = dMONgndTtl + dMON
                dTUEgndTtl = dTUEgndTtl + dTUE
                dWEDgndTtl = dWEDgndTtl + dWED
                dTHUgndTtl = dTHUgndTtl + dTHU
                dFRIgndTtl = dFRIgndTtl + dFRI
                dSATgndTtl = dSATgndTtl + dSAT
                dSUNgndTtl = dSUNgndTtl + dSUN
                dTtlWeekgndTtl = dTtlWeekgndTtl + dTtlWeek
                
                'Reset for next claim
                dtDay = NULL_DATE
                dMON = 0
                dTUE = 0
                dWED = 0
                dTHU = 0
                dFRI = 0
                dSAT = 0
                dSUN = 0
                dTtlWeek = 0
                
                'Check for Last Record
                If bLastRecord Then
                    Exit Do
                End If
            End If
            lAssignmentsID = mRS!AssignmentsID
            sStatus = mRS!Status
            sCLIENTNUM = mRS!CLIENTNUM
            dtDay = IIf(IsNull(mRS!ActDate), NULL_DATE, mRS!ActDate)
            If dtDay <> NULL_DATE Then
                Select Case Weekday(dtDay)
                    Case VbDayOfWeek.vbMonday
                        dMON = dMON + IIf(IsNull(mRS!ServiceTime), 0, mRS!ServiceTime)
                    Case VbDayOfWeek.vbTuesday
                        dTUE = dTUE + IIf(IsNull(mRS!ServiceTime), 0, mRS!ServiceTime)
                    Case VbDayOfWeek.vbWednesday
                        dWED = dWED + IIf(IsNull(mRS!ServiceTime), 0, mRS!ServiceTime)
                    Case VbDayOfWeek.vbThursday
                        dTHU = dTHU + IIf(IsNull(mRS!ServiceTime), 0, mRS!ServiceTime)
                    Case VbDayOfWeek.vbFriday
                        dFRI = dFRI + IIf(IsNull(mRS!ServiceTime), 0, mRS!ServiceTime)
                    Case VbDayOfWeek.vbSaturday
                        dSAT = dSAT + IIf(IsNull(mRS!ServiceTime), 0, mRS!ServiceTime)
                    Case VbDayOfWeek.vbSunday
                        dSUN = dSUN + IIf(IsNull(mRS!ServiceTime), 0, mRS!ServiceTime)
                End Select
                dTtlWeek = dTtlWeek + IIf(IsNull(mRS!ServiceTime), 0, mRS!ServiceTime)
            End If
            mRS.MoveNext
        Loop
    End If
    mRS.Close
    Set mRS = Nothing
    
    'Finally populate the Header/Footer Totals
     With PopulateudtarRpt
        .f_HeadMon = dMONgndTtl
        .f_HeadTue = dTUEgndTtl
        .f_HeadWed = dWEDgndTtl
        .f_HeadThu = dTHUgndTtl
        .f_HeadFri = dFRIgndTtl
        .f_HeadSat = dSATgndTtl
        .f_HeadSun = dSUNgndTtl
        .f_HeadWeekTtl = dTtlWeekgndTtl
        .f_FootMon = dMONgndTtl
        .f_FootTue = dTUEgndTtl
        .f_FootWed = dWEDgndTtl
        .f_FootThu = dTHUgndTtl
        .f_FootFri = dFRIgndTtl
        .f_FootSat = dSATgndTtl
        .f_FootSun = dSUNgndTtl
        .f_FootWeekTtl = dTtlWeekgndTtl
    End With
    
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



