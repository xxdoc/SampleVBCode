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
Private marRpt As ECrptABC_arIndemnityReport.arIndemnityReport

Public Type udtIndemnity01
    f_RTIndemnityID As Long
    f_Payment As String
    f_IsPreviousPayment As Boolean
    f_ClassOfLoss As String
    f_ClassOfLossCode As String
    f_TypeOfLoss As String
    f_typeOfLossCode As String
    f_Description As String
    f_ReplacementCost As Currency
    f_RecoverableDep As Currency
    f_NonRecoverableDep As Currency
    f_AcvClaim As Currency
    f_SpecialLimits As Currency
    f_IsAddAmountOfInsurance As Boolean
    f_ExcessLimits As Currency
    f_Miscellaneous As Currency
    f_ACVLessExcessLimits As Currency
End Type

Private Type udtarRpt
    f_Insured As String
    f_ClientClaimNo As String
    f_ADJNAME As String
    f_gndTtlCatName As String
    f_gndTtlSiteLocation As String
    f_CloseDate As Date
    f_InspectedDate As Date
    f_RH_FullCostRepairReplace As Currency
    f_RH_RecDep As Currency
    f_RH_NonRecDep As Currency
    f_RH_ACV As Currency
    f_RH_Deductible As Currency
    f_RH_LessExcessLimits As Currency
    f_RH_LessMiscellaneous As Currency
    f_RH_NetACVC As Currency
    coludtIndemnity01 As Collection
End Type

Private mbExportXML As Boolean
Private mbExportXMLOnly As Boolean

Private Property Get clsRptDB_MiscDelimParamsCol() As Collection
    Dim bCleanup As Boolean
    'Add Report Parameters that are not updated by DB Fields.
    If marRpt Is Nothing Then
        bCleanup = True
        Set marRpt = New ECrptABC_arIndemnityReport.arIndemnityReport
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
    Dim MyParams As rptMiscDelimParam
    Dim sXML As String
    Dim sXMLXportPath As String
    Dim lErrNum As Long
    Dim sErrDesc As String

    Set marRpt = New ECrptABC_arIndemnityReport.arIndemnityReport
    
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
        marRpt.SetProperty "f_Insured", .f_Insured, vbString
        marRpt.SetProperty "f_ClientClaimNo", .f_ClientClaimNo, vbString
        marRpt.SetProperty "f_ADJNAME", .f_ADJNAME, vbString
        marRpt.SetProperty "f_gndTtlCatName", .f_gndTtlCatName, vbString
        marRpt.SetProperty "f_gndTtlSiteLocation", .f_gndTtlSiteLocation, vbString
        marRpt.SetProperty "f_CloseDate", .f_CloseDate, vbDate
        marRpt.SetProperty "f_InspectedDate", .f_InspectedDate, vbDate
        marRpt.SetProperty "f_RH_FullCostRepairReplace", .f_RH_FullCostRepairReplace, vbCurrency
        marRpt.SetProperty "f_RH_RecDep", .f_RH_RecDep, vbCurrency
        marRpt.SetProperty "f_RH_NonRecDep", .f_RH_NonRecDep, vbCurrency
        marRpt.SetProperty "f_RH_ACV", .f_RH_ACV, vbCurrency
        marRpt.SetProperty "f_RH_Deductible", .f_RH_Deductible, vbCurrency
        marRpt.SetProperty "f_RH_LessExcessLimits", .f_RH_LessExcessLimits, vbCurrency
        marRpt.SetProperty "f_RH_LessMiscellaneous", .f_RH_LessMiscellaneous, vbCurrency
        marRpt.SetProperty "f_RH_NetACVC", .f_RH_NetACVC, vbCurrency
        marRpt.SetProperty "coludtIndemnity01", .coludtIndemnity01, vbObject
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
        .Printer.Orientation = ddOLandscape 'this one is LAND SCAPE !!!
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
    moLRs.CLEANUP
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
    Dim MyIndemnity As udtIndemnity01

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

    'B.  get the Indemnity Header Info
    sSQL = "SELECT "
    sSQL = sSQL & "A.Insured As Insured, "
    sSQL = sSQL & "A.CLIENTNUM As ClientClaimNo, "
    sSQL = sSQL & "A.CloseDate, "
    sSQL = sSQL & "A.InspectedDate, "
    sSQL = sSQL & "( "
    sSQL = sSQL & "SELECT   LastName + ', ' + FirstName "
    sSQL = sSQL & "FROM     USERS "
    sSQL = sSQL & "WHERE    USERSID = "
                            sSQL = sSQL & "( "
                            sSQL = sSQL & "SELECT   UsersID "
                            sSQL = sSQL & "FROM     ClientCoAdjusterSpec "
                            sSQL = sSQL & "WHERE    ClientCoAdjusterSpecID = A.AdjusterSpecID "
                            sSQL = sSQL & ") "
    sSQL = sSQL & ") As ADJName, "
    sSQL = sSQL & "( "
    sSQL = sSQL & "SELECT   Name + ' - ' + Description "
    sSQL = sSQL & "FROM     CAT "
    sSQL = sSQL & "WHERE CATID = "
                sSQL = sSQL & "( "
                sSQL = sSQL & "SELECT   CATID "
                sSQL = sSQL & "FROM     ClientCompanyCatSpec "
                sSQL = sSQL & "WHERE    ClientCompanyCatSpecID = A.ClientCompanyCatSpecID "
                sSQL = sSQL & ") "
    sSQL = sSQL & ") As gndTtlCatName, "
    sSQL = sSQL & "( "
    sSQL = sSQL & "SELECT   SACity + ', ' + SAState "
    sSQL = sSQL & "FROM     ClientCompanyCat "
    sSQL = sSQL & "WHERE    ClientCompanyID = "
                sSQL = sSQL & "( "
                sSQL = sSQL & "SELECT   ClientCompanyID "
                sSQL = sSQL & "FROM     ClientCompanyCatSpec "
                sSQL = sSQL & "WHERE    ClientCompanyCatSpecID = A.ClientCompanyCatSpecID "
                sSQL = sSQL & ") "
    sSQL = sSQL & "AND  CATID = "
                sSQL = sSQL & "( "
                sSQL = sSQL & "SELECT   CATID "
                sSQL = sSQL & "FROM     ClientCompanyCatSpec "
                sSQL = sSQL & "WHERE    ClientCompanyCatSpecID = A.ClientCompanyCatSpecID "
                sSQL = sSQL & ") "
    sSQL = sSQL & ") As gndTtlSiteLocation, "
    sSQL = sSQL & "SUM(RTI.ReplacementCost) As RH_FullCostRepairReplace, "
    sSQL = sSQL & "SUM(RecoverableDep) As RH_RecDep, "
    sSQL = sSQL & "SUM(NonRecoverableDep) As RH_NonRecDep, "
    sSQL = sSQL & "SUM(ACVClaim) As RH_ACV, "
    sSQL = sSQL & "SUM(AppliedDeductible) As RH_Deductible, "
    sSQL = sSQL & "SUM(ExcessLimits) AS RH_LessExcessLimits, "
    sSQL = sSQL & "SUM(Miscellaneous) AS RH_LessMiscellaneous, "
    sSQL = sSQL & "(SUM(ACVLessExcessLimits) - SUM(AppliedDeductible)) As RH_NetACVC "
    sSQL = sSQL & "FROM Assignments A "
    sSQL = sSQL & "LEFT JOIN RTIndemnity RTI ON (A.AssignmentsID = RTI.AssignmentsID) "
    sSQL = sSQL & "WHERE A.AssignmentsID = " & msAssignmentsID & " "
    sSQL = sSQL & "AND A.IsDeleted = 0 "
    sSQL = sSQL & "AND (RTI.IsDeleted = 0 Or RTI.IsDeleted Is Null) "
    sSQL = sSQL & "AND (RTI.IsPreviousPayment = 0 Or RTI.IsPreviousPayment Is Null) "
    sSQL = sSQL & "AND (RTI.ClassOfLossID Not IN ( "
                                sSQL = sSQL & "SELECT   ClassOfLossID "
                                sSQL = sSQL & "FROM     ClassOfLoss "
                                sSQL = sSQL & "WHERE    ClassTypeID IN ( "
                                                        sSQL = sSQL & "SELECT   ClassTypeID "
                                                        sSQL = sSQL & "FROM     ClassType "
                                                        sSQL = sSQL & "WHERE    Class = 'OTHER' "
                                                        sSQL = sSQL & ") "
                                sSQL = sSQL & ") "
     sSQL = sSQL & "Or RTI.ClassOfLossID Is Null) "
    sSQL = sSQL & "GROUP BY A.Insured, A.CLIENTNUM, A.AdjusterSpecID , A.ClientCompanyCatSpecID, A.CloseDate, A.InspectedDate "
    
    Set mRS = New ADODB.Recordset
    mRS.CursorLocation = adUseClient
    mRS.Open sSQL, mConn, adOpenForwardOnly, adLockReadOnly
    Set mRS.ActiveConnection = Nothing
    With PopulateudtarRpt
            .f_Insured = vbNullString
            .f_ClientClaimNo = vbNullString
            .f_ADJNAME = vbNullString
            .f_gndTtlCatName = vbNullString
            .f_gndTtlSiteLocation = vbNullString
            .f_CloseDate = NULL_DATE
            .f_InspectedDate = NULL_DATE
            .f_RH_FullCostRepairReplace = 0
            .f_RH_RecDep = 0
            .f_RH_NonRecDep = 0
            .f_RH_ACV = 0
            .f_RH_Deductible = 0
            .f_RH_LessExcessLimits = 0
            .f_RH_LessMiscellaneous = 0
            .f_RH_NetACVC = 0
            On Error Resume Next
            .f_Insured = mRS!Insured
            .f_ClientClaimNo = mRS!ClientClaimNo
            .f_ADJNAME = mRS!ADJNAME
            .f_gndTtlCatName = mRS!gndTtlCatName
            .f_gndTtlSiteLocation = mRS!gndTtlSiteLocation
            .f_CloseDate = mRS!CloseDate
            .f_InspectedDate = mRS!InspectedDate
            .f_RH_FullCostRepairReplace = mRS!RH_FullCostRepairReplace
            .f_RH_RecDep = mRS!RH_RecDep
            .f_RH_NonRecDep = mRS!RH_NonRecDep
            .f_RH_ACV = mRS!RH_ACV
            .f_RH_Deductible = mRS!RH_Deductible
            .f_RH_LessExcessLimits = mRS!RH_LessExcessLimits
            .f_RH_LessMiscellaneous = mRS!RH_LessMiscellaneous
            .f_RH_NetACVC = mRS!RH_NetACVC
            On Error GoTo EH
        End With

    'Get Indemnity RS for collection population.
    sSQL = "SELECT RETIndem.* FROM( "
    sSQL = sSQL & "SELECT "
    sSQL = sSQL & "RTI.RTIndemnityID, "
    sSQL = sSQL & "( "
    If goUtil.gMainDB Is Nothing Then
        '------------------------------------BEGIN SQL SERVER DB-------------------------
        '
        '--------------------------------------------------------------------------------
'        sSQL = sSQL & "CASE WHEN  RTC.CheckNum Is Not Null "
'        sSQL = sSQL & "THEN "
'                    sSQL = sSQL & "( "
'                    sSQL = sSQL & "Cast(RTC.CheckNum As Varchar(10)) + ' Of ' + "
'                    sSQL = sSQL & "(SELECT Cast(MAX(CheckNum) as Varchar(10)) FROM RTChecks WHERE AssignmentsID = " & msAssignmentsID & ") "
'                    sSQL = sSQL & ") "
'        sSQL = sSQL & "ELSE '' "
'        sSQL = sSQL & "END "

        sSQL = sSQL & "CASE WHEN  RTC.CheckNum Is Not Null "
        sSQL = sSQL & "THEN "
                    sSQL = sSQL & "( "
                    sSQL = sSQL & "Cast(RTC.CheckNum As Varchar(10)) "
                    sSQL = sSQL & ") "
        sSQL = sSQL & "ELSE '' "
        sSQL = sSQL & "END "
        
        '------------------------------------END SQL SERVER DB---------------------------
        '
        '--------------------------------------------------------------------------------
    Else
        '------------------------------------BEGIN Access DB----------------------------
        '
        '-------------------------------------------------------------------------------
'        sSQL = sSQL & "IIF(RTC.CheckNum Is Not Null, "
'        'Then
'        sSQL = sSQL & "Cstr(RTC.CheckNum) & ' Of ' & "
'        sSQL = sSQL & "( "
'        sSQL = sSQL & "SELECT MAX(CheckNum) "
'        sSQL = sSQL & "FROM RTChecks "
'        sSQL = sSQL & "WHERE [AssignmentsID] = " & msAssignmentsID & " "
'        sSQL = sSQL & "), "
'        'else
'        sSQL = sSQL & "'') "

        sSQL = sSQL & "IIF(RTC.CheckNum Is Not Null, "
        'Then
        sSQL = sSQL & "Cstr(RTC.CheckNum) "
        sSQL = sSQL & ", "
        'else
        sSQL = sSQL & "'') "
        '------------------------------------END Access DB------------------------------
        '
        '-------------------------------------------------------------------------------
    End If
   
    sSQL = sSQL & ") As Payment, "
    sSQL = sSQL & "RTI.IsPreviousPayment,"
    sSQL = sSQL & "(CT.Description + ' (' + COl.Description + ')') As ClassOfLoss, "
    sSQL = sSQL & "COL.Code As ClassOfLossCode, "
    sSQL = sSQL & "TOL.TypeOfLoss + ' (' + TOL.Description + ')' As TypeOfLoss, "
    sSQL = sSQL & "TOL.Code As typeOfLossCode, "
    sSQL = sSQL & "RTI.Description, "
    sSQL = sSQL & "RTI.ReplacementCost, "
    sSQL = sSQL & "RTI.RecoverableDep, "
    sSQL = sSQL & "RTI.NonRecoverableDep, "
    sSQL = sSQL & "RTI.AcvClaim, "
    sSQL = sSQL & "RTI.SpecialLimits, "
    sSQL = sSQL & "RTI.IsAddAmountOfInsurance, "
    sSQL = sSQL & "RTI.ExcessLimits, "
    sSQL = sSQL & "RTI.Miscellaneous, "
    sSQL = sSQL & "RTI.ACVLessExcessLimits "
    sSQL = sSQL & "FROM (((RTIndemnity AS RTI LEFT JOIN RTChecks AS RTC ON RTI.RTChecksID = RTC.RTChecksID) "
    sSQL = sSQL & "LEFT JOIN TypeOfLoss AS TOL ON RTI.TypeOfLossID = TOL.TypeOfLossID) "
    sSQL = sSQL & "LEFT JOIN ClassOfLoss AS COL ON RTI.ClassOfLossID = COL.ClassOfLossID) "
    sSQL = sSQL & "LEFT JOIN ClassType AS CT ON COL.ClassTypeID = CT.ClassTypeID "
    sSQL = sSQL & "WHERE RTI.AssignmentsID = " & msAssignmentsID & " "
    sSQL = sSQL & ") As  RETIndem "
    sSQL = sSQL & "ORDER BY RETIndem.[Payment], RETIndem.[IsPreviousPayment], ABS(RETIndem.[RTIndemnityID]) "
    

    Set mRS = New ADODB.Recordset
    mRS.CursorLocation = adUseClient
    mRS.Open sSQL, mConn, adOpenForwardOnly, adLockReadOnly
    Set mRS.ActiveConnection = Nothing

    Set PopulateudtarRpt.coludtIndemnity01 = New Collection
    If Not mRS.EOF Then
        mRS.MoveFirst
        'First need to build the Photo Path. The RS will only contain the
        'Name of the photo. depending upon what DB working with the actual
        'Photo path will differ.
        Do Until mRS.EOF
            'SQL Server the Photo path will be
            With MyIndemnity
                .f_RTIndemnityID = 0
                .f_Payment = vbNullString
                .f_IsPreviousPayment = 0
                .f_ClassOfLoss = vbNullString
                .f_ClassOfLossCode = vbNullString
                .f_TypeOfLoss = vbNullString
                .f_typeOfLossCode = vbNullString
                .f_Description = vbNullString
                .f_ReplacementCost = 0
                .f_RecoverableDep = 0
                .f_NonRecoverableDep = 0
                .f_AcvClaim = 0
                .f_SpecialLimits = 0
                .f_IsAddAmountOfInsurance = 0
                .f_ExcessLimits = 0
                .f_Miscellaneous = 0
                .f_ACVLessExcessLimits = 0
                On Error Resume Next
                .f_RTIndemnityID = mRS!RTIndemnityID
                .f_Payment = mRS!Payment
                .f_IsPreviousPayment = mRS!IsPreviousPayment
                .f_ClassOfLoss = mRS!ClassOfLoss
                .f_ClassOfLossCode = mRS!ClassOfLossCode
                .f_TypeOfLoss = mRS!TypeOfLoss
                .f_typeOfLossCode = mRS!typeOfLossCode
                .f_Description = mRS!Description
                .f_ReplacementCost = mRS!ReplacementCost
                .f_RecoverableDep = mRS!RecoverableDep
                .f_NonRecoverableDep = mRS!NonRecoverableDep
                .f_AcvClaim = mRS!AcvClaim
                .f_SpecialLimits = mRS!SpecialLimits
                .f_IsAddAmountOfInsurance = mRS!IsAddAmountOfInsurance
                .f_ExcessLimits = mRS!ExcessLimits
                .f_Miscellaneous = mRS!Miscellaneous
                .f_ACVLessExcessLimits = mRS!ACVLessExcessLimits
                On Error GoTo EH
            End With
            PopulateudtarRpt.coludtIndemnity01.Add MyIndemnity, CStr(MyIndemnity.f_RTIndemnityID)
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
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsRptDB_ClassName & vbCrLf & "Private Function PopulateUDTRCAABC"
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
    sMiscReportParamName = "MiscReportParam06"
    '2/8/2004 MiscReportParam , and MiscReportParam01 to MiscReportParam30
    
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





