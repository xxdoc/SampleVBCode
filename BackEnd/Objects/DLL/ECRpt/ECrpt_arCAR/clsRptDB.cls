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
Private marRpt As ECrpt_arCAR.arCAR

Public Type udtCAR01
    AssignmentsID As Long
    f_Adjuster As String
    f_CatCode As String
    f_ClaimNumber As String
    f_Status As String
    f_RCVClaim As Currency
    F_RecDep As Currency
    f_NonRecDep As Currency
    f_Deductible As Currency
    f_ExcessLimits As Currency
    f_NetACVClaim As Currency
End Type

Private Type udtarRpt
    f_PrintDate As Date
    f_ADJName As String
    f_gndTtlCatName As String
    f_gndTtlSiteLocation As String
    f_RH_TotalNumClaims As Long
    f_RH_NumClosedClaims As Long
    f_RH_AvgRCVClaim As Currency
    f_RH_AvgRecDep As Currency
    f_RH_AvgNonRecDep As Currency
    f_RH_AvgDeductible As Currency
    f_RH_AvgExcessLimits As Currency
    f_RH_AvgNetACVClaim As Currency
    f_RH_ReplacementCostLessDeductible As Currency
    f_RH_AvgNetACVPlusRecDep As Currency
    f_RH_Closed As Long
    f_RH_Open As Long
    f_RH_ClosedPct As Double
    f_RH_OpenPct As Double
    f_RH_TotalNumClaims2 As Long
    f_RF_ClosedTtl_RCVClaim As Currency
    f_RF_ClosedTtl_RecDep As Currency
    f_RF_ClosedTtl_NonRecDep As Currency
    f_RF_ClosedTtl_Deductible As Currency
    f_RF_ClosedTtl_ExcessLimits As Currency
    f_RF_ClosedTtl_NetACVClaim As Currency
    f_RF_OpenTtl_RCVClaim As Currency
    f_RF_OpenTtl_RecDep As Currency
    f_RF_OpenTtl_NonRecDep As Currency
    f_RF_OpenTtl_Deductible As Currency
    f_RF_OpenTtl_ExcessLimits As Currency
    f_RF_OpenTtl_NetACVClaim As Currency
    f_RF_GrandTtl_RCVClaim As Currency
    f_RF_GrandTtl_RecDep As Currency
    f_RF_GrandTtl_NonRecDep As Currency
    f_RF_GrandTtl_Deductible As Currency
    f_RF_GrandTtl_ExcessLimits As Currency
    f_RF_GrandTtl_NetACVClaim As Currency
    coludtCAR01 As Collection
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
        Set marRpt = New ECrpt_arCAR.arCAR
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
    
    Set marRpt = New ECrpt_arCAR.arCAR
    
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
        marRpt.SetProperty "f_gndTtlCatName", .f_gndTtlCatName, vbString
        marRpt.SetProperty "f_gndTtlSiteLocation", .f_gndTtlSiteLocation, vbString
        marRpt.SetProperty "f_RH_TotalNumClaims", .f_RH_TotalNumClaims, vbLong
        marRpt.SetProperty "f_RH_NumClosedClaims", .f_RH_NumClosedClaims, vbLong
        marRpt.SetProperty "f_RH_AvgRCVClaim", .f_RH_AvgRCVClaim, vbCurrency
        marRpt.SetProperty "f_RH_AvgRecDep", .f_RH_AvgRecDep, vbCurrency
        marRpt.SetProperty "f_RH_AvgNonRecDep", .f_RH_AvgNonRecDep, vbCurrency
        marRpt.SetProperty "f_RH_AvgDeductible", .f_RH_AvgDeductible, vbCurrency
        marRpt.SetProperty "f_RH_AvgExcessLimits", .f_RH_AvgExcessLimits, vbCurrency
        marRpt.SetProperty "f_RH_AvgNetACVClaim", .f_RH_AvgNetACVClaim, vbCurrency
        marRpt.SetProperty "f_RH_ReplacementCostLessDeductible", .f_RH_ReplacementCostLessDeductible, vbCurrency
        marRpt.SetProperty "f_RH_AvgNetACVPlusRecDep", .f_RH_AvgNetACVPlusRecDep, vbCurrency
        marRpt.SetProperty "f_RH_Closed", .f_RH_Closed, vbLong
        marRpt.SetProperty "f_RH_Open", .f_RH_Open, vbLong
        marRpt.SetProperty "f_RH_ClosedPct", .f_RH_ClosedPct, vbDouble
        marRpt.SetProperty "f_RH_OpenPct", .f_RH_OpenPct, vbDouble
        marRpt.SetProperty "f_RH_TotalNumClaims2", .f_RH_TotalNumClaims2, vbLong
        marRpt.SetProperty "f_RF_ClosedTtl_RCVClaim", .f_RF_ClosedTtl_RCVClaim, vbCurrency
        marRpt.SetProperty "f_RF_ClosedTtl_RecDep", .f_RF_ClosedTtl_RecDep, vbCurrency
        marRpt.SetProperty "f_RF_ClosedTtl_NonRecDep", .f_RF_ClosedTtl_NonRecDep, vbCurrency
        marRpt.SetProperty "f_RF_ClosedTtl_Deductible", .f_RF_ClosedTtl_Deductible, vbCurrency
        marRpt.SetProperty "f_RF_ClosedTtl_ExcessLimits", .f_RF_ClosedTtl_ExcessLimits, vbCurrency
        marRpt.SetProperty "f_RF_ClosedTtl_NetACVClaim", .f_RF_ClosedTtl_NetACVClaim, vbCurrency
        marRpt.SetProperty "f_RF_OpenTtl_RCVClaim", .f_RF_OpenTtl_RCVClaim, vbCurrency
        marRpt.SetProperty "f_RF_OpenTtl_RecDep", .f_RF_OpenTtl_RecDep, vbCurrency
        marRpt.SetProperty "f_RF_OpenTtl_NonRecDep", .f_RF_OpenTtl_NonRecDep, vbCurrency
        marRpt.SetProperty "f_RF_OpenTtl_Deductible", .f_RF_OpenTtl_Deductible, vbCurrency
        marRpt.SetProperty "f_RF_OpenTtl_ExcessLimits", .f_RF_OpenTtl_ExcessLimits, vbCurrency
        marRpt.SetProperty "f_RF_OpenTtl_NetACVClaim", .f_RF_OpenTtl_NetACVClaim, vbCurrency
        marRpt.SetProperty "f_RF_GrandTtl_RCVClaim", .f_RF_GrandTtl_RCVClaim, vbCurrency
        marRpt.SetProperty "f_RF_GrandTtl_RecDep", .f_RF_GrandTtl_RecDep, vbCurrency
        marRpt.SetProperty "f_RF_GrandTtl_NonRecDep", .f_RF_GrandTtl_NonRecDep, vbCurrency
        marRpt.SetProperty "f_RF_GrandTtl_Deductible", .f_RF_GrandTtl_Deductible, vbCurrency
        marRpt.SetProperty "f_RF_GrandTtl_ExcessLimits", .f_RF_GrandTtl_ExcessLimits, vbCurrency
        marRpt.SetProperty "f_RF_GrandTtl_NetACVClaim", .f_RF_GrandTtl_NetACVClaim, vbCurrency
        marRpt.SetProperty "coludtCAR01", .coludtCAR01, vbObject
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
    Dim sSQLWhere As String
    Dim sSQLRTIndemnityAnd As String
    Dim colParams As Collection
    Dim vParams As Variant
    Dim MyParams As rptMiscDelimParam
    Dim MyCar As udtCAR01
    
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
    

    sSQL = sSQL & "WHERE ClientCompanyCatSpecID IN "
                                        sSQL = sSQL & "( "
                                        sSQL = sSQL & "SELECT   ClientCompanyCatSpecID "
                                        sSQL = sSQL & "FROM     ClientCompanyCatSpec "
                                        sSQL = sSQL & "WHERE    ClientCompanyID = " & msClientCompanyID & " "
                                        sSQL = sSQL & "AND      CATID = " & msCATID & " "
                                        sSQL = sSQL & ") "
    If msUSERSID <> "-1" Then
        sSQL = sSQL & "AND AdjusterSpecID IN "
                                            sSQL = sSQL & "( "
                                            sSQL = sSQL & "SELECT   ClientCoAdjusterSpecID "
                                            sSQL = sSQL & "FROM     ClientCoAdjusterSpec "
                                            sSQL = sSQL & "WHERE    USERSID = " & msUSERSID & " "
                                            sSQL = sSQL & ") "
    End If
    sSQL = sSQL & "AND Assignments.IsDeleted = 0 "
    sSQL = sSQL & "AND (Assignments.Reassigned = 0 Or Assignments.Reassigned Is Null) "
    sSQL = sSQL & "AND Assignments.CLIENTNUM Not Like '%TEST%' "
    sSQLWhere = sSQL
    
    'Build RTIndemnity And Statement
    sSQL = "AND (RTIndemnity.IsDeleted = 0 Or RTIndemnity.IsDeleted Is Null) "
    sSQL = sSQL & "AND (RTIndemnity.IsPreviousPayment = 0 Or RTIndemnity.IsPreviousPayment Is Null) "
    sSQL = sSQL & "AND (RTIndemnity.ClassOfLossID Not IN ( "
                         sSQL = sSQL & "SELECT   ClassOfLossID "
                         sSQL = sSQL & "FROM     ClassOfLoss "
                         sSQL = sSQL & "WHERE    ClassTypeID IN ( "
                                                 sSQL = sSQL & "SELECT   ClassTypeID "
                                                 sSQL = sSQL & "FROM     ClassType "
                                                 sSQL = sSQL & "WHERE    Class = 'OTHER' "
                                                 sSQL = sSQL & ") "
                         sSQL = sSQL & ") "
    sSQL = sSQL & "Or RTIndemnity.ClassOfLossID Is Null) "
    sSQLRTIndemnityAnd = sSQL
    
    'Get thr CAR Report Header Footer Info
    sSQL = "SELECT "
    sSQL = sSQL & "( "
     'Check for UserID if its "-1" then don't filter
    If msUSERSID <> "-1" Then
        sSQL = sSQL & "SELECT   LastName + ', ' + FirstName "
        sSQL = sSQL & "FROM     USERS "
        sSQL = sSQL & "WHERE    USERSID = " & msUSERSID & " "
    Else
        sSQL = sSQL & "'ALL' "
    End If
    sSQL = sSQL & ") As ADJName, "
    sSQL = sSQL & "( "
    sSQL = sSQL & "SELECT   Name + ' - ' + Description "
    sSQL = sSQL & "FROM     CAT "
    sSQL = sSQL & "WHERE    CATID = " & msCATID & " "
    sSQL = sSQL & ") As gndTtlCatName, "
    sSQL = sSQL & "( "
    sSQL = sSQL & "SELECT   SACity + ', ' + SAState "
    sSQL = sSQL & "FROM     ClientCompanyCat "
    sSQL = sSQL & "WHERE    ClientCompanyID = " & msClientCompanyID & " "
    sSQL = sSQL & "AND      CATID = " & msCATID & " "
    sSQL = sSQL & ") As gndTtlSiteLocation, "
    sSQL = sSQL & "COUNT(A.AssignmentsID) As TotalNumClaims, "
    'Closed
    sSQL = sSQL & "( "
    sSQL = sSQL & "SELECT   COUNT(AssignmentsID) As CountOFAssignmentsID "
    sSQL = sSQL & "FROM     Assignments "
    sSQL = sSQL & sSQLWhere
    sSQL = sSQL & "AND StatusID = " & AssgnStatus.iAssignmentsStatus_CLOSED & " "
    sSQL = sSQL & ") As Closed, "
    'Open
    sSQL = sSQL & "( "
    sSQL = sSQL & "SELECT   COUNT(AssignmentsID) As CountOFAssignmentsID "
    sSQL = sSQL & "FROM     Assignments "
    sSQL = sSQL & sSQLWhere
    sSQL = sSQL & "AND StatusID <> " & AssgnStatus.iAssignmentsStatus_CLOSED & " "
    sSQL = sSQL & ") As [Open], "
    
    'Closed Totals
    '1. ClosedTtl_RCVClaim
    sSQL = sSQL & "( "
    sSQL = sSQL & "SELECT   SUM(RTIndemnity.ReplacementCost) As SumReplacementCost "
    sSQL = sSQL & "FROM     Assignments "
    sSQL = sSQL & "LEFT JOIN RTIndemnity ON (Assignments.AssignmentsID = RTIndemnity.AssignmentsID) "
    sSQL = sSQL & sSQLWhere
    sSQL = sSQL & "AND StatusID = " & AssgnStatus.iAssignmentsStatus_CLOSED & " "
    sSQL = sSQL & sSQLRTIndemnityAnd
    sSQL = sSQL & ") As ClosedTtl_RCVClaim, "
    '2. ClosedTtl_RecDep
    sSQL = sSQL & "( "
    sSQL = sSQL & "SELECT   SUM(RTIndemnity.RecoverableDep) As SumRecoverableDep "
    sSQL = sSQL & "FROM     Assignments "
    sSQL = sSQL & "LEFT JOIN RTIndemnity ON (Assignments.AssignmentsID = RTIndemnity.AssignmentsID) "
    sSQL = sSQL & sSQLWhere
    sSQL = sSQL & "AND StatusID = " & AssgnStatus.iAssignmentsStatus_CLOSED & " "
    sSQL = sSQL & sSQLRTIndemnityAnd
    sSQL = sSQL & ") As ClosedTtl_RecDep, "
    '3. ClosedTtl_NonRecDep
    sSQL = sSQL & "( "
    sSQL = sSQL & "SELECT   SUM(RTIndemnity.NonRecoverableDep) As SumNonRecoverableDep "
    sSQL = sSQL & "FROM     Assignments "
    sSQL = sSQL & "LEFT JOIN RTIndemnity ON (Assignments.AssignmentsID = RTIndemnity.AssignmentsID) "
    sSQL = sSQL & sSQLWhere
    sSQL = sSQL & "AND StatusID = " & AssgnStatus.iAssignmentsStatus_CLOSED & " "
    sSQL = sSQL & sSQLRTIndemnityAnd
    sSQL = sSQL & ") As ClosedTtl_NonRecDep, "
    '4. ClosedTtl_Deductible
    sSQL = sSQL & "( "
    sSQL = sSQL & "SELECT  SUM(Assignments.Deductible) As SumDeductible "
    sSQL = sSQL & "FROM     Assignments "
    sSQL = sSQL & sSQLWhere
    sSQL = sSQL & "AND StatusID = " & AssgnStatus.iAssignmentsStatus_CLOSED & " "
    sSQL = sSQL & ") As ClosedTtl_Deductible, "
    '5. ClosedTtl_ExcessLimits
    sSQL = sSQL & "( "
    sSQL = sSQL & "SELECT   SUM(RTIndemnity.ExcessLimits) As SumExcessLimits "
    sSQL = sSQL & "FROM     Assignments "
    sSQL = sSQL & "LEFT JOIN RTIndemnity ON (Assignments.AssignmentsID = RTIndemnity.AssignmentsID) "
    sSQL = sSQL & sSQLWhere
    sSQL = sSQL & "AND StatusID = " & AssgnStatus.iAssignmentsStatus_CLOSED & " "
    sSQL = sSQL & sSQLRTIndemnityAnd
    sSQL = sSQL & ") As ClosedTtl_ExcessLimits, "
    '6. ClosedTtl_NetACVClaim
    sSQL = sSQL & "( "
    sSQL = sSQL & "SELECT   SUM(RTIndemnity.ACVLessExcessLimits) As SumACVLessExcessLimits "
    sSQL = sSQL & "FROM     Assignments "
    sSQL = sSQL & "LEFT JOIN RTIndemnity ON (Assignments.AssignmentsID = RTIndemnity.AssignmentsID) "
    sSQL = sSQL & sSQLWhere
    sSQL = sSQL & "AND StatusID = " & AssgnStatus.iAssignmentsStatus_CLOSED & " "
    sSQL = sSQL & sSQLRTIndemnityAnd
    sSQL = sSQL & ") As ClosedTtl_NetACVClaim, "
    
    'Open Totals
    '1. OpenTtl_RCVClaim
    sSQL = sSQL & "( "
    sSQL = sSQL & "SELECT   SUM(RTIndemnity.ReplacementCost) As SumReplacementCost "
    sSQL = sSQL & "FROM     Assignments "
    sSQL = sSQL & "LEFT JOIN RTIndemnity ON (Assignments.AssignmentsID = RTIndemnity.AssignmentsID) "
    sSQL = sSQL & sSQLWhere
    sSQL = sSQL & "AND StatusID <> " & AssgnStatus.iAssignmentsStatus_CLOSED & " "
    sSQL = sSQL & sSQLRTIndemnityAnd
    sSQL = sSQL & ") As OpenTtl_RCVClaim, "
    '2. OpenTtl_RecDep
    sSQL = sSQL & "( "
    sSQL = sSQL & "SELECT   SUM(RTIndemnity.RecoverableDep) As SumRecoverableDep "
    sSQL = sSQL & "FROM     Assignments "
    sSQL = sSQL & "LEFT JOIN RTIndemnity ON (Assignments.AssignmentsID = RTIndemnity.AssignmentsID) "
    sSQL = sSQL & sSQLWhere
    sSQL = sSQL & "AND StatusID <> " & AssgnStatus.iAssignmentsStatus_CLOSED & " "
    sSQL = sSQL & sSQLRTIndemnityAnd
    sSQL = sSQL & ") As OpenTtl_RecDep, "
    '3. OpenTtl_NonRecDep
    sSQL = sSQL & "( "
    sSQL = sSQL & "SELECT   SUM(RTIndemnity.NonRecoverableDep) As SumNonRecoverableDep "
    sSQL = sSQL & "FROM     Assignments "
    sSQL = sSQL & "LEFT JOIN RTIndemnity ON (Assignments.AssignmentsID = RTIndemnity.AssignmentsID) "
    sSQL = sSQL & sSQLWhere
    sSQL = sSQL & "AND StatusID <> " & AssgnStatus.iAssignmentsStatus_CLOSED & " "
    sSQL = sSQL & sSQLRTIndemnityAnd
    sSQL = sSQL & ") As OpenTtl_NonRecDep, "
    '4. OpenTtl_Deductible
    sSQL = sSQL & "( "
    sSQL = sSQL & "SELECT SUM(Assignments.Deductible) As SumDeductible "
    sSQL = sSQL & "FROM     Assignments "
    sSQL = sSQL & sSQLWhere
    sSQL = sSQL & "AND StatusID <> " & AssgnStatus.iAssignmentsStatus_CLOSED & " "
    sSQL = sSQL & ") As OpenTtl_Deductible, "
    '5. ClosedTtl_ExcessLimits
    sSQL = sSQL & "( "
    sSQL = sSQL & "SELECT   SUM(RTIndemnity.ExcessLimits) As SumExcessLimits "
    sSQL = sSQL & "FROM     Assignments "
    sSQL = sSQL & "LEFT JOIN RTIndemnity ON (Assignments.AssignmentsID = RTIndemnity.AssignmentsID) "
    sSQL = sSQL & sSQLWhere
    sSQL = sSQL & "AND StatusID <> " & AssgnStatus.iAssignmentsStatus_CLOSED & " "
    sSQL = sSQL & sSQLRTIndemnityAnd
    sSQL = sSQL & ") As OpenTtl_ExcessLimits, "
    '6. OpenTtl_NetACVClaim
    sSQL = sSQL & "( "
    sSQL = sSQL & "SELECT   SUM(RTIndemnity.ACVLessExcessLimits) As SumACVLessExcessLimits "
    sSQL = sSQL & "FROM     Assignments "
    sSQL = sSQL & "LEFT JOIN RTIndemnity ON (Assignments.AssignmentsID = RTIndemnity.AssignmentsID) "
    sSQL = sSQL & sSQLWhere
    sSQL = sSQL & "AND StatusID <> " & AssgnStatus.iAssignmentsStatus_CLOSED & " "
    sSQL = sSQL & sSQLRTIndemnityAnd
    sSQL = sSQL & ") As OpenTtl_NetACVClaim "
    sSQL = sSQL & "FROM Assignments A "
    sSQL = sSQL & "WHERE A.ClientCompanyCatSpecID IN "
                                    sSQL = sSQL & "( "
                                    sSQL = sSQL & "SELECT   ClientCompanyCatSpecID "
                                    sSQL = sSQL & "FROM     ClientCompanyCatSpec "
                                    sSQL = sSQL & "WHERE    ClientCompanyID = " & msClientCompanyID & " "
                                    sSQL = sSQL & "AND      CATID = " & msCATID & " "
                                    sSQL = sSQL & ") "
    'Check for UserID if its "-1" then don't filter
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
    If Not mRS.EOF Then
        mRS.MoveFirst
        With PopulateudtarRpt
            .f_PrintDate = NULL_DATE
            .f_ADJName = vbNullString
            .f_gndTtlCatName = vbNullString
            .f_gndTtlSiteLocation = vbNullString
            .f_RH_TotalNumClaims = 0
            .f_RH_NumClosedClaims = 0
            .f_RH_AvgRCVClaim = 0
            .f_RH_AvgRecDep = 0
            .f_RH_AvgNonRecDep = 0
            .f_RH_AvgDeductible = 0
            .f_RH_AvgExcessLimits = 0
            .f_RH_AvgNetACVClaim = 0
            .f_RH_ReplacementCostLessDeductible = 0
            .f_RH_AvgNetACVPlusRecDep = 0
            .f_RH_Closed = 0
            .f_RH_Open = 0
            .f_RH_ClosedPct = 0
            .f_RH_OpenPct = 0
            .f_RH_TotalNumClaims2 = 0
            .f_RF_ClosedTtl_RCVClaim = 0
            .f_RF_ClosedTtl_RecDep = 0
            .f_RF_ClosedTtl_NonRecDep = 0
            .f_RF_ClosedTtl_Deductible = 0
            .f_RF_ClosedTtl_ExcessLimits = 0
            .f_RF_ClosedTtl_NetACVClaim = 0
            .f_RF_OpenTtl_RCVClaim = 0
            .f_RF_OpenTtl_RecDep = 0
            .f_RF_OpenTtl_NonRecDep = 0
            .f_RF_OpenTtl_Deductible = 0
            .f_RF_OpenTtl_ExcessLimits = 0
            .f_RF_OpenTtl_NetACVClaim = 0
            .f_RF_GrandTtl_RCVClaim = 0
            .f_RF_GrandTtl_RecDep = 0
            .f_RF_GrandTtl_NonRecDep = 0
            .f_RF_GrandTtl_Deductible = 0
            .f_RF_GrandTtl_ExcessLimits = 0
            .f_RF_GrandTtl_NetACVClaim = 0
            On Error Resume Next
            .f_PrintDate = Now()
            .f_ADJName = mRS!ADJName
            .f_gndTtlCatName = mRS!gndTtlCatName
            .f_gndTtlSiteLocation = mRS!gndTtlSiteLocation
            .f_RH_TotalNumClaims = mRS!TotalNumClaims
            .f_RH_NumClosedClaims = mRS!Closed
            .f_RH_TotalNumClaims2 = mRS!TotalNumClaims
            .f_RH_Closed = mRS!Closed
            .f_RH_Open = mRS!Open
            .f_RF_ClosedTtl_RCVClaim = mRS!ClosedTtl_RCVClaim
            .f_RF_ClosedTtl_RecDep = mRS!ClosedTtl_RecDep
            .f_RF_ClosedTtl_NonRecDep = mRS!ClosedTtl_NonRecDep
            .f_RF_ClosedTtl_Deductible = mRS!ClosedTtl_Deductible
            .f_RF_ClosedTtl_ExcessLimits = mRS!ClosedTtl_ExcessLimits
            .f_RF_ClosedTtl_NetACVClaim = mRS!ClosedTtl_NetACVClaim - .f_RF_ClosedTtl_Deductible
            .f_RF_OpenTtl_RCVClaim = mRS!OpenTtl_RCVClaim
            .f_RF_OpenTtl_RecDep = mRS!OpenTtl_RecDep
            .f_RF_OpenTtl_NonRecDep = mRS!OpenTtl_NonRecDep
            .f_RF_OpenTtl_Deductible = mRS!OpenTtl_Deductible
            .f_RF_OpenTtl_ExcessLimits = mRS!OpenTtl_ExcessLimits
            .f_RF_OpenTtl_NetACVClaim = mRS!OpenTtl_NetACVClaim - .f_RF_OpenTtl_Deductible
            .f_RF_GrandTtl_RCVClaim = .f_RF_ClosedTtl_RCVClaim + .f_RF_OpenTtl_RCVClaim
            .f_RF_GrandTtl_RecDep = .f_RF_ClosedTtl_RecDep + .f_RF_OpenTtl_RecDep
            .f_RF_GrandTtl_NonRecDep = .f_RF_ClosedTtl_NonRecDep + .f_RF_OpenTtl_NonRecDep
            .f_RF_GrandTtl_Deductible = .f_RF_ClosedTtl_Deductible + .f_RF_OpenTtl_Deductible
            .f_RF_GrandTtl_ExcessLimits = .f_RF_ClosedTtl_ExcessLimits + .f_RF_OpenTtl_ExcessLimits
            .f_RF_GrandTtl_NetACVClaim = .f_RF_ClosedTtl_NetACVClaim + .f_RF_OpenTtl_NetACVClaim
            'Closed Claims Only (Averages)
            'Be sure to check for Div by 0 errors
            If .f_RH_Closed = 0 Then
                .f_RH_AvgRCVClaim = 0
                .f_RH_AvgRecDep = 0
                .f_RH_AvgNonRecDep = 0
                .f_RH_AvgDeductible = 0
                .f_RH_AvgExcessLimits = 0
                .f_RH_AvgNetACVClaim = 0
            Else
                .f_RH_AvgRCVClaim = .f_RF_ClosedTtl_RCVClaim / .f_RH_Closed
                .f_RH_AvgRecDep = .f_RF_ClosedTtl_RecDep / .f_RH_Closed
                .f_RH_AvgNonRecDep = .f_RF_ClosedTtl_NonRecDep / .f_RH_Closed
                .f_RH_AvgDeductible = .f_RF_ClosedTtl_Deductible / .f_RH_Closed
                .f_RH_AvgExcessLimits = .f_RF_ClosedTtl_ExcessLimits / .f_RH_Closed
                .f_RH_AvgNetACVClaim = .f_RF_ClosedTtl_NetACVClaim / .f_RH_Closed
            End If
            .f_RH_ReplacementCostLessDeductible = .f_RH_AvgRCVClaim - .f_RH_AvgDeductible
            .f_RH_AvgNetACVPlusRecDep = .f_RH_AvgNetACVClaim + .f_RH_AvgRecDep
            If .f_RH_TotalNumClaims = 0 Then
                .f_RH_ClosedPct = 0
                .f_RH_OpenPct = 0
            Else
                .f_RH_ClosedPct = .f_RH_Closed / .f_RH_TotalNumClaims
                .f_RH_OpenPct = .f_RH_Open / .f_RH_TotalNumClaims
            End If
            
            On Error GoTo EH
        End With
    End If
        
    'Get the Detail RS
    sSQL = "SELECT RetCar.* FROM ( "
    
    sSQL = sSQL & "SELECT "
    sSQL = sSQL & "A.AssignmentsID, "
    sSQL = sSQL & "( "
    sSQL = sSQL & "SELECT   LastName + ', ' + FirstName + ' (' + UserName + ')' "
    sSQL = sSQL & "FROM     USERS "
    sSQL = sSQL & "WHERE    USERSID = "
                            sSQL = sSQL & "( "
                            sSQL = sSQL & "SELECT   UsersID "
                            sSQL = sSQL & "FROM     ClientCoAdjusterSpec "
                            sSQL = sSQL & "WHERE    ClientCoAdjusterSpecID = A.AdjusterSpecID "
                            sSQL = sSQL & ") "
    sSQL = sSQL & ") As Adjuster, "
    sSQL = sSQL & "( "
    sSQL = sSQL & "SELECT  CatCode "
    sSQL = sSQL & "FROM     ClientCompanyCatSpec "
    sSQL = sSQL & "WHERE    ClientCompanyCatSpecID = A.ClientCompanyCatSpecID "
    sSQL = sSQL & ") As CatCode, "
    sSQL = sSQL & "A.CLIENTNUM As ClaimNumber, "
    sSQL = sSQL & "( "
    sSQL = sSQL & "SELECT  StatusAlias As Status "
    sSQL = sSQL & "FROM     Status "
    sSQL = sSQL & "WHERE    StatusID = A.[StatusID] "
    sSQL = sSQL & ") As Status, "
    sSQL = sSQL & "( "
    If goUtil.gMainDB Is Nothing Then
        '------------------------------------BEGIN SQL SERVER DB-------------------------
        '
        '--------------------------------------------------------------------------------
        sSQL = sSQL & "CASE    WHEN SUM(RTI.ReplacementCost) Is Null "
        sSQL = sSQL & "         THEN 0 "
        sSQL = sSQL & "         ELSE SUM(RTI.ReplacementCost) "
        sSQL = sSQL & "         END "
        '------------------------------------END SQL SERVER DB---------------------------
        '
        '--------------------------------------------------------------------------------
    Else
        '------------------------------------BEGIN Access DB----------------------------
        '
        '-------------------------------------------------------------------------------
        sSQL = sSQL & "IIF(SUM(RTI.ReplacementCost) Is Null, "
        'Then
        sSQL = sSQL & "0, "
        'Else
        sSQL = sSQL & "SUM(RTI.ReplacementCost) "
        'End
        sSQL = sSQL & ") "
        '------------------------------------END Access DB------------------------------
        '
        '-------------------------------------------------------------------------------
    End If
    sSQL = sSQL & ") As RCVCLaim, "
    sSQL = sSQL & "( "
    If goUtil.gMainDB Is Nothing Then
        '------------------------------------BEGIN SQL SERVER DB-------------------------
        '
        '--------------------------------------------------------------------------------
        sSQL = sSQL & "CASE    WHEN SUM(RTI.RecoverableDep) Is Null "
        sSQL = sSQL & "         THEN 0 "
        sSQL = sSQL & "         ELSE SUM(RTI.RecoverableDep) "
        sSQL = sSQL & "         END "
        '------------------------------------END SQL SERVER DB---------------------------
        '
        '--------------------------------------------------------------------------------
    Else
        '------------------------------------BEGIN Access DB----------------------------
        '
        '-------------------------------------------------------------------------------
        sSQL = sSQL & "IIF(SUM(RTI.RecoverableDep) Is Null, "
        'Then
        sSQL = sSQL & "0, "
        'Else
        sSQL = sSQL & "SUM(RTI.RecoverableDep) "
        'End
        sSQL = sSQL & ") "
        '------------------------------------END Access DB------------------------------
        '
        '-------------------------------------------------------------------------------
    End If
    sSQL = sSQL & ") As RecDep, "
    sSQL = sSQL & "( "
    If goUtil.gMainDB Is Nothing Then
        '------------------------------------BEGIN SQL SERVER DB-------------------------
        '
        '--------------------------------------------------------------------------------
        sSQL = sSQL & "CASE    WHEN SUM(RTI.NonRecoverableDep) Is Null "
        sSQL = sSQL & "         THEN 0 "
        sSQL = sSQL & "         ELSE SUM(RTI.NonRecoverableDep) "
        sSQL = sSQL & "         END "
        '------------------------------------END SQL SERVER DB---------------------------
        '
        '--------------------------------------------------------------------------------
    Else
        '------------------------------------BEGIN Access DB----------------------------
        '
        '-------------------------------------------------------------------------------
        sSQL = sSQL & "IIF(SUM(RTI.NonRecoverableDep) Is Null, "
        'Then
        sSQL = sSQL & "0, "
        'Else
        sSQL = sSQL & "SUM(RTI.NonRecoverableDep) "
        'End
        sSQL = sSQL & ") "
        '------------------------------------END Access DB------------------------------
        '
        '-------------------------------------------------------------------------------
    End If
    sSQL = sSQL & ") As NonRecDep, "
    sSQL = sSQL & "A.Deductible As Deductible, "
    sSQL = sSQL & "( "
    If goUtil.gMainDB Is Nothing Then
        '------------------------------------BEGIN SQL SERVER DB-------------------------
        '
        '--------------------------------------------------------------------------------
        sSQL = sSQL & "CASE    WHEN SUM(RTI.ExcessLimits) Is Null "
        sSQL = sSQL & "         THEN 0 "
        sSQL = sSQL & "         ELSE SUM(RTI.ExcessLimits) "
        sSQL = sSQL & "         END "
        '------------------------------------END SQL SERVER DB---------------------------
        '
        '--------------------------------------------------------------------------------
    Else
        '------------------------------------BEGIN Access DB----------------------------
        '
        '-------------------------------------------------------------------------------
        sSQL = sSQL & "IIF(SUM(RTI.ExcessLimits) Is Null, "
        'Then
        sSQL = sSQL & "0, "
        'Else
        sSQL = sSQL & "SUM(RTI.ExcessLimits) "
        'End
        sSQL = sSQL & ") "
        '------------------------------------END Access DB------------------------------
        '
        '-------------------------------------------------------------------------------
    End If
    sSQL = sSQL & ") As ExcessLimits, "
    sSQL = sSQL & "( "
    If goUtil.gMainDB Is Nothing Then
        '------------------------------------BEGIN SQL SERVER DB-------------------------
        '
        '--------------------------------------------------------------------------------
        sSQL = sSQL & "CASE    WHEN SUM(ACVLessExcessLimits) Is Null "
        sSQL = sSQL & "         THEN 0 - A.Deductible "
        sSQL = sSQL & "         ELSE SUM(RTI.ACVLessExcessLimits) - A.Deductible "
        sSQL = sSQL & "         END "
        '------------------------------------END SQL SERVER DB---------------------------
        '
        '--------------------------------------------------------------------------------
    Else
        '------------------------------------BEGIN Access DB----------------------------
        '
        '-------------------------------------------------------------------------------
        sSQL = sSQL & "IIF(SUM(ACVLessExcessLimits) Is Null, "
        'Then
        sSQL = sSQL & "0 - A.Deductible, "
        'Else
        sSQL = sSQL & "SUM(RTI.ACVLessExcessLimits) - A.Deductible "
        'End
        sSQL = sSQL & ") "
        '------------------------------------END Access DB------------------------------
        '
        '-------------------------------------------------------------------------------
    End If
    sSQL = sSQL & ") As NetACVClaim "
    sSQL = sSQL & "FROM Assignments A "
    sSQL = sSQL & "LEFT JOIN RTIndemnity RTI ON (A.AssignmentsID = RTI.AssignmentsID) "
    sSQL = sSQL & "WHERE A.ClientCompanyCatSpecID IN "
                                        sSQL = sSQL & "( "
                                        sSQL = sSQL & "SELECT   ClientCompanyCatSpecID "
                                        sSQL = sSQL & "FROM     ClientCompanyCatSpec "
                                        sSQL = sSQL & "WHERE    ClientCompanyID = " & msClientCompanyID & " "
                                        sSQL = sSQL & "AND      CATID = " & msCATID & " "
                                        sSQL = sSQL & ") "
    'Check for UserID if its "-1" then don't filter
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
    sSQL = sSQL & "GROUP BY A.AssignmentsID, A.ClientCompanyCatSpecID, A.AdjusterSpecID, A.CLIENTNUM , A.StatusID, A.Deductible "
    sSQL = sSQL & ") As RetCar "
    sSQL = sSQL & "ORDER BY Adjuster, ClaimNumber "
    
    Set PopulateudtarRpt.coludtCAR01 = New Collection
    
    If mbHideDetails Then
        mRS.Close
        Set mRS = Nothing
        Exit Function
    End If
    
    Set mRS = New ADODB.Recordset
    mRS.CursorLocation = adUseClient
    mRS.Open sSQL, mConn, adOpenStatic, adLockReadOnly
    Set mRS.ActiveConnection = Nothing
    
    
    If Not mRS.EOF Then
        mRS.MoveFirst
        'First need to build the Photo Path. The RS will only contain the
        'Name of the photo. depending upon what DB working with the actual
        'Photo path will differ.
        Do Until mRS.EOF
            'SQL Server the Photo path will be
            With MyCar
                .AssignmentsID = 0
                .f_Adjuster = vbNullString
                .f_CatCode = vbNullString
                .f_ClaimNumber = vbNullString
                .f_Status = vbNullString
                .f_RCVClaim = 0
                .F_RecDep = 0
                .f_NonRecDep = 0
                .f_Deductible = 0
                .f_ExcessLimits = 0
                .f_NetACVClaim = 0
                On Error Resume Next
                .AssignmentsID = mRS!AssignmentsID
                .f_Adjuster = mRS!Adjuster
                .f_CatCode = mRS!CatCode
                .f_ClaimNumber = mRS!ClaimNumber
                .f_Status = mRS!Status
                .f_RCVClaim = mRS!RCVCLaim
                .F_RecDep = mRS!RecDep
                .f_NonRecDep = mRS!NonRecDep
                .f_Deductible = mRS!Deductible
                .f_ExcessLimits = mRS!ExcessLimits
                .f_NetACVClaim = mRS!NetACVClaim
                On Error GoTo EH
            End With
            PopulateudtarRpt.coludtCAR01.Add MyCar, CStr(MyCar.AssignmentsID)
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
    
    'This Report does not use Parameters stored in DB
    'since it is not an Assignment specific Report
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






