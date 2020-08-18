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
Private marRpt As ECRptStateFarm_arRptIB02.arRptIB02

Public Type udtFeeItems01
    fsftTypeNum As Long
    fsftName As String
    fsftDescription As String
    fsftFeeAmount As Currency
    fsftIsExpense As Boolean
    fsftMaxNumberOfItems As Long
    fsftMaxFeeAmount As Long
    fsftIsMiscAmount As Boolean
    fsftUseFormula As Boolean
    fsftVBFormula As String
    NumberOfItems As Long
    Amount As Currency
    Comment As String
End Type

Private Type udtarRpt
    bPrintOnIB As Boolean
    CopyName  As String
    Comments As String
    f01_sSubToCarrier  As String
    f02_sIBNumber As String
    f04_sCatCode As String
    f05_sLocation  As String
    f05a_sState As String
    f06_dtDateClosed  As Date
    f07_sAdjusterName As String
    f08_sAdjCRID  As String
    f09_sSALN  As String
    f09a_sPolicyNo  As String
    f10_sInsuredName  As String
    f11_sLossLocation As String
    f12_dtDateOfLoss  As Date
    f13_cGrossLoss  As Currency
    f14_cDepreciation  As Currency
    f14a_sSupplement  As String
    f14b_sRebilled  As String
    f15_cDeductible  As Currency
    f15a_cLessExcessLimits  As Currency
    f15b_sExcessLimDesc  As String
    f15c_cLessMiscellaneous  As Currency
    f15d_cMiscellaneousDesc  As String
    f16_cNetClaim  As Currency
    f17_cServiceFee  As Currency
    f17a_cMiscServiceFee  As Currency
    f18_sServiceFeeComment As String
    f18a_sMiscServiceFeeComment  As String
    f19_cOutBuildingsFee  As Currency
    f19a_iOutBuildingsCount  As Integer
    f19b_cOutBuildingsAmount  As Currency
    f20_sOutBuildingsFeeComment As String
    f21_cTwoStoryCharge  As Currency
    f22_cSteepCharge  As Currency
    f23_cInteriorDamageCharge As Currency
    f24_cExternalDamageBGCharge  As Currency
    f25_cServiceFeeSubTotal As Currency
    f26_iPhotoCount As Integer
    f27_cPhotoFee  As Currency
    f28_iOther  As Integer
    f29_cOtherFee  As Currency
    f29a_sMiscFeeComment  As String
    f29b_cMiscFee  As Currency
    f30_cTotalExpenses  As Currency
    f31_dTaxPercent  As Double
    f32_cTaxAmount  As Currency
    f33_cTotalAdjustingFee  As Currency
    f42_ClassOfClaim  As String
    f43_sCauseOfLoss  As String
    f50_sInsuredPayeeName  As String
    f51_sPayeeNames  As String
    f52_sAddress  As String
    f53_cAmountOfCheck  As Currency
    f54_sCatCode  As String
    coludtFeeItems01 As Collection
End Type

Private msCopyName As String
Private msIBID As String
Private msIBHistoryID As String
Private msSupplement As String
Private mbExportXML As Boolean
Private mbExportXMLOnly As Boolean

Private Property Get clsRptDB_MiscDelimParamsCol() As Collection
    Dim bCleanup As Boolean
    'Add Report Parameters that are not updated by DB Fields.
    If marRpt Is Nothing Then
        bCleanup = True
        Set marRpt = New ECRptStateFarm_arRptIB02.arRptIB02
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

Public Property Let CopyName(psCopyName As String)
    msCopyName = psCopyName
End Property
Public Property Get CopyName() As String
    CopyName = msCopyName
End Property

Public Property Let IBID(psIBID As String)
    msIBID = psIBID
End Property
Public Property Get IBID() As String
    IBID = msIBID
End Property

Public Property Let IBHistoryID(psIBHistoryID As String)
    msIBHistoryID = psIBHistoryID
End Property
Public Property Get IBHistoryID() As String
    IBHistoryID = msIBHistoryID
End Property

Public Property Let Supplement(psSupplement As String)
    msSupplement = psSupplement
End Property
Public Property Get Supplement() As String
    Supplement = msSupplement
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
        mConn.ConnectionTimeout = 0
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
    
    Set marRpt = New ECRptStateFarm_arRptIB02.arRptIB02
    
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
         marRpt.SetProperty "bPrintOnIB", .bPrintOnIB, vbBoolean
         marRpt.SetProperty "CopyName", .CopyName, vbString
         marRpt.SetProperty "Comments", .Comments, vbString
         marRpt.SetProperty "f01_sSubToCarrier", .f01_sSubToCarrier, vbString
         marRpt.SetProperty "f02_sIBNumber", .f02_sIBNumber, vbString
         marRpt.SetProperty "f04_sCatCode", .f04_sCatCode, vbString
         marRpt.SetProperty "f05_sLocation", .f05_sLocation, vbString
         marRpt.SetProperty "f05a_sState", .f05a_sState, vbString
         marRpt.SetProperty "f06_dtDateClosed", .f06_dtDateClosed, vbDate
         marRpt.SetProperty "f07_sAdjusterName", .f07_sAdjusterName, vbString
         marRpt.SetProperty "f08_sAdjCRID", .f08_sAdjCRID, vbString
         marRpt.SetProperty "f09_sSALN", .f09_sSALN, vbString
         marRpt.SetProperty "f09a_sPolicyNo", .f09a_sPolicyNo, vbString
         marRpt.SetProperty "f10_sInsuredName", .f10_sInsuredName, vbString
         marRpt.SetProperty "f11_sLossLocation", .f11_sLossLocation, vbString
         marRpt.SetProperty "f12_dtDateOfLoss", .f12_dtDateOfLoss, vbDate
         marRpt.SetProperty "f13_cGrossLoss", .f13_cGrossLoss, vbCurrency
         marRpt.SetProperty "f14_cDepreciation", .f14_cDepreciation, vbCurrency
         marRpt.SetProperty "f14a_sSupplement", .f14a_sSupplement, vbString
         marRpt.SetProperty "f14b_sRebilled", .f14b_sRebilled, vbString
         marRpt.SetProperty "f15_cDeductible", .f15_cDeductible, vbCurrency
         marRpt.SetProperty "f15a_cLessExcessLimits", .f15a_cLessExcessLimits, vbCurrency
         marRpt.SetProperty "f15b_sExcessLimDesc", .f15b_sExcessLimDesc, vbString
         marRpt.SetProperty "f15c_cLessMiscellaneous", .f15c_cLessMiscellaneous, vbCurrency
         marRpt.SetProperty "f15d_cMiscellaneousDesc", .f15d_cMiscellaneousDesc, vbString
         marRpt.SetProperty "f16_cNetClaim", .f16_cNetClaim, vbCurrency
         marRpt.SetProperty "f17_cServiceFee", .f17_cServiceFee, vbCurrency
         marRpt.SetProperty "f17a_cMiscServiceFee", .f17a_cMiscServiceFee, vbCurrency
         marRpt.SetProperty "f18_sServiceFeeComment", .f18_sServiceFeeComment, vbString
         marRpt.SetProperty "f18a_sMiscServiceFeeComment", .f18a_sMiscServiceFeeComment, vbString
         marRpt.SetProperty "f19_cOutBuildingsFee", .f19_cOutBuildingsFee, vbCurrency
         marRpt.SetProperty "f19a_iOutBuildingsCount", .f19a_iOutBuildingsCount, vbInteger
         marRpt.SetProperty "f19b_cOutBuildingsAmount", .f19b_cOutBuildingsAmount, vbCurrency
         marRpt.SetProperty "f20_sOutBuildingsFeeComment", .f20_sOutBuildingsFeeComment, vbString
         marRpt.SetProperty "f21_cTwoStoryCharge", .f21_cTwoStoryCharge, vbCurrency
         marRpt.SetProperty "f22_cSteepCharge", .f22_cSteepCharge, vbCurrency
         marRpt.SetProperty "f23_cInteriorDamageCharge", .f23_cInteriorDamageCharge, vbCurrency
         marRpt.SetProperty "f24_cExternalDamageBGCharge", .f24_cExternalDamageBGCharge, vbCurrency
         marRpt.SetProperty "f25_cServiceFeeSubTotal", .f25_cServiceFeeSubTotal, vbCurrency
         marRpt.SetProperty "f26_iPhotoCount", .f26_iPhotoCount, vbInteger
         marRpt.SetProperty "f27_cPhotoFee", .f27_cPhotoFee, vbCurrency
         marRpt.SetProperty "f28_iOther", .f28_iOther, vbInteger
         marRpt.SetProperty "f29_cOtherFee", .f29_cOtherFee, vbCurrency
         marRpt.SetProperty "f29a_sMiscFeeComment", .f29a_sMiscFeeComment, vbString
         marRpt.SetProperty "f29b_cMiscFee", .f29b_cMiscFee, vbCurrency
         marRpt.SetProperty "f30_cTotalExpenses", .f30_cTotalExpenses, vbCurrency
         marRpt.SetProperty "f31_dTaxPercent", .f31_dTaxPercent, vbDouble
         marRpt.SetProperty "f32_cTaxAmount", .f32_cTaxAmount, vbCurrency
         marRpt.SetProperty "f33_cTotalAdjustingFee", .f33_cTotalAdjustingFee, vbCurrency
         marRpt.SetProperty "f42_ClassOfClaim", .f42_ClassOfClaim, vbString
         marRpt.SetProperty "f43_sCauseOfLoss", .f43_sCauseOfLoss, vbString
         marRpt.SetProperty "f50_sInsuredPayeeName", .f50_sInsuredPayeeName, vbString
         marRpt.SetProperty "f51_sPayeeNames", .f51_sPayeeNames, vbString
         marRpt.SetProperty "f52_sAddress", .f52_sAddress, vbString
         marRpt.SetProperty "f53_cAmountOfCheck", .f53_cAmountOfCheck, vbCurrency
         marRpt.SetProperty "f54_sCatCode", .f54_sCatCode, vbString
         marRpt.SetProperty "coludtFeeItems01", .coludtFeeItems01, vbObject
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
    Set udtarRpt.coludtFeeItems01 = Nothing
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
    Dim lErrNum As Long
    Dim sErrDesc As String
    'USed to Accesss AddlCheck Report(Payment Requests)
    Dim lMainSPVerison As Long
    Dim larRptAddlChkVersion As Long
    Dim RSSPVersion As ADODB.Recordset
    Dim sParams As String
    Dim colParams As Collection
    Dim vParams As Variant
    Dim MyParams As MiscReportParam
    Dim oCarList As V2ECKeyBoard.clsCarLists
    'For Fee items
    Dim RSFeeItems As ADODB.Recordset
    Dim MyFeeItem As udtFeeItems01
    Dim bAddFeeItem As Boolean
    Dim sSevCode As String
    Dim sPrevSevCode As String
    Dim sAddedSevCodes As String
    Dim sFeeScheduleID As String
    

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
    
    
    sSQL = "SELECT "
    sSQL = sSQL & "( "
    If goUtil.gMainDB Is Nothing Then
        '------------------------------------BEGIN SQL SERVER DB-------------------------
        '
        '--------------------------------------------------------------------------------
        sSQL = sSQL & "CASE   WHEN RTChecks.[PrintOnIB] Is Null "
        sSQL = sSQL & "THEN    0 "
        sSQL = sSQL & "Else "
        sSQL = sSQL & "RTChecks.[PrintOnIB] "
        sSQL = sSQL & "End "
        '------------------------------------END SQL SERVER DB-------------------------
        '
        '------------------------------------------------------------------------------
    Else
        '------------------------------------BEGIN Access DB-------------------------
        '
        '----------------------------------------------------------------------------
        sSQL = sSQL & "IIF(RTChecks.[PrintOnIB] Is Null, "
        'Then
        sSQL = sSQL & "0, "
        'else
        sSQL = sSQL & "RTChecks.[PrintOnIB]) "
        '------------------------------------END Access DB---------------------------
        '
        '----------------------------------------------------------------------------
    End If
    sSQL = sSQL & ")As bPrintOnIB, "
    sSQL = sSQL & "IB.[IB01_sSubToCarrier] As rt01_sSubToCarrier, "
    sSQL = sSQL & "( "
    If goUtil.gMainDB Is Nothing Then
        '------------------------------------BEGIN SQL SERVER DB-------------------------
        '
        '--------------------------------------------------------------------------------
        sSQL = sSQL & "CASE WHEN    IB.[IB14a_sSupplement] > 0 And IB.[IB14b_sRebilled] > 0 "
        sSQL = sSQL & "THEN IB.[IB02_sIBNumber] + 'S' + cast(IB.[IB14a_sSupplement] As varchar(4)) + 'R' + cast(IB.[IB14b_sRebilled] As varchar(4)) "
        sSQL = sSQL & "ELSE ( "
                    sSQL = sSQL & "CASE WHEN    IB.[IB14a_sSupplement] > 0 "
                    sSQL = sSQL & "THEN IB.[IB02_sIBNumber] + 'S' + cast(IB.[IB14a_sSupplement] As varchar(4)) "
                    sSQL = sSQL & "ELSE ( "
                                sSQL = sSQL & "CASE   WHEN    IB.[IB14b_sRebilled] > 0 "
                                sSQL = sSQL & "THEN IB.[IB02_sIBNumber] + 'R' + cast(IB.[IB14b_sRebilled] As varchar(4)) "
                                sSQL = sSQL & "ELSE IB.[IB02_sIBNumber] "
                                sSQL = sSQL & "END "
                    sSQL = sSQL & ") "
                    sSQL = sSQL & "END "
        sSQL = sSQL & ") "
        sSQL = sSQL & "END "
        '------------------------------------END SQL SERVER DB-------------------------
        '
        '------------------------------------------------------------------------------
    Else
        '------------------------------------BEGIN Access DB-------------------------
        '
        '----------------------------------------------------------------------------
        sSQL = sSQL & "IIF( IB.[IB14a_sSupplement] > 0 And IB.[IB14b_sRebilled] > 0, "
        'Then
        sSQL = sSQL & "IB.[IB02_sIBNumber] + 'S' + Cstr(IB.[IB14a_sSupplement]) + 'R' + Cstr(IB.[IB14b_sRebilled]), "
        'else
                    sSQL = sSQL & "IIF(IB.[IB14a_sSupplement] > 0, "
                    'Then
                    sSQL = sSQL & "IB.[IB02_sIBNumber] + 'S' + Cstr(IB.[IB14a_sSupplement]) , "
                    'else
                                sSQL = sSQL & "IIF(IB.[IB14b_sRebilled] > 0, "
                                'Then
                                sSQL = sSQL & "IB.[IB02_sIBNumber] + 'R' + Cstr(IB.[IB14b_sRebilled]) , "
                                'else
                                sSQL = sSQL & "IB.[IB02_sIBNumber]) "
                    sSQL = sSQL & ") "
        sSQL = sSQL & ") "
        '------------------------------------END Access DB---------------------------
        '
        '----------------------------------------------------------------------------
    End If
    sSQL = sSQL & ") As RT02_sIBNumber, "
    sSQL = sSQL & "CCCS.[CatCode] As rt04_sCatCode, "
    sSQL = sSQL & "IB.[IB05_sLocation] As rt05_sLocation, "
    sSQL = sSQL & "IB.[IB05a_sState] As rt05a_sState, "
    sSQL = sSQL & "IB.[IB06_dtDateClosed] As rt06_dtDateClosed, "
    sSQL = sSQL & "IB.[IB07_sAdjusterName] As rt07_sAdjusterName, "
    sSQL = sSQL & "CCAS.[ACID] As rt08_sAdjCRID, "
    sSQL = sSQL & "IB.[IB09_sSALN] As rt09_sSALN, "
    sSQL = sSQL & "IB.[IB09a_sPolicyNo] As rt09a_sPolicyNo, "
    sSQL = sSQL & "IB.[IB10_sInsuredName] As rt10_sInsuredName, "
    sSQL = sSQL & "IB.[IB11_sLossLocation] As rt11_sLossLocation, "
    sSQL = sSQL & "IB.[IB12_dtDateOfLoss] As rt12_dtDateOfLoss, "
    sSQL = sSQL & "IB.[IB13_cGrossLoss] As rt13_cGrossLoss, "
    sSQL = sSQL & "IB.[IB14_cDepreciation] As rt14_cDepreciation, "
    sSQL = sSQL & "IB.[IB14a_sSupplement] As rt14a_sSupplement, "
    sSQL = sSQL & "IB.[IB14b_sRebilled] As rt14b_sRebilled, "
    sSQL = sSQL & "IB.[IB15_cDeductible] As rt15_cDeductible, "
    sSQL = sSQL & "IB.[IB15a_cLessExcessLimits] As rt15a_cLessExcessLimits, "
    sSQL = sSQL & "IB.[IB15b_sExcessLimDesc] As rt15b_sExcessLimDesc, "
    sSQL = sSQL & "IB.[IB15c_cLessMiscellaneous] As rt15c_cLessMiscellaneous, "
    sSQL = sSQL & "IB.[IB15d_cMiscellaneousDesc] As rt15d_cMiscellaneousDesc, "
    sSQL = sSQL & "IB.[IB16_cNetClaim] As rt16_cNetClaim, "
    sSQL = sSQL & "IB.[IB17_cServiceFee] As rt17_cServiceFee, "
    sSQL = sSQL & "IB.[IB17a_cMiscServiceFee] As rt17a_cMiscServiceFee, "
    sSQL = sSQL & "IB.[IB18_sServiceFeeComment] As rt18_sServiceFeeComment, "
    sSQL = sSQL & "IB.[IB18a_sMiscServiceFeeComment] As rt18a_sMiscServiceFeeComment, "
    sSQL = sSQL & "IB.[IB25_cServiceFeeSubTotal] As rt25_cServiceFeeSubTotal, "
    sSQL = sSQL & "IB.[IB29a_sMiscExpenseFeeComment] As rt29a_sMiscFeeComment, "
    sSQL = sSQL & "IB.[IB29b_cMiscExpenseFee] As rt29b_cMiscFee, "
    sSQL = sSQL & "IB.[IB30_cTotalExpenses] As rt30_cTotalExpenses, "
    sSQL = sSQL & "IB.[IB31_dTaxPercent] As rt31_dTaxPercent, "
    sSQL = sSQL & "IB.[IB32_cTaxAmount] As rt32_cTaxAmount, "
    sSQL = sSQL & "IB.[IB33_cTotalAdjustingFee] As rt33_cTotalAdjustingFee, "
    sSQL = sSQL & "IB.[Comments] As Comments, "
    sSQL = sSQL & "IB.[FeeScheduleID], "
    sSQL = sSQL & "( "
    sSQL = sSQL & "SELECT  Code "
    sSQL = sSQL & "From ClassOfLoss "
    sSQL = sSQL & "Where ClassOfLossID = RTChecks.[RT42_ClassOfLossID] "
    sSQL = sSQL & ") As RT42_sClassOfClaim, "
    sSQL = sSQL & "( "
    sSQL = sSQL & "SELECT  Code "
    sSQL = sSQL & "From TypeOfLoss "
    sSQL = sSQL & "Where TypeOfLossID = RTChecks.[RT43_TypeOfLossID] "
    sSQL = sSQL & ") As RT43_sCauseOfLoss, "
    sSQL = sSQL & "RTChecks.[RT50_sInsuredPayeeName] As rt50_sInsuredPayeeName, "
    sSQL = sSQL & "RTChecks.[RT51_sPayeeNames] As rt51_sPayeeNames, "
    sSQL = sSQL & "RTChecks.[RT52_sAddress] As rt52_sAddress, "
    sSQL = sSQL & "RTChecks.[RT53_cAmountOfCheck] As rt53_cAmountOfCheck, "
    sSQL = sSQL & "RTChecks.[RTChecksID], "
    sSQL = sSQL & "RTChecks.[CheckNum], "
    sSQL = sSQL & "CCCS.[CatCode] As rt54_sCatCode "
    If msIBHistoryID <> vbNullString Then
        sSQL = sSQL & "From ((((IBHistory IB "
    ElseIf msIBID <> vbNullString Then
        sSQL = sSQL & "From ((((IB "
    ElseIf msIBHistoryID = vbNullString And msIBID = vbNullString Then
        'If getting from the RTIB Table (Real Time Billing The Bill before it moves to IB Table)
        'Need to update all the IB.[IB  to IB.[RT
        sSQL = Replace(sSQL, "IB.[IB", "IB.[RT", , , vbTextCompare)
        sSQL = sSQL & "From ((((RTIB IB "
    End If
    sSQL = sSQL & "INNER JOIN Assignments A ON IB.AssignmentsID = A.AssignmentsID) "
    sSQL = sSQL & "INNER JOIN ClientCompanyCatSpec CCCS ON CCCS.ClientCompanyCatSpecID = A.ClientCompanyCatSpecID) "
    sSQL = sSQL & "INNER JOIN ClientCoAdjusterSpec CCAS ON CCAS.ClientCoAdjusterSpecID = A.AdjusterSpecIDDisplay) "
    sSQL = sSQL & "LEFT JOIN RTChecks ON (IB.AssignmentsID = RTChecks.AssignmentsID And IB.BillingCountID = RTChecks.BillingCountID And RTChecks.CheckNum = 1) ) "
    
    sSQL = sSQL & "WHERE IB.AssignmentsID  = " & msAssignmentsID & " "
    If msIBHistoryID <> vbNullString Then
        sSQL = sSQL & "AND IB.IBHistoryID = " & msIBHistoryID & " "
    ElseIf msIBID <> vbNullString Then
        sSQL = sSQL & "AND IB.IBID = " & msIBID & " "
    End If
            
    
    Set mRS = New ADODB.Recordset
    mRS.CursorLocation = adUseClient
    mRS.Open sSQL, mConn, adOpenForwardOnly, adLockReadOnly
    Set mRS.ActiveConnection = Nothing
    
    If Not mRS.EOF Then
        mRS.MoveFirst
        With PopulateudtarRpt
            On Error Resume Next
            .bPrintOnIB = mRS!bPrintOnIB
            .CopyName = msCopyName
            .Comments = mRS!Comments
            sFeeScheduleID = mRS!FeeScheduleID
            .f01_sSubToCarrier = mRS!rt01_sSubToCarrier
            .f02_sIBNumber = mRS!RT02_sIBNumber
            .f04_sCatCode = mRS!rt04_sCatCode
            .f05_sLocation = mRS!rt05_sLocation
            .f05a_sState = mRS!rt05a_sState
            .f06_dtDateClosed = mRS.Fields("rt06_dtDateClosed")
            .f07_sAdjusterName = mRS!rt07_sAdjusterName
            .f08_sAdjCRID = mRS!rt08_sAdjCRID
            .f09_sSALN = mRS!rt09_sSALN
            .f09a_sPolicyNo = mRS!rt09a_sPolicyNo
            .f10_sInsuredName = mRS!rt10_sInsuredName
            .f11_sLossLocation = mRS!rt11_sLossLocation
            .f12_dtDateOfLoss = mRS!rt12_dtDateOfLoss
            .f13_cGrossLoss = mRS!rt13_cGrossLoss
            .f14_cDepreciation = mRS!rt14_cDepreciation
            .f14a_sSupplement = mRS!rt14a_sSupplement
            .f14b_sRebilled = mRS!rt14b_sRebilled
            .f15_cDeductible = mRS!rt15_cDeductible
            .f15a_cLessExcessLimits = mRS!rt15a_cLessExcessLimits
            .f15b_sExcessLimDesc = mRS!rt15b_sExcessLimDesc
            .f15c_cLessMiscellaneous = mRS!rt15c_cLessMiscellaneous
            .f15d_cMiscellaneousDesc = mRS!rt15d_cMiscellaneousDesc
            .f16_cNetClaim = mRS!rt16_cNetClaim
            .f17_cServiceFee = mRS!rt17_cServiceFee
            .f17a_cMiscServiceFee = mRS!rt17a_cMiscServiceFee
            .f18_sServiceFeeComment = mRS!rt18_sServiceFeeComment
            .f18a_sMiscServiceFeeComment = mRS!rt18a_sMiscServiceFeeComment
            .f25_cServiceFeeSubTotal = mRS!rt25_cServiceFeeSubTotal
            .f29a_sMiscFeeComment = mRS!rt29a_sMiscFeeComment
            .f29b_cMiscFee = mRS!rt29b_cMiscFee
            .f30_cTotalExpenses = mRS!rt30_cTotalExpenses
            .f31_dTaxPercent = mRS!rt31_dTaxPercent
            .f32_cTaxAmount = mRS!rt32_cTaxAmount
            .f33_cTotalAdjustingFee = mRS!rt33_cTotalAdjustingFee
            .f42_ClassOfClaim = mRS!RT42_sClassOfClaim
            .f43_sCauseOfLoss = mRS!RT43_sCauseOfLoss
            .f50_sInsuredPayeeName = mRS!rt50_sInsuredPayeeName
            .f51_sPayeeNames = mRS!rt51_sPayeeNames
            .f52_sAddress = mRS!rt52_sAddress
            .f53_cAmountOfCheck = mRS!rt53_cAmountOfCheck
            .f54_sCatCode = mRS!rt54_sCatCode
            On Error GoTo EH
        End With
        
        
        'Need to Get the Params that are associated with the Check Num 1
        'if it is Flaged to Print On IB
        If mRS!bPrintOnIB And Not IsNull(mRS!RTChecksID) And Not IsNull(mRS!CheckNum) Then
            'Need to get the SPVersion from the Assignments Table
            sSQL = "SELECT SPVersion "
            sSQL = sSQL & "FROM Assignments "
            sSQL = sSQL & "WHERE AssignmentsID = " & msAssignmentsID & " "
            
            Set RSSPVersion = New ADODB.Recordset
            RSSPVersion.CursorLocation = adUseClient
            RSSPVersion.Open sSQL, mConn, adOpenForwardOnly, adLockReadOnly
            Set RSSPVersion.ActiveConnection = Nothing
            
            If Not RSSPVersion.EOF Then
                lMainSPVerison = RSSPVersion.Fields("SPVersion").Value
            End If
            larRptAddlChkVersion = goUtil.GetApplicationVersionNumber(lMainSPVerison, "ECRptStateFarm_arRptAddlChk", mConn)
            
            sParams = sParams & "psAssignmentsID=" & msAssignmentsID & "|"
            sParams = sParams & "pRTChecksID=" & mRS.Fields("RTChecksID").Value & "|"
            sParams = sParams & "pbPreview=" & "False" & "|"
            sParams = sParams & "pbGetObjectOnly=" & "True" & "|"
            sParams = sParams & "psCheckNum=" & mRS.Fields("CheckNum").Value & "|"
            
            'Create the Carlist for ABC
            Set oCarList = CreateObject("V2ECcarABC.clsLists")
            Set colParams = oCarList.GetARMiscDelimParamsCol("ECRptStateFarm_arRptAddlChk.clsLists", larRptAddlChkVersion, sParams)
            
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
        End If
        
        'Populate the Service and Expense Fees HERE !!!
        'Depending upon wethere this IB is closed or Current Will be
        'Pulling from IBfee (Closed) or RTIBFee (Current)
        If msIBID <> vbNullString Then
            sSQL = GetClosedIBFeeSQL(msAssignmentsID, msIBID, sFeeScheduleID)
        ElseIf msIBHistoryID = vbNullString And msIBID = vbNullString Then
            sSQL = GetOpenIBFeeSQL(msAssignmentsID, sFeeScheduleID)
        Else
            'No history on Fee Items At this time!
            sSQL = vbNullString
        End If
        
        If sSQL = vbNullString Then
            GoTo CLEAN_UP
        End If
        
        Set RSFeeItems = New ADODB.Recordset
        RSFeeItems.CursorLocation = adUseClient
        RSFeeItems.Open sSQL, mConn, adOpenForwardOnly, adLockReadOnly
        Set RSFeeItems.ActiveConnection = Nothing
        
        If RSFeeItems.RecordCount = 0 Then
            GoTo CLEAN_UP
        End If
        
        RSFeeItems.MoveFirst
        Set PopulateudtarRpt.coludtFeeItems01 = New Collection

        Do Until RSFeeItems.EOF
            With MyFeeItem
                .Amount = goUtil.IsNullIsVbNullString(RSFeeItems.Fields("Amount"))
                .Comment = goUtil.IsNullIsVbNullString(RSFeeItems.Fields("Comment"))
                .fsftDescription = goUtil.IsNullIsVbNullString(RSFeeItems.Fields("fsftDescription"))
                .fsftFeeAmount = goUtil.IsNullIsVbNullString(RSFeeItems.Fields("fsftFeeAmount"))
                .fsftIsExpense = goUtil.IsNullIsVbNullString(RSFeeItems.Fields("fsftIsExpense"))
                .fsftMaxNumberOfItems = goUtil.IsNullIsVbNullString(RSFeeItems.Fields("fsftMaxNumberOfItems"))
                .fsftMaxFeeAmount = goUtil.IsNullIsVbNullString(RSFeeItems.Fields("fsftMaxFeeAmount"))
                .fsftIsMiscAmount = goUtil.IsNullIsVbNullString(RSFeeItems.Fields("fsftIsMiscAmount"))
                .fsftUseFormula = goUtil.IsNullIsVbNullString(RSFeeItems.Fields("fsftUseFormula"))
                .fsftVBFormula = goUtil.IsNullIsVbNullString(RSFeeItems.Fields("fsftVBFormula"))
                .fsftName = goUtil.IsNullIsVbNullString(RSFeeItems.Fields("fsftName"))
                .fsftTypeNum = goUtil.IsNullIsVbNullString(RSFeeItems.Fields("fsftTypeNum"))
                .NumberOfItems = goUtil.IsNullIsVbNullString(RSFeeItems.Fields("NumberOfItems"))
            End With
            
            sSevCode = vbNullString
            If InStr(1, MyFeeItem.fsftVBFormula, "OVERRIDES_", vbTextCompare) > 0 Then
                bAddFeeItem = False
            ElseIf StrComp(Left(MyFeeItem.fsftName, 4), "SEV_", vbTextCompare) = 0 Then
                'CHeck for Severity Codes
                sSevCode = Mid(MyFeeItem.fsftName, 7, Len(MyFeeItem.fsftName))
                'If the amount of this Sev code is > 0 then Add it
                If MyFeeItem.Amount > 0 Then
                    bAddFeeItem = True
                    sAddedSevCodes = sAddedSevCodes & sSevCode & vbCrLf
                ElseIf StrComp(Left(MyFeeItem.fsftName, 6), "SEV_1_", vbTextCompare) = 0 Then
                    'If this is the SEV_1 field check for previous Adding of lower severity
                    'if found do not add Sev 1
                    If InStr(1, sAddedSevCodes, sSevCode & vbCrLf, vbTextCompare) = 0 Then
                        'Change the Desc to Take Out the Sev(01)
                        MyFeeItem.fsftDescription = Replace(MyFeeItem.fsftDescription, "Sev(1) -", vbNullString, , , vbTextCompare)
                        bAddFeeItem = True
                    Else
                        bAddFeeItem = False
                    End If
                Else
                    bAddFeeItem = False
                End If
                
            Else
                bAddFeeItem = True
            End If
            
            If bAddFeeItem Then
                PopulateudtarRpt.coludtFeeItems01.Add MyFeeItem, """" & CStr(MyFeeItem.fsftTypeNum) & """"
            End If
            
            If sSevCode <> vbNullString Then
                sPrevSevCode = sSevCode
            End If

            RSFeeItems.MoveNext
        Loop

    End If
CLEAN_UP:
    'Cleanup
    Set oCarList = Nothing
    Set colParams = Nothing
    Set mRS = Nothing
    Set RSSPVersion = Nothing
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Set RS = Nothing
    Set mRS = Nothing
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsRptDB_ClassName & vbCrLf & "Private Function PopulateudtarRpt"
End Function

Private Function GetClosedIBFeeSQL(psAssignmentsID As String, psIBID As String, psFeeScheduleID As String) As String
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim sSQL As String
   
    
    sSQL = "SELECT FSFT.[TypeNum] As [fsftTypeNum], "
    sSQL = sSQL & "FSFT.[Name] As [fsftName], "
    sSQL = sSQL & "FSFT.[Description] As [fsftDescription], "
    sSQL = sSQL & "FSFT.[FeeAmount] As [fsftFeeAmount], "
    sSQL = sSQL & "FSFT.[IsExpense] As [fsftIsExpense], "
    sSQL = sSQL & "FSFT.[MaxNumberOfItems] As [fsftMaxNumberOfItems], "
    sSQL = sSQL & "FSFT.[MaxFeeAmount] As [fsftMaxFeeAmount], "
    sSQL = sSQL & "FSFT.[IsMiscAmount] As [fsftIsMiscAmount], "
    sSQL = sSQL & "FSFT.[UseFormula] As [fsftUseFormula], "
    sSQL = sSQL & "FSFT.[VBFormula] As [fsftVBFormula], "
    sSQL = sSQL & "IBFee.[NumberOfItems], "
    sSQL = sSQL & "IBFee.[Amount], "
    sSQL = sSQL & "IBFee.[Comment] "
    sSQL = sSQL & "FROM IBFee INNER JOIN FeeScheduleFeeTypes FSFT ON IBFee.[FeeScheduleFeeTypesID] = FSFT.[FeeScheduleFeeTypesID] "
    sSQL = sSQL & "WHERE IBFee.[AssignmentsID] = " & psAssignmentsID & " "
    sSQL = sSQL & "AND IBFee.[IBID] = " & psIBID & " "
    sSQL = sSQL & "AND FSFT.[FeeScheduleID] = " & psFeeScheduleID & " "
'    sSQL = sSQL & "AND FSFT.[FeeScheduleID] = ( "
'                            sSQL = sSQL & "SELECT   [FeeScheduleID] "
'                            sSQL = sSQL & "FROM     ClientCompanyCat "
'                            sSQL = sSQL & "WHERE    [ClientCompanyID] = ( "
'                                                        sSQL = sSQL & "SELECT   [ClientCompanyID] "
'                                                        sSQL = sSQL & "FROM     ClientCompanyCatSpec "
'                                                        sSQL = sSQL & "WHERE    [ClientCompanyCatSpecID] = ( "
'                                                                                            sSQL = sSQL & "SELECT [ClientCompanyCatSpecID] "
'                                                                                            sSQL = sSQL & "FROM     Assignments "
'                                                                                            sSQL = sSQL & "WHERE    AssignmentsID = " & psAssignmentsID & " "
'                                                                                            sSQL = sSQL & ") "
'                                                        sSQL = sSQL & ") "
'                            sSQL = sSQL & "AND      [CATID] = ( "
'                                                sSQL = sSQL & "SELECT [CATID] "
'                                                sSQL = sSQL & "FROM ClientCompanyCatSpec "
'                                                sSQL = sSQL & "WHERE    [ClientCompanyCatSpecID] = ( "
'                                                                                    sSQL = sSQL & "SELECT [ClientCompanyCatSpecID] "
'                                                                                    sSQL = sSQL & "FROM     Assignments "
'                                                                                    sSQL = sSQL & "WHERE    AssignmentsID = " & psAssignmentsID & " "
'                                                                                    sSQL = sSQL & ") "
'                                                sSQL = sSQL & ") "
'                            sSQL = sSQL & ") "
    sSQL = sSQL & "ORDER BY  FSFT.[TypeNum] "
    
    GetClosedIBFeeSQL = sSQL
    
Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Set RS = Nothing
    Set mRS = Nothing
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsRptDB_ClassName & vbCrLf & "Private Function GetClosedIBFeeSQL"
End Function

Private Function GetOpenIBFeeSQL(psAssignmentsID As String, psFeeScheduleID As String) As String
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim sSQL As String
    
    sSQL = "SELECT FSFT.[TypeNum] As [fsftTypeNum], "
    sSQL = sSQL & "FSFT.[Name] As [fsftName], "
    sSQL = sSQL & "FSFT.[Description] As [fsftDescription], "
    sSQL = sSQL & "FSFT.[FeeAmount] As [fsftFeeAmount], "
    sSQL = sSQL & "FSFT.[IsExpense] As [fsftIsExpense], "
    sSQL = sSQL & "FSFT.[MaxNumberOfItems] As [fsftMaxNumberOfItems], "
    sSQL = sSQL & "FSFT.[MaxFeeAmount] As [fsftMaxFeeAmount], "
    sSQL = sSQL & "FSFT.[IsMiscAmount] As [fsftIsMiscAmount], "
    sSQL = sSQL & "FSFT.[UseFormula] As [fsftUseFormula], "
    sSQL = sSQL & "FSFT.[VBFormula] As [fsftVBFormula], "
    sSQL = sSQL & "RTIBFee.[NumberOfItems], "
    sSQL = sSQL & "RTIBFee.[Amount], "
    sSQL = sSQL & "RTIBFee.[Comment] "
    sSQL = sSQL & "FROM RTIBFee INNER JOIN FeeScheduleFeeTypes FSFT ON RTIBFee.[FeeScheduleFeeTypesID] = FSFT.[FeeScheduleFeeTypesID] "
    sSQL = sSQL & "WHERE RTIBFee.[AssignmentsID] = " & psAssignmentsID & " "
    sSQL = sSQL & "AND FSFT.[FeeScheduleID] = " & psFeeScheduleID & " "
'    sSQL = sSQL & "AND FSFT.[FeeScheduleID] = ( "
'                            sSQL = sSQL & "SELECT   [FeeScheduleID] "
'                            sSQL = sSQL & "FROM     ClientCompanyCat "
'                            sSQL = sSQL & "WHERE    [ClientCompanyID] = ( "
'                                                        sSQL = sSQL & "SELECT   [ClientCompanyID] "
'                                                        sSQL = sSQL & "FROM     ClientCompanyCatSpec "
'                                                        sSQL = sSQL & "WHERE    [ClientCompanyCatSpecID] = ( "
'                                                                                            sSQL = sSQL & "SELECT [ClientCompanyCatSpecID] "
'                                                                                            sSQL = sSQL & "FROM     Assignments "
'                                                                                            sSQL = sSQL & "WHERE    AssignmentsID = " & psAssignmentsID & " "
'                                                                                            sSQL = sSQL & ") "
'                                                        sSQL = sSQL & ") "
'                            sSQL = sSQL & "AND      [CATID] = ( "
'                                                sSQL = sSQL & "SELECT [CATID] "
'                                                sSQL = sSQL & "FROM ClientCompanyCatSpec "
'                                                sSQL = sSQL & "WHERE    [ClientCompanyCatSpecID] = ( "
'                                                                                    sSQL = sSQL & "SELECT [ClientCompanyCatSpecID] "
'                                                                                    sSQL = sSQL & "FROM     Assignments "
'                                                                                    sSQL = sSQL & "WHERE    AssignmentsID = " & psAssignmentsID & " "
'                                                                                    sSQL = sSQL & ") "
'                                                sSQL = sSQL & ") "
'                            sSQL = sSQL & ") "
    sSQL = sSQL & "ORDER BY FSFT.[TypeNum] "
    
    GetOpenIBFeeSQL = sSQL
    
Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Set RS = Nothing
    Set mRS = Nothing
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsRptDB_ClassName & vbCrLf & "Private Function GetOpenIBFeeSQL"
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
    sMiscReportParamName = "MiscReportParam07"
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
    'IB needs to get the Number that matches the Supplement
    sSQL = sSQL & "AND Number = " & msSupplement & " "
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
                    .Number = msSupplement
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
