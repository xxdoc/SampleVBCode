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
Private marRpt As ECRptABC_arActivityLog.arActivityLog

Public Type udtActivityLog01
    f_RTActivityLogID As Long
    f_ActDate As Date
    f_ActTime As Date
    f_ServiceTime As Double
    f_IB14a_sSupplement As String
    f_ActText As String
    f_PageBreakAfter As Boolean
    f_BlankPageAfter As Boolean
    f_BlankRowsAfter As Long
End Type

Private Type udtarRpt
    f_Insured As String
    f_ClientClaimNo As String
    f_ADJNAME As String
    f_ReceivedDate As Date
    f_ContactDate As Date
    f_InspectedDate As Date
    f_AL01_sPresentDurringInspectionY As String
    f_AL01_sPresentDurringInspectionN As String
    f_AL02_sExplainedEstimateY As String
    f_AL02_sExplainedEstimateN As String
    f_AL03_sExplainedRCVY As String
    f_AL03_sExplainedRCVN As String
    f_AL03_sExplainedRCVNA As String
    f_AL04_sConfirmMortgageeIsCorrectY As String
    f_AL04_sConfirmMortgageeIsCorrectN As String
    f_AL04_sConfirmMortgageeIsCorrectNA As String
    f_AL05_sExplainedMortgageeChecksY As String
    f_AL05_sExplainedMortgageeChecksN As String
    f_AL05_sExplainedMortgageeChecksNA As String
    f_AL06_sConfirmedCoverageY As String
    f_AL06_sConfirmedCoverageN As String
    f_AL07_sPriorLossY As String
    f_AL07_sPriorLossN As String
    f_AL07_sPriorLossNA As String
    f_AL08_sSalvageYesNo As String
    f_AL09_sSubrogationYesNo As String
    f_ServiceTimeSum As Double
    f_CloseDate As Date
    f_IBNUM As String
    PrintBlankReport As Boolean
    coludtActivityLog01 As Collection
End Type

Private mbPrintBlankReport As Boolean
Private mbExportXML As Boolean
Private mbExportXMLOnly As Boolean

Public Property Let PrintBlankReport(pbFlag As Boolean)
    mbPrintBlankReport = pbFlag
End Property
Public Property Get PrintBlankReport() As Boolean
    PrintBlankReport = mbPrintBlankReport
End Property

Private Property Get clsRptDB_MiscDelimParamsCol() As Collection
    Dim bCleanup As Boolean
    'Add Report Parameters that are not updated by DB Fields.
    If marRpt Is Nothing Then
        bCleanup = True
        Set marRpt = New ECRptABC_arActivityLog.arActivityLog
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
'    Set mCurDB = poDB
End Property
Private Property Let clsRptDB_CurDB(poDB As DAO.Database)
'    Set mCurDB = poDB
End Property
Private Property Get clsRptDB_CurDB() As DAO.Database
'    Set clsRptDB_CurDB = mCurDB
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
    If Not mRS Is Nothing Then
        mRS.Close
        Set mRS = Nothing
    End If
    If Not mConn Is Nothing Then
        mConn.Close
        Set mConn = Nothing
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
    
    Set marRpt = New ECRptABC_arActivityLog.arActivityLog
    
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
        marRpt.SetProperty "f_ReceivedDate", .f_ReceivedDate, vbDate
        marRpt.SetProperty "f_ContactDate", .f_ContactDate, vbDate
        marRpt.SetProperty "f_InspectedDate", .f_InspectedDate, vbDate
        marRpt.SetProperty "f_AL01_sPresentDurringInspectionY", .f_AL01_sPresentDurringInspectionY, vbString
        marRpt.SetProperty "f_AL01_sPresentDurringInspectionN", .f_AL01_sPresentDurringInspectionN, vbString
        marRpt.SetProperty "f_AL02_sExplainedEstimateY", .f_AL02_sExplainedEstimateY, vbString
        marRpt.SetProperty "f_AL02_sExplainedEstimateN", .f_AL02_sExplainedEstimateN, vbString
        marRpt.SetProperty "f_AL03_sExplainedRCVY", .f_AL03_sExplainedRCVY, vbString
        marRpt.SetProperty "f_AL03_sExplainedRCVN", .f_AL03_sExplainedRCVN, vbString
        marRpt.SetProperty "f_AL03_sExplainedRCVNA", .f_AL03_sExplainedRCVNA, vbString
        marRpt.SetProperty "f_AL04_sConfirmMortgageeIsCorrectY", .f_AL04_sConfirmMortgageeIsCorrectY, vbString
        marRpt.SetProperty "f_AL04_sConfirmMortgageeIsCorrectN", .f_AL04_sConfirmMortgageeIsCorrectN, vbString
        marRpt.SetProperty "f_AL04_sConfirmMortgageeIsCorrectNA", .f_AL04_sConfirmMortgageeIsCorrectNA, vbString
        marRpt.SetProperty "f_AL05_sExplainedMortgageeChecksY", .f_AL05_sExplainedMortgageeChecksY, vbString
        marRpt.SetProperty "f_AL05_sExplainedMortgageeChecksN", .f_AL05_sExplainedMortgageeChecksN, vbString
        marRpt.SetProperty "f_AL05_sExplainedMortgageeChecksNA", .f_AL05_sExplainedMortgageeChecksNA, vbString
        marRpt.SetProperty "f_AL06_sConfirmedCoverageY", .f_AL06_sConfirmedCoverageY, vbString
        marRpt.SetProperty "f_AL06_sConfirmedCoverageN", .f_AL06_sConfirmedCoverageN, vbString
        marRpt.SetProperty "f_AL07_sPriorLossY", .f_AL07_sPriorLossY, vbString
        marRpt.SetProperty "f_AL07_sPriorLossN", .f_AL07_sPriorLossN, vbString
        marRpt.SetProperty "f_AL07_sPriorLossNA", .f_AL07_sPriorLossNA, vbString
        marRpt.SetProperty "f_AL08_sSalvageYesNo", .f_AL08_sSalvageYesNo, vbString
        marRpt.SetProperty "f_AL09_sSubrogationYesNo", .f_AL09_sSubrogationYesNo, vbString
        marRpt.SetProperty "f_ServiceTimeSum", .f_ServiceTimeSum, vbDouble
        marRpt.SetProperty "f_CloseDate", .f_CloseDate, vbDate
        marRpt.SetProperty "f_IBNUM", .f_IBNUM, vbString
        marRpt.SetProperty "PrintBlankReport", .PrintBlankReport, vbBoolean
        marRpt.SetProperty "coludtActivityLog01", .coludtActivityLog01, vbObject
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
    Dim MyActLog As udtActivityLog01
    
    'For Blank Activity Log
    Dim lCount As Long
    
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

    'B.  get the Activity Log Header Info
    sSQL = "SELECT "
    sSQL = sSQL & "A.Insured As Insured, "
    sSQL = sSQL & "A.CLIENTNUM As ClientClaimNo, "
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
    sSQL = sSQL & "A.ReceivedDate As ReceivedDate, "
    sSQL = sSQL & "A.ContactDate As ContactDate, "
    sSQL = sSQL & "A.InspectedDate As InspectedDate, "
    sSQL = sSQL & "A.CloseDate As CloseDate, "
    sSQL = sSQL & "A.IBNUM as IBNUM, "
    If goUtil.gMainDB Is Nothing Then
        '------------------------------------BEGIN SQL SERVER DB-------------------------
        '
        '--------------------------------------------------------------------------------
        'AL01_sPresentDurringInspection
        sSQL = sSQL & "( "
        sSQL = sSQL & "CASE WHEN RTALI.AL01_sPresentDurringInspection = 1 "
        sSQL = sSQL & "THEN 'X' "
        sSQL = sSQL & "ELSE '' "
        sSQL = sSQL & "END "
        sSQL = sSQL & ") As AL01_sPresentDurringInspectionY, "
        sSQL = sSQL & "( "
        sSQL = sSQL & "CASE WHEN RTALI.AL01_sPresentDurringInspection = 0 "
        sSQL = sSQL & "THEN 'X' "
        sSQL = sSQL & "ELSE '' "
        sSQL = sSQL & "END "
        sSQL = sSQL & ") As AL01_sPresentDurringInspectionN, "
        'AL02_sExplainedEstimate
        sSQL = sSQL & "( "
        sSQL = sSQL & "CASE WHEN RTALI.AL02_sExplainedEstimate = 1 "
        sSQL = sSQL & "THEN 'X' "
        sSQL = sSQL & "ELSE '' "
        sSQL = sSQL & "END "
        sSQL = sSQL & ") As AL02_sExplainedEstimateY, "
        sSQL = sSQL & "( "
        sSQL = sSQL & "CASE WHEN RTALI.AL02_sExplainedEstimate = 0 "
        sSQL = sSQL & "THEN 'X' "
        sSQL = sSQL & "ELSE '' "
        sSQL = sSQL & "END "
        sSQL = sSQL & ") As AL02_sExplainedEstimateN, "
        'AL03_sExplainedRCV
        sSQL = sSQL & "( "
        sSQL = sSQL & "CASE WHEN RTALI.AL03_sExplainedRCV = 1 "
        sSQL = sSQL & "THEN 'X' "
        sSQL = sSQL & "ELSE '' "
        sSQL = sSQL & "END "
        sSQL = sSQL & ") As AL03_sExplainedRCVY, "
        sSQL = sSQL & "( "
        sSQL = sSQL & "CASE WHEN RTALI.AL03_sExplainedRCV = 0 And  RTALI.AL03_sExplainedRCVNA = 0 "
        sSQL = sSQL & "THEN 'X' "
        sSQL = sSQL & "ELSE '' "
        sSQL = sSQL & "END "
        sSQL = sSQL & ") As AL03_sExplainedRCVN, "
        sSQL = sSQL & "( "
        sSQL = sSQL & "CASE WHEN RTALI.AL03_sExplainedRCVNA = 1 "
        sSQL = sSQL & "THEN 'X' "
        sSQL = sSQL & "ELSE '' "
        sSQL = sSQL & "END "
        sSQL = sSQL & ") As AL03_sExplainedRCVNA, "
        'AL04_sConfirmMortgageeIsCorrect
        sSQL = sSQL & "( "
        sSQL = sSQL & "CASE WHEN RTALI.AL04_sConfirmMortgageeIsCorrect = 1 "
        sSQL = sSQL & "THEN 'X' "
        sSQL = sSQL & "ELSE '' "
        sSQL = sSQL & "END "
        sSQL = sSQL & ") As AL04_sConfirmMortgageeIsCorrectY, "
        sSQL = sSQL & "( "
        sSQL = sSQL & "CASE WHEN RTALI.AL04_sConfirmMortgageeIsCorrect = 0 And RTALI.AL04_sConfirmMortgageeIsCorrectNA = 0 "
        sSQL = sSQL & "THEN 'X' "
        sSQL = sSQL & "ELSE '' "
        sSQL = sSQL & "END "
        sSQL = sSQL & ") As AL04_sConfirmMortgageeIsCorrectN, "
        sSQL = sSQL & "( "
        sSQL = sSQL & "CASE WHEN RTALI.AL04_sConfirmMortgageeIsCorrectNA = 1 "
        sSQL = sSQL & "THEN 'X' "
        sSQL = sSQL & "ELSE '' "
        sSQL = sSQL & "END "
        sSQL = sSQL & ") As AL04_sConfirmMortgageeIsCorrectNA, "
        'AL05_sExplainedMortgageeChecksY
        sSQL = sSQL & "( "
        sSQL = sSQL & "CASE WHEN RTALI.AL05_sExplainedMortgageeChecks = 1 "
        sSQL = sSQL & "THEN 'X' "
        sSQL = sSQL & "ELSE '' "
        sSQL = sSQL & "END "
        sSQL = sSQL & ") As AL05_sExplainedMortgageeChecksY, "
        sSQL = sSQL & "( "
        sSQL = sSQL & "CASE WHEN RTALI.AL05_sExplainedMortgageeChecks = 0 And RTALI.AL05_sExplainedMortgageeChecksNA = 0 "
        sSQL = sSQL & "THEN 'X' "
        sSQL = sSQL & "ELSE '' "
        sSQL = sSQL & "END "
        sSQL = sSQL & ") As AL05_sExplainedMortgageeChecksN, "
        sSQL = sSQL & "( "
        sSQL = sSQL & "CASE WHEN RTALI.AL05_sExplainedMortgageeChecksNA = 1 "
        sSQL = sSQL & "THEN 'X' "
        sSQL = sSQL & "ELSE '' "
        sSQL = sSQL & "END "
        sSQL = sSQL & ") As AL05_sExplainedMortgageeChecksNA, "
        'AL06_sConfirmedCoverageY
        sSQL = sSQL & "( "
        sSQL = sSQL & "CASE WHEN RTALI.AL06_sConfirmedCoverage = 1 "
        sSQL = sSQL & "THEN 'X' "
        sSQL = sSQL & "ELSE '' "
        sSQL = sSQL & "END "
        sSQL = sSQL & ") As AL06_sConfirmedCoverageY, "
        sSQL = sSQL & "( "
        sSQL = sSQL & "CASE WHEN RTALI.AL06_sConfirmedCoverage = 0 "
        sSQL = sSQL & "THEN 'X' "
        sSQL = sSQL & "ELSE '' "
        sSQL = sSQL & "END "
        sSQL = sSQL & ") As AL06_sConfirmedCoverageN, "
        'AL07_sPriorLossY
        sSQL = sSQL & "( "
        sSQL = sSQL & "CASE WHEN RTALI.AL07_sPriorLoss = 1 "
        sSQL = sSQL & "THEN 'X' "
        sSQL = sSQL & "ELSE '' "
        sSQL = sSQL & "END "
        sSQL = sSQL & ") As AL07_sPriorLossY, "
        sSQL = sSQL & "( "
        sSQL = sSQL & "CASE WHEN RTALI.AL07_sPriorLoss = 0 And RTALI.AL07_sPriorLossNA = 0 "
        sSQL = sSQL & "THEN 'X' "
        sSQL = sSQL & "ELSE '' "
        sSQL = sSQL & "END "
        sSQL = sSQL & ") As AL07_sPriorLossN, "
        sSQL = sSQL & "( "
        sSQL = sSQL & "CASE WHEN RTALI.AL07_sPriorLossNA = 1 "
        sSQL = sSQL & "THEN 'X' "
        sSQL = sSQL & "ELSE '' "
        sSQL = sSQL & "END "
        sSQL = sSQL & ") As AL07_sPriorLossNA, "
        'AL08_sSalvage
        sSQL = sSQL & "( "
        sSQL = sSQL & "CASE WHEN RTALI.AL08_sSalvage = 1 "
        sSQL = sSQL & "THEN 'YES' "
        sSQL = sSQL & "ELSE 'NO' "
        sSQL = sSQL & "END "
        sSQL = sSQL & ") As AL08_sSalvageYesNo, "
        'AL09_sSubrogation
        sSQL = sSQL & "( "
        sSQL = sSQL & "CASE WHEN RTALI.AL09_sSubrogation = 1 "
        sSQL = sSQL & "THEN 'YES' "
        sSQL = sSQL & "ELSE 'NO' "
        sSQL = sSQL & "END "
        sSQL = sSQL & ") As AL09_sSubrogationYesNo, "
        
        '------------------------------------END SQL SERVER DB-------------------------
        '
        '------------------------------------------------------------------------------
    Else
        '------------------------------------BEGIN Access DB-------------------------
        '
        '----------------------------------------------------------------------------
                'AL01_sPresentDurringInspection
        sSQL = sSQL & "( "
        sSQL = sSQL & "IIF(RTALI.AL01_sPresentDurringInspection, 'X', '') "
        sSQL = sSQL & ") As AL01_sPresentDurringInspectionY, "
        sSQL = sSQL & "( "
        sSQL = sSQL & "IIF(Not RTALI.AL01_sPresentDurringInspection, 'X', '') "
        sSQL = sSQL & ") As AL01_sPresentDurringInspectionN, "
        'AL02_sExplainedEstimate
        sSQL = sSQL & "( "
        sSQL = sSQL & "IIF(RTALI.AL02_sExplainedEstimate, 'X', '') "
        sSQL = sSQL & ") As AL02_sExplainedEstimateY, "
        sSQL = sSQL & "( "
        sSQL = sSQL & "IIF(Not RTALI.AL02_sExplainedEstimate, 'X', '') "
        sSQL = sSQL & ") As AL02_sExplainedEstimateN, "
        'AL03_sExplainedRCV
        sSQL = sSQL & "( "
        sSQL = sSQL & "IIF(RTALI.AL03_sExplainedRCV, 'X', '') "
        sSQL = sSQL & ") As AL03_sExplainedRCVY, "
        sSQL = sSQL & "( "
        sSQL = sSQL & "IIF(Not RTALI.AL03_sExplainedRCV And Not RTALI.AL03_sExplainedRCVNA , 'X', '') "
        sSQL = sSQL & ") As AL03_sExplainedRCVN, "
        sSQL = sSQL & "( "
        sSQL = sSQL & "IIF(RTALI.AL03_sExplainedRCVNA, 'X', '') "
        sSQL = sSQL & ") As AL03_sExplainedRCVNA, "
        'AL04_sConfirmMortgageeIsCorrect
        sSQL = sSQL & "( "
        sSQL = sSQL & "IIF(RTALI.AL04_sConfirmMortgageeIsCorrect, 'X', '') "
        sSQL = sSQL & ") As AL04_sConfirmMortgageeIsCorrectY, "
        sSQL = sSQL & "( "
        sSQL = sSQL & "IIF(Not RTALI.AL04_sConfirmMortgageeIsCorrect And Not RTALI.AL04_sConfirmMortgageeIsCorrectNA, 'X', '') "
        sSQL = sSQL & ") As AL04_sConfirmMortgageeIsCorrectN, "
        sSQL = sSQL & "( "
        sSQL = sSQL & "IIF(RTALI.AL04_sConfirmMortgageeIsCorrectNA, 'X', '') "
        sSQL = sSQL & ") As AL04_sConfirmMortgageeIsCorrectNA, "
        'AL05_sExplainedMortgageeChecksY
        sSQL = sSQL & "( "
        sSQL = sSQL & "IIF(RTALI.AL05_sExplainedMortgageeChecks, 'X', '') "
        sSQL = sSQL & ") As AL05_sExplainedMortgageeChecksY, "
        sSQL = sSQL & "( "
        sSQL = sSQL & "IIF(Not RTALI.AL05_sExplainedMortgageeChecks And Not RTALI.AL05_sExplainedMortgageeChecksNA, 'X', '') "
        sSQL = sSQL & ") As AL05_sExplainedMortgageeChecksN, "
        sSQL = sSQL & "( "
        sSQL = sSQL & "IIF(RTALI.AL05_sExplainedMortgageeChecksNA, 'X', '') "
        sSQL = sSQL & ") As AL05_sExplainedMortgageeChecksNA, "
        'AL06_sConfirmedCoverageY
        sSQL = sSQL & "( "
        sSQL = sSQL & "IIF(RTALI.AL06_sConfirmedCoverage, 'X', '') "
        sSQL = sSQL & ") As AL06_sConfirmedCoverageY, "
        sSQL = sSQL & "( "
        sSQL = sSQL & "IIF(Not RTALI.AL06_sConfirmedCoverage, 'X', '') "
        sSQL = sSQL & ") As AL06_sConfirmedCoverageN, "
        'AL07_sPriorLossY
        sSQL = sSQL & "( "
        sSQL = sSQL & "IIF(RTALI.AL07_sPriorLoss, 'X', '') "
        sSQL = sSQL & ") As AL07_sPriorLossY, "
        sSQL = sSQL & "( "
        sSQL = sSQL & "IIF(Not RTALI.AL07_sPriorLoss And Not RTALI.AL07_sPriorLossNA, 'X', '') "
        sSQL = sSQL & ") As AL07_sPriorLossN, "
        sSQL = sSQL & "( "
        sSQL = sSQL & "IIF(RTALI.AL07_sPriorLossNA, 'X', '') "
        sSQL = sSQL & ") As AL07_sPriorLossNA, "
        'AL08_sSalvage
        sSQL = sSQL & "( "
        sSQL = sSQL & "IIF(RTALI.AL08_sSalvage, 'YES', 'NO') "
        sSQL = sSQL & ") As AL08_sSalvageYesNo, "
        'AL09_sSubrogation
        sSQL = sSQL & "( "
        sSQL = sSQL & "IIF(RTALI.AL09_sSubrogation, 'YES', 'NO') "
        sSQL = sSQL & ") As AL09_sSubrogationYesNo, "
        '------------------------------------END Access DB---------------------------
        '
        '----------------------------------------------------------------------------
    End If
    sSQL = sSQL & "( "
    If goUtil.gMainDB Is Nothing Then
        '------------------------------------BEGIN SQL SERVER DB-------------------------
        '
        '--------------------------------------------------------------------------------
        sSQL = sSQL & "SELECT   Round(Sum([ServiceTime]), 2) As ServiceTimeSum "
        sSQL = sSQL & "FROM     RTActivityLog "
        sSQL = sSQL & "WHERE    [AssignmentsID] = A.[AssignmentsID] "
        sSQL = sSQL & "AND [IsDeleted] = 0 "
        sSQL = sSQL & "AND [IsMgrEntry] = 0 "
        '------------------------------------END SQL SERVER DB-------------------------
        '
        '------------------------------------------------------------------------------
    Else
        '------------------------------------BEGIN Access DB-------------------------
        '
        '----------------------------------------------------------------------------
        sSQL = sSQL & "SELECT   Format(Sum([ServiceTime]), '0.00') As ServiceTimeSum "
        sSQL = sSQL & "FROM     RTActivityLog "
        sSQL = sSQL & "WHERE    [AssignmentsID] = A.[AssignmentsID] "
        sSQL = sSQL & "AND [IsDeleted] = 0 "
        sSQL = sSQL & "AND [IsMgrEntry] = 0 "
        '------------------------------------END Access DB---------------------------
        '
        '----------------------------------------------------------------------------
    End If
    
    sSQL = sSQL & ") As ServiceTimeSum "
    sSQL = sSQL & "FROM Assignments A "
    sSQL = sSQL & "INNER JOIN RTActivityLogInfo RTALI ON RTALI.AssignmentsID = A.AssignmentsID "
    sSQL = sSQL & "WHERE A.AssignmentsID = " & msAssignmentsID & " "
    
    Set mRS = New ADODB.Recordset
    mRS.CursorLocation = adUseClient
    mRS.Open sSQL, mConn, adOpenForwardOnly, adLockReadOnly
    Set mRS.ActiveConnection = Nothing
    
    With PopulateudtarRpt
        .f_Insured = vbNullString
        .f_ClientClaimNo = vbNullString
        .f_ADJNAME = vbNullString
        .f_ReceivedDate = NULL_DATE
        .f_ContactDate = NULL_DATE
        .f_InspectedDate = NULL_DATE
        .f_AL01_sPresentDurringInspectionY = vbNullString
        .f_AL01_sPresentDurringInspectionN = vbNullString
        .f_AL02_sExplainedEstimateY = vbNullString
        .f_AL02_sExplainedEstimateN = vbNullString
        .f_AL03_sExplainedRCVY = vbNullString
        .f_AL03_sExplainedRCVN = vbNullString
        .f_AL03_sExplainedRCVNA = vbNullString
        .f_AL04_sConfirmMortgageeIsCorrectY = vbNullString
        .f_AL04_sConfirmMortgageeIsCorrectN = vbNullString
        .f_AL04_sConfirmMortgageeIsCorrectNA = vbNullString
        .f_AL05_sExplainedMortgageeChecksY = vbNullString
        .f_AL05_sExplainedMortgageeChecksN = vbNullString
        .f_AL05_sExplainedMortgageeChecksNA = vbNullString
        .f_AL06_sConfirmedCoverageY = vbNullString
        .f_AL06_sConfirmedCoverageN = vbNullString
        .f_AL07_sPriorLossY = vbNullString
        .f_AL07_sPriorLossN = vbNullString
        .f_AL07_sPriorLossNA = vbNullString
        .f_AL08_sSalvageYesNo = vbNullString
        .f_AL09_sSubrogationYesNo = vbNullString
        .f_ServiceTimeSum = 0
        .f_CloseDate = NULL_DATE
        .f_IBNUM = vbNullString
        On Error Resume Next
        .f_Insured = mRS!Insured
        .f_ClientClaimNo = mRS!ClientClaimNo
        .f_ADJNAME = mRS!ADJNAME
        .f_ReceivedDate = mRS!ReceivedDate
        .f_ContactDate = mRS!ContactDate
        .f_InspectedDate = mRS!InspectedDate
        .f_AL01_sPresentDurringInspectionY = mRS!AL01_sPresentDurringInspectionY
        .f_AL01_sPresentDurringInspectionN = mRS!AL01_sPresentDurringInspectionN
        .f_AL02_sExplainedEstimateY = mRS!AL02_sExplainedEstimateY
        .f_AL02_sExplainedEstimateN = mRS!AL02_sExplainedEstimateN
        .f_AL03_sExplainedRCVY = mRS!AL03_sExplainedRCVY
        .f_AL03_sExplainedRCVN = mRS!AL03_sExplainedRCVN
        .f_AL03_sExplainedRCVNA = mRS!AL03_sExplainedRCVNA
        .f_AL04_sConfirmMortgageeIsCorrectY = mRS!AL04_sConfirmMortgageeIsCorrectY
        .f_AL04_sConfirmMortgageeIsCorrectN = mRS!AL04_sConfirmMortgageeIsCorrectN
        .f_AL04_sConfirmMortgageeIsCorrectNA = mRS!AL04_sConfirmMortgageeIsCorrectNA
        .f_AL05_sExplainedMortgageeChecksY = mRS!AL05_sExplainedMortgageeChecksY
        .f_AL05_sExplainedMortgageeChecksN = mRS!AL05_sExplainedMortgageeChecksN
        .f_AL05_sExplainedMortgageeChecksNA = mRS!AL05_sExplainedMortgageeChecksNA
        .f_AL06_sConfirmedCoverageY = mRS!AL06_sConfirmedCoverageY
        .f_AL06_sConfirmedCoverageN = mRS!AL06_sConfirmedCoverageN
        .f_AL07_sPriorLossY = mRS!AL07_sPriorLossY
        .f_AL07_sPriorLossN = mRS!AL07_sPriorLossN
        .f_AL07_sPriorLossNA = mRS!AL07_sPriorLossNA
        .f_AL08_sSalvageYesNo = mRS!AL08_sSalvageYesNo
        .f_AL09_sSubrogationYesNo = mRS!AL09_sSubrogationYesNo
        .f_ServiceTimeSum = mRS!ServiceTimeSum
        .f_CloseDate = mRS!CloseDate
        .f_IBNUM = mRS!IBNUM
        .PrintBlankReport = mbPrintBlankReport
        On Error GoTo EH
    End With
    
    If mbPrintBlankReport Then
        GoTo PRINT_BLANK_REPORT
    End If
    
    sSQL = "SELECT "
    sSQL = sSQL & "AL.RTActivityLogID, "
    sSQL = sSQL & "AL.ActDate, "
    sSQL = sSQL & "AL.ActText, "
    sSQL = sSQL & "AL.ActTime, "
    sSQL = sSQL & "AL.BlankPageAfter, "
    sSQL = sSQL & "AL.BlankRowsAfter, "
    sSQL = sSQL & "AL.PageBreakAfter, "
    sSQL = sSQL & "AL.ServiceTime, "
    sSQL = sSQL & "( "
    If goUtil.gMainDB Is Nothing Then
        '------------------------------------BEGIN SQL SERVER DB-------------------------
        '
        '--------------------------------------------------------------------------------
        sSQL = sSQL & "CASE WHEN RTIB.RT14a_sSupplement Is Null Or IB.IB14a_sSupplement Is Not Null "
        sSQL = sSQL & "THEN ( "
                sSQL = sSQL & "CASE WHEN IB.IB14a_sSupplement Is Null Or IB.IB14a_sSupplement = 0 "
                sSQL = sSQL & "THEN 'IB' "
                sSQL = sSQL & "ELSE 'IB' + Cast(IB.IB14a_sSupplement As Varchar(10)) "
                sSQL = sSQL & "END "
                sSQL = sSQL & ") "
        sSQL = sSQL & "ELSE ( "
                sSQL = sSQL & "CASE WHEN RTIB.RT14a_sSupplement Is Null Or RTIB.RT14a_sSupplement = 0 "
                sSQL = sSQL & "THEN 'IB' "
                sSQL = sSQL & "ELSE 'IB' + Cast(RTIB.RT14a_sSupplement As Varchar(10)) "
                sSQL = sSQL & "END "
                sSQL = sSQL & ") "
        sSQL = sSQL & "END "
        '------------------------------------END SQL SERVER DB---------------------------
        '
        '--------------------------------------------------------------------------------
    Else
        '------------------------------------BEGIN Access DB----------------------------
        '
        '-------------------------------------------------------------------------------
        sSQL = sSQL & "SELECT 'IB' & IIF(BC.Supplement > 0 ,'S' & Format(BC.Supplement,'00') ,'' ) & IIF(BC.Rebill > 0 ,'R' & Format(BC.Rebill,'00') ,'' ) "
        sSQL = sSQL & "FROM     BillingCount BC "
        sSQL = sSQL & "WHERE    BC.BillingCountID = AL.BillingCountID "
        sSQL = sSQL & "AND      BC.AssignmentsID = AL.AssignmentsID "
        '------------------------------------END Access DB------------------------------
        '
        '-------------------------------------------------------------------------------
    End If
    sSQL = sSQL & ") AS IB14a_sSupplement "
    sSQL = sSQL & "FROM "
    sSQL = sSQL & "( "
    sSQL = sSQL & "RTActivityLog AL "
    sSQL = sSQL & "LEFT JOIN RTIB  ON (AL.AssignmentsID = RTIB.AssignmentsID AND AL.BillingCountID = RTIB.BillingCountID) "
    sSQL = sSQL & ") "
    sSQL = sSQL & "LEFT JOIN IB    ON (AL.AssignmentsID = IB.AssignmentsID AND AL.BillingCountID = IB.BillingCountID) "
    sSQL = sSQL & "WHERE AL.AssignmentsID = " & msAssignmentsID & " "
    sSQL = sSQL & "AND AL.IsDeleted = 0 "
    sSQL = sSQL & "ORDER BY AL.[ActTime], ABS(AL.[RTActivityLogID]) "
    
    Set mRS = New ADODB.Recordset
    mRS.CursorLocation = adUseClient
    mRS.Open sSQL, mConn, adOpenForwardOnly, adLockReadOnly
    Set mRS.ActiveConnection = Nothing
    
    Set PopulateudtarRpt.coludtActivityLog01 = New Collection
    If Not mRS.EOF Then
        mRS.MoveFirst
        'First need to build the Photo Path. The RS will only contain the
        'Name of the photo. depending upon what DB working with the actual
        'Photo path will differ.
        Do Until mRS.EOF
            'SQL Server the Photo path will be
            With MyActLog
                .f_RTActivityLogID = 0
                .f_ActDate = NULL_DATE
                .f_ActText = vbNullString
                .f_ActTime = NULL_DATE
                .f_BlankPageAfter = 0
                .f_BlankRowsAfter = 0
                .f_IB14a_sSupplement = vbNullString
                .f_PageBreakAfter = 0
                .f_ServiceTime = 0
                On Error Resume Next
                .f_RTActivityLogID = mRS!RTActivityLogID
                .f_ActDate = mRS!ActDate
                .f_ActText = mRS!ActText
                .f_ActTime = mRS!ActTime
                .f_BlankPageAfter = mRS!BlankPageAfter
                .f_BlankRowsAfter = mRS!BlankRowsAfter
                .f_IB14a_sSupplement = mRS!IB14a_sSupplement
                .f_PageBreakAfter = mRS!PageBreakAfter
                .f_ServiceTime = mRS!ServiceTime
                On Error GoTo EH
            End With
            PopulateudtarRpt.coludtActivityLog01.Add MyActLog, CStr(MyActLog.f_RTActivityLogID)
            mRS.MoveNext
        Loop
    End If
    
PRINT_BLANK_REPORT:
    If mbPrintBlankReport Then
        Set PopulateudtarRpt.coludtActivityLog01 = New Collection
        For lCount = 1 To 12
            With MyActLog
                .f_RTActivityLogID = lCount
                .f_ActDate = NULL_DATE
                .f_ActText = vbNullString
                .f_ActTime = NULL_DATE
                .f_BlankPageAfter = 0
                .f_BlankRowsAfter = 0
                .f_IB14a_sSupplement = 0
                .f_PageBreakAfter = 0
                .f_ServiceTime = 0
                On Error GoTo EH
            End With
            PopulateudtarRpt.coludtActivityLog01.Add MyActLog, CStr(MyActLog.f_RTActivityLogID)
        Next
    End If

    mRS.Close
    Set mRS = Nothing
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
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
    sMiscReportParamName = "MiscReportParam05"
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



