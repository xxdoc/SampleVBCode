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
Private marRpt As ECrpt_arCommission.arCommission

Public Type udtCommission01
    AssignmentsID As Long
    f_CLIENTNUM As String
    f_Insured As String
    f_ServiceFee As Currency
    f_OutBldg As Currency
    f_OtherFees As Currency
    f_Photos As Currency
    f_Mileage As Currency
    f_MiscExp As Currency
    f_Tax As Currency
    f_Total As Currency
End Type

Private Type udtarRpt
    CommissionPercentage As Double
    f_PrintDate As Date
    f_ADJName As String
    f_ServiceFeeSub As Currency
    f_OutBldgSub As Currency
    f_OtherFeesSub As Currency
    f_PhotosSub As Currency
    f_MileageSub As Currency
    f_MiscExpSub As Currency
    f_TaxSub As Currency
    f_TotalSub As Currency
    f_gndTtlCommOnServices As Currency
    f_gndTtlPhotosAndMileage As Currency
    f_gndTtlNoOfBilledFiles As Long
    f_gndTtlAvgCommPerFile As Currency
    f_gndTtlTotalCommissions As Currency
    f_gndTtlCatName As String
    f_gndTtlSiteLocation As String
    coludtCommission01 As Collection
End Type

Private msClientCompanyID As String
Private msCATID As String
Private msUSERSID As String
Private mdCommissionPercentage As Double
Private mbHideDetails As Boolean
Private mbExportXML As Boolean
Private mbExportXMLOnly As Boolean

Public Property Let HideDetails(pbFlag As Boolean)
    mbHideDetails = pbFlag
End Property
Public Property Get HideDetails() As Boolean
    HideDetails = mbHideDetails
End Property

Public Property Let CommissionPercentage(pdCommissionPercentage As Double)
    mdCommissionPercentage = pdCommissionPercentage
End Property
Public Property Get CommissionPercentage() As Double
    CommissionPercentage = mdCommissionPercentage
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
        Set marRpt = New ECrpt_arCommission.arCommission
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
    
    Set marRpt = New ECrpt_arCommission.arCommission
    
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
        marRpt.SetProperty "f_ServiceFeeSub", .f_ServiceFeeSub, vbCurrency
'        marRpt.SetProperty "f_OutBldgSub", .f_OutBldgSub, vbCurrency
        marRpt.SetProperty "f_OtherFeesSub", .f_OtherFeesSub, vbCurrency
'        marRpt.SetProperty "f_PhotosSub", .f_PhotosSub, vbCurrency
'        marRpt.SetProperty "f_MileageSub", .f_MileageSub, vbCurrency
        marRpt.SetProperty "f_MiscExpSub", .f_MiscExpSub, vbCurrency
        marRpt.SetProperty "f_TaxSub", .f_TaxSub, vbCurrency
        marRpt.SetProperty "f_TotalSub", .f_TotalSub, vbCurrency
        marRpt.SetProperty "f_gndTtlCommOnServices", .f_gndTtlCommOnServices, vbCurrency
        marRpt.SetProperty "f_gndTtlPhotosAndMileage", .f_gndTtlPhotosAndMileage, vbCurrency
        marRpt.SetProperty "f_gndTtlNoOfBilledFiles", .f_gndTtlNoOfBilledFiles, vbLong
        marRpt.SetProperty "f_gndTtlAvgCommPerFile", .f_gndTtlAvgCommPerFile, vbCurrency
        marRpt.SetProperty "f_gndTtlTotalCommissions", .f_gndTtlTotalCommissions, vbCurrency
        marRpt.SetProperty "f_gndTtlCatName", .f_gndTtlCatName, vbString
        marRpt.SetProperty "f_gndTtlSiteLocation", .f_gndTtlSiteLocation, vbString
        marRpt.SetProperty "coludtCommission01", .coludtCommission01, vbObject
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
    Dim MyCommision As udtCommission01
    
    'Gnd Ttl Fee Vars
    Dim cMygndTtlCommOnServices As Currency
    Dim cMygndTtlPhotosAndMileage As Currency
    Dim lMygndTtlNoOfBilledFiles As Long
    Dim cMygndTtlTotalCommissions As Currency
    Dim cMygndTtlAvgCommPerFile As Currency
    
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
    
    
    'B.  get the Commision Report Info
    sSQL = "SELECT "
    sSQL = sSQL & "( "
    sSQL = sSQL & "SELECT   LastName + ', ' + FirstName "
    sSQL = sSQL & "FROM     USERS "
    sSQL = sSQL & "WHERE    USERSID = " & msUSERSID & " "
    sSQL = sSQL & ") As ADJName, "
    sSQL = sSQL & "SUM(IB.IB17_cServiceFee) As ServiceFeeSub, "
    sSQL = sSQL & "( "
    sSQL = sSQL & "SUM(IB.IB25_cServiceFeeSubTotal) - "
    sSQL = sSQL & "SUM(IB.IB17_cServiceFee)  "
    sSQL = sSQL & ") As OtherFeesSub, "
    sSQL = sSQL & "SUM(IB.IB30_cTotalExpenses) As MiscExpSub, "
    sSQL = sSQL & "SUM(IB.IB32_cTaxAmount) As TaxSub, "
    sSQL = sSQL & "SUM(IB.IB33_cTotalAdjustingFee) As TotalSub, "
    sSQL = sSQL & "COUNT(IB.IBID) As gndTtNoOfBilledFiles, "
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
    sSQL = sSQL & "LEFT JOIN IB ON A.AssignMentsID = IB.AssignmentsID "
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
    sSQL = sSQL & "AND A.IsDeleted = 0 "
    sSQL = sSQL & "AND (A.Reassigned = 0 Or A.Reassigned Is Null) "
    sSQL = sSQL & "AND A.CLIENTNUM Not Like '%TEST%' "
    sSQL = sSQL & "AND (IB.Void = 0 Or IB.Void Is Null) "

    
    Set mRS = New ADODB.Recordset
    mRS.CursorLocation = adUseClient
    mRS.Open sSQL, mConn, adOpenStatic, adLockReadOnly
    Set mRS.ActiveConnection = Nothing
    With PopulateudtarRpt
            .CommissionPercentage = 0
            .f_PrintDate = NULL_DATE
            .f_ADJName = vbNullString
            .f_ServiceFeeSub = 0
            .f_OutBldgSub = 0
            .f_OtherFeesSub = 0
            .f_PhotosSub = 0
            .f_MileageSub = 0
            .f_MiscExpSub = 0
            .f_TaxSub = 0
            .f_TotalSub = 0
            .f_gndTtlCommOnServices = 0
            .f_gndTtlPhotosAndMileage = 0
            .f_gndTtlNoOfBilledFiles = 0
            .f_gndTtlAvgCommPerFile = 0
            .f_gndTtlTotalCommissions = 0
            .f_gndTtlCatName = vbNullString
            .f_gndTtlSiteLocation = vbNullString
            On Error Resume Next
            .CommissionPercentage = mdCommissionPercentage
            'Figure Gnd Ttl Fees Here
            '1 gndTtlCommOnServices
            cMygndTtlCommOnServices = mRS!ServiceFeeSub + mRS!OtherFeesSub
            cMygndTtlCommOnServices = cMygndTtlCommOnServices * mdCommissionPercentage
            
            '2.gndTtPhotosAndMileage
            cMygndTtlPhotosAndMileage = mRS!MiscExpSub
            
            '3. gndTtlNoOfBilledFiles
            lMygndTtlNoOfBilledFiles = mRS!gndTtNoOfBilledFiles
            
            '4.gndTtTotalCommissions
            cMygndTtlTotalCommissions = cMygndTtlCommOnServices + cMygndTtlPhotosAndMileage
            
            '5. gndTtAvgCommPerFile
            'Check for Div by Zero
            If lMygndTtlNoOfBilledFiles <= 0 Then
                cMygndTtlAvgCommPerFile = 0
            Else
                cMygndTtlAvgCommPerFile = cMygndTtlTotalCommissions / lMygndTtlNoOfBilledFiles
            End If
            
            .f_PrintDate = Now()
            .f_ADJName = mRS!ADJName
            .f_ServiceFeeSub = mRS!ServiceFeeSub
            .f_OtherFeesSub = mRS!OtherFeesSub
            .f_MiscExpSub = mRS!MiscExpSub
            .f_TaxSub = mRS!TaxSub
            .f_TotalSub = mRS!TotalSub
            .f_gndTtlCommOnServices = cMygndTtlCommOnServices
            .f_gndTtlPhotosAndMileage = cMygndTtlPhotosAndMileage
            .f_gndTtlNoOfBilledFiles = lMygndTtlNoOfBilledFiles
            .f_gndTtlAvgCommPerFile = cMygndTtlAvgCommPerFile
            .f_gndTtlTotalCommissions = cMygndTtlTotalCommissions
            .f_gndTtlCatName = mRS!gndTtlCatName
            .f_gndTtlSiteLocation = mRS!gndTtlSiteLocation
            On Error GoTo EH
        End With
    
    sSQL = "SELECT "
    sSQL = sSQL & "A.AssignmentsID, "
    sSQL = sSQL & "A.CLIENTNUM, "
    sSQL = sSQL & "A.Insured, "
    sSQL = sSQL & "SUM(IB.IB17_cServiceFee) As ServiceFee, "
    sSQL = sSQL & "( "
    sSQL = sSQL & "SUM(IB.IB25_cServiceFeeSubTotal) - "
    sSQL = sSQL & "SUM(IB.IB17_cServiceFee)  "
    sSQL = sSQL & ") As OtherFees, "
    sSQL = sSQL & "SUM(IB.IB30_cTotalExpenses) As MiscExp, "
    sSQL = sSQL & "SUM(IB.IB32_cTaxAmount) As Tax, "
    sSQL = sSQL & "SUM(IB.IB33_cTotalAdjustingFee) As Total "
    sSQL = sSQL & "FROM Assignments A "
    sSQL = sSQL & "LEFT JOIN IB ON A.AssignMentsID = IB.AssignmentsID "
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
    sSQL = sSQL & "AND A.IsDeleted = 0 "
    sSQL = sSQL & "AND (A.Reassigned = 0 Or A.Reassigned Is Null) "
    sSQL = sSQL & "AND A.CLIENTNUM Not Like '%TEST%' "
    sSQL = sSQL & "AND (IB.Void = 0 Or IB.Void Is Null) "
    sSQL = sSQL & "GROUP BY A.AssignmentsID, A.CLIENTNUM, A.Insured "
    sSQL = sSQL & "ORDER BY A.CLIENTNUM "
    
    Set mRS = New ADODB.Recordset
    mRS.CursorLocation = adUseClient
    mRS.Open sSQL, mConn, adOpenStatic, adLockReadOnly
    Set mRS.ActiveConnection = Nothing
    
    Set PopulateudtarRpt.coludtCommission01 = New Collection
    
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
            With MyCommision
                .AssignmentsID = 0
                .f_CLIENTNUM = vbNullString
                .f_Insured = vbNullString
                .f_ServiceFee = 0
                .f_OutBldg = 0
                .f_OtherFees = 0
                .f_Photos = 0
                .f_Mileage = 0
                .f_MiscExp = 0
                .f_Tax = 0
                .f_Total = 0
                On Error Resume Next
                .AssignmentsID = mRS!AssignmentsID
                .f_CLIENTNUM = mRS!CLIENTNUM
                .f_Insured = mRS!Insured
                .f_ServiceFee = mRS!ServiceFee
                .f_OtherFees = mRS!OtherFees
                .f_MiscExp = mRS!MiscExp
                .f_Tax = mRS!Tax
                .f_Total = mRS!Total
                On Error GoTo EH
            End With
            PopulateudtarRpt.coludtCommission01.Add MyCommision, CStr(MyCommision.AssignmentsID)
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





