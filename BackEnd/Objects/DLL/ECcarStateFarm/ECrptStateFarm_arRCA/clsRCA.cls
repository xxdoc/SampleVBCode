VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRCA"
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

'Line Items
'Line 01 Policy Limits
Private mcDwellPolicyLim As Currency         '{Assignments.BuildingPolicyLimit}/{rtfarmer.CI14_cBu}
Private mcContPolicyLim As Currency          '{Assignments.PersonalPropertyLimit}/{rtfarmer.CI15_cCo}
'Line 02 Full Replacement Cost of the said property at the time of the loss was
Private mcDwellRCSaid As Currency            '{Assignments.RCSaidBuilding}/{rtfarmer.CI17_cBu}
Private mcContRCSaid As Currency             '{Assignments.RCSaidPersonal}/{rtfarmer.CI18_cCo}
'Line 03 The Full Cost of Repair or Replacement is
Private mcDwellFullRC As Currency            '{indemnity.ReplacementCost} where {left(indemnity.class,2) ='01'}/{rtfarmerindem.cReplaceme} where {left(rtfarmerindem.sClass,2) ='01'}
Private mcContFullRC As Currency             '{indemnity.ReplacementCost}where {left(indemnity.class,2) ='02'/{rtfarmerindem.cReplaceme} where {left(rtfarmerindem.sClass,2) ='02'}
'Line 04 Recov Depreciation
Private mcDwellRecDep As Currency            '{indemnity.RecoverableDepreciation} where {left(indemnity.class,2) ='01'}/{rtfarmerindem.cRecoverab}where {left(rtfarmerindem.sClass,2) ='01'}
Private mcContRecDep As Currency             '{indemnity.RecoverableDepreciation} where {left(indemnity.class,2) ='02'}/{rtfarmerindem.cRecoverab}where {left(rtfarmerindem.sClass,2) ='02'}
'Line 05 Non Recov Depreciation
Private mcDwellNonRecDep As Currency         '{indemnity.NonRecoverableDepreciation} where {left(indemnity.class,2) ='01'}/{rtfarmerindem.cNonRecove}where {left(rtfarmerindem.sClass,2) ='01'}
Private mcContNonRecDep As Currency          '{indemnity.NonRecoverableDepreciation} where {left(indemnity.class,2) ='02'}/{rtfarmerindem.cNonRecove}where {left(rtfarmerindem.sClass,2) ='02'}
'Line 06 ACV Actual Cash Value loss is (Line 3 minus Line 4 and Line 5)
Private mcDwellACV As Currency               '{indemnity.ACVClaim} where {left(indemnity.class,2) ='01'}/{rtfarmerindem.cACVClaim}where {left(rtfarmerindem.sClass,2) ='01'}
Private mcContACV As Currency                '{indemnity.ACVClaim} where {left(indemnity.class,2) ='02'}/{rtfarmerindem.cACVClaim}where {left(rtfarmerindem.sClass,2) ='02'}
'XS Limits
Private mcDwellXSLim As Currency             '{indemnity.LessExcessLimit} where {left(indemnity.class,2) ='01'}/{rtfarmerindem.cExcessLim}where {left(rtfarmerindem.sClass,2) ='01'}
Private mcContXSLim As Currency              '{indemnity.LessExcessLimit} where {left(indemnity.class,2) ='02'}/{rtfarmerindem.cExcessLim}where {left(rtfarmerindem.sClass,2) ='02'}
'Acv Less XS Limits (Deductible not applied yet)
Private mcDwellACVLessXSLim As Currency      '{indemnity.ACVLessExcessLimit} where {left(indemnity.class,2) ='01'}/{rtfarmerindem.cACVLessEx}where {left(rtfarmerindem.sClass,2) ='01'}
Private mcContACVLessXSLim As Currency       '{indemnity.ACVLessExcessLimit} where {left(indemnity.class,2) ='02'}/{rtfarmerindem.cACVLessEx}where {left(rtfarmerindem.sClass,2) ='02'}
'Line07 Dwelling Less Prior Payments
'<!--- Leave these 0 value for first version--->
Private mcDwellPrevPaymnet As Currency
Private mcContPrevPayment As Currency
'Line 08 Deductible
'See --->PopulateInitDeductibles
Private mcDeductible As Currency             '{Assignments.Deductible}/{rtfarmer.RT15_c}
Private mcDwellDed As Currency               'See --->PopulateInitDeductibles
Private mcContDed As Currency                'See --->PopulateInitDeductibles
'Line 09 Actual Cash Value Claim is (Line 6 minus Line 7 and Line 8)
Private mcDwellACVC As Currency
Private mcContACVC As Currency
'Line 10 Supplemental Claim, to be filed in accordance with the terms and conditions of the Replacement Cost Coverage
Private mcLineSuppClaim As Currency

Private Type udtarRpt
    f_RCAOther As String
    f_RT10_sInsuredName As String
    f_RT07_sAdjusterName As String
    f_CI11sPolicyNumber As String
    f_RT09_sSALN As String
    f_RT12_dtDateOfLoss As Date
    f_RT11_sLossLocation As String
    f_Dwell01 As Currency
    f_Dwell02 As Currency
    f_Dwell03 As Currency
    f_Dwell04 As Currency
    f_Dwell05 As Currency
    f_Dwell06 As Currency
    f_Dwell07 As Currency
    f_Dwell08 As Currency
    f_Dwell09 As Currency
    f_Cont01 As Currency
    f_Cont02 As Currency
    f_Cont03 As Currency
    f_Cont04 As Currency
    f_Cont05 As Currency
    f_Cont06 As Currency
    f_Cont07 As Currency
    f_Cont08 As Currency
    f_Cont09 As Currency
    f_Line10 As Currency
End Type
Private marRpt As ECrptStateFarm_arRCA.arRCA
Private msWithinDays As String
Private mbExportXML As Boolean
Private mbExportXMLOnly As Boolean

Private Property Get clsRptDB_MiscDelimParamsCol() As Collection
    Dim bCleanup As Boolean
    'Add Report Parameters that are not updated by DB Fields.
    If marRpt Is Nothing Then
        bCleanup = True
        Set marRpt = New ECrptStateFarm_arRCA.arRCA
    End If
    Set clsRptDB_MiscDelimParamsCol = GetMiscDelimParamsCol()
    If bCleanup Then
        Set mConn = Nothing
        Set mRS = Nothing
        Unload marRpt
        Set marRpt = Nothing
    End If
End Property

Public Property Let WithinDays(psWithinDays As String)
    msWithinDays = psWithinDays
End Property
Public Property Get WithinDays() As String
    WithinDays = msWithinDays
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



Public Function clsRptDB_PrintARReport() As Boolean

    On Error GoTo EH
    Dim udtRCA As udtarRpt
    Dim colParams As Collection
    Dim vParams As Variant
    Dim MyParams As MiscReportParam
    Dim sXML As String
    Dim sXMLXportPath As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    Set marRpt = New ECrptStateFarm_arRCA.arRCA
    
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
    
    udtRCA = PopulateudtarRpt
    
    With udtRCA
         marRpt.SetProperty "f_RCAOther", .f_RCAOther, vbString
         marRpt.SetProperty "f_RT10_sInsuredName", .f_RT10_sInsuredName, vbString
         marRpt.SetProperty "f_RT07_sAdjusterName", .f_RT07_sAdjusterName, vbString
         marRpt.SetProperty "f_CI11sPolicyNumber", .f_CI11sPolicyNumber, vbString
         marRpt.SetProperty "f_RT09_sSALN", .f_RT09_sSALN, vbString
         marRpt.SetProperty "f_RT12_dtDateOfLoss", .f_RT12_dtDateOfLoss, vbString
         marRpt.SetProperty "f_RT11_sLossLocation", .f_RT11_sLossLocation, vbString
         marRpt.SetProperty "f_Dwell01", .f_Dwell01, vbCurrency
         marRpt.SetProperty "f_Cont01", .f_Cont01, vbCurrency
         marRpt.SetProperty "f_Dwell02", .f_Dwell02, vbCurrency
         marRpt.SetProperty "f_Cont02", .f_Cont02, vbCurrency
         marRpt.SetProperty "f_Dwell03", .f_Dwell03, vbCurrency
         marRpt.SetProperty "f_Cont03", .f_Cont03, vbCurrency
         marRpt.SetProperty "f_Dwell04", .f_Dwell04, vbCurrency
         marRpt.SetProperty "f_Cont04", .f_Cont04, vbCurrency
         marRpt.SetProperty "f_Dwell05", .f_Dwell05, vbCurrency
         marRpt.SetProperty "f_Cont05", .f_Cont05, vbCurrency
         marRpt.SetProperty "f_Dwell06", .f_Dwell06, vbCurrency
         marRpt.SetProperty "f_Cont06", .f_Cont06, vbCurrency
         marRpt.SetProperty "f_Dwell07", .f_Dwell07, vbCurrency
         marRpt.SetProperty "f_Cont07", .f_Cont07, vbCurrency
         marRpt.SetProperty "f_Dwell08", .f_Dwell08, vbCurrency
         marRpt.SetProperty "f_Cont08", .f_Cont08, vbCurrency
         marRpt.SetProperty "f_Dwell09", .f_Dwell09, vbCurrency
         marRpt.SetProperty "f_Cont09", .f_Cont09, vbCurrency
         marRpt.SetProperty "f_Line10", .f_Line10, vbCurrency
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
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsRptDB_ClassName & vbCrLf & "Public Function PrintRCA"
End Function

Private Function PopulateudtarRpt() As udtarRpt
    On Error GoTo EH
    Dim sSQL As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
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
    
    With PopulateudtarRpt

        
        'C.
        'Next get Header Info and Policy Limits.
        sSQL = "SELECT A.Insured As InsuredNames, "
        sSQL = sSQL & "A.PAState As PropertyState, "
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
        sSQL = sSQL & "A.PolicyNo As PolicyNumber, "
        sSQL = sSQL & "A.LossDate AS LossDate, "
        sSQL = sSQL & "A.PropertyAddress As LossLocation, "
        sSQL = sSQL & "( "
        sSQL = sSQL & "SELECT   SUM(P.LimitAmount) "
        sSQL = sSQL & "FROM     PolicyLimits P "
        sSQL = sSQL & "WHERE    P.AssignmentsID = A.AssignmentsID "
        sSQL = sSQL & "AND      P.ClassTypeID IN "
                                sSQL = sSQL & "( "
                                sSQL = sSQL & "SELECT   ClassTypeID "
                                sSQL = sSQL & "FROM     ClassType "
                                'A = COV. A - BUILDING
                                'E = COV. E - COMMERCIAL BUILDING
                                sSQL = sSQL & "WHERE    Class IN('A','E') "
                                sSQL = sSQL & ") "
        sSQL = sSQL & ") As BuildingPolicyLimit, "
        sSQL = sSQL & "( "
        sSQL = sSQL & "SELECT   SUM(P.LimitAmount) "
        sSQL = sSQL & "FROM     PolicyLimits P "
        sSQL = sSQL & "WHERE    P.AssignmentsID = A.AssignmentsID "
        sSQL = sSQL & "AND      P.ClassTypeID IN "
                                sSQL = sSQL & "( "
                                sSQL = sSQL & "SELECT   ClassTypeID "
                                sSQL = sSQL & "FROM     ClassType "
                                'C = COV. C - PERSONAL PROPERTY COVERAGES
                                sSQL = sSQL & "WHERE    Class IN('C') "
                                sSQL = sSQL & ") "
        sSQL = sSQL & ") As PersonalPropertyLimit, "
        sSQL = sSQL & "( "
        sSQL = sSQL & "SELECT   SUM(P.RCSaidProp) "
        sSQL = sSQL & "FROM     PolicyLimits P "
        sSQL = sSQL & "WHERE    P.AssignmentsID = A.AssignmentsID "
        sSQL = sSQL & "AND      P.ClassTypeID IN "
                                sSQL = sSQL & "( "
                                sSQL = sSQL & "SELECT   ClassTypeID "
                                sSQL = sSQL & "FROM     ClassType "
                                'A = COV. A - BUILDING
                                'E = COV. E - COMMERCIAL BUILDING
                                sSQL = sSQL & "WHERE    Class IN('A','E') "
                                sSQL = sSQL & ") "
        sSQL = sSQL & ") As RCSaidBuilding, "
        sSQL = sSQL & "( "
        sSQL = sSQL & "SELECT   SUM(P.RCSaidProp) "
        sSQL = sSQL & "FROM     PolicyLimits P "
        sSQL = sSQL & "WHERE    P.AssignmentsID = A.AssignmentsID "
        sSQL = sSQL & "AND      P.ClassTypeID IN "
                                sSQL = sSQL & "( "
                                sSQL = sSQL & "SELECT   ClassTypeID "
                                sSQL = sSQL & "FROM     ClassType "
                                'C = COV. C - PERSONAL PROPERTY COVERAGES
                                sSQL = sSQL & "WHERE    Class IN('C') "
                                sSQL = sSQL & ") "
        sSQL = sSQL & ") As RCSaidPersonal, "
        sSQL = sSQL & "A.Deductible As Deductible, "
        sSQL = sSQL & "A.CLIENTNUM "
        sSQL = sSQL & "FROM Assignments A "
        sSQL = sSQL & "WHERE A.AssignmentsID = " & msAssignmentsID & " "
        
        Set mRS = New ADODB.Recordset
        mRS.CursorLocation = adUseClient
        mRS.Open sSQL, mConn, adOpenStatic, adLockReadOnly
        Set mRS.ActiveConnection = Nothing
        
        If Not mRS.EOF Then
            'Rest of the Header info
            .f_RT10_sInsuredName = IIf(IsNull(mRS!InsuredNames), vbNullString, mRS!InsuredNames)
            .f_RT07_sAdjusterName = IIf(IsNull(mRS!ADJName), vbNullString, mRS!ADJName)
            .f_CI11sPolicyNumber = IIf(IsNull(mRS!PolicyNumber), vbNullString, mRS!PolicyNumber)
            .f_RT09_sSALN = IIf(IsNull(mRS!CLIENTNUM), vbNullString, mRS!CLIENTNUM)
            .f_RT12_dtDateOfLoss = IIf(IsNull(mRS!LossDate), NULL_DATE, IIf(IsDate(mRS!LossDate), mRS!LossDate, NULL_DATE))
            .f_RT11_sLossLocation = IIf(IsNull(mRS!LossLocation), vbNullString, mRS!LossLocation)
            
            'Start Line Items
            '1.     Full Amount of Insurance applicable to the
            '       property for which claim is presented was
            mcDwellPolicyLim = IIf(IsNull(mRS!BuildingPolicyLimit), 0, mRS!BuildingPolicyLimit)
            mcContPolicyLim = IIf(IsNull(mRS!PersonalPropertyLimit), 0, mRS!PersonalPropertyLimit)
            
            '2.     Full Replacement Cost of the said property
            '       at the time of the loss was
            mcDwellRCSaid = IIf(IsNull(mRS!RCSaidBuilding), 0, mRS!RCSaidBuilding)
            mcContRCSaid = IIf(IsNull(mRS!RCSaidPersonal), 0, mRS!RCSaidPersonal)
            '
            'Populate the initial Deductible
            mcDeductible = IIf(IsNull(mRS!Deductible), 0, mRS!Deductible)
        End If
        mRS.Close
    End With
    Set mRS = Nothing
    
    With PopulateudtarRpt
        'Line Item 01
        .f_Dwell01 = mcDwellPolicyLim
        .f_Cont01 = mcContPolicyLim
        
        'Line Item 02
        .f_Dwell02 = mcDwellRCSaid
        .f_Cont02 = mcContRCSaid
        
        'Line Items 3 --> 6 get from PolulateIndemValues
        PopulateIndemValues
        
        
        .f_Dwell03 = mcDwellFullRC
        .f_Cont03 = mcContFullRC
        .f_Dwell04 = mcDwellRecDep
        .f_Cont04 = mcContRecDep
        .f_Dwell05 = mcDwellNonRecDep
        .f_Cont05 = mcContNonRecDep
        .f_Dwell06 = mcDwellACV
        .f_Cont06 = mcContACV
        
        '7.     Less Prior Payments
        'These are not poulated in 1st version
        mcDwellPrevPaymnet = 0
        mcContPrevPayment = 0
        .f_Dwell07 = mcDwellPrevPaymnet
        .f_Cont07 = mcContPrevPayment
        
        '8.     Less Deductible
        'First get the initial deductible values. Need these first to figure all deductibles.
        PopulateInitDeductibles
        .f_Dwell08 = GetDwellDed
        .f_Cont08 = GetContDed
        
        '9.     Actual Cash Value Claim is (Line 6 minus Line 7 and Line 8)
        .f_Dwell09 = GetDwellACV
        .f_Cont09 = GetContACV
        
        '10.        Supplemental Claim, to be filed in accordance with the
        '           terms and conditions of the Replacement Cost Coverage
        '           within ____ days of date of loss as shown above, will not
        '           exceed .... ______
        '           (This figure will be that portion of the amounts shown on
        '           Line 4 which is recoverable)
        .f_Line10 = GetSuppClaim
    End With

    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Set RS = Nothing
    Set mRS = Nothing
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsRptDB_ClassName & vbCrLf & "Private Function PopulateudtarRpt"
End Function

Private Sub PopulateIndemValues()
    '    'Line 03 The Full Cost of Repair or Replacement is
'    mcDwellFullRC             '{indemnity.ReplacementCost} where {left(indemnity.class,2) ='01'}/{rtfarmerindem.cReplaceme} where {left(rtfarmerindem.sClass,2) ='01'}
'    mcContFullRC              '{indemnity.ReplacementCost}where {left(indemnity.class,2) ='02'/{rtfarmerindem.cReplaceme} where {left(rtfarmerindem.sClass,2) ='02'}
'    'Line 04 Recov Depreciation
'    mcDwellRecDep            '{indemnity.RecoverableDepreciation} where {left(indemnity.class,2) ='01'}/{rtfarmerindem.cRecoverab}where {left(rtfarmerindem.sClass,2) ='01'}
'    mcContRecDep             '{indemnity.RecoverableDepreciation} where {left(indemnity.class,2) ='02'}/{rtfarmerindem.cRecoverab}where {left(rtfarmerindem.sClass,2) ='02'}
'    'Line 05 Non Recov Depreciation
'    mcDwellNonRecDep         '{indemnity.NonRecoverableDepreciation} where {left(indemnity.class,2) ='01'}/{rtfarmerindem.cNonRecove}where {left(rtfarmerindem.sClass,2) ='01'}
'    mcContNonRecDep          '{indemnity.NonRecoverableDepreciation} where {left(indemnity.class,2) ='02'}/{rtfarmerindem.cNonRecove}where {left(rtfarmerindem.sClass,2) ='02'}
'    'Line 06 ACV Actual Cash Value loss is (Line 3 minus Line 4 and Line 5)
'    mcDwellACV              '{indemnity.ACVClaim} where {left(indemnity.class,2) ='01'}/{rtfarmerindem.cACVClaim}where {left(rtfarmerindem.sClass,2) ='01'}
'    mcContACV               '{indemnity.ACVClaim} where {left(indemnity.class,2) ='02'}/{rtfarmerindem.cACVClaim}where {left(rtfarmerindem.sClass,2) ='02'}
'    'XS Limits
'    mcDwellXSLim              '{indemnity.LessExcessLimit} where {left(indemnity.class,2) ='01'}/{rtfarmerindem.cExcessLim}where {left(rtfarmerindem.sClass,2) ='01'}
'    mcContXSLim              '{indemnity.LessExcessLimit} where {left(indemnity.class,2) ='02'}/{rtfarmerindem.cExcessLim}where {left(rtfarmerindem.sClass,2) ='02'}
'    'Acv Less XS Limits (Deductible not applied yet)
'    mcDwellACVLessXSLim      '{indemnity.ACVLessExcessLimit} where {left(indemnity.class,2) ='01'}/{rtfarmerindem.cACVLessEx}where {left(rtfarmerindem.sClass,2) ='01'}
'    mcContACVLessXSLim        '{indemnity.ACVLessExcessLimit} where {left(indemnity.class,2) ='02'}/{rtfarmerindem.cACVLessEx}where {left(rtfarmerindem.sClass,2) ='02'}
    On Error GoTo EH
    Dim lClassCount As Long
    Dim sClass As String
    Dim RS As DAO.Recordset
    Dim sSQL As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    'Make sure these are clear
    mcDwellFullRC = 0
    mcDwellRecDep = 0
    mcDwellNonRecDep = 0
    mcDwellACV = 0
    mcDwellXSLim = 0
    mcDwellACVLessXSLim = 0
    mcContFullRC = 0
    mcContRecDep = 0
    mcContNonRecDep = 0
    mcContACV = 0
    mcContXSLim = 0
    mcContACVLessXSLim = 0
    
    'Adjuster Access DB
    If Not mCurDB Is Nothing Then
        For lClassCount = 1 To 2
            'A = COV. A - BUILDING
            'E = COV. E - COMMERCIAL BUILDING
            'C = COV. C - PERSONAL PROPERTY COVERAGES
            sClass = Switch(lClassCount = 1, "'A', 'E'", _
                            lClassCount = 2, "'C'", _
                            True, vbNullString)
            sSQL = "SELECT A.ReplacementCost, "
            sSQL = sSQL & "A.RecoverableDep, "
            sSQL = sSQL & "A.NonRecoverableDep, "
            sSQL = sSQL & "A.ACVClaim, "
            sSQL = sSQL & "A.ExcessLimits, "
            sSQL = sSQL & "A.ACVLessExcessLimits "
            sSQL = sSQL & "FROM     RTIndemnity A "
            sSQL = sSQL & "WHERE    A.AssignmentsID = " & msAssignmentsID & " "
            sSQL = sSQL & "AND      A.ClassOfLossID IN "
                                    sSQL = sSQL & "( "
                                    sSQL = sSQL & "SELECT   ClassOfLossID "
                                    sSQL = sSQL & "FROM     ClassOfLoss "
                                    sSQL = sSQL & "WHERE    ClassTypeID IN "
                                                            sSQL = sSQL & "( "
                                                            sSQL = sSQL & "SELECT   ClassTypeID "
                                                            sSQL = sSQL & "FROM     ClassType "
                                                            sSQL = sSQL & "WHERE    Class IN(" & sClass & ") "
                                                            sSQL = sSQL & ") "
                                    sSQL = sSQL & ") "
            Set RS = mCurDB.OpenRecordset(sSQL)
            If Not RS.EOF Then
                RS.MoveFirst
                Do Until RS.EOF
                    If lClassCount = 1 Then
                        mcDwellFullRC = mcDwellFullRC + IIf(IsNull(RS!ReplacementCost), 0, RS!ReplacementCost)
                        mcDwellRecDep = mcDwellRecDep + IIf(IsNull(RS!RecoverableDep), 0, RS!RecoverableDep)
                        mcDwellNonRecDep = mcDwellNonRecDep + IIf(IsNull(RS!NonRecoverableDep), 0, RS!NonRecoverableDep)
                        mcDwellACV = mcDwellACV + IIf(IsNull(RS!ACVClaim), 0, RS!ACVClaim)
                        mcDwellXSLim = mcDwellXSLim + IIf(IsNull(RS!ExcessLimits), 0, RS!ExcessLimits)
                        mcDwellACVLessXSLim = mcDwellACVLessXSLim + IIf(IsNull(RS!ACVLessExcessLimits), 0, RS!ACVLessExcessLimits)
                    ElseIf lClassCount = 2 Then
                        mcContFullRC = mcContFullRC + IIf(IsNull(RS!ReplacementCost), 0, RS!ReplacementCost)
                        mcContRecDep = mcContRecDep + IIf(IsNull(RS!RecoverableDep), 0, RS!RecoverableDep)
                        mcContNonRecDep = mcContNonRecDep + IIf(IsNull(RS!NonRecoverableDep), 0, RS!NonRecoverableDep)
                        mcContACV = mcContACV + IIf(IsNull(RS!ACVClaim), 0, RS!ACVClaim)
                        mcContXSLim = mcContXSLim + IIf(IsNull(RS!ExcessLimits), 0, RS!ExcessLimits)
                        mcContACVLessXSLim = mcContACVLessXSLim + IIf(IsNull(RS!ACVLessExcessLimits), 0, RS!ACVLessExcessLimits)
                    End If
                    RS.MoveNext
                Loop
            End If
            RS.Close
        Next
        Set RS = Nothing
    Else ' Web DB
        For lClassCount = 1 To 2
            'A = COV. A - BUILDING
            'E = COV. E - COMMERCIAL BUILDING
            'C = COV. C - PERSONAL PROPERTY COVERAGES
            sClass = Switch(lClassCount = 1, "'A', 'E'", _
                            lClassCount = 2, "'C'", _
                            True, vbNullString)
            sSQL = "SELECT A.ReplacementCost, "
            sSQL = sSQL & "A.RecoverableDep, "
            sSQL = sSQL & "A.NonRecoverableDep, "
            sSQL = sSQL & "A.ACVClaim, "
            sSQL = sSQL & "A.ExcessLimits, "
            sSQL = sSQL & "A.ACVLessExcessLimits "
            sSQL = sSQL & "FROM     RTIndemnity A "
            sSQL = sSQL & "WHERE    A.AssignmentsID = " & msAssignmentsID & " "
            sSQL = sSQL & "AND      A.ClassOfLossID IN "
                                    sSQL = sSQL & "( "
                                    sSQL = sSQL & "SELECT   ClassOfLossID "
                                    sSQL = sSQL & "FROM     ClassOfLoss "
                                    sSQL = sSQL & "WHERE    ClassTypeID IN "
                                                            sSQL = sSQL & "( "
                                                            sSQL = sSQL & "SELECT   ClassTypeID "
                                                            sSQL = sSQL & "FROM     ClassType "
                                                            sSQL = sSQL & "WHERE    Class IN(" & sClass & ") "
                                                            sSQL = sSQL & ") "
                                    sSQL = sSQL & ") "
            
            Set mRS = New ADODB.Recordset
            mRS.CursorLocation = adUseClient
            mRS.Open sSQL, mConn, adOpenStatic, adLockReadOnly
            Set mRS.ActiveConnection = Nothing
            
            If Not mRS.EOF Then
                mRS.MoveFirst
                Do Until mRS.EOF
                    If lClassCount = 1 Then
                        mcDwellFullRC = mcDwellFullRC + IIf(IsNull(mRS!ReplacementCost), 0, mRS!ReplacementCost)
                        mcDwellRecDep = mcDwellRecDep + IIf(IsNull(mRS!RecoverableDep), 0, mRS!RecoverableDep)
                        mcDwellNonRecDep = mcDwellNonRecDep + IIf(IsNull(mRS!NonRecoverableDep), 0, mRS!NonRecoverableDep)
                        mcDwellACV = mcDwellACV + IIf(IsNull(mRS!ACVClaim), 0, mRS!ACVClaim)
                        mcDwellXSLim = mcDwellXSLim + IIf(IsNull(mRS!ExcessLimits), 0, mRS!ExcessLimits)
                        mcDwellACVLessXSLim = mcDwellACVLessXSLim + IIf(IsNull(mRS!ACVLessExcessLimits), 0, mRS!ACVLessExcessLimits)
                    ElseIf lClassCount = 2 Then
                        mcContFullRC = mcContFullRC + IIf(IsNull(mRS!ReplacementCost), 0, mRS!ReplacementCost)
                        mcContRecDep = mcContRecDep + IIf(IsNull(mRS!RecoverableDep), 0, mRS!RecoverableDep)
                        mcContNonRecDep = mcContNonRecDep + IIf(IsNull(mRS!NonRecoverableDep), 0, mRS!NonRecoverableDep)
                        mcContACV = mcContACV + IIf(IsNull(mRS!ACVClaim), 0, mRS!ACVClaim)
                        mcContXSLim = mcContXSLim + IIf(IsNull(mRS!ExcessLimits), 0, mRS!ExcessLimits)
                        mcContACVLessXSLim = mcContACVLessXSLim + IIf(IsNull(mRS!ACVLessExcessLimits), 0, mRS!ACVLessExcessLimits)
                    End If
                    mRS.MoveNext
                Loop
            End If
            mRS.Close
        Next
        Set RS = Nothing
    End If
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsRptDB_ClassName & vbCrLf & "Private Sub PopulateIndemValues"
End Sub

Private Sub PopulateInitDeductibles()
    On Error GoTo EH
    Dim cDed As Currency
    Dim cRemainDed As Currency
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    'Set the deductible to the initial deductible
    cDed = mcDeductible
    
    'Figure the Dwelling Deductible
    If mcDwellACVLessXSLim > cDed Then
        mcDwellDed = cDed - mcDwellXSLim
        If mcDwellDed < 0 Then
            mcDwellDed = 0
        End If
    Else
        cRemainDed = cDed - mcDwellACVLessXSLim
        mcDwellDed = cDed - cRemainDed
    End If
    
    'Set the deductible to the Remaining deductible
    cDed = cRemainDed
    
    'Figure the Contents Deductible
    If mcContACVLessXSLim > cDed Then
        mcContDed = cDed - mcContXSLim
        If mcContDed < 0 Then
            mcContDed = 0
        End If
    Else
        cRemainDed = cDed - mcContACVLessXSLim
        mcContDed = cDed - cRemainDed
    End If
    
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsRptDB_ClassName & vbCrLf & "Private Sub PopulateInitDeductibles"
End Sub

Private Function GetDwellDed() As Currency
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    'Line 8

    If mcDwellFullRC = 0 Then
        GetDwellDed = 0
    ElseIf mcContACV = 0 Then
        GetDwellDed = mcDwellDed
    ElseIf GetDwellDed1 + GetContDed1 <= mcDeductible Then
        GetDwellDed = mcDwellDed
    End If
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsRptDB_ClassName & vbCrLf & "Private Function GetDwellDed"
End Function

Private Function GetContDed() As Currency
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    'Line 8 Contents

    If mcContFullRC = 0 Then
        GetContDed = 0
    ElseIf GetDwellDed1 = mcDeductible Then
        GetContDed = 0
    ElseIf GetDwellDed1 = 0 Then
        GetContDed = mcContDed
    ElseIf GetDwellDed1 + GetContDed1 < mcDeductible Then
        GetContDed = mcDeductible - GetDwellDed1
    ElseIf GetDwellDed1 < mcDeductible Then
        GetContDed = GetContDed1
    End If
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsRptDB_ClassName & vbCrLf & "Private Function GetContDed"
End Function

Private Function GetDwellDed1() As Currency
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    If mcDwellACV > mcDeductible Then
        GetDwellDed1 = mcDeductible
    Else
        GetDwellDed1 = mcDwellACV
    End If
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsRptDB_ClassName & vbCrLf & "Private Function GetDwellDed1"
End Function

Private Function GetContDed1() As Currency
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    If mcContACV < mcDeductible - GetDwellDed1 Then
        GetContDed1 = GetACVClaimCont
    ElseIf GetDwellDed1 < mcDeductible Then
        GetContDed1 = mcDeductible - GetDwellDed1
    ElseIf GetDwellDed1 = mcDeductible Then
        GetContDed1 = 0
    ElseIf mcContACV < mcDeductible - GetDwellDed1 Then
        GetContDed1 = mcContACV
    End If
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsRptDB_ClassName & vbCrLf & "Private Function GetContDed1"
End Function

Private Function GetACVClaimCont() As Currency
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    GetACVClaimCont = mcContACV - mcContDed
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsRptDB_ClassName & vbCrLf & "Private Function GetACVClaimCont"
End Function

Private Function GetDwellACV() As Currency
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    'Line 9 Dwelling
    GetDwellACV = mcDwellACV - mcDwellDed
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsRptDB_ClassName & vbCrLf & "Private Function GetDwellACV"
End Function

Private Function GetContACV() As Currency
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    'Line 9 Contents
    GetContACV = mcContACV - mcContDed
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsRptDB_ClassName & vbCrLf & "Private Function GetContACV"
End Function

Private Function GetSuppClaim() As Currency
    On Error GoTo EH
    'Line 10
    Dim cTotRecDep As Currency
    Dim cUnusedDed As Currency
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    cTotRecDep = mcDwellRecDep + mcContRecDep
    cUnusedDed = mcDeductible - mcDwellDed - mcContDed
    
    GetSuppClaim = cTotRecDep - cUnusedDed
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsRptDB_ClassName & vbCrLf & "Private Function GetSuppClaim"
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
    sMiscReportParamName = "MiscReportParam02"
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


