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
Private marRpt As ECRptStateFarm_arRptPhotos.arRptPhotos

Public Type udtPhotos01
    RTPhotoLogID As Long
    fPhotoNo As String
    fPhotodate As Date
    fDesc As String
    imgPhotoPath As String
End Type

Private Type udtarRpt
    f_Name As String
    f_Description As String
    f_Insured As String
    f_ClientClaimNo As String
    f_PolicyNo As String
    f_CatCode As String
    f_DateOfLoss As String
    f_Adjuster As String
    f_DateInspected As Date
    coludtPhotos01 As Collection
End Type
Private msPhotoReposPath As String
Private msNumber As String
Private mbExportXML As Boolean
Private mbExportXMLOnly As Boolean

Public Property Let Number(psNumber As String)
    msNumber = psNumber
End Property
Public Property Get Number() As String
    Number = msNumber
End Property


Public Property Let PhotoReposPath(psPhotoReposPath As String)
    msPhotoReposPath = psPhotoReposPath
End Property
Public Property Get PhotoReposPath() As String
    PhotoReposPath = msPhotoReposPath
End Property

Private Property Get clsRptDB_MiscDelimParamsCol() As Collection
    Dim bCleanup As Boolean
    'Add Report Parameters that are not updated by DB Fields.
    If marRpt Is Nothing Then
        bCleanup = True
        Set marRpt = New ECRptStateFarm_arRptPhotos.arRptPhotos
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
    Dim bPhotoErrors As Boolean
    
    Set marRpt = New ECRptStateFarm_arRptPhotos.arRptPhotos
    
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
        marRpt.SetProperty "f_Name", .f_Name, vbString
        marRpt.SetProperty "f_Description", .f_Description, vbString
        marRpt.SetProperty "f_Insured", .f_Insured, vbString
        marRpt.SetProperty "f_ClientClaimNo", .f_ClientClaimNo, vbString
        marRpt.SetProperty "f_PolicyNo", .f_PolicyNo, vbString
        marRpt.SetProperty "f_CatCode", .f_CatCode, vbString
        marRpt.SetProperty "f_DateOfLoss", .f_DateOfLoss, vbString
        marRpt.SetProperty "f_Adjuster", .f_Adjuster, vbString
        marRpt.SetProperty "f_DateInspected", .f_DateInspected, vbDate
        marRpt.SetProperty "coludtPhotos01", .coludtPhotos01, vbObject
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
    
    'Check for photo Errors
    If Not marRpt.PhotoErrorsCol Is Nothing Then
        FlagServerPhotoErrors marRpt.PhotoErrorsCol
        marRpt.CleanErrorCol
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
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim MyPhoto As udtPhotos01
    Dim sPhotoPath As String
    Dim sPhotoName As String
    
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

    'First get the Photos Header Info
    sSQL = "SELECT "
    sSQL = sSQL & "( "
    If msNumber <> 0 Then
        sSQL = sSQL & "SELECT   Name "
    Else
        sSQL = sSQL & "SELECT Top 1 'ALL' As Name "
    End If
    sSQL = sSQL & "FROM     RTPhotoReport "
    sSQL = sSQL & "WHERE    AssignmentsID = " & msAssignmentsID & " "
    If msNumber <> 0 Then
        sSQL = sSQL & "AND      Number = " & msNumber & " "
    End If
    sSQL = sSQL & ") As Name, "
    sSQL = sSQL & "( "
    If msNumber <> 0 Then
        sSQL = sSQL & "SELECT   Description "
    Else
        sSQL = sSQL & "SELECT  Top 1 'ALL' As Description "
    End If
    sSQL = sSQL & "FROM     RTPhotoReport "
    sSQL = sSQL & "WHERE    AssignmentsID = " & msAssignmentsID & " "
    If msNumber <> 0 Then
        sSQL = sSQL & "AND      Number = " & msNumber & " "
    End If
    sSQL = sSQL & ") As Description, "
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
    sSQL = sSQL & "SELECT  CatCode "
    sSQL = sSQL & "FROM     ClientCompanyCatSpec "
    sSQL = sSQL & "WHERE    ClientCompanyCatSpecID = A.ClientCompanyCatSpecID "
    sSQL = sSQL & ") As ClientCompanyCatSpecCatCode, "
    sSQL = sSQL & "A.CLIENTNUM As ClientClaimNo, "
    sSQL = sSQL & "A.InspectedDate As DateInspected, "
    sSQL = sSQL & "A.LossDate As DateOfLoss, "
    sSQL = sSQL & "A.Insured As Insured, "
    sSQL = sSQL & "A.PolicyNo As PolicyNo "
    sSQL = sSQL & "FROM Assignments A "
    sSQL = sSQL & "WHERE    AssignmentsID = " & msAssignmentsID & " "
    
    Set mRS = New ADODB.Recordset
    mRS.CursorLocation = adUseClient
    mRS.Open sSQL, mConn, adOpenForwardOnly, adLockReadOnly
    Set mRS.ActiveConnection = Nothing
    
    With PopulateudtarRpt
        .f_Name = vbNullString
        .f_Description = vbNullString
        .f_Adjuster = vbNullString
        .f_CatCode = vbNullString
        .f_ClientClaimNo = vbNullString
        .f_DateInspected = NULL_DATE
        .f_DateOfLoss = vbNullString
        .f_Insured = vbNullString
        .f_PolicyNo = vbNullString
        On Error Resume Next
        .f_Name = mRS!Name
        .f_Description = mRS!Description
        .f_Adjuster = mRS!ADJName
        .f_CatCode = mRS!ClientCompanyCatSpecCatCode
        .f_ClientClaimNo = mRS!ClientClaimNo
        .f_DateInspected = mRS!DateInspected
        .f_DateOfLoss = mRS!DateOfLoss
        .f_Insured = mRS!Insured
        .f_PolicyNo = mRS!PolicyNo
        On Error GoTo EH
    End With
    
    sSQL = "SELECT "
    sSQL = sSQL & "PL.[RTPhotoLogID], "
    sSQL = sSQL & "PL.[Description] As [Desc], "
    sSQL = sSQL & "PL.[PhotoDate] As Photodate, "
    sSQL = sSQL & "PL.[SortOrder] As PhotoNo, "
    sSQL = sSQL & "PL.[PhotoName] As imgPhotoPath "
    sSQL = sSQL & "FROM RTPhotoLog PL "
    sSQL = sSQL & "WHERE PL.AssignmentsID = " & msAssignmentsID & " "
    sSQL = sSQL & "AND PL.IsDeleted = 0 "
    If msNumber <> 0 Then
        sSQL = sSQL & "AND      RTPhotoReportID = "
        sSQL = sSQL & "( "
        sSQL = sSQL & "SELECT   RTPhotoReportID "
        sSQL = sSQL & "FROM     RTPhotoReport "
        sSQL = sSQL & "WHERE    AssignmentsID = " & msAssignmentsID & " "
        sSQL = sSQL & "AND      Number = " & msNumber & " "
        sSQL = sSQL & ") "
    End If
    sSQL = sSQL & "ORDER BY PL.SortOrder "
    
    Set mRS = New ADODB.Recordset
    mRS.CursorLocation = adUseClient
    mRS.Open sSQL, mConn, adOpenForwardOnly, adLockReadOnly
    Set mRS.ActiveConnection = Nothing
    
    Set PopulateudtarRpt.coludtPhotos01 = New Collection
    If Not mRS.EOF Then
        mRS.MoveFirst
        'First need to build the Photo Path. The RS will only contain the
        'Name of the photo. depending upon what DB working with the actual
        'Photo path will differ.
        Do Until mRS.EOF
            If goUtil.gMainDB Is Nothing Then
                'SQL Server the Photo path will be
                sPhotoName = Trim(mRS!imgPhotoPath)
                sPhotoPath = Mid(sPhotoName, InStr(1, sPhotoName, "_", vbBinaryCompare) + 1, 4) & "\"
                sPhotoPath = sPhotoPath & Mid(sPhotoName, InStr(1, sPhotoName, "_", vbBinaryCompare) + 5, 2) & "\" & sPhotoName
                sPhotoPath = msPhotoReposPath & sPhotoPath
            Else
                'Access DB
                sPhotoName = Trim(mRS!imgPhotoPath)
                sPhotoPath = goUtil.PhotoReposPath & sPhotoName
            End If
            With MyPhoto
                .RTPhotoLogID = 0
                .fDesc = vbNullString
                .fPhotodate = NULL_DATE
                .fPhotoNo = vbNullString
                .imgPhotoPath = vbNullString
                On Error Resume Next
                .RTPhotoLogID = mRS!RTPhotoLogID
                .fDesc = Trim(mRS!Desc)
                .fPhotodate = mRS!Photodate
                .fPhotoNo = mRS!PhotoNo
                .imgPhotoPath = sPhotoPath
                On Error GoTo EH
            End With
            PopulateudtarRpt.coludtPhotos01.Add MyPhoto, MyPhoto.imgPhotoPath
            mRS.MoveNext
        Loop
    End If
    mRS.Close
    Set mRS = Nothing
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Set mRS = Nothing
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsRptDB_ClassName & vbCrLf & "Private Function PopulateudtarRpt"
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
    sMiscReportParamName = "MiscReportParam08"
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
    'Need to Include the Number since this is a Multi Item
    If msNumber <> 0 Then
        sSQL = sSQL & "AND Number = " & msNumber & " "
    End If
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
                    .Number = msNumber
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

Private Function FlagServerPhotoErrors(pColPhotoErrors As Collection) As Boolean
    On Error GoTo EH
    Dim sSQL As String
    Dim colPhotos As Collection
    Dim vPhoto As Variant
    Dim MyPhoto As udtPhotos01
    Dim sPhotoPath As String
    Dim sPhotoName As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    'Check to see if this is Access DB.
    'If it is need to open connection to it
    If Not goUtil.gMainDB Is Nothing Then
        'Only flag photo erros on Server...
        'Force Adj Client App to Upload Photos again
        'photos may have become corrupted durring previous adj upload attempt.
        'IE adjuster lost connectin in middle of update or had spotty connection
        'in middle of a photo upload.
        Exit Function
     Else
        'WEB Server Connection
         clsRptDB_OpenConnection
    End If
    
    Set colPhotos = pColPhotoErrors
    
    For Each vPhoto In colPhotos
        MyPhoto = vPhoto
        sSQL = "UPDATE RTPhotoLog SET "
        sSQL = sSQL & "[DownLoadMe] = 1, "
        sSQL = sSQL & "[UpLoadPhoto]= 1, "
        sSQL = sSQL & "[UpLoadPhotoThumb]= 1, "
        sSQL = sSQL & "[UpLoadPhotoHighRes]= 1, "
        sSQL = sSQL & "[UpLoadMe]= 1 "
        sSQL = sSQL & "WHERE [AssignmentsID] = " & msAssignmentsID & " "
        sSQL = sSQL & "AND [RTPhotoLogID] = " & CStr(MyPhoto.RTPhotoLogID) & " "
        
        mConn.Execute sSQL
    Next
     
    'cleanup
    Set colPhotos = Nothing
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Set mRS = Nothing
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsRptDB_ClassName & vbCrLf & "Private Function FlagServerPhotoErrorst"
End Function

