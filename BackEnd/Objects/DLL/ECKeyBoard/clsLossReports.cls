VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLossReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'EVENTS
'Unload the Loss Report
Public Event CleanUpLossReports()

'Populate ADO Loss Report Recordset
'This Event will be sent back to what ever application is trying to Load Loss Reports
'As of 1.19.2004 that would be Webcontrol or Easy Claim.  Webcontrol will populate
'The exact same ADO Recordset Object that Easy Claim does.  Each application is Responsible
'for the DB connections necesary to get the Loss Report Data.
Public Event PopulateLossReportsRS(ByRef RSLossReports As ADODB.Recordset)

'PrintAppDoc is  an event raised in the implementing Class (or form)
'The implementing class will be responsible for printing these additional docs
Public Event PrintAppDoc(ByVal colAppDocs As Collection, ByVal sAPPDocClientClaimNo As String)

'Update WebControl or Easy Claim Database
Public Event UpdateDB(ByVal oLossReport As V2ECKeyBoard.clsCarLR)

Public Event GetAssignByZip(ByRef bAssignByZIP As Boolean)

''This can be fired whenever we do DB update then need to refresh
''any List Views or whatever
'Public Event RePopulate(ByVal FormName As String)

'Memory Clean up Event
Public Event MemoryCleanUpAlert()
'Meneory Clean up Finished
Public Event MemoryCleanUpFinished()

'Error Event
Public Event ErrorMess(ByVal Mess As String)

'VVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV
Private msAPPEXEName As String 'THIS IDENTIFIES CARRRIER OBJECT !! IE FARMERS, AMFAM, ETC. ETC.
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'ID's
Private mdNextUKADJ As Double ' Holds Next Number if adding UNKNOWN Adjusters into adjuster collection
Private mdIBSuffix As Double 'Holds IbSuffix
Private mlNextID As Long 'Used to Store ID Used in SQL Building

'Collections
Private mcolAdjusters As Collection 'Used to look up against Adjuster Table to see if Adjuster Is missing
Private mcolRawLossReports As Collection 'This is a collection of Loss reports (Varying formats) That have just been processed
Private mcolAppDocs As Collection 'Contains a collection Docs that may be selected to be appended to each print job
Private mcolParameters As Collection 'Collection of Misc Parameters

'Forms
Private mfrmLoss As frmLossReports 'This form shows Loss reports to be printed

'Objects
Private moFV As V2ECKeyBoard.clsFileVersion

'Flags
Private mbAppDocFlag As Boolean
Private mbMissingAdjusters As Boolean

'Chain Report Memeber variables
Private moChainReport As Object ' Will Hold initial Active report and then Additional Reports will be added to its SubReport Collection.
Private msChainReportName As String
Private msChainLeftName As String
Private msChainRightName As String
Private msChainAdjuster As String
Private mbIgnoreProcessRawDataErrors As Boolean
'Private moPDFDoc As CAcroPDDoc 'Create using Late Binding ONLY!!!
'Private moPDFAVDoc As CAcroAVDoc  'Create using Late Binding ONLY!!!

Public Property Get IgnoreProcessRawDataErrors() As Boolean
    IgnoreProcessRawDataErrors = mbIgnoreProcessRawDataErrors
End Property
Public Property Let IgnoreProcessRawDataErrors(pbFlag As Boolean)
    mbIgnoreProcessRawDataErrors = pbFlag
End Property

Public Property Get ClassName() As String
    ClassName = App.EXEName & ".clsLossReports"
End Property

Public Property Let APPEXEName(psAppEXEName As String)
    msAPPEXEName = psAppEXEName
End Property
Public Property Get APPEXEName() As String
    APPEXEName = msAPPEXEName
End Property

Public Property Get ParametersCol() As Collection
    Set ParametersCol = mcolParameters
End Property
    
Public Property Get NextID() As Long
    On Error GoTo EH
    Dim Param As udtParameter
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    If mlNextID = 0 Then
        Param = mcolParameters("MAX_ID")
        mlNextID = Param.ParamValue
    End If
    mlNextID = mlNextID + 1
    NextID = mlNextID
    Exit Property
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Public Property Get NextID"
End Property

Public Property Get IBSuffix() As String
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    'BGS 11.20.2001 Since we are entering in multiple ClientBatchesDB all at the same time
    'the IBID has to increment starting with the initial time Hack
    If mdIBSuffix = 0 Then
        mdIBSuffix = Format(Now, "YYMMDDHHMMSS")
        IBSuffix = CStr(mdIBSuffix)
    Else 'BGS if we already have the first suffix need to increment the next one
        mdIBSuffix = mdIBSuffix + 1
        IBSuffix = CStr(mdIBSuffix)
    End If
    
    Exit Property
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Public Property Get IBSuffix"
End Property
Public Property Get NextUKADJ() As Long
    mdNextUKADJ = mdNextUKADJ + 1
    NextUKADJ = mdNextUKADJ
End Property

Public Property Let ChainReportName(psName As String)
    msChainReportName = psName
End Property
Public Property Get ChainReportName() As String
    If msChainReportName = vbNullString Then
        ChainReportName = msChainLeftName & msChainRightName
    Else
        ChainReportName = msChainReportName
    End If
End Property

Public Property Let ChainAdjuster(psName As String)
    msChainAdjuster = psName
End Property
Public Property Get ChainAdjuster() As String
    ChainAdjuster = msChainAdjuster
End Property

Public Property Let ChainReport(poActiveReport As Object)
    Set moChainReport = poActiveReport
End Property
Public Property Set ChainReport(poActiveReport As Object)
    Set moChainReport = poActiveReport
End Property
Public Property Get ChainReport() As Object
    Set ChainReport = moChainReport
End Property

Public Property Let MissingADJ(pbFlag As Boolean)
    mbMissingAdjusters = pbFlag
End Property
Public Property Get MissingADJ() As Boolean
    MissingADJ = mbMissingAdjusters
End Property

Public Property Let AppDocsCol(pcolAppDocs As Collection)
    Set mcolAppDocs = pcolAppDocs
    If Not mcolAppDocs Is Nothing Then
        mbAppDocFlag = True
    End If
End Property
Public Property Set AppDocsCol(pcolAppDocs As Collection)
    Set mcolAppDocs = pcolAppDocs
    If Not mcolAppDocs Is Nothing Then
        mbAppDocFlag = True
    End If
End Property
'Only expose this internally as External will only be sent through the Public Event PrintAppDoc
Public Property Get AppDocsCol() As Collection
    Set AppDocsCol = mcolAppDocs
End Property

Public Property Get RawLossReportsCol() As Collection
    Set RawLossReportsCol = mcolRawLossReports
End Property

Public Property Get AppDocFlag() As Boolean
    AppDocFlag = mbAppDocFlag
End Property

Public Property Get AdjustersCol() As Collection
    Set AdjustersCol = mcolAdjusters
End Property
Public Property Let AdjustersCol(pcolAdjusters As Collection)
    Set mcolAdjusters = pcolAdjusters
End Property
Public Property Set AdjustersCol(pcolAdjusters As Collection)
    Set mcolAdjusters = pcolAdjusters
End Property

Public Property Get LossReportsForm() As Object
    Set LossReportsForm = mfrmLoss
End Property

Public Property Get goForms() As Object
    Set goForms = Forms
End Property

Public Function BuildChain(poActiveReport As Object, plChainType As Long, _
                           psCaption As String, psAdjuster As String) As Boolean
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    'Pass in first and Next in link to be added to moChainReport
    'msChainReportName
    Select Case plChainType
        Case ChainType.FirstInChain
SETCHAIN:
            Set moChainReport = poActiveReport
            msChainLeftName = psCaption
            msChainAdjuster = psAdjuster
        Case ChainType.NextLink
            If moChainReport Is Nothing Then
                GoTo SETCHAIN
            End If
            moChainReport.AddChainReport poActiveReport
            msChainRightName = " To " & psCaption
    End Select
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Public Function BuildChain"
End Function

Public Function AddRawLossReport(poLR As V2ECKeyBoard.clsCarLR, Optional pbMissingAdjuster As Boolean = False) As Boolean
    On Error GoTo EH
    Dim bError As Boolean
    Dim sMess As String
    Dim sPath As String
    Dim lRet As Long
    Dim sErrorMess As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    'Be sure to only set the RawLossReport col here
    'This is important if we are showing the loss reports
    'just after process Raw Reports vs showing Loss Reports any other time
    If mcolRawLossReports Is Nothing Then
        Set mcolRawLossReports = New Collection
    End If
    
    'Be sure to set the reference to this instance of
    'Loss reports class
    Set poLR.LRs = Nothing
    Set poLR.LRs = Me
    
    mcolRawLossReports.Add poLR, poLR.PrnKey
    
    'Need to Raise Event to Update an Implementing App's(IE WebControl EasyClaim) Database.
    RaiseEvent UpdateDB(poLR)
    
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    If lErrNum = ERROR_CODE_KEY_ALREADY_ASSOCIATED Then
        'The raw data we are processing may have multiple
        'instances of the exact same Loss.
        'We will change the name of the file and save it as
        'A duplicate.
        bError = True
        sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
        sMess = sMess & "ERROR # " & Err.Number & vbCrLf
        sMess = sMess & sErrDesc & vbCrLf
        sMess = sMess & "Report Format: " & poLR.ClassName & vbCrLf
        sMess = sMess & """" & poLR.PrnKey & """" & vbCrLf
        sMess = sMess & "This is a duplicate Loss Report assigned to: " & poLR.ACID & vbCrLf
        sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
        If StrComp(poLR.ClassName, "V2ECCarFarmers.clsLossXML01", vbTextCompare) = 0 Then
            Err.Clear
            poLR.Status = "AddToExistingLoss"
            RaiseEvent UpdateDB(poLR)
            Exit Function
        Else
            Err.Clear
        End If
        RaiseEvent ErrorMess(sMess)
        sErrorMess = "Click ""Yes"" to not prompt for subsequent Errors and continue processing." & vbCrLf & vbCrLf
        sErrorMess = sErrorMess & "Click ""No"" to continue processing as well as continue to prompt subsequent Errors." & vbCrLf & vbCrLf
        sErrorMess = sErrorMess & "Click ""Cancel"" to STOP processing all together."
        If Not IgnoreProcessRawDataErrors Then
            lRet = MsgBox(sMess & vbCrLf & vbCrLf & vbCrLf & sErrorMess, vbYesNoCancel, "Error")
            If lRet = vbYes Then
                IgnoreProcessRawDataErrors = True
            ElseIf lRet = vbNo Then
                 IgnoreProcessRawDataErrors = False
            ElseIf lRet = vbCancel Then
                IgnoreProcessRawDataErrors = True
                poLR.AbortProcessRawData = True
            End If
        End If
        Exit Function
    End If
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Public Function AddRawLossReport"
End Function

Public Sub FireCleanUpLossReports()
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    RaiseEvent CleanUpLossReports
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Public Sub FireCleanUpLossReports"
End Sub

Public Sub FirePrintAppDoc(ByVal psClientClaimNo As String)
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    RaiseEvent PrintAppDoc(ByVal mcolAppDocs, psClientClaimNo)
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Public Sub FirePrintAppDoc"
End Sub

Public Sub FireMemoryCleanUpAlert()
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    RaiseEvent MemoryCleanUpAlert
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Public Sub FireMemoryCleanUpAlert"
End Sub

Public Sub FireMemoryCleanUpFinished()
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    RaiseEvent MemoryCleanUpFinished
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Public Sub FireMemoryCleanUpFinished"
End Sub


Public Sub FireErrorMess(psMess As String)
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    RaiseEvent ErrorMess(ByVal psMess)
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Public Sub FireErrorMess"
End Sub

Public Sub FireGetAssignByZip(pbAssignByZip As Boolean)
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    RaiseEvent GetAssignByZip(pbAssignByZip)
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Public Sub FireGetAssignByZip"
End Sub



Public Sub ResetParametersCol()
    On Error GoTo EH
    Dim oMiscObject As Object
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    'Cleanup Parametere collection
    'It may have Objects in it so clean up them
    If Not mcolParameters Is Nothing Then
        On Error Resume Next
        For Each oMiscObject In mcolParameters
            oMiscObject.CLEANUP
            Set oMiscObject = Nothing
        Next
        On Error GoTo EH
        Set mcolParameters = Nothing
    End If
    
    Set mcolParameters = New Collection
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Public Sub ResetParametersCol"
End Sub

Public Function RemoveParameter(pParam As udtParameter) As Boolean
    On Error GoTo EH
    
    If Not mcolParameters Is Nothing Then
        mcolParameters.Remove pParam.ParamName
        RemoveParameter = True
    End If
    
    Exit Function
EH:
    RemoveParameter = False
End Function

Public Function AddParameter(pParam As udtParameter) As Boolean
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    RemoveParameter pParam
    
    If mcolParameters Is Nothing Then
        Set mcolParameters = New Collection
    End If
    
    mcolParameters.Add pParam, pParam.ParamName
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Public Function AddParameter"
End Function

Public Function FindParameter(psName As String) As udtParameter
    On Error GoTo EH
    Dim vParam As Variant
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    If Not mcolParameters Is Nothing Then
        'if this errors then the
        'Parameter is missing
        vParam = mcolParameters(psName)
    Else
        Exit Function
    End If
    
    FindParameter = vParam
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Public Function FindParameter"
End Function

Public Sub ResetAdjusterCol()
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    Set mcolAdjusters = Nothing
    Set mcolAdjusters = New Collection
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Public Sub ResetAdjusterCol"
End Sub

Public Function AddAdjuster(pAdjuster As udtAdjuster) As Boolean
    On Error GoTo EH
    Dim sMess As String
    
    RemoveAdjuster pAdjuster
    
    If Not mcolAdjusters Is Nothing Then
        'Use FACT as key
        mcolAdjusters.Add pAdjuster, pAdjuster.Key
        AddAdjuster = True
    End If
    
    Exit Function
EH:
    If Err.Number = ERROR_CODE_KEY_ALREADY_ASSOCIATED Then
        sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
        sMess = sMess & "ERROR # " & Err.Number & vbCrLf
        sMess = sMess & Err.Description & vbCrLf
        sMess = sMess & "FACT: " & pAdjuster.FACT & vbCrLf
        sMess = sMess & "CRID: " & pAdjuster.CRID & vbCrLf
        sMess = sMess & "EMAIL: " & pAdjuster.EmailAddress & vbCrLf
        sMess = sMess & "Is a duplicate entry for..." & vbCrLf
        sMess = sMess & "FACT: " & mcolAdjusters(pAdjuster.Key).FACT & vbCrLf
        sMess = sMess & "CRID: " & mcolAdjusters(pAdjuster.Key).CRID & vbCrLf
        sMess = sMess & "EMAIL: " & mcolAdjusters(pAdjuster.Key).EmailAddress & vbCrLf
        sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
        
        RaiseEvent ErrorMess(sMess)
    End If
    
    AddAdjuster = False
End Function

Public Function RemoveAdjuster(pAdjuster As udtAdjuster) As Boolean
    On Error GoTo EH
    
    If Not mcolAdjusters Is Nothing Then
        mcolAdjusters.Remove pAdjuster.Key
        RemoveAdjuster = True
    End If
    
    Exit Function
EH:
    RemoveAdjuster = False
    
End Function

Public Function FindAdjuster(psFact As String, Optional pbMissing As Boolean) As udtAdjuster
    On Error GoTo EH
    Dim vAdjuster As Variant
    Dim bMissing As Boolean
    
    'Set pbmissing to false
    pbMissing = bMissing
    
    If Not mcolAdjusters Is Nothing Then
        'if this errors then the
        'adjuster is missing
        vAdjuster = mcolAdjusters.Item(psFact)
    Else
        Exit Function
    End If
    If bMissing Then
        FindAdjuster.FACT = psFact
        FindAdjuster.CRID = psFact
        FindAdjuster.Key = psFact
        FindAdjuster.IsDirty = IsDirty.AddMe
    Else
        FindAdjuster = vAdjuster
    End If
    
    If bMissing Then
        'Set the member flag
        mbMissingAdjusters = True
    End If
    
    Exit Function
EH:
    Err.Clear
    bMissing = True
    pbMissing = bMissing
    Resume Next
End Function

Public Function FindOriginAdjuster(psFact As String, Optional pbMissing As Boolean, _
                                   Optional pcolOrigADJ As Collection) As udtAdjuster
    On Error GoTo EH
    Dim vAdjuster As Variant
    Dim bMissing As Boolean
    'set pbmissing to false
    pbMissing = bMissing
    
    If Not pcolOrigADJ Is Nothing Then
        'if this errors then the
        'adjuster is missing
        vAdjuster = pcolOrigADJ(psFact)
    Else
        Exit Function
    End If
    If bMissing Then
        FindOriginAdjuster.FACT = psFact
        FindOriginAdjuster.CRID = psFact
        FindOriginAdjuster.Key = psFact
    Else
        FindOriginAdjuster = vAdjuster
    End If
    
    Exit Function
EH:
    Err.Clear
    bMissing = True
    pbMissing = bMissing
    Resume Next
End Function
    
Public Function ProcessRawData(psClass As String, psDataPath As String, _
                               psOutPath As String, _
                               Optional poPB As Object, _
                               Optional potxtProgMess As Object) As Boolean
    On Error GoTo EH
    Dim oLR As V2ECKeyBoard.clsCarLR
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    If psClass = vbNullString Then
        Set oLR = New V2ECKeyBoard.clsLossUnknown
    Else
        Set oLR = CreateObject(APPEXEName & "." & psClass)
    End If
    
    oLR.SetUtilObject goUtil
    
    
    Set oLR.LRs = Me  'Reference this instance of clsLossReports inside oLR
    
    'After the Loss Report object has processed the Raw Data the
    'mcolRawLossReports will be populated
    If oLR.ProcessRawData(psDataPath, psOutPath, poPB, potxtProgMess) Then
        ProcessRawData = True
    End If
    
    oLR.CLEANUP
    Set oLR = Nothing
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Public Function ProcessRawData"
End Function

Public Sub ShowLossReports(Optional psPDFOutPath As String, _
                            Optional pProgForm As V2ECKeyBoard.clsProgForm, _
                            Optional pbResetLossCollection As Boolean = True)
    On Error GoTo EH
    Dim sMess As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    If goUtil.utFindSetForm(Forms, "frmLossReports", mfrmLoss) Then
       mfrmLoss.Show vbModeless
       mfrmLoss.WindowState = vbNormal
       Set mfrmLoss.LossReportsCol = Nothing
       mfrmLoss.Populate
       Exit Sub
    End If
    
    Set mfrmLoss = New frmLossReports
    
    With mfrmLoss
        'Set properties
        .PDFOutPath = psPDFOutPath
        Set .ProgForm = pProgForm
        Set .LRs = Me
        'If we just processed some reports use the Raw Loss Reports Collection
        'Saves time having to GetThe loss reports collection from Scratch
        If Not mcolRawLossReports Is Nothing Then
            If pbResetLossCollection Then
                Set mcolRawLossReports = Nothing
            Else
                Set .LossReportsCol = mcolRawLossReports
            End If
        End If
        'Load & Show
        If Not goUtil.gfrmECTray Is Nothing Then
            goUtil.gfrmECTray.FlagLoadingLossReports = True
        End If
        Load mfrmLoss
        If Not goUtil.gfrmECTray Is Nothing Then
            goUtil.gfrmECTray.FlagLoadingLossReports = False
        End If
        
        .Show vbModeless
     
    End With
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Public Sub ShowLossReports"
End Sub

Public Sub ShowHelpViewer(psHelpHTML As String, psHelpCaption As String, Optional pbDel As Boolean)
    On Error GoTo EH
    Dim MyHelpForm As frmHelpViewer
    Dim sHelpFile As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim bPDFFile As Boolean
        
    
    'Check for Temp folder
    If Not goUtil.utFileExists("C:\Temp", True) Then
        On Error GoTo 0
        goUtil.utMakeDir "C:\Temp"
        On Error GoTo EH
    End If
    'Set the file path
    If goUtil.utFileExists("C:\Temp", True) Then
        If InStr(1, psHelpHTML, ".pdf", vbTextCompare) > 0 Then
            sHelpFile = "C:\Temp\PDFViewer_" & CStr(timeGetTime()) & ".pdf"
            bPDFFile = True
        Else
            sHelpFile = "C:\Temp\HelpViewer_" & Format(Now, "YYMMDDHHMMSS") & ".html"
        End If
    Else
        sHelpFile = App.Path & "\HelpViewer_" & Format(Now, "YYMMDDHHMMSS") & ".html"
    End If
    
    If bPDFFile Then
        goUtil.utCopyFile psHelpHTML, sHelpFile
    Else
        'Save Help file
        goUtil.utSaveFileData sHelpFile, psHelpHTML
    End If
    
    
    Set MyHelpForm = New frmHelpViewer
    
    With MyHelpForm
        .Helpfile = sHelpFile
        .HelpCaption = psHelpCaption
        Load MyHelpForm
        .Show vbModal
        Unload MyHelpForm
        Set MyHelpForm = Nothing
    End With
    
    'Get rid of it
    If pbDel Then
        If goUtil.utFileExists(sHelpFile) Then
            SetAttr sHelpFile, vbNormal
            goUtil.utDeleteFile sHelpFile
        End If
    End If
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Public Sub ShowHelpViewer"
End Sub

Public Function CLEANUP() As Boolean
    On Error GoTo EH
    Dim oLR As V2ECKeyBoard.clsCarLR
    Dim oMiscObject As Object
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    'Cleanup Parametere collection
    'It may have Objects in it so clean up them
    If Not mcolParameters Is Nothing Then
        On Error Resume Next
        For Each oMiscObject In mcolParameters
            oMiscObject.CLEANUP
            Set oMiscObject = Nothing
        Next
        On Error GoTo EH
        Set mcolParameters = Nothing
    End If
    
    Set mcolParameters = Nothing
    Set moChainReport = Nothing
    Set mcolAdjusters = Nothing
    Set moFV = Nothing
    'Clean up each Loss Report object within Loss reports Collection
    'This should free up allocated memory  better than if we just set the
    'Collection To nothing
    If Not mcolRawLossReports Is Nothing Then
        For Each oLR In mcolRawLossReports
            If Not oLR Is Nothing Then
                oLR.CLEANUP
                Set oLR = Nothing
            End If
        Next
        Set mcolRawLossReports = Nothing
    End If
    
    If Not mfrmLoss Is Nothing Then
        mfrmLoss.CLEANUP
        Unload mfrmLoss
        Set mfrmLoss = Nothing
    End If
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Public Function CleanUp"
End Function

Public Function GetLossReportCol(Optional poPB As V2ECKeyBoard.clsProgForm) As Collection
    On Error GoTo EH
    Dim colLossReports As Collection
    Dim RSLossReports As ADODB.Recordset
    Dim sPrnKey As String
    Dim sPrnData As String
    Dim sClass As String
    Dim lPos As Long
    Dim oLR As V2ECKeyBoard.clsCarLR
    Dim bPB As Boolean 'True if using Prog Bar
    Dim lMax As Long
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    Dim sAttachName As String
    Dim sAttachPath As String
    Dim sAttachReposPath As String
    
    
    'SET the Loss Report RecordSet
    Set RSLossReports = New ADODB.Recordset
    RaiseEvent PopulateLossReportsRS(RSLossReports)
    If Not RSLossReports Is Nothing Then
        If Not RSLossReports.State = adStateClosed Then
            If Not RSLossReports.EOF Then
                RSLossReports.MoveFirst
            Else
                GoTo CLEANUP
            End If
        Else
            GoTo CLEANUP
        End If
    Else
        GoTo CLEANUP
    End If

    'Set ProgBar flag
    If Not poPB Is Nothing Then
        bPB = True
    End If

    'Progress bar
    If bPB Then
        poPB.ShowForm True
        poPB.PBarFile.Max = RSLossReports.RecordCount
        poPB.PBarFile.Value = 0
    End If

    'Go through every file and create the appropriate class object
    'and add that object to the Loss Reports Collection colLossReports
    Do Until RSLossReports.EOF
        sPrnData = goUtil.IsNullIsVbNullString(RSLossReports.Fields("LossReport"))
        sClass = goUtil.IsNullIsVbNullString(RSLossReports.Fields("LRFormat"))
        If InStr(1, left(sClass, 5), "ECcar", vbTextCompare) = 1 Then
            sClass = Replace(sClass, left(sClass, 5), goUtil.gsCarPrefix, , , vbTextCompare)
        End If
        
        'If we get an Error then go to Unknown
        On Error Resume Next
        If InStr(1, sClass, ".", vbBinaryCompare) > 0 Then
            Set oLR = CreateObject(sClass)
        Else
            GoTo UNKNOWN_LOSS
        End If
        If Err.Number > 0 Then GoTo UNKNOWN_LOSS
        On Error GoTo EH
        GoTo SET_LOSSREPORT
UNKNOWN_LOSS:
        If Err.Number > 0 Then
            Err.Clear
        End If
        On Error GoTo EH
        Set oLR = New V2ECKeyBoard.clsLossUnknown
SET_LOSSREPORT:
        'Set up Loss Report Object with Data
        With oLR
            .SetUtilObject goUtil
            Set .LRs = Me
            .PrnData = sPrnData
            .PopulateMe
            'Since getting this lossreport from the database,
            'need to populate some items with database fields even though PopuateME()
            'will have already filled in these items.
            'This is because the loss report contains orginal data.
            'The database will contain the latest info
            'the original loss report prndata WILL NOT BE UPDATED...
            'Unless it contains PDF FileName.
            .Adjuster = goUtil.IsNullIsVbNullString(RSLossReports.Fields("AdjusterSpecUserName"))
            .ACID = goUtil.IsNullIsVbNullString(RSLossReports.Fields("AdjusterSpecACID"))
            .AssignmentType = goUtil.IsNullIsVbNullString(RSLossReports.Fields("AssignmentTypeType"))
            .CatCode = goUtil.IsNullIsVbNullString(RSLossReports.Fields("ClientCompanyCatSpecCatCode"))
            .CatName = goUtil.IsNullIsVbNullString(RSLossReports.Fields("CatName"))
            .CLIENTNUM = goUtil.IsNullIsVbNullString(RSLossReports.Fields("CLIENTNUM"))
            .DateAssign = goUtil.IsNullIsVbNullString(RSLossReports.Fields("AssignedDate"))
            .HomePhone = goUtil.IsNullIsVbNullString(RSLossReports.Fields("HomePhone"))
            .IBNUM = goUtil.IsNullIsVbNullString(RSLossReports.Fields("IBNUM"))
            .InsuredName = goUtil.IsNullIsVbNullString(RSLossReports.Fields("Insured"))
            .LossType = goUtil.IsNullIsVbNullString(RSLossReports.Fields("AssignmentTypeType"))
            .Status = goUtil.IsNullIsVbNullString(RSLossReports.Fields("StatusStatus"))
            .TypeOfACID = goUtil.IsNullIsVbNullString(RSLossReports.Fields("AdjusterSpecAcidDescription"))
            .WorkPhone = goUtil.IsNullIsVbNullString(RSLossReports.Fields("BusinessPhone"))
            .OleType = goUtil.IsNullIsVbNullString(RSLossReports.Fields("LRFormat"))
            'If the OleType is PDF then need to Build the Main Path to the PDF file in front of the
            'PDF File Name.
            If InStr(1, .OleType, "OLEType_pdf", vbTextCompare) > 0 Then
                If InStr(1, goUtil.gsAppEXEName, "WebControl", vbTextCompare) > 0 Then
                    sAttachReposPath = GetSetting(goUtil.gsAppEXEName, "Dir", "FTPSitePath", vbNullString)
                Else
                    sAttachReposPath = goUtil.gsInstallDir
                End If
                'check for backslash in path add it if there is none
                If StrComp(Right(sAttachReposPath, 1), "\", vbBinaryCompare) <> 0 Then
                    sAttachReposPath = sAttachReposPath & "\"
                End If
                sAttachReposPath = sAttachReposPath & "AttachRepos\"
                'If on Web Server need to account for YY MM folder structure
                If InStr(1, goUtil.gsAppEXEName, "WebControl", vbTextCompare) > 0 Then
                    sAttachName = .PrnData
                    sAttachPath = Mid(sAttachName, InStr(1, sAttachName, "_", vbBinaryCompare) + 1, 4) & "\"
                    sAttachPath = sAttachPath & Mid(sAttachName, InStr(1, sAttachName, "_", vbBinaryCompare) + 5, 2) & "\"
                    sAttachPath = sAttachReposPath & sAttachPath
                    sAttachReposPath = sAttachPath
                End If
                .PrnData = sAttachReposPath & .PrnData
            End If
            '"*ADJUSTERUSERNAME*" & "_" & "*ACID*" & "_" & "*IBNUMBER*" & "_" & "*CLIENTNUM*"
            .PrnKey = .Adjuster & "_" & .ACID & "_" & .IBNUM & "_" & .CLIENTNUM
        End With

        'Init Collection first time
        If colLossReports Is Nothing Then
            Set colLossReports = New Collection
        End If
        'Use file Path as the Key (Use this to quickly recall this
        'particular Report Object from this collection later)
        colLossReports.Add oLR, oLR.PrnKey

        If bPB Then
            poPB.lblFileText = oLR.PrnKey
            poPB.RefreshMe
            poPB.PBarFile.Value = poPB.PBarFile.Value + 1
            'See if the user Cancled the process
            DoEvents
            Sleep 10
            If poPB.CancelMe Or goUtil Is Nothing Then
                poPB.CancelMe = False
                GoTo CLEANUP
            End If
        End If
        Set oLR = Nothing
        RSLossReports.MoveNext
    Loop

    If bPB Then
        'Hide the Form but Ignore modality
        poPB.ShowForm False
    End If

    'After don adding Loss Reports objects set this function
    If Not colLossReports Is Nothing Then
        Set GetLossReportCol = colLossReports
    End If

CLEANUP:
    Set RSLossReports = Nothing
    Set colLossReports = Nothing
    Set oLR = Nothing
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Set RSLossReports = Nothing
    Set colLossReports = Nothing
    Set oLR = Nothing
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "public Function GetLossReportCol"
End Function

Public Function GetCreateDate(psFilePath As String) As String
    On Error GoTo EH
    Dim FI As V2ECKeyBoard.FILE_INFORMATION
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    FI = moFV.GetFileInformation(psFilePath)
    GetCreateDate = CStr(Format(FI.dtCreationDate, "MM/DD/YYYY"))
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Public Function GetCreateDate"
End Function

Public Function LoadAppDoc() As Boolean
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    LoadAppDoc = True
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Public Function LoadAppDoc"
End Function

Public Function PrintToSpool(psFilePath As String, psData As String) As Boolean
    On Error GoTo EH
    Dim sFileTemp As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    If InStr(1, psFilePath, ".txt", vbTextCompare) > 0 Then
        sFileTemp = Replace(psFilePath, ".txt", ".tmp")
    Else
        sFileTemp = psFilePath & ".tmp"
    End If
    
    If goUtil.utFileExists(sFileTemp) Then
        SetAttr sFileTemp, vbNormal
        Kill sFileTemp
    End If
    
    goUtil.utSaveFileData sFileTemp, psData
    
    Call SpoolFile(sFileTemp)
    Kill sFileTemp
    
    PrintToSpool = True
       
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Public Function PrintToSpool"
End Function

Public Function ClosePrn() As Long
    ClosePrn = SpoolClosePrn
End Function

Public Function OpenPrn(Optional psPrnDeviceName As String = vbNullString, Optional pbGetDeviceNameOnly As Boolean) As Long
    On Error GoTo EH
    Dim prn As Printer
    Dim sLastSelectedPrinter As String
    Dim Param As V2ECKeyBoard.udtParameter
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    'Find the selected printer against printer list.
    'Do this because it is possible that the printer list
    'could have changed from the last time they loaded the form
    For Each prn In Printers
       If InStr(1, mfrmLoss.cboSelectPrinter.Text, prn.DeviceName) = 1 _
       And Right(mfrmLoss.cboSelectPrinter.Text, Len(prn.Port)) = prn.Port Then
            If psPrnDeviceName = vbNullString Then
                psPrnDeviceName = prn.DeviceName
            End If
            If Not pbGetDeviceNameOnly Then
                OpenPrn = SpoolOpenPrn(psPrnDeviceName)
            End If
            'Need to update Parameters Collection with the current Selected
            'Printer Object
            Param.ParamName = "PRINTER_OBJECT"
            Set Param.ParamValue = prn
            AddParameter Param
    
            Exit Function ' Bail here we are done
       End If
    Next prn
    
    'If we get to here then the user must have changed the printer collection
    'reset it so they do not have to unload and reload the Form.
    'Fill the Printers Combo Box
    mfrmLoss.cboSelectPrinter.Clear
    For Each prn In Printers
        mfrmLoss.cboSelectPrinter.AddItem prn.DeviceName & " on " & prn.Port
    Next prn
    
    sLastSelectedPrinter = GetSetting(App.EXEName, "Printer", "LossReport", vbNullString)
    'Select Deafult printer from list
    SelectDefaultPrinter mfrmLoss.cboSelectPrinter, sLastSelectedPrinter
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Public Function OpenPrn"
End Function

Public Function PrintGenText(poLR As V2ECKeyBoard.clsCarLR, psPrnDeviceName As String, pbPreview As Boolean, _
                             pbPDF As Boolean, plHwnd As Long, plChainType As Long) As Boolean
    On Error GoTo EH
    Dim MyGenText As arGenText
    Dim sPDFPath As String
    Dim sCaption As String
    Dim dDays As Double
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    
    'Build Caption
    sCaption = "Loss Report " & Chr(160) & " (" & poLR.PrnKey & ")"
    
    'BGS 5.9.2002 LOOK FOR <HTML> TAG.
    'Launch the frmHelpViewer. and do not process HTML like other Gen text Reports
    If InStr(1, poLR.PrnData, "<HTML>", vbTextCompare) > 0 Then
        If plChainType = ChainType.NotChain Then
            ShowHelpViewer poLR.PrnData, sCaption, True
        End If
        GoTo CLEANUP
    End If
    
    Set MyGenText = New arGenText
    With MyGenText
        .Printer.Orientation = ddOPortrait
        On Error Resume Next
        .Printer.DeviceName = psPrnDeviceName
        If Err.Number > 0 Then
            Err.Clear
            On Error GoTo EH
        End If
        .fText.Height = poLR.GenTextHeight
        .fText.Font.Name = poLR.GenTextFontName
        .fText.Font.Size = poLR.GenTextFontSize
        Set .LR = poLR 'Set the reports Loss report Object
    End With
    
    'Check to see if we are building a Chain
    If plChainType = ChainType.FirstInChain Or plChainType = ChainType.NextLink Then
        BuildChain MyGenText, plChainType, sCaption, poLR.Adjuster
        GoTo CLEANUP
    End If
    
    'If we are previewing then we need to use ARV object
    If pbPreview Then
        If goUtil.gARV Is Nothing Then
            Set goUtil.gARV = CreateObject("V2ARViewer.clsARViewer")
        End If
        With goUtil.gARV
            .SetUtilObject goUtil
            MyGenText.Run 'Run it before it gets to ARV object
            .objARvReport = MyGenText
            .sRptTitle = sCaption 'Use 160 to limit the Form Posn settings
            .ShowReport vbModeless
        End With
        Screen.MousePointer = vbDefault
    Else 'If we not preview then check to see if this is going to PDF
        If pbPDF Then
            'Set PDF File
            sPDFPath = Trim(mfrmLoss.txtPDFPath.Text)
            sPDFPath = sPDFPath & sCaption
            ExportFile MyGenText, sPDFPath, ExportType.ARPdf 'Export to PDF
        Else
            MyGenText.PrintReport False 'Print without showing print dialog
        End If
    End If
    
CLEANUP:

    PrintGenText = True
    
    'Clean up
    Set MyGenText = Nothing
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Public Function PrintGenText"
End Function

Public Function ExportGenText(poLR As V2ECKeyBoard.clsCarLR, pXportType As ExportType) As Boolean
    On Error GoTo EH
    Dim MyGenText As arGenText
    Dim sXportPath As String
    
    Set MyGenText = New arGenText
    
    With MyGenText
        .Printer.Orientation = ddOPortrait
        .fText.Height = poLR.GenTextHeight
        .fText.Font.Name = poLR.GenTextFontName
        .fText.Font.Size = poLR.GenTextFontSize
        Set .LR = poLR 'Set the reports Loss report Object
    End With
    
    sXportPath = poLR.PrnKey
    If ExportFile(MyGenText, sXportPath, pXportType) Then
        ExportGenText = True
    End If
    
    'Clean up
    Set MyGenText = Nothing
    
    Exit Function
EH:
    Err.Clear
End Function

Public Function PrintOlePDFDoc(poLR As V2ECKeyBoard.clsCarLR, psPrnDeviceName As String, pbPreview As Boolean, _
                             pbPDF As Boolean, plHwnd As Long, plChainType As Long) As Boolean
    On Error GoTo EH
    Dim sCaption As String
    Dim sPDFPath As String
    Dim slpOperation As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    'Build Caption
    sCaption = "Loss Report " & Chr(160) & " (" & poLR.PrnKey & ")"
    If pbPDF Then
        'Set PDF File
        sPDFPath = Trim(mfrmLoss.txtPDFPath.Text)
        goUtil.utCopyFile poLR.PrnData, sPDFPath & sCaption & ".pdf"
    Else
        If pbPreview Then
            slpOperation = "OPEN"
        Else
            slpOperation = "PRINT"
        End If
        goUtil.utShellExecute , slpOperation, poLR.PrnData, , , vbNormalFocus, True, True, True, sCaption
    End If
    
CLEANUP:

    PrintOlePDFDoc = True
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Public Function PrintOlePDFDoc"
End Function

Public Function ExportOlePDFDoc(poLR As V2ECKeyBoard.clsCarLR, pXportType As ExportType) As Boolean
    On Error GoTo EH
'    Dim MyOlePDF As arPDFDoc
'    Dim sXportPath As String
'
'    Set MyOlePDF = New arPDFDoc
'
'    With MyOlePDF
'        .Printer.Orientation = ddOPortrait
'        Set .LR = poLR 'Set the reports Loss report Object
'    End With
'
'    sXportPath = poLR.PrnKey
'    If ExportFile(MyOlePDF, sXportPath, pXportType) Then
'        ExportOlePDFDoc = True
'    End If
'
'    'Clean up
'    Set MyOlePDF = Nothing
    
    Exit Function
EH:
    Err.Clear
End Function

Private Sub Class_Initialize()
    Set moFV = New V2ECKeyBoard.clsFileVersion
    mdNextUKADJ = Format(Now(), "YYMMDDHHMMSS") 'Set this number so it will always be unique
End Sub

Private Function ERROR_CODE_KEY_ALREADY_ASSOCIATED() As Long
    'This key is already associated with an element of this collection
    ERROR_CODE_KEY_ALREADY_ASSOCIATED = 457
End Function

Public Function ExportFile(poAR As Object, psFileName As String, plExpType As Long, Optional plPhotoQuality As Long = 50) As Boolean
    On Error GoTo EH
    Dim oExport As Object
    Dim sFilename As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    sFilename = psFileName
    
    If Not poAR Is Nothing Then
        poAR.Run 'Run the report without print or preview
        If poAR.Pages.Count > 0 Then
            'Get rid of the existsing Ext Name.
            If InStr(1, Right(sFilename, 5), ".", vbBinaryCompare) > 0 Then
                sFilename = left(sFilename, InStrRev(sFilename, ".", , vbBinaryCompare) - 1)
            End If
            Select Case plExpType
                Case ExportType.ARExcel
                    Set oExport = CreateObject("ActiveReportsExcelExport.ARExportExcel")
                    'No photo quality exists for this type
                    sFilename = sFilename & ".xls"
                Case ExportType.ARHtml
                    Set oExport = CreateObject("ActiveReportsHTMLExport.HTMLexport")
                    oExport.JPEGQuality = plPhotoQuality
                    sFilename = sFilename & ".html"
                Case ExportType.ARPdf
                    Set oExport = CreateObject("ActiveReportsPDFExport.ARExportPDF")
                    oExport.JPGQuality = plPhotoQuality
                    sFilename = sFilename & ".pdf"
                Case ExportType.ARRtf
                    Set oExport = CreateObject("ActiveReportsRTFExport.ARExportRTF")
                    'No photo quality exists for this type
                    sFilename = sFilename & ".rtf"
                Case ExportType.ARText
                    Set oExport = CreateObject("ActiveReportsTextExport.ARExportText")
                    'No photo quality exists for this type
                    sFilename = sFilename & ".txt"
            End Select
            oExport.FileName = sFilename 'Set the Export Filename
            oExport.Export poAR.Pages 'Export to one of the above Late bound Types
        End If
    End If
    
    'Cleanup
    Set oExport = Nothing
    ExportFile = True
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Set oExport = Nothing
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Private Function ExportFile"
End Function

Public Function CFCreateExport(ByVal psPrnData As String, ByVal psExportPath As String, ByVal pXportType As Long) As Boolean
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim bCreateExport As Boolean
    'Since CF MX does not recognize Public Enum ExportType
    'Need to use this Wrapper Function which passes Long instead
    
    CFCreateExport = CreateExport(psPrnData, psExportPath, pXportType)
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Public Function CFCreateExport"
End Function

Public Function CreateExport(psPrnData As String, psExportPath As String, pXportType As ExportType) As Boolean
    On Error GoTo EH
    Dim sPrnData As String
    Dim sClass As String
    Dim lPos As Long
    Dim oLR As V2ECKeyBoard.clsCarLR
    Dim sFarmersLRClass As String 'If there is no App.exename in CLass name
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    sPrnData = psPrnData
    
    If sPrnData = vbNullString Then
        Exit Function
    End If
    
    lPos = InStr(1, sPrnData, vbCrLf) - 1
    If lPos < 0 Then
        lPos = 0
    End If
    
    sClass = Trim(left(sPrnData, lPos))
    
    'BGS 4.11.2002
    'Check to see if there is App.exename in there ('.')
    'If not then it was made with an older version clsCarULFarm
    'when it was still in ECKeyBoard So we need to append ECcarFarmers app name.
    If InStr(1, sClass, ".", vbBinaryCompare) = 0 Then
        sFarmersLRClass = sClass
        sClass = "V2ECcarFarmers." & sFarmersLRClass
    ElseIf InStr(1, sClass, "ECFarmers.", vbTextCompare) > 0 Then
        sFarmersLRClass = sClass
        sClass = Replace(sClass, "ECFarmers.", "V2ECcarFarmers.", , 1, vbTextCompare)
    ElseIf InStr(1, sClass, "ECCarFarmers.", vbTextCompare) > 0 And InStr(1, sClass, "V2ECCarFarmers.", vbTextCompare) = 0 Then
        sFarmersLRClass = sClass
        sClass = Replace(sClass, "ECCarFarmers.", "V2ECcarFarmers.", , 1, vbTextCompare)
    Else
        sFarmersLRClass = vbNullString
    End If
    
    'If we get an Error then go to Unknown
    On Error Resume Next
    Set oLR = CreateObject(sClass)
    If Err.Number > 0 Then GoTo UNKNOWN_LOSS
    'Remove the class marker frm the prn data
    If sFarmersLRClass = vbNullString Then
        sPrnData = Replace(sPrnData, sClass & vbCrLf, vbNullString, 1)
    Else
        sPrnData = Replace(sPrnData, sFarmersLRClass & vbCrLf, vbNullString, 1)
    End If
    On Error GoTo EH
    GoTo SET_LOSSREPORT
UNKNOWN_LOSS:
    If Err.Number > 0 Then
        Err.Clear
    End If
    On Error GoTo EH
    Set oLR = New V2ECKeyBoard.clsLossUnknown
SET_LOSSREPORT:
    'Set up Loss Report Object with Data
    With oLR
        If goUtil Is Nothing Then
            Set goUtil = New V2ECKeyBoard.clsUtil
        End If
        .SetUtilObject goUtil
        Set .LRs = Me
        .PrnKey = psExportPath
        .PrnData = sPrnData
        .PopulateMe
    End With
       
    If oLR.ExportME(Me, pXportType) Then
        CreateExport = True
    End If
    
    Set oLR = Nothing
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Public Function CreateExport"
End Function

Public Function SetUtilObject(oUtilObject As Object) As Boolean
    On Error Resume Next
    Set goUtil = oUtilObject
    If Err.Number = 0 Then
        SetUtilObject = True
    Else
        SetUtilObject = False
    End If
End Function
