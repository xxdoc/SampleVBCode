VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLists"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Carrier Lists Polymorphism
'Carrier Lists InterFace Definiton
'Use Implements V2ECKeyBoard.clsCarLists in ALL clsLists For All Carrier Objects (Not in V2ECKeyBoard.clsLists, However!)

Implements V2ECKeyBoard.clsCarLists

Private mFrmCancel As frmCancel
Private moARReport As Object
Private moReportLists As V2ECKeyBoard.clsRptLists
Private moGUI As V2ECKeyBoard.clsCarGUI
Private WithEvents moLoss As V2ECKeyBoard.clsLossReports ' New Loss Notices 4.11.2002
Attribute moLoss.VB_VarHelpID = -1

Public Property Get clsCarLists_ClassName() As String
    clsCarLists_ClassName = App.EXEName & ".clsLists"
End Property

Public Property Get clsCarLists_goForms() As Object
    Set clsCarLists_goForms = moGUI.goForms
End Property

Public Property Let clsCarLists_LossReports(poLoss As V2ECKeyBoard.clsLossReports)
    Set moLoss = poLoss
End Property
Public Property Set clsCarLists_LossReports(poLoss As V2ECKeyBoard.clsLossReports)
    Set moLoss = poLoss
End Property
Public Property Get clsCarLists_LossReports() As V2ECKeyBoard.clsLossReports
    Set clsCarLists_LossReports = moLoss
End Property

Public Property Let clsCarLists_ARReport(poAR As Object)
    Set moARReport = poAR
End Property
Public Property Set clsCarLists_ARReport(poAR As Object)
    Set moARReport = poAR
End Property
Public Property Get clsCarLists_ARReport() As Object
    Set clsCarLists_ARReport = moARReport
End Property

Public Property Get clsCarLists_ECDBVersion() As String
    'Set the Carrier DB version here
        '12.14.2004 This is not in use at this time
        'All Carriers will use ECMain.mdb NOT a specific Carrier DB
    clsCarLists_ECDBVersion = 1
End Property

Public Property Get clsCarLists_FormatList() As Collection
    Set clsCarLists_FormatList = New Collection
    'Lists all the available Report format classes
    'available in this Carrrier object.
    With clsCarLists_FormatList
        .Add "clsLossASN", "clsLossASN"
        .Add "clsLossCCMS", "clsLossCCMS"
        .Add "clsLossXML01", "clsLossXML01"
    End With
    
End Property

Public Property Get clsCarLists_ActionList() As Collection
    On Error GoTo EH
    'Dim oULTWIA As New clsULTWIA
    'Dim oULTWIART As New clsULTWIART
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    'Lists all the available "ACTION" Classes
    'Action= classes that
    
    Set clsCarLists_ActionList = New Collection
    
    'With clsCarLists_ActionList
     '   .Add oULTWIA, oULTWIA.clsCarUL_ClassName
      '  .Add oULTWIART, oULTWIART.clsCarUL_ClassName
    'End With
    Exit Property
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLists_ClassName & vbCrLf & "Public Property Get clsCarLists_ActionList"
End Property

Public Function clsCarLists_GetARMiscDelimParamsCol(psReportName As String, _
                                                        Optional plVersion As Long = -1, _
                                                        Optional psDelimParams As String = vbNullString) As Collection
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    'Check the Util object
    If goUtil Is Nothing Then
        Set goUtil = New V2ECKeyBoard.clsUtil
    End If
    'Set ref to itself
    goUtil.SetUtilObject goUtil
    
    Select Case UCase(psReportName)
        Case Else
            'Need to Create the Object Late Binding
            'Need to Create the Object Late Binding
            If Not moReportLists Is Nothing Then
                moReportLists.CLEANUP
                Set moReportLists = Nothing
            End If
            Set moReportLists = CreateObject(psReportName)
            moReportLists.SetUtilObject goUtil
            Set clsCarLists_GetARMiscDelimParamsCol = moReportLists.GetARMiscDelimParamsCol(plVersion, psDelimParams)
            moReportLists.CLEANUP
            Set moReportLists = Nothing
    End Select
    
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Set clsCarLists_GetARMiscDelimParamsCol = Nothing
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLists_ClassName & vbCrLf & "Public Function GetARMiscDelimParamsCol"
End Function

Public Function clsCarLists_GetARMiscDelimParamsStr(psReportName As String, _
                                                        Optional plVersion As Long = -1, _
                                                        Optional psDelimParams As String = vbNullString) As String
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim oCol As Collection
    Dim vParams As Variant
    Dim MyParams As rptMiscDelimParam
    Dim sBuildDelim As String
    
    'Check the Util object
    If goUtil Is Nothing Then
        Set goUtil = New V2ECKeyBoard.clsUtil
    End If
    'Set ref to itself
    goUtil.SetUtilObject goUtil
    
    Select Case UCase(psReportName)
        Case Else
            'Need to Create the Object Late Binding
            Set moReportLists = CreateObject(psReportName)
            moReportLists.SetUtilObject goUtil
            Set oCol = moReportLists.GetARMiscDelimParamsCol(plVersion, psDelimParams)
    End Select
    
    'After Getting the Collection convert it to a Delim String
    If Not oCol Is Nothing Then
        sBuildDelim = moReportLists.ClassName & "_DELIMPARAMS" & COLUMN_DELIM
        For Each vParams In oCol
            MyParams = vParams
            With MyParams
                sBuildDelim = sBuildDelim & "ProjectName=" & .ProjectName & "|"
                sBuildDelim = sBuildDelim & "ClassName=" & .ClassName & "|"
                sBuildDelim = sBuildDelim & "ParamName=" & .ParamName & "|"
                sBuildDelim = sBuildDelim & "ParamCaption=" & .ParamCaption & "|"
                sBuildDelim = sBuildDelim & "ParamDataType=" & .ParamDataType & "|"
                sBuildDelim = sBuildDelim & "ParamValue=" & .ParamValue & "|"
                sBuildDelim = sBuildDelim & "^"
            End With
        Next
        sBuildDelim = sBuildDelim & RECORD_DELIM
    End If
    
    clsCarLists_GetARMiscDelimParamsStr = sBuildDelim
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLists_ClassName & vbCrLf & "Public Function GetARMiscDelimParamsStr"
End Function

Public Function clsCarLists_GetARReport(psReportName As String, _
                                        Optional plVersion As Long = -1, _
                                        Optional psDelimParams As String = vbNullString) As Object
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    'Check the Util object
    If goUtil Is Nothing Then
        Set goUtil = New V2ECKeyBoard.clsUtil
    End If
    'Set ref to itself
    goUtil.SetUtilObject goUtil
    
    Select Case UCase(psReportName)
        Case Else
            'Need to Create the Object Late Binding
            If Not moReportLists Is Nothing Then
                moReportLists.CLEANUP
                Set moReportLists = Nothing
            End If
            Set moReportLists = CreateObject(psReportName)
            moReportLists.SetUtilObject goUtil
            moReportLists.GetARReport plVersion, psDelimParams
    End Select
    
    'Set ref to this Lists Instance
    If Not moReportLists.ARReportObject Is Nothing Then
        Set clsCarLists_GetARReport = moReportLists.ARReportObject
        Set clsCarLists_GetARReport.Lists = Me
        Set clsCarLists_ARReport = clsCarLists_GetARReport
    End If
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLists_ClassName & vbCrLf & "Public Function GetARReport"
End Function

Public Sub clsCarLists_SetARReportProperty(psName As String, pvValue As Variant, pType As VbVarType)
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    If Not moARReport Is Nothing Then
        moARReport.SetProperty psName, pvValue, pType
    End If
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLists_ClassName & vbCrLf & "Public Sub SetARReportProperty"
End Sub

Public Function clsCarLists_CFExportARReport(psXportPath As String, pXportType As ExportType) As Boolean
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    If Not moARReport Is Nothing Then
        clsCarLists_CFExportARReport = moARReport.ExportME(psXportPath, pXportType)
    End If
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    clsCarLists_CFExportARReport = False
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLists_ClassName & vbCrLf & "Public Function CFExportARReport"
End Function

Public Function clsCarLists_ExportARReport(poARReport As Object, psXportPath As String, pXportType As ExportType) As Boolean
    On Error GoTo EH
    Dim oLRs As V2ECKeyBoard.clsLossReports
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    If Not poARReport Is Nothing Then
        Set moARReport = poARReport
    End If
    If Not moARReport Is Nothing Then
        Set oLRs = New V2ECKeyBoard.clsLossReports
    Else
        Exit Function
    End If
    With moARReport
        .Printer.Orientation = ddOPortrait
    End With
    
    If oLRs.ExportFile(moARReport, psXportPath, pXportType) Then
        clsCarLists_ExportARReport = True
    Else
        clsCarLists_ExportARReport = False
    End If
    
    'Clean up
    Unload moARReport
    Set moARReport = Nothing
    oLRs.CLEANUP
    Set oLRs = Nothing
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    clsCarLists_ExportARReport = False
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLists_ClassName & vbCrLf & "Public Function ExportME"
End Function

Public Function clsCarLists_SetGlobalObjects(pcolObjects As Collection) As Boolean
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    Set goUtil = pcolObjects("goUtil")
    Set moGUI = New ECguiTWIA.clsGUI
    moGUI.SetGlobalObjects pcolObjects
    
    clsCarLists_SetGlobalObjects = True
   
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    clsCarLists_SetGlobalObjects = False
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLists_ClassName & vbCrLf & "Public Function clsCarLists_SetGlobalObjects"
End Function


Public Sub clsCarLists_PopulateECTree(poNodX As Object, poECTree As Object, psCarCatKey As String, psCar As String, psCAT As String)
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
'    Dim sCarList As String
'    'fill in Carrier Specific Tree Items
'    sCarList = psCarCatKey & "|CarList"
'    'Add Carrier List Folder
'    Set poNodX = poECTree.Nodes.Add(psCarCatKey, 4, sCarList, psCar & " List", PicTree.A09_ClosedFolder)
'
'    'Add options to Cat Maintenance Folder
'    Set poNodX = poECTree.Nodes.Add(sCarList, 4, sCarList & "+Item01", "Item 01", PicTree.A31_Item1)
'    Set poNodX = poECTree.Nodes.Add(sCarList, 4, sCarList & "+Item02", "Item 02", PicTree.A32_Item2)
'    Set poNodX = poECTree.Nodes.Add(sCarList, 4, sCarList & "+Item03", "Item 03", PicTree.A33_Item3)
'    Set poNodX = poECTree.Nodes.Add(sCarList, 4, sCarList & "+Item04", "Item 04", PicTree.A34_Item4)
'    Set poNodX = poECTree.Nodes.Add(sCarList, 4, sCarList & "+Item05", "Item 05", PicTree.A35_Item5)
'    Set poNodX = poECTree.Nodes.Add(sCarList, 4, sCarList & "+Item06", "Item 06", PicTree.A36_Item6)
'    Set poNodX = poECTree.Nodes.Add(sCarList, 4, sCarList & "+Item07", "Item 07", PicTree.A37_Item7)
'    Set poNodX = poECTree.Nodes.Add(sCarList, 4, sCarList & "+Item08", "Item 08", PicTree.A38_Item8)
'    Set poNodX = poECTree.Nodes.Add(sCarList, 4, sCarList & "+Item09", "Item 09", PicTree.A39_Item9)
'    'Add Items to Item09
'    Set poNodX = poECTree.Nodes.Add(sCarList & "+Item09", 4, sCarList & "+Item09+Item01", "Item 1", PicTree.A31_Item1)
'    Set poNodX = poECTree.Nodes.Add(sCarList & "+Item09", 4, sCarList & "+Item09+Item02", "Item 2", PicTree.A32_Item2)
'    Set poNodX = poECTree.Nodes.Add(sCarList & "+Item09", 4, sCarList & "+Item09+Item03", "Item 3", PicTree.A33_Item3)
'    'Add Item 9 SubItems Folder
'    Set poNodX = poECTree.Nodes.Add(sCarList & "+Item09", 4, sCarList & "+Item09+SubItems", "Sub Items", PicTree.A09_ClosedFolder)
'    'Add Options to Folder 9
'    Set poNodX = poECTree.Nodes.Add(sCarList & "+Item09+SubItems", 4, sCarList & "+Item09+SubItems+Item01", "Item 01", PicTree.A31_Item1)
'    Set poNodX = poECTree.Nodes.Add(sCarList & "+Item09+SubItems", 4, sCarList & "+Item09+SubItems+Item02", "Item 02", PicTree.A32_Item2)
'    Set poNodX = poECTree.Nodes.Add(sCarList & "+Item09+SubItems", 4, sCarList & "+Item09+SubItems+Item03", "Item 03", PicTree.A33_Item3)
'    Set poNodX = poECTree.Nodes.Add(sCarList & "+Item09+SubItems", 4, sCarList & "+Item09+SubItems+Item04", "Item 04", PicTree.A34_Item4)
'    Set poNodX = poECTree.Nodes.Add(sCarList, 4, sCarList & "+Item10", "Item 10", PicTree.A40_Item10)
'
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLists_ClassName & vbCrLf & "Public Sub PopulateTree"
End Sub

Public Function clsCarLists_CarListCommand(psCar As String, psCAT As String, psCommand As String) As Boolean
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
'    Dim oGui As ECguiTWIA.clsGUI
'    Set oGui = moGUI
'
'
'
'    Set oGui = Nothing
'
    clsCarLists_CarListCommand = True
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    clsCarLists_CarListCommand = False
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLists_ClassName & vbCrLf & "Public Function CarListCommand"
End Function

Public Function clsCarLists_CatCommand(psCar As String, psCAT As String, psCommand As String) As Boolean
    On Error GoTo EH
    Dim oForm As Form
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    'Execute Command
    If StrComp(psCommand, "ClaimsListView", vbTextCompare) = 0 Then
        moGUI.ShowForm "frmClaimsList", vbModeless
    ElseIf StrComp(psCommand, "AddClaim", vbTextCompare) = 0 Then
        
    ElseIf StrComp(psCommand, "FeeSchedule", vbTextCompare) = 0 Then
        moGUI.ShowForm "frmFeeSchedule", vbModeless
    ElseIf StrComp(psCommand, "CatPreferences", vbTextCompare) = 0 Then
    
    ElseIf StrComp(psCommand, "Communications", vbTextCompare) = 0 Then
        goUtil.utShowFTP goUtil.gsInstallDir & "\ECFTP.EXE " & goUtil.gsCommandString
    ElseIf StrComp(psCommand, "ViewLossReports", vbTextCompare) = 0 Then
        
        If goUtil.utFindSetForm(goUtil.goForms, "frmLossReports", oForm) Then
            oForm.Show
            oForm.WindowState = vbNormal
        Else
            clsCarLists_LoadLoss False
        End If
    ElseIf StrComp(psCommand, "ViewAdjusterReports", vbTextCompare) = 0 Then
         moGUI.ShowForm "frmAdjusterReports", vbModeless
    ElseIf StrComp(psCommand, "MakeCurrent", vbTextCompare) = 0 Then
       
    End If
    Set oForm = Nothing
    clsCarLists_CatCommand = True
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    clsCarLists_CatCommand = False
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLists_ClassName & vbCrLf & "Public Function CatCommand"
End Function

Public Function clsCarLists_SendToXactimate(Optional psCLIENTNUM As String = vbNullString) As Boolean
    On Error GoTo EH
    Dim XProj As V2ECKeyBoard.udtXactProject
    Dim vXProj As Variant
    Dim oConn As New ADODB.Connection
    Dim adoRS As ADODB.Recordset
    Dim sSQL As String
    Dim colXProjects As Collection
    Dim sCity As String
    Dim sZip As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim sTemp As String
    
    sSQL = "SELECT A.*, "
    sSQL = sSQL & "S.Status, "
    sSQL = sSQL & "CCCS.CatCode "
    sSQL = sSQL & "FROM (Assignments A "
    sSQL = sSQL & "INNER JOIN STATUS S ON A.StatusID = S.StatusID) "
    sSQL = sSQL & "INNER JOIN CLIENTCOMPANYCATSPEC CCCS ON (A.ClientCompanyCatSpecID = CCCS.ClientCompanyCatSpecID) "
    'If they are in a specific claim then only get this one claim to be sent to
    'xactimate. Otherwise we are sending all projects to xactimate.
    If psCLIENTNUM <> vbNullString Then
        sSQL = sSQL & "WHERE A.CLIENTNUM = '" & goUtil.utCleanSQLString(psCLIENTNUM) & "' "
    Else
        sSQL = sSQL & "WHERE A.ClientCompanyCatSpecID IN "
        sSQL = sSQL & "( "
        sSQL = sSQL & "SELECT  ClientCompanyCatSpecID "
        sSQL = sSQL & "FROM ClientCompanyCatSpec "
        sSQL = sSQL & "WHERE ClientCompanyID = " & goUtil.gsCurCar & " "
        sSQL = sSQL & "AND     CATID = " & goUtil.gsCurCat & " "
        sSQL = sSQL & ") "
    End If
    sSQL = sSQL & "AND A.AdjusterSpecID IN "
    sSQL = sSQL & "( "
    sSQL = sSQL & "SELECT  ClientCoAdjusterSpecID "
    sSQL = sSQL & "FROM ClientCoAdjusterSpec "
    sSQL = sSQL & "Where ClientCompanyID = " & goUtil.gsCurCar & " "
    sSQL = sSQL & "AND UsersID = " & goUtil.gsCurUsersID & " "
    sSQL = sSQL & ") "
    sSQL = sSQL & "AND A.Reassigned = False "
    sSQL = sSQL & "AND A.IsDeleted = False "
    sSQL = sSQL & "AND A.IsLocked = False "
    sSQL = sSQL & "ORDER BY S.Status, A.CLIENTNUM "
    
    
    Set oConn = New ADODB.Connection
    Set adoRS = New ADODB.Recordset
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    'Use Disconnected Record Set on asUseClient Cusor ONLY !
    adoRS.CursorLocation = adUseClient
    adoRS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set adoRS.ActiveConnection = Nothing

    If Not adoRS.EOF Then
        adoRS.MoveFirst
    Else
        clsCarLists_SendToXactimate = True 'No errors
        MsgBox "No Available Records.", vbInformation + vbOKOnly, "Send To Xactimate"
        GoTo SKIPXACT
    End If
    
    Set goUtil.goXact = New V2ECKeyBoard.clsXact
    goUtil.goXact.StartHwnd = goUtil.gfrmECTray.hWnd
    
    If Not goUtil.goXact.LookUpLoaded Then
        goUtil.goXact.CLEANUP
        Set goUtil.goXact = Nothing
        clsCarLists_SendToXactimate = True
        GoTo SKIPXACT
    End If
    
    'BGS 3.1.2002 OK it might help if there is some more comments here now that I have 2 seconds.
    'First we are loading Xproj user defined type with all the Admin information that will
    'be typed into xactimate.
    
    Do Until adoRS.EOF
        With XProj
            .SentToXact = IIf(IsNull(adoRS!SentToXact), False, adoRS!SentToXact)
            With .Main
                .Name = IIf(IsNull(adoRS!Insured), vbNullString, adoRS!Insured)
                .Street = IIf(IsNull(adoRS!PAStreet), vbNullString, adoRS!PAStreet)
                .City = IIf(IsNull(adoRS!PACity), vbNullString, adoRS!PACity)
                .State = IIf(IsNull(adoRS!PAState), vbNullString, adoRS!PAState)
                .Zip = IIf(IsNull(adoRS!PAZIP), vbNullString, Format(adoRS!PAZIP, "00000"))
                .Zip4 = IIf(IsNull(adoRS!PAZIP4), vbNullString, Format(adoRS!PAZIP4, "0000"))
                .DateEntered = Format(Now, "MM/DD/YYYY")
                .DateInspected = IIf(IsNull(adoRS!InspectedDate), vbNullString, adoRS!InspectedDate)
                .DateOfLoss = IIf(IsNull(adoRS!LossDate), vbNullString, adoRS!LossDate)
                .DateReceived = IIf(IsNull(adoRS!ReceivedDate), vbNullString, adoRS!ReceivedDate)
            End With
            With .Loss
                .ClaimNumber = IIf(IsNull(adoRS!CLIENTNUM), vbNullString, adoRS!CLIENTNUM)
                .PolicyNumber = IIf(IsNull(adoRS!PolicyNo), vbNullString, adoRS!PolicyNo)
                .Deductible = IIf(IsNull(adoRS!Deductible), vbNullString, adoRS!Deductible)
                .TypeOfLoss = IIf(IsNull(adoRS!XactTypeOfLoss), vbNullString, adoRS!XactTypeOfLoss)
                .CatCode = IIf(IsNull(adoRS!CatCode), vbNullString, adoRS!CatCode)
            End With
        End With
        'Add each UDT into the Collection inside the Xact class
        'This UDT will convert into a variant which is important
        'because you can only use Variant to do For Each loop on a collection
        'of UDTs.  Kind of funcky I know but it works just fine.
        If Not goUtil Is Nothing Then
            If Not goUtil.goXact Is Nothing Then
                goUtil.goXact.AddXProject XProj
            Else
                GoTo SKIPXACT
            End If
        Else
            
        End If
        adoRS.MoveNext
    Loop
    
    'BGS 3.1.2002 We are done with the RS here so we can
    'clear it so if say someone clicks on make cat current while
    'we are in the middle of typing info into xactimate we won't get
    'error
    Set adoRS = Nothing
    If Not goUtil Is Nothing Then
        If goUtil.goXact Is Nothing Then
            GoTo SKIPXACT
        End If
    Else
        GoTo SKIPXACT
    End If
    
    'BGS we still have not launched Xactimate yet,
    'after ValidateXactprojects Passes then we update
    'Easy Claim DB then check to see if we need to launch xactimate
    
    If goUtil.goXact.ValidateXactProjects() Then
        'BGS 2/1/2002 need to check for any dirty projects
        'means they were changed due to the validation process
        'and update Assignemtnes table with new data.
        If Not goUtil.goXact.GetFromExport Then
            Set colXProjects = goUtil.goXact.XactProjects
            For Each vXProj In colXProjects
                XProj = vXProj
                If XProj.IsDirty Then
                    With XProj
                        sSQL = "UPDATE Assignments "
                        sSQL = sSQL & "SET "
                        sSQL = sSQL & "Assignments.Insured = '" & goUtil.utCleanSQLString(.Main.Name) & "', "
                        sSQL = sSQL & "Assignments.PAStreet = '" & goUtil.utCleanSQLString(.Main.Street) & "', "
                        sCity = Replace(goUtil.utCleanSQLString(.Main.City), Chr(32), Chr(160))
                        sSQL = sSQL & "Assignments.PACity = '" & sCity & "', "
                        sSQL = sSQL & "Assignments.PAState = '" & goUtil.utCleanSQLString(.Main.State) & "', "
                        sSQL = sSQL & "Assignments.PAZip = " & .Main.Zip & ", "
                        sSQL = sSQL & "Assignments.PAZip4 = " & .Main.Zip4 & ", "
                        sSQL = sSQL & "Assignments.PropertyAddress = '" & goUtil.utCleanSQLString(goUtil.utJoinToAddress(.Main.Street, sCity, .Main.State, sZip)) & "', "
                        sTemp = goUtil.utValidDate(.Main.DateInspected)
                        sSQL = sSQL & "Assignments.InspectedDate = " & IIf(sTemp = NULL_DATE, "null", "#" & sTemp & "#") & ", "
                        sTemp = goUtil.utValidDate(.Main.DateOfLoss)
                        sSQL = sSQL & "Assignments.LossDate =  " & IIf(sTemp = NULL_DATE, "null", "#" & sTemp & "#") & ", "
                        sTemp = goUtil.utValidDate(.Main.DateReceived)
                        sSQL = sSQL & "Assignments.ReceivedDate = " & IIf(sTemp = NULL_DATE, "null", "#" & sTemp & "#") & ", "
                        sSQL = sSQL & "Assignments.CLIENTNUM = '" & goUtil.utCleanSQLString(.Loss.ClaimNumber) & "', "
                        sSQL = sSQL & "Assignments.PolicyNo = '" & goUtil.utCleanSQLString(.Loss.PolicyNumber) & "', "
                        sSQL = sSQL & "Assignments.Deductible = '" & goUtil.utCleanSQLString(.Loss.Deductible) & "', "
                        sSQL = sSQL & "Assignments.XactTypeOfLoss = '" & goUtil.utCleanSQLString(.Loss.TypeOfLoss) & "', "
                        sSQL = sSQL & "Assignments.UpLoadMe = True "
                        sSQL = sSQL & "WHERE "
                        sSQL = sSQL & "Assignments.CLIENTNUM = '" & goUtil.utCleanSQLString(.Loss.ClaimNumber) & "' "
                        oConn.Execute sSQL
                    End With
                End If
            Next
        End If
        
        'Display the Cancel Form
        Set mFrmCancel = New frmCancel
        Set mFrmCancel.Util = goUtil
        Load mFrmCancel
        mFrmCancel.Visible = True
        mFrmCancel.lblMess.Caption = "Please wait! Sending files..."
        
        If Not goUtil Is Nothing Then
            If goUtil.goXact Is Nothing Then
               GoTo SKIPXACT
            End If
        Else
            GoTo SKIPXACT
        End If
        
        'If we are not skiping all then we need to SendToXactimate
        'SendToXact will launch Xactimate if it isn't already loaded and set focus to it
        If Not goUtil.goXact.SkipAll Then
            If goUtil.goXact.SendToXact() Then
                'OK we need a check here incase user tries to unload in middle of sending
                If Not goUtil Is Nothing Then
                    If goUtil.goXact Is Nothing Then
                       GoTo SKIPXACT
                    End If
                Else
                    GoTo SKIPXACT
                End If
                
                goUtil.goXact.LookForWindow left(goUtil.gfrmECTray.Caption, 10)
                
                If goUtil.goXact.SendToExport Then
                    MsgBox "Project(s) Exported to: " & vbCrLf & vbCrLf & goUtil.goXact.ExportFilePath, vbInformation + vbOKOnly, "Send To Xactimate Export File"
                Else
                    MsgBox "Project(s) Sent to Xactimate!", vbInformation + vbOKOnly, "Send To Xactimate"
                End If
                
            Else
                'OK we need a check here incase user tries to unload in middle of sending
                If Not goUtil Is Nothing Then
                    If goUtil.goXact Is Nothing Then
                        GoTo SKIPXACT
                    End If
                Else
                    GoTo SKIPXACT
                End If
                
                goUtil.goXact.LookForWindow left(goUtil.gfrmECTray.Caption, 10)
                If goUtil.goXact.SendToExport Then
                    MsgBox "Project(s) Failed to Export to: " & vbCrLf & vbCrLf & goUtil.goXact.ExportFilePath, vbExclamation + vbOKOnly, "Send To Xactimate Export File"
                Else
                    MsgBox "Project(s) Not Sent to Xactimate!", vbExclamation + vbOKOnly, "Send To Xactimate"
                End If
            End If
        End If
    End If
    clsCarLists_SendToXactimate = True
SKIPXACT:
    If Not mFrmCancel Is Nothing Then
        Unload mFrmCancel
        Set mFrmCancel = Nothing
    End If
    Set colXProjects = Nothing
    Set adoRS = Nothing
    Set oConn = Nothing
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    If Not mFrmCancel Is Nothing Then
        Unload mFrmCancel
        Set mFrmCancel = Nothing
    End If
    clsCarLists_SendToXactimate = False
    If Not goUtil.goXact Is Nothing Then
        goUtil.goXact.CLEANUP
    End If
    Set goUtil.goXact = Nothing
    Set colXProjects = Nothing
    Set adoRS = Nothing
    Set oConn = Nothing
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLists_ClassName & vbCrLf & "Public Function SendToXactimate"
End Function

Public Function clsCarLists_FlagSentToXactimate(pvXProj As Variant) As Boolean
    On Error GoTo EH
    Dim XProj As V2ECKeyBoard.udtXactProject
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    XProj = pvXProj
    
    With XProj
        sSQL = "UPDATE Assignments "
        sSQL = sSQL & "SET "
        'Yes Yes don't blame me for using InspectedWindowScreens as Send to Xact Flag
        'Only update this flag if we are successfull sending it to Xactimate
        sSQL = sSQL & "Assignments.SentToXact = " & .SentToXact & "  "
        sSQL = sSQL & "WHERE "
        sSQL = sSQL & "Assignments.CLIENTNUM = '" & goUtil.utCleanSQLString(.Loss.ClaimNumber) & "' "
        If Not goUtil Is Nothing Then
            If Not goUtil.gMainDB Is Nothing Then
                Set oConn = New ADODB.Connection
                goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
                oConn.Execute sSQL
            Else
                Exit Function
            End If
        End If
    End With
    
    clsCarLists_FlagSentToXactimate = True
    Set oConn = Nothing
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    clsCarLists_FlagSentToXactimate = False
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLists_ClassName & vbCrLf & "Public Function FlagSentToXactimate"
End Function

'<---------------------------Implement New Loss Notice Here-------------------->
Public Sub clsCarLists_LoadLoss(pbUpdateDB As Boolean)
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
'    If Not moLoss Is Nothing Then
'        Exit Sub
'    End If
    
    With goUtil.goProgForm
        .LoadForm
        .Caption = "Loss Report Progress"
        .framTableText = vbNullString
        .framRecordText = vbNullString
        .framFileText = "Loss Reports"
        .cmdCancelEnable = True
        .ShowForm
    End With
   
    Set moLoss = New V2ECKeyBoard.clsLossReports
    'Pass in the Utility object
    moLoss.SetUtilObject goUtil
    'Add the Append Documents
    clsCarLists_ADDAppDocs
    'If updating Database then Add Parameters
    If pbUpdateDB Then
        clsCarLists_ADDParameters
    End If
   
    'Show Loss Reports
    moLoss.ShowLossReports , goUtil.goProgForm
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Set moLoss = Nothing
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLists_ClassName & vbCrLf & "Public Sub LoadLoss"
End Sub

Public Sub clsCarLists_ADDAppDocs()
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    '---------------->NOTE<-----------------------------
    'Maximum of 14 App docs may be added to the collection
    '---------------->NOTE<-----------------------------
    If Not goUtil.goCurCarList Is Nothing Then
        Set moLoss.AppDocsCol = clsCarLists_GetAppDocsCol
    End If

    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLists_ClassName & vbCrLf & "Public Sub ADDAppDocs"
End Sub

Public Sub clsCarLists_ADDParameters()
    On Error GoTo EH
    Dim Param As V2ECKeyBoard.udtParameter
    Dim sFName As String 'First Name Adjuster
    Dim sLName As String 'Last Name Adjuster
    Dim dTaxPercent As Double
    Dim sTypeOfLoss As String
    Dim sSiteAddress As String
    Dim sSiteState As String
    Dim sSiteCity As String
    Dim sSiteZip As String
    Dim lMaxID As Long
    Dim sCRID As String
    Dim sCATCode As String
    Dim sSSN As String
    Dim sEmail As String
    Dim sContactPhone As String
    Dim sAppVSInfo As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    'Add Cat Code
    sCATCode = GetSetting(App.EXEName, "GENERAL", "CAT_CODE", 0#)
    Param.ParamName = "CAT_CODE"
    Param.ParamValue = sCATCode
    moLoss.AddParameter Param

    'Add "TAX_PERCENT"
    dTaxPercent = GetSetting(App.EXEName, "FEES", "TAX_PERCENT", 0#)
    Param.ParamName = "TAX_PERCENT"
    Param.ParamValue = dTaxPercent
    moLoss.AddParameter Param

    'Add "IB_PREFIX"
    'First lettter of First and first letter of Last
    sFName = UCase(GetSetting(App.EXEName, "GENERAL", "ADJUSTOR_FIRST_NAME", vbNullString))
    sLName = UCase(GetSetting(App.EXEName, "GENERAL", "ADJUSTOR_LAST_NAME", vbNullString))
    Param.ParamName = "IB_PREFIX"
    Param.ParamValue = left(sFName, 1) & left(sLName, 1)
    moLoss.AddParameter Param

    'Add First Name
    Param.ParamName = "ADJ_FIRST_NAME"
    Param.ParamValue = sFName
    moLoss.AddParameter Param

    'Add Last Name
    Param.ParamName = "ADJ_LAST_NAME"
    Param.ParamValue = sLName
    moLoss.AddParameter Param

    'Add CRID
    sCRID = GetSetting(App.EXEName, "GENERAL", "CLAIM_REP_ID_NO", vbNullString)
    Param.ParamName = "CRID"
    Param.ParamValue = sCRID
    moLoss.AddParameter Param

    'Add SSN
    sSSN = GetSetting("ECS", "WEB_SECURITY", "USERNAME", vbNullString)
    Param.ParamName = "ADJ_SSN"
    Param.ParamValue = sSSN
    moLoss.AddParameter Param

    'Add Email
    sEmail = GetSetting(App.EXEName, "GENERAL", "ADJ_EMAIL", vbNullString)
    Param.ParamName = "ADJ_EMAIL"
    Param.ParamValue = sEmail
    moLoss.AddParameter Param

    'Add Contact Phone
    sContactPhone = GetSetting(App.EXEName, "GENERAL", "ADJ_CONTACT_PHONE", vbNullString)
    Param.ParamName = "ADJ_CONTACT_PHONE"
    Param.ParamValue = sContactPhone
    moLoss.AddParameter Param

    'Add AppVSInfo
    sAppVSInfo = goUtil.utGetAppVSInfo(App.EXEName, App.Path)
    Param.ParamName = "ADJ_AppVSInfo"
    Param.ParamValue = sAppVSInfo
    moLoss.AddParameter Param

    'Add "TYPE_OF_LOSS"
    sTypeOfLoss = GetSetting(App.EXEName, "GENERAL", "TYPE_OF_LOSS", vbNullString)
    Param.ParamName = "TYPE_OF_LOSS"
    Param.ParamValue = sTypeOfLoss
    moLoss.AddParameter Param

    'Add SiteAddress
    sSiteAddress = GetSetting(App.EXEName, "GENERAL", "ADDRESS_1", vbNullString)
    Param.ParamName = "SITE_ADDRESS"
    Param.ParamValue = sSiteAddress
    moLoss.AddParameter Param

    'Add SiteState
    sSiteState = GetSetting(App.EXEName, "GENERAL", "STATE", vbNullString)
    Param.ParamName = "SITE_STATE"
    Param.ParamValue = sSiteState
    moLoss.AddParameter Param

    'Add SiteCity
    sSiteCity = GetSetting(App.EXEName, "GENERAL", "CITY", vbNullString)
    Param.ParamName = "SITE_CITY"
    Param.ParamValue = sSiteCity
    moLoss.AddParameter Param

    'Add SiteZip
    sSiteZip = GetSetting(App.EXEName, "GENERAL", "ZIP", vbNullString)
    Param.ParamName = "SITE_ZIP"
    Param.ParamValue = sSiteZip
    moLoss.AddParameter Param

    'Need to Get Max ID from Assignemnts Table
    'Because of bad AutoNum
    goUtil.utBadAutoNum , , lMaxID
    Param.ParamName = "MAX_ID"
    Param.ParamValue = lMaxID
    moLoss.AddParameter Param

    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLists_ClassName & vbCrLf & "Public Sub ADDParameters"
End Sub

Public Function clsCarLists_GetAppDocsCol() As Collection
    On Error GoTo EH
    Dim sDocName As String
    Dim MyAppDoc As V2ECKeyBoard.udtAppDoc
    Dim lcount As Long
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    Set clsCarLists_GetAppDocsCol = New Collection
    
    '---------------->NOTE<-----------------------------
    'Maximum of 14 App docs may be added to the collection
    '---------------->NOTE<-----------------------------
    'Put the Appended documents in the order they should be printed
    'after each Loss report is printed
    
    For lcount = 1 To 14
    'lcount = 3, "TWIA Claims Explainer",
        sDocName = Switch(lcount = 1, "WorkSheet Diagram", _
                          lcount = 2, "Activity Log (Blank)", _
                          lcount = 3, vbNullString, _
                          lcount = 4, vbNullString, _
                          lcount = 5, vbNullString, _
                          lcount = 6, vbNullString, _
                          lcount = 7, vbNullString, _
                          lcount = 8, vbNullString, _
                          lcount = 9, vbNullString, _
                          lcount = 10, vbNullString, _
                          lcount = 11, vbNullString, _
                          lcount = 12, vbNullString, _
                          lcount = 13, vbNullString, _
                          lcount = 14, vbNullString)
        If sDocName <> vbNullString Then
            'All Appdocs will start off not selected to print
            MyAppDoc.Selected = False
            MyAppDoc.DocName = sDocName
            clsCarLists_GetAppDocsCol.Add MyAppDoc, sDocName
        End If
    Next
        
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLists_ClassName & vbCrLf & "Public Function GetAppDocsCol"
End Function

Public Function clsCarLists_AddLogEntries(psClaimNo As String, pdtActDate As Date, _
                              Optional pbAddPrefix As Boolean, _
                              Optional psPolicyDesc As String, _
                              Optional psBuildLim As String, _
                              Optional psContLim As String, _
                              Optional psDeductible As String) As Boolean
                              
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
'    Dim sLName As String
'    Dim sFName As String
'    Dim sPrefix As String
'    Dim sSQL As String
'    Dim RS As Recordset
'    Dim bInsertLog As Boolean
'    Dim sIBNum As String
'    Dim lMaxID As Long
'    Dim sAct As String 'Activitiy text
'
'    clsCarLists_AddLogEntries = True
'
'    sLName = UCase(get_INIString("GENERAL", "ADJUSTOR_LAST_NAME", 0))
'    sFName = UCase(get_INIString("GENERAL", "ADJUSTOR_FIRST_NAME", 0))
'
'    sPrefix = left(sFName, 1) & left(sLName, 1)
'
'    If pbAddPrefix Then
'        sIBNum = sPrefix & psClaimNo
'    Else
'        sIBNum = psClaimNo
'    End If
'
'    '6.27.2002 before we insert a log entry and init the service fee types(Fee schedule or time entry),
'    'we need to be sure that the claimno exists in the Assignments Table.  If we don't check, then we
'    'will end up putting a log entry for a claim that already exists.  It won't hurt
'    'anything except take up space needlessly, so lets make sure we don't take up space.
'    sSQL = "SELECT A.ClaimNo "
'    sSQL = sSQL & "FROM " & DB_ASSIGNMENTS & " As A "
'    sSQL = sSQL & "WHERE A.ClaimNo = '" & CleanString(sIBNum) & "' "
'
'    Set RS = gCurDB.OpenRecordset(sSQL)
'
'    If Not RS.EOF Then
'        RS.MoveFirst
'        If sIBNum = IIf(IsNull(RS!ClaimNo), vbNullString, RS!ClaimNo) Then
'            bInsertLog = True
'        End If
'    End If
'    Set RS = Nothing
'    If bInsertLog Then
'        'BGS 3.1.2002 151  "Duplicate IB"
'        BadAutoNum , DB_ACTIVITY, lMaxID
'        sSQL = "INSERT INTO " & DB_ACTIVITY & " (ID, "
'        sSQL = sSQL & "ClaimNumber, "
'        sSQL = sSQL & "ActivityDate, "
'        sSQL = sSQL & "ActivityTime, "
'        sSQL = sSQL & "ActivityText, "
'        sSQL = sSQL & "RTUpLoadMe ) "
'
'        sSQL = sSQL & "Values (" & lMaxID + 1 & ", "
'        sSQL = sSQL & "'" & CleanString(sIBNum) & "', "
'        sSQL = sSQL & "'" & Format(pdtActDate, "YYYYMMDD") & "', "
'        sSQL = sSQL & "'" & Format(Now, "HH:MM") & "', "
'        'BGS 4.15.2002 Add in Policy descritiona and Building Cont limits
'        'To the first line
'        sAct = "Claim received.  "
'        If psPolicyDesc <> vbNullString Then
'            sAct = sAct & "Policy Description: " & psPolicyDesc & ".  "
'        End If
'        If psBuildLim <> vbNullString Then
'            sAct = sAct & "Bldg Limit: " & psBuildLim & ".  "
'        End If
'        If psContLim <> vbNullString Then
'            sAct = sAct & "Cont Limit: " & psContLim & ".  "
'        End If
'        If psDeductible <> vbNullString Then
'            sAct = sAct & "Deductible: " & psDeductible & ".  "
'        End If
'        sSQL = sSQL & "'" & sAct & "', "
'        sSQL = sSQL & "True ) "
'
'        gCurDB.Execute sSQL
'        If gCurDB.RecordsAffected > 0 Then
'            '7.3.2002 Need to Get rid of any RTDelete Actions if they are the
'            'Same as the ID we are adding for this table.  This Will avoid the
'            'possibility that the user will delete the last record then add another
'            'record with the same Id that is marked to be deleted on the Server.
'            UpdateRTDelete lMaxID + 1, "ActLog", sIBNum, True
'        End If
'
'        'BGS after adding a log entry we can also init Fee Schedule items
'        'this is a good place to do this call since entering the initial
'        'log entry is the same time we need to set up initial fee settings
'
'        InitFeeSchedValues sIBNum
'    End If

    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    clsCarLists_AddLogEntries = False
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLists_ClassName & vbCrLf & "Public Function AddLogEntries"
End Function

Private Function clsCarLists_SetUtilObject(oUtilObject As Object) As Boolean

    On Error Resume Next
    Set goUtil = oUtilObject
    If Err.Number = 0 Then
        clsCarLists_SetUtilObject = True
    Else
        clsCarLists_SetUtilObject = False
    End If

End Function

Public Sub clsCarLists_ShowAllForms(oECTray As Object, psNavPos As String, Optional psSkipForms As String)
    moGUI.ShowAllForms oECTray, psNavPos, psSkipForms
End Sub

Public Sub clsCarLists_HideAllForms(Optional psSkipForms As String)
    moGUI.HideAllForms psSkipForms
End Sub

Public Function clsCarLists_CLEANUP() As Boolean
    On Error GoTo EH
    Dim oForms As Object
    Dim oForm As Form
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    'Chekc for any Util Forms opend by this Carrier
    Set oForms = goUtil.goForms
    
    For Each oForm In oForms
        Debug.Print oForm.Name
        Unload oForm
        Set oForm = Nothing
    Next
    Set oForms = Nothing
     
    If Not moGUI Is Nothing Then
        moGUI.CLEANUP
        Set moGUI = Nothing
    End If
    
    If Not moLoss Is Nothing Then
        moLoss.CLEANUP
        Set moLoss = Nothing
    End If
    
    'Clean Up Report Objects
    If Not moARReport Is Nothing Then
        Unload moARReport
        Set moARReport = Nothing
    End If
    If Not moReportLists Is Nothing Then
        moReportLists.CLEANUP
        Set moReportLists = Nothing
    End If
    
    clsCarLists_CLEANUP = True
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    clsCarLists_CLEANUP = False
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLists_ClassName & vbCrLf & "Public Function CLEANUP"
End Function

Private Sub moLoss_CleanUpLossReports()
    On Error GoTo EH
    If Not moLoss Is Nothing Then
        moLoss.CLEANUP
        Set moLoss = Nothing
    End If
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, clsCarLists_ClassName, "Private Sub moLoss_CleanUpLossReports", False
End Sub

Private Sub moLoss_ErrorMess(ByVal Mess As String)
    On Error GoTo EH
    Dim vMess As Variant
    Dim lcount As Long
    vMess = Split(Mess, vbCrLf)
    For lcount = LBound(vMess) To UBound(vMess)
        If goUtil.utFormExists(Forms, "frmCommStatus") Then
            goUtil.gfrmCommStatus.lstProcess.AddItem vMess(lcount)
        End If
    Next lcount
    
    Err.Raise -999, , Mess
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, clsCarLists_ClassName, "Private Sub moLoss_ErrorMess", False
End Sub

Private Sub moLoss_PopulateLossReportsRS(RSLossReports As ADODB.Recordset)
    On Error GoTo EH
    Dim oConn As New ADODB.Connection
    Dim sSQL As String
   
    Set oConn = New ADODB.Connection
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    'Need to Get Companies, not client Companies
    'that the DB User Name has access to...
    sSQL = "SELECT A.*, "
    sSQL = sSQL & "( "
    sSQL = sSQL & "SELECT UserName "
    sSQL = sSQL & "FROM USERS "
    sSQL = sSQL & "WHERE UsersID =  " & goUtil.gsCurUsersID & " "
    sSQL = sSQL & ") As AdjusterSpecUserName, "
    sSQL = sSQL & "( "
    sSQL = sSQL & "SELECT  ACID "
    sSQL = sSQL & "FROM ClientCoAdjusterSpec "
    sSQL = sSQL & "Where ClientCoAdjusterSpecID = A.[AdjusterSPecID] "
    sSQL = sSQL & ") As AdjusterSpecACID, "
    sSQL = sSQL & "( "
    sSQL = sSQL & "SELECT  ACIDDescription "
    sSQL = sSQL & "FROM ClientCoAdjusterSpec "
    sSQL = sSQL & "Where ClientCoAdjusterSpecID = A.[AdjusterSPecID] "
    sSQL = sSQL & ") As AdjusterSpecAcidDescription, "
    sSQL = sSQL & "( "
    sSQL = sSQL & "SELECT  Type "
    sSQL = sSQL & "FROM    AssignmentType "
    sSQL = sSQL & "WHERE   AssignmentTypeID = A.[AssignmentTypeID] "
    sSQL = sSQL & ") As AssignmentTypeType, "
    sSQL = sSQL & "( "
    sSQL = sSQL & "SELECT  CatCode "
    sSQL = sSQL & "FROM    ClientCompanyCatSpec "
    sSQL = sSQL & "WHERE   ClientCompanyCatSpecID = A.[ClientCompanyCatSpecID] "
    sSQL = sSQL & ") As ClientCompanyCatSpecCatCode, "
    sSQL = sSQL & "( "
    sSQL = sSQL & "SELECT Name "
    sSQL = sSQL & "FROM CAT "
    sSQL = sSQL & "WHERE CATID = " & goUtil.gsCurCat & " "
    sSQL = sSQL & ") As CatName, "
    sSQL = sSQL & "S.Status As StatusStatus, "
    sSQL = sSQL & "CCCS.CatCode "
    sSQL = sSQL & "FROM (Assignments A "
    sSQL = sSQL & "INNER JOIN STATUS S ON A.StatusID = S.StatusID) "
    sSQL = sSQL & "INNER JOIN CLIENTCOMPANYCATSPEC CCCS ON (A.ClientCompanyCatSpecID = CCCS.ClientCompanyCatSpecID) "
    'If they are in a specific claim then only get this one claim to be sent to
    'xactimate. Otherwise we are sending all projects to xactimate.
    sSQL = sSQL & "WHERE A.ClientCompanyCatSpecID IN "
    sSQL = sSQL & "( "
    sSQL = sSQL & "SELECT  ClientCompanyCatSpecID "
    sSQL = sSQL & "FROM ClientCompanyCatSpec "
    sSQL = sSQL & "WHERE ClientCompanyID = " & goUtil.gsCurCar & " "
    sSQL = sSQL & "AND     CATID = " & goUtil.gsCurCat & " "
    sSQL = sSQL & ") "
    sSQL = sSQL & "AND A.AdjusterSpecID IN "
    sSQL = sSQL & "( "
    sSQL = sSQL & "SELECT  ClientCoAdjusterSpecID "
    sSQL = sSQL & "FROM ClientCoAdjusterSpec "
    sSQL = sSQL & "Where ClientCompanyID = " & goUtil.gsCurCar & " "
    sSQL = sSQL & "AND UsersID = " & goUtil.gsCurUsersID & " "
    sSQL = sSQL & ") "
    sSQL = sSQL & "AND A.Reassigned = False "
    sSQL = sSQL & "AND A.IsDeleted = False "
'    sSQL = sSQL & "AND A.IsLocked = False "
    sSQL = sSQL & "ORDER BY S.Status, A.CLIENTNUM "
    
    'Use Disconnected Record Set on asUseClient Cusor ONLY !
    'Need this so RecordCount will populate and All fields will populate
    RSLossReports.CursorLocation = adUseClient
    RSLossReports.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set RSLossReports.ActiveConnection = Nothing
    
CLEAN_UP:
    Set oConn = Nothing
   
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, clsCarLists_ClassName, "Private Sub moLoss_PopulateLossReportsRS", True
End Sub

Private Sub moLoss_PrintAppDoc(ByVal colAppDocs As Collection, ByVal sAPPDocClientClaimNo As String)
    On Error GoTo EH
    Dim sReportName As String
    Dim MyAppDoc As V2ECKeyBoard.udtAppDoc
    Dim vAppDoc As Variant
    Dim oConn As ADODB.Connection
    Dim RS As ADODB.Recordset
    Dim sSQL As String
    Dim lMainSPVersion As Long
    Dim lVersion As Long
    Dim sAssignmentsID As String
    Dim sCLIENTNUM As String
    Dim sParams As String
    Dim MyAdjReport As Object
    Dim MyPrn As Printer
    Dim oCarList As V2ECKeyBoard.clsCarLists
    
    sCLIENTNUM = sAPPDocClientClaimNo
    
    If sCLIENTNUM = vbNullString Then
        Exit Sub
    End If
    
    Set oConn = New ADODB.Connection
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    'need to get the Main SPVersion  and AssignemtsID for this Assignment
    sSQL = "SELECT A.AssignmentsID, "
    sSQL = sSQL & "A.SPVersion "
    sSQL = sSQL & "FROM Assignments A "
    sSQL = sSQL & "INNER JOIN CLIENTCOMPANYCATSPEC CCCS ON (A.ClientCompanyCatSpecID = CCCS.ClientCompanyCatSpecID) "
    sSQL = sSQL & "WHERE A.ClientCompanyCatSpecID IN "
    sSQL = sSQL & "( "
    sSQL = sSQL & "SELECT  ClientCompanyCatSpecID "
    sSQL = sSQL & "FROM ClientCompanyCatSpec "
    sSQL = sSQL & "WHERE ClientCompanyID = " & goUtil.gsCurCar & " "
    sSQL = sSQL & "AND     CATID = " & goUtil.gsCurCat & " "
    sSQL = sSQL & ") "
    sSQL = sSQL & "AND A.AdjusterSpecID IN "
    sSQL = sSQL & "( "
    sSQL = sSQL & "SELECT  ClientCoAdjusterSpecID "
    sSQL = sSQL & "FROM ClientCoAdjusterSpec "
    sSQL = sSQL & "Where ClientCompanyID = " & goUtil.gsCurCar & " "
    sSQL = sSQL & "AND UsersID = " & goUtil.gsCurUsersID & " "
    sSQL = sSQL & ") "
    sSQL = sSQL & "AND ( "
    sSQL = sSQL & "IIF(Instr(1, A.CLIENTNUM, '-') > 0 ,"
    'Then
    sSQL = sSQL & "Left(A.CLIENTNUM, Instr(1, A.CLIENTNUM, '-')-1) + Mid(A.CLIENTNUM, Instr(1, A.CLIENTNUM, '-')+1), "
    'Else
    sSQL = sSQL & "A.CLIENTNUM) "
    sSQL = sSQL & ") = '" & goUtil.utCleanSQLString(sCLIENTNUM) & "' "
    
    Set RS = New ADODB.Recordset
    RS.CursorLocation = adUseClient
    RS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set RS.ActiveConnection = Nothing
    
    If RS.RecordCount = 0 Then
        Exit Sub
    End If
    
    Set MyPrn = moLoss.ParametersCol("PRINTER_OBJECT").ParamValue
    lMainSPVersion = RS.Fields("SPVersion").Value
    sAssignmentsID = RS.Fields("AssignmentsID").Value
    
    
    For Each vAppDoc In colAppDocs
        MyAppDoc = vAppDoc
        If MyAppDoc.Selected Then
            sReportName = MyAppDoc.DocName
            
            'These Params should always be the same for each report.
            sParams = vbNullString
            sParams = sParams & "psAssignmentsID=" & sAssignmentsID & "|"
            sParams = sParams & "pbPreview=" & "False" & "|"
            sParams = sParams & "pbGetObjectOnly=" & "True" & "|"
            
            Set oCarList = CreateObject("V2ECcarTWIA.clsLists")
            
            If StrComp(sReportName, "WorkSheet Diagram", vbTextCompare) = 0 Then
                'WorkSheet Diagram
                lVersion = goUtil.GetApplicationVersionNumber(lMainSPVersion, "ECrptTWIA_arWorkSheetDiag", oConn)
                Set MyAdjReport = oCarList.GetARReport("ECrptTWIA_arWorkSheetDiag.clsLists", lVersion, sParams)
            
            ElseIf StrComp(sReportName, "Activity Log (Blank)", vbTextCompare) = 0 Then
                'Activity Log (Blank)
                'Add Param for Blank Report
                sParams = sParams & "pbPrintBlankReport=" & "True" & "|"
                lVersion = goUtil.GetApplicationVersionNumber(lMainSPVersion, "ECRptTWIA_arActivityLog", oConn)
                Set MyAdjReport = oCarList.GetARReport("ECRptTWIA_arActivityLog.clsLists", lVersion, sParams)
                
            ElseIf StrComp(sReportName, "TWIA Claims Explainer", vbTextCompare) = 0 Then
                'TWIA Claims Explainer
                lVersion = goUtil.GetApplicationVersionNumber(lMainSPVersion, "ECrptTWIA_arClaimsExplain", oConn)
                Set MyAdjReport = oCarList.GetARReport("ECrptTWIA_arClaimsExplain.clsLists", lVersion, sParams)
                'need two copies
                MyAdjReport.Printer.Copies = 2
            End If
            
            'Set the Printer Device
             On Error Resume Next
            MyAdjReport.Printer.DeviceName = MyPrn.DeviceName
            If Err.Number > 0 Then
                Err.Clear
            End If
            MyAdjReport.Printer.Port = MyPrn.Port
            If Err.Number > 0 Then
                Err.Clear
            End If
            On Error GoTo EH
            
            'Prin the Report
            MyAdjReport.PrintReport False 'False means no Print Dialog to be shown
            oCarList.CLEANUP
            Set oCarList = Nothing
        End If
    Next
    
    'Clean up
    Set oConn = Nothing
    Set RS = Nothing
    Set MyPrn = Nothing
    Exit Sub
EH:
    Screen.MousePointer = vbDefault
   goUtil.utErrorLog Err, App.EXEName, clsCarLists_ClassName, "Private Sub moLoss_PrintAppDoc", True
End Sub

Private Sub moLoss_UpdateDB(ByVal oLossReport As V2ECKeyBoard.clsCarLR)
'    On Error GoTo EH
'    Dim sSQL As String
'    Dim sMess As String
'    Dim sSQLError As String
'    Dim sIBNum As String
'    Dim sPolicyDesc As String
'    Dim sBuildLim As String
'    Dim sContLim As String
'    Dim sDeductible As String
'    Dim Param As V2ECKeyBoard.udtParameter
'    Dim lVersion As Long
'    Dim sAppEXEName As String
'
'    lVersion = CLng(App.Major & App.Minor & App.Revision)
'    sAppEXEName = App.EXEName
'
'    sSQL = oLossReport.GetLRSQL(sAppEXEName, lVersion, sMess)
'    goUtil.gMainDB.Execute sSQL
'
'    'IB NUM
'    Param = moLoss.ParametersCol("IB_NUMBER")
'    sIBNum = Param.ParamValue
'    'Policy Desc
'    Param = moLoss.ParametersCol("POLICY_DESC")
'    sPolicyDesc = Param.ParamValue
'    'Building Lim
'    Param = moLoss.ParametersCol("BUILD_LIM")
'    sBuildLim = Param.ParamValue
'    'Contents Lim
'    Param = moLoss.ParametersCol("CONT_LIM")
'    sContLim = Param.ParamValue
'    'Deductible
'    Param = moLoss.ParametersCol("DEDUCTIBLE")
'    sDeductible = Param.ParamValue
'    'Add Log Entry
'    AddLogEntries sIBNum, Now, , sPolicyDesc, sBuildLim, sContLim, sDeductible
'
'    Exit Sub
'EH:
'    sMess = sMess & "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
'    sMess = sMess & "ERROR # " & Err.Number & vbCrLf
'    sMess = sMess & Err.Description & vbCrLf
'    sMess = sMess & "Report Format: " & oLossReport.ClassName & vbCrLf
'    sMess = sMess & "Could not Add """ & oLossReport.PrnPath & """ into Database." & vbCrLf
'    sMess = sMess & "This Loss Report is assigned to: " & oLossReport.Adjuster & "_" & oLossReport.CRID & vbCrLf
'    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
'    moLoss_ErrorMess sMess
'    Err.Clear
End Sub

