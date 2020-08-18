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
    'Set the Safeco DB version here
    clsCarLists_ECDBVersion = 1
End Property

Public Property Get clsCarLists_FormatList() As Collection
    Set clsCarLists_FormatList = New Collection
    'Lists all the available Report format classes
    'available in this Carrrier object.
    With clsCarLists_FormatList
        
    End With
    
End Property

Public Property Get clsCarLists_ActionList() As Collection
    On Error GoTo EH
    Dim oULSafeco As New clsULSafeco
    Dim oULSafecoRT As New clsULSafecoRT
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    'Lists all the available "ACTION" Classes
    'Action= classes that
    
    Set clsCarLists_ActionList = New Collection
    
    With clsCarLists_ActionList
        .Add oULSafeco, oULSafeco.clsCarUL_ClassName
        .Add oULSafecoRT, oULSafecoRT.clsCarUL_ClassName
    End With
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
            Set moReportLists = CreateObject(psReportName)
            moReportLists.SetUtilObject goUtil
            Set clsCarLists_GetARMiscDelimParamsCol = moReportLists.GetARMiscDelimParamsCol(plVersion, psDelimParams)
    End Select
    
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
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
    clsCarLists_GetARReport = Nothing
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
    oLRs.CleanUp
    Set oLRs = Nothing
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    clsCarLists_ExportARReport = False
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLists_ClassName & vbCrLf & "Public Function ExportME"
End Function

Public Function clsCarLists_SetGlobalObjects(pcolObjects As Collection) As Boolean
    On Error Resume Next
    
    Set goUtil = pcolObjects("goUtil")
    
    Set moGUI = New ECguiSafeco.clsGUI
    moGUI.SetGlobalObjects pcolObjects
    
    clsCarLists_SetGlobalObjects = True
    
End Function

Public Sub clsCarLists_PopulateECTree(poNodX As Object, poECTree As Object, psCarCatKey As String, psCar As String, psCAT As String)
    On Error GoTo EH
    Dim sCarList As String
    Dim lErrNum As Long
    Dim sErrDesc As String
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
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLists_ClassName & vbCrLf & "Public Sub PopulateTree"
End Sub

Public Function clsCarLists_CarListCommand(psCar As String, psCAT As String, psCommand As String) As Boolean
    On Error GoTo EH
    Dim oGui As ECguiSafeco.clsGUI
    Dim lErrNum As Long
    Dim sErrDesc As String
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
    
    ElseIf StrComp(psCommand, "MakeCurrent", vbTextCompare) = 0 Then
       
    End If
    clsCarLists_CatCommand = True
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    clsCarLists_CatCommand = False
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLists_ClassName & vbCrLf & "Public Function CatCommand"
End Function

Public Function clsCarLists_SendToXactimate(Optional psClientClaimNo As String = vbNullString) As Boolean
    On Error GoTo EH
    Dim XProj As V2ECKeyBoard.udtXactProject
    Dim vXProj As Variant
    Dim RS As DAO.Recordset
    Dim sSQL As String
    Dim colXProjects As Collection
    Dim sCity As String
    Dim sZip As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    sSQL = "SELECT * "
    sSQL = sSQL & "FROM " & DB_ASSIGNMENTS & " A "
    'If they are in a specific claim then only get this one claim to be sent to
    'xactimate. Otherwise we are sending all projects to xactimate.
    If psClientClaimNo <> vbNullString Then
        sSQL = sSQL & "WHERE A.ClientClaimNo = '" & goUtil.utCleanSQLString(psClientClaimNo) & "' "
        sSQL = sSQL & "AND A.Status In('NEW', 'REOPEN') "
    Else
        sSQL = sSQL & "WHERE A.Status In('NEW', 'REOPEN') "
    End If
    sSQL = sSQL & "ORDER BY A.Status, A.ClaimNo "
    Set RS = goUtil.gCurDB.OpenRecordset(sSQL)

    If Not RS.EOF Then
        RS.MoveFirst
    Else
        clsCarLists_SendToXactimate = True 'No errors
        GoTo SKIPXACT
    End If
    
    Set goUtil.goXact = New V2ECKeyBoard.clsXact
    goUtil.goXact.StartHwnd = goUtil.gfrmECTray.hWnd
    
    If Not goUtil.goXact.LookUpLoaded Then
        goUtil.goXact.CleanUp
        Set goUtil.goXact = Nothing
        clsCarLists_SendToXactimate = True
        GoTo SKIPXACT
    End If
    
    'BGS 3.1.2002 OK it might help if there is some more comments here now that I have 2 seconds.
    'First we are loading Xproj user defined type with all the Admin information that will
    'be typed into xactimate.
    
    Do Until RS.EOF
        With XProj
            .SentToXact = IIf(IsNull(RS!InspectedWindowsScreens), False, RS!InspectedWindowsScreens)
            With .Main
                .Name = IIf(IsNull(RS!NewInsuredNames), vbNullString, RS!NewInsuredNames)
                .Street = IIf(IsNull(RS!LossLocAddress), vbNullString, RS!LossLocAddress)
                .City = IIf(IsNull(RS!LossLocCity), vbNullString, RS!LossLocCity)
                .State = IIf(IsNull(RS!LossLocState), vbNullString, RS!LossLocState)
                .Zip = IIf(IsNull(RS!LossLocZip), vbNullString, RS!LossLocZip)
                .DateEntered = Format(Now, "MM/DD/YYYY")
                .DateInspected = IIf(IsNull(RS!InspectedDate), vbNullString, Format(Mid(RS!InspectedDate, 5, 2) & "/" & Mid(RS!InspectedDate, 7, 2) & "/" & left(RS!InspectedDate, 4), "MM/DD/YYYY"))
                .DateOfLoss = IIf(IsNull(RS!LossDate), vbNullString, Format(Mid(RS!LossDate, 5, 2) & "/" & Mid(RS!LossDate, 7, 2) & "/" & left(RS!LossDate, 4), "MM/DD/YYYY"))
                .DateReceived = IIf(IsNull(RS!ReceivedDate), vbNullString, Format(Mid(RS!ReceivedDate, 5, 2) & "/" & Mid(RS!ReceivedDate, 7, 2) & "/" & left(RS!ReceivedDate, 4), "MM/DD/YYYY"))
            End With
            With .Loss
                .ClaimNumber = IIf(IsNull(RS!ClientClaimNo), vbNullString, RS!ClientClaimNo)
                .PolicyNumber = IIf(IsNull(RS!PolicyNumber), vbNullString, RS!PolicyNumber)
                .Deductible = IIf(IsNull(RS!Deductible), vbNullString, RS!Deductible)
                .TypeOfLoss = IIf(IsNull(RS!TypeOfLoss), vbNullString, RS!TypeOfLoss)
                .CatCode = IIf(IsNull(RS!CatCode), vbNullString, RS!CatCode)
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
        RS.MoveNext
    Loop
    
    'BGS 3.1.2002 We are done with the RS here so we can
    'clear it so if say someone clicks on make cat current while
    'we are in the middle of typing info into xactimate we won't get
    'error
    Set RS = Nothing
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
                If vXProj.IsDirty Then
                    With vXProj
                        sSQL = "UPDATE " & DB_ASSIGNMENTS & " "
                        sSQL = sSQL & "SET "
                        sSQL = sSQL & DB_ASSIGNMENTS & ".NewInsuredNames = '" & goUtil.utCleanSQLString(.Main.Name) & "', "
                        sSQL = sSQL & DB_ASSIGNMENTS & ".LossLocAddress = '" & goUtil.utCleanSQLString(.Main.Street) & "', "
                        sCity = Replace(goUtil.utCleanSQLString(.Main.City), Chr(32), Chr(160))
                        sSQL = sSQL & DB_ASSIGNMENTS & ".LossLocCity = '" & sCity & "', "
                        sSQL = sSQL & DB_ASSIGNMENTS & ".LossLocState = '" & goUtil.utCleanSQLString(.Main.State) & "', "
                        sZip = .Main.Zip
                        sZip = goUtil.utCleanSQLString(sZip)
                        sZip = Replace(sZip, "00000-0000", vbNullString)
                        sSQL = sSQL & DB_ASSIGNMENTS & ".LossLocZip = '" & sZip & "', "
                        sSQL = sSQL & DB_ASSIGNMENTS & ".NewLossLocation = '" & goUtil.utCleanSQLString(goUtil.utJoinToAddress(.Main.Street, sCity, .Main.State, sZip)) & "', "
                        sSQL = sSQL & DB_ASSIGNMENTS & ".InspectedDate = '" & goUtil.utCleanSQLString(Format(.Main.DateInspected, "YYYYMMDD")) & "', "
                        sSQL = sSQL & DB_ASSIGNMENTS & ".LossDate = '" & goUtil.utCleanSQLString(Format(.Main.DateOfLoss, "YYYYMMDD")) & "', "
                        sSQL = sSQL & DB_ASSIGNMENTS & ".ReceivedDate = '" & goUtil.utCleanSQLString(Format(.Main.DateReceived, "YYYYMMDD")) & "', "
                        sSQL = sSQL & DB_ASSIGNMENTS & ".ClientClaimNo = '" & goUtil.utCleanSQLString(.Loss.ClaimNumber) & "', "
                        sSQL = sSQL & DB_ASSIGNMENTS & ".PolicyNumber = '" & goUtil.utCleanSQLString(.Loss.PolicyNumber) & "', "
                        sSQL = sSQL & DB_ASSIGNMENTS & ".Deductible = '" & goUtil.utCleanSQLString(.Loss.Deductible) & "', "
                        sSQL = sSQL & DB_ASSIGNMENTS & ".TypeOfLoss = '" & goUtil.utCleanSQLString(.Loss.TypeOfLoss) & "', "
                        sSQL = sSQL & DB_ASSIGNMENTS & ".CatCode = '" & goUtil.utCleanSQLString(.Loss.CatCode) & "' "
                        sSQL = sSQL & "WHERE "
                        sSQL = sSQL & DB_ASSIGNMENTS & ".ClientClaimNo = '" & goUtil.utCleanSQLString(.Loss.ClaimNumber) & "' "
                        goUtil.gCurDB.Execute sSQL
                    End With
                End If
            Next
        End If
        
        'Display the Cancel Form
        Set mFrmCancel = New frmCancel
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
    Set RS = Nothing
    
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
        goUtil.goXact.CleanUp
    End If
    Set goUtil.goXact = Nothing
    Set colXProjects = Nothing
    Set RS = Nothing
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLists_ClassName & vbCrLf & "Public Function SendToXactimate"
End Function

Public Function clsCarLists_FlagSentToXactimate(pvXProj As Variant) As Boolean
    On Error GoTo EH
    Dim XProj As V2ECKeyBoard.udtXactProject
    Dim sSQL As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    XProj = pvXProj
    
    With XProj
        sSQL = "UPDATE " & DB_ASSIGNMENTS & " "
        sSQL = sSQL & "SET "
        'Yes Yes don't blame me for using InspectedWindowScreens as Send to Xact Flag
        'Only update this flag if we are successfull sending it to Xactimate
        sSQL = sSQL & DB_ASSIGNMENTS & ".InspectedWindowsScreens = " & .SentToXact & "  "
        sSQL = sSQL & "WHERE "
        sSQL = sSQL & DB_ASSIGNMENTS & ".ClientClaimNo = '" & goUtil.utCleanSQLString(.Loss.ClaimNumber) & "' "
        If Not goUtil Is Nothing Then
            If Not goUtil.gCurDB Is Nothing Then
                goUtil.gCurDB.Execute sSQL
            Else
                Exit Function
            End If
        End If
    End With
    
    clsCarLists_FlagSentToXactimate = True
    
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
    Set moLoss = Nothing
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
    Dim lCount As Long
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    Set clsCarLists_GetAppDocsCol = New Collection
    
    '---------------->NOTE<-----------------------------
    'Maximum of 14 App docs may be added to the collection
    '---------------->NOTE<-----------------------------
    'Put the Appended documents in the order they should be printed
    'after each Loss report is printed
    
    For lCount = 1 To 14
        sDocName = Switch(lCount = 1, "Test01", _
                          lCount = 2, "Test02", _
                          lCount = 3, "test03", _
                          lCount = 4, vbNullString, _
                          lCount = 5, vbNullString, _
                          lCount = 6, vbNullString, _
                          lCount = 7, vbNullString, _
                          lCount = 8, vbNullString, _
                          lCount = 9, vbNullString, _
                          lCount = 10, vbNullString, _
                          lCount = 11, vbNullString, _
                          lCount = 12, vbNullString, _
                          lCount = 13, vbNullString, _
                          lCount = 14, vbNullString)
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

Public Function clsCarLists_CleanUp() As Boolean
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
        moGUI.CleanUp
        Set moGUI = Nothing
    End If
    
    If Not moLoss Is Nothing Then
        moLoss.CleanUp
        Set moLoss = Nothing
    End If
     If Not moReportLists Is Nothing Then
        moReportLists.CleanUp
        Set moReportLists = Nothing
    End If
    
    Set goUtil.gARV = Nothing
    Set goUtil = Nothing
    clsCarLists_CleanUp = True
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    clsCarLists_CleanUp = False
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLists_ClassName & vbCrLf & "Public Function CLEANUP"
End Function

Private Sub moLoss_CleanUpLossReports()
    If Not moLoss Is Nothing Then
        moLoss.CleanUp
        Set moLoss = Nothing
    End If
End Sub

Private Sub moLoss_ErrorMess(ByVal Mess As String)
    On Error GoTo EH
    Dim vMess As Variant
    Dim lCount As Long
    vMess = Split(Mess, vbCrLf)
    For lCount = LBound(vMess) To UBound(vMess)
        If goUtil.utFormExists(Forms, "frmCommStatus") Then
            goUtil.gfrmCommStatus.lstProcess.AddItem vMess(lCount)
        End If
    Next lCount
    
    Err.Raise -999, , Mess
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, clsCarLists_ClassName, "Private Sub moLoss_ErrorMess", False
End Sub

Private Sub moLoss_PrintAppDoc(ByVal colAppDocs As Collection, ByVal sAPPDocClientClaimNo As String)
'    On Error GoTo EH
'    Dim sSelFormula As String
'    Dim sReportName As String
'    Dim iPos As Integer
'    Dim sMess As String
'    Dim MyAppDoc As V2ECKeyBoard.udtAppDoc
'    Dim vAppDoc As Variant
'    Dim MyPrn As Printer
'
'    frmCrystal.Report1.Destination = crptToPrinter
'    'Save Printer Settings
'    Set MyPrn = moLoss.ParametersCol("PRINTER_OBJECT").ParamValue
'
'    If frmCrystal.Report1.PrinterName <> MyPrn.DeviceName Then
'        frmCrystal.Report1.PrinterName = MyPrn.DeviceName
'        SaveSetting App.EXEName, "PRINTER", "PRINTER_NAME", frmCrystal.Report1.PrinterName
'    End If
'    If frmCrystal.Report1.PrinterPort <> MyPrn.Port Then
'        frmCrystal.Report1.PrinterPort = MyPrn.Port
'        SaveSetting App.EXEName, "PRINTER", "PRINTER_PORT", frmCrystal.Report1.PrinterPort
'    End If
'    If frmCrystal.Report1.PrinterDriver <> MyPrn.DriverName Then
'        frmCrystal.Report1.PrinterDriver = MyPrn.DriverName
'        SaveSetting App.EXEName, "PRINTER", "PRINTER_DRIVER", frmCrystal.Report1.PrinterDriver
'    End If
'
'
'    For Each vAppDoc In colAppDocs
'        MyAppDoc = vAppDoc
'
'        If MyAppDoc.Selected Then
'            sReportName = MyAppDoc.DocName
'            With frmCrystal.Report1
'                Select Case sReportName
'                    '1 data file
'                    Case "FarmersActivityLogBlank.rpt", "Commission Report.rpt", "FarmersContactLetter.rpt", "FarmersLetter.rpt", "FarmersNCCAdvice.rpt", "IBReport.rpt", "LossReport.rpt", "PrintLoss.rpt", "ReplacementCostAgreementFarmers.rpt", "Production Report.rpt", "ReplacementCostAgreementFarmersTX.rpt", "WorksheetFarmers.rpt", "WorksheetFarmersBlank.rpt"
'                        .DataFiles(0) = CurrentDirectory & "\Eberls.mdb"
'                        .DataFiles(1) = vbNullString
'                        .DataFiles(2) = vbNullString
'                        sSelFormula = "{Assignments.ClientClaimNo} = " & Chr(34) & sAPPDocClientClaimNo & Chr(34) & " "
'
'                    Case Else 'If the report name is not found then the rpeort name changed need to fix it
'                        MsgBox "Can't find Data settings for Report '" & sReportName & "'", vbInformation, "Report Settings not found"
'                        Exit Sub
'                End Select
'
'                .ReportFileName = CurrentDirectory & "\" & sReportName
'                .SelectionFormula = sSelFormula
'                Select Case sReportName
'                    Case "FarmersLetter.rpt"
'                        'Issue 179 4.5.2002 Claim Letter should print 2 copies
'                        'OB Kay B Lets print 2 copies
'                        .CopiesToPrinter = 2
'                End Select
'                'BGS launch the report
'                .Action = PRINT_REPORT
'                .CopiesToPrinter = 1 'Always set this back to 1 Copy after printing
'            End With
'        End If
'    Next
'    'Clean up
'    Set MyPrn = Nothing
'    Exit Sub
'EH:
'    Screen.MousePointer = vbDefault
'    If Err.Number = 20545 Then
'        MsgBox "Please set up a default printer for Crystal Reports first.", vbInformation, "Default Printer Not Found"
'        EberlsMain.mnuPrinterSetUp_Click
'        Exit Sub
'    End If
'    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
'    sMess = sMess & "ERROR # " & Err.Number & vbCrLf
'    sMess = sMess & Err.Description & vbCrLf
'    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
'    moLoss_ErrorMess sMess
'    Err.Clear
'    Set MyPrn = Nothing
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
'    goUtil.gCurDB.Execute sSQL
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

