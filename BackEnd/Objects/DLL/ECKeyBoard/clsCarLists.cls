VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCarLists"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Carrier Lists Polymorphism
'Carrier Lists InterFace Definiton
'Use Implements V2ECKeyBoard.clsCarLists in ALL clsLists For All Carrier Objects (Not in V2ECKeyBoard.clsLists, However!)

'Implements V2ECKeyBoard.clsCarLists
'
'Private moARReport As Object
'Private moGUI As ECgui???.clsGUI
'Loss Reports Object
'Private WithEvents moLoss As V2ECKeyBoard.clsLossReports ' New Loss Notices 4.11.2002
Public Property Get ClassName() As String
'    clsCarLists_ClassName = App.EXEName & ".clsLists"
End Property

Public Property Get goForms() As Object
'    Set clsCarLists_goForms = mogui.goForms
End Property

Public Property Let LossReports(poLoss As V2ECKeyBoard.clsLossReports)
    'Set moLoss = poLoss
End Property
Public Property Set LossReports(poLoss As V2ECKeyBoard.clsLossReports)
    'Set moLoss = poLoss
End Property
Public Property Get LossReports() As V2ECKeyBoard.clsLossReports
    'Set clsCarLists_LossReports = moLoss
End Property

Public Property Let ARReport(poAR As Object)
'    Set moARReport = poAR
End Property
Public Property Set ARReport(poAR As Object)
'    Set moARReport = poAR
End Property
Public Property Get ARReport() As Object
'    Set clsCarLists_ARReport = moARReport
End Property

Public Property Get ECDBVersion() As String
'    'Set the Carrier (??Farmers) DB version here
'    clsCarLists_ECDBVersion = ??
End Property

Public Property Get FormatList() As Collection
'    Set FormatList = New Collection
'    'Lists all the available Report format classes
'    'available in this Carrrier object.
'    With clsCarLists_FormatList
'        .Add "clsLoss???", "clsLoss???"
'        .Add "clsLoss???", "clsLoss???"
'    End With
    
End Property

Public Property Get ActionList() As Collection
'    Set ActionList = New Collection
'    'Lists all the available "ACTION" Classes
'    'Action= classes that
'    With clsCarLists_ActionList
'        .Add "clsCarUL???", "clsCarUL???"
'        .Add "clsCarUL???RT", "clsCarUL???RT"
'    End With
End Property

Public Function GetARMiscDelimParamsCol(psReportName As String, _
                                    Optional plVersion As Long = -1, _
                                    Optional psDelimParams As String = vbNullString) As Collection
'    On Error GoTo EH
'    Dim lErrNum As Long
'    Dim sErrDesc As String

    
'    Exit Function
'EH:
'    lErrNum = Err.Number
'    sErrDesc = Err.Description
'    Err.Raise lerrnum, , serrdesc & vbCrLf & clsCarLists_ClassName & vbCrLf & "Public Function GetARMiscDelimParamsCol"
End Function

Public Function GetARMiscDelimParamsStr(psReportName As String, _
                                    Optional plVersion As Long = -1, _
                                    Optional psDelimParams As String = vbNullString) As String
'    On Error GoTo EH
'    Dim lErrNum As Long
'    Dim sErrDesc As String

    
'    Exit Function
'EH:
'    lErrNum = Err.Number
'    sErrDesc = Err.Description
'    Err.Raise lerrnum, , serrdesc & vbCrLf & clsCarLists_ClassName & vbCrLf & "Public Function GetARMiscDelimParamsStr"
End Function

Public Function GetARReport(psReportName As String, _
                            Optional plVersion As Long = -1, _
                            Optional psDelimParams As String = vbNullString) As Object
'    On Error GoTo EH
'    Dim lErrNum As Long
'    Dim sErrDesc As String

    
'
'    Select Case UCase(psReportName)
'        Case "???ReportName???"
'            Set clsCarLists_GetARReport = New ECcar???.arRpt???ReportName
'    End Select
'
'    'Set ref to this Lists Instance
'    Set clsCarLists_GetARReport.Lists = Me
'    Set ARReport = clsCarLists_GetARReport
'    Exit Function
'EH:
'    lErrNum = Err.Number
'    sErrDesc = Err.Description
'    Err.Raise lerrnum, , serrdesc & vbCrLf & clsCarLists_ClassName & vbCrLf & "Public Function GetARReport"
End Function

Public Sub SetARReportProperty(psName As String, pvValue As Variant, pType As VbVarType)
'    On Error GoTo EH
'    Dim lErrNum As Long
'    Dim sErrDesc As String

'
'    If Not moARReport Is Nothing Then
'        moARReport.SetProperty psName, pvValue, pType
'    End If
'    Exit Sub
'EH:
'    lErrNum = Err.Number
'    sErrDesc = Err.Description
'    Err.Raise lerrnum, , serrdesc & vbCrLf & clsCarLists_ClassName & vbCrLf & "Public Sub SetARReportProperty"
End Sub

Public Function CFExportARReport(psXportPath As String, pXportType As ExportType) As Boolean
'    On Error GoTo EH
'    Dim lErrNum As Long
'    Dim sErrDesc As String

    
'    If Not moARReport Is Nothing Then
'        clsCarLists_CFExportARReport = moARReport.ExportME(psXportPath, pXportType)
'    End If
'    Exit Function
'EH:
'    lErrNum = Err.Number
'    sErrDesc = Err.Description
'     Err.Raise lerrnum, , serrdesc & vbCrLf & clsCarLists_ClassName & vbCrLf & "Public Function CFExportARReport"
End Function

Public Function ExportARReport(poARReport As Object, psXportPath As String, pXportType As ExportType) As Boolean
'    On Error GoTo EH
'    Dim oLRs As V2ECKeyBoard.clsLossReports
'    Dim oAR As Object
'    Dim lErrNum As Long
'    Dim sErrDesc As String

    
'
'    If Not poARReport Is Nothing Then
'        Set moARReport = poARReport
'    End If
'    If Not moARReport Is Nothing Then
'        Set oLRs = New V2ECKeyBoard.clsLossReports
'    Else
'        Exit Function
'    End If
'    With moARReport
'        .Printer.Orientation = ddOPortrait
'    End With
'
'    If oLRs.ExportFile(moARReport, psXportPath, pXportType) Then
'        clsCarLists_ExportARReport = True
'    End If
'
'    'Clean up
'    Unload moARReport
'    Set moARReport = Nothing
'    oLRs.CLEANUP
'    Set oLRs = Nothing
'
'    Exit Function
'EH:
'    lErrNum = Err.Number
'    sErrDesc = Err.Description
'    Err.Raise lerrnum, , serresc & vbCrLf & clsCarLists_ClassName & vbCrLf & "Public Function ExportME"
End Function

Public Function SetGlobalObjects(pcolObjects As Collection) As Boolean
'    On Error Resume Next
'
'    Set goUtil = pcolObjects("goUtil")
'
'
'    Set moGUI = New ECgui???.clsGUI
'    moGUI.SetGlobalObjects pcolObjects
    
End Function


Public Sub PopulateECTree(poNodX As Object, poECTree As Object, psCarCatKey As String, psCar As String, psCAT As String)
'    On Error GoTo EH
'    Dim sCarList As String
'    Dim lErrNum As Long
'    Dim sErrDesc As String
    
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
'    Exit Sub
'EH:
'    lErrNum = Err.Number
'    sErrDesc = Err.Description
'    Err.Raise lerrnum, , serrdesc & vbCrLf & clsCarLists_ClassName & vbCrLf & "Public Sub PopulateTree"
End Sub

Public Function CarListCommand(psCar As String, psCAT As String, psCommand As String) As Boolean
'    On Error GoTo EH
'    Dim oGui As ECgui???.clsGUI
'    Dim lErrNum As Long
'    Dim sErrDesc As String

    
'    Set oGui = moGUI
'
'
'
'    Set oGui = Nothing
'
'    CarListCommand = True
'    Exit Function
'EH:
'    lErrNum = Err.Number
'    sErrDesc = Err.Description
'    Err.Raise lerrnum, , serrdesc & vbCrLf & clsCarLists_ClassName & vbCrLf & "Public Function CarListCommand"
End Function

Public Function CatCommand(psCar As String, psCAT As String, psCommand As String) As Boolean

End Function

Public Function SendToXactimate(Optional psClientClaimNo As String = vbNullString) As Boolean
    
End Function

Public Function FlagSentToXactimate(pvXProj As Variant) As Boolean
    
End Function

'<---------------------------Implement New Loss Notice Here-------------------->
Public Sub LoadLoss(pbUpdateDB As Boolean)

End Sub

Public Sub ADDAppDocs()
'    On Error GoTo EH
'    Dim lErrNum As Long
'    Dim sErrDesc As String

    
'    '---------------->NOTE<-----------------------------
'    'Maximum of 14 App docs may be added to the collection
'    '---------------->NOTE<-----------------------------
'    If Not goUtil.goCurCarList Is Nothing Then
'        Set moLoss.AppDocsCol = goUtil.goCurCarList.GetAppDocsCol
'    End If
'
'    Exit Sub
'EH:
'    lErrNum = Err.Number
'    sErrDesc = Err.Description
'    Err.Raise lerrnum, , serrdesc & vbCrLf & clsCarLists_ClassName & vbCrLf & "Public Sub ADDAppDocs"
End Sub

Public Sub ADDParameters()
'    On Error GoTo EH
'    Dim Param As V2ECKeyBoard.udtParameter
'    Dim sFName As String 'First Name Adjuster
'    Dim sLName As String 'Last Name Adjuster
'    Dim dTaxPercent As Double
'    Dim sTypeOfLoss As String
'    Dim sSiteAddress As String
'    Dim sSiteState As String
'    Dim sSiteCity As String
'    Dim sSiteZip As String
'    Dim lMaxID As Long
'    Dim sUserName As String
'    Dim sCATCode As String
'    Dim sSSN As String
'    Dim sEmail As String
'    Dim sContactPhone As String
'    Dim sAppVSInfo As String
'    Dim lErrNum As Long
'    Dim sErrDesc As String

    
'
'    'Add Cat Code
'    sCATCode = GetSetting(App.EXEName, "GENERAL", "CAT_CODE", 0#)
'    Param.ParamName = "CAT_CODE"
'    Param.ParamValue = sCATCode
'    moLoss.AddParameter Param
'
'    'Add "TAX_PERCENT"
'    dTaxPercent = GetSetting(App.EXEName, "FEES", "TAX_PERCENT", 0#)
'    Param.ParamName = "TAX_PERCENT"
'    Param.ParamValue = dTaxPercent
'    moLoss.AddParameter Param
'
'    'Add "IB_PREFIX"
'    'First lettter of First and first letter of Last
'    sFName = UCase(GetSetting(App.EXEName, "GENERAL", "ADJUSTOR_FIRST_NAME", vbNullString))
'    sLName = UCase(GetSetting(App.EXEName, "GENERAL", "ADJUSTOR_LAST_NAME", vbNullString))
'    Param.ParamName = "IB_PREFIX"
'    Param.ParamValue = Left(sFName, 1) & Left(sLName, 1)
'    moLoss.AddParameter Param
'
'    'Add First Name
'    Param.ParamName = "ADJ_FIRST_NAME"
'    Param.ParamValue = sFName
'    moLoss.AddParameter Param
'
'    'Add Last Name
'    Param.ParamName = "ADJ_LAST_NAME"
'    Param.ParamValue = sLName
'    moLoss.AddParameter Param
'
'    'Add UserName
'    sCRID = GetSetting(App.EXEName, "GENERAL", "USER_NAME", vbNullString)
'    Param.ParamName = "CRID"
'    Param.ParamValue = sCRID
'    moLoss.AddParameter Param
'
'    'Add SSN
'    sSSN = GetSetting("ECS", "WEB_SECURITY", "SSN", vbNullString)
'    Param.ParamName = "ADJ_SSN"
'    Param.ParamValue = sSSN
'    moLoss.AddParameter Param
'
'    'Add Email
'    sEmail = GetSetting(App.EXEName, "GENERAL", "ADJ_EMAIL", vbNullString)
'    Param.ParamName = "ADJ_EMAIL"
'    Param.ParamValue = sEmail
'    moLoss.AddParameter Param
'
'    'Add Contact Phone
'    sContactPhone = GetSetting(App.EXEName, "GENERAL", "ADJ_CONTACT_PHONE", vbNullString)
'    Param.ParamName = "ADJ_CONTACT_PHONE"
'    Param.ParamValue = sContactPhone
'    moLoss.AddParameter Param
'
'    'Add AppVSInfo
'    sAppVSInfo = goUtil.utGetAppVSInfo(App.EXEName, App.Path)
'    Param.ParamName = "ADJ_AppVSInfo"
'    Param.ParamValue = sAppVSInfo
'    moLoss.AddParameter Param
'
'    'Add "TYPE_OF_LOSS"
'    sTypeOfLoss = GetSetting(App.EXEName, "GENERAL", "TYPE_OF_LOSS", vbNullString)
'    Param.ParamName = "TYPE_OF_LOSS"
'    Param.ParamValue = sTypeOfLoss
'    moLoss.AddParameter Param
'
'    'Add SiteAddress
'    sSiteAddress = GetSetting(App.EXEName, "GENERAL", "ADDRESS_1", vbNullString)
'    Param.ParamName = "SITE_ADDRESS"
'    Param.ParamValue = sSiteAddress
'    moLoss.AddParameter Param
'
'    'Add SiteState
'    sSiteState = GetSetting(App.EXEName, "GENERAL", "STATE", vbNullString)
'    Param.ParamName = "SITE_STATE"
'    Param.ParamValue = sSiteState
'    moLoss.AddParameter Param
'
'    'Add SiteCity
'    sSiteCity = GetSetting(App.EXEName, "GENERAL", "CITY", vbNullString)
'    Param.ParamName = "SITE_CITY"
'    Param.ParamValue = sSiteCity
'    moLoss.AddParameter Param
'
'    'Add SiteZip
'    sSiteZip = GetSetting(App.EXEName, "GENERAL", "ZIP", vbNullString)
'    Param.ParamName = "SITE_ZIP"
'    Param.ParamValue = sSiteZip
'    moLoss.AddParameter Param
'
'    'Need to Get Max ID from Assignemnts Table
'    'Because of bad AutoNum
'    goUtil.utBadAutoNum , , lMaxID
'    Param.ParamName = "MAX_ID"
'    Param.ParamValue = lMaxID
'    moLoss.AddParameter Param
'
'    Exit Sub
'EH:
'    lErrNum = Err.Number
'    sErrDesc = Err.Description
'    Err.Raise lerrnum, , serrdesc & vbCrLf & clsCarLists_ClassName & vbCrLf & "Public Sub ADDParameters"
End Sub

Public Function GetAppDocsCol() As Collection
'    On Error GoTo EH
'    Dim sDocName As String
'    Dim MyAppDoc As V2ECKeyBoard.udtAppDoc
'    Dim lcount As Long
'    Dim lErrNum As Long
'    Dim sErrDesc As String


    
'
'    Set GetAppDocsCol = New Collection
'
'    '---------------->NOTE<-----------------------------
'    'Maximum of 14 App docs may be added to the collection
'    '---------------->NOTE<-----------------------------
'    'Put the Appended documents in the order they should be printed
'    'after each Loss report is printed
'
'    For lcount = 1 To 14
'        sDocName = Switch(lcount = 1, vbNullString, _
'                          lcount = 2, vbNullString, _
'                          lcount = 3, vbNullString, _
'                          lcount = 4, vbNullString, _
'                          lcount = 5, vbNullString, _
'                          lcount = 6, vbNullString, _
'                          lcount = 7, vbNullString, _
'                          lcount = 8, vbNullString, _
'                          lcount = 9, vbNullString, _
'                          lcount = 10, vbNullString, _
'                          lcount = 11, vbNullString, _
'                          lcount = 12, vbNullString, _
'                          lcount = 13, vbNullString, _
'                          lcount = 14, vbNullString)
'        If sDocName <> vbNullString Then
'            'All Appdocs will start off not selected to print
'            MyAppDoc.Selected = False
'            MyAppDoc.DocName = sDocName
'            GetAppDocsCol.Add MyAppDoc, sDocName
'        End If
'    Next
'
'    Exit Function
'EH:
'    lErrNum = Err.Number
'    sErrDesc = Err.Description
'    Err.Raise lerrnum, , serrdesc & vbCrLf & clsCarLists_ClassName & vbCrLf & "Public Function GetAppDocsCol"
End Function

Public Function AddLogEntries(psClaimNo As String, pdtActDate As Date, _
                              Optional pbAddPrefix As Boolean, _
                              Optional psPolicyDesc As String, _
                              Optional psBuildLim As String, _
                              Optional psContLim As String, _
                              Optional psDeductible As String) As Boolean

End Function

Public Sub ShowAllForms(oECTray As Object, psNavPos As String, Optional psSkipForms As String)
'    moGUI.ShowAllForms oECTray, psNavPos, psSkipForms
End Sub

Public Sub HideAllForms(Optional psSkipForms As String)
'    moGUI.HideAllForms psSkipForms
End Sub

Public Function SetUtilObject(oUtilObject As Object) As Boolean
'    On Error Resume Next
'    Set goUtil = oUtilObject
'    If Err.Number = 0 Then
'        SetUtilObject = True
'    Else
'        SetUtilObject = False
'    End If
End Function

Public Function CLEANUP() As Boolean
'    On Error GoTo EH
'    Dim lErrNum As Long
'    Dim sErrDesc As String

    
'
'    If Not moGUI Is Nothing Then
'        moGUI.CLEANUP
'        Set moGUI = Nothing
'    End If
'
'    Set gARV = Nothing
'    Set goUtil = Nothing
'
'    CLEANUP = True
'    Exit Function
'EH:
'    lErrNum = Err.Number
'    sErrDesc = Err.Description
'    Err.Raise lerrnum, , serrdesc & vbCrLf & clsCarLists_ClassName & vbCrLf & "Public Function CLEANUP"
End Function

