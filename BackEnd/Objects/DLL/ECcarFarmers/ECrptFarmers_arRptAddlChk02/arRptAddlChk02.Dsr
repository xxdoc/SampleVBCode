VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} arRptAddlChk02 
   Caption         =   "AddlCheck"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19606
   SectionData     =   "arRptAddlChk02.dsx":0000
End
Attribute VB_Name = "arRptAddlChk02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Report Items
Private moLists As V2ECKeyBoard.clsCarLists
Private mcolProperty As Collection
Private mcolPropertyKeys As Collection

'BGS 12.27.2000 Need this to see if the Report is still active.
Private mbActiveFlag As Boolean
'Chain Reports
'Private mbChainPgBrk As Boolean
Private mbChainFlag As Boolean
Private mlChainCount As Long
Private mcolChainReports As Collection ' Contains Reports Chained to it to be added to the Sub Report object
Private moChainReport As Object

Public Property Let ChainReport(poActReport As Object)
    Set moChainReport = poActReport
End Property
Public Property Set ChainReport(poActReport As Object)
    Set moChainReport = poActReport
End Property
Public Property Get ChainReport() As Object
    Set ChainReport = moChainReport
End Property

Public Property Let Lists(poLists As V2ECKeyBoard.clsCarLists)
    Set moLists = poLists
End Property
Public Property Set Lists(poLists As V2ECKeyBoard.clsCarLists)
    Set moLists = poLists
End Property
Public Property Get Lists() As V2ECKeyBoard.clsCarLists
    Set Lists = moLists
End Property

Public Property Get ClassName() As String
    ClassName = App.EXEName & "." & Me.Name
End Property

Public Property Get ActiveFlag() As Boolean
    ActiveFlag = mbActiveFlag
End Property
Public Property Let ActiveFlag(pbFlag As Boolean)
    mbActiveFlag = pbFlag
End Property

Public Sub SetProperty(psName As String, pvValue As Variant, pType As VbVarType)
    On Error GoTo EH
    Dim sValue As String
    Dim vNewValue As Variant
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    sValue = RTrim(CStr(pvValue))
    'Replace any carriage return flags with vbCrLf
    sValue = Replace(sValue, F_VBCRLF, vbCrLf)
    
    Select Case pType
        Case VbVarType.vbDate
            If IsDate(sValue) Then
                vNewValue = CDate(sValue)
            Else
                vNewValue = CDate(NULL_DATE)
            End If
        Case VbVarType.vbCurrency
            vNewValue = CCur(sValue)
        Case VbVarType.vbString
            vNewValue = CStr(sValue)
        Case VbVarType.vbInteger
            vNewValue = CInt(sValue)
        Case VbVarType.vbBoolean
            vNewValue = CBool(sValue)
        Case VbVarType.vbLong
            vNewValue = CLng(sValue)
        Case VbVarType.vbDouble
            vNewValue = CDbl(sValue)
        Case VbVarType.vbSingle
            vNewValue = CSng(sValue)
        Case VbVarType.vbUserDefinedType
            vNewValue = CStr(sValue)
    End Select
    
    If mcolProperty Is Nothing Then
        Set mcolProperty = New Collection
        Set mcolPropertyKeys = New Collection
    End If
    RemoveProperty psName
    mcolProperty.Add vNewValue, psName
    mcolPropertyKeys.Add psName, psName
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Public Sub SetProperty"
End Sub

Public Function GetProperty(psName As String) As Variant
    On Error GoTo EH
    
    If Not mcolProperty Is Nothing Then
        GetProperty = mcolProperty(psName)
    Else
        GetProperty = psName & ": Property collection not set!"
    End If
    
    Exit Function
EH:
    GetProperty = vbNullString
    Err.Clear
End Function

Public Function RemoveProperty(psName As String) As Boolean
    On Error GoTo EH
    If Not mcolProperty Is Nothing Then
        mcolProperty.Remove psName
        mcolPropertyKeys.Remove psName
        RemoveProperty = True
    End If
    Exit Function
EH:
    Err.Clear
End Function

Public Function ExportME(psXportPath As String, pXportType As ExportType) As Boolean
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    ExportME = Lists.ExportARReport(Me, psXportPath, pXportType)
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Public Function ExportME"
End Function

Private Sub ActiveReport_ReportStart()
    On Error GoTo EH
    Dim oField As Object
    Dim sTemp As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    For Each oField In Me.Detail.Controls
        If Mid(CStr(oField.Name), 2, 1) = "_" Then
            sTemp = GetProperty(CStr(oField.Name))
            Select Case UCase(sTemp)
                Case "YES", "TRUE", "-1"
                    sTemp = "X"
                Case "NO", "FALSE", "0"
                    sTemp = vbNullString
            End Select
            Select Case oField.Name
                Case f_p049_dtDate.Name, f_p051_dtApproveDate.Name
                    If sTemp = NULL_DATE Then
                        sTemp = vbNullString
                    End If
                Case f_RT53_cAmountOfCheck.Name, f_p046_cTexasRoofDepreciation.Name
                    sTemp = Format(sTemp, "$##,##0.00")
                Case f_p016_SocialSecNum1.Name
                    sTemp = Format(sTemp, "000-00-0000")
            End Select
            oField.Text = sTemp
        Else
            If GetProperty(CStr(oField.Name)) <> vbNullString Then
                oField.Text = GetProperty(CStr(oField.Name))
            End If
        End If
    Next
    
    'Set the Chain flag if we have any
    If Not mcolChainReports Is Nothing Then
        If Not mbChainFlag Then
            mbChainFlag = True
            mlChainCount = 1
        End If
    Else
        mbChainFlag = False
    End If
    
    'If we have Chained Reports...
    If mbChainFlag Then
        Set moChainReport = mcolChainReports(mlChainCount)
        'Start the daisy linking here
        SetNextChainReport mlChainCount, mcolChainReports
        'Set the ref to sub reports in this Report
        Set subChain.object = moChainReport
    Else
        If Not moChainReport Is Nothing Then
            'Set the ref to sub reports in this Report
            Set subChain.object = moChainReport
        End If
        
    End If
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Private Sub ActiveReport_ReportStart"
End Sub

Private Sub ActiveReport_ReportEnd()
    On Error Resume Next
    Dim oAR As Object
    'Clean up chain reports collection and objects
    If Not mcolChainReports Is Nothing Then
        For Each oAR In mcolChainReports
            Unload oAR
            Set oAR = Nothing
        Next
        Set mcolChainReports = Nothing
        Unload moChainReport
        Set moChainReport = Nothing
    End If
    
'    If Not mcolProperty Is Nothing Then
'        Set mcolProperty = Nothing
'    End If
    
End Sub

'For Chained Reports
Public Sub AddChainReport(poActiveReport As Object)
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    If mcolChainReports Is Nothing Then
        Set mcolChainReports = New Collection
    End If
    
    mcolChainReports.Add poActiveReport
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Public Sub AddChainReport"
End Sub
'For Chained Reports
Public Sub SetNextChainReport(plChainCount As Long, pcolChainReports As Collection)
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    If plChainCount + 1 <= pcolChainReports.Count Then
        Set pcolChainReports(plChainCount).ChainReport = pcolChainReports(plChainCount + 1)
        plChainCount = plChainCount + 1
        'Do daisy again
        pcolChainReports(plChainCount - 1).SetNextChainReport plChainCount, pcolChainReports
    End If
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Public Sub SetNextChainReport"
End Sub

Private Sub ActiveReport_Terminate()
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    Set moLists = Nothing
    Set mcolProperty = Nothing
    Set mcolPropertyKeys = Nothing
    Set mcolChainReports = Nothing
    Set moChainReport = Nothing
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Private Sub ActiveReport_Terminate"
End Sub

Private Sub Detail_Format()
CHAINED_REPORTS:
    If Not moChainReport Is Nothing Then
        subChain.Visible = True
        ReportFooter.Visible = True
    Else
        subChain.Visible = False
        ReportFooter.Visible = False
    End If
End Sub

Public Function GetXMLExport() As String
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    'Export Report Collection Items
    Dim oMySer As WDDXSerializer        'Allaire's WDDX serializer
    Dim oMyRS As WDDXRecordset          'Allaire's WDDX Recordset
    Dim oMyStruct As WDDXStruct         'Allaire's WDDX Structure (Cold Fusion Strucuture type)
    Dim oIndemnityPaymentRS As WDDXRecordset
    Dim sTemp As String
    Dim sPayeeLine3 As String
    Dim sPayeeLine4 As String
    Dim lCount As Long
    Dim sColName As String
    Dim vValue As Variant 'Use Variant to Get Variant Equiv Data Type.
    Dim sAddress As String
    Dim sStreet As String
    Dim sCity As String
    Dim sState As String
    Dim sZip As String
    
    'Make sure the Collection items exist
    If mcolProperty Is Nothing Or mcolPropertyKeys Is Nothing Then
        Exit Function
    End If
    If mcolProperty.Count = 0 Or mcolPropertyKeys.Count = 0 Then
        Exit Function
    End If
    
    'Create a WDDX RS
    Set oMyRS = New WDDXRecordset
    For lCount = 1 To mcolPropertyKeys.Count
        'Use the Keys Collection to create Column Names
        sColName = mcolPropertyKeys(lCount)
        oMyRS.addColumn sColName
    Next
    
    'Only one row for the Data RS
    oMyRS.addRows 1
    'Set the Col values for this one row
    For lCount = 1 To mcolProperty.Count
        sColName = mcolPropertyKeys(lCount)
        'Use Variant Value to Get Data type
        vValue = mcolProperty(lCount)
        oMyRS.setField 1, sColName, vValue
    Next
    
    '6.15.2005 BGS Need to Add IndemnityPaymentRS to the Wddx so
    'It can be harvested when sending Documents to Farmers
'    <var name="IndemnityPaymentRS">
'- <recordset rowCount="1" fieldNames="PaymentGUID,ContactPayeeId,PayeeLineOne,PayeeLineTwo,PayeeLineThree,PayeeLineFour,PaymentAmount,CorrespondenceRequired,Repairable,TotalLoss,TexasRoofDepreciation,TexasCoverageCode,CashInLieu,OwnerRetained">
'- <var name="DocumentPropertiesRS">
'- <recordset rowCount="1" fieldNames="UniqueID,UnitNumber,SequenceNumber,TotalDocs,GUID,SubType,Description,State">
    Set oIndemnityPaymentRS = New WDDXRecordset
    oIndemnityPaymentRS.addColumn "PaymentGUID"
    oIndemnityPaymentRS.addColumn "ContactPayeeId"
    oIndemnityPaymentRS.addColumn "PayeeLineOne"
    oIndemnityPaymentRS.addColumn "PayeeLineTwo"
    oIndemnityPaymentRS.addColumn "PayeeLineThree"
    oIndemnityPaymentRS.addColumn "PayeeLineFour"
    oIndemnityPaymentRS.addColumn "PaymentAmount"
    oIndemnityPaymentRS.addColumn "CorrespondenceRequired"
    oIndemnityPaymentRS.addColumn "Repairable"
    oIndemnityPaymentRS.addColumn "TotalLoss"
    oIndemnityPaymentRS.addColumn "TexasRoofDepreciation"
    oIndemnityPaymentRS.addColumn "TexasCoverageCode"
    oIndemnityPaymentRS.addColumn "CashInLieu"
    oIndemnityPaymentRS.addColumn "OwnerRetained"
    oIndemnityPaymentRS.addRows 1
    oIndemnityPaymentRS.setField 1, "PaymentGUID", "[ENTER_PaymentGUID_PackageItemGUID]"
    sTemp = oMyRS.getField(1, "f_p057_CRNVar_ContactPayeeId")
    sTemp = Mid(sTemp, InStrRev(sTemp, "_", , vbBinaryCompare) + 1)
    oIndemnityPaymentRS.setField 1, "ContactPayeeId", CleanXML(sTemp)
    sTemp = oMyRS.getField(1, "f_RT51_sPayeeNames")
    oIndemnityPaymentRS.setField 1, "PayeeLineOne", CleanXML(sTemp)
    sTemp = oMyRS.getField(1, "f_RT50_sInsuredPayeeName")
    oIndemnityPaymentRS.setField 1, "PayeeLineTwo", CleanXML(sTemp)
    sAddress = oMyRS.getField(1, "f_RT52_sAddress")
    goUtil.utFillAddressFields sAddress, sZip, sState, sCity, sStreet
    If sStreet <> vbNullString Then
        oIndemnityPaymentRS.setField 1, "PayeeLineThree", CleanXML(sStreet)
    Else
        oIndemnityPaymentRS.setField 1, "PayeeLineThree", "[ENTER_PayeeLineThree_MAStreet]"
    End If
    If sCity <> vbNullString And sState <> vbNullString And sZip <> vbNullString Then
        sTemp = sCity & ", " & sState & ", " & sZip
        oIndemnityPaymentRS.setField 1, "PayeeLineFour", CleanXML(sTemp)
    Else
        oIndemnityPaymentRS.setField 1, "PayeeLineFour", "[ENTER_PayeeLineFour_MACity,MAState,MAZip]"
    End If
    sTemp = oMyRS.getField(1, "f_RT53_cAmountOfCheck")
    sTemp = Format(sTemp, "0.00")
    oIndemnityPaymentRS.setField 1, "PaymentAmount", CleanXML(sTemp)
    sTemp = oMyRS.getField(1, "f_p00a_CorrespondenceRequiredYN")
    If CBool(sTemp) Then
        sTemp = "Y"
    Else
        sTemp = "N"
    End If
    oIndemnityPaymentRS.setField 1, "CorrespondenceRequired", CleanXML(sTemp)
    sTemp = oMyRS.getField(1, "f_p040_sTotalLoss")
    If CBool(sTemp) Then
        sTemp = "N"
    Else
        sTemp = "Y"
    End If
    oIndemnityPaymentRS.setField 1, "Repairable", CleanXML(sTemp)
    If sTemp = "Y" Then
        sTemp = "N"
    Else
        sTemp = "Y"
    End If
    oIndemnityPaymentRS.setField 1, "TotalLoss", CleanXML(sTemp)
    sTemp = oMyRS.getField(1, "f_p046_cTexasRoofDepreciation")
    sTemp = Format(sTemp, "0.00")
    oIndemnityPaymentRS.setField 1, "TexasRoofDepreciation", CleanXML(sTemp)
    sTemp = oMyRS.getField(1, "f_p0000_sTexasSubCovCode")
    oIndemnityPaymentRS.setField 1, "TexasCoverageCode", CleanXML(sTemp)
    sTemp = oMyRS.getField(1, "f_p041_sCashInLieu")
    If CBool(sTemp) Then
        sTemp = "Y"
    Else
        sTemp = "N"
    End If
    oIndemnityPaymentRS.setField 1, "CashInLieu", CleanXML(sTemp)
    sTemp = oMyRS.getField(1, "f_p042_sOwnerRetainSalvage")
    If CBool(sTemp) Then
        sTemp = "Y"
    Else
        sTemp = "N"
    End If
    oIndemnityPaymentRS.setField 1, "OwnerRetained", CleanXML(sTemp)
    
    'Create WDDX Structure
    Set oMyStruct = New WDDXStruct
    
    oMyStruct.setProp "ClassName", ClassName
    oMyStruct.setProp "DataRS", oMyRS
    oMyStruct.setProp "IndemnityPaymentRS", oIndemnityPaymentRS
    Set oMySer = New WDDXSerializer
    
    GetXMLExport = oMySer.serialize(oMyStruct)
    
    'Cleanup
    Set oIndemnityPaymentRS = Nothing
    Set oMyRS = Nothing
    Set oMyStruct = Nothing
    Set oMySer = Nothing
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Private Sub ActiveReport_Terminate"
End Function

Private Function CleanXML(psXML As String) As String
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim sXML As String
    Dim oSer As WDDXSerializer
    
    Set oSer = New WDDXSerializer
    sXML = psXML
    'First get rid of Yucky chars
    sXML = Replace(sXML, Chr(160), Chr(32), , , vbBinaryCompare)
    sXML = Replace(sXML, vbCrLf, Chr(32), , , vbBinaryCompare)
    sXML = oSer.serialize(sXML)
   
    sXML = Mid(sXML, InStr(1, sXML, "<string>", vbBinaryCompare) + 8)
    sXML = Left(sXML, InStr(1, sXML, "</string>", vbBinaryCompare) - 1)
    CleanXML = sXML
    
    Set oSer = Nothing
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Set oSer = Nothing
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Private Function CleanXML"
End Function
