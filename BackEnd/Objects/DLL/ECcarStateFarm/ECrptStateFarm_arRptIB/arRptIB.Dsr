VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} arRptIB 
   Caption         =   "IB"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19606
   SectionData     =   "arRptIB.dsx":0000
End
Attribute VB_Name = "arRptIB"
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
    
    If IsNull(pvValue) Then
        pvValue = vbNullString
    End If
    If Not IsObject(pvValue) Then
        sValue = RTrim(CStr(pvValue))
        'Replace any carriage return flags with vbCrLf
        sValue = Replace(sValue, F_VBCRLF, vbCrLf)
    End If
    
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
        Case VbVarType.vbObject
            Set vNewValue = pvValue
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
    Dim sTemp As String
    Dim cTotalServiceFee As Currency
    Dim cTotalExpense As Currency
    Dim oField As Object
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim vMyFeeItem As Variant
    Dim MyFeeItem As udtFeeItems01
    Dim lServiceFeeCount As Long
    Dim lExpenseFeeCount As Long
    
    If GetProperty("fSubmitTo") = vbNullString Then
        fSubmitTo.Text = "State Farm"
    Else
        fSubmitTo.Text = GetProperty("fSubmitTo")
    End If
    
    If GetProperty("RemitTo") = vbNullString Then
        fRemitTo.Text = "Eberl's Temporary Services, Inc."
    Else
        fRemitTo.Text = GetProperty("RemitTo")
    End If
    
    If GetProperty("RemitAddress") = vbNullString Then
        sTemp = "7276 W. Mansfield Avenue" & vbCrLf
        sTemp = sTemp & "Lakewood, CO  80235" & vbCrLf
        sTemp = sTemp & "(303) 988-6286    Fax (303) 986-2771" & vbCrLf
        sTemp = sTemp & "Tax I.D. #84-1258811" & vbCrLf
        sTemp = sTemp & "Please return ""Remittance"" Copy to receive proper credit"
        fRemitAddress.Text = sTemp
    Else
        fRemitAddress.Text = GetProperty("RemitAddress")
    End If
    With Me
        .fIBNumber = GetProperty("fIBNumber")
        If GetProperty("f06_dtDateClosed") <> NULL_DATE Then
            .fDateClosed = Format(GetProperty("fDateClosed"), "mm/dd/yy")
        Else
            .fDateClosed = vbNullString
        End If
        .fAdjuster = GetProperty("fAdjuster")
        .fStateFarmID = GetProperty("fStateFarmID")
        
        '----------------Claim Information----------------------
        .fClaimNo = GetProperty("fClaimNo")
        'Supplement
        .fSupplement = GetProperty("fSupplement")
        If .fSupplement = "0" Then
            .fSupplement = vbNullString
            .chkSupplement1.Value = False
            .chkSupplement2.Value = False
        Else
            .chkSupplement1.Value = True
            .chkSupplement2.Value = True
        End If
        .fSupplementExplain = GetProperty("fSupplementExplain")
        .fAdditionalLoss = Format(GetProperty("fAdditionalLoss"), "$#,###,###,##0.00")
        
        .fMultiClaimBldgUnitNum = Trim(GetProperty("fMultiClaimBldgUnitNum"))
        If .fMultiClaimBldgUnitNum = vbNullString Then
            .chkMultiClaim.Value = False
        Else
            .chkMultiClaim.Value = True
        End If
        
        'Rebill
        .fRebilled = GetProperty("fRebilled")
        If .fRebilled = "0" Then
            .fRebilled = vbNullString
            .chkRebilled1.Value = False
            .chkRebilled1.Value = False
        Else
            .chkRebilled1.Value = True
            .chkRebilled1.Value = True
        End If
        .fOrigIBIBNumber = GetProperty("fOrigIBIBNumber")
        .fOrigIBTotalFee = Format(GetProperty("fOrigIBTotalFee"), "$#,###,###,##0.00")
        .fRebillExplain = GetProperty("fRebillExplain")
        
        .fPolicyNo = GetProperty("fPolicyNo")
        .fInsured = GetProperty("fInsured")
        .fLossLoc1 = GetProperty("fLossLoc1")
        .fLossLoc2 = GetProperty("fLossLoc2")
        .fLossLocCity = GetProperty("fLossLocCity")
        .fLossLocState = GetProperty("fLossLocState")
        .fLossLocZipcode = GetProperty("fLossLocZipcode")
        If GetProperty("fDateOfLoss") <> NULL_DATE Then
            .fDateOfLoss = Format(GetProperty("fDateOfLoss"), "mm/dd/yy")
        Else
            .fDateOfLoss = vbNullString
        End If
        .fGrossLoss = Format(GetProperty("fGrossLoss"), "$#,###,###,##0.00")
        
        
        '-----------------Fee Structure---------------------
        .fCatCode = GetProperty("fCatCode")
        .fSeverityCode = GetProperty("fSeverityCode")
        .fServiceFeeBase = Format(GetProperty("fServiceFeeBase"), "$#,###,###,##0.00")
        .fServiceFeeCovAExterior = Format(GetProperty("fServiceFeeCovAExterior"), "$#,###,###,##0.00")
        .fServiceFeeCovAFraming = Format(GetProperty("fServiceFeeCovAFraming"), "$#,###,###,##0.00")
        .fServiceFeeCovAInterior = Format(GetProperty("fServiceFeeCovAInterior"), "$#,###,###,##0.00")
        .fServiceFeeCovB = Format(GetProperty("fServiceFeeCovB"), "$#,###,###,##0.00")
        .fServiceFeeALE = Format(GetProperty("fServiceFeeALE"), "$#,###,###,##0.00")
        .fServiceFeeOutBuildings = Format(GetProperty("fServiceFeeOutBuildings"), "$#,###,###,##0.00")
        .fServiceFeeSteepCharge = Format(GetProperty("fServiceFeeSteepCharge"), "$#,###,###,##0.00")
        .fServiceFeeTwoStory = Format(GetProperty("fServiceFeeTwoStory"), "$#,###,###,##0.00")
        .fServiceFeeMoreThan50Squares = Format(GetProperty("fServiceFeeMoreThan50Squares"), "$#,###,###,##0.00")
        .fServiceFeeWoodSlateTileConRoof = Format(GetProperty("fServiceFeeWoodSlateTileConRoof"), "$#,###,###,##0.00")
        .fAddlFeesExplain = GetProperty("fAddlFeesExplain")
        .fServiceFeeAddl = Format(GetProperty("fServiceFeeAddl"), "$#,###,###,##0.00")
        .fServiceFeeTotal = Format(GetProperty("fServiceFeeTotal"), "$#,###,###,##0.00")
        .fPagerPhoneExpenseExplain = GetProperty("fPagerPhoneExpenseExplain")
        .fPagerPhoneExpense = Format(GetProperty("fPagerPhoneExpense"), "$#,###,###,##0.00")
        .fOtherExpenseExplain = GetProperty("fOtherExpenseExplain")
        .fOtherExpense = Format(GetProperty("fOtherExpense"), "$#,###,###,##0.00")
        .fSumTotalServiceFeeAndExpense = Format(GetProperty("fSumTotalServiceFeeAndExpense"), "$#,###,###,##0.00")
        .fTaxPercent = Format(GetProperty("fTaxPercent"), "0.000)")
        .fTaxesTotal = Format(GetProperty("fTaxesTotal"), "$#,###,###,##0.00")
        .fTotalFee = Format(GetProperty("fTotalFee"), "$#,###,###,##0.00")
        .fCopyName = GetProperty("fCopyName")
    End With
        
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
    mbActiveFlag = True
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
'        ReportFooter.Visible = True
    Else
        subChain.Visible = False
'        ReportFooter.Visible = False
    End If
End Sub

Public Function GetXMLExport() As String
    On Error GoTo EH

    'Export Report Collection Items
    Dim oMySer As WDDXSerializer        'Allaire's WDDX serializer
    Dim oMyRS As WDDXRecordset          'Allaire's WDDX Recordset
    Dim oMyStruct As WDDXStruct         'Allaire's WDDX Structure (Cold Fusion Strucuture type)
    Dim lCount As Long
    Dim sColName As String
    Dim vValue As Variant 'Use Variant to Get Variant Equiv Data Type.
    
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
        'Do not Add the Collection of Logs
        If StrComp(sColName, "coludtFeeItems01", vbTextCompare) <> 0 Then
            oMyRS.addColumn sColName
        End If
    Next
    
    'Only one row for the Data RS
    oMyRS.addRows 1
    'Set the Col values for this one row
    For lCount = 1 To mcolProperty.Count
        sColName = mcolPropertyKeys(lCount)
        'Use Variant Value to Get Data type
        If StrComp(sColName, "coludtFeeItems01", vbTextCompare) <> 0 Then
            vValue = mcolProperty(lCount)
            oMyRS.setField 1, sColName, vValue
        End If
    Next
    
    
    'Create WDDX Structure
    Set oMyStruct = New WDDXStruct
    
    oMyStruct.setProp "ClassName", ClassName
    oMyStruct.setProp "DataRS", oMyRS
    
    Set oMySer = New WDDXSerializer
    
    GetXMLExport = oMySer.serialize(oMyStruct)
    
    'Cleanup
    Set oMyRS = Nothing
    Set oMyStruct = Nothing
    Set oMySer = Nothing
    
    Exit Function
EH:
    GetXMLExport = "Class Name: " & ClassName & vbCrLf
    GetXMLExport = GetXMLExport & "Error # " & Err.Number & vbCrLf
    GetXMLExport = GetXMLExport & "Description: " & vbCrLf
    GetXMLExport = GetXMLExport & Err.Description
End Function

Private Sub ReportHeader_Format()

End Sub
