VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} arRptIB02 
   Caption         =   "IB"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19606
   SectionData     =   "arRptIB02.dsx":0000
End
Attribute VB_Name = "arRptIB02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Report Items
Private moLists As V2ECKeyBoard.clsCarLists
Private mcolProperty As Collection
Private mcolPropertyKeys As Collection
Private msubAttachCheck As arsubAttachCheck 'Attatched CheckReq
Private msubFeeExplain As arsubFeeExplainer 'Fee Explainer Sub Report

'BGS 12.27.2000 Need this to see if the Report is still active.
Private mbActiveFlag As Boolean
'Chain Reports
'Private mbChainPgBrk As Boolean
Private mbChainFlag As Boolean
Private mlChainCount As Long
Private mcolChainReports As Collection ' Contains Reports Chained to it to be added to the Sub Report object
Private mcoludtFeeItems01 As Collection
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
    
    'Set the Fee Items Collection
    Set mcoludtFeeItems01 = mcolProperty("coludtFeeItems01")
    
    'Create Fee Explainer Sub Report
    Set msubFeeExplain = New arsubFeeExplainer
    Set msubFeeExplain.FeeItemsCol = mcoludtFeeItems01
    Set subFeeExplainer.object = msubFeeExplain.object
    
    f_Comments.Text = GetProperty("Comments")
    
    If GetProperty("f01_sSubToCarrier") = vbNullString Then
        f01SubmitTo.Text = "ABC Insurance Company"
    Else
        f01SubmitTo.Text = GetProperty("f01_sSubToCarrier")
    End If
    
    If GetProperty("RemitTo") = vbNullString Then
        fRemitTo.Text = "Eberls Temporary Services, Inc."
    Else
        fRemitTo.Text = GetProperty("RemitTo")
    End If
    
    If GetProperty("RemitAddress") = vbNullString Then
        sTemp = "7276 W. Mansfield Avenue" & vbCrLf
        sTemp = sTemp & "Lakewood, CO  80235" & vbCrLf
        sTemp = sTemp & "303.988.6286    Fax 303.986.2771" & vbCrLf
        sTemp = sTemp & "Tax I.D. #84-1258811" & vbCrLf
        sTemp = sTemp & "Please include Copy with Check to receive proper credit"
        fRemitAddress.Text = sTemp
    Else
        fRemitAddress.Text = GetProperty("RemitAddress")
    End If
    With Me
        .f02IB = GetProperty("f02_sIBNumber")
        .f04CatCd = GetProperty("f04_sCatCode")
        .f05City = GetProperty("f05_sLocation")
        .f05aState = GetProperty("f05a_sState")
        If GetProperty("f06_dtDateClosed") <> NULL_DATE Then
            .f06DateClosed = Format(GetProperty("f06_dtDateClosed"), "mm/dd/yy")
        Else
            .f06DateClosed = vbNullString
        End If
        .f07Adjuster = GetProperty("f07_sAdjusterName")
        .f08AdjCRID = GetProperty("f08_sAdjCRID")
        .f09SALN = GetProperty("f09_sSALN")
        .f09asPolicyNo = GetProperty("f09a_sPolicyNo")
        .f10InsuredName = GetProperty("f10_sInsuredName")
        .f11LossLocation = GetProperty("f11_sLossLocation")
        If GetProperty("f12_dtDateOfLoss") <> NULL_DATE Then
            .f12DateOfLoss = Format(GetProperty("f12_dtDateOfLoss"), "mm/dd/yy")
        Else
            .f12DateOfLoss = vbNullString
        End If
        .f13GrossLoss = Format(GetProperty("f13_cGrossLoss"), "$#,###,###,##0.00")
        .f14Depreciation = Format(GetProperty("f14_cDepreciation"), "$#,###,###,##0.00")
        .f14aSupplement = GetProperty("f14a_sSupplement")
        If .f14aSupplement = "0" Then
            .f14aSupplement = vbNullString
        End If
        .f14bRebilled = GetProperty("f14b_sRebilled")
        If .f14bRebilled = "0" Then
            .f14bRebilled = vbNullString
        End If
        .f15Deductible = Format(GetProperty("f15_cDeductible"), "$#,###,###,##0.00")
        .f15aExcessLim = Format(GetProperty("f15a_cLessExcessLimits"), "$#,###,###,##0.00")
        .f15bExcessLimDesc = GetProperty("f15b_sExcessLimDesc")
        .f15cMiscellaneous = Format(GetProperty("f15c_cLessMiscellaneous"), "$#,###,###,##0.00")
        .f15dMiscellaneousDesc = GetProperty("f15d_cMiscellaneousDesc")
        .f16NetClaim = Format(GetProperty("f16_cNetClaim"), "$#,###,###,##0.00")
        .f17ServiceFee = Format(GetProperty("f17_cServiceFee"), "$#,###,###,##0.00")
        msubFeeExplain.ServiceFee = Format(GetProperty("f17_cServiceFee"), "$#,###,###,##0.00")

        msubFeeExplain.MiscServiceFee = Format(GetProperty("f17a_cMiscServiceFee"), "$#,###,###,##0.00")
        .f18ServiceFeeComment = GetProperty("f18_sServiceFeeComment")
        msubFeeExplain.ServiceFeeComment = .f18ServiceFeeComment
        msubFeeExplain.MiscServiceFeeComment = GetProperty("f18a_sMiscServiceFeeComment")
        cTotalServiceFee = GetProperty("f25_cServiceFeeSubTotal")
'        .f25ServiceFeeSubTotal = Format(cTotalServiceFee, "$#,###,###,##0.00")
        msubFeeExplain.MiscExpenseFeeComment = GetProperty("f29a_sMiscFeeComment")
        msubFeeExplain.MiscExpenseFee = Format(GetProperty("f29b_cMiscFee"), "$#,###,###,##0.00")
        cTotalExpense = GetProperty("f30_cTotalExpenses")
'        .f30TotalExpenses = Format(cTotalExpense, "$#,###,###,##0.00")
'        .f_TotalServiceAndExpense = Format(cTotalServiceFee + cTotalExpense, "$#,###,###,##0.00")
        msubFeeExplain.TaxPercent = GetProperty("f31_dTaxPercent")
        msubFeeExplain.TaxAmount = GetProperty("f32_cTaxAmount")
        msubFeeExplain.TotalAdjustingFee = GetProperty("f33_cTotalAdjustingFee")
        If GetProperty("f_p001_AccountCode") = vbNullString Then
            msubFeeExplain.AccountCode = "Account code 36 Sub Account 89"
        Else
            msubFeeExplain.AccountCode = GetProperty("f_p001_AccountCode")
        End If
        'Indemnity Section
        'Hide it if the printOnIb flag is false
        If CBool(GetProperty("bPrintOnIB")) Then
            Set msubAttachCheck = New arsubAttachCheck
            Set msubAttachCheck.PropertyCol = mcolProperty
            Set subAttachCheck.object = msubAttachCheck.object
            subAttachCheck.Visible = True
        Else
            subAttachCheck.Visible = False
        End If
        .fCopyName = GetProperty("CopyName")
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
    
'    If Not mcolProperty Is Nothing Then
'        Set mcolProperty = Nothing
'    End If
    Set subFeeExplainer.object = Nothing
    Set msubFeeExplain = Nothing
    Set subAttachCheck.object = Nothing
    Set msubAttachCheck = Nothing
    
    
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
    Set mcoludtFeeItems01 = Nothing
    
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
    'IB needs to include Collection of Fee Items
    Dim MyFeeItem  As udtFeeItems01
    Dim oFeeItemRS As WDDXRecordset
    
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

    '****BEGIN***IB needs to include Collection of Fee Items***
    Set mcoludtFeeItems01 = mcolProperty("coludtFeeItems01")
    If Not mcoludtFeeItems01 Is Nothing Then
        If mcoludtFeeItems01.Count > 0 Then
            Set oFeeItemRS = New WDDXRecordset
            'Add Colmn names
            oFeeItemRS.addColumn "Amount"
            oFeeItemRS.addColumn "Comment"
            oFeeItemRS.addColumn "fsftDescription"
            oFeeItemRS.addColumn "fsftFeeAmount"
            oFeeItemRS.addColumn "fsftIsExpense"
            oFeeItemRS.addColumn "fsftIsMiscAmount"
            oFeeItemRS.addColumn "fsftMaxFeeAmount"
            oFeeItemRS.addColumn "fsftMaxNumberOfItems"
            oFeeItemRS.addColumn "fsftName"
            oFeeItemRS.addColumn "fsftTypeNum"
            oFeeItemRS.addColumn "fsftUseFormula"
            oFeeItemRS.addColumn "fsftVBFormula"
            oFeeItemRS.addColumn "NumberOfItems"
            
            'Add the same number in collection
            oFeeItemRS.addRows mcoludtFeeItems01.Count
        End If
    
    End If
         
    For lCount = 1 To mcoludtFeeItems01.Count
        MyFeeItem = mcoludtFeeItems01(lCount)
        With MyFeeItem
            vValue = .Amount
            oFeeItemRS.setField lCount, "Amount", vValue
            vValue = .Comment
            oFeeItemRS.setField lCount, "Comment", vValue
            vValue = .fsftDescription
            oFeeItemRS.setField lCount, "fsftDescription", vValue
            vValue = .fsftFeeAmount
            oFeeItemRS.setField lCount, "fsftFeeAmount", vValue
            vValue = .fsftIsExpense
            oFeeItemRS.setField lCount, "fsftIsExpense", vValue
            vValue = .fsftIsMiscAmount
            oFeeItemRS.setField lCount, "fsftIsMiscAmount", vValue
            vValue = .fsftMaxFeeAmount
            oFeeItemRS.setField lCount, "fsftMaxFeeAmount", vValue
            vValue = .fsftMaxNumberOfItems
            oFeeItemRS.setField lCount, "fsftMaxNumberOfItems", vValue
            vValue = .fsftName
            oFeeItemRS.setField lCount, "fsftName", vValue
            vValue = .fsftTypeNum
            oFeeItemRS.setField lCount, "fsftTypeNum", vValue
            vValue = .fsftUseFormula
            oFeeItemRS.setField lCount, "fsftUseFormula", vValue
            vValue = .fsftVBFormula
            oFeeItemRS.setField lCount, "fsftVBFormula", vValue
            vValue = .NumberOfItems
            oFeeItemRS.setField lCount, "NumberOfItems", vValue
        End With
    Next
    '****END***IB needs to include Collection of Fee Items***
    
    'Create WDDX Structure
    Set oMyStruct = New WDDXStruct
    
    oMyStruct.setProp "ClassName", ClassName
    oMyStruct.setProp "DataRS", oMyRS
    If Not oFeeItemRS Is Nothing Then
        oMyStruct.setProp "FeeItemRS", oFeeItemRS
    End If
    
    Set oMySer = New WDDXSerializer
    
    GetXMLExport = oMySer.serialize(oMyStruct)
    
    'Cleanup
    Set oMyRS = Nothing
    Set oFeeItemRS = Nothing
    Set oMyStruct = Nothing
    Set oMySer = Nothing
    
    Exit Function
EH:
    GetXMLExport = "Class Name: " & ClassName & vbCrLf
    GetXMLExport = GetXMLExport & "Error # " & Err.Number & vbCrLf
    GetXMLExport = GetXMLExport & "Description: " & vbCrLf
    GetXMLExport = GetXMLExport & Err.Description
End Function

