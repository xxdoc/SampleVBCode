VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} arRptIBFarmers 
   Caption         =   "IB"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19606
   SectionData     =   "arRptIBFarmers.dsx":0000
End
Attribute VB_Name = "arRptIBFarmers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Report Items
Private moLists As V2ECKeyBoard.clsCarLists
Private mcolProperty As Collection

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
    ClassName = App.EXEName & "." & "arRptIBFarmers"
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
    End Select
    
    If mcolProperty Is Nothing Then
        Set mcolProperty = New Collection
    End If
    RemoveProperty psName
    mcolProperty.Add vNewValue, psName
    
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
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    If GetProperty("f01_sSubToCarrier") = vbNullString Then
        f01SubmitTo.Text = "Farmers Insurance Company"
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
        .f10InsuredName = GetProperty("f10_sInsuredName")
        .f11LossLocation = GetProperty("f11_sLossLocation")
        If GetProperty("f12_dtDateOfLoss") <> NULL_DATE Then
            .f12DateOfLoss = Format(GetProperty("f12_dtDateOfLoss"), "mm/dd/yy")
        Else
            .f12DateOfLoss = vbNullString
        End If
        .f13GrossLoss = Format(GetProperty("f13_cGrossLoss"), "$##,##0.00")
        .f14Depreciation = Format(GetProperty("f14_cDepreciation"), "$##,##0.00")
        .f14aSupplement = GetProperty("f14a_sSupplement")
        .f14bRebilled = GetProperty("f14b_sRebilled")
        .f15Deductible = Format(GetProperty("f15_cDeductible"), "$##,##0.00")
        .f15aExcessLim = Format(GetProperty("f15a_cLessExcessLimits"), "$##,##0.00")
        .f15bExcessLimDesc = GetProperty("f15b_sExcessLimDesc")
        .f16NetClaim = Format(GetProperty("f16_cNetClaim"), "$##,##0.00")
        .f17ServiceFee = Format(GetProperty("f17_cServiceFee"), "$##,##0.00")
        .f17aMiscServiceFee = Format(GetProperty("f17a_cMiscServiceFee"), "$##,##0.00")
        .f18ServiceFeeComment = GetProperty("f18_sServiceFeeComment")
        .f18aMiscServiceFeeComment = GetProperty("f18a_sMiscServiceFeeComment")
        'Set outbuildings Amount
        If GetProperty("f19b_cOutBuildingsAmount") = 0 And GetProperty("f20_sOutBuildingsFeeComment") = vbNullString Then
            SetProperty "f19b_cOutBuildingsAmount", 20, vbCurrency
        End If
        If GetProperty("f19_cOutBuildingsFee") > 0 And GetProperty("f20_sOutBuildingsFeeComment") = vbNullString Then
            If GetProperty("f19a_iOutBuildingsCount") = 0 And GetProperty("f19b_cOutBuildingsAmount") > 0 Then
                If GetProperty("f19b_cOutBuildingsAmount") <= GetProperty("f19_cOutBuildingsFee") Then
                    SetProperty "f19a_iOutBuildingsCount", GetProperty("f19_cOutBuildingsFee") / GetProperty("f19b_cOutBuildingsAmount"), vbInteger
                End If
            End If
        End If
        .fOutBuildings = GetProperty("f19a_iOutBuildingsCount")
        .fOutBuildAmount = Format(GetProperty("f19b_cOutBuildingsAmount"), "$##,##0.00")
        .f19OutBuildingsFee = Format(GetProperty("f19_cOutBuildingsFee"), "$##,##0.00")
        .f20OutBuildingsComment = GetProperty("f20_sOutBuildingsFeeComment")
        .f21TwoStoryFee = Format(GetProperty("f21_cTwoStoryCharge"), "$##,##0.00")
        .f22SteepFee = Format(GetProperty("f22_cSteepCharge"), "$##,##0.00")
        .f23InteriorFee = Format(GetProperty("f23_cInteriorDamageCharge"), "$##,##0.00")
        .f24ExteriorFee = Format(GetProperty("f24_cExternalDamageBGCharge"), "$##,##0.00")
        'Add up all the service charges
        cTotalServiceFee = GetProperty("f17_cServiceFee") + GetProperty("f19_cOutBuildingsFee") + GetProperty("f21_cTwoStoryCharge") + GetProperty("f22_cSteepCharge") + GetProperty("f23_cInteriorDamageCharge") + GetProperty("f24_cExternalDamageBGCharge") + GetProperty("f17a_cMiscServiceFee")
        .f25ServiceFeeSubTotal = Format(cTotalServiceFee, "$##,##0.00")
        .f26Photos = GetProperty("f26_iPhotoCount")
        .f27PhotoFee = Format(GetProperty("f27_cPhotoFee"), "$##,##0.00")
        .f28Other = GetProperty("f28_iOther")
        .f29OtherFee = Format(GetProperty("f29_cOtherFee"), "$##,##0.00")
        .f29aMiscFeeComment = GetProperty("f29a_sMiscFeeComment")
        .f29bMiscFee = Format(GetProperty("f29b_cMiscFee"), "$##,##0.00")
        If GetProperty("f29_cOtherFee") > 0 And GetProperty("f28_iOther") > 0 Then
            .fotherAmount = Format(GetProperty("f29_cOtherFee") / GetProperty("f28_iOther"), "$##,##0.00")
        Else
            .fotherAmount = "0.00"
        End If
        cTotalExpense = GetProperty("f27_cPhotoFee") + GetProperty("f29_cOtherFee") + GetProperty("f29b_cMiscFee")
        .f30TotalExpenses = Format(cTotalExpense, "$##,##0.00")
        .f31TaxPercent = Format(GetProperty("f31_dTaxPercent"), "0.000")
        .f32Tax = Format(GetProperty("f32_cTaxAmount"), "$##,##0.00")
        .f33TotalAdjFee = Format(GetProperty("f33_cTotalAdjustingFee"), "$##,##0.00")
        If GetProperty("f33a_sAccountCode") = vbNullString Then
            .fAccountCode = "Account code 36 Sub Account 89"
        Else
            .fAccountCode = GetProperty("f33a_sAccountCode")
        End If
        'Indemnity Section
        'Hide it if the printOnIb flag is false
        If CBool(GetProperty("PrintOnIB")) Then
            Me.shapeHideIdemnity.Visible = False
            .f34Property = GetProperty("f34_sPaymentForProperty")
            .f35Auto = GetProperty("f35_sPaymentForAuto")
            .f36Final = GetProperty("f36_sPaymentForFinal")
            .f37Partial = GetProperty("f37_sPaymentForPartial")
            .f38AddSupYes = GetProperty("f38_sPaymentIsAddSupplementY")
            .f39AddSupNo = GetProperty("f39_sPaymentIsAddSupplementN")
            .f40PayAssYes = GetProperty("f40_sPaymentIsAssociatedY")
            .f41PayAssNo = GetProperty("f41_sPaymentISAssociatedN")
            .f42ClassOfClaim = GetProperty("f42_sClassOfClaim")
            .f43CauseOfLoss = GetProperty("f43_sCauseOfLoss")
            .f44TexasSubCovCode = GetProperty("f44_sTexasSubCovCode")
            .f45TexasSuffix = GetProperty("f45_sTexasSuffix")
            .f46TexasRoofDep = Format(GetProperty("f46_cTexasRoofDepreciation"), "$##,##0.00")
            .f47Building = GetProperty("f47_sTypeOfPropLossBuilding")
            .f48Contents = GetProperty("f48_sTypeOfPropLossContents")
            .f49ALE = GetProperty("f49_sTypeOfPropLossALE")
            .fOtherPropLoss = GetProperty("f49a_sOtherPropLoss")
            .f50InsuredPayee = GetProperty("f50_sInsuredPayeeName")
            .f51PayeeNames = GetProperty("f51_sPayeeNames")
            .f52Address = GetProperty("f52_sAddress")
            .f53AmountOfCheck = Format(GetProperty("f53_cAmountOfCheck"), "$##,##0.00")
            .f54CatCode = GetProperty("f54_sCatCode")
            .f55FieldHand = GetProperty("f55_sFieldHandled")
            .f56TotalLoss = GetProperty("f56_sTotalLoss")
            .f57CashinLieu = GetProperty("f57_sCashInLieu")
            .f58OwnRetSalvage = GetProperty("f58_sOwnerRetainSalvage")
            .f59Subrogation = GetProperty("f59_sSub")
            .f60Salvage = GetProperty("f60_sSalvage")
            .f61Instructions = GetProperty("f61_sInstructions")
            .f62ReqBy = GetProperty("f62_sRequestedBy")
            If GetProperty("f63_dtDate") <> NULL_DATE Then
                .f63ReqDate = Format(GetProperty("f63_dtDate"), "mm/dd/yy")
            Else
                .f63ReqDate = vbNullString
            End If
            .fApproveBy = GetProperty("f64_sApproveBy")
            If GetProperty("f65_dtApproveDate") <> NULL_DATE Then
                .fApproveDate = Format(GetProperty("f65_dtApproveDate"), "mm/dd/yy")
            Else
                .fApproveDate = vbNullString
            End If
            .fIssuedBy = GetProperty("f66_sIssuedBy")
            .fRetrievedBy = GetProperty("f67_sRetrievedBy")
        Else
            Me.shapeHideIdemnity.Visible = True
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
    
    If Not mcolProperty Is Nothing Then
        Set mcolProperty = Nothing
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

