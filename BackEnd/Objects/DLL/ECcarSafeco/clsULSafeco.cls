VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsULSafeco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Carrier Upload Polymorphism
'Carrier Upload InterFace Definiton
'Use Implements V2ECKeyBoard.clsCarUL in ALL Upload Classes For All Carrier Objects

Implements V2ECKeyBoard.clsCarUL
'--------------------------------------------------------------------------------------
'If fields need to be added to this class must update the following:
'1. The back end DB must have the Fields duh
'2. udtIB_Safeco (Try to use the Same convention)
'3. Update Private Function PopulateUDT
'4. Update Private Function BuildClaimsSQLAB
'5. Update Public Function PopulateFields
'
'Each function that needs to be updated should be Easy, no imbeded SQL
'Each Field Item will be another Line of code. Thats why using the NumAlpha convention
'is good for maintenance, ie future additions. As well using the Num Alpha allows for
'Databases that will not allow long field names
'--------------------------------------------------------------------------------------


Private Type udtIB_Safeco
    IB00_dtContactedDate As Date            'Date contacted (Not on IB Report)
    IB00_lSSN As String                      'ADJ SSN (Not on IB Report)
    IB01_sSubToCarrier As String            'Safeco
    IB02_sIBNumber As String                '{Checks.ClaimNo} + " " + {@Rebilled} + {@Supplement}
    IB03_sCheckNUM As String                'Checks.CheckNUM
    IB04_sCatCode As String                 'Ass.CatCode
    IB05_sLocation As String                'Ass.SiteCity
    IB05a_sState As String                  'Ass.SiteState
    IB06_dtDateClosed As Date                '{Billing.ClosedDate}[5 to 6] + "/" + {Billing.ClosedDate}[7 to 8] + "/" + {Billing.ClosedDate}[1 to 4]
    IB07_sAdjusterName As String            'Trim({Billing.AdjustorFirstName}) + " " + trim({Billing.AdjustorLastName})
    IB08_sAdjCRID As String                 'ass.claimrepidno
    IB09_sSALN As String                     'ass.clientclaimno
    IB09a_sPolicyNumber As String
    IB10_sInsuredName As String              'ass.newinsurednames
    IB11_sLossLocation  As String            'ass.NewLoassLocation
    IB12_dtDateOfLoss As Date                 '{Billing.LossDate}[5 to 6] + "/" + {Billing.LossDate}[7 to 8] + "/" + {Billing.LossDate}[3 to 4]
    IB13_cGrossLoss As Currency              'ass.FullCostOfRepair
    IB14_cDepreciation As Currency           '{Billing.RecoverableDepreciation}+{Billing.NonRecovDepr}
    IB14a_sSupplement As String              'if {Checks.Supplement} >0 then "X" else ""
    IB14b_sRebilled As String                'if {Checks.Rebilled} = true then "X" else ""
    IB15_cDeductible As Currency             'ass.Deductible
    IB15a_cLessExcessLimits As Currency      'ass.LessExcessLimits
    IB15b_sExcessLimDesc As String           'Checks.ExcessLimDesc
    IB16_cNetClaim As Currency               'if {@TotalNetActualCashValueClaim} <= 0 then 0 else {@TotalNetActualCashValueClaim}
    IB17_cServiceFee As Currency             'ass.ServiceFee
    IB17a_cMiscServiceFee As Currency        'Ass.MiscServiceFee
    IB18_sServiceFeeComment As String        'if {Billing.BillingHours} <> "" then "(" + {Billing.BillingHours} + " hours @ $75)" else ""
    IB18a_sMiscServiceFeeComment As String   'ass.MiscServiceFeeComment
    IB19_cOutBuildingsFee As Currency        'ass.OutBuildingsFee
    IB19a_iOutBuildingsCount As Integer      'ass.Outbuildings
    IB19b_cOutBuildingsAmount As Currency    'Reg Setting Amount per outbuilding
    IB20_sOutBuildingsFeeComment As String   'if {Billing.BillingFenceOnly} = "1" then "(Fence Only)" else ""
    IB21_cTwoStoryCharge As Currency         'ass.TwoStoryDamage
    IB22_cSteepCharge As Currency            'ass.SteepCharge
    IB23_cInteriorDamageCharge As Currency   'ass.BillingInteriorDamageDollars
    IB24_cExternalDamageBGCharge As Currency 'ass.BillingExteriorDamageDollars
    IB25_cServiceFeeSubTotal As Currency     '{Billing.ServiceFee}+{Billing.OutBuildingsFee}+{Billing.TwoStoryDamage}+{Billing.SteepCharge}+{Billing.BillingInteriorDamageDollars}+{Billing.BillingExteriorDamageDollars}
    IB26_iPhotoCount As Integer              'ass.photographCount
    IB27_cPhotoFee As Currency               'ass.photographfee
    IB28_iOther As Integer                   'ass.miles
    IB29_cOtherFee As Currency               'ass.MilesFee
    IB29a_sMiscExpenseFeeComment As String   'ass.MiscFeesComment
    IB29b_cMiscExpenseFee As Currency        'ass.MiscFees
    IB30_cTotalExpenses As Currency          '{Billing.PhotographFee}+{Billing.MilesFee}
    IB31_dTaxPercent As Double               'ass.TaxPercent
    IB32_cTaxAmount As Currency              'ass.TaxPercentAmount
    IB33_cTotalAdjustingFee As Currency      'ass.InvoiceTotal
    IB33a_sAccountCode As String             'Reg setting
    IB34_sPaymentForProperty As String      'if {Checks.ForProperty} = True then "X" else ""
    IB35_sPaymentForAuto As String          'if {Checks.ForAuto} = true then "X" else ""
    IB36_sPaymentForFinal As String         'if {Checks.FinalPayment} = true then "X" else ""
    IB37_sPaymentForPartial As String      'if {Checks.PartialPayment} = true then "X" else ""
    IB38_sPaymentIsAddSupplementY As String 'if {Checks.ForAdditionalSupplement} = true then "Y" else ""
    IB39_sPaymentIsAddSupplementN As String 'if {Checks.ForAdditionalSupplement} = false then "X" else ""
    IB40_sPaymentIsAssociatedY As String    'if {Checks.PaymentAssociated} = true then "X" else ""
    IB41_sPaymentISAssociatedN As String   'if {Checks.PaymentAssociated}= False then "X" else ""
    IB42_sClassOfClaim As String             '{Checks.TypeOfLoss}[1 to 2]
    IB43_sCauseOfLoss As String              '{Checks.CauseOfLoss}[1 to 2]
    IB44_sTexasSubCovCode As String          'Checks.texasSubCode
    IB45_sTexasSuffix As String              'Checks.texasSuffix
    IB46_cTexasRoofDepreciation As Currency 'Checks.TexasRoofDepreciation
    IB47_sTypeOfPropLossBuilding As String  'if {Checks.TypeOfLoss}[1 to 2]  = "01" then "X" else ""
    IB48_sTypeOfPropLossContents As String  'if {Checks.TypeOfLoss}[1 to 2]  = "02" then "X" else ""
    IB49_sTypeOfPropLossALE As String       'if {Checks.TypeOfLoss}[1 to 2]  = "03" then "X" else ""
    IB49a_sOtherPropLoss As String          'Used in future vs 2.0
    IB50_sInsuredPayeeName As String         'if isnull({Billing.NewSecondInsuredNames}) or trim({Billing.NewSecondInsuredNames}) = "" then {Billing.NewInsuredNames} else trim({Billing.NewInsuredNames}) + " and " + {Billing.NewSecondInsuredNames}
    IB51_sPayeeNames As String               'if {Checks.IncludeMorgOnDraft}= TRUE then {Billing.MortgageCompany} else ""
    IB52_sAddress As String                  'ass.NewMailing Address
    IB53_cAmountOfCheck As Currency          'Checks.AmountOfCheck
    IB54_sCatCode As String                  'ass.CatCode
    IB55_sFieldHandled As String            'if{Checks.FieldHandled}= true then "Yes" else "No"
    IB56_sTotalLoss As String              'if {Checks.TotalLoss} = true then "YES" else "NO"
    IB57_sCashInLieu As String              'if {Checks.CashInLieu}= true then "Yes" else "No"
    IB58_sOwnerRetainSalvage As String      'if{Checks.OwnerRetainedSalvage}= true then "Yes" else "No"
    IB59_sSub As String                     'if {Checks.Subrogation} = true then "Yes" else "No"
    IB60_sSalvage As String                 'if {Checks.Salvage} = true then "Yes" else "No"
    IB61_sInstructions As String             'Checks.Instructions
    IB62_sRequestedBy As String              'Trim({Billing.AdjustorFirstName}) + " " + trim({Billing.AdjustorLastName})
    IB63_dtDate As Date                       'Have to think about this one. Right now this is figured whenever the Adjuster Prints the Report
    IB64_sApproveBy As String               'Used in future vs 2.0
    IB65_dtApproveDate As Date              'Used in future vs 2.0
    IB66_sIssuedBy As String                'Used in future vs 2.0
    IB67_sRetrievedBy As String             'Used in future vs 2.0
End Type

Private msClassName As String
Private msIBTableName As String
Private msIBNumber As String
Private mudtIBSafeco As udtIB_Safeco
Private mbDoDBUpdate As Boolean

Private Function clsCarUL_SetUtilObject(oUtilObject As Object) As Boolean

    On Error Resume Next
    Set goUtil = oUtilObject
    If Err.Number = 0 Then
        clsCarUL_SetUtilObject = True
    Else
        clsCarUL_SetUtilObject = False
    End If

End Function

Public Property Get clsCarUL_ULType() As V2ECKeyBoard.UploadType
    clsCarUL_ULType = ClosedIB
End Property

Public Property Let clsCarUL_DoDBUpdate(pbFlag As Boolean)
    mbDoDBUpdate = pbFlag
End Property
Public Property Get clsCarUL_DoDBUpdate() As Boolean
    clsCarUL_DoDBUpdate = mbDoDBUpdate
End Property

Public Property Get clsCarUL_ClassName() As String
    clsCarUL_ClassName = App.EXEName & "." & msClassName
End Property

Public Property Get clsCarUL_TableName() As String
    clsCarUL_TableName = msIBTableName
End Property

Public Property Get clsCarUL_IBNumber() As String
    clsCarUL_IBNumber = msIBNumber
End Property

Public Function clsCarUL_PopulateFields(psClaimNo As String, psCheckNum As String, pCurDB As Object, pclsUL As V2ECKeyBoard.clsUpload) As Boolean
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    clsCarUL_PopulateFields = True
    
    'BGS 10.31.2001 If we use Dbase tables to store these fields
    'we can't have field names bigger than 10 chars.  So thats why
    'the field name is just the IB Number.  If you need to know what
    'the field refers to have to have a printout This function
    If PopulateUDT(psClaimNo, psCheckNum, pCurDB) Then
    
        With mudtIBSafeco
            'BGS 10.31.2001 update the class with the ibnumber
            msIBNumber = .IB02_sIBNumber
            pclsUL.AddField .IB00_dtContactedDate, "IB00_dtCD" 'Date Contacted
            pclsUL.AddField .IB00_lSSN, "IB00_lSSN", , vbLong
            pclsUL.AddField .IB01_sSubToCarrier, "IB01_s"
            pclsUL.AddField .IB02_sIBNumber, "IB02_s"
            pclsUL.AddField .IB03_sCheckNUM, "IB03_s"
            pclsUL.AddField .IB04_sCatCode, "IB04_s"
            pclsUL.AddField .IB05_sLocation, "IB05_s"
            pclsUL.AddField .IB05a_sState, "IB05a_s"
            pclsUL.AddField .IB06_dtDateClosed, "IB06_dt"
            pclsUL.AddField .IB07_sAdjusterName, "IB07_s"
            pclsUL.AddField .IB08_sAdjCRID, "IB08_s"
            pclsUL.AddField .IB09_sSALN, "IB09_s"
            pclsUL.AddField .IB09a_sPolicyNumber, "IB09a_s"
            pclsUL.AddField .IB10_sInsuredName, "IB10_s"
            pclsUL.AddField .IB11_sLossLocation, "IB11_s"
            pclsUL.AddField .IB12_dtDateOfLoss, "IB12_dt"
            pclsUL.AddField .IB13_cGrossLoss, "IB13_c"
            pclsUL.AddField .IB14_cDepreciation, "IB14_c"
            pclsUL.AddField .IB14a_sSupplement, "IB14a_s"
            pclsUL.AddField .IB14b_sRebilled, "IB14b_s"
            pclsUL.AddField .IB15_cDeductible, "IB15_c"
            pclsUL.AddField .IB15a_cLessExcessLimits, "IB15a_c"
            pclsUL.AddField .IB15b_sExcessLimDesc, "IB15b_s"
            pclsUL.AddField .IB16_cNetClaim, "IB16_c"
            pclsUL.AddField .IB17_cServiceFee, "IB17_c"
            pclsUL.AddField .IB17a_cMiscServiceFee, "IB17a_c"
            pclsUL.AddField .IB18_sServiceFeeComment, "IB18_s"
            pclsUL.AddField .IB18a_sMiscServiceFeeComment, "IB18a_s"
            pclsUL.AddField .IB19_cOutBuildingsFee, "IB19_c"
            pclsUL.AddField .IB19a_iOutBuildingsCount, "IB19a_i"
            pclsUL.AddField .IB19b_cOutBuildingsAmount, "IB19b_c"
            pclsUL.AddField .IB20_sOutBuildingsFeeComment, "IB20_s"
            pclsUL.AddField .IB21_cTwoStoryCharge, "IB21_c"
            pclsUL.AddField .IB22_cSteepCharge, "IB22_c"
            pclsUL.AddField .IB23_cInteriorDamageCharge, "IB23_c"
            pclsUL.AddField .IB24_cExternalDamageBGCharge, "IB24_c"
            pclsUL.AddField .IB25_cServiceFeeSubTotal, "IB25_c"
            pclsUL.AddField .IB26_iPhotoCount, "IB26_i"
            pclsUL.AddField .IB27_cPhotoFee, "IB27_c"
            pclsUL.AddField .IB28_iOther, "IB28_i"
            pclsUL.AddField .IB29_cOtherFee, "IB29_c"
            pclsUL.AddField .IB29a_sMiscExpenseFeeComment, "IB29a_s"
            pclsUL.AddField .IB29b_cMiscExpenseFee, "IB29b_c"
            pclsUL.AddField .IB30_cTotalExpenses, "IB30_c"
            pclsUL.AddField .IB31_dTaxPercent, "IB31_d"
            pclsUL.AddField .IB32_cTaxAmount, "IB32_c"
            pclsUL.AddField .IB33_cTotalAdjustingFee, "IB33_c"
            pclsUL.AddField .IB33a_sAccountCode, "IB33a_s"
            pclsUL.AddField .IB34_sPaymentForProperty, "IB34_s"
            pclsUL.AddField .IB35_sPaymentForAuto, "IB35_s"
            pclsUL.AddField .IB36_sPaymentForFinal, "IB36_s"
            pclsUL.AddField .IB37_sPaymentForPartial, "IB37_s"
            pclsUL.AddField .IB38_sPaymentIsAddSupplementY, "IB38_s"
            pclsUL.AddField .IB39_sPaymentIsAddSupplementN, "IB39_s"
            pclsUL.AddField .IB40_sPaymentIsAssociatedY, "IB40_s"
            pclsUL.AddField .IB41_sPaymentISAssociatedN, "IB41_s"
            pclsUL.AddField .IB42_sClassOfClaim, "IB42_s"
            pclsUL.AddField .IB43_sCauseOfLoss, "IB43_s"
            pclsUL.AddField .IB44_sTexasSubCovCode, "IB44_s"
            pclsUL.AddField .IB45_sTexasSuffix, "IB45_s"
            pclsUL.AddField .IB46_cTexasRoofDepreciation, "IB46_c"
            pclsUL.AddField .IB47_sTypeOfPropLossBuilding, "IB47_s"
            pclsUL.AddField .IB48_sTypeOfPropLossContents, "IB48_s"
            pclsUL.AddField .IB49_sTypeOfPropLossALE, "IB49_s"
            pclsUL.AddField .IB49a_sOtherPropLoss, "IB49a_s"
            pclsUL.AddField .IB50_sInsuredPayeeName, "IB50_s"
            pclsUL.AddField .IB51_sPayeeNames, "IB51_s"
            pclsUL.AddField .IB52_sAddress, "IB52_s"
            pclsUL.AddField .IB53_cAmountOfCheck, "IB53_c"
            pclsUL.AddField .IB54_sCatCode, "IB54_s"
            pclsUL.AddField .IB55_sFieldHandled, "IB55_s"
            pclsUL.AddField .IB56_sTotalLoss, "IB56_s"
            pclsUL.AddField .IB57_sCashInLieu, "IB57_s"
            pclsUL.AddField .IB58_sOwnerRetainSalvage, "IB58_s"
            pclsUL.AddField .IB59_sSub, "IB59_s"
            pclsUL.AddField .IB60_sSalvage, "IB60_s"
            pclsUL.AddField .IB61_sInstructions, "IB61_s"
            pclsUL.AddField .IB62_sRequestedBy, "IB62_s"
            pclsUL.AddField .IB63_dtDate, "IB63_dt"
            pclsUL.AddField .IB64_sApproveBy, "IB64_s"
            pclsUL.AddField .IB65_dtApproveDate, "IB65_dt"
            pclsUL.AddField .IB66_sIssuedBy, "IB66_s"
            pclsUL.AddField .IB67_sRetrievedBy, "IB67_s"
        End With
    Else
        clsCarUL_PopulateFields = False
    End If
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    clsCarUL_PopulateFields = False
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarUL_ClassName & vbCrLf & " Public Function PopulateFields"
End Function

Public Function clsCarUL_UpdateBatchesUDT(poBatches As V2ECKeyBoard.clsBatches, poUL As V2ECKeyBoard.clsUpload) As Boolean
    On Error GoTo EH
    Dim udtBatches As udtBatches
    Dim oField As V2ECKeyBoard.clsField
    Dim sClaimNo As String
    Dim sCheckNUM As String
    'Error message
    Dim sMess As String
    
    clsCarUL_UpdateBatchesUDT = True
    With udtBatches
        
        For Each oField In poUL.Fields
            Select Case UCase(oField.FieldName)
                Case "IB00_DTCD" ' 4.22.2002 Update the WebControl DB
                    .dtContacted = oField.Value
                Case "IB00_LSSN" '11.21.2001 SSN used to match up Cat_no / Site
                    .lSSN = oField.Value
                Case "IB13_C" '_cGrossLoss"
                    .cGrossLoss = oField.Value
                Case "IB32_C" '_cTaxAmount"
                    .cTaxesTotal = oField.Value
                Case "IB33_C" '_cTotalAdjustingFee"
                    .cTotalFee = oField.Value
                Case "IB25_C" '_cServiceFeeSubTotal"
                    .cTotalServiceFees = oField.Value
                Case "IB63_DT" '_dtDate"
                    .dtDate = oField.Value
                Case "IB06_DT" '_dtDateClosed"
                    .dtDateClosed = CDate(oField.Value)
                Case "IB12_DT" '_dtDateOfLoss"
                    .dtDateOfLoss = oField.Value
                Case "IB08_S" '_sAdjCRID"
                    .sAdjuster_I = oField.Value
                Case "IB07_S" '_sAdjusterName"
                    .sAdjuster_N = oField.Value
                Case "IB04_S" '_sCatCode"
                    .sCATCode = oField.Value
                Case "IB09_S" '_sSALN"
                    .sClaimNumber = oField.Value
                Case "IB02_S" '_sIBNumber"
                    .sIBNumber = oField.Value
                    sClaimNo = oField.Value
                Case "IB10_S" '_sInsuredName"
                    .sInsuredName = oField.Value
                Case "IB05_S" '_sLocation"
                    .sLossLocationCity = oField.Value
                Case "IB05A_S" '_sState"
                    .sLossLocationState = oField.Value
                Case "IB03_S" '_sCheckNUM"
                    sCheckNUM = oField.Value
                Case "IB30_C" 'IB30_cTotalExpenses, "IB30_c"
                    .cMiscCharges = oField.Value
            End Select
        Next
        'Set this flag to false since the RT upload will do this for us.
        'But leave the option to set the Flag. This may be usefull in the future ?
        mbDoDBUpdate = False
        
        .sBatchID = GetBatchID(sClaimNo, sCheckNUM)
        .sCarrierCode = "SC"
        .cAdministrativeFee = 0
        
    End With
    
    poBatches.Batches = udtBatches
    
    'BGS 11.21.2001 This will be a log Table of all transactions (Approach)
    If Not poBatches.UpdateBatchesTable Then
        clsCarUL_UpdateBatchesUDT = False
        Exit Function
    End If
    'This is the Claims Information update (Claims)
    If Not poBatches.UpdateBillDataTable Then
        clsCarUL_UpdateBatchesUDT = False
        Exit Function
    End If
    'This will update the IB Table (Approach)
    If Not poBatches.UpdateIBTable(Me, poUL) Then
        clsCarUL_UpdateBatchesUDT = False
        Exit Function
    End If
    
    Exit Function
EH:
    clsCarUL_UpdateBatchesUDT = False
    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
    sMess = sMess & "ERROR # " & Err.Number & vbCrLf
    sMess = sMess & Err.Description & vbCrLf
    sMess = sMess & clsCarUL_ClassName & " Public Function UpdateBatches" & vbCrLf
    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
    poUL.FireErrorMess sMess
End Function

Public Sub clsCarUL_UpdateDB(poBatches As V2ECKeyBoard.clsBatches, poUL As V2ECKeyBoard.clsUpload)
    On Error GoTo EH
    Dim sMess As String
    If mbDoDBUpdate Then
        poUL.FireUpdateDB poBatches.Batches
    End If
    
    Exit Sub
EH:
    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
    sMess = sMess & "ERROR # " & Err.Number & vbCrLf
    sMess = sMess & Err.Description & vbCrLf
    sMess = sMess & clsCarUL_ClassName & " Public Sub UpdateDB" & vbCrLf
    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
    poUL.FireErrorMess sMess
End Sub

Public Function clsCarUL_GetTotalFee(psClaimNumber As String, psIBNumber As String, poBatch As V2ECKeyBoard.clsBatches, poUL As V2ECKeyBoard.clsUpload) As Currency
    clsCarUL_GetTotalFee = 0
End Function

Private Function PopulateUDT(psClaimNo As String, psCheckNum As String, pCurDB As Object) As Boolean
    On Error GoTo EH
    Dim RS As Recordset
    Dim sSQL As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    PopulateUDT = True
    
    'BGS 10.12.2001 Need to populate the IB Safeco udt with records for the
    'Claimno and Checknum Provided
    sSQL = "SELECT " & BuildClaimsSQLAB & " "
    sSQL = sSQL & "FROM Billing A LEFT JOIN Checks B "
    sSQL = sSQL & "ON A.ClaimNo = B.ClaimNo "
    sSQL = sSQL & "WHERE A.ClaimNo = '" & goUtil.utCleanSQLString(psClaimNo) & "' "
    sSQL = sSQL & "AND B.CheckNum = '" & goUtil.utCleanSQLString(psCheckNum) & "' "
    sSQL = sSQL & "AND A.ID = B.BillingID "
    
    Set RS = pCurDB.OpenRecordset(sSQL)
    
    With mudtIBSafeco
        If Not RS.EOF Then
            RS.MoveFirst
            .IB00_dtContactedDate = IIf(IsNull(RS!IB00_dtContactedDate) Or RS!IB00_dtContactedDate = "//", NULL_DATE, RS!IB00_dtContactedDate)
            .IB00_lSSN = GetSetting("ECS", "WEB_SECURITY", "USERNAME", NO_SSN)
            'BGS 10.11.2001 Put in the Carrier Header Here
            .IB01_sSubToCarrier = "Safeco"
            .IB02_sIBNumber = IIf(IsNull(RS!IB02_sIBNumber), vbNullString, RS!IB02_sIBNumber)
            .IB03_sCheckNUM = IIf(IsNull(RS!IB03_sCheckNUM), vbNullString, RS!IB03_sCheckNUM)
            'BGS 5.13.2002 Ony use first 2 Chars of Cat Code
            .IB04_sCatCode = IIf(IsNull(RS!IB04_sCatCode), vbNullString, left(RS!IB04_sCatCode, 2))
            .IB05_sLocation = IIf(IsNull(RS!IB05_sLocation), vbNullString, RS!IB05_sLocation)
            .IB05a_sState = IIf(IsNull(RS!IB05a_sState), vbNullString, RS!IB05a_sState)
            .IB06_dtDateClosed = IIf(IsNull(RS!IB06_dtDateClosed) Or RS!IB06_dtDateClosed = "//", NULL_DATE, RS!IB06_dtDateClosed)
            .IB07_sAdjusterName = IIf(IsNull(RS!IB07_sAdjusterName), vbNullString, RS!IB07_sAdjusterName)
            .IB08_sAdjCRID = IIf(IsNull(RS!IB08_sAdjCRID), vbNullString, RS!IB08_sAdjCRID)
            .IB09_sSALN = IIf(IsNull(RS!IB09_sSALN), vbNullString, RS!IB09_sSALN)
            .IB09a_sPolicyNumber = IIf(IsNull(RS!IB09a_sPolicyNumber), vbNullString, RS!IB09a_sPolicyNumber)
            .IB10_sInsuredName = IIf(IsNull(RS!IB10_sInsuredName), vbNullString, RS!IB10_sInsuredName)
            .IB11_sLossLocation = IIf(IsNull(RS!IB11_sLossLocation), vbNullString, RS!IB11_sLossLocation)
            .IB12_dtDateOfLoss = IIf(IsNull(RS!IB12_dtDateOfLoss) Or RS!IB12_dtDateOfLoss = "//", NULL_DATE, RS!IB12_dtDateOfLoss)
            .IB13_cGrossLoss = IIf(IsNull(RS!IB13_cGrossLoss), 0, RS!IB13_cGrossLoss)
            .IB14_cDepreciation = IIf(IsNull(RS!IB14_cDepreciation), 0, RS!IB14_cDepreciation)
            .IB14a_sSupplement = IIf(IsNull(RS!IB14a_sSupplement), vbNullString, RS!IB14a_sSupplement)
            .IB14b_sRebilled = IIf(IsNull(RS!IB14b_sRebilled), vbNullString, RS!IB14b_sRebilled)
            .IB15_cDeductible = IIf(IsNull(RS!IB15_cDeductible), 0, RS!IB15_cDeductible)
            .IB15a_cLessExcessLimits = IIf(IsNull(RS!IB15a_cLessExcessLimits), 0, RS!IB15a_cLessExcessLimits)
            .IB15b_sExcessLimDesc = IIf(IsNull(RS!IB15b_sExcessLimDesc), vbNullString, RS!IB15b_sExcessLimDesc)
            .IB16_cNetClaim = IIf(IsNull(RS!IB16_cNetClaim), 0, RS!IB16_cNetClaim)
            .IB17_cServiceFee = IIf(IsNull(RS!IB17_cServiceFee), 0, RS!IB17_cServiceFee)
            .IB17a_cMiscServiceFee = IIf(IsNull(RS!IB17a_cMiscServiceFee), 0, RS!IB17a_cMiscServiceFee)
            .IB18_sServiceFeeComment = IIf(IsNull(RS!IB18_sServiceFeeComment), vbNullString, RS!IB18_sServiceFeeComment)
            .IB18a_sMiscServiceFeeComment = IIf(IsNull(RS!IB18a_sMiscServiceFeeComment), vbNullString, RS!IB18a_sMiscServiceFeeComment)
            .IB19_cOutBuildingsFee = IIf(IsNull(RS!IB19_cOutBuildingsFee), 0, RS!IB19_cOutBuildingsFee)
            .IB19a_iOutBuildingsCount = IIf(IsNull(RS!IB19a_iOutBuildingsCount), 0, RS!IB19a_iOutBuildingsCount)
            .IB19b_cOutBuildingsAmount = GetSetting("EBERLSSC", "FEES", "OUT_BUILDING_CHARGE", "20.00")
            .IB20_sOutBuildingsFeeComment = IIf(IsNull(RS!IB20_sOutBuildingsFeeComment), vbNullString, RS!IB20_sOutBuildingsFeeComment)
            .IB21_cTwoStoryCharge = IIf(IsNull(RS!IB21_cTwoStoryCharge), 0, RS!IB21_cTwoStoryCharge)
            .IB22_cSteepCharge = IIf(IsNull(RS!IB22_cSteepCharge), 0, RS!IB22_cSteepCharge)
            .IB23_cInteriorDamageCharge = IIf(IsNull(RS!IB23_cInteriorDamageCharge), 0, RS!IB23_cInteriorDamageCharge)
            .IB24_cExternalDamageBGCharge = IIf(IsNull(RS!IB24_cExternalDamageBGCharge), 0, RS!IB24_cExternalDamageBGCharge)
            .IB25_cServiceFeeSubTotal = IIf(IsNull(RS!IB25_cServiceFeeSubTotal), 0, RS!IB25_cServiceFeeSubTotal)
            .IB26_iPhotoCount = IIf(IsNull(RS!IB26_iPhotoCount), 0, RS!IB26_iPhotoCount)
            .IB27_cPhotoFee = IIf(IsNull(RS!IB27_cPhotoFee), 0, RS!IB27_cPhotoFee)
            .IB28_iOther = IIf(IsNull(RS!IB28_iOther), 0, RS!IB28_iOther)
            .IB29_cOtherFee = IIf(IsNull(RS!IB29_cOtherFee), 0, RS!IB29_cOtherFee)
            .IB29a_sMiscExpenseFeeComment = IIf(IsNull(RS!IB29a_sMiscExpenseFeeComment), vbNullString, RS!IB29a_sMiscExpenseFeeComment)
            .IB29b_cMiscExpenseFee = IIf(IsNull(RS!IB29b_cMiscExpenseFee), 0, RS!IB29b_cMiscExpenseFee)
            .IB30_cTotalExpenses = IIf(IsNull(RS!IB30_cTotalExpenses), 0, RS!IB30_cTotalExpenses)
            .IB31_dTaxPercent = IIf(IsNull(RS!IB31_dTaxPercent), 0, RS!IB31_dTaxPercent)
            .IB32_cTaxAmount = IIf(IsNull(RS!IB32_cTaxAmount), 0, RS!IB32_cTaxAmount)
            .IB33_cTotalAdjustingFee = IIf(IsNull(RS!IB33_cTotalAdjustingFee), 0, RS!IB33_cTotalAdjustingFee)
            .IB33a_sAccountCode = GetSetting("EBERLSSC", "FEES", "ACCOUNT_CODE", " ")
            .IB34_sPaymentForProperty = IIf(IsNull(RS!IB34_sPaymentForProperty), vbNullString, RS!IB34_sPaymentForProperty)
            .IB35_sPaymentForAuto = IIf(IsNull(RS!IB35_sPaymentForAuto), vbNullString, RS!IB35_sPaymentForAuto)
            .IB36_sPaymentForFinal = IIf(IsNull(RS!IB36_sPaymentForFinal), vbNullString, RS!IB36_sPaymentForFinal)
            .IB37_sPaymentForPartial = IIf(IsNull(RS!IB37_sPaymentForPartial), vbNullString, RS!IB37_sPaymentForPartial)
            .IB38_sPaymentIsAddSupplementY = IIf(IsNull(RS!IB38_sPaymentIsAddSupplementY), vbNullString, RS!IB38_sPaymentIsAddSupplementY)
            .IB39_sPaymentIsAddSupplementN = IIf(IsNull(RS!IB39_sPaymentIsAddSupplementN), vbNullString, RS!IB39_sPaymentIsAddSupplementN)
            .IB40_sPaymentIsAssociatedY = IIf(IsNull(RS!IB40_sPaymentIsAssociatedY), vbNullString, RS!IB40_sPaymentIsAssociatedY)
            .IB41_sPaymentISAssociatedN = IIf(IsNull(RS!IB41_sPaymentISAssociatedN), vbNullString, RS!IB41_sPaymentISAssociatedN)
            .IB42_sClassOfClaim = IIf(IsNull(RS!IB42_sClassOfClaim), vbNullString, RS!IB42_sClassOfClaim)
            .IB43_sCauseOfLoss = IIf(IsNull(RS!IB43_sCauseOfLoss), vbNullString, RS!IB43_sCauseOfLoss)
            .IB45_sTexasSuffix = IIf(IsNull(RS!IB45_sTexasSuffix), vbNullString, RS!IB45_sTexasSuffix)
            .IB46_cTexasRoofDepreciation = IIf(IsNull(RS!IB46_cTexasRoofDepreciation), vbNullString, RS!IB46_cTexasRoofDepreciation)
            .IB47_sTypeOfPropLossBuilding = IIf(IsNull(RS!IB47_sTypeOfPropLossBuilding), vbNullString, RS!IB47_sTypeOfPropLossBuilding)
            .IB48_sTypeOfPropLossContents = IIf(IsNull(RS!IB48_sTypeOfPropLossContents), vbNullString, RS!IB48_sTypeOfPropLossContents)
            .IB49_sTypeOfPropLossALE = IIf(IsNull(RS!IB49_sTypeOfPropLossALE), vbNullString, RS!IB49_sTypeOfPropLossALE)
            .IB49a_sOtherPropLoss = vbNullString    'Future use in vs2.0
            .IB50_sInsuredPayeeName = IIf(IsNull(RS!IB50_sInsuredPayeeName), vbNullString, RS!IB50_sInsuredPayeeName)
            .IB51_sPayeeNames = IIf(IsNull(RS!IB51_sPayeeNames), vbNullString, RS!IB51_sPayeeNames)
            .IB52_sAddress = IIf(IsNull(RS!IB52_sAddress), vbNullString, RS!IB52_sAddress)
            .IB53_cAmountOfCheck = IIf(IsNull(RS!IB53_cAmountOfCheck), 0, RS!IB53_cAmountOfCheck)
            .IB54_sCatCode = IIf(IsNull(RS!IB54_sCatCode), vbNullString, RS!IB54_sCatCode)
            .IB55_sFieldHandled = IIf(IsNull(RS!IB55_sFieldHandled), vbNullString, RS!IB55_sFieldHandled)
            .IB56_sTotalLoss = IIf(IsNull(RS!IB56_sTotalLoss), vbNullString, RS!IB56_sTotalLoss)
            .IB57_sCashInLieu = IIf(IsNull(RS!IB57_sCashInLieu), vbNullString, RS!IB57_sCashInLieu)
            .IB58_sOwnerRetainSalvage = IIf(IsNull(RS!IB58_sOwnerRetainSalvage), vbNullString, RS!IB58_sOwnerRetainSalvage)
            .IB59_sSub = IIf(IsNull(RS!IB59_sSub), vbNullString, RS!IB59_sSub)
            .IB60_sSalvage = IIf(IsNull(RS!IB60_sSalvage), vbNullString, RS!IB60_sSalvage)
            .IB61_sInstructions = IIf(IsNull(RS!IB61_sInstructions), vbNullString, RS!IB61_sInstructions)
            .IB62_sRequestedBy = IIf(IsNull(RS!IB62_sRequestedBy), vbNullString, RS!IB62_sRequestedBy)
            .IB63_dtDate = IIf(IsNull(RS!IB63_dtDate), NULL_DATE, RS!IB63_dtDate)
            .IB64_sApproveBy = vbNullString     'Future use in vs2.0
            .IB65_dtApproveDate = NULL_DATE     'Future use in vs2.0
            .IB66_sIssuedBy = vbNullString      'Future use in vs2.0
            .IB67_sRetrievedBy = vbNullString   'Future use in vs2.0
        End If
    End With
    RS.Close
    Set RS = Nothing
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    PopulateUDT = False
    Set RS = Nothing
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarUL_ClassName & vbCrLf & " Private Function PopulateUDT"
End Function

Private Function BuildClaimsSQLAB() As String
    On Error GoTo EH
    Dim sSQL As String
    Dim sT As String
    Dim dtDate As Date
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    dtDate = Format(Now, "MM/DD/YY")
    
    sSQL = "Mid(A.ContactedDate,5,2) & '/' & Mid(A.ContactedDate,7,2) & '/' & Left(A.ContactedDate,4) AS IB00_dtContactedDate, "
    sSQL = sSQL & "B.ClaimNo & IIF(B.Supplement > 0, 'S' & B.Supplement & '.', '') & IIF(B.Rebilled > 0, 'R' & B.Rebilled & '.', '')  AS IB02_sIBNumber, "
    sSQL = sSQL & "B.CheckNUM AS IB03_sCheckNUM, "
    sSQL = sSQL & "A.CatCode AS IB04_sCatCode, "
    sSQL = sSQL & "A.SiteCity AS IB05_sLocation, "
    sSQL = sSQL & "A.SiteState AS IB05a_sState, "
    sSQL = sSQL & "Mid(A.ClosedDate,5,2) & '/' & Mid(A.ClosedDate,7,2) & '/' & Left(A.ClosedDate,4) AS IB06_dtDateClosed, "
    sSQL = sSQL & "A.AdjustorFirstName & ' ' & Trim(A.AdjustorLastName) AS IB07_sAdjusterName, "
    sSQL = sSQL & "A.ClaimRepIDNO AS IB08_sAdjCRID, "
    sSQL = sSQL & "A.ClientClaimNo AS IB09_sSALN, "
    sSQL = sSQL & "A.PolicyNumber As IB09a_sPolicyNumber, "
    sSQL = sSQL & "A.NewInsuredNames AS IB10_sInsuredName, "
    sSQL = sSQL & "A.NewLossLocation AS IB11_sLossLocation, "
    sSQL = sSQL & "Mid(A.LossDate,5,2) & '/' & Mid(A.LossDate,7,2) & '/' & Mid(A.LossDate,3,2) AS IB12_dtDateOfLoss, "
    sSQL = sSQL & "A.FullCostOfRepair AS IB13_cGrossLoss, "
    sSQL = sSQL & "A.RecoverableDepreciation + A.NonRecovDepr AS IB14_cDepreciation, "
    sSQL = sSQL & "IIF(B.Supplement > 0, 'X', '') AS IB14a_sSupplement, "
    sSQL = sSQL & "IIF(B.Rebilled >0, 'X', '') AS IB14b_sRebilled, "
    sSQL = sSQL & "A.Deductible AS IB15_cDeductible, "
    sSQL = sSQL & "A.LessExcessLimits AS IB15a_cLessExcessLimits, "
    sSQL = sSQL & "B.ExcessLimDesc as IB15b_sExcessLimDesc, "
    sSQL = sSQL & "IIF(A.NetActualCashValueClaim <=0, 0, A.NetActualCashValueClaim) AS IB16_cNetClaim, "
    sSQL = sSQL & "A.ServiceFee AS IB17_cServiceFee, "
    sSQL = sSQL & "A.MiscServiceFee As IB17a_cMiscServiceFee, "
    sSQL = sSQL & "IIF(A.BillingHours <> '', IIF(CCur(A.BillingHours) > 0, '(' & A.BillingHours & ' hours @ $' & A.ServiceFee / CCur(A.BillingHours) & ')', ''), '') AS IB18_sServiceFeeComment, "
    sSQL = sSQL & "A.MiscServiceFeeComment As IB18a_sMiscServiceFeeComment, "
    sSQL = sSQL & "A.OutBuildingsFee AS IB19_cOutBuildingsFee, "
    sSQL = sSQL & "A.OutBuildings AS IB19a_iOutBuildingsCount, "
    sSQL = sSQL & "IIF(A.BillingFenceOnly = '1', '(Fence Only)', '') AS IB20_sOutBuildingsFeeComment, "
    sSQL = sSQL & "A.TwoStoryDamage AS IB21_cTwoStoryCharge, "
    sSQL = sSQL & "A.SteepCharge AS IB22_cSteepCharge, "
    sSQL = sSQL & "A.BillingInteriorDamageDollars AS IB23_cInteriorDamageCharge, "
    sSQL = sSQL & "A.BillingExteriorDamageDollars AS IB24_cExternalDamageBGCharge, "
    sSQL = sSQL & "A.ServiceFee + A.OutBuildingsFee + A.TwoStoryDamage + A.SteepCharge + A.BillingInteriorDamageDollars + A.BillingExteriorDamageDollars + A.MiscServiceFee AS IB25_cServiceFeeSubTotal, "
    sSQL = sSQL & "A.PhotographCount AS IB26_iPhotoCount, "
    sSQL = sSQL & "A.PhotographFee AS IB27_cPhotoFee, "
    sSQL = sSQL & "A.Miles AS IB28_iOther, "
    sSQL = sSQL & "A.MilesFee AS IB29_cOtherFee, "
    sSQL = sSQL & "A.MiscFeesComment As IB29a_sMiscExpenseFeeComment, "
    sSQL = sSQL & "A.MiscFees As IB29b_cMiscExpenseFee, "
    sSQL = sSQL & "A.PhotographFee + A.MilesFee + A.MiscFees AS IB30_cTotalExpenses, "
    sSQL = sSQL & "A.TaxPercent AS IB31_dTaxPercent, "
    sSQL = sSQL & "A.TaxPercentAmount AS IB32_cTaxAmount, "
    sSQL = sSQL & "A.InvoiceTotal AS IB33_cTotalAdjustingFee, "
    sSQL = sSQL & "IIF(B.ForProperty = True,'X','') AS IB34_sPaymentForProperty, "
    sSQL = sSQL & "IIf(B.ForAuto = True,'X','') AS IB35_sPaymentForAuto, "
    sSQL = sSQL & "IIF(B.FinalPayment = True,'X','') AS IB36_sPaymentForFinal, "
    sSQL = sSQL & "IIF(B.PartialPayment = True,'X','') AS IB37_sPaymentForPartial, "
    sSQL = sSQL & "IIF(B.ForAdditionalSupplement = True,'X','') AS IB38_sPaymentIsAddSupplementY, "
    sSQL = sSQL & "IIF(B.ForAdditionalSupplement = False,'X','') AS IB39_sPaymentIsAddSupplementN, "
    sSQL = sSQL & "IIF(B.PaymentAssociated = True,'X','') AS IB40_sPaymentIsAssociatedY, "
    sSQL = sSQL & "IIF(B.PaymentAssociated = False,'X','') AS IB41_sPaymentISAssociatedN, "
    sSQL = sSQL & "Left(B.TypeOfLoss,2) AS IB42_sClassOfClaim, "
    sSQL = sSQL & "Left(B.CauseOfLoss,2) AS IB43_sCauseOfLoss, "
    sSQL = sSQL & "B.TexasSubCode AS IB44_sTexasSubCovCode, "
    sSQL = sSQL & "B.TexasSuffix AS IB45_sTexasSuffix, "
    sSQL = sSQL & "B.TexasRoofDepreciation AS IB46_cTexasRoofDepreciation, "
    sSQL = sSQL & "IIF(Left(B.TypeOfLoss,2)= '01','X','') AS IB47_sTypeOfPropLossBuilding, "
    sSQL = sSQL & "IIF(Left(B.TypeOfLoss,2)= '02','X','') AS IB48_sTypeOfPropLossContents, "
    'BGS 2.28.2002 143  Class of Loss - 88 s/b ALE
    'Todd let me know that 88 should be the ALE code, not 89
    'Its been 89 since forever so we will change it to 88 but we still
    'need to account for 89 being used previously.
    '8.30.2002 Use 03 for Safeco
    sSQL = sSQL & "IIF(Left(B.TypeOfLoss,2)= '03' Or Left(B.TypeOfLoss,2)= '03','X','') AS IB49_sTypeOfPropLossALE, "
    sSQL = sSQL & "IIF(IsNull(A.NewSecondInsuredNames) or Trim(A.NewSecondInsuredNames) = '', A.NewInsuredNames, Trim(A.NewInsuredNames) & ' And ' & A.NewSecondInsuredNames) AS IB50_sInsuredPayeeName, "
    sSQL = sSQL & "IIF(B.IncludeMorgOnDraft = TRUE, A.MortgageCompany, '') AS IB51_sPayeeNames, "
    sSQL = sSQL & "A.NewMailingAddress AS IB52_sAddress, "
    sSQL = sSQL & "B.AmountOfCheck AS IB53_cAmountOfCheck, "
    sSQL = sSQL & "A.CatCode AS IB54_sCatCode, "
    sSQL = sSQL & "IIF(B.FieldHandled = True, 'Yes', 'No') AS IB55_sFieldHandled, "
    sSQL = sSQL & "IIF(B.TotalLoss = True, 'YES', 'NO') AS IB56_sTotalLoss, "
    sSQL = sSQL & "IIF(B.CashInLieu = True, 'Yes', 'No') AS IB57_sCashInLieu, "
    sSQL = sSQL & "IIF(B.OwnerRetainedSalvage = True, 'Yes', 'No') AS IB58_sOwnerRetainSalvage, "
    sSQL = sSQL & "IIF(B.Subrogation = True, 'Yes', 'No') AS IB59_sSub, "
    sSQL = sSQL & "IIF(B.Salvage = True, 'Yes', 'No') AS IB60_sSalvage, "
    sSQL = sSQL & "B.Instructions AS IB61_sInstructions, "
    sSQL = sSQL & "Trim(A.AdjustorFirstName) & ' ' & Trim(A.AdjustorLastName) AS IB62_sRequestedBy, "
    sSQL = sSQL & "#" & dtDate & "# AS IB63_dtDate"
    
    BuildClaimsSQLAB = sSQL
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarUL_ClassName & vbCrLf & " Private Function BuildClaimsSQLAB"
End Function

Private Function GetBatchID(psClaimNo As String, psCheckNum As String) As String
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    Randomize Timer
    GetBatchID = Format(Now(), "mmddyyhhmmss") & (Int(Rnd * 9)) & (Int(Rnd * 9)) & (Int(Rnd * 9)) & (Int(Rnd * 9)) & "E"

    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarUL_ClassName & vbCrLf & " Private Function GetBatchID"
End Function

Private Sub Class_Initialize()
    msClassName = "clsULSafeco"
    msIBTableName = "IBSafeco"
End Sub
