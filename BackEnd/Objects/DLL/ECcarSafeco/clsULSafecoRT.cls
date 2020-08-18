VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsULSafecoRT"
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
'2. udtSafecoRT (Try to use the Same convention)
'3. Update Private Function PopulateUDT
'4. Update Private Function BuildClaimsSQLAB
'5. Update Public Function PopulateFields
'
'Each function that needs to be updated should be Easy, no imbeded SQL
'Each Field Item will be another Line of code. Thats why using the NumAlpha convention
'is good for maintenance, ie future additions. As well using the Num Alpha allows for
'Databases that will not allow long field names
'--------------------------------------------------------------------------------------


'All items not Beginning with RT are (Not on IB REBORT)
Private Type udtSafecoRT
    'Claim Info
    CI01_sStatus As String
    CI02_sSiteAddress As String
    CI03_sSiteCity As String
    CI04_sSiteState As String
    CI05_sSiteZip As String
    CI06_dtInspectedDate As Date            'Date Inspected
    CI07_dtReceivedDate As Date             'Date Received
    CI08_dtContactedDate As Date            'Date contacted
    CI09_dtAssignedDate As Date             'AssignedDate
    CI10_sTypeOfLoss As String              'TypeOfLoss
    CI11_sPolicyNumber As String            'PolicyNumber
    CI12_sMailAddress As String             'MailAddress
    CI13_sMortgageeName As String           'MortgageeName
    CI14_cBuildingLimits As Currency        'BuildingLimits
    CI15_cContentsLimits As Currency        'ContentsLimits
    CI16_cAPSLimits As Currency
    CI17_cBuildingRCLimits As Currency
    CI18_cContentsRCLimits As Currency
    CI19_cAPSRCLimits As Currency
    'IndemInfo
    II01_cRecovDep As Currency
    II02_cNonRecovDep As Currency
    'Billing Info
    LID As Long                       'BillingID (Check ID) (Not on IB Report)
    RT00_lSSN As String                      'ADJ SSN (Not on IB Report)
    RT01_sSubToCarrier As String            'Safeco
    RT02_sIBNumber As String                '{Checks.ClaimNo} + " " + {@Rebilled} + {@Supplement}
    RT03_sCheckNUM As String                'Checks.CheckNUM
    RT04_sCatCode As String                 'Ass.CatCode
    RT05_sLocation As String                'Ass.SiteCity
    RT05a_sState As String                  'Ass.SiteState
    RT06_dtDateClosed As Date                '{Billing.ClosedDate}[5 to 6] + "/" + {Billing.ClosedDate}[7 to 8] + "/" + {Billing.ClosedDate}[1 to 4]
    RT07_sAdjusterName As String            'Trim({Billing.AdjustorFirstName}) + " " + trim({Billing.AdjustorLastName})
    RT08_sAdjCRID As String                 'ass.claimrepidno
    RT09_sSALN As String                     'ass.clientclaimno
    RT09a_sPolicyNumber As String
    RT10_sInsuredName As String              'ass.newinsurednames
    RT11_sLossLocation  As String            'ass.NewLossLocation
    RT12_dtDateOfLoss As Date                 '{Billing.LossDate}[5 to 6] + "/" + {Billing.LossDate}[7 to 8] + "/" + {Billing.LossDate}[3 to 4]
    RT13_cGrossLoss As Currency              'ass.FullCostOfRepair
    RT14_cDepreciation As Currency           '{Billing.RecoverableDepreciation}+{Billing.NonRecovDepr}
    RT14a_sSupplement As String              'if {Checks.Supplement} >0 then "X" else ""
    RT14b_sRebilled As String                'if {Checks.Rebilled} = true then "X" else ""
    RT15_cDeductible As Currency             'ass.Deductible
    RT15a_cLessExcessLimits As Currency      'ass.LessExcessLimits
    RT15b_sExcessLimDesc As String           'Checks.ExcessLimDesc
    RT16_cNetClaim As Currency               'if {@TotalNetActualCashValueClaim} <= 0 then 0 else {@TotalNetActualCashValueClaim}
    RT17_cServiceFee As Currency             'ass.ServiceFee
    RT17a_cMiscServiceFee As Currency        'Ass.MiscServiceFee
    RT18_sServiceFeeComment As String        'if {Billing.BillingHours} <> "" then "(" + {Billing.BillingHours} + " hours @ $75)" else ""
    RT18a_sMiscServiceFeeComment As String   'ass.MiscServiceFeeComment
    RT19_cOutBuildingsFee As Currency        'ass.OutBuildingsFee
    RT19a_iOutBuildingsCount As Integer      'ass.Outbuildings
    RT19b_cOutBuildingsAmount As Currency    'Reg Setting Amount per outbuilding
    RT20_sOutBuildingsFeeComment As String   'if {Billing.BillingFenceOnly} = "1" then "(Fence Only)" else ""
    RT21_cTwoStoryCharge As Currency         'ass.TwoStoryDamage
    RT22_cSteepCharge As Currency            'ass.SteepCharge
    RT23_cInteriorDamageCharge As Currency   'ass.BillingInteriorDamageDollars
    RT24_cExternalDamageBGCharge As Currency 'ass.BillingExteriorDamageDollars
    RT25_cServiceFeeSubTotal As Currency     '{Billing.ServiceFee}+{Billing.OutBuildingsFee}+{Billing.TwoStoryDamage}+{Billing.SteepCharge}+{Billing.BillingInteriorDamageDollars}+{Billing.BillingExteriorDamageDollars}
    RT26_iPhotoCount As Integer              'ass.photographCount
    RT27_cPhotoFee As Currency               'ass.photographfee
    RT28_iOther As Integer                   'ass.miles
    RT29_cOtherFee As Currency               'ass.MilesFee
    RT29a_sMiscExpenseFeeComment As String   'ass.MiscFeesComment
    RT29b_cMiscExpenseFee As Currency        'ass.MiscFees
    RT30_cTotalExpenses As Currency          '{Billing.PhotographFee}+{Billing.MilesFee}
    RT31_dTaxPercent As Double               'ass.TaxPercent
    RT32_cTaxAmount As Currency              'ass.TaxPercentAmount
    RT33_cTotalAdjustingFee As Currency      'ass.InvoiceTotal
    RT33a_sAccountCode As String             'Reg setting
    RT34_sPaymentForProperty As String      'if {Checks.ForProperty} = True then "X" else ""
    RT35_sPaymentForAuto As String          'if {Checks.ForAuto} = true then "X" else ""
    RT36_sPaymentForFinal As String         'if {Checks.FinalPayment} = true then "X" else ""
    RT37_sPaymentForPartial As String      'if {Checks.PartialPayment} = true then "X" else ""
    RT38_sPaymentIsAddSupplementY As String 'if {Checks.ForAdditionalSupplement} = true then "Y" else ""
    RT39_sPaymentIsAddSupplementN As String 'if {Checks.ForAdditionalSupplement} = false then "X" else ""
    RT40_sPaymentIsAssociatedY As String    'if {Checks.PaymentAssociated} = true then "X" else ""
    RT41_sPaymentISAssociatedN As String   'if {Checks.PaymentAssociated}= False then "X" else ""
    RT42_sClassOfClaim As String             '{Checks.TypeOfLoss}[1 to 2]
    RT43_sCauseOfLoss As String              '{Checks.CauseOfLoss}[1 to 2]
    RT44_sTexasSubCovCode As String          'Checks.texasSubCode
    RT45_sTexasSuffix As String              'Checks.texasSuffix
    RT46_cTexasRoofDepreciation As Currency 'Checks.TexasRoofDepreciation
    RT47_sTypeOfPropLossBuilding As String  'if {Checks.TypeOfLoss}[1 to 2]  = "01" then "X" else ""
    RT48_sTypeOfPropLossContents As String  'if {Checks.TypeOfLoss}[1 to 2]  = "02" then "X" else ""
    RT49_sTypeOfPropLossALE As String       'if {Checks.TypeOfLoss}[1 to 2]  = "03" then "X" else ""
    RT49a_sOtherPropLoss As String          'Used in future vs 2.0
    RT50_sInsuredPayeeName As String         'if isnull({Billing.NewSecondInsuredNames}) or trim({Billing.NewSecondInsuredNames}) = "" then {Billing.NewInsuredNames} else trim({Billing.NewInsuredNames}) + " and " + {Billing.NewSecondInsuredNames}
    RT51_sPayeeNames As String               'if {Checks.IncludeMorgOnDraft}= TRUE then {Billing.MortgageCompany} else ""
    RT52_sAddress As String                  'ass.NewMailing Address
    RT53_cAmountOfCheck As Currency          'Checks.AmountOfCheck
    RT54_sCatCode As String                  'ass.CatCode
    RT55_sFieldHandled As String            'if{Checks.FieldHandled}= true then "Yes" else "No"
    RT56_sTotalLoss As String              'if {Checks.TotalLoss} = true then "YES" else "NO"
    RT57_sCashInLieu As String              'if {Checks.CashInLieu}= true then "Yes" else "No"
    RT58_sOwnerRetainSalvage As String      'if{Checks.OwnerRetainedSalvage}= true then "Yes" else "No"
    RT59_sSub As String                     'if {Checks.Subrogation} = true then "Yes" else "No"
    RT60_sSalvage As String                 'if {Checks.Salvage} = true then "Yes" else "No"
    RT61_sInstructions As String             'Checks.Instructions
    RT62_sRequestedBy As String              'Trim({Billing.AdjustorFirstName}) + " " + trim({Billing.AdjustorLastName})
    RT63_dtDate As Date                       'Have to think about this one. Right now this is figured whenever the Adjuster Prints the Report
    RT64_sApproveBy As String               'Used in future vs 2.0
    RT65_dtApproveDate As Date              'Used in future vs 2.0
    RT66_sIssuedBy As String                'Used in future vs 2.0
    RT67_sRetrievedBy As String             'Used in future vs 2.0
End Type
'
'RT Types for ActLog, Photos, Indemnity, Attachments, etc etc
''udtBatchesRT will have a collection for each of these types
Public Type RTSafecoActLog
    sIBNumber As String
    sClientClaimNumber As String
    lActID As Long
    dtActDate As Date
    sActTime As String
    sActText As String
    dServiceTime As Double
    sAssociation As String
End Type

Public Type RTSafecoPhotoLog
    sIBNumber As String
    sClientClaimNumber As String
    lPhotoID As Long
    dtPhotoDate As Date
    sPhotoName As String
    sDescription As String
    lSortOrder As Long
End Type

Public Type RTSafecoIndemnity
    sIBNumber As String
    sClientClaimNumber As String
    lIndemID As Long
    sIndemType As String
    sClass As String
    sCauseOfLoss As String
    sDescription As String
    cReplacementCost As Currency
    cRecoverableDepreciation As Currency
    cNonRecoverableDepreciation As Currency
    cACVClaim As Currency
    cExcessLimit As Currency
    cACVLessExcessLimit As Currency
End Type

Private msClassName As String
Private msTableName As String
Private msIBNumber As String
Private mudtSafecoRT As udtSafecoRT
Private mcolRTActLog As Collection
Private mcolRTPhotoLog As Collection
Private mcolRTIndem As Collection
Private mcolRTDel As Collection
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
    clsCarUL_ULType = RealTime
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
    clsCarUL_TableName = msTableName
End Property

Public Property Get clsCarUL_IBNumber() As String
    clsCarUL_IBNumber = msIBNumber
End Property

Public Function clsCarUL_PopulateFields(psClaimNo As String, psCheckNum As String, pCurDB As Object, pclsUL As V2ECKeyBoard.clsUpload) As Boolean
    On Error GoTo EH
    Dim lTableCount As Long
    Dim vItem As Variant
    Dim MyActLog As RTSafecoActLog
    Dim MyPhotoLog As RTSafecoPhotoLog
    Dim MyIndem As RTSafecoIndemnity
    Dim MyRTDelAction As V2ECKeyBoard.RTDeleteAction
    Dim sDelAllTables As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    clsCarUL_PopulateFields = True
    
    'BGS 10.31.2001 If we use Dbase tables to store these fields
    'we can't have field names bigger than 10 chars.  So thats why
    'the field name is just the IB Number.  If you need to know what
    'the field refers to have to have a printout This function
    If PopulateUDT(psClaimNo, psCheckNum, pCurDB) Then
    
        With mudtSafecoRT
            'BGS 10.31.2001 update the class with the ibnumber
            msIBNumber = .RT02_sIBNumber
            lTableCount = 1
            pclsUL.AddField "RTSafeco", "UPDATE_RT_TABLE" & ":" & lTableCount, , vbUserDefinedType
            If UCase(left(.RT03_sCheckNUM, 2)) = "IB" Then
                pclsUL.AddField .CI01_sStatus, "CI01_sSt"
                pclsUL.AddField .CI02_sSiteAddress, "CI02_sSi"
                pclsUL.AddField .CI03_sSiteCity, "CI03_sSi"
                pclsUL.AddField .CI04_sSiteState, "CI04_sSi"
                pclsUL.AddField .CI05_sSiteZip, "CI05_sSi"
                pclsUL.AddField .CI06_dtInspectedDate, "CI06_dtIn"          'Date Inspected
                pclsUL.AddField .CI07_dtReceivedDate, "CI07_dtRe"           'Date Received
                pclsUL.AddField .CI08_dtContactedDate, "CI08_dtCo"          'Date contacted
                pclsUL.AddField .CI09_dtAssignedDate, "CI09_dtAs"           'AssignedDate
                pclsUL.AddField .CI10_sTypeOfLoss, "CI10_sTy"               'TypeOfLoss
                pclsUL.AddField .CI11_sPolicyNumber, "CI11_sPo"             'PolicyNumber
                pclsUL.AddField .CI12_sMailAddress, "CI12_sMa"              'MailAddress
                pclsUL.AddField .CI13_sMortgageeName, "CI13_sMo"            'MortgageeName
                pclsUL.AddField .CI14_cBuildingLimits, "CI14_cBu"           'BuildingLimits
                pclsUL.AddField .CI15_cContentsLimits, "CI15_cCo"           'ContentsLimits
                pclsUL.AddField .CI16_cAPSLimits, "CI16_cAP"
                pclsUL.AddField .CI17_cBuildingRCLimits, "CI17_cBu"
                pclsUL.AddField .CI18_cContentsRCLimits, "CI18_cCo"
                pclsUL.AddField .CI19_cAPSRCLimits, "CI19_cAP"
                'IndemInfo
                pclsUL.AddField .II01_cRecovDep, "II01_cRe"
                pclsUL.AddField .II02_cNonRecovDep, "II02_cNo"
            End If
            'Billing Info
            pclsUL.AddField .LID, "LID"                          'Billing ID (CheckID)
            pclsUL.AddField .RT00_lSSN, "RT00_lSSN", , vbLong
            pclsUL.AddField .RT01_sSubToCarrier, "RT01_s"
            pclsUL.AddField .RT02_sIBNumber, "IB_NAME=sIBNumber"
            pclsUL.AddField .RT03_sCheckNUM, "CHECK_NAME=RT03_s"
            pclsUL.AddField .RT04_sCatCode, "RT04_s"
            pclsUL.AddField .RT05_sLocation, "RT05_s"
            pclsUL.AddField .RT05a_sState, "RT05a_s"
            pclsUL.AddField .RT06_dtDateClosed, "RT06_dt"
            pclsUL.AddField .RT07_sAdjusterName, "RT07_s"
            pclsUL.AddField .RT08_sAdjCRID, "RT08_s"
            pclsUL.AddField .RT09_sSALN, "CLIENT_CLAIM_NAME=sClientCla"
            pclsUL.AddField .RT09a_sPolicyNumber, "RT09a_s"
            pclsUL.AddField .RT10_sInsuredName, "RT10_s"
            pclsUL.AddField .RT11_sLossLocation, "RT11_s"
            pclsUL.AddField .RT12_dtDateOfLoss, "RT12_dt"
            pclsUL.AddField .RT13_cGrossLoss, "RT13_c"
            pclsUL.AddField .RT14_cDepreciation, "RT14_c"
            pclsUL.AddField .RT14a_sSupplement, "RT14a_s"
            pclsUL.AddField .RT14b_sRebilled, "RT14b_s"
            pclsUL.AddField .RT15_cDeductible, "RT15_c"
            pclsUL.AddField .RT15a_cLessExcessLimits, "RT15a_c"
            pclsUL.AddField .RT15b_sExcessLimDesc, "RT15b_s"
            pclsUL.AddField .RT16_cNetClaim, "RT16_c"
            pclsUL.AddField .RT17_cServiceFee, "RT17_c"
            pclsUL.AddField .RT17a_cMiscServiceFee, "RT17a_c"
            pclsUL.AddField .RT18_sServiceFeeComment, "RT18_s"
            pclsUL.AddField .RT18a_sMiscServiceFeeComment, "RT18a_s"
            pclsUL.AddField .RT19_cOutBuildingsFee, "RT19_c"
            pclsUL.AddField .RT19a_iOutBuildingsCount, "RT19a_i"
            pclsUL.AddField .RT19b_cOutBuildingsAmount, "RT19b_c"
            pclsUL.AddField .RT20_sOutBuildingsFeeComment, "RT20_s"
            pclsUL.AddField .RT21_cTwoStoryCharge, "RT21_c"
            pclsUL.AddField .RT22_cSteepCharge, "RT22_c"
            pclsUL.AddField .RT23_cInteriorDamageCharge, "RT23_c"
            pclsUL.AddField .RT24_cExternalDamageBGCharge, "RT24_c"
            pclsUL.AddField .RT25_cServiceFeeSubTotal, "RT25_c"
            pclsUL.AddField .RT26_iPhotoCount, "RT26_i"
            pclsUL.AddField .RT27_cPhotoFee, "RT27_c"
            pclsUL.AddField .RT28_iOther, "RT28_i"
            pclsUL.AddField .RT29_cOtherFee, "RT29_c"
            pclsUL.AddField .RT29a_sMiscExpenseFeeComment, "RT29a_s"
            pclsUL.AddField .RT29b_cMiscExpenseFee, "RT29b_c"
            pclsUL.AddField .RT30_cTotalExpenses, "RT30_c"
            pclsUL.AddField .RT31_dTaxPercent, "RT31_d"
            pclsUL.AddField .RT32_cTaxAmount, "RT32_c"
            pclsUL.AddField .RT33_cTotalAdjustingFee, "RT33_c"
            pclsUL.AddField .RT33a_sAccountCode, "RT33a_s"
            pclsUL.AddField .RT34_sPaymentForProperty, "RT34_s"
            pclsUL.AddField .RT35_sPaymentForAuto, "RT35_s"
            pclsUL.AddField .RT36_sPaymentForFinal, "RT36_s"
            pclsUL.AddField .RT37_sPaymentForPartial, "RT37_s"
            pclsUL.AddField .RT38_sPaymentIsAddSupplementY, "RT38_s"
            pclsUL.AddField .RT39_sPaymentIsAddSupplementN, "RT39_s"
            pclsUL.AddField .RT40_sPaymentIsAssociatedY, "RT40_s"
            pclsUL.AddField .RT41_sPaymentISAssociatedN, "RT41_s"
            pclsUL.AddField .RT42_sClassOfClaim, "RT42_s"
            pclsUL.AddField .RT43_sCauseOfLoss, "RT43_s"
            pclsUL.AddField .RT44_sTexasSubCovCode, "RT44_s"
            pclsUL.AddField .RT45_sTexasSuffix, "RT45_s"
            pclsUL.AddField .RT46_cTexasRoofDepreciation, "RT46_c"
            pclsUL.AddField .RT47_sTypeOfPropLossBuilding, "RT47_s"
            pclsUL.AddField .RT48_sTypeOfPropLossContents, "RT48_s"
            pclsUL.AddField .RT49_sTypeOfPropLossALE, "RT49_s"
            pclsUL.AddField .RT49a_sOtherPropLoss, "RT49a_s"
            pclsUL.AddField .RT50_sInsuredPayeeName, "RT50_s"
            pclsUL.AddField .RT51_sPayeeNames, "RT51_s"
            pclsUL.AddField .RT52_sAddress, "RT52_s"
            pclsUL.AddField .RT53_cAmountOfCheck, "RT53_c"
            pclsUL.AddField .RT54_sCatCode, "RT54_s"
            pclsUL.AddField .RT55_sFieldHandled, "RT55_s"
            pclsUL.AddField .RT56_sTotalLoss, "RT56_s"
            pclsUL.AddField .RT57_sCashInLieu, "RT57_s"
            pclsUL.AddField .RT58_sOwnerRetainSalvage, "RT58_s"
            pclsUL.AddField .RT59_sSub, "RT59_s"
            pclsUL.AddField .RT60_sSalvage, "RT60_s"
            pclsUL.AddField .RT61_sInstructions, "RT61_s"
            pclsUL.AddField .RT62_sRequestedBy, "RT62_s"
            pclsUL.AddField .RT63_dtDate, "RT63_dt"
            pclsUL.AddField .RT64_sApproveBy, "RT64_s"
            pclsUL.AddField .RT65_dtApproveDate, "RT65_dt"
            pclsUL.AddField .RT66_sIssuedBy, "RT66_s"
            pclsUL.AddField .RT67_sRetrievedBy, "RT67_s"
            pclsUL.AddField "RTSafeco", "END_OF_RT_TABLE" & ":" & lTableCount, , vbUserDefinedType
            'Activity Log (Stored in Table RTSafecoActLog)
            'Need to flag the following so clsBatches will know where to
            'get them when processing on the server side:
            'IB_NAME=
            'CLIENT_CLAIM_NAME=
            If Not mcolRTActLog Is Nothing Then
                If mcolRTActLog.Count > 0 Then
                    For Each vItem In mcolRTActLog
                        lTableCount = lTableCount + 1
                        pclsUL.AddField "RTSafecoActLog", "UPDATE_RT_TABLE" & ":" & lTableCount, , vbUserDefinedType
                        MyActLog = vItem
                        With MyActLog
                            pclsUL.AddField .dServiceTime, "dServiceTi" & ":" & lTableCount
                            pclsUL.AddField .dtActDate, "dtActDate" & ":" & lTableCount
                            pclsUL.AddField .lActID, "lID" & ":" & lTableCount
                            pclsUL.AddField .sActText, "sActText" & ":" & lTableCount
                            pclsUL.AddField .sActTime, "sActTime" & ":" & lTableCount
                            pclsUL.AddField .sAssociation, "sAssociati" & ":" & lTableCount
                            pclsUL.AddField .sClientClaimNumber, "CLIENT_CLAIM_NAME=sClientCla" & ":" & lTableCount
                            pclsUL.AddField .sIBNumber, "IB_NAME=sIBNumber" & ":" & lTableCount
                        End With
                        pclsUL.AddField "RTSafecoActLog", "END_OF_RT_TABLE" & ":" & lTableCount, , vbUserDefinedType
                    Next
                    Set mcolRTActLog = Nothing
                End If
            End If
            'Photo Log (Stored in Table RTSafecoPhotoLog)
            'Need to flag the following so clsBatches will know where to
            'get them when processing on the server side:
            'IB_NAME=
            'CLIENT_CLAIM_NAME=
            If Not mcolRTPhotoLog Is Nothing Then
                If mcolRTPhotoLog.Count > 0 Then
                    For Each vItem In mcolRTPhotoLog
                        lTableCount = lTableCount + 1
                        pclsUL.AddField "RTSafecoPhotoLog", "UPDATE_RT_TABLE" & ":" & lTableCount, , vbUserDefinedType
                        MyPhotoLog = vItem
                        With MyPhotoLog
                            pclsUL.AddField .dtPhotoDate, "dtPhotoDat" & ":" & lTableCount
                            pclsUL.AddField .lPhotoID, "lID" & ":" & lTableCount
                            pclsUL.AddField .lSortOrder, "lSortOrder" & ":" & lTableCount
                            pclsUL.AddField .sClientClaimNumber, "CLIENT_CLAIM_NAME=sClientCla" & ":" & lTableCount
                            pclsUL.AddField .sDescription, "sDescripti" & ":" & lTableCount
                            pclsUL.AddField .sIBNumber, "IB_NAME=sIBNumber" & ":" & lTableCount
                            pclsUL.AddField .sPhotoName, "sPhotoName" & ":" & lTableCount
                        End With
                        pclsUL.AddField "RTSafecoPhotoLog", "END_OF_RT_TABLE" & ":" & lTableCount, , vbUserDefinedType
                    Next
                    Set mcolRTPhotoLog = Nothing
                End If
            End If
            'Indemnity (Stored in Table RTSafecoIndem)
            'Need to flag the following so clsBatches will know where to
            'get them when processing on the server side:
            'IB_NAME=
            'CLIENT_CLAIM_NAME=
            If Not mcolRTIndem Is Nothing Then
                If mcolRTIndem.Count > 0 Then
                    For Each vItem In mcolRTIndem
                        lTableCount = lTableCount + 1
                        pclsUL.AddField "RTSafecoIndem", "UPDATE_RT_TABLE" & ":" & lTableCount, , vbUserDefinedType
                        MyIndem = vItem
                        With MyIndem
                            pclsUL.AddField .cACVClaim, "cACVClaim" & ":" & lTableCount
                            pclsUL.AddField .cACVLessExcessLimit, "cACVLessEx" & ":" & lTableCount
                            pclsUL.AddField .cExcessLimit, "cExcessLim" & ":" & lTableCount
                            pclsUL.AddField .cNonRecoverableDepreciation, "cNonRecove" & ":" & lTableCount
                            pclsUL.AddField .cRecoverableDepreciation, "cRecoverab" & ":" & lTableCount
                            pclsUL.AddField .cReplacementCost, "cReplaceme" & ":" & lTableCount
                            pclsUL.AddField .lIndemID, "lID" & ":" & lTableCount
                            pclsUL.AddField .sCauseOfLoss, "sCauseOfLo" & ":" & lTableCount
                            pclsUL.AddField .sClass, "sClass" & ":" & lTableCount
                            pclsUL.AddField .sClientClaimNumber, "CLIENT_CLAIM_NAME=sClientCla" & ":" & lTableCount
                            pclsUL.AddField .sDescription, "sDescripti" & ":" & lTableCount
                            pclsUL.AddField .sIBNumber, "IB_NAME=sIBNumber" & ":" & lTableCount
                            pclsUL.AddField .sIndemType, "sIndemType" & ":" & lTableCount
                        End With
                        pclsUL.AddField "RTSafecoIndem", "END_OF_RT_TABLE" & ":" & lTableCount, , vbUserDefinedType
                    Next
                    Set mcolRTIndem = Nothing
                End If
            End If
            'Delete Action
            'Need to flag the following so clsBatches will know where to
            'get them when processing on the server side:
            'IB_NAME=
            'CLIENT_CLAIM_NAME=
            'TABLE_NAME=
            If Not mcolRTDel Is Nothing Then
                If mcolRTDel.Count > 0 Then
                    For Each vItem In mcolRTDel
                        lTableCount = lTableCount + 1
                        pclsUL.AddField "RTSafeco_DeleteAction", "DELETE_FROM_RT_TABLE" & ":" & lTableCount, , vbUserDefinedType
                        MyRTDelAction = vItem
                        With MyRTDelAction
                            pclsUL.AddField .ID, "ID" & ":" & lTableCount
                            pclsUL.AddField .sClientClaimNumber, "CLIENT_CLAIM_NAME=sClientCla" & ":" & lTableCount
                            pclsUL.AddField .sIBNumber, "IB_NAME=sIBNumber" & ":" & lTableCount
                            If InStr(1, .sTableName, "(DELETE_CLAIM)", vbTextCompare) > 0 Then
                                sDelAllTables = "RTSafeco,RTSafecoActLog,RTSafecoPhotoLog,RTSafecoIndem"
                                pclsUL.AddField .sTableName & "RT:" & sDelAllTables, "TABLE_NAME=sTableName" & ":" & lTableCount
                            Else
                                pclsUL.AddField "RTSafeco" & .sTableName, "TABLE_NAME=sTableName" & ":" & lTableCount
                            End If
                        End With
                        pclsUL.AddField "RTSafeco_DeleteAction", "END_OF_RT_TABLE" & ":" & lTableCount, , vbUserDefinedType
                    Next
                    Set mcolRTPhotoLog = Nothing
                End If
            End If
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
    Dim udtBatchesRT As V2ECKeyBoard.udtBatchesRT
    Dim oField As V2ECKeyBoard.clsField
    Dim sCheckNUM As String
    'Issue 244  9.10.2002 Production Report should indicate Status Deleted
    'Since Delete Actions are attached to other RT Claim Updates need
    'to extrapulate Delete action and at the same time store main claim info
    'in temp variables. After firing a db Update Event we can place the
    'Temp variables back into Batches UDT
    Dim bDeleteClaim As Boolean         'This is true if there is an attached Delete Action for another Claim Number (Different from the actual Claim Number)
    Dim udtDelAction As V2ECKeyBoard.RTDeleteAction
    Dim vDelAction As Variant
    Dim colDelAction As Collection
    'Error message
    Dim sMess As String
    
    clsCarUL_UpdateBatchesUDT = True
    With udtBatchesRT
        'Init this value
        .cMiscCharges = 0
        For Each oField In poUL.Fields
            Select Case UCase(oField.FieldName)
                Case "CI01_SST"
                    .sStatus = oField.Value
                Case "CI02_SSI"
                    .sSiteAddress = oField.Value
                Case "CI03_SSI"
                    .sSiteCity = oField.Value
                Case "CI04_SSI"
                    .sSiteState = oField.Value
                Case "CI05_SSI"
                    .sSiteZip = oField.Value
                Case "CI06_DTIN"
                    .dtInspectedDate = oField.Value
                Case "CI07_DTRE"
                    .dtReceivedDate = oField.Value
                Case "CI08_DTCO"
                    .dtContactedDate = oField.Value
                Case "CI09_DTAS"
                    .dtAssignedDate = oField.Value
                Case "CI10_STY"
                    .sTypeOfLoss = oField.Value
                Case "CI11_SPO"
                    .sPolicyNumber = oField.Value
                Case "CI12_SMA"
                    .sMailAddress = oField.Value
                Case "RT11_S"
                    .sLossLocation = oField.Value
                Case "CI13_SMO"
                    .sMortgageeName = oField.Value
                Case "CI14_CBU"
                    .cBuildingLimits = oField.Value
                Case "CI15_CCO"
                    .cContentsLimits = oField.Value
                Case "CI16_CAP"
                    .cAPSLimits = oField.Value
                Case "CI17_CBU"
                    .cBuildingRCLimits = oField.Value
                Case "CI18_CCO"
                    .cContentsRCLimits = oField.Value
                Case "CI19_CAP"
                    .cAPSRCLimits = oField.Value
                Case "RT15_C"
                    .sDeductibles = oField.Value
                Case "RT00_LSSN" '11.21.2001 SSN used to match up Cat_no / Site
                    .lSSN = oField.Value
                Case "RT13_C" '_cGrossLoss"
                    .cGrossLoss = oField.Value
                Case "RT32_C" '_cTaxAmount"
                    .cTaxesTotal = oField.Value
                Case "RT33_C" '_cTotalAdjustingFee"
                    .cTotalFee = oField.Value
                Case "RT25_C" '_cServiceFeeSubTotal"
                    .cTotalServiceFees = oField.Value
                Case "RT63_DT" '_dtDate"
                    .dtDate = oField.Value
                Case "RT06_DT" '_dtDateClosed"
                    .dtDateClosed = CDate(oField.Value)
                Case "RT12_DT" '_dtDateOfLoss"
                    .dtDateOfLoss = oField.Value
                Case "RT08_S" '_sAdjCRID"
                    .sAdjuster_I = oField.Value
                Case "RT07_S" '_sAdjusterName"
                    .sAdjuster_N = oField.Value
                Case "RT04_S" '_sCatCode"
                    .sCATCode = oField.Value
                Case "CLIENT_CLAIM_NAME=SCLIENTCLA" '_sSALN"
                    .sClaimNumber = oField.Value
                Case "IB_NAME=SIBNUMBER" '_sIBNumber"
                    .sIBNumber = oField.Value
                Case "RT10_S" '_sInsuredName"
                    .sInsuredName = oField.Value
                Case "RT05_S" '_sLocation"
                    .sLossLocationCity = oField.Value
                Case "RT05A_S" '_sState"
                    .sLossLocationState = oField.Value
                Case "CHECK_NAME=RT03_S" '_sCheckNUM"
                    sCheckNUM = oField.Value
                Case "RT30_C" 'RT30_cTotalExpenses, "RT30_c"
                    .cMiscCharges = oField.Value
            End Select
            
            'Issue 244  9.10.2002 Production Report should indicate Status Deleted
            Select Case UCase(left(oField.FieldName, InStrRev(oField.FieldName, ":", , vbBinaryCompare)))
                Case "ID:"
                    If oField.Value = "-999" Then
                        'Set this flag true
                        bDeleteClaim = True
                    End If
                Case "CLIENT_CLAIM_NAME=SCLIENTCLA:"
                    If bDeleteClaim Then
                        udtDelAction.sClientClaimNumber = oField.Value
                    End If
                Case "IB_NAME=SIBNUMBER:"
                    If bDeleteClaim Then
                        udtDelAction.sIBNumber = oField.Value
                    End If
            End Select
            If bDeleteClaim Then
                If udtDelAction.sIBNumber <> vbNullString Then
                    If colDelAction Is Nothing Then
                        Set colDelAction = New Collection
                    End If
                    colDelAction.Add udtDelAction, udtDelAction.sIBNumber
                    'Reset for next Deleteaction if Multiple
                    bDeleteClaim = False
                    udtDelAction.sClientClaimNumber = vbNullString
                    udtDelAction.sIBNumber = vbNullString
                End If
            End If
        Next
        .sCarrierCode = "SC"
        'Set the DBUpdate flag
        If UCase(left(sCheckNUM, 2)) = "IB" Then
            mbDoDBUpdate = True
        Else
            mbDoDBUpdate = False
        End If
        .sRTID = GetRTID
        .cAdministrativeFee = 0
        
    End With
    
    poBatches.BatchesRT = udtBatchesRT
    
    'This will update the RT Tables (Approach)
    If Not poBatches.UpdateAllRTTables(Me, poUL) Then
        clsCarUL_UpdateBatchesUDT = False
        Exit Function
    Else
        'Issue 244  9.10.2002 Production Report should indicate Status Deleted
        If Not colDelAction Is Nothing Then
            For Each vDelAction In colDelAction
                udtDelAction = vDelAction
                With udtBatchesRT
                    .sIBNumber = udtDelAction.sIBNumber
                    .sClaimNumber = udtDelAction.sClientClaimNumber
                    .sStatus = "DELETED"
                End With
                poUL.FireUpdateDBRT udtBatchesRT, poBatches, Me
            Next
            'Clean up
             Set colDelAction = Nothing
        End If
    End If
    
    Exit Function
EH:
    clsCarUL_UpdateBatchesUDT = False
    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
    sMess = sMess & "ERROR # " & Err.Number & vbCrLf
    sMess = sMess & Err.Description & vbCrLf
    sMess = sMess & clsCarUL_ClassName & " Public Function UpdateBatchesUDT" & vbCrLf
    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
    poUL.FireErrorMess sMess
End Function

Public Sub clsCarUL_UpdateDB(poBatches As V2ECKeyBoard.clsBatches, poUL As V2ECKeyBoard.clsUpload)
    On Error GoTo EH
    Dim sMess As String
    If mbDoDBUpdate Then
        poUL.FireUpdateDBRT poBatches.BatchesRT, poBatches, Me
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
    On Error GoTo EH
    Dim sMess As String
    Dim sSQL As String
    Dim DCRS As ADODB.Recordset  'Disconnected recordset
    Dim cFee As Currency
    Dim cLastFee As Currency
    Dim cTotalFee As Currency
    Dim sIB As String
    
    'Can only Get the Total Fee from the IB Table Since this
    'Is where the Final Fee is tabulated. This means that the Total
    'fee will remain 0 until the Claim is actualy Closed (IE IB uploaded)
    
    '10.3.2002 Added order by checknum
    sSQL = "SELECT "
    sSQL = sSQL & "IB33_C AS FEE, "
    sSQL = sSQL & "RIGHT(RTRIM(IB02_S), 6) AS IB "
    sSQL = sSQL & "FROM IBSafeco "
    sSQL = sSQL & "WHERE IB09_S = '" & goUtil.utCleanSQLString(psClaimNumber) & "' "
    sSQL = sSQL & "AND LEFT(IB02_S, " & Len(psIBNumber) & ") = '" & goUtil.utCleanSQLString(psIBNumber) & "' "
    sSQL = sSQL & "ORDER BY IB03_S "
    
    Set DCRS = poBatch.OpenDCRS(sSQL, WebControlDB)
    
    If Not DCRS.EOF Then
        DCRS.MoveFirst
        Do Until DCRS.EOF
            sIB = IIf(IsNull(DCRS!IB), vbNullString, DCRS!IB)
            cFee = IIf(IsNull(DCRS!FEE), 0, DCRS!FEE)
            If InStr(1, sIB, "R", vbTextCompare) > 0 Then
                cTotalFee = cTotalFee - cLastFee
                cTotalFee = cTotalFee + cFee
            Else
                cTotalFee = cTotalFee + cFee
            End If
            cLastFee = cFee
            DCRS.MoveNext
        Loop
    End If
    
    clsCarUL_GetTotalFee = cTotalFee
    'Cleanup
    DCRS.Close
    Set DCRS = Nothing
    
    Exit Function
EH:
    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
    sMess = sMess & "ERROR # " & Err.Number & vbCrLf
    sMess = sMess & Err.Description & vbCrLf
    sMess = sMess & clsCarUL_ClassName & " Public Function GetRTTotalFee" & vbCrLf
    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
    poUL.FireErrorMess sMess
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
    sSQL = sSQL & "FROM Assignments As A LEFT JOIN Checks As B "
    sSQL = sSQL & "ON A.ClaimNo = B.ClaimNo "
    sSQL = sSQL & "WHERE A.ClaimNo = '" & goUtil.utCleanSQLString(psClaimNo) & "' "
    sSQL = sSQL & "AND B.CheckNum = '" & goUtil.utCleanSQLString(psCheckNum) & "' "
    
    Set RS = pCurDB.OpenRecordset(sSQL)
    
    With mudtSafecoRT
        If Not RS.EOF Then
            RS.MoveFirst
            'Claim Info
            .CI01_sStatus = IIf(IsNull(RS!CI01_sStatus), vbNullString, RS!CI01_sStatus)
            .CI02_sSiteAddress = IIf(IsNull(RS!CI02_sSiteAddress), vbNullString, RS!CI02_sSiteAddress)
            .CI03_sSiteCity = IIf(IsNull(RS!CI03_sSiteCity), vbNullString, RS!CI03_sSiteCity)
            .CI04_sSiteState = IIf(IsNull(RS!CI04_sSiteState), vbNullString, RS!CI04_sSiteState)
            .CI05_sSiteZip = IIf(IsNull(RS!CI05_sSiteZip), vbNullString, RS!CI05_sSiteZip)
            .CI06_dtInspectedDate = IIf(IsNull(RS!CI06_dtInspectedDate) Or RS!CI06_dtInspectedDate = "//", NULL_DATE, RS!CI06_dtInspectedDate)           'Date Inspected
            .CI07_dtReceivedDate = IIf(IsNull(RS!CI07_dtReceivedDate) Or RS!CI07_dtReceivedDate = "//", NULL_DATE, RS!CI07_dtReceivedDate)             'Date Received
            .CI08_dtContactedDate = IIf(IsNull(RS!CI08_dtContactedDate) Or RS!CI08_dtContactedDate = "//", NULL_DATE, RS!CI08_dtContactedDate)            'Date contacted
            .CI09_dtAssignedDate = IIf(IsNull(RS!CI09_dtAssignedDate) Or RS!CI09_dtAssignedDate = "//", NULL_DATE, RS!CI09_dtAssignedDate)             'AssignedDate
            .CI10_sTypeOfLoss = IIf(IsNull(RS!CI10_sTypeOfLoss), vbNullString, RS!CI10_sTypeOfLoss)               'TypeOfLoss
            .CI11_sPolicyNumber = IIf(IsNull(RS!CI11_sPolicyNumber), vbNullString, RS!CI11_sPolicyNumber)             'PolicyNumber
            .CI12_sMailAddress = IIf(IsNull(RS!CI12_sMailAddress), vbNullString, RS!CI12_sMailAddress)              'MailAddress
            .CI13_sMortgageeName = IIf(IsNull(RS!CI13_sMortgageeName), vbNullString, RS!CI13_sMortgageeName)            'MortgageeName
            .CI14_cBuildingLimits = IIf(IsNull(RS!CI14_cBuildingLimits), 0, RS!CI14_cBuildingLimits)          'BuildingLimits
            .CI15_cContentsLimits = IIf(IsNull(RS!CI15_cContentsLimits), 0, RS!CI15_cContentsLimits)        'ContentsLimits
            .CI16_cAPSLimits = IIf(IsNull(RS!CI16_cAPSLimits), 0, RS!CI16_cAPSLimits)
            .CI17_cBuildingRCLimits = IIf(IsNull(RS!CI17_cBuildingRCLimits), 0, RS!CI17_cBuildingRCLimits)
            .CI18_cContentsRCLimits = IIf(IsNull(RS!CI18_cContentsRCLimits), 0, RS!CI18_cContentsRCLimits)
            .CI19_cAPSRCLimits = IIf(IsNull(RS!CI19_cAPSRCLimits), 0, RS!CI19_cAPSRCLimits)
            'IndemInfo
            .II01_cRecovDep = IIf(IsNull(RS!II01_cRecovDep), 0, RS!II01_cRecovDep)
            .II02_cNonRecovDep = IIf(IsNull(RS!II02_cNonRecovDep), 0, RS!II02_cNonRecovDep)
            'Billing Info
            .LID = IIf(IsNull(RS!LID), 0, RS!LID)
            .RT00_lSSN = GetSetting("ECS", "WEB_SECURITY", "USERNAME", NO_SSN)
            'BGS 10.11.2001 Put in the Carrier Header Here
            .RT01_sSubToCarrier = "Safeco"
            .RT02_sIBNumber = IIf(IsNull(RS!RT02_sIBNumber), vbNullString, RS!RT02_sIBNumber)
            .RT03_sCheckNUM = IIf(IsNull(RS!RT03_sCheckNUM), vbNullString, RS!RT03_sCheckNUM)
            'BGS 5.13.2002 Ony use first 2 Chars of Cat Code
            .RT04_sCatCode = IIf(IsNull(RS!RT04_sCatCode), vbNullString, left(RS!RT04_sCatCode, 2))
            .RT05_sLocation = IIf(IsNull(RS!RT05_sLocation), vbNullString, RS!RT05_sLocation)
            .RT05a_sState = IIf(IsNull(RS!RT05a_sState), vbNullString, RS!RT05a_sState)
            .RT06_dtDateClosed = IIf(IsNull(RS!RT06_dtDateClosed) Or RS!RT06_dtDateClosed = "//", NULL_DATE, RS!RT06_dtDateClosed)
            .RT07_sAdjusterName = IIf(IsNull(RS!RT07_sAdjusterName), vbNullString, RS!RT07_sAdjusterName)
            .RT08_sAdjCRID = IIf(IsNull(RS!RT08_sAdjCRID), vbNullString, RS!RT08_sAdjCRID)
            .RT09_sSALN = IIf(IsNull(RS!RT09_sSALN), vbNullString, RS!RT09_sSALN)
            .RT09a_sPolicyNumber = IIf(IsNull(RS!RT09a_sPolicyNumber), vbNullString, RS!RT09a_sPolicyNumber)
            .RT10_sInsuredName = IIf(IsNull(RS!RT10_sInsuredName), vbNullString, RS!RT10_sInsuredName)
            .RT11_sLossLocation = IIf(IsNull(RS!RT11_sLossLocation), vbNullString, RS!RT11_sLossLocation)
            .RT12_dtDateOfLoss = IIf(IsNull(RS!RT12_dtDateOfLoss) Or RS!RT12_dtDateOfLoss = "//", NULL_DATE, RS!RT12_dtDateOfLoss)
            .RT13_cGrossLoss = IIf(IsNull(RS!RT13_cGrossLoss), 0, RS!RT13_cGrossLoss)
            .RT14_cDepreciation = IIf(IsNull(RS!RT14_cDepreciation), 0, RS!RT14_cDepreciation)
            .RT14a_sSupplement = IIf(IsNull(RS!RT14a_sSupplement), vbNullString, RS!RT14a_sSupplement)
            .RT14b_sRebilled = IIf(IsNull(RS!RT14b_sRebilled), vbNullString, RS!RT14b_sRebilled)
            .RT15_cDeductible = IIf(IsNull(RS!RT15_cDeductible), 0, RS!RT15_cDeductible)
            .RT15a_cLessExcessLimits = IIf(IsNull(RS!RT15a_cLessExcessLimits), 0, RS!RT15a_cLessExcessLimits)
            .RT15b_sExcessLimDesc = IIf(IsNull(RS!RT15b_sExcessLimDesc), vbNullString, RS!RT15b_sExcessLimDesc)
            .RT16_cNetClaim = IIf(IsNull(RS!RT16_cNetClaim), 0, RS!RT16_cNetClaim)
            .RT17_cServiceFee = IIf(IsNull(RS!RT17_cServiceFee), 0, RS!RT17_cServiceFee)
            .RT17a_cMiscServiceFee = IIf(IsNull(RS!RT17a_cMiscServiceFee), 0, RS!RT17a_cMiscServiceFee)
            .RT18_sServiceFeeComment = IIf(IsNull(RS!RT18_sServiceFeeComment), vbNullString, RS!RT18_sServiceFeeComment)
            .RT18a_sMiscServiceFeeComment = IIf(IsNull(RS!RT18a_sMiscServiceFeeComment), vbNullString, RS!RT18a_sMiscServiceFeeComment)
            .RT19_cOutBuildingsFee = IIf(IsNull(RS!RT19_cOutBuildingsFee), 0, RS!RT19_cOutBuildingsFee)
            .RT19a_iOutBuildingsCount = IIf(IsNull(RS!RT19a_iOutBuildingsCount), 0, RS!RT19a_iOutBuildingsCount)
            .RT19b_cOutBuildingsAmount = GetSetting("EBERLSSC", "FEES", "OUT_BUILDING_CHARGE", "20.00")
            .RT20_sOutBuildingsFeeComment = IIf(IsNull(RS!RT20_sOutBuildingsFeeComment), vbNullString, RS!RT20_sOutBuildingsFeeComment)
            .RT21_cTwoStoryCharge = IIf(IsNull(RS!RT21_cTwoStoryCharge), 0, RS!RT21_cTwoStoryCharge)
            .RT22_cSteepCharge = IIf(IsNull(RS!RT22_cSteepCharge), 0, RS!RT22_cSteepCharge)
            .RT23_cInteriorDamageCharge = IIf(IsNull(RS!RT23_cInteriorDamageCharge), 0, RS!RT23_cInteriorDamageCharge)
            .RT24_cExternalDamageBGCharge = IIf(IsNull(RS!RT24_cExternalDamageBGCharge), 0, RS!RT24_cExternalDamageBGCharge)
            .RT25_cServiceFeeSubTotal = IIf(IsNull(RS!RT25_cServiceFeeSubTotal), 0, RS!RT25_cServiceFeeSubTotal)
            .RT26_iPhotoCount = IIf(IsNull(RS!RT26_iPhotoCount), 0, RS!RT26_iPhotoCount)
            .RT27_cPhotoFee = IIf(IsNull(RS!RT27_cPhotoFee), 0, RS!RT27_cPhotoFee)
            .RT28_iOther = IIf(IsNull(RS!RT28_iOther), 0, RS!RT28_iOther)
            .RT29_cOtherFee = IIf(IsNull(RS!RT29_cOtherFee), 0, RS!RT29_cOtherFee)
            .RT29a_sMiscExpenseFeeComment = IIf(IsNull(RS!RT29a_sMiscExpenseFeeComment), vbNullString, RS!RT29a_sMiscExpenseFeeComment)
            .RT29b_cMiscExpenseFee = IIf(IsNull(RS!RT29b_cMiscExpenseFee), 0, RS!RT29b_cMiscExpenseFee)
            .RT30_cTotalExpenses = IIf(IsNull(RS!RT30_cTotalExpenses), 0, RS!RT30_cTotalExpenses)
            .RT31_dTaxPercent = IIf(IsNull(RS!RT31_dTaxPercent), 0, RS!RT31_dTaxPercent)
            .RT32_cTaxAmount = IIf(IsNull(RS!RT32_cTaxAmount), 0, RS!RT32_cTaxAmount)
            .RT33_cTotalAdjustingFee = IIf(IsNull(RS!RT33_cTotalAdjustingFee), 0, RS!RT33_cTotalAdjustingFee)
            .RT33a_sAccountCode = GetSetting("EBERLSSC", "FEES", "ACCOUNT_CODE", " ")
            .RT34_sPaymentForProperty = IIf(IsNull(RS!RT34_sPaymentForProperty), vbNullString, RS!RT34_sPaymentForProperty)
            .RT35_sPaymentForAuto = IIf(IsNull(RS!RT35_sPaymentForAuto), vbNullString, RS!RT35_sPaymentForAuto)
            .RT36_sPaymentForFinal = IIf(IsNull(RS!RT36_sPaymentForFinal), vbNullString, RS!RT36_sPaymentForFinal)
            .RT37_sPaymentForPartial = IIf(IsNull(RS!RT37_sPaymentForPartial), vbNullString, RS!RT37_sPaymentForPartial)
            .RT38_sPaymentIsAddSupplementY = IIf(IsNull(RS!RT38_sPaymentIsAddSupplementY), vbNullString, RS!RT38_sPaymentIsAddSupplementY)
            .RT39_sPaymentIsAddSupplementN = IIf(IsNull(RS!RT39_sPaymentIsAddSupplementN), vbNullString, RS!RT39_sPaymentIsAddSupplementN)
            .RT40_sPaymentIsAssociatedY = IIf(IsNull(RS!RT40_sPaymentIsAssociatedY), vbNullString, RS!RT40_sPaymentIsAssociatedY)
            .RT41_sPaymentISAssociatedN = IIf(IsNull(RS!RT41_sPaymentISAssociatedN), vbNullString, RS!RT41_sPaymentISAssociatedN)
            .RT42_sClassOfClaim = IIf(IsNull(RS!RT42_sClassOfClaim), vbNullString, RS!RT42_sClassOfClaim)
            .RT43_sCauseOfLoss = IIf(IsNull(RS!RT43_sCauseOfLoss), vbNullString, RS!RT43_sCauseOfLoss)
            .RT45_sTexasSuffix = IIf(IsNull(RS!RT45_sTexasSuffix), vbNullString, RS!RT45_sTexasSuffix)
            .RT46_cTexasRoofDepreciation = IIf(IsNull(RS!RT46_cTexasRoofDepreciation), vbNullString, RS!RT46_cTexasRoofDepreciation)
            .RT47_sTypeOfPropLossBuilding = IIf(IsNull(RS!RT47_sTypeOfPropLossBuilding), vbNullString, RS!RT47_sTypeOfPropLossBuilding)
            .RT48_sTypeOfPropLossContents = IIf(IsNull(RS!RT48_sTypeOfPropLossContents), vbNullString, RS!RT48_sTypeOfPropLossContents)
            .RT49_sTypeOfPropLossALE = IIf(IsNull(RS!RT49_sTypeOfPropLossALE), vbNullString, RS!RT49_sTypeOfPropLossALE)
            .RT49a_sOtherPropLoss = vbNullString    'Future use in vs2.0
            .RT50_sInsuredPayeeName = IIf(IsNull(RS!RT50_sInsuredPayeeName), vbNullString, RS!RT50_sInsuredPayeeName)
            .RT51_sPayeeNames = IIf(IsNull(RS!RT51_sPayeeNames), vbNullString, RS!RT51_sPayeeNames)
            .RT52_sAddress = IIf(IsNull(RS!RT52_sAddress), vbNullString, RS!RT52_sAddress)
            .RT53_cAmountOfCheck = IIf(IsNull(RS!RT53_cAmountOfCheck), 0, RS!RT53_cAmountOfCheck)
            .RT54_sCatCode = IIf(IsNull(RS!RT54_sCatCode), vbNullString, RS!RT54_sCatCode)
            .RT55_sFieldHandled = IIf(IsNull(RS!RT55_sFieldHandled), vbNullString, RS!RT55_sFieldHandled)
            .RT56_sTotalLoss = IIf(IsNull(RS!RT56_sTotalLoss), vbNullString, RS!RT56_sTotalLoss)
            .RT57_sCashInLieu = IIf(IsNull(RS!RT57_sCashInLieu), vbNullString, RS!RT57_sCashInLieu)
            .RT58_sOwnerRetainSalvage = IIf(IsNull(RS!RT58_sOwnerRetainSalvage), vbNullString, RS!RT58_sOwnerRetainSalvage)
            .RT59_sSub = IIf(IsNull(RS!RT59_sSub), vbNullString, RS!RT59_sSub)
            .RT60_sSalvage = IIf(IsNull(RS!RT60_sSalvage), vbNullString, RS!RT60_sSalvage)
            .RT61_sInstructions = IIf(IsNull(RS!RT61_sInstructions), vbNullString, RS!RT61_sInstructions)
            .RT62_sRequestedBy = IIf(IsNull(RS!RT62_sRequestedBy), vbNullString, RS!RT62_sRequestedBy)
            .RT63_dtDate = IIf(IsNull(RS!RT63_dtDate), NULL_DATE, RS!RT63_dtDate)
            .RT64_sApproveBy = vbNullString     'Future use in vs2.0
            .RT65_dtApproveDate = NULL_DATE     'Future use in vs2.0
            .RT66_sIssuedBy = vbNullString      'Future use in vs2.0
            .RT67_sRetrievedBy = vbNullString   'Future use in vs2.0
        End If
    End With
    
    RS.Close
    Set RS = Nothing
    
    'Populate Collections
    PopulateRTActLog psClaimNo, pCurDB
    PopulateRTPhotoLog psClaimNo, pCurDB
    PopulateRTIndem psClaimNo, pCurDB
    PopulateRTDelAction psClaimNo, pCurDB
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    PopulateUDT = False
    Set RS = Nothing
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarUL_ClassName & vbCrLf & " Private Function PopulateUDT"
End Function

Private Function GetClientClaimNo(psClaimNo As String, pCurDB As Object) As String
    On Error GoTo EH
    Dim sSQL As String
    Dim RS As Recordset
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    sSQL = "SELECT A.ClientClaimNo "
    sSQL = sSQL & "FROM " & DB_ASSIGNMENTS & " As A "
    sSQL = sSQL & "WHERE A.ClaimNo = '" & goUtil.utCleanSQLString(psClaimNo) & "' "
    
    Set RS = pCurDB.OpenRecordset(sSQL)
    
    If Not RS.EOF Then
        RS.MoveFirst
        GetClientClaimNo = IIf(IsNull(RS!ClientClaimNo), vbNullString, RS!ClientClaimNo)
    End If
    Set RS = Nothing
    
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarUL_ClassName & vbCrLf & " Private Function GetClientClaimNo"
End Function

Private Function BuildClaimsSQLAB() As String
    On Error GoTo EH
    Dim sSQL As String
    Dim sT As String
    Dim dtDate As Date
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    dtDate = Format(Now, "MM/DD/YY")
    'Claim Info
    sSQL = "A.Status As CI01_sStatus, "
    sSQL = sSQL & "A.SiteAddress As CI02_sSiteAddress, "
    sSQL = sSQL & "A.SiteCity As CI03_sSiteCity, "
    sSQL = sSQL & "A.SiteState As CI04_sSiteState, "
    sSQL = sSQL & "A.SiteZip As CI05_sSiteZip, "
    sSQL = sSQL & "Mid(A.InspectedDate,5,2) & '/' & Mid(A.InspectedDate,7,2) & '/' & Left(A.InspectedDate,4) As CI06_dtInspectedDate, "
    sSQL = sSQL & "Mid(A.ReceivedDate,5,2) & '/' & Mid(A.ReceivedDate,7,2) & '/' & Left(A.ReceivedDate,4) As CI07_dtReceivedDate, "
    sSQL = sSQL & "Mid(A.ContactedDate,5,2) & '/' & Mid(A.ContactedDate,7,2) & '/' & Left(A.ContactedDate,4) As CI08_dtContactedDate, "
    sSQL = sSQL & "Mid(A.AssignedDate,5,2) & '/' & Mid(A.AssignedDate,7,2) & '/' & Left(A.AssignedDate,4) As CI09_dtAssignedDate, "
    sSQL = sSQL & "A.TypeOfLoss As CI10_sTypeOfLoss, "
    sSQL = sSQL & "A.PolicyNumber As CI11_sPolicyNumber, "
    sSQL = sSQL & "A.NewMailingAddress As CI12_sMailAddress, "
    sSQL = sSQL & "A.MortgageCompany As CI13_sMortgageeName, "
    sSQL = sSQL & "A.BuildingPolicyLimit As CI14_cBuildingLimits, "
    sSQL = sSQL & "A.PersonalPropertyLimit As CI15_cContentsLimits, "
    sSQL = sSQL & "A.APSPolicyLimit As CI16_cAPSLimits, "
    sSQL = sSQL & "A.RCSaidBuilding As CI17_cBuildingRCLimits, "
    sSQL = sSQL & "A.RCSaidPersonal As CI18_cContentsRCLimits, "
    sSQL = sSQL & "A.RCSaidAPS As CI19_cAPSRCLimits, "
    'IndemInfo
    sSQL = sSQL & "A.RecoverableDepreciation As II01_cRecovDep, "
    sSQL = sSQL & "A.NonRecovDepr As II02_cNonRecovDep, "
    'BillingInfo
    sSQL = sSQL & "B.ID As LID, "
    sSQL = sSQL & "B.ClaimNo & IIF(B.Supplement > 0, 'S' & B.Supplement & '.', '') & IIF(B.Rebilled > 0, 'R' & B.Rebilled & '.', '') As RT02_sIBNumber, "
    sSQL = sSQL & "B.CheckNUM As RT03_sCheckNUM, "
    sSQL = sSQL & "A.CatCode As RT04_sCatCode, "
    sSQL = sSQL & "A.SiteCity As RT05_sLocation, "
    sSQL = sSQL & "A.SiteState As RT05a_sState, "
    sSQL = sSQL & "Mid(A.ClosedDate,5,2) & '/' & Mid(A.ClosedDate,7,2) & '/' & Left(A.ClosedDate,4) As RT06_dtDateClosed, "
    sSQL = sSQL & "A.AdjustorFirstName & ' ' & Trim(A.AdjustorLastName) As RT07_sAdjusterName, "
    sSQL = sSQL & "A.ClaimRepIDNO As RT08_sAdjCRID, "
    sSQL = sSQL & "A.ClientClaimNo As RT09_sSALN, "
    sSQL = sSQL & "A.PolicyNumber As RT09a_sPolicyNumber, "
    sSQL = sSQL & "A.NewInsuredNames As RT10_sInsuredName, "
    sSQL = sSQL & "A.NewLossLocation As RT11_sLossLocation, "
    sSQL = sSQL & "Mid(A.LossDate,5,2) & '/' & Mid(A.LossDate,7,2) & '/' & Mid(A.LossDate,3,2) As RT12_dtDateOfLoss, "
    sSQL = sSQL & "A.FullCostOfRepair As RT13_cGrossLoss, "
    sSQL = sSQL & "A.RecoverableDepreciation + A.NonRecovDepr As RT14_cDepreciation, "
    sSQL = sSQL & "IIF(B.Supplement > 0, 'X', '') As RT14a_sSupplement, "
    sSQL = sSQL & "IIF(B.Rebilled >0, 'X', '') As RT14b_sRebilled, "
    sSQL = sSQL & "A.Deductible As RT15_cDeductible, "
    sSQL = sSQL & "A.LessExcessLimits As RT15a_cLessExcessLimits, "
    sSQL = sSQL & "B.ExcessLimDesc As RT15b_sExcessLimDesc, "
    sSQL = sSQL & "IIF(A.NetActualCashValueClaim <=0, 0, A.NetActualCashValueClaim) As RT16_cNetClaim, "
    sSQL = sSQL & "A.ServiceFee As RT17_cServiceFee, "
    sSQL = sSQL & "A.MiscServiceFee As RT17a_cMiscServiceFee, "
    sSQL = sSQL & "IIF(A.BillingHours <> '', IIF(CCur(A.BillingHours) > 0, '(' & A.BillingHours & ' hours @ $' & A.ServiceFee / CCur(A.BillingHours) & ')', ''), '') AS RT18_sServiceFeeComment, "
    sSQL = sSQL & "A.MiscServiceFeeComment As RT18a_sMiscServiceFeeComment, "
    sSQL = sSQL & "A.OutBuildingsFee As RT19_cOutBuildingsFee, "
    sSQL = sSQL & "A.OutBuildings AS RT19a_iOutBuildingsCount, "
    sSQL = sSQL & "IIF(A.BillingFenceOnly = '1', '(Fence Only)', '') As RT20_sOutBuildingsFeeComment, "
    sSQL = sSQL & "A.TwoStoryDamage As RT21_cTwoStoryCharge, "
    sSQL = sSQL & "A.SteepCharge As RT22_cSteepCharge, "
    sSQL = sSQL & "A.BillingInteriorDamageDollars As RT23_cInteriorDamageCharge, "
    sSQL = sSQL & "A.BillingExteriorDamageDollars As RT24_cExternalDamageBGCharge, "
    sSQL = sSQL & "A.ServiceFee + A.OutBuildingsFee + A.TwoStoryDamage + A.SteepCharge + A.BillingInteriorDamageDollars + A.BillingExteriorDamageDollars + A.MiscServiceFee As RT25_cServiceFeeSubTotal, "
    sSQL = sSQL & "A.PhotographCount As RT26_iPhotoCount, "
    sSQL = sSQL & "A.PhotographFee As RT27_cPhotoFee, "
    sSQL = sSQL & "A.Miles As RT28_iOther, "
    sSQL = sSQL & "A.MilesFee As RT29_cOtherFee, "
    sSQL = sSQL & "A.MiscFeesComment As RT29a_sMiscExpenseFeeComment, "
    sSQL = sSQL & "A.MiscFees As RT29b_cMiscExpenseFee, "
    sSQL = sSQL & "A.PhotographFee + A.MilesFee + A.MiscFees As RT30_cTotalExpenses, "
    sSQL = sSQL & "A.TaxPercent As RT31_dTaxPercent, "
    sSQL = sSQL & "A.TaxPercentAmount As RT32_cTaxAmount, "
    sSQL = sSQL & "A.InvoiceTotal As RT33_cTotalAdjustingFee, "
    sSQL = sSQL & "IIF(B.ForProperty = True,'X','') As RT34_sPaymentForProperty, "
    sSQL = sSQL & "IIf(B.ForAuto = True,'X','') As RT35_sPaymentForAuto, "
    sSQL = sSQL & "IIF(B.FinalPayment = True,'X','') As RT36_sPaymentForFinal, "
    sSQL = sSQL & "IIF(B.PartialPayment = True,'X','') As RT37_sPaymentForPartial, "
    sSQL = sSQL & "IIF(B.ForAdditionalSupplement = True,'X','') As RT38_sPaymentIsAddSupplementY, "
    sSQL = sSQL & "IIF(B.ForAdditionalSupplement = False,'X','') As RT39_sPaymentIsAddSupplementN, "
    sSQL = sSQL & "IIF(B.PaymentAssociated = True,'X','') As RT40_sPaymentIsAssociatedY, "
    sSQL = sSQL & "IIF(B.PaymentAssociated = False,'X','') As RT41_sPaymentISAssociatedN, "
    sSQL = sSQL & "Left(B.TypeOfLoss,2) As RT42_sClassOfClaim, "
    sSQL = sSQL & "Left(B.CauseOfLoss,2) As RT43_sCauseOfLoss, "
    sSQL = sSQL & "B.TexasSubCode As RT44_sTexasSubCovCode, "
    sSQL = sSQL & "B.TexasSuffix As RT45_sTexasSuffix, "
    sSQL = sSQL & "B.TexasRoofDepreciation As RT46_cTexasRoofDepreciation, "
    sSQL = sSQL & "IIF(Left(B.TypeOfLoss,2)= '01','X','') As RT47_sTypeOfPropLossBuilding, "
    sSQL = sSQL & "IIF(Left(B.TypeOfLoss,2)= '02','X','') As RT48_sTypeOfPropLossContents, "
    'BGS 2.28.2002 143  Class of Loss - 88 s/b ALE
    '8.30.2002 Use 03 for Safeco
    sSQL = sSQL & "IIF(Left(B.TypeOfLoss,2)= '03' Or Left(B.TypeOfLoss,2)= '03','X','') As RT49_sTypeOfPropLossALE, "
    sSQL = sSQL & "IIF(IsNull(A.NewSecondInsuredNames) or Trim(A.NewSecondInsuredNames) = '', A.NewInsuredNames, Trim(A.NewInsuredNames) & ' And ' & A.NewSecondInsuredNames) As RT50_sInsuredPayeeName, "
    sSQL = sSQL & "IIF(B.IncludeMorgOnDraft = TRUE, A.MortgageCompany, '') As RT51_sPayeeNames, "
    sSQL = sSQL & "A.NewMailingAddress As RT52_sAddress, "
    sSQL = sSQL & "B.AmountOfCheck As RT53_cAmountOfCheck, "
    sSQL = sSQL & "A.CatCode As RT54_sCatCode, "
    sSQL = sSQL & "IIF(B.FieldHandled = True, 'Yes', 'No') As RT55_sFieldHandled, "
    sSQL = sSQL & "IIF(B.TotalLoss = True, 'Yes', 'No') As RT56_sTotalLoss, "
    sSQL = sSQL & "IIF(B.CashInLieu = True, 'Yes', 'No') As RT57_sCashInLieu, "
    sSQL = sSQL & "IIF(B.OwnerRetainedSalvage = True, 'Yes', 'No') As RT58_sOwnerRetainSalvage, "
    sSQL = sSQL & "IIF(B.Subrogation = True, 'Yes', 'No') As RT59_sSub, "
    sSQL = sSQL & "IIF(B.Salvage = True, 'Yes', 'No') As RT60_sSalvage, "
    sSQL = sSQL & "B.Instructions As RT61_sInstructions, "
    sSQL = sSQL & "Trim(A.AdjustorFirstName) & ' ' & Trim(A.AdjustorLastName) As RT62_sRequestedBy, "
    sSQL = sSQL & "#" & dtDate & "# As RT63_dtDate"
    
    BuildClaimsSQLAB = sSQL
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarUL_ClassName & vbCrLf & " Private Function BuildClaimsSQLAB"
End Function

Private Sub PopulateRTActLog(psClaimNo As String, pCurDB As Object)
    On Error GoTo EH
    Dim RS As Recordset
    Dim sSQL As String
    Dim MyActLog As RTSafecoActLog
    Dim sClientClaim As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    sClientClaim = GetClientClaimNo(psClaimNo, pCurDB)
    Set mcolRTActLog = New Collection
    sSQL = "SELECT A.ID, "
    sSQL = sSQL & "A.ClaimNumber, "
    sSQL = sSQL & "Mid(A.ActivityDate,5,2) & '/' & Mid(A.ActivityDate,7,2) & '/' & Left(A.ActivityDate,4) As ActDate, "
    sSQL = sSQL & "A.ActivityTime, "
    sSQL = sSQL & "A.ActivityText, "
    sSQL = sSQL & "A.ServiceTime, "
    sSQL = sSQL & "A.CheckID, "
    sSQL = sSQL & "A.Association "
    sSQL = sSQL & "FROM ActivityLog As A "
    sSQL = sSQL & "WHERE A.ClaimNumber = '" & goUtil.utCleanSQLString(psClaimNo) & "' "
    sSQL = sSQL & "AND A.RTUpLoadMe = True "
   
    Set RS = pCurDB.OpenRecordset(sSQL)
    If Not RS.EOF Then
        RS.MoveFirst
        Do Until RS.EOF
            With MyActLog
                .dServiceTime = IIf(IsNull(RS!ServiceTime), 0, RS!ServiceTime)
                .dtActDate = IIf(IsNull(RS!ActDate) Or RS!ActDate = "//", NULL_DATE, RS!ActDate)
                .lActID = IIf(IsNull(RS!ID), 0, RS!ID)
                .sActText = IIf(IsNull(RS!ActivityText), vbNullString, RS!ActivityText)
                .sActTime = IIf(IsNull(RS!ActivityTime), vbNullString, RS!ActivityTime)
                .sAssociation = IIf(IsNull(RS!Association), vbNullString, RS!Association)
                .sClientClaimNumber = sClientClaim
                .sIBNumber = psClaimNo
            End With
            mcolRTActLog.Add MyActLog, CStr(MyActLog.lActID)
            RS.MoveNext
        Loop
    End If
    
    Set RS = Nothing
    
    'Need to update RT flag
    sSQL = "UPDATE ActivityLog SET "
    sSQL = sSQL & "ActivityLog.RTUpLoadMe = False "
    sSQL = sSQL & "WHERE ActivityLog.ClaimNumber = '" & goUtil.utCleanSQLString(psClaimNo) & "' "
    pCurDB.Execute sSQL
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarUL_ClassName & vbCrLf & " Private Sub PopulateRTActLog"
End Sub

Private Sub PopulateRTPhotoLog(psClaimNo As String, pCurDB As Object)
    On Error GoTo EH
    Dim RS As Recordset
    Dim sSQL As String
    Dim MyPhotoLog As RTSafecoPhotoLog
    Dim sClientClaim As String
    Dim sPhotoName As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    sClientClaim = GetClientClaimNo(psClaimNo, pCurDB)
    Set mcolRTPhotoLog = New Collection
    sSQL = "SELECT A.ID, "
    sSQL = sSQL & "A.ClaimNo, "
    sSQL = sSQL & "Mid(A.PhotoDate,5,2) & '/' & Mid(A.PhotoDate,7,2) & '/' & Left(A.PhotoDate,4) As PhotoDate, "
    sSQL = sSQL & "A.PhotoPath, "
    sSQL = sSQL & "A.Description, "
    sSQL = sSQL & "A.SortOrder "
    sSQL = sSQL & "FROM PhotoLog As A "
    sSQL = sSQL & "WHERE A.ClaimNo = '" & goUtil.utCleanSQLString(psClaimNo) & "' "
    sSQL = sSQL & "AND A.RTUpLoadMe = True "
   
    Set RS = pCurDB.OpenRecordset(sSQL)
    If Not RS.EOF Then
        RS.MoveFirst
        Do Until RS.EOF
            With MyPhotoLog
                If IsNull(RS!photopath) Then
                    sPhotoName = vbNullString
                Else
                    sPhotoName = RS!photopath
                    sPhotoName = Mid(sPhotoName, InStrRev(sPhotoName, "\") + 1)
                    sPhotoName = Replace(sPhotoName, "_0.jpg", "_1.jpg")
                End If
                .dtPhotoDate = IIf(IsNull(RS!PhotoDate) Or RS!PhotoDate = "//", NULL_DATE, RS!PhotoDate)
                .lPhotoID = IIf(IsNull(RS!ID), 0, RS!ID)
                .lSortOrder = IIf(IsNull(RS!SortOrder), 0, RS!SortOrder)
                .sClientClaimNumber = sClientClaim
                .sDescription = IIf(IsNull(RS!Description), vbNullString, RS!Description)
                .sIBNumber = psClaimNo
                .sPhotoName = sPhotoName
            End With
            mcolRTPhotoLog.Add MyPhotoLog, CStr(MyPhotoLog.lPhotoID)
            RS.MoveNext
        Loop
    End If
    
    Set RS = Nothing
    
    'Need to update RT flag
    sSQL = "UPDATE PhotoLog SET "
    sSQL = sSQL & "PhotoLog.RTUpLoadMe = False "
    sSQL = sSQL & "WHERE PhotoLog.ClaimNo = '" & goUtil.utCleanSQLString(psClaimNo) & "' "
    pCurDB.Execute sSQL
   
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarUL_ClassName & vbCrLf & " Private Sub PopulateRTPhotoLog"
End Sub

Private Sub PopulateRTIndem(psClaimNo As String, pCurDB As Object)
    On Error GoTo EH
    Dim RS As Recordset
    Dim sSQL As String
    Dim MyIndem As RTSafecoIndemnity
    Dim sClientClaim As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    sClientClaim = GetClientClaimNo(psClaimNo, pCurDB)
    Set mcolRTIndem = New Collection
    sSQL = "SELECT A.ID, "
    sSQL = sSQL & "A.ClaimNo, "
    sSQL = sSQL & "A.IndemType, "
    sSQL = sSQL & "A.Class, "
    sSQL = sSQL & "A.CauseOfLoss, "
    sSQL = sSQL & "A.Description, "
    sSQL = sSQL & "A.ReplacementCost, "
    sSQL = sSQL & "A.RecoverableDepreciation, "
    sSQL = sSQL & "A.NonRecoverableDepreciation, "
    sSQL = sSQL & "A.ACVClaim, "
    sSQL = sSQL & "A.LessExcessLimit, "
    sSQL = sSQL & "A.ACVLessExcessLimit "
    sSQL = sSQL & "FROM Indemnity As A "
    sSQL = sSQL & "WHERE A.ClaimNo = '" & goUtil.utCleanSQLString(psClaimNo) & "' "
    sSQL = sSQL & "AND A.RTUpLoadMe = True "
   
    Set RS = pCurDB.OpenRecordset(sSQL)
    If Not RS.EOF Then
        RS.MoveFirst
        Do Until RS.EOF
            With MyIndem
                .cACVClaim = IIf(IsNull(RS!ACVClaim), 0, RS!ACVClaim)
                .cACVLessExcessLimit = IIf(IsNull(RS!ACVLessExcessLimit), 0, RS!ACVLessExcessLimit)
                .cExcessLimit = IIf(IsNull(RS!LessExcessLimit), 0, RS!LessExcessLimit)
                .cNonRecoverableDepreciation = IIf(IsNull(RS!NonRecoverableDepreciation), 0, RS!NonRecoverableDepreciation)
                .cRecoverableDepreciation = IIf(IsNull(RS!RecoverableDepreciation), 0, RS!RecoverableDepreciation)
                .cReplacementCost = IIf(IsNull(RS!ReplacementCost), 0, RS!ReplacementCost)
                .lIndemID = IIf(IsNull(RS!ID), 0, RS!ID)
                .sCauseOfLoss = IIf(IsNull(RS!CauseOfLoss), vbNullString, RS!CauseOfLoss)
                .sClass = IIf(IsNull(RS!Class), vbNullString, RS!Class)
                .sClientClaimNumber = sClientClaim
                .sDescription = IIf(IsNull(RS!Description), vbNullString, RS!Description)
                .sIBNumber = psClaimNo
                .sIndemType = IIf(IsNull(RS!IndemType), vbNullString, RS!IndemType)
            End With
            mcolRTIndem.Add MyIndem, CStr(MyIndem.lIndemID)
            RS.MoveNext
        Loop
    End If
    
    Set RS = Nothing
    
    'Need to update RT flag
    sSQL = "UPDATE Indemnity SET "
    sSQL = sSQL & "Indemnity.RTUpLoadMe = False "
    sSQL = sSQL & "WHERE Indemnity.ClaimNo = '" & goUtil.utCleanSQLString(psClaimNo) & "' "
    pCurDB.Execute sSQL
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarUL_ClassName & vbCrLf & " Private Sub PopulateRTIndem"
End Sub

Private Sub PopulateRTDelAction(psClaimNo As String, pCurDB As Object)
    On Error GoTo EH
    Dim RS As Recordset
    Dim sSQL As String
    Dim MyRTDel As V2ECKeyBoard.RTDeleteAction
    Dim sClientClaim As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    Set mcolRTDel = New Collection
    sSQL = "SELECT A.ID, "
    sSQL = sSQL & "A.TableName, "
    sSQL = sSQL & "A.IBNumber, "
    sSQL = sSQL & "A.ClientClaimNumber "
    sSQL = sSQL & "FROM RTDelete As A "
    sSQL = sSQL & "WHERE A.IBNumber = '" & goUtil.utCleanSQLString(psClaimNo) & "' "
    sSQL = sSQL & "OR Instr(1, A.TableName, 'DELETE_CLAIM') > 0 "
   
    Set RS = pCurDB.OpenRecordset(sSQL)
    If Not RS.EOF Then
        RS.MoveFirst
        Do Until RS.EOF
            With MyRTDel
                .ID = IIf(IsNull(RS!ID), 0, RS!ID)
                .sClientClaimNumber = IIf(IsNull(RS!ClientClaimNumber), vbNullString, RS!ClientClaimNumber)
                .sIBNumber = IIf(IsNull(RS!IBNumber), vbNullString, RS!IBNumber)
                .sTableName = IIf(IsNull(RS!TableName), vbNullString, RS!TableName)
            End With
            mcolRTDel.Add MyRTDel
            RS.MoveNext
        Loop
    End If
    
    Set RS = Nothing
    
    'Need to update RT flag
    sSQL = "DELETE * FROM RTDelete As A "
    sSQL = sSQL & "WHERE A.IBNumber = '" & goUtil.utCleanSQLString(psClaimNo) & "' "
    sSQL = sSQL & "OR Instr(1, A.TableName, 'DELETE_CLAIM') > 0 "
    pCurDB.Execute sSQL
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarUL_ClassName & vbCrLf & " Private Sub PopulateRTDelAction"
End Sub

Private Function GetRTID() As String
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    Randomize Timer
    GetRTID = Format(Now(), "mmddyyhhmmss") & (Int(Rnd * 9)) & (Int(Rnd * 9)) & (Int(Rnd * 9)) & (Int(Rnd * 9)) & "RT"

    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarUL_ClassName & vbCrLf & " Private Function GetRTID"
End Function

Private Sub Class_Initialize()
    msClassName = "clsULSafecoRT"
    msTableName = "RTSafeco"
End Sub
  
Private Sub Class_Terminate()
    Set mcolRTActLog = Nothing
    Set mcolRTIndem = Nothing
    Set mcolRTPhotoLog = Nothing
    Set mcolRTDel = Nothing
End Sub
