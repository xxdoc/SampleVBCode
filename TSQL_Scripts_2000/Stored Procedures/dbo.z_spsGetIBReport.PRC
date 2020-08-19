SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spsGetIBReport]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spsGetIBReport]
GO

/******************************************************************
*
*    	Description:	Return IB for Report Population
*	
*	Author: 	Brad Skidmore
*	Date: 		5/7/2004
*
******************************************************************/
CREATE PROCEDURE [dbo].[z_spsGetIBReport]
	@AssignmentsID	int,
	@IBID		int=Null,
	@IBHistoryID	int=Null
AS

IF @IBID Is Null And @IBHistoryID Is Null
BEGIN
	SELECT 	
		(CASE 	WHEN RTChecks.[PrintOnIB] Is Null
			THEN 	0
			ELSE
				RTChecks.[PrintOnIB]
			END
		) 						As bPrintOnIB,
		RTIB.[RT01_sSubToCarrier] 			As rt01_s,
		(CASE	WHEN 	RTIB.[RT14a_sSupplement] > 0 And RTIB.[RT14b_sRebilled] > 0
			THEN	RTIB.[RT02_sIBNumber] + 'S' + cast(RTIB.[RT14a_sSupplement] As varchar(4)) + 'R' + cast(RTIB.[RT14b_sRebilled] As varchar(4))
			ELSE	(CASE	WHEN 	RTIB.[RT14a_sSupplement] > 0
					THEN	RTIB.[RT02_sIBNumber] + 'S' + cast(RTIB.[RT14a_sSupplement] As varchar(4))
					ELSE	(CASE	WHEN	RTIB.[RT14b_sRebilled] > 0
							THEN	RTIB.[RT02_sIBNumber] + 'R' + cast(RTIB.[RT14b_sRebilled] As varchar(4))
							ELSE	RTIB.[RT02_sIBNumber]
							END
						)
					END
				)
			END
		)						As RT02_sIBNumber,
		CCCS.[CatCode] 					As rt04_s,	
		RTIB.[RT05_sLocation] 				As rt05_s,	
		RTIB.[RT05a_sState] 				As rt05a_s,	
		RTIB.[RT06_dtDateClosed] 			As rt06_dt,	
		RTIB.[RT07_sAdjusterName]  			As rt07_s,	
		CCAS.[ACID] 					As rt08_s,	
		RTIB.[RT09_sSALN] 				As rt09_sSALN,
		RTIB.[RT10_sInsuredName] 			As rt10_s,	
		RTIB.[RT11_sLossLocation] 			As rt11_s,	
		RTIB.[RT12_dtDateOfLoss] 			As rt12_dt,	
		RTIB.[RT13_cGrossLoss] 				As rt13_c,	
		RTIB.[RT14_cDepreciation] 			As rt14_c,	
		RTIB.[RT14a_sSupplement] 			As rt14a_s,	
		RTIB.[RT14b_sRebilled] 				As rt14b_s,	
		RTIB.[RT15_cDeductible] 			As rt15_c,	
		RTIB.[RT15a_cLessExcessLimits] 			As rt15a_c,	
		RTIB.[RT15b_sExcessLimDesc] 			As rt15b_s,	
		RTIB.[RT15c_cLessMiscellaneous] 		As rt15c_c,	
		RTIB.[RT15d_cMiscellaneousDesc] 		As rt15d_c,	
		RTIB.[RT16_cNetClaim] 				As rt16_c,	
		RTIB.[RT17_cServiceFee] 			As rt17_c,
		RTIB.[RT17a_cMiscServiceFee] 			As rt17a_c,
		RTIB.[RT18_sServiceFeeComment] 			As rt18_s,
		RTIB.[RT18a_sMiscServiceFeeComment] 		As rt18a_s,
		RTIB.[RT25_cServiceFeeSubTotal] 		As rt25_c,	
		RTIB.[RT29a_sMiscExpenseFeeComment] 		As rt29a_s,
		RTIB.[RT29b_cMiscExpenseFee]  			As rt29b_c,
		RTIB.[RT30_cTotalExpenses]  			As rt30_c,	
		RTIB.[RT31_dTaxPercent]  			As rt31_d,	
		RTIB.[RT32_cTaxAmount]  			As rt32_c,	
		RTIB.[RT33_cTotalAdjustingFee]  		As rt33_c,
		RTIB.[RT33a_sAccountCode]  			As rt33a_s,	
		(	
			SELECT 	Code
			FROM	ClassOfLoss
			WHERE	ClassOfLossID = RTChecks.[RT42_ClassOfLossID]
		) 						As rt42_s,	
		(	
			SELECT 	Code
			FROM	TypeOfLoss
			WHERE	TypeOfLossID =RTChecks.[RT43_TypeOfLossID] 	
		) 						As rt43_s,	
		RTChecks.[RT50_sInsuredPayeeName] 		As rt50_s,	
		RTChecks.[RT51_sPayeeNames]  			As rt51_s,	
		RTChecks.[RT52_sAddress]  			As rt52_s,	
		RTChecks.[RT53_cAmountOfCheck] 			As rt53_c,	
		CCCS.[CatCode]  				As rt54_s	
	FROM 		RTIB
	INNER JOIN 	Assignments A	ON
					RTIB.AssignmentsID = A.AssignmentsID
	LEFT JOIN	RTChecks 	ON 
					RTIB.AssignmentsID = RTChecks.AssignmentsID 
					And RTIB.BillingCountID = RTChecks.BillingCountID
					And RTChecks.CheckNum = 1
	INNER JOIN 	ClientCompanyCatSpec CCCS 	ON
							CCCS.ClientCompanyCatSpecID = A.ClientCompanyCatSpecID
	INNER JOIN 	ClientCoAdjusterSpec CCAS 	ON
							CCAS.ClientCoAdjusterSpecID = A.AdjusterSpecIDDisplay							
					
	WHERE 		RTIB.AssignmentsID  = @AssignmentsID
END
ELSE IF @IBHistoryID Is null
BEGIN
	SELECT 		
		(CASE 	WHEN RTChecks.[PrintOnIB] Is Null
			THEN 	0
			ELSE
				RTChecks.[PrintOnIB]
			END
		) 						As bPrintOnIB,
		IB.[IB01_sSubToCarrier] 			As rt01_s,	
		(CASE	WHEN 	IB.[IB14a_sSupplement] > 0 And IB.[IB14b_sRebilled] > 0
			THEN	IB.[IB02_sIBNumber] + 'S' + cast(IB.[IB14a_sSupplement] As varchar(4)) + 'R' + cast(IB.[IB14b_sRebilled] As varchar(4))
			ELSE	(CASE	WHEN 	IB.[IB14a_sSupplement] > 0
					THEN	IB.[IB02_sIBNumber] + 'S' + cast(IB.[IB14a_sSupplement] As varchar(4))
					ELSE	(CASE	WHEN	IB.[IB14b_sRebilled] > 0
							THEN	IB.[IB02_sIBNumber] + 'R' + cast(IB.[IB14b_sRebilled] As varchar(4))
							ELSE	IB.[IB02_sIBNumber]
							END
						)
					END
				)
			END
		)						As RT02_sIBNumber,
		CCCS.[CatCode] 					As rt04_s,	
		IB.[IB05_sLocation] 				As rt05_s,	
		IB.[IB05a_sState] 				As rt05a_s,	
		IB.[IB06_dtDateClosed] 				As rt06_dt,	
		IB.[IB07_sAdjusterName]  			As rt07_s,	
		CCAS.[ACID] 					As rt08_s,	
		IB.[IB09_sSALN] 				As rt09_sSALN,
		IB.[IB10_sInsuredName] 				As rt10_s,	
		IB.[IB11_sLossLocation] 			As rt11_s,	
		IB.[IB12_dtDateOfLoss] 				As rt12_dt,	
		IB.[IB13_cGrossLoss] 				As rt13_c,	
		IB.[IB14_cDepreciation] 			As rt14_c,	
		IB.[IB14a_sSupplement] 				As rt14a_s,	
		IB.[IB14b_sRebilled] 				As rt14b_s,	
		IB.[IB15_cDeductible] 				As rt15_c,	
		IB.[IB15a_cLessExcessLimits] 			As rt15a_c,	
		IB.[IB15b_sExcessLimDesc] 			As rt15b_s,	
		IB.[IB15c_cLessMiscellaneous] 			As rt15c_c,	
		IB.[IB15d_cMiscellaneousDesc] 			As rt15d_c,	
		IB.[IB16_cNetClaim] 				As rt16_c,	
		IB.[IB17_cServiceFee] 				As rt17_c,
		IB.[IB17a_cMiscServiceFee] 			As rt17a_c,
		IB.[IB18_sServiceFeeComment] 			As rt18_s,
		IB.[IB18a_sMiscServiceFeeComment] 		As rt18a_s,
		IB.[IB25_cServiceFeeSubTotal] 			As rt25_c,	
		IB.[IB29a_sMiscExpenseFeeComment] 		As rt29a_s,
		IB.[IB29b_cMiscExpenseFee]  			As rt29b_c,
		IB.[IB30_cTotalExpenses]  			As rt30_c,	
		IB.[IB31_dTaxPercent]  				As rt31_d,	
		IB.[IB32_cTaxAmount]  				As rt32_c,	
		IB.[IB33_cTotalAdjustingFee]  			As rt33_c,
		IB.[IB33a_sAccountCode]  			As rt33a_s,	
		(	
			SELECT 	Code
			FROM	ClassOfLoss
			WHERE	ClassOfLossID = RTChecks.[RT42_ClassOfLossID]
		) 						As rt42_s,	
		(	
			SELECT 	Code
			FROM	TypeOfLoss
			WHERE	TypeOfLossID =RTChecks.[RT43_TypeOfLossID] 	
		) 						As rt43_s,	
		RTChecks.[RT50_sInsuredPayeeName] 		As rt50_s,	
		RTChecks.[RT51_sPayeeNames]  			As rt51_s,	
		RTChecks.[RT52_sAddress]  			As rt52_s,	
		RTChecks.[RT53_cAmountOfCheck] 			As rt53_c,	
		CCCS.[CatCode]  				As rt54_s	
	FROM 		IB
	INNER JOIN 	Assignments A	ON
					IB.AssignmentsID = A.AssignmentsID
	LEFT JOIN	RTChecks 	ON 
					IB.AssignmentsID = RTChecks.AssignmentsID 
					And IB.BillingCountID = RTChecks.BillingCountID
					And RTChecks.CheckNum = 1
	INNER JOIN 	ClientCompanyCatSpec CCCS 	ON
							CCCS.ClientCompanyCatSpecID = A.ClientCompanyCatSpecID
	INNER JOIN 	ClientCoAdjusterSpec CCAS 	ON
							CCAS.ClientCoAdjusterSpecID = A.AdjusterSpecIDDisplay							
					
	WHERE 		IB.AssignmentsID  = @AssignmentsID
			And IB.IBID = @IBID
END
ELSE IF @IBHistoryID Is Not Null
BEGIN
	SELECT 		
		(CASE 	WHEN RTChecks.[PrintOnIB] Is Null
			THEN 	0
			ELSE
				RTChecks.[PrintOnIB]
			END
		) 						As bPrintOnIB,
		IB.[IB01_sSubToCarrier] 			As rt01_s,	
		(CASE	WHEN 	IB.[IB14a_sSupplement] > 0 And IB.[IB14b_sRebilled] > 0
			THEN	IB.[IB02_sIBNumber] + 'S' + cast(IB.[IB14a_sSupplement] As varchar(4)) + 'R' + cast(IB.[IB14b_sRebilled] As varchar(4))
			ELSE	(CASE	WHEN 	IB.[IB14a_sSupplement] > 0
					THEN	IB.[IB02_sIBNumber] + 'S' + cast(IB.[IB14a_sSupplement] As varchar(4))
					ELSE	(CASE	WHEN	IB.[IB14b_sRebilled] > 0
							THEN	IB.[IB02_sIBNumber] + 'R' + cast(IB.[IB14b_sRebilled] As varchar(4))
							ELSE	IB.[IB02_sIBNumber]
							END
						)
					END
				)
			END
		)						As RT02_sIBNumber,
		CCCS.[CatCode] 					As rt04_s,	
		IB.[IB05_sLocation] 				As rt05_s,	
		IB.[IB05a_sState] 				As rt05a_s,	
		IB.[IB06_dtDateClosed] 				As rt06_dt,	
		IB.[IB07_sAdjusterName]  			As rt07_s,	
		CCAS.[ACID] 					As rt08_s,	
		IB.[IB09_sSALN] 				As rt09_sSALN,
		IB.[IB10_sInsuredName] 				As rt10_s,	
		IB.[IB11_sLossLocation] 			As rt11_s,	
		IB.[IB12_dtDateOfLoss] 				As rt12_dt,	
		IB.[IB13_cGrossLoss] 				As rt13_c,	
		IB.[IB14_cDepreciation] 			As rt14_c,	
		IB.[IB14a_sSupplement] 				As rt14a_s,	
		IB.[IB14b_sRebilled] 				As rt14b_s,	
		IB.[IB15_cDeductible] 				As rt15_c,	
		IB.[IB15a_cLessExcessLimits] 			As rt15a_c,	
		IB.[IB15b_sExcessLimDesc] 			As rt15b_s,	
		IB.[IB15c_cLessMiscellaneous] 			As rt15c_c,	
		IB.[IB15d_cMiscellaneousDesc] 			As rt15d_c,	
		IB.[IB16_cNetClaim] 				As rt16_c,	
		IB.[IB17_cServiceFee] 				As rt17_c,
		IB.[IB17a_cMiscServiceFee] 			As rt17a_c,
		IB.[IB18_sServiceFeeComment] 			As rt18_s,
		IB.[IB18a_sMiscServiceFeeComment] 		As rt18a_s,
		IB.[IB25_cServiceFeeSubTotal] 			As rt25_c,	
		IB.[IB29a_sMiscExpenseFeeComment] 		As rt29a_s,
		IB.[IB29b_cMiscExpenseFee]  			As rt29b_c,
		IB.[IB30_cTotalExpenses]  			As rt30_c,	
		IB.[IB31_dTaxPercent]  				As rt31_d,	
		IB.[IB32_cTaxAmount]  				As rt32_c,	
		IB.[IB33_cTotalAdjustingFee]  			As rt33_c,
		IB.[IB33a_sAccountCode]  			As rt33a_s,	
		(	
			SELECT 	Code
			FROM	ClassOfLoss
			WHERE	ClassOfLossID = RTChecks.[RT42_ClassOfLossID]
		) 						As rt42_s,	
		(	
			SELECT 	Code
			FROM	TypeOfLoss
			WHERE	TypeOfLossID =RTChecks.[RT43_TypeOfLossID] 	
		) 						As rt43_s,	
		RTChecks.[RT50_sInsuredPayeeName] 		As rt50_s,	
		RTChecks.[RT51_sPayeeNames]  			As rt51_s,	
		RTChecks.[RT52_sAddress]  			As rt52_s,	
		RTChecks.[RT53_cAmountOfCheck] 			As rt53_c,	
		CCCS.[CatCode]			  		As rt54_s
	FROM 		IBHistory IB
	INNER JOIN 	Assignments A	ON
					IB.AssignmentsID = A.AssignmentsID
	LEFT JOIN	RTChecks 	ON 
					IB.AssignmentsID = RTChecks.AssignmentsID 
					And IB.BillingCountID = RTChecks.BillingCountID
					And RTChecks.CheckNum = 1
	INNER JOIN 	ClientCompanyCatSpec CCCS 	ON
							CCCS.ClientCompanyCatSpecID = A.ClientCompanyCatSpecID
	INNER JOIN 	ClientCoAdjusterSpec CCAS 	ON
							CCAS.ClientCoAdjusterSpecID = A.AdjusterSpecIDDisplay							
					
	WHERE 		IB.AssignmentsID  = @AssignmentsID
			And IB.IBHistoryID = @IBHistoryID
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[z_spsGetIBReport]  TO [CFUSER]
GO

