SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spsGetRTChecksReport]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spsGetRTChecksReport]
GO

/******************************************************************
*
*    	Description:	Return RTChecks for Report Population
*	
*	Author: 	Brad Skidmore
*	Date: 		5/19/2004
*
******************************************************************/
CREATE PROCEDURE [dbo].[z_spsGetRTChecksReport]
	@AssignmentsID		int,
	@RTChecksID		int=Null,
	@RTChecksHistoryID	int=Null
AS

IF @RTChecksHistoryID Is null
BEGIN
	SELECT 		
		(CASE 	WHEN RTChecks.[PrintOnIB] Is Null
			THEN 	0
			ELSE
				RTChecks.[PrintOnIB]
			END
		) 						As bPrintOnIB,
		CCCS.[CatCode] 					As rt04_s,	
		CCAS.[ACID] 					As rt08_s,
		RTChecks.CheckNum				As rt03_s,
		(	SELECT 	Count(AssignmentsID) 
			FROM 	RTChecks
			WHERE	AssignmentsID = @AssignmentsID
			And 	IsDeleted = 0	
		)						As CountOfChecks,
		RTChecks.RT50_sInsuredPayeeName			As rt10_s,
		A.CLIENTNUM					As sClientCla, 			
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
	FROM 		RTChecks
	INNER JOIN 	Assignments A	ON
					RTChecks.AssignmentsID = A.AssignmentsID
	INNER JOIN 	ClientCompanyCatSpec CCCS 	ON
							CCCS.ClientCompanyCatSpecID = A.ClientCompanyCatSpecID
	INNER JOIN 	ClientCoAdjusterSpec CCAS 	ON
							CCAS.ClientCoAdjusterSpecID = A.AdjusterSpecIDDisplay							
					
	WHERE 		RTChecks.AssignmentsID  = @AssignmentsID
			And RTChecks.RTChecksID = @RTChecksID
END
ELSE IF @RTChecksHistoryID Is Not Null
BEGIN
	SELECT 		
		(CASE 	WHEN RTChecks.[PrintOnIB] Is Null
			THEN 	0
			ELSE
				RTChecks.[PrintOnIB]
			END
		) 						As bPrintOnIB,
		CCCS.[CatCode] 					As rt04_s,	
		CCAS.[ACID] 					As rt08_s,
		RTChecks.CheckNum				As rt03_s,
		(	SELECT 	Count(AssignmentsID) 
			FROM 	RTChecks
			WHERE	AssignmentsID = @AssignmentsID
			And 	IsDeleted = 0	
		)						As CountOfChecks,
		RTChecks.RT50_sInsuredPayeeName			As rt10_s,
		A.CLIENTNUM					As sClientCla, 			
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
	FROM 		RTChecksHistory RTChecks
	INNER JOIN 	Assignments A	ON
					RTChecks.AssignmentsID = A.AssignmentsID
	INNER JOIN 	ClientCompanyCatSpec CCCS 	ON
							CCCS.ClientCompanyCatSpecID = A.ClientCompanyCatSpecID
	INNER JOIN 	ClientCoAdjusterSpec CCAS 	ON
							CCAS.ClientCoAdjusterSpecID = A.AdjusterSpecIDDisplay							
					
	WHERE 		RTChecks.AssignmentsID  = @AssignmentsID
			And RTChecks.RTChecksHistoryID = @RTChecksHistoryID
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[z_spsGetRTChecksReport]  TO [CFUSER]
GO

