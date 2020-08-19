SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spsGetBillAssignmentIBData]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spsGetBillAssignmentIBData]
GO

/******************************************************************
*
*    	Description:	Get Bill Assignment IB Data
*	
*	Author: 	Brad Skidmore
*	Date: 		03/28/2005
*	Date:		4/21/2005 BGS Added ClaimNumber (AKA CLIENTNUM)
******************************************************************/

CREATE PROCEDURE [dbo].[z_spsGetBillAssignmentIBData]
	@ClientCompanyCatSpecID		int,
	@AdjusterSpecID			int,
	@CLIENTNUM			varchar(20),
	@PolicyNo			varchar(50),
	@Supplement			int
AS

DECLARE @BillAssignmentID	int
DECLARE	@BillBillingCountID	int
DECLARE @sClientCompanyDBName 	varchar(30)

--Need to figure out wich Bill Assignment this is
SET @BillAssignmentID = 	(
					SELECT 	[BillAssignmentID]
					FROM 	BillAssignment
					WHERE	[ClientCompanyCatSpecID] = @ClientCompanyCatSpecID
					AND	[AdjusterSpecID] = @AdjusterSpecID
					AND	[CLIENTNUM] = @CLIENTNUM
					AND	[PolicyNo] = @PolicyNo
				)	
--Then Get the Bill BillingCountID ID for the appropriate IB Supplement
-- 0 = First IB 1+ = Supplementals
SET @BillBillingCountID = 	(
					SELECT 	[BillBillingCountID]
					FROM 	BillBillingCount
					WHERE	[BillAssignmentID] = @BillAssignmentID
					AND	[Supplement] = @Supplement
				)	
--Need to figure out which IB table to pull the IB Data From
SET @sClientCompanyDBName = 	(
				SELECT	[DBName]
				FROM	Company C
				WHERE 	[IsClientOF] = 	(
							SELECT 	[CompanyID]
							FROM	Company
							WHERE 	[DBName] = 'Eberls'
							)
				AND	[CompanyID] =	(
							SELECT 	CCCS.[ClientCompanyID]	
							FROM	ClientCompanyCatSpec CCCS 
							WHERE	CCCS.[ClientCompanyCatSpecID] = @ClientCompanyCatSpecID
							)
				)

IF @sClientCompanyDBName = 'StateFarm' 
BEGIN
	if @BillBillingCountID Is Not Null
	BEGIN
		SELECT	IB.*,
			BA.[AdjusterSpecID],
			BA.[CLIENTNUM] As [CLAIMNUMBER],
			U.[FirstName] + ' ' + U.[LastName] As [AdjusterName]
		FROM	IBStateFarm IB
			Inner Join BillAssignment BA ON BA.BillAssignmentID = IB.BillAssignmentID
			Inner Join ClientCoAdjusterSpec CCAS ON BA.AdjusterSpecID = CCAS.ClientCoAdjusterSpecID
			Inner Join Users U ON U.UsersID = CCAS.UsersID
		WHERE	IB.[BillAssignmentID] = @BillAssignmentID
		AND 	[BillBillingCountID] = @BillBillingCountID
	END
	ELSE
	BEGIN
		SELECT	IB.*,
			BA.[AdjusterSpecID],
			BA.[CLIENTNUM] As [CLAIMNUMBER],
			U.[FirstName] + ' ' + U.[LastName] As [AdjusterName] 	
		FROM	IBStateFarm IB
			Inner Join BillAssignment BA ON BA.BillAssignmentID = IB.BillAssignmentID
			Inner Join ClientCoAdjusterSpec CCAS ON BA.AdjusterSpecID = CCAS.ClientCoAdjusterSpecID
			Inner Join Users U ON U.UsersID = CCAS.UsersID
		WHERE	IB.[BillAssignmentID] = @BillAssignmentID
		Order By [Supplement], [Rebilled]
	END
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[z_spsGetBillAssignmentIBData]  TO [CFUSER]
GO

