SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spsGetBillAssignmentACIDDropDown]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spsGetBillAssignmentACIDDropDown]
GO

/******************************************************************
*
*    	Description:	Get Bill Assignment ACID DropDown
*	
*	Author: 	Brad Skidmore
*	Date: 		03/31/2005
******************************************************************/

CREATE PROCEDURE [dbo].[z_spsGetBillAssignmentACIDDropDown]
	@BillAssignmentID		int
AS

	SELECT	
		[ClientCoAdjusterSpecID] As [AdjusterSpecID],
		[ACID],
		[ACIDDescription]
	FROM	ClientCoAdjusterSpec
	WHERE	[ClientCompanyID] IN 	(
						SELECT 	[ClientCompanyID]
						FROM	ClientCoAdjusterSpec CCAS
						Inner Join BillAssignment BA ON BA.[AdjusterSpecID] = CCAS.[ClientCoAdjusterSpecID]
						WHERE	BA.[BillAssignmentID] = @BillAssignmentID
					)
	ORDER BY [ACID]


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[z_spsGetBillAssignmentACIDDropDown]  TO [CFUSER]
GO

