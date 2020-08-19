SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spsGetBillAssignmentCATDropDown]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spsGetBillAssignmentCATDropDown]
GO

/******************************************************************
*
*    	Description:	Get Bill Assignment CAT DropDown
*	
*	Author: 	Brad Skidmore
*	Date: 		03/31/2005
******************************************************************/

CREATE PROCEDURE [dbo].[z_spsGetBillAssignmentCATDropDown]
	@BillAssignmentID		int
AS

	SELECT	
		[ClientCompanyCatSpecID],
		[CatCode],
		[Comments]
	FROM	ClientCompanyCatSpec
	WHERE	[ClientCompanyID] IN 	(
						SELECT 	[ClientCompanyID]
						FROM	ClientCompanyCatSpec CCCS
						Inner Join BillAssignment BA ON BA.[ClientCompanyCatSpecID] = CCCS.[ClientCompanyCatSpecID]
						WHERE	BA.[BillAssignmentID] = @BillAssignmentID
					)
	ORDER BY [CatCode]

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[z_spsGetBillAssignmentCATDropDown]  TO [CFUSER]
GO

