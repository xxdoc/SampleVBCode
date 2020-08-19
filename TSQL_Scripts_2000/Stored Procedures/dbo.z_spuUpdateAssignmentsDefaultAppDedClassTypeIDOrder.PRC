SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spuUpdateAssignmentsDefaultAppDedClassTypeIDOrder]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spuUpdateAssignmentsDefaultAppDedClassTypeIDOrder]
GO

/******************************************************************
*
*    	Description:	Update Assignments DefaultAppDedClassTypeIDOrder
*			
*	Author: 	Brad Skidmore
*	Date: 		4/13/2004
*	
******************************************************************/
CREATE PROCEDURE [dbo].[z_spuUpdateAssignmentsDefaultAppDedClassTypeIDOrder]
	@UID					int,
	@DefaultAppDedClassTypeIDOrder		varchar(255),
	@ClientCompanyID			int,
	@FeeScheduleID				int
	
AS
UPDATE ASSIGNMENTS	SET 	AppDedClassTypeIDOrder 	= @DefaultAppDedClassTypeIDOrder,
				AdminComments	   	= 'Admin Updated Default Apply Deductible By Class Type Order',
				DateLastUpdated 	= GetDate(),
				UpdateByUserID		= @UID
WHERE	ClientCompanyCatSpecID IN 	(
					SELECT 	ClientCompanyCatSpecID
					FROM	ClientCompanyCatSpec
					WHERE	CATID IN	(
								SELECT 	CATID
								FROM 	ClientCompanyCat
								WHERE  	ClientCompanyID = @ClientCompanyID
								AND 	FeeScheduleID = @FeeScheduleID
								AND 	InactiveDate Is Null
								)
					AND 	ClientCompanyID = @ClientCompanyID
					AND 	InactiveDate Is Null
					)

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[z_spuUpdateAssignmentsDefaultAppDedClassTypeIDOrder]  TO [CFUSER]
GO

