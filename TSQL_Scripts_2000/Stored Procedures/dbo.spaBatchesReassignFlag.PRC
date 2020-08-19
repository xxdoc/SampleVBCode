SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spaBatchesReassignFlag]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[spaBatchesReassignFlag]
GO

/******************************************************************
*
*    	Description:	Flag and comment All E-Bills Under @ClientClaimNum 
*			in Batches Table that it has been Reasigned.
*	
*	Author: 	Brad Skidmore
*	Date: 		06/04/2003
*
******************************************************************/
CREATE PROCEDURE [dbo].[spaBatchesReassignFlag] 
	@UsersID	int,
	@AssignmentsID	int,
	@Comments		Varchar(100)
AS

Set @Comments = 'Admin User:' + (SELECT UserName FROM Users WHERE UsersID = @UsersID) + ' ' + @Comments

UPDATE Batches SET	[Reassigned] = Reassigned + 1,
			[Comments] = @Comments,
			[DateLastupdated] = GetDate(),
			[UpdateByUserID] = @UsersID
WHERE 		AssignmentsID = @AssignmentsID

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[spaBatchesReassignFlag]  TO [CFUSER]
GO

