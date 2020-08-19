SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spuUpdateUsersComments]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spuUpdateUsersComments]
GO

/******************************************************************
*
*    	Description:	Update Users Comments
*			
*	
*	Author: 	Brad Skidmore
*	Date: 		4/29/2004
******************************************************************/

CREATE PROCEDURE [dbo].[z_spuUpdateUsersComments]
	@UID			int,
	@UsersID		Int=null,
	@Comments		varchar(1000)
AS
UPDATE USERS SET 	
			Comments = @Comments,
			DateLastUpdated = GetDate(),
			UpdateByUserID = @UID
FROM USERS
WHERE USERS.USERSID = @UsersID

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[z_spuUpdateUsersComments]  TO [CFUSER]
GO

