SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spsBuildUsersCommentsHistory]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spsBuildUsersCommentsHistory]
GO

/******************************************************************
*
*    	Description:	Build Comments history for Users/CompanyUsers
*			
*	
*	Author: 	Brad Skidmore
*	Date: 		10/14/2003
*
******************************************************************/

CREATE PROCEDURE [dbo].[z_spsBuildUsersCommentsHistory]
	@bHideDeleted	bit =1, 
	@UID		int,
	@UsersID	int=null,
	@CompanyID	int=null
AS
IF @CompanyID Is Null
BEGIN
	SELECT 	UBComHis.Comments,
		UBComHis.DateLastUpdated,
		(	SELECT 	UserName	
			FROM	Users
			WHERE 	UsersID = UBComHis.UpdateByUserID
		) As UpdateByUserName
	FROM UsersHistory UBComHis
	WHERE UBComHis.UsersID = @UsersID
	ORDER BY UBComHis.DateLastUpdated
END
ELSE
BEGIN
	SELECT 	CUBComHis.Comments,
		CUBComHis.DateLastUpdated,
		(	SELECT 	UserName	
			FROM	Users
			WHERE 	UsersID = CUBComHis.UpdateByUserID
		) As UpdateByUserName
	FROM CompanyUsersHistory CUBComHis
	WHERE 	CUBComHis.UsersID = @UsersID
	AND 	CUBComHis.CompanyID = @CompanyID
	ORDER BY CUBComHis.DateLastUpdated
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[z_spsBuildUsersCommentsHistory]  TO [CFUSER]
GO

