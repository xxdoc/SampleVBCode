SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spsGetECFTPUsersHistory]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spsGetECFTPUsersHistory]
GO


/******************************************************************
*
*    	Description:	Get Easy Claim FTP Users History
*			
*	
*	Author: 	Brad Skidmore
*	Date: 		4/20/2005
*
******************************************************************/
CREATE PROCEDURE [dbo].[z_spsGetECFTPUsersHistory]
	@UID			int,
	@UsersID		Int=null,
	@DaysAgo		int=7,
	@UseHistoryTable	bit=1
AS
BEGIN
	IF @UseHistoryTable = 1 
	BEGIN
		SELECT 
			[UsersHistoryID] ,
			[UsersID],
			[UserName],
			[PassWord],
			[FirstName],
			[LastName],
			[SSN],
			[Email],
			[ContactPhone],
			[EmergencyPhone],
			[Address],
			[City],
			[State],
			[Zip],
			[ZIP4],
			[OtherPostCode],
			[Active],
			[ActiveDate],
			[InactiveDate],
			[SecurityLevel],
			[Comments],
			[DateLastUpdated],
			[UpdateByUserID]
		FROM 	UsersHistory
		WHERE 	[UsersID] = @UsersID 	
		AND [comments] like 'Easy Claim Update%'
		AND [DateLastUpdated] >= getdate()- @DaysAgo
		Order By [DateLastUpdated]
	END
	ELSE
	BEGIN
		SELECT 
			[UsersID] As [UsersHistoryID] ,
			[UsersID],
			[UserName],
			[PassWord],
			[FirstName],
			[LastName],
			[SSN],
			[Email],
			[ContactPhone],
			[EmergencyPhone],
			[Address],
			[City],
			[State],
			[Zip],
			[ZIP4],
			[OtherPostCode],
			[Active],
			[ActiveDate],
			[InactiveDate],
			[SecurityLevel],
			[Comments],
			[DateLastUpdated],
			[UpdateByUserID]
		FROM 	Users
		WHERE 	[UsersID] = @UsersID 	
		AND [comments] like 'Easy Claim Update%'
		AND [DateLastUpdated] >= getdate()- @DaysAgo
		Order By [DateLastUpdated]
	END
END


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[z_spsGetECFTPUsersHistory]  TO [CFUSER]
GO

