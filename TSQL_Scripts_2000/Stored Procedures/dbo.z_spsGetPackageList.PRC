SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spsGetPackageList]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spsGetPackageList]
GO

/******************************************************************
*
*    	Description:	Get Package List
*			
*	
*	Author: 	Brad Skidmore
*	Date: 		2/18/2005
*
******************************************************************/
CREATE PROCEDURE [dbo].[z_spsGetPackageList]
	@UID			int,
	@AssignmentsID		Int=null,
	@bGetHistory		bit=0
AS
BEGIN
	IF @bGetHistory = 0
	BEGIN
		SELECT 
			[PackageID],
			[AssignmentsID],
			[ID],
			[IDAssignments],
			[CreateDate],
			[PackageStatus],
			[Name],
			[Description],
			[Number],
			[SendMe],
			[SentDate],
			[SentToEmail],
			[IsDeleted],
			[DownLoadMe],
			[UpLoadMe],
			[AdminComments],
			dbo.FormatDate([DateLastUpdated],'MM/dd/yyy HHHH:nn') As [DateLastUpdated],
			[UpdateByUserID],
			(
			SELECT Top 1 [USERNAME]
			FROM	Users
			WHERE	[UsersID] = P.[UpdateByUserID]
			) As UpdateByUserName
		FROM 	Package P
		WHERE 	[AssignmentsID] = @AssignmentsID
		AND	[Number] = 1
	END
	IF @bGetHistory = 1
	BEGIN
		SELECT 
			[PackageHistoryID],
			[PackageID],
			[AssignmentsID],
			[ID],
			[IDAssignments],
			[CreateDate],
			[PackageStatus],
			[Name],
			[Description],
			[Number],
			[SendMe],
			[SentDate],
			[SentToEmail],
			[IsDeleted],
			[DownLoadMe],
			[UpLoadMe],
			[AdminComments],
			dbo.FormatDate([DateLastUpdated],'MM/dd/yyy HHHH:nn') As [DateLastUpdated],
			[UpdateByUserID],
			(
			SELECT Top 1[USERNAME]
			FROM	Users
			WHERE	[UsersID] = PH.[UpdateByUserID]
			) As UpdateByUserName
		FROM 	PackageHistory PH
		WHERE 	[AssignmentsID] = @AssignmentsID
		AND	[Number] = 1
		ORDER BY [DateLastUpdated] Desc
	END
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[z_spsGetPackageList]  TO [CFUSER]
GO

