SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spsGetUserProfileInfo]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spsGetUserProfileInfo]
GO

/******************************************************************
*
*    	Description:	User Profile Info
*			
*	
*	Author: 	Brad Skidmore
*	Date: 		09/29/2003
*
******************************************************************/


CREATE PROCEDURE [dbo].[z_spsGetUserProfileInfo]
	@bHideDeleted		bit =1, 
	@UID			int,
	@UserProfileID		Int=null,
	@OrderBy		varchar(500)=null,
	@GroupBy		varchar(500)=null,
	@SearchBy		varchar(1000)=null,
--updating the Company Table
	@bDelete		bit=0,
	@bAdd			bit=0,
	@bEdit			bit=0,
--Update variables
	@TableName 		varchar(30)=null,
	@Description		varchar(255)=null,
	@SortOrder		int=null,
	@DateLastUpdated	datetime=Null,
--ViewHistory
	@bViewHistory		bit=NULL,
	@FromDate		varchar(50)=Null,
	@ToDate			varchar(50)=Null

AS
DECLARE @sSQL Varchar(8000)
DECLARE @CurDate 		DateTime
SET @CurDate = GetDate()
--See if the DateLastupdated is null
IF IsDate(@DateLastUpdated) =0
BEGIN
	SET @DateLastUpdated = GetDate()
END
--Clean Dates
SET @FromDate =			dbo.CleanFromOrToDate(@CurDate, @FromDate, 1)
SET @ToDate = 			dbo.CleanFromOrToDate(@CurDate, @ToDate, 0)

IF @UserProfileid Is Not Null And @bViewHistory is Null
BEGIN
	IF @bDelete=0 And @bAdd=0 And @bEdit=0
	BEGIN
		SELECT 
			[UserProfileID] ,
			[TableName],
			[Description],
			[Active] ,
			[SortOrder],
			[DateLastUpdated],
			[UpdateByUserID] 
		FROM	UserProfile
		WHERE 	[UserProfileID] = @UserProfileID
		AND 	[Active] =1
		ORDER BY [Active] Desc, [TableName]
	END
	ELSE
	BEGIN
		IF @bDelete =1
		BEGIN
			UPDATE UserProfile SET 	[Active]=	(CASE 	WHEN [Active]= 0
									THEN 1
									ELSE 0
									END
								),
						[DateLastUpdated]=@DateLastUpdated ,
						[UpdateByUserID]=@UID
			WHERE UserProfileID = @UserProfileid
		
		END
		ELSE 
		BEGIN
			IF @bAdd=1
			BEGIN
				INSERT INTO UserProfile(
							[TableName],
							[Description],
							[Active] ,
							[SortOrder],
							[DateLastUpdated],
							[UpdateByUserID] 
						)
				SELECT 		@TableName As [TableName],
							@Description As [Description],
							1 As [Active] ,
							@SortOrder As [SortOrder],
							@DateLastUpdated As [DateLastUpdated],
							@UID As [UpdateByUserID] 
			END
			ELSE 
			BEGIN
				IF @bEdit=1
				BEGIN
					Update UserProfile Set 	[TableName] = @TableName ,
								[Description] = @Description,
								[SortOrder] = @SortOrder,
								[DateLastUpdated] = @DateLastUpdated ,
								[UpdateByUserID] = @UID 
					WHERE UserProfileID = @UserProfileID
				END
			END
		END
	END
END
ELSE
BEGIN
	IF @GroupBy Is Not Null
	BEGIN
		SET @sSQL = 'SELECT RetUserProfile.' + @GroupBy + ' '
	END
	ELSE
	BEGIN
		SET @sSQL = 'SELECT RetUserProfile.* '
	END	
	SET @sSQL = @sSQL + 'FROM( ' 
	SET @sSQL = @sSQL + 'SELECT '
	SET @sSQL = @sSQL + '[UserProfileID], '
	SET @sSQL = @sSQL + '[TableName] , '
	SET @sSQL = @sSQL + '[Description] , '
	SET @sSQL = @sSQL + '[Active] , '
	SET @sSQL = @sSQL + '[SortOrder], '
	SET @sSQL = @sSQL + '[DateLastUpdated] , '
	SET @sSQL = @sSQL + '(SELECT	USERNAME '
	SET @sSQL = @sSQL + 'FROM 		USERS '
	SET @sSQL = @sSQL + 'WHERE 	USERSID = UP.[UpdateByUserID]) As [UpdateByUserName],  '
	SET @sSQL = @sSQL + '[UpdateByUserID]  '


	
	IF @bViewHistory Is Null
	BEGIN
		SET @sSQL = @sSQL + 'FROM UserProfile UP '
		SET @sSQL = @sSQL + ') RetUserProfile '
		IF @bHideDeleted = 1
		BEGIN
			SET @sSQL = @sSQL + 'WHERE Active = 1 '
		END
		ELSE
		BEGIN
			SET @sSQL = @sSQL + 'WHERE Active Is Not Null '
		END
		--Searching
		IF @SearchBy Is Not Null
		BEGIN
			SET @sSQL = @sSQL + @SearchBy + ' '
		END
	END
	ELSE
	BEGIN
		SET @sSQL = @sSQL + 'FROM UserProfileHistory UP '
		SET @sSQL = @sSQL + ') RetUserProfile '
		SET @sSQL = @sSQL + 'WHERE [DateLastUpdated] >= Convert(datetime,''' + @FromDate + ''') '
		SET @sSQL = @sSQL + 'AND [DateLastUpdated] <= Convert(datetime,''' + @ToDate + ''') '
		IF @UserProfileid <> -1
		BEGIN
			SET @sSQL = @sSQL + 'AND [UserProfileID] = ' + Cast(@UserProfileID as varchar(10)) + ' '
		END
		--Searching
		IF @SearchBy Is Not Null
		BEGIN
			SET @sSQL = @sSQL + @SearchBy + ' '
		END
	END
	IF @bViewHistory Is Null
	BEGIN
		--Grouping
		IF @GroupBy Is Not Null
		BEGIN
			SET @sSQL = @sSQL + 'GROUP BY ' + @GroupBy + ' '			
		END
		---Sorting 
		IF @GroupBy Is Null
		BEGIN
			IF @OrderBy Is Null OR @OrderBy = 'Null'
			BEGIN
				SET @sSQL = @sSQL + 'ORDER BY [Active] Desc, [SortOrder] '
			END
			ELSE
			BEGIN
				SET @sSQL = @sSQL + 'ORDER BY  [Active] Desc, ' + @OrderBy + ' '
			END
		END
	END
	ELSE
	BEGIN
		--Grouping
		IF @GroupBy Is Not Null
		BEGIN
			SET @sSQL = @sSQL + 'GROUP BY ' + @GroupBy + ' '			
		END
		---Sorting
		IF @GroupBy Is Null
		BEGIN
			SET @sSQL = @sSQL + 'ORDER BY  ' + @OrderBy + ' '
		END
	END
	EXEC(@sSQL)
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[z_spsGetUserProfileInfo]  TO [CFUSER]
GO

