SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spsGetStatusInfo]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spsGetStatusInfo]
GO

/******************************************************************
*
*    	Description:	Status Info
*			
*	
*	Author: 	Brad Skidmore
*	Date: 		11/10/2003
*
******************************************************************/


CREATE PROCEDURE [dbo].[z_spsGetStatusInfo]
	@bHideDeleted		bit =1, 
	@UID			int,
	@StatusID		Int=null,
	@OrderBy		varchar(500)=null,
	@GroupBy		varchar(500)=null,
	@SearchBy		varchar(1000)=null,
--updating the TypeOfLoss Table
	@bDelete		bit=0,
	@bAdd			bit=0,
	@bEdit			bit=0,
--Update variables
	@StatusAlias		varchar(20)=null,
	@Status			varchar(20)=null,
	@Description		varchar(100)=null,
	@AdminComments		varchar(1000)=null,
	@IsDeleted		bit=null,
	@DateLastUpdated 	datetime =null,
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

IF @StatusID Is Not Null And @bViewHistory is Null
BEGIN
	IF @bDelete=0 And @bAdd=0 And @bEdit=0
	BEGIN
		SELECT 
			[StatusID],
			[StatusAlias],
			[Status],
			[Description],
			[AdminComments],
			[IsDeleted],
			[DateLastUpdated],
			[UpdateByUserID]
		FROM	Status
		WHERE 	[StatusID] = @StatusID
	END
	ELSE
	BEGIN
		IF @bDelete =1
		BEGIN
			UPDATE Status SET 	[IsDeleted]=	(CASE 	WHEN [IsDeleted]= 0
									THEN 1
									ELSE 0
									END
								),
						[DateLastUpdated]=@DateLastUpdated ,
						[UpdateByUserID]=@UID
			WHERE StatusID = @StatusID
		
		END
		ELSE 
		BEGIN
			IF @bAdd=1
			BEGIN
				INSERT INTO Status(
							[StatusAlias],
							[Status],
							[Description],
							[AdminComments],
							[IsDeleted],
							[DateLastUpdated],
							[UpdateByUserID]
						)
				SELECT 			
							@StatusAlias As [StatusAlias],
							@Status As [Status],
							@Description As [Description],
							@AdminComments As [AdmninComments],
							@IsDeleted As [IsDeleted] ,
							@DateLastUpdated As [DateLastUpdated],
							@UID As [UpdateByUserID] 
					
			END
			ELSE 
			BEGIN
				IF @bEdit=1
				BEGIN
					Update Status Set 	
								[StatusAlias] = @StatusAlias ,
								[Status] = @Status ,
								[Description] = @Description ,
								[AdminComments] = @AdminComments ,
								[IsDeleted] = @IsDeleted   ,
								[DateLastUpdated] = @DateLastUpdated  ,
								[UpdateByUserID] = @UID   
					WHERE 	StatusID = @StatusID
				END
			END
		END
	END
END
ELSE
BEGIN
	IF @GroupBy Is Not Null
	BEGIN
		SET @sSQL = 'SELECT RetStatus.' + @GroupBy + ' '
	END
	ELSE
	BEGIN
		SET @sSQL = 'SELECT RetStatus.* '
	END	
	SET @sSQL = @sSQL + 'FROM( ' 
	SET @sSQL = @sSQL + 'SELECT '
	SET @sSQL = @sSQL + '[StatusID], '
	SET @sSQL = @sSQL + '[StatusAlias], '
	SET @sSQL = @sSQL + '[Status], '
	SET @sSQL = @sSQL + '[Description], '
	SET @sSQL = @sSQL + '[AdminComments], '
	SET @sSQL = @sSQL + '[IsDeleted] , '
	SET @sSQL = @sSQL + '[DateLastUpdated] , '
	SET @sSQL = @sSQL + '(SELECT	USERNAME '
	SET @sSQL = @sSQL + 'FROM 		USERS '
	SET @sSQL = @sSQL + 'WHERE 	USERSID = S.[UpdateByUserID]) As [UpdateByUserName],  '
	SET @sSQL = @sSQL + '[UpdateByUserID]  '

	IF @bViewHistory Is Null
	BEGIN
		SET @sSQL = @sSQL + 'FROM Status S '
		SET @sSQL = @sSQL + ') RetStatus '
		IF @bHideDeleted = 1
		BEGIN
			SET @sSQL = @sSQL + 'WHERE IsDeleted = 0 '
		END
		ELSE
		BEGIN
			SET @sSQL = @sSQL + 'WHERE IsDeleted Is Not Null '
		END
		--Searching
		IF @SearchBy Is Not Null
		BEGIN
			SET @sSQL = @sSQL + @SearchBy + ' '
		END
	END
	ELSE
	BEGIN
		SET @sSQL = @sSQL + 'FROM StatusHistory S '
		SET @sSQL = @sSQL + ') RetStatus '
		SET @sSQL = @sSQL + 'WHERE [DateLastUpdated] >= Convert(datetime,''' + @FromDate + ''') '
		SET @sSQL = @sSQL + 'AND [DateLastUpdated] <= Convert(datetime,''' + @ToDate + ''') '
		IF @StatusID <> -1
		BEGIN
			SET @sSQL = @sSQL + 'AND [StatusID] = ' + Cast(@Statusid as varchar(10)) + ' '
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
				SET @sSQL = @sSQL + 'ORDER BY [IsDeleted] Asc, [StatusAlias] '
			END
			ELSE
			BEGIN
				SET @sSQL = @sSQL + 'ORDER BY  [IsDeleted] Asc, ' + @OrderBy + ' '
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

GRANT  EXECUTE  ON [dbo].[z_spsGetStatusInfo]  TO [CFUSER]
GO

