SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spsGetRTAttachmentsInfo]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spsGetRTAttachmentsInfo]
GO

/******************************************************************
*
*    	Description:	RT Attachments Info
*			
*	
*	Author: 	Brad Skidmore
*	Date: 		5/25/2004
*
******************************************************************/
CREATE PROCEDURE [dbo].[z_spsGetRTAttachmentsInfo]
	@bHideDeleted		bit =1, 
	@UID			int,
	@RTAttachmentsID	Int=null,
	@OrderBy		varchar(500)=null,
	@GroupBy		varchar(500)=null,
	@SearchBy		varchar(1000)=null,
--updating the TypeOfLoss Table
	@bDelete		bit=0,
	@bAdd			bit=0,
	@bEdit			bit=0,
--Update variables
	@AssignmentsID		int=null,
	@ID			int=null,
	@IDAssignments		int=null,
	@AttachDate		datetime=null,
	@SortOrder		int=null,
	@Description		varchar(254)=null,
	@AttachName		varchar(50)=null,
	@Attachment		varchar(254)=null,
	@DownloadAttachment	bit=null,
	@UpLoadAttachment	bit=null,
	@IsDeleted		bit=null,
	@DownLoadMe		bit=null,
	@UpLoadMe		bit=null,
	@AdminComments		varchar(1000)=null,
	@DateLastUpdated	datetime=null,
--ViewHistory
	@bViewHistory		bit=NULL,
	@FromDate		varchar(50)=Null,
	@ToDate			varchar(50)=Null,
--IncludeFlags
	@bTableFieldsOnly	bit=0
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

IF @RTAttachmentsID Is Not Null And @bViewHistory is Null
BEGIN
	IF @bDelete=0 And @bAdd=0 And @bEdit=0
	BEGIN
		SELECT 
			[RTAttachmentsID],
			[AssignmentsID],
			[ID] ,
			[IDAssignments],
			[AttachDate] ,
			[SortOrder] ,
			[Description],
			[AttachName] ,
			[Attachment] ,
			[DownloadAttachment],
			[UpLoadAttachment] ,
			[IsDeleted] ,
			[DownLoadMe] ,
			[UpLoadMe] ,
			[AdminComments],
			[DateLastUpdated],
			[UpdateByUserID] 
		FROM	RTAttachments
		WHERE 	[RTAttachmentsID] = @RTAttachmentsID
	END
	ELSE
	BEGIN
		IF @bDelete =1
		BEGIN
			UPDATE RTAttachments SET 	[IsDeleted]=	(CASE 	WHEN [IsDeleted]= 0
										THEN 1
										ELSE 0
										END
									),
							[DateLastUpdated]=@DateLastUpdated ,
							[UpdateByUserID]=@UID
			WHERE RTAttachmentsID = @RTAttachmentsID
		
		END
		ELSE 
		BEGIN
			IF @bAdd=1
			BEGIN
				INSERT INTO RTAttachments(
						[AssignmentsID],
						[ID] ,
						[IDAssignments],
						[AttachDate] ,
						[SortOrder] ,
						[Description],
						[AttachName] ,
						[Attachment] ,
						[DownloadAttachment],
						[UpLoadAttachment] ,
						[IsDeleted] ,
						[DownLoadMe] ,
						[UpLoadMe] ,
						[AdminComments],
						[DateLastUpdated],
						[UpdateByUserID]  
						)
				SELECT 			
						@AssignmentsID As [AssignmentsID],
						@ID As [ID] ,
						@IDAssignments As [IDAssignments],
						@AttachDate As [AttachDate] ,
						@SortOrder As [SortOrder] ,
						@Description As [Description],
						@AttachName As [AttachName] ,
						@Attachment As [Attachment] ,
						@DownloadAttachment As [DownloadAttachment],
						@UpLoadAttachment As [UpLoadAttachment] ,
						@IsDeleted As [IsDeleted] ,
						@DownLoadMe As [DownLoadMe] ,
						@UpLoadMe As [UpLoadMe] ,
						@AdminComments As [AdminComments],
						@DateLastUpdated As [DateLastUpdated],
						@UID As [UpdateByUserID] 
			--Need to update ID if it is null
			Update RTAttachments Set 	[ID] = [RTAttachmentsID],
							[IDAssignments] = [AssignmentsID]
			WHERE [ID] Is Null
					
			END
			ELSE 
			BEGIN
				IF @bEdit=1
				BEGIN
					Update RTAttachments Set 	
							[AssignmentsID] = @AssignmentsID,
							[ID] = @ID ,
							[IDAssignments] = @IDAssignments,
							[AttachDate] = @AttachDate ,
							[SortOrder] = @SortOrder ,
							[Description] = @Description,
							[AttachName] = @AttachName ,
							[Attachment] = @Attachment ,
							[DownloadAttachment] = @DownloadAttachment,
							[UpLoadAttachment] = @UpLoadAttachment ,
							[IsDeleted] = @IsDeleted ,
							[DownLoadMe] = @DownLoadMe ,
							[UpLoadMe] = @UpLoadMe ,
							[AdminComments] = @AdminComments,
							[DateLastUpdated] = @DateLastUpdated  ,
							[UpdateByUserID] = @UID   
					WHERE RTAttachmentsID = @RTAttachmentsID
				END
			END
		END
	END
END
ELSE
BEGIN
	IF @GroupBy Is Not Null
	BEGIN
		SET @sSQL = 'SELECT RetRTAttachments.' + @GroupBy + ' '
	END
	ELSE
	BEGIN
		SET @sSQL = 'SELECT RetRTAttachments.* '
	END	
	SET @sSQL = @sSQL + 'FROM( ' 
	SET @sSQL = @sSQL + 'SELECT '
	SET @sSQL = @sSQL + '[RTAttachmentsID], '
	SET @sSQL = @sSQL + '[AssignmentsID], '
	SET @sSQL = @sSQL + '[ID] , '
	SET @sSQL = @sSQL + '[IDAssignments], '
	SET @sSQL = @sSQL + '[AttachDate] , '
	SET @sSQL = @sSQL + '[SortOrder] , '
	SET @sSQL = @sSQL + '[Description], '
	SET @sSQL = @sSQL + '[AttachName] , '
	SET @sSQL = @sSQL + '[Attachment] , '
	SET @sSQL = @sSQL + '[DownloadAttachment], '
	SET @sSQL = @sSQL + '[UpLoadAttachment] , '
	SET @sSQL = @sSQL + '[IsDeleted] , '
	SET @sSQL = @sSQL + '[DownLoadMe] , '
	SET @sSQL = @sSQL + '[UpLoadMe] , '
	SET @sSQL = @sSQL + '[AdminComments], '
	SET @sSQL = @sSQL + '[DateLastUpdated] , ' 
	IF @bTableFieldsOnly = 0
	BEGIN
		SET @sSQL = @sSQL + '(SELECT	USERNAME '
		SET @sSQL = @sSQL + 'FROM 	USERS '
		SET @sSQL = @sSQL + 'WHERE 	USERSID = S.[UpdateByUserID]) As [UpdateByUserName],  '
	END
	SET @sSQL = @sSQL + '[UpdateByUserID]  '

	IF @bViewHistory Is Null
	BEGIN
		SET @sSQL = @sSQL + 'FROM RTAttachments S '
		SET @sSQL = @sSQL + ') RetRTAttachments '
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
		SET @sSQL = @sSQL + 'FROM RTAttachmentsHistory S '
		SET @sSQL = @sSQL + ') RetRTAttachments '
		SET @sSQL = @sSQL + 'WHERE [DateLastUpdated] >= Convert(datetime,''' + @FromDate + ''') '
		SET @sSQL = @sSQL + 'AND [DateLastUpdated] <= Convert(datetime,''' + @ToDate + ''') '
		IF @AssignmentsID <> -1
		BEGIN
			SET @sSQL = @sSQL + 'AND [AssignmentsID] = ' + Cast(@AssignmentsID as varchar(10)) + ' '
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
				IF @bTableFieldsOnly = 0
				BEGIN
					SET @sSQL = @sSQL + 'ORDER BY [IsDeleted] Asc '
				END
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
	print @sSQL
	EXEC(@sSQL)
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[z_spsGetRTAttachmentsInfo]  TO [CFUSER]
GO

