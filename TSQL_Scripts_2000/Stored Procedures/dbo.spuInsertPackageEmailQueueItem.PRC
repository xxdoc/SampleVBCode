SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spuInsertPackageEmailQueueItem]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[spuInsertPackageEmailQueueItem]
GO

/******************************************************************
*
*    	Description:	Get Packges Qued for Email...
*			To be processed
*	
*	Author: 	Brad Skidmore
*	Date: 		07/28/2005
*
******************************************************************/
CREATE PROCEDURE [dbo].[spuInsertPackageEmailQueueItem] 
	@insPackageID		int,
	@insAssignmentsID	int,
	@insPackageStatus	varchar(100)='',
	@insEmailTo		varchar(255)='',
	@insEmailCC		varchar(255)='',
	@insEmailBCC		varchar(255)='',
	@insEmailSubject	varchar(255)='',
	@insEmailMess		varchar(500)='',
	@insPackageItemIDList	varchar(8000)='',
	@insUpdateByUserID 	int,
	@insDateDontSendUntil	datetime=null
AS	
	DECLARE @CurDate			As datetime
	DECLARE @EditPackageEmailQueueID	As int	

	DECLARE @PackageID 		As int
	DECLARE @AssignmentsID 		As int
	DECLARE @PackageStatus		As varchar(100)
	DECLARE @EmailTo		As varchar(255)
	DECLARE @EmailCC		As varchar(255)
	DECLARE @EmailBCC		As varchar(255)
	DECLARE @EmailSubject		As varchar(255)
	DECLARE @EmailMess		As varchar(500)
	DECLARE @PackageItemIDList	As varchar(8000)
	DECLARE @DateLastUpdated	As datetime
	DECLARE @UpdateByUserID		As int
	DECLARE @DateDontSendUntil	AS datetime
	--Set the Vars
	SET @CurDate = GetDate()	
	
	--@RTActivityLogID 		This gets set after an insert
	SET @PackageID 			= @insPackageID
	SET @AssignmentsID 		= @insAssignmentsID	
	SET @PackageStatus		= @insPackageStatus
	SET @EmailTo			= @insEmailTo
	SET @EmailCC			= @insEmailCC
	SET @EmailBCC			= @insEmailBCC
	SET @EmailSubject		= @insEmailSubject
	SET @EmailMess			= @insEmailMess
	SET @PackageItemIDList		= @insPackageItemIDList	
	SET @DateLastUpdated		= @CurDate		
	--If the Passed in UID is Null then... Get the UsersID For CFUSER
	IF @insUpdateByUserID Is Null
	BEGIN
		SET @UpdateByUserID = (SELECT [UsersID] FROM USERS WHERE USERNAME = 'CFUSER')
		
	END
	ELSE
	BEGIN
		SET @UpdateByUserID 		= @insUpdateByUserID
	END
	--If the DateDontSendUntil is null then set it to the current date
	IF @insDateDontSendUntil Is Null
	BEGIN
		SET @DateDontSendUntil = @CurDate
	END
	ELSE
	BEGIN
		SET @DateDontSendUntil = @insDateDontSendUntil
	END

	-- If the @DateDontSendUntil > current date that means this item has a delay
	-- If this Item has a delay, and @PackageStatus = '[ADMIN SENDING EMAIL]' 
	-- that means there may be a previous item under the same AssignmentsID.  
	-- If so then Need to Add on to the exisitng pending item.
	-- This will mean only one email will be sent contiaing multiple items
	IF @DateDontSendUntil > @CurDate And @PackageStatus = '[ADMIN SENDING EMAIL]'
	BEGIN
		SELECT 	TOP 1 @EditPackageEmailQueueID = [PackageEmailQueueID]
		FROM	PackageEmailQueue
		WHERE	[AssignmentsID] = @AssignmentsID
		AND	[DateDontSendUntil] > [DateLastUpdated]
		AND 	[PackageStatus] = '[ADMIN SENDING EMAIL]'
		Order By [DateLastUpdated] Desc
	END

	--Insert the Package Email Queue Item
	--If the @EditAssignmentsID is null that means there is nothing previously enetered for this
	--particular item.  So do the insert as usual.
	IF @EditPackageEmailQueueID Is Null
	BEGIN
		INSERT INTO PackageEmailQueue
		(
		[PackageID],
		[AssignmentsID],
		[PackageStatus],
		[EmailTo],
		[EmailCC],
		[EmailBCC],
		[EmailSubject],
		[EmailMess],
		[PackageItemIDList],
		[DateLastUpdated],
		[UpdateByUserID],
		[DateDontSendUntil]
		)
	SELECT	
		@PackageID 		As [PackageID],
		@AssignmentsID 		As [AssignmentsID],
		@PackageStatus 		As [PackageStatus],
		@EmailTo 		As [EmailTo],
		@EmailCC 		As [EmailCC],
		@EmailBCC 		As [EmailBCC],
		@EmailSubject 		As [EmailSubject],
		@EmailMess 		As [EmailMess],
		@PackageItemIDList 	As [PackageItemIDList],
		@DateLastUpdated 	As [DateLastUpdated],
		@UpdateByUserID 	As [UpdateByUserID],
		@DateDontSendUntil	As [DateDontSendUntil]
	END
	ELSE --Otherwise do an edit to the exisitng item
	BEGIN
		UPDATE PackageEmailQueue SET
			[PackageID] 		= @PackageID,
			[AssignmentsID] 	= @AssignmentsID,
			[PackageStatus]		= @PackageStatus,
			[EmailTo]		= @EmailTo,
			[EmailCC]		= @EmailCC,
			[EmailBCC]		= @EmailBCC,
			[EmailSubject]		= [EmailSubject] + '*',
			[EmailMess]		= [EmailMess]+ CHAR(13) + CHAR(10) + @EmailSubject,
			[PackageItemIDList]	= [PackageItemIDList] + ', ' + @PackageItemIDList,
			[DateLastUpdated]	= @DateLastUpdated,
			[UpdateByUserID]	= @UpdateByUserID,
			[DateDontSendUntil]	= @DateDontSendUntil
		WHERE	[PackageEmailQueueID]	= @EditPackageEmailQueueID
	END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[spuInsertPackageEmailQueueItem]  TO [CFUSER]
GO

