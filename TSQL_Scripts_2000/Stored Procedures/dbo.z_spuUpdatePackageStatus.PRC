SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spuUpdatePackageStatus]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spuUpdatePackageStatus]
GO

/******************************************************************
*
*    	Description:	Update package Status Approval Rejection, Comments
*			
*	
*	Author: 	Brad Skidmore
*	Date: 		6/11/2005
******************************************************************/
CREATE PROCEDURE [dbo].[z_spuUpdatePackageStatus]
	@UID			int,
	@AssignmentsID		int,
	@PackageNumber		int,
	@SendMe			bit=null,
	@StatusID		int=null,
	@AdminComments		varchar(1000)=null,
	@SentToEmail		varchar(50)=null,
	@SentDate		DateTime=null,
	@ApprovePackageItems	bit=null,
	--Updating Status via at the Document Level
	@PackageItemGUID	varchar(50)=null,
	@bClientUpdate		bit=0,
	@IsClientCoReject	bit=null,
	@ClientCoRejectDate	datetime=null,
	@ClientCoRejectDesc	varchar(500)=null,
	@IsClientCoDelete	bit=null,
	@ClientCoDeleteDate	datetime=null,
	@ClientCoDeleteDesc	varchar(500)=null,
	@IsClientCoApprove	bit=null,
	@ClientCoApproveDate	datetime=null,
	@ClientCoApproveDesc	varchar(500)=null
AS	
	--Vars
	DECLARE @PackageStatus varchar(100)
	DECLARE @DateLastUpdated DateTime
	DECLARE @DateDontSendUntil DateTime --Set this to be 30 minutes into the future

	--If the Client is updating Documents then 
	--Update the Status to CLIENTResponse StatusID 	
	IF @bClientUpdate = 1
	BEGIN
		--Always use CFUSER for USerID
		SET @UID = (SELECT [UsersID] FROM Users WHERE USerName = 'CFUSER')

		IF @IsClientCoReject = 0
		BEGIN
			--Check for the Status Approved string inside the admin comments
			IF @AdminComments Is Not Null
			BEGIN
				IF CHARINDEX('Status: Approved', @AdminComments) > 0
				BEGIN
					SET @StatusID = 	(
								SELECT	[StatusID]
								FROM	Status
								WHERE	[Status]	= 'CLIENTApproved'
								)
				END
			END
			IF @StatusID Is Null
			BEGIN
				SET @StatusID = 	(
							SELECT	[StatusID]
							FROM	Status
							WHERE	[Status]	= 'CLIENTResponse'
							)
			END
		END
		ELSE
		BEGIN
			-- If the Client Is Rejecting then need to 
			-- Give the Client Reject Satus.
			SET @StatusID = 	(
						SELECT	[StatusID]
						FROM	Status
						WHERE	[Status]	= 'REJECT02'
						)
		END
		--Also need to get the AssignmentsID and the PackageID for
		--This Particular Document via the GUID
		SELECT 	@AssignmentsID = PKI.[AssignmentsID],
			@PackageNumber = PKI.[Number]
		FROM	PackageItem PKI
		WHERE	PKI.[PackageItemGUID] 	= @PackageItemGUID
	END
	SET @PackageStatus = 	(
				SELECT	[Description]
				FROM	Status
				WHERE	[StatusID]	= @StatusID
				)

	SET @DateLastUpdated = GetDate()

	--Update the package Table
	IF @bClientUpdate = 1
	BEGIN
		UPDATE Package SET
			[PackageStatus]		= (CASE WHEN @PackageStatus Is Null 	THEN [PackageStatus] 	ELSE @PackageStatus 	END),
			[SendMe] 		= (CASE WHEN @SendMe Is Null 		THEN [SendMe] 		ELSE @SendMe		END),
			[SentDate] 		= (CASE WHEN @SentDate Is Null		THEN [SentDate]		ELSE @SentDate		END),
			[SentToEmail] 		= (CASE WHEN @SentToEmail Is Null	THEN [SentToEmail]	ELSE @SentToEmail	END),
			[DownLoadMe] 		= 1,
			[AdminComments] 	= (CASE WHEN @AdminComments Is Null	THEN [AdminComments]	ELSE @AdminComments	END),
			[DateLastUpdated] 	= @DateLastUpdated,
			[UpdateByUserID] 	= @UID
		WHERE	[AssignmentsID] 	= @AssignmentsID
		AND	[IsDeleted] 		= 0
		AND 	[Number] 		= @PackageNumber

		--When A client Updates need to also update the Assignments Table
		UPDATE 	Assignments SET 
			[StatusID] 		= @StatusID,
			[DownLoadMe] 		= 1,
			[DateLastUpdated] 	= @DateLastUpdated,
			[UpdateByUserID] 	= @UID
		WHERE	[AssignmentsID] 	= @AssignmentsID
	END
	ELSE
	BEGIN
		UPDATE Package SET
			[PackageStatus]		= (CASE WHEN @PackageStatus Is Null 	THEN [PackageStatus] 	ELSE @PackageStatus 	END),
			[SendMe] 		= (CASE WHEN @SendMe Is Null 		THEN [SendMe] 		ELSE @SendMe		END),
			[SentDate] 		= (CASE WHEN @SentDate Is Null		THEN [SentDate]		ELSE @SentDate		END),
			[SentToEmail] 		= (CASE WHEN @SentToEmail Is Null	THEN [SentToEmail]	ELSE @SentToEmail	END),
			[DownLoadMe] 		= 1,
			[AdminComments] 	= (CASE WHEN @AdminComments Is Null	THEN [AdminComments]	ELSE @AdminComments	END),
			[DateLastUpdated] 	= @DateLastUpdated,
			[UpdateByUserID] 	= @UID
		WHERE	[AssignmentsID] 	= @AssignmentsID
		AND	[IsDeleted] 		= 0
		AND 	[Number] 		= @PackageNumber
	END
	
	
	IF @bClientUpdate = 0
	BEGIN
		--Get this PackageID
		DECLARE @PackageID  int
		SET  @PackageID = 	(
					SELECT 	[PackageID]
					FROM	Package
					WHERE	[AssignmentsID] 	= @AssignmentsID
					AND	[IsDeleted] 		= 0
					AND 	[Number] 		= @PackageNumber
					)
		
		--Mark Items in PackageItem as Approved when appropriate
		IF @ApprovePackageItems Is Not Null
		BEGIN
			--Update the numerous package items
			UPDATE PackageItem SET
				[IsCoApprove] 		= @ApprovePackageItems,
				[CoApproveDate] 	= @DateLastUpdated,
				[DownLoadMe] 		= 1,
				[DateLastUpdated] 	= @DateLastUpdated,
				[UpdateByUserID] 	= @UID
			WHERE	[AssignmentsID] 	= @AssignmentsID
			AND	[PackageID]		= @PackageID
			AND	[IsDeleted] 		= 0
		END
		
		--Mark items in packageItems as Sent when appropriate
		IF @SentToEmail Is Not Null And @SentDate Is Not Null
		BEGIN
			--Update the numerous package items
			UPDATE PackageItem SET
				[SendMe]		= 0,
				[SentDate]		= @SentDate,
				[DownLoadMe] 		= 1,
				[DateLastUpdated] 	= @DateLastUpdated,
				[UpdateByUserID] 	= @UID
			WHERE	[AssignmentsID] 	= @AssignmentsID
			AND	[PackageID]		= @PackageID
			AND	[IsDeleted] 		= 0
			AND	[SendMe]		= 1
			AND	[IsClientCoApprove]	= 0
			AND	[IsClientCoDelete]	= 0
		END
	END
	ELSE
	BEGIN
		--Client Updates of Actual Doument Items
		UPDATE PackageItem SET
			[IsClientCoReject] 	=(CASE WHEN @IsClientCoReject 	Is Null 	THEN [IsClientCoReject] 	ELSE @IsClientCoReject 			END),
			[ClientCoRejectDate] 	=(CASE WHEN @ClientCoRejectDate Is Null 	THEN [ClientCoRejectDate] 	ELSE @ClientCoRejectDate		END),
			[ClientCoRejectDesc] 	=(CASE WHEN @ClientCoRejectDesc Is Null 	THEN [ClientCoRejectDesc] 	ELSE Left(@ClientCoRejectDesc,100) 	END),
			[IsClientCoDelete] 	=(CASE WHEN @IsClientCoDelete 	Is Null 	THEN [IsClientCoDelete] 	ELSE @IsClientCoDelete 			END),
			[ClientCoDeleteDate] 	=(CASE WHEN @ClientCoDeleteDate Is Null 	THEN [ClientCoDeleteDate] 	ELSE @ClientCoDeleteDate 		END),
			[ClientCoDeleteDesc] 	=(CASE WHEN @ClientCoDeleteDesc Is Null 	THEN [ClientCoDeleteDesc] 	ELSE Left(@ClientCoDeleteDesc,100) 	END),
			[IsClientCoApprove] 	=(CASE WHEN @IsClientCoApprove 	Is Null 	THEN [IsClientCoApprove] 	ELSE @IsClientCoApprove 		END),
			[ClientCoApproveDate] 	=(CASE WHEN @ClientCoApproveDate Is Null	THEN [ClientCoApproveDate] 	ELSE @ClientCoApproveDate 		END),
			[ClientCoApproveDesc] 	=(CASE WHEN @ClientCoApproveDesc Is Null 	THEN [ClientCoApproveDesc] 	ELSE Left(@ClientCoApproveDesc,100)	END),
			[DownLoadMe] 		= 1,
			[DateLastUpdated] 	= @DateLastUpdated,
			[UpdateByUserID] 	= @UID
		WHERE	[PackageItemGUID] 	= @PackageItemGUID
		
		DECLARE @sClientCoAction Varchar(500)
		IF @IsClientCoReject Is not Null 
		BEGIN
			IF @IsClientCoReject = 1
			BEGIN
				SET @sClientCoAction = @ClientCoRejectDesc
			END			
		END
		IF @IsClientCoDelete Is not Null 
		BEGIN
			IF @IsClientCoDelete = 1
			BEGIN
				SET @sClientCoAction = @ClientCoDeleteDesc
			END			
		END
		IF @sClientCoAction Is Not Null
		BEGIN
			--9.12.2005 BGS When the Client Is Rejecting a Document... also need to update the 
			-- Activity log with this action, and email the assigned manager and the Adjuster email
			-- Associated with this rejected document.
			DECLARE @EmailPackageID 	int
			DECLARE @EmailAssignmentsID 	int
			DECLARE @EmailToMGR 		Varchar(50)
			DECLARE @EmailCCAdj 		Varchar(50)
			DECLARE @EmailSubject		varchar(255)
			DECLARE @EmailMess		varchar(500)
			DECLARE @EmailPackageItemIDList varchar(800)
			DECLARE	@EmailItemCLIENTNUM	varchar(50)
			DECLARE @EmailItemInsured	Varchar(50)
			DECLARE @EmailItemAdjuster	varchar(100)
			DECLARE @EmailItemAdjFName	varchar(50)
			DECLARE @EmailItemAdjLName	varchar(50)
			DECLARE @EmailItemSortOrder	varchar(10)
			DECLARE @EmailItemName		varchar(50)				
			DECLARE	@EmailItemDesc		varchar(255)

			--First get the Assignments Id (will be used to lookup adj and mgr email)
			--And the packageID via lookup using GUID
			SELECT 	@EmailPackageID 	= [PackageID], 
				@EmailAssignmentsID 	= [AssignmentsID],
				@EmailPackageItemIDList = Cast([PackageItemID] As Varchar(20)),
				@EmailItemSortOrder	= Cast([SortOrder] As Varchar(10)),
				@EmailItemName		= [Name],
				@EmailItemDesc		= [Description]
			FROM 	PackageItem 		
			WHERE 	[PackageItemGUID] 	= @PackageItemGUID
			
			--Now Get the Adjuster And Manager Emails
			SELECT 	
				@EmailCCAdj = (SELECT [Email] From USERS WHERE [UsersID] = CCAS.UsersID),
				@EmailToMGR = (SELECT [Email] From USERS WHERE [UsersID] = URptMgr.[ReportsToUsersID]),
				@EmailItemCLIENTNUM = A.[CLIENTNUM],
				@EmailItemInsured = A.[Insured],
				@EmailItemAdjuster = (SELECT [UserName] From Users WHERE [UsersID] = CCAS.UsersID),
				@EmailItemAdjFName = (SELECT [FirstName] From Users WHERE [UsersID] = CCAS.UsersID),
				@EmailItemAdjLName = (SELECT [LastName] From Users WHERE [UsersID] = CCAS.UsersID)
			FROM	Assignments A
				Inner join ClientCoAdjusterSpec CCAS ON A.AdjusterSpecID = CCAS.ClientCoAdjusterSpecID
				Inner join Adjuster ADJ ON CCAS.[ClientCompanyID] = ADJ.[CompanyID] And CCAS.[UsersID] = ADJ.[UsersID]
				Inner Join UserReportsToManager URptMgr ON ADJ.[CompanyID] = URptMgr.[CompanyID] And  ADJ.[UsersID] = URptMgr.[UsersID] And URptMgr.[Active] = 1
			Where  A.[AssignmentsID] = @EmailAssignmentsID
			
			IF @EmailToMGR Is null
			BEGIN
				SET @EmailToMGR = ''
			END
			IF @EmailCCAdj Is Null
			BEGIN
				SET @EmailCCAdj = ''
			END

			-- Build the Subject
			IF @IsClientCoDelete Is not Null
			BEGIN
				IF @IsClientCoDelete = 1
				BEGIN
					SET @EmailSubject = 'CLIENT DELETED: ' + @EmailItemCLIENTNUM 
				END
			END
			IF @IsClientCoReject Is not Null
			BEGIN
				IF @IsClientCoReject = 1
				BEGIN
					SET @EmailSubject = 'CLIENT REJECTED: ' + @EmailItemCLIENTNUM
				END 
			END
			
			SET @EmailSubject = @EmailSubject +  ' DOC: ' + @EmailItemSortOrder + ' - ' + @EmailItemName + ' - ' + @EmailItemDesc
			SET @EmailSubject = @EmailSubject +  ' ADJ: ' + @EmailItemAdjuster 
	 		
			
			IF @IsClientCoDelete Is not Null
			BEGIN
				IF @IsClientCoDelete = 1
				BEGIN
					SET @EmailMess = 'CLIENT DELETED: ' + @EmailItemCLIENTNUM + CHAR(13) + CHAR(10)
				END
			END
			IF @IsClientCoReject Is not Null
			BEGIN
				IF @IsClientCoReject = 1
				BEGIN
					SET @EmailMess = 'CLIENT REJECTED: ' + @EmailItemCLIENTNUM + CHAR(13) + CHAR(10)
				END 
			END
			SET @EmailMess = @EmailMess + 'INSURED: ' + @EmailItemInsured + CHAR(13) + CHAR(10)
			SET @EmailMess = @EmailMess + 'ADJ: ' + @EmailItemAdjFName + ' ' + @EmailItemAdjLName + ' (' + @EmailItemAdjuster + ') - ' + @EmailCCAdj + CHAR(13) + CHAR(10)
			SET @EmailMess = @EmailMess + 'DOC: ' + @EmailItemSortOrder + CHAR(13) + CHAR(10)
			SET @EmailMess = @EmailMess + 'Name: ' + @EmailItemName + CHAR(13) + CHAR(10)
			SET @EmailMess = @EmailMess + 'Desc: ' + @EmailItemDesc + CHAR(13) + CHAR(10)
			SET @EmailMess = @EmailMess + '[-------CLIENT MESSAGE------]'  + CHAR(13) + CHAR(10) + CHAR(13) + CHAR(10)
			
			SET @EmailMess = @EmailMess + @sClientCoAction
			--When sending an email containing a rejection or deletion...
			-- need to set a delay using @DateDontSendUntil  this will
			-- allow for subsequent rejections under the same assignment to
			--be added to the same email waiting in the message queue.
			-- set this to 30 minutes into the future  
			SET @DateDontSendUntil = DateAdd(mi, 30, GetDate())

			EXEC spuInsertPackageEmailQueueItem
			@EmailPackageID,			--@insPackageID		int,
			@EmailAssignmentsID,			--@insAssignmentsID	int,
			'[ADMIN SENDING EMAIL]',		--@insPackageStatus	varchar(100)='',
			@EmailToMGR,				--@insEmailTo		varchar(255)='',
			@EmailCCAdj,				--@insEmailCC		varchar(255)='',
			'EberlsClaimSystem@eberls.com',		--@insEmailBCC		varchar(255)='',
			@EmailSubject,				--@insEmailSubject	varchar(255)='',
			@EmailMess,				--@insEmailMess		varchar(500)='',
			@EmailPackageItemIDList, 		--@insPackageItemIDList	varchar(8000)='',
			@UID,					--@insUpdateByUserID 	int,
			@DateDontSendUntil			--datetime=null

			--after entered into queue, Insert Activity Entry for it too !
			EXEC z_spuInsertActivityLogItem
			@UID,			-- @insUpdateByUserID	int=null,
			@EmailAssignmentsID,	-- @insAssignmentsID	int,
			@EmailMess
		END
	END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[z_spuUpdatePackageStatus]  TO [CFUSER]
GO

