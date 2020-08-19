SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spuInsertActivityLogItem]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spuInsertActivityLogItem]
GO

/******************************************************************
*
*    	Description:	Insert Activity Log item 
*			
*	
*	Author: 	Brad Skidmore
*	Date: 		7/14/2005
******************************************************************/
CREATE PROCEDURE [dbo].[z_spuInsertActivityLogItem]
	@insUpdateByUserID	int=null,
	@insAssignmentsID	int,
	@insActText		varchar(8000),
	@insIsMgrEntry		bit=0,
	@insAdminComments	varchar(1000)=null,
	@insBillingCountID	int=null,
	@insServiceTime		numeric(10, 2)=null,
	@insActDate		datetime=null,
	@insActTime		datetime=null,
	@insPageBreakAfter	bit=0,
	@insBlankPageAfter	bit=0,
	@insBlankRowsAfter	int=0,
	@insIsDeleted		bit=0,
	@insDownLoadMe		bit=1,
	@insUpLoadMe		bit=0,
	@insDateLastUpdated 	datetime=null
AS	
	DECLARE @CurDate			datetime

	--RTActivityLog INSERT
	DECLARE	@RTActivityLogID		int
	DECLARE	@AssignmentsID			int
	DECLARE	@BillingCountID			int
	DECLARE	@ID				int
	DECLARE	@IDAssignments			int
	DECLARE	@IDBillingCount			int
	DECLARE	@ServiceTime			numeric(10, 2)
	DECLARE	@ActDate			datetime
	DECLARE	@ActText			varchar(8000)
	DECLARE	@ActTime			datetime
	DECLARE	@PageBreakAfter			bit
	DECLARE	@BlankPageAfter			bit
	DECLARE	@BlankRowsAfter			int
	DECLARE	@IsMgrEntry			bit
	DECLARE	@IsDeleted			bit
	DECLARE	@DownLoadMe			bit
	DECLARE	@UpLoadMe			bit
	DECLARE	@AdminComments			varchar(1000)
	DECLARE @DateLastUpdated 		datetime
	DECLARE	@UpdateByUserID			int

	--Set the Vars
	SET @CurDate = GetDate()	
	
	--@RTActivityLogID 		This gets set after an insert
	SET @AssignmentsID 		= @insAssignmentsID
	SET @BillingCountID 		= @insBillingCountID	
	--@ID 				This gets set the @RTActivityLogID
	SET @IDAssignments		= @insAssignmentsID	
	SET @IDBillingCount		= @insBillingCountID	
	SET @ServiceTime 		= @insServiceTime			
	IF @insActDate Is Null 
	BEGIN
		SET @ActDate 		=dbo.CleanFromOrToDate(@CurDate, @ActDate, 1)
	END
	ELSE
	BEGIN
		SET @ActDate 		= @insActDate	
	END		
	SET @ActText 			= @insActText						
	IF @insActTime Is Null 
	BEGIN
		SET @ActTime 		= @CurDate
	END
	ELSE
	BEGIN
		SET @ActTime 		= @insActTime	
	END				
	SET @PageBreakAfter 		= @insPageBreakAfter	
	SET @BlankPageAfter 		= @insBlankPageAfter	
	SET @BlankRowsAfter 		= @insBlankRowsAfter	
	SET @IsMgrEntry 		= @insIsMgrEntry		
	SET @IsDeleted 			= @insIsDeleted		
	SET @DownLoadMe 		= @insDownLoadMe			
	SET @UpLoadMe 			= @insUpLoadMe				
	SET @AdminComments 		= @insAdminComments		
	IF @insDateLastUpdated Is Null 
	BEGIN
		SET @DateLastUpdated	= @CurDate
	END
	ELSE
	BEGIN
		SET @DateLastUpdated	= @insDateLastUpdated 	
	END				
	--If the Passed in UID is Null then... Get the UsersID For CFUSER
	IF @insUpdateByUserID Is Null
	BEGIN
		SET @UpdateByUserID = (SELECT [UsersID] FROM USERS WHERE USERNAME = 'CFUSER')
		
	END
	ELSE
	BEGIN
		SET @UpdateByUserID 		= @insUpdateByUserID
	END

	--Insert the Activity Entry
	INSERT INTO RTActivityLog
		(
		[AssignmentsID],
		[BillingCountID],
		[ID],
		[IDAssignments],
		[IDBillingCount],
		[ServiceTime],
		[ActDate],
		[ActText],
		[ActTime],
		[PageBreakAfter],
		[BlankPageAfter],
		[BlankRowsAfter],
		[IsMgrEntry],
		[IsDeleted],
		[DownLoadMe],
		[UpLoadMe],
		[AdminComments],
		[DateLastUpdated],
		[UpdateByUserID] 
		)
	SELECT	
		 @AssignmentsID As [AssignmentsID],
		 @BillingCountID As [BillingCountID],
		 @ID As [ID],
		 @IDAssignments As [IDAssignments],
		 @IDBillingCount As [IDBillingCount],
		 @ServiceTime As [ServiceTime],
		 @ActDate As [ActDate],
		 @ActText As [ActText],
		 @ActTime As [ActTime],
		 @PageBreakAfter As [PageBreakAfter],
		 @BlankPageAfter As [BlankPageAfter],
		 @BlankRowsAfter As [BlankRowsAfter],
		 @IsMgrEntry As [IsMgrEntry],
		 @IsDeleted As [IsDeleted],
		 @DownLoadMe As [DownLoadMe],
		 @UpLoadMe As [UpLoadMe],
		 @AdminComments As [AdminComments],
		 @DateLastUpdated As [DateLastUpdated],
		 @UpdateByUserID As [UpdateByUserID]
	--Need to Update the [ID] to be the Current Identity
	SET @RTActivityLogID = IDENT_CURRENT('RTActivityLog')
	UPDATE RTActivityLog SET [ID] = @RTActivityLogID
	WHERE 	[RTActivityLogID] = @RTActivityLogID
	AND  	[AssignmentsID] = @AssignmentsID

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[z_spuInsertActivityLogItem]  TO [CFUSER]
GO

