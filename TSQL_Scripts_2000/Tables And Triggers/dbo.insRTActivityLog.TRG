SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[insRTActivityLog]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[insRTActivityLog]
GO

/******************************************************************
*
*    	Description:	Allows for snagging Items to do upates elsewhere 
*			and then allow the actual Update to take place.
*			
*	Author: 	Brad Skidmore
*	Date: 		5/25/2005 
*
******************************************************************/
CREATE TRIGGER insRTActivityLog
ON dbo.RTActivityLog
INSTEAD OF INSERT
AS
--5.19.2005 BGS  Farmers XML01 Need to check for certain Activity Updates
--Only do this if one Record assignment at a time is being updated
IF (SELECT COUNT(INS.[ID]) FROM INSERTED INS) = 1
BEGIN
	--Check for Farmers XML01 Activity Updates
	DECLARE	@AssignmentsID	as Int
	DECLARE @sLRFormat 	as varchar(50)
	DECLARE	@ActDate 	as datetime
	DECLARE	@Comment	as varchar(1500)
	DECLARE @CompletedDate 	as datetime
	DECLARE	@FromDate	as datetime
	DECLARE	@ToDate		as datetime
	DECLARE	@UID		as Int
	
	SELECT 	@AssignmentsID =	A.[AssignmentsID],
		@sLRFormat = 		A.[LRFormat],
		@ActDate = 		INS.[ActDate],
		@CompletedDate = 	INS.[ActTime],
		@Comment =		LTrim(RTrim(Convert(varchar(1500),INS.[ActText]))),
		@UID =			INS.[UpdateByUserID]
	FROM 	INSERTED INS 
		INNER JOIN Assignments A ON A.[AssignmentsID] = INS.[AssignmentsID]
	WHERE	A.[LRFormat] = 'V2ECcarFarmers.clsLossXML01'
	
	IF @sLRFormat Is Not Null 
	BEGIN
		--Clean Dates
		SET @FromDate =	dbo.CleanFromOrToDate(@ActDate, @FromDate, 1)
		SET @ToDate = dbo.CleanFromOrToDate(@ActDate, @ToDate, 0)
		--Update the Activity Que table!!!
		UPDATE 	FarmersXML01_ActivityUpdateQue SET
			[Comment] = @Comment,
			[CompletedDate] = @CompletedDate,
			[GMTCreated] = GetUTCDate(),
			[SendToCRN] = 1,
			[DateLastUpdated] = GetDate(),
			[UpdateByUserID] = @UID				
		WHERE	[AssignmentsID] = @AssignmentsID
		AND 	[Type] IN ('First Contact','Field Inspection')
		AND 	[CompletedDate] >= @FromDate
		AND 	[CompletedDate] <= @ToDate
		AND 	[SendToCRN]=1
	END
END

-- Then do the Insert!
INSERT INTO RTActivityLog 	(
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
			INS.[AssignmentsID],
			INS.[BillingCountID],
			INS.[ID],
			INS.[IDAssignments],
			INS.[IDBillingCount],
			INS.[ServiceTime],
			INS.[ActDate],
			INS.[ActText],
			INS.[ActTime],
			INS.[PageBreakAfter],
			INS.[BlankPageAfter],
			INS.[BlankRowsAfter],
			INS.[IsMgrEntry],
			INS.[IsDeleted],
			INS.[DownloadMe],
			INS.[UpLoadMe],
			INS.[AdminComments],
			INS.[DateLastUpdated],
			INS.[UpdateByUserID]
	FROM INSERTED INS

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

