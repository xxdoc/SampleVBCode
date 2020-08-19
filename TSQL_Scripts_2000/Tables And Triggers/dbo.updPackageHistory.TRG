SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[updPackageHistory]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[updPackageHistory]
GO

CREATE TRIGGER updPackageHistory
ON dbo.Package
INSTEAD OF UPDATE
AS
-- Check for any changes to critical items.  If there
-- are no changes then don't create a friggin history of same data
IF (SELECT COUNT(INS.[ID]) FROM INSERTED INS) = 1
BEGIN
	--Check each critical field to see if it has changed
	--If any of them have then allow the History write
	--If not then don't allow the history
	DECLARE @bAllowHistory		bit
	
	--	INSERTED
	DECLARE @InsAssignmentsID 	int
	DECLARE @InsID 			int
	DECLARE @InsIDAssignments 	int
	DECLARE @InsCreateDate 		DateTime
	DECLARE @InsPackageStatus 	varchar(100)
	DECLARE @InsName 		varchar(50)
	DECLARE @InsDescription 	varchar(50)
	DECLARE @InsNumber 		int
	DECLARE @InsSendMe		bit
	DECLARE @InsSentDate 		DateTime
	DECLARE @InsSentToEmail 	varchar(50)
	DECLARE	@InsIsDeleted		bit
	DECLARE	@InsAdminComments	varchar(1000)
	--	DELETED
	DECLARE @DelAssignmentsID 	int
	DECLARE @DelID 			int
	DECLARE @DelIDAssignments 	int
	DECLARE @DelCreateDate 		DateTime
	DECLARE @DelPackageStatus 	varchar(100)
	DECLARE @DelName 		varchar(50)
	DECLARE @DelDescription 	varchar(50)
	DECLARE @DelNumber 		int
	DECLARE @DelSendMe		bit
	DECLARE @DelSentDate 		DateTime
	DECLARE @DelSentToEmail 	varchar(50)
	DECLARE	@DelIsDeleted		bit
	DECLARE	@DelAdminComments	varchar(1000)

	--Select the Friggin Inserted
	SELECT 
		@InsAssignmentsID 	= INS.[AssignmentsID],
		@InsID  		= INS.[ID],			
		@InsIDAssignments 	= INS.[IDAssignments],
		@InsCreateDate 	 	= INS.[CreateDate],	
		@InsPackageStatus 	= INS.[PackageStatus],
		@InsName 	 	= INS.[Name],
		@InsDescription 	= INS.[Description],
		@InsNumber 	 	= INS.[Number],	
		@InsSendMe	 	= INS.[SendMe],	
		@InsSentDate 	 	= INS.[SentDate],	
		@InsSentToEmail 	= INS.[SentToEmail],
		@InsIsDeleted	 	= INS.[IsDeleted],	
		@InsAdminComments	= INS.[AdminComments]
	FROM INSERTED INS
	
	--Select the Friggin Deleted
	SELECT 
		@DelAssignmentsID 	= DEL.[AssignmentsID],
		@DelID  		= DEL.[ID],			
		@DelIDAssignments 	= DEL.[IDAssignments],
		@DelCreateDate 	 	= DEL.[CreateDate],	
		@DelPackageStatus 	= DEL.[PackageStatus],
		@DelName 	 	= DEL.[Name],	
		@DelDescription 	= DEL.[Description],
		@DelNumber 	 	= DEL.[Number],	
		@DelSendMe	 	= DEL.[SendMe],	
		@DelSentDate 	 	= DEL.[SentDate],	
		@DelSentToEmail 	= DEL.[SentToEmail],
		@DelIsDeleted	 	= DEL.[IsDeleted],	
		@DelAdminComments	= DEL.[AdminComments]
	FROM DELETED DEL
	
	IF 	@InsAssignmentsID <> @DelAssignmentsID
		OR
		@InsID <> @DelID
		OR			
		@InsIDAssignments <> @DelIDAssignments
		OR
		@InsCreateDate <> @DelCreateDate
		OR
		@InsPackageStatus <> @DelPackageStatus
		OR
		@InsName <> @DelName
		OR
		@InsDescription <> @DelDescription
		OR
		@InsNumber <> @DelNumber
		OR
		@InsSendMe <> @DelSendMe
		OR
		@InsSentDate <> @DelSentDate
		OR
		@InsSentToEmail <> @DelSentToEmail
		OR
		@InsIsDeleted <> @DelIsDeleted
		OR
		@InsAdminComments <> @DelAdminComments
	BEGIN
		SET @bAllowHistory = 1
	END
	ELSE
	BEGIN
		SET @bAllowHistory = 0
	END
	--If this one little measly flag is on then Allow the History
	IF @bAllowHistory = 1
	BEGIN
		INSERT INTO PackageHistory
		SELECT del.* 
		FROM DELETED del
	END
END
ELSE
BEGIN
	--IF more than one record is being updated then allow the History to occur
	--without checking every stinking one
	INSERT INTO PackageHistory
	SELECT del.* 
	FROM DELETED del
END

--Allow the Actual Update to occur
Update Package SET
	[AssignmentsID]		= INS.[AssignmentsID],
	[ID]			= INS.[ID],
	[IDAssignments]		= INS.[IDAssignments],
	[CreateDate]		= INS.[CreateDate],
	[PackageStatus]		= INS.[PackageStatus],
	[Name]			= INS.[Name],
	[Description]		= INS.[Description],
	[Number]		= INS.[Number],
	[SendMe]		= INS.[SendMe],
	[SentDate]		= INS.[SentDate],
	[SentToEmail]		= INS.[SentToEmail],
	[IsDeleted]		= INS.[IsDeleted],
	[DownLoadMe]		= INS.[DownLoadMe],
	[UpLoadMe]		= INS.[UpLoadMe],
	[AdminComments]		= INS.[AdminComments],
	[DateLastUpdated]	= INS.[DateLastUpdated],
	[UpdateByUserID]	= INS.[UpdateByUserID]
FROM Package PKG INNER JOIN INSERTED INS ON PKG.PackageID = INS.PackageID

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

