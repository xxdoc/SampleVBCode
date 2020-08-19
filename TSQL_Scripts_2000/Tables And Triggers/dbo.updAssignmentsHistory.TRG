SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[updAssignmentsHistory]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[updAssignmentsHistory]
GO

--Because of the Restriction on Triggers when Ntext fields ...
/*
In a DELETE, INSERT, or UPDATE trigger, SQL Server does not allow text, ntext, or image column references 
in the inserted and deleted tables if the compatibility level is equal to 70. The text, ntext, and image values 
in the inserted and deleted tables cannot be accessed. To retrieve the new value in either an INSERT or UPDATE 
trigger, join the inserted table with the original update table. When the compatibility level is 65 or lower, 
null values are returned for inserted or deleted text, ntext, or image columns that allow null values; zero-length 
strings are returned if the columns are not nullable. 
If the compatibility level is 80 or higher, SQL Server allows the update of text, ntext, or image columns through 
the INSTEAD OF trigger on tables or views.
*/
--Use Instead OF to get around the above restriction
CREATE TRIGGER updAssignmentsHistory
ON dbo.Assignments
INSTEAD OF UPDATE
AS

DECLARE @bAllowHistory 			bit
DECLARE @CLIENTResponseStatusID 	int
DECLARE @REJECT02StatusID		int
DECLARE @InsThisStatusID		int
DECLARE @DelThisStatusID		int
--If Inserting a Status that is CLIENTResponse
--Or REJECT02
--Means Client is updating docuements
--Do not insert a history item for this Assignment 
--Only do this if one assignment at a time is being updated
-- IF 	(SELECT COUNT(INS.[ID]) FROM INSERTED INS) = 1
-- BEGIN
-- 	IF (SELECT	[UserName] FROM USERS WHERE [UsersID] = (SELECT INS.[UpdateByUserID] FROM INSERTED INS)) = 'CFUSER'
-- 	BEGIN
-- 		--Set the Client Resposne StatusID
-- 		SET @CLIENTResponseStatusID = 	(
-- 						SELECT	[StatusID]
-- 						FROM	Status
-- 						WHERE	[Status]	= 'CLIENTResponse'
-- 						)
-- 		--Set the Client REJECT02 StatusID
-- 		SET @REJECT02StatusID = 	(
-- 						SELECT	[StatusID]
-- 						FROM	Status
-- 						WHERE	[Status]	= 'REJECT02'
-- 						)
-- 		--Select the Friggin Inserted
-- 		SELECT 
-- 			@InsThisStatusID	= INS.[StatusID]
-- 		FROM INSERTED INS
-- 		
-- 		--Select the Friggin Deleted
-- 		SELECT 
-- 			@DelThisStatusID 	= DEL.[StatusID]
-- 		FROM DELETED DEL
-- 		IF 	@InsThisStatusID <> @DelThisStatusID
-- 		BEGIN
-- 			SET @bAllowHistory = 1
-- 		END
-- 		ELSE
-- 		BEGIN
-- 			--Only Disallow History When the Deleted StatusID Matches the Insterted Statusid
-- 			--And the Inserted StatusID = CLIENTResponse or REJECT02
-- 			IF 	@InsThisStatusID = @CLIENTResponseStatusID
-- 				OR
-- 				@InsThisStatusID = @REJECT02StatusID
-- 			BEGIN
-- 				SET @bAllowHistory = 0
-- 			END
-- 			ELSE
-- 			BEGIN
-- 				SET @bAllowHistory = 1
-- 			END
-- 			
-- 		END
-- 	END
-- END
-- ELSE
-- BEGIN
-- 	SET @bAllowHistory =1
-- END
-- 	
-- IF @bAllowHistory = 1
-- BEGIN
-- 	INSERT INTO AssignmentsHistory	
-- 		SELECT del.*
-- 		FROM DELETED del
-- END
--Allow all updates to be inserted into the history table
-- Just have to live with the billion updates from Farmers for now.
INSERT INTO AssignmentsHistory	
SELECT del.*
FROM DELETED del

--5.19.2005 BGS  Farmers XML01 Need to check for certain updates to ContactDate and InspectedDate
-- When these changes occur need to update the FarmersXML01_ActivityUpdateQue with updated items
-- and then flag them to be sent off.
--As well, Changes to Property address items for Automobile (as of 5.19.2005 this is the only type to send updates this way)

--Only do this if one assignment at a time is being updated
IF (SELECT COUNT(INS.[ID]) FROM INSERTED INS) = 1
BEGIN
	--- Ignore updates from CFUSER !!!
	IF 	(
			(SELECT INS.[LRFormat] FROM INSERTED INS) = 'V2ECcarFarmers.clsLossXML01'
			AND
			(SELECT	[UserName] FROM USERS WHERE [UsersID] = (SELECT INS.[UpdateByUserID] FROM INSERTED INS)) <> 'CFUSER'
		)
		Or
		(
			(SELECT DEL.[LRFormat] FROM DELETED DEL) = 'V2ECcarFarmers.clsLossXML01'
			AND
			(SELECT	[UserName] FROM USERS WHERE [UsersID] = (SELECT INS.[UpdateByUserID] FROM INSERTED INS)) <> 'CFUSER'
		)
	BEGIN
		--Inserted
		DECLARE @INSContactDate as datetime
		DECLARE @INSInspectedDate as datetime
		DECLARE @INSPAStreet as varchar(150)
		DECLARE	@INSPACity as varchar(50)
		DECLARE @INSPAState as varchar(50)
		DECLARE @INSPAZIP as int 
		DECLARE @INSPAZIP4 as int
		DECLARE	@INSPAOtherPostCode as varchar(20)
		--DELETED
		DECLARE @DELContactDate as datetime
		DECLARE @DELInspectedDate as datetime
		DECLARE @DELPAStreet as varchar(150)
		DECLARE	@DELPACity as varchar(50)
		DECLARE @DELPAState as varchar(50)
		DECLARE @DELPAZIP as int 
		DECLARE @DELPAZIP4 as int
		DECLARE	@DELPAOtherPostCode as varchar(20)
		--ID vars
		DECLARE @AssignmentsID as int
		DECLARE @UID as int
		--- Property Que Vars
		DECLARE @LocationAddress as varchar(200)
		DECLARE @LocationCity  as varchar(50)
		DECLARE @LocationState  as varchar(50)
		DECLARE @LocationPostalCode  as varchar(30)
		DECLARE @GMTUpdated as datetime
		--Activity Que Vars
		DECLARE @Description as varchar(100)
		DECLARE @Comment as varchar(1500)
		DECLARE @CompletedDate as datetime
		DECLARE @GMTCreated as datetime
		DECLARE @ToDate as datetime


		--First Set all the Inserted and Deleted Vars so they
		--Can be compared!
		--ID vars
		SET @AssignmentsID = 	(SELECT INS.[AssignmentsID] FROM INSERTED INS)
		SET @UID =		(SELECT INS.[UpdateByUserID] FROM INSERTED INS)	
		--Inserted
		SET @INSContactDate = 	(SELECT INS.[ContactDate] FROM INSERTED INS)
		SET @INSInspectedDate = (SELECT INS.[InspectedDate] FROM INSERTED INS)
		SET @INSPAStreet = 	(SELECT INS.[PAStreet] FROM INSERTED INS)
		SET @INSPACity = 	(SELECT INS.[PACity] FROM INSERTED INS)
		SET @INSPAState = 	(SELECT INS.[PAState] FROM INSERTED INS)
		SET @INSPAZIP = 	(SELECT INS.[PAZIP] FROM INSERTED INS)
		SET @INSPAZIP4 = 	(SELECT INS.[PAZIP4] FROM INSERTED INS)
		SET @INSPAOtherPostCode = (SELECT INS.[PAOtherPostCode] FROM INSERTED INS)
		--Deleted
		SET @DELContactDate = 	(SELECT DEL.[ContactDate] FROM DELETED DEL)

		SET @DELInspectedDate = (SELECT DEL.[InspectedDate] FROM DELETED DEL)
		SET @DELPAStreet = 	(SELECT DEL.[PAStreet] FROM DELETED DEL)
		SET @DELPACity = 	(SELECT DEL.[PACity] FROM DELETED DEL)
		SET @DELPAState = 	(SELECT DEL.[PAState] FROM DELETED DEL)
		SET @DELPAZIP = 	(SELECT DEL.[PAZIP] FROM DELETED DEL)
		SET @DELPAZIP4 = 	(SELECT DEL.[PAZIP4] FROM DELETED DEL)
		SET @DELPAOtherPostCode = (SELECT DEL.[PAOtherPostCode] FROM DELETED DEL)
		
		---Update Activities Applies to both Auto and Property
		--Check for Contacted Date
		IF @INSContactDate Is Not Null
		BEGIN
			IF @DELContactDate Is Null
			BEGIN
				SET @DELContactDate = GetDate()
			END
		END
		ELSE
		BEGIN
			SET @INSContactDate = GetDate()
			SET @DELContactDate = @INSContactDate
		END
		IF @INSContactDate <> @DELContactDate
		BEGIN
			SET @ToDate = dbo.CleanFromOrToDate(@INSContactDate, @ToDate, 0)
			--Get some Comments if any
			SELECT  TOP 1 	@Comment = Convert(varchar(1500),[ActText]),
					@CompletedDate = [ActTime]
			FROM	RTActivityLog
			WHERE	[AssignmentsID] = @AssignmentsID
			AND 	[ActDate] >= @INSContactDate
			AND	[ActDate] <= @ToDate
			AND 	[IsDeleted] = 0 
			AND	[IsMgrEntry] = 0 
			ORDER BY [DateLastUpdated] 
			IF @Comment Is Null
			BEGIN
				SET @Comment = ''
			END
			
			--Convert the Date to GMT or UTC
			SET @GMTCreated = GETUTCDATE() 
			--Make Description
			IF @CompletedDate Is Null
			BEGIN
				SET @CompletedDate = @INSContactDate
			END		
			SET @Description = 'Contacted Insured '
			--Update the Activity Que table!!!
			UPDATE 	FarmersXML01_ActivityUpdateQue SET
				[Description] = @Description,
				[Comment] = @Comment,
				[CompletedDate] = @CompletedDate,
				[GMTCreated] = @GMTCreated,
				[SendToCRN] = 1,
				[DateLastUpdated] = GetDate(),
				[UpdateByUserID] = @UID				
			WHERE	[AssignmentsID] = @AssignmentsID
			AND 	[Type] = 'First Contact'
		END
		--Check for Inspected Date
		SET @CompletedDate = Null
		IF @INSInspectedDate Is Not Null
		BEGIN
			IF @DELInspectedDate Is Null
			BEGIN
				SET @DELInspectedDate = GetDate()
			END
		END
		ELSE
		BEGIN
			SET @INSInspectedDate = GetDate()
			SET @DELInspectedDate = @INSInspectedDate
		END
		IF @INSInspectedDate <> @DELInspectedDate
		BEGIN
			SET @ToDate = dbo.CleanFromOrToDate(@INSInspectedDate, @ToDate, 0)
			--Get some Comments if any
			SELECT  TOP 1 	@Comment = Convert(varchar(1500),[ActText]),
					@CompletedDate = [ActTime]
			FROM	RTActivityLog
			WHERE	[AssignmentsID] = @AssignmentsID
			AND 	[ActDate] >= @INSInspectedDate
			AND	[ActDate] <= @ToDate
			AND 	[IsDeleted] = 0 
			AND	[IsMgrEntry] = 0 
			ORDER BY [DateLastUpdated] 
			IF @Comment Is Null
			BEGIN
				SET @Comment = ''
			END
			
			--Convert the Date to GMT or UTC
			SET @GMTCreated = GETUTCDATE() 
			--Make Description
			IF @CompletedDate Is Null
			BEGIN
				SET @CompletedDate = @INSInspectedDate
			END		
			SET @Description = 'Inspection Completed ' 
			--Update the Activity Que table!!!
			UPDATE 	FarmersXML01_ActivityUpdateQue SET
				[Description] = @Description,
				[Comment] = @Comment,
				[CompletedDate] = @CompletedDate,
				[GMTCreated] = @GMTCreated,
				[SendToCRN] = 1,
				[DateLastUpdated] = GetDate(),
				[UpdateByUserID] = @UID				
			WHERE	[AssignmentsID] = @AssignmentsID
			AND 	[Type] = 'Field Inspection'
		END
		--For Property updates (onlt applies to Automobile at this time)
		IF (SELECT INS.[AssignmentTypeID] FROM INSERTED INS) = 	(
									SELECT 	[AssignmentTypeID]
									FROM 	AssignmentType
									WHERE	[Type] = 'Auto'
									)
		BEGIN
			--If there are any differences need to update entire Property Location
			IF 	@INSPAStreet <> @DELPAStreet 
				OR @INSPACity <> @DELPACity 
				Or @INSPAState <> @DELPAState 
				Or @INSPAZIP <> @DELPAZIP 
				Or @INSPAZIP4 <> @DELPAZIP4
				Or @INSPAOtherPostCode <> @DELPAOtherPostCode
			BEGIN
				--- Build the Property Field
				SET @LocationAddress =  	Convert(varchar(200),@INSPAStreet)
				SET @LocationCity =  		Convert(varchar(50),@INSPACity)
				SET @LocationState =  		Convert(varchar(50),@INSPAState)
				IF @INSPAZIP4 > 0
				BEGIN
					SET @LocationPostalCode =  	Convert(varchar(10),@INSPAZIP) + '-' + Convert(varchar(4),@INSPAZIP4)
				END
				ELSE
				BEGIN
					IF @INSPAZIP > 0
					BEGIN
						SET @LocationPostalCode =  	Convert(varchar(10),@INSPAZIP)
					END
					ELSE
					BEGIN
						SET @LocationPostalCode = @INSPAOtherPostCode
					END
				END
				
				SET @GMTUpdated = 		GetUTCDate()
				UPDATE 	FarmersXML01_LocationUpdateQue SET
					[LocationAddress] = @LocationAddress,
					[LocationCity] = @LocationCity,
					[LocationState] = @LocationState,
					[LocationPostalCode] = @LocationPostalCode,
					[GMTUpdated] = @GMTUpdated,
					[SendToCRN] = 1,
					[DateLastUpdated] = GetDate(),
					[UpdateByUserID] = @UID				
				WHERE	[AssignmentsID] = @AssignmentsID
			END
		END
	END
END

--Check for Adjuster Requesting Approval for Claim Package.
--IF the adjuster is Flagged as [SingleFileSendAuthority] under the 
--AdjusterUsersSoftware Table.  Then need to Change the status directly to PendingDelivery...
--on the Assignments Table to Be Pending Delivery update the Package Table
--Only do this if one assignment at a time is being updated
DECLARE @INSStatusID 			int
DECLARE @DELStatusID 			int
DECLARE @StatusName 			varchar(20)
DECLARE @ADJUSERSID 			int
DECLARE @CFUSERSID			int 
DECLARE	@SingleFileSendAuthority	bit
DECLARE	@UseThisStatusID		int
DECLARE @PendingDeliveryStatusID	int
DECLARE @InsAssignmentsID		int 
IF (SELECT COUNT(INS.[ID]) FROM INSERTED INS) = 1
BEGIN
	--First Check to see if the Inserted Status Is <> to the current status...
	SET @INSStatusID = (SELECT INS.[StatusID] FROM INSERTED INS)
	SET @DELStatusID = (SELECT DEL.[StatusID] FROM DELETED DEL)
	IF @INSStatusID <> @DELStatusID
	BEGIN
		--Check the Inserted StatusID to see if it is a Request for Approval
		--If it is then Check for the Adjuster being flagged for [SingleFileSendAuthority]
		SET @StatusName = (SELECT [Status] FROM Status WHERE [StatusID] = @INSStatusID ) 
		IF @StatusName Is Not Null
		BEGIN
			IF @StatusName = 'APPROVALRequest'
			BEGIN
				--Only do this when the USER is updating records!  
				--is the same user that is Assigned to the record!
				SET @CFUSERSID = (SELECT INS.[UpdateByUserID] FROM INSERTED INS)
				SET @ADJUSERSID = (SELECT [UsersID] FROM ClientCoAdjusterSpec WHERE [ClientCoAdjusterSpecID] = (SELECT INS.[AdjusterSpecID] FROM INSERTED INS) ) 
				IF @CFUSERSID Is Not Null And @ADJUSERSID Is Not Null
				BEGIN
					IF @ADJUSERSID = @CFUSERSID
					BEGIN
						SET @SingleFileSendAuthority = (SELECT [SingleFileSendAuthority] FROM AdjusterUsersSoftware WHERE [UsersID] = @ADJUSERSID )
						IF @SingleFileSendAuthority Is Not Null
						BEGIN
							IF @SingleFileSendAuthority = 1
							BEGIN
								SET @InsAssignmentsID = (SELECT INS.[AssignmentsID] FROM INSERTED INS)
								SET @PendingDeliveryStatusID = (SELECT [StatusID] FROM Status WHERE [Status] = 'PENDINGDelivery' )
								SET @UseThisStatusID = @PendingDeliveryStatusID
								-- Also Update the Package Table with this Good News
								EXEC 	z_spuUpdatePackageStatus
									@ADJUSERSID				,	--int,
									@InsAssignmentsID			,	--int,
									1					,	--int,
									1					,	--bit=null,
									@UseThisStatusID			,	--int=null,
									'[SingleFileSendAuthority] is ON!'	,	--varchar(1000)=null,
									null					,	--varchar(50)=null,
									null					,	--DateTime=null,
									1						--bit=null
							END
						END
					END
				END
			END
		END
	END
END

-- Now that the History table was updated first...
--Allow the original update to process...
IF (SELECT COUNT(INS.[ID]) FROM INSERTED INS) > 1
BEGIN
	Update Assignments SET
		[ID]= 			INS.ID,
		[AssignmentTypeID]=	INS.AssignmentTypeID,
		[ClientCompanyCatSpecID]=INS.ClientCompanyCatSpecID,
		[AdjusterSpecID]= 	INS.AdjusterSpecID,
		[AdjusterSpecIDDisplay]= INS.AdjusterSpecIDDisplay,
		[SPVersion]=		INS.SPVersion,
		[IBNUM]= 		INS.IBNUM,
		[CLIENTNUM]= 		INS.CLIENTNUM,
		[PolicyNo]= 		INS.PolicyNo,
		[PolicyDescription]= 	INS.PolicyDescription,
		[Insured]= 		INS.Insured,
		[MailingAddress]= 	INS.MailingAddress,
		[MAStreet]= 		INS.MAStreet,
		[MACity]=		INS.MACity,
		[MAState]= 		INS.MAState,
		[MAZIP]= 		INS.MAZIP,
		[MAZIP4]= 		INS.MAZIP4,
		[MAOtherPostCode]=	INS.[MAOtherPostCode],
		[HomePhone]= 		INS.HomePhone,
		[BusinessPhone]= 	INS.BusinessPhone,
		[PropertyAddress]= 	INS.PropertyAddress,
		[PAStreet]= 		INS.PAStreet,
		[PACity]= 		INS.PACity,
		[PAState]= 		INS.PAState,
		[PAZIP]= 		INS.PAZIP,
		[PAZIP4]= 		INS.PAZIP4,
		[PAOtherPostCode]=	INS.[PAOtherPostCode],	
		[MortgageeName]= 	INS.MortgageeName,
		[AgentNo]= 		INS.AgentNo,
		[ReportedBy]= 		INS.ReportedBy,
		[ReportedByPhone]= 	INS.ReportedByPhone,
		[Deductible]= 		INS.Deductible,
		[AppDedClassTypeIDOrder]= INS.AppDedClassTypeIDOrder,
		[LRFormat]= 		INS.LRFormat,
		[LossReport]= 		INS.LossReport,
		[LRPrintedDate]=	INS.LRPrintedDate,
		[DownLoadLossReport]=INS.DownLoadLossReport,
		[UpLoadLossReport]=	INS.UpLoadLossReport,
		[StatusID]= 		(CASE WHEN @UseThisStatusID Is Not Null THEN @UseThisStatusID ELSE INS.StatusID END),
		[TypeOfLossID]=		INS.[TypeOfLossID],
		[XactTypeOfLoss]=	INS.[XactTypeOfLoss],
		[SentToXact]=		INS.[SentToXact],
		[LossDate]= 		INS.LossDate,
		[AssignedDate]= 	INS.AssignedDate,
		[ReceivedDate]=		INS.[ReceivedDate],
		[ContactDate]= 		INS.ContactDate,
		[InspectedDate]=	INS.[InspectedDate],
		[CloseDate]= 		INS.CloseDate,
		[Reassigned]= 		INS.Reassigned,
		[DateReassigned]= 	INS.DateReassigned,
		[RAAdjusterSpecID]= 	INS.RAAdjusterSpecID,
		[IsLocked]=		INS.IsLocked,
		[IsDeleted]=		INS.IsDeleted,
		[DownloadMe]=		(CASE WHEN @UseThisStatusID Is Not Null THEN 1 ELSE INS.DownloadME END),
		[UpLoadMe]=		INS.UpLoadME,
		[DownloadAll]=		INS.DownloadAll,
		[UpLoadAll]=		INS.UpLoadAll,
		[AdminComments]=	INS.AdminComments,
		[MiscDelimSettings]=	INS.MiscDelimSettings,
		[DateLastUpdated]= 	INS.DateLastUpdated,
		[UpdateByUserID]= 	INS.UpdateByUserID 
	FROM Assignments A INNER JOIN INSERTED INS ON A.AssignmentsID = INS.AssignmentsID
END
ELSE
BEGIN
	IF 	(
			(SELECT INS.[LRFormat] FROM INSERTED INS) <> 'V2ECcarFarmers.clsLossXML01'
			And
			(SELECT DEL.[LRFormat] FROM DELETED DEL) <> 'V2ECcarFarmers.clsLossXML01'
		)
	BEGIN
		Update Assignments SET
			[ID]= 			INS.ID,
			[AssignmentTypeID]=	INS.AssignmentTypeID,
			[ClientCompanyCatSpecID]=INS.ClientCompanyCatSpecID,
			[AdjusterSpecID]= 	INS.AdjusterSpecID,
			[AdjusterSpecIDDisplay]= INS.AdjusterSpecIDDisplay,
			[SPVersion]=		INS.SPVersion,
			[IBNUM]= 		INS.IBNUM,
			[CLIENTNUM]= 		INS.CLIENTNUM,
			[PolicyNo]= 		INS.PolicyNo,
			[PolicyDescription]= 	INS.PolicyDescription,
			[Insured]= 		INS.Insured,
			[MailingAddress]= 	INS.MailingAddress,
			[MAStreet]= 		INS.MAStreet,
			[MACity]=		INS.MACity,
			[MAState]= 		INS.MAState,
			[MAZIP]= 		INS.MAZIP,
			[MAZIP4]= 		INS.MAZIP4,
			[MAOtherPostCode]=	INS.[MAOtherPostCode],
			[HomePhone]= 		INS.HomePhone,
			[BusinessPhone]= 	INS.BusinessPhone,
			[PropertyAddress]= 	INS.PropertyAddress,
			[PAStreet]= 		INS.PAStreet,
			[PACity]= 		INS.PACity,
			[PAState]= 		INS.PAState,
			[PAZIP]= 		INS.PAZIP,
			[PAZIP4]= 		INS.PAZIP4,
			[PAOtherPostCode]=	INS.[PAOtherPostCode],	
			[MortgageeName]= 	INS.MortgageeName,
			[AgentNo]= 		INS.AgentNo,
			[ReportedBy]= 		INS.ReportedBy,
			[ReportedByPhone]= 	INS.ReportedByPhone,
			[Deductible]= 		INS.Deductible,
			[AppDedClassTypeIDOrder]= INS.AppDedClassTypeIDOrder,
			[LRFormat]= 		INS.LRFormat,
			[LossReport]= 		INS.LossReport,
			[LRPrintedDate]=	INS.LRPrintedDate,
			[DownLoadLossReport]=INS.DownLoadLossReport,
			[UpLoadLossReport]=	INS.UpLoadLossReport,
			[StatusID]= 		(CASE WHEN @UseThisStatusID Is Not Null THEN @UseThisStatusID ELSE INS.StatusID END),
			[TypeOfLossID]=		INS.[TypeOfLossID],
			[XactTypeOfLoss]=	INS.[XactTypeOfLoss],
			[SentToXact]=		INS.[SentToXact],
			[LossDate]= 		INS.LossDate,
			[AssignedDate]= 	INS.AssignedDate,
			[ReceivedDate]=		INS.[ReceivedDate],
			[ContactDate]= 		INS.ContactDate,
			[InspectedDate]=	INS.[InspectedDate],
			[CloseDate]= 		INS.CloseDate,
			[Reassigned]= 		INS.Reassigned,
			[DateReassigned]= 	INS.DateReassigned,
			[RAAdjusterSpecID]= 	INS.RAAdjusterSpecID,
			[IsLocked]=		INS.IsLocked,
			[IsDeleted]=		INS.IsDeleted,
			[DownloadMe]=		(CASE WHEN @UseThisStatusID Is Not Null THEN 1 ELSE INS.DownloadME END),
			[UpLoadMe]=		INS.UpLoadME,
			[DownloadAll]=		INS.DownloadAll,
			[UpLoadAll]=		INS.UpLoadAll,
			[AdminComments]=	INS.AdminComments,
			[MiscDelimSettings]=	INS.MiscDelimSettings,
			[DateLastUpdated]= 	INS.DateLastUpdated,
			[UpdateByUserID]= 	INS.UpdateByUserID 
		FROM Assignments A INNER JOIN INSERTED INS ON A.AssignmentsID = INS.AssignmentsID
	END
	ELSE
	BEGIN
		-- V2ECcarFarmers.clsLossXML01
		-- If this is Update by CFUSER need to allow updates to Loss Report Field
		IF 	(
				(
					(SELECT INS.[LRFormat] FROM INSERTED INS) = 'V2ECcarFarmers.clsLossXML01'
					AND
					(SELECT	[UserName] FROM USERS WHERE [UsersID] = (SELECT INS.[UpdateByUserID] FROM INSERTED INS)) = 'CFUSER'
				)
				Or
				(
					(SELECT DEL.[LRFormat] FROM DELETED DEL) = 'V2ECcarFarmers.clsLossXML01'
					AND
					(SELECT	[UserName] FROM USERS WHERE [UsersID] = (SELECT INS.[UpdateByUserID] FROM INSERTED INS)) = 'CFUSER'
				)
			)
		BEGIN
			Update Assignments SET
				[ID]= 			INS.ID,
				[AssignmentTypeID]=	INS.AssignmentTypeID,
				[ClientCompanyCatSpecID]=INS.ClientCompanyCatSpecID,
				[AdjusterSpecID]= 	INS.AdjusterSpecID,
				[AdjusterSpecIDDisplay]= INS.AdjusterSpecIDDisplay,
				[SPVersion]=		INS.SPVersion,
				[IBNUM]= 		INS.IBNUM,
				[CLIENTNUM]= 		INS.CLIENTNUM,
				[PolicyNo]= 		INS.PolicyNo,
				[PolicyDescription]= 	INS.PolicyDescription,
				[Insured]= 		INS.Insured,
				[MailingAddress]= 	INS.MailingAddress,
				[MAStreet]= 		INS.MAStreet,
				[MACity]=		INS.MACity,
				[MAState]= 		INS.MAState,
				[MAZIP]= 		INS.MAZIP,

				[MAZIP4]= 		INS.MAZIP4,
				[MAOtherPostCode]=	INS.[MAOtherPostCode],
				[HomePhone]= 		INS.HomePhone,
				[BusinessPhone]= 	INS.BusinessPhone,
				[PropertyAddress]= 	INS.PropertyAddress,
				[PAStreet]= 		INS.PAStreet,
				[PACity]= 		INS.PACity,
				[PAState]= 		INS.PAState,
				[PAZIP]= 		INS.PAZIP,
				[PAZIP4]= 		INS.PAZIP4,
				[PAOtherPostCode]=	INS.[PAOtherPostCode],	
				[MortgageeName]= 	INS.MortgageeName,
				[AgentNo]= 		INS.AgentNo,
				[ReportedBy]= 		INS.ReportedBy,
				[ReportedByPhone]= 	INS.ReportedByPhone,
				[Deductible]= 		INS.Deductible,
				[AppDedClassTypeIDOrder]= INS.AppDedClassTypeIDOrder,
				[LRFormat]= 		INS.LRFormat,
				[LossReport]= 		INS.LossReport,
				[LRPrintedDate]=	INS.LRPrintedDate,
				[DownLoadLossReport]=INS.DownLoadLossReport,
				[UpLoadLossReport]=	INS.UpLoadLossReport,
				[StatusID]= 		(CASE WHEN @UseThisStatusID Is Not Null THEN @UseThisStatusID ELSE INS.StatusID END),
				[TypeOfLossID]=		INS.[TypeOfLossID],
				[XactTypeOfLoss]=	INS.[XactTypeOfLoss],
				[SentToXact]=		INS.[SentToXact],
				[LossDate]= 		INS.LossDate,
				[AssignedDate]= 	INS.AssignedDate,
				[ReceivedDate]=		INS.[ReceivedDate],
				[ContactDate]= 		INS.ContactDate,
				[InspectedDate]=	INS.[InspectedDate],
				[CloseDate]= 		INS.CloseDate,
				[Reassigned]= 		INS.Reassigned,
				[DateReassigned]= 	INS.DateReassigned,
				[RAAdjusterSpecID]= 	INS.RAAdjusterSpecID,
				[IsLocked]=		INS.IsLocked,
				[IsDeleted]=		INS.IsDeleted,
				[DownloadMe]=		(CASE WHEN @UseThisStatusID Is Not Null THEN 1 ELSE INS.DownloadME END),
				[UpLoadMe]=		INS.UpLoadME,
				[DownloadAll]=		INS.DownloadAll,
				[UpLoadAll]=		INS.UpLoadAll,
				[AdminComments]=	INS.AdminComments,
				[MiscDelimSettings]=	INS.MiscDelimSettings,
				[DateLastUpdated]= 	INS.DateLastUpdated,
				[UpdateByUserID]= 	INS.UpdateByUserID 
			FROM Assignments A INNER JOIN INSERTED INS ON A.AssignmentsID = INS.AssignmentsID
			RETURN --Exit this SP HERE 
		END
		-- If this is Update other than by CFUSER need to NOT ! allow  updates to Loss Report Field
		IF 	(
				(
					(SELECT INS.[LRFormat] FROM INSERTED INS) = 'V2ECcarFarmers.clsLossXML01'
					AND
					(SELECT	[UserName] FROM USERS WHERE [UsersID] = (SELECT INS.[UpdateByUserID] FROM INSERTED INS)) <> 'CFUSER'
				)
				Or
				(
					(SELECT DEL.[LRFormat] FROM DELETED DEL) = 'V2ECcarFarmers.clsLossXML01'
					AND
					(SELECT	[UserName] FROM USERS WHERE [UsersID] = (SELECT INS.[UpdateByUserID] FROM INSERTED INS)) <> 'CFUSER'
				)
			)
		BEGIN
			Update Assignments SET
				[ID]= 			INS.ID,
				[AssignmentTypeID]=	INS.AssignmentTypeID,
				[ClientCompanyCatSpecID]=INS.ClientCompanyCatSpecID,
				[AdjusterSpecID]= 	INS.AdjusterSpecID,
				[AdjusterSpecIDDisplay]= INS.AdjusterSpecIDDisplay,
				[SPVersion]=		INS.SPVersion,
				[IBNUM]= 		INS.IBNUM,
				[CLIENTNUM]= 		INS.CLIENTNUM,
				[PolicyNo]= 		INS.PolicyNo,
				[PolicyDescription]= 	INS.PolicyDescription,
				[Insured]= 		INS.Insured,
				[MailingAddress]= 	INS.MailingAddress,
				[MAStreet]= 		INS.MAStreet,
				[MACity]=		INS.MACity,
				[MAState]= 		INS.MAState,
				[MAZIP]= 		INS.MAZIP,
				[MAZIP4]= 		INS.MAZIP4,
				[MAOtherPostCode]=	INS.[MAOtherPostCode],
				[HomePhone]= 		INS.HomePhone,
				[BusinessPhone]= 	INS.BusinessPhone,
				[PropertyAddress]= 	INS.PropertyAddress,
				[PAStreet]= 		INS.PAStreet,
				[PACity]= 		INS.PACity,
				[PAState]= 		INS.PAState,
				[PAZIP]= 		INS.PAZIP,
				[PAZIP4]= 		INS.PAZIP4,
				[PAOtherPostCode]=	INS.[PAOtherPostCode],	
				[MortgageeName]= 	INS.MortgageeName,
				[AgentNo]= 		INS.AgentNo,
				[ReportedBy]= 		INS.ReportedBy,
				[ReportedByPhone]= 	INS.ReportedByPhone,
				[Deductible]= 		INS.Deductible,
				[AppDedClassTypeIDOrder]= INS.AppDedClassTypeIDOrder,
	-- 			[LRFormat]= 		DEL.LRFormat, --If this is Update other than by CFUSER need to NOT ! allow  updates to Loss Report Field
	-- 			[LossReport]= 		DEL.LossReport, --If this is Update other than by CFUSER need to NOT ! allow  updates to Loss Report Field
				[LRPrintedDate]=	INS.LRPrintedDate,
				[DownLoadLossReport]=INS.DownLoadLossReport,
				[UpLoadLossReport]=	INS.UpLoadLossReport,
				[StatusID]= 		(CASE WHEN @UseThisStatusID Is Not Null THEN @UseThisStatusID ELSE INS.StatusID END),
				[TypeOfLossID]=		INS.[TypeOfLossID],
				[XactTypeOfLoss]=	INS.[XactTypeOfLoss],
				[SentToXact]=		INS.[SentToXact],
				[LossDate]= 		INS.LossDate,
				[AssignedDate]= 	INS.AssignedDate,
				[ReceivedDate]=		INS.[ReceivedDate],
				[ContactDate]= 		INS.ContactDate,
				[InspectedDate]=	INS.[InspectedDate],
				[CloseDate]= 		INS.CloseDate,
				[Reassigned]= 		INS.Reassigned,
				[DateReassigned]= 	INS.DateReassigned,
				[RAAdjusterSpecID]= 	INS.RAAdjusterSpecID,
				[IsLocked]=		INS.IsLocked,
				[IsDeleted]=		INS.IsDeleted,
				[DownloadMe]=		(CASE WHEN @UseThisStatusID Is Not Null THEN 1 ELSE INS.DownloadME END),
				[UpLoadMe]=		INS.UpLoadME,
				[DownloadAll]=		INS.DownloadAll,
				[UpLoadAll]=		INS.UpLoadAll,
				[AdminComments]=	INS.AdminComments,
				[MiscDelimSettings]=	INS.MiscDelimSettings,
				[DateLastUpdated]= 	INS.DateLastUpdated,
				[UpdateByUserID]= 	INS.UpdateByUserID 
			FROM Assignments A 	INNER JOIN INSERTED INS ON A.AssignmentsID = INS.AssignmentsID
			RETURN --Exit this SP HERE 
		END
	END
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

