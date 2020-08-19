SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spuInsertAssignment]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spuInsertAssignment]
GO

/******************************************************************
*
*    	Description:	Insert Assignments 
*	
*	Author: 	Brad Skidmore
*	Date: 		1/7/2004
*	Date:		4.27.2005 BGS Farmers XML Update Modification
*	Date"		10.12.2005 BGS When Farmers UNDELETES a UNIT, need to update the Assignment close date to null
******************************************************************/

CREATE PROCEDURE [dbo].[z_spuInsertAssignment]
	@ClientCompanyCode		varchar(20), --Code From Company table for the Client Company
	@ID				int=null,
	@AssignmentTypeID		int,
	@ClientCompanyCatSpecID		int,
	@AdjusterSpecID 		int,
	@AdjusterSpecIDDisplay 		int,
	@SPVersion			int,
	@IBNUM 				varchar(20)=null,
	@CLIENTNUM 			varchar(20),
	@PolicyNo 			varchar(50)=null,
	@PolicyDescription		varchar(100)=null,
	@Insured 			varchar(100)=null,
	@MailingAddress 		varchar(255)=null,
	@MAStreet 			varchar(150)=null,
	@MACity 			varchar(50)=null,
	@MAState 			varchar(50)=null,
	@MAZIP 				int=null,
	@MAZIP4 			int=null,
	@MAOtherPostCode 		varchar(20)=null,
	@HomePhone 			varchar(50)=null,
	@BusinessPhone			varchar(50)=null,
	@PropertyAddress 		varchar(255)=null,
	@PAStreet 			varchar(150)=null,
	@PACity 			varchar(50)=null,
	@PAState 			varchar(50)=null,
	@PAZIP 				int=null,
	@PAZIP4 			int=null,
	@PAOtherPostCode 		varchar(20)=null,
	@MortgageeName 			varchar(100)=null,
	@AgentNo 			varchar(50)=null,
	@ReportedBy 			varchar(100)=null,
	@ReportedByPhone 		varchar(50)=null,
	@Deductible 			money=null,
	@AppDedClassTypeIDOrder 	varchar(255)=null,
	@LRFormat 			varchar(50)=null,
	@LossReportPrnData		ntext=null,
	@DownLoadLossReport		bit=1, --Inserting an Assignment or Updating The LossReport Need to Flag the Download for the Loss Report
	@UpLoadLossReport		bit=0,
	@StatusID 			int,
	@TypeOfLossID 			int=null,
	@XactTypeOfLoss 		varchar(50)=null,
	@SentToXact 			bit,
	@LossDate 			datetime=null,
	@AssignedDate 			datetime=null,
	@ReceivedDate 			datetime=null,
	@ContactDate 			datetime=null,
	@InspectedDate 			datetime=null,
	@CloseDate 			datetime=null,
	@Reassigned 			bit=0,
	@DateReassigned 		datetime=null,
	@RAAdjusterSpecID 		int=null,
	@IsLocked 			bit=0,
	@IsDeleted 			bit=0,
	@DownLoadMe 			bit=1, --Inserting an Assignment need to Flag the Entire Record for Downloading
	@UpLoadMe			bit=0,
	@AdminComments 			varchar(1000)=null,
	@MiscDelimSettings		text,--MiscDelimString TEXT TYPE!(This String will hold delim data of misc items applicable for each assignment)
	@PolicyLimitsList		text=Null,--Policy Limits Delimited String to be parsed by SplitPolicyLimits UDF
	@DateLastUpdated 		datetime,
	@UpdateByUserID 		int,
--****************************************************************
--*		BEGIN 	FarmersXML01_LocationUpdateQue
--*		5.10.2005 BGS Need to Insert an
--*
--****************************************************************
	--Output the new AssignmentsID
	@AssignmentsID 		int=Null OUTPUT
--****************************************************************
--*		END 	FarmersXML01_LocationUpdateQue
--****************************************************************	
AS

DECLARE @ptrLossReport	binary(16) --Pointer for ntext field update
--Farmers Adding Units to exisiting Assignment
DECLARE @FarmersXMLAddUnit		varchar(200)		
DECLARE @FarmersUnitRowID		varchar(20)
DECLARE	@PolicyLimitsID			int
DECLARE @FarmersUnitIsDeleted		bit
DECLARE @strDateTime			varchar(50)
DECLARE @sTemp				varchar(500)

SET @strDateTime = Cast(@DateLastUpdated As Varchar(50))

--First check to see if the CLIENTNUM IS already in Assignments
-- Assigned to the same adjuster.  If it is then Just update the Loss Report with this new one.
-- When the Loss Report is updated, the trigger on Assignemts will update it's history table.
SET @AssignmentsID = 	(	SELECT 	TOP 1 AssignmentsID 
				FROM 	Assignments 
				WHERE  	CLIENTNUM = @CLIENTNUM
				AND 	AdjusterSpecID IN	(
									SELECT 	ClientCoAdjusterSpecID
									FROM 	ClientCoAdjusterSpec
									WHERE	USERSID	IN	(
													SELECT 	USERSID
													FROM	ClientCoAdjusterSpec
													WHERE	ClientCoAdjusterSpecID = @AdjusterSpecID
												)
								) 
			)
---If the Property ZIPCODE is 0 OR Null, then need to Use Mailing Address
---As The Property Address
IF (@PAZIP Is Null or @PAZIP = 0) AND (@MAZIP Is Not Null) AND (@MAZIP <> 0) AND (@LRFormat <> 'V2ECCarFarmers.clsLossXML01')
BEGIN
	--Update the Admin Comments using MAiling Address for Property
	--Because the Property Address is Invalid	
	IF @AdminComments Is Not Null
	BEGIN
		SET @AdminComments = @AdminComments + 'MAILING ADDRESS USED AS PROPERTY ADDRESS'
	END
	ELSE
	BEGIN
		SET @AdminComments = 'MAILING ADDRESS USED AS PROPERTY ADDRESS'
	END
	--Leave the @MailingAddress (Orginal Entire Mailing address Data)
	--Just update the Mailing Address parts.
	SET @PAStreet = @MAStreet
	SET @PACity = @MACity
	SET @PAState = @MAState
	SET @PAZIP = @MAZIP 
	SET @PAZIP4 = @MAZIP4
END

IF @AssignmentsID IS NULL
BEGIN
	INSERT INTO Assignments 	(
					[ID],
					[AssignmentTypeID],
					[ClientCompanyCatSpecID],
					[AdjusterSpecID],
					[AdjusterSpecIDDisplay],
					[SPVersion],
					[IBNUM],
					[CLIENTNUM],
					[PolicyNo],
					[PolicyDescription],
					[Insured],
					[MailingAddress],
					[MAStreet],
					[MACity],
					[MAState],
					[MAZIP],
					[MAZIP4],
					[MAOtherPostCode],
					[HomePhone],
					[BusinessPhone],
					[PropertyAddress],
					[PAStreet],
					[PACity],
					[PAState],
					[PAZIP],
					[PAZIP4],
					[PAOtherPostCode],
					[MortgageeName],
					[AgentNo],
					[ReportedBy],
					[ReportedByPhone],
					[Deductible],
					[AppDedClassTypeIDOrder],
					[LRFormat],
					[LossReport],
					[DownLoadLossReport],
					[UpLoadLossReport],
					[StatusID],
					[TypeOfLossID],
					[XactTypeOfLoss],
					[SentToXact],
					[LossDate],
					[AssignedDate],
					[ReceivedDate],
					[ContactDate],
					[InspectedDate],
					[CloseDate],
					[Reassigned],
					[DateReassigned],
					[RAAdjusterSpecID],
					[IsLocked],
					[IsDeleted],
					[DownLoadMe],
					[UpLoadMe],
					[AdminComments],
					[MiscDelimSettings],
					[DateLastUpdated],
					[UpdateByUserID]
					)
		SELECT			@ID AS [ID],
					@AssignmentTypeID AS [AssignmentTypeID],
					@ClientCompanyCatSpecID As [ClientCompanyCatSpecID],
					@AdjusterSpecID As [AdjusterSpecID],
					@AdjusterSpecIDDisplay As [AdjusterSpecIDDisplay],
					@SPVersion AS [SPVersion],
					@IBNUM As [IBNUM],
					@CLIENTNUM As [CLIENTNUM],
					@PolicyNo As [PolicyNo],
					@PolicyDescription As [PolicyDescription],
					@Insured As [Insured],
					@MailingAddress As [MailingAddress],
					@MAStreet As [MAStreet],
					@MACity As [MACity],
					@MAState As [MAState],
					@MAZIP As [MAZIP],
					@MAZIP4 As [MAZIP4],
					@MAOtherPostCode As [MAOtherPostCode],
					@HomePhone As [HomePhone],
					@BusinessPhone As [BusinessPhone],
					@PropertyAddress As [PropertyAddress],
					@PAStreet As [PAStreet],
					@PACity As [PACity],
					@PAState As [PAState],
					@PAZIP As [PAZIP],
					@PAZIP4 As [PAZIP4],
					@PAOtherPostCode As [PAOtherPostCode],
					@MortgageeName As [MortgageeName],
					@AgentNo As [AgentNo],
					@ReportedBy As [ReportedBy],
					@ReportedByPhone As [ReportedByPhone],
					@Deductible As [Deductible],
					@AppDedClassTypeIDOrder As [AppDedClassTypeIDOrder],
					@LRFormat As [LRFormat],
					@LossReportPrnData AS [LossReport],
					@DownLoadLossReport As [@DownLoadLossReport],
					@UpLoadLossReport As [@UpLoadLossReport],
					@StatusID As [StatusID],
					@TypeOfLossID As [TypeOfLossID],
					@XactTypeOfLoss As [XactTypeOfLoss],
					@SentToXact As [SentToXact],
					@LossDate As [LossDate],
					@AssignedDate As [AssignedDate],
					@ReceivedDate As [ReceivedDate],
					@ContactDate As [ContactDate],
					@InspectedDate As [InspectedDate],
					@CloseDate As [CloseDate],
					@Reassigned As [Reassigned],
					@DateReassigned As [DateReassigned],
					@RAAdjusterSpecID As [RAAdjusterSpecID],
					@IsLocked As [IsLocked],
					@IsDeleted As [IsDeleted],
					@DownLoadMe As [DownLoadMe],
					@UpLoadMe As [UpLoadMe],
					@AdminComments As [AdminComments],
					@MiscDelimSettings As [MiscDelimSettings],
					@DateLastUpdated As [DateLastUpdated],
					@UpdateByUserID As [UpdateByUserID]
	
	SET @AssignmentsID = (Select IDENT_CURRENT('Assignments') as AssignmentsID)
	
	
	-- After this Assignments has been inserted need to Update it with 
	--AN IBNUMBER (USE the Client Company code and Assignment ID)
	--And
	--ID (ID is what Easy Claim Adjuster Database uses for the Primary Key)
	--When an adjuster creates an Assignment manually, ID is the same Primary Key
	--It is very important to set the Seed of for AssignmentsID On SQL Server assignments Table to start ar 20001.
	--This is because the ID on Easy Claim Assignments table needs room to add manual Claims per CAT
	--20000 should be a generous amount of IDs per Cat.
	--Also note that When Reassigning a Claim from one adjuster top the next...
	--The original assignment Record needs to be copied and inserted as a New Assignment Record
	--Creating a NEW AssignmentID and ID to be downloaded by the Adjuster it was reassigned to.
	UPDATE Assignments SET 
				[ID]	=	(	CASE 	WHEN [ID] Is Null
							THEN @AssignmentsID
							ELSE [ID]
							END
						),
				[IBNUM] = 	(	CASE 	WHEN [IBNUM] Is Null
							THEN @ClientCompanyCode + Convert(varchar(20),@AssignmentsID)
							ELSE [IBNUM]
							END
						)
	
	WHERE AssignmentsID = @AssignmentsID

	--After Updating Assignment info need to populate Policy Limits if any
	IF @PolicyLimitsList Is Not Null
	BEGIN
		EXEC z_spuUpdatePolicyLimitsFromSplit
			@PolicyLimitsList	,--Policy Limits Delimited String to be parsed by SplitPolicyLimits UDF
			@AssignmentsID		,
			@DateLastUpdated 	,
			@UpdateByUserID 	
	END
END
ELSE
BEGIN
	--4.27.2005 BGS Check for Farmers XML Updates
	IF @LRFormat = 'V2ECCarFarmers.clsLossXML01'
	BEGIN 
		--First See if this Unit being added has not been previously added. 
		-- If This unit already exisits for this Assigned Loss (same adjuster...)
		-- Need to give an Error stating the case 
		SELECT	@FarmersUnitRowID = PL.[FarmersUnitRowID]
		FROM 	dbo.SplitPolicyLimits(@PolicyLimitsList, default, default) As PL
		
		SELECT 	TOP 1 
			@PolicyLimitsID = [PolicyLimitsID],
			@FarmersUnitIsDeleted = [IsDeleted]
		FROM	PolicyLimits 
		WHERE	[AssignmentsID] = @AssignmentsID
		AND	[FarmersUnitRowID] = @FarmersUnitRowID

		IF @PolicyLimitsID Is Not Null
		BEGIN
			-- If the Unit Already Exists for this Assignment....
			-- And the already Exisiting Unit Is Deleted...
			-- That Means this Add Unit is Really an Undelete 
			IF @FarmersUnitIsDeleted = 1
			BEGIN
				SET @sTemp = 'UNDELETE UNIT CLIENT REOPEN ' + @strDateTime
				---No Error will be raised
				UPDATE PolicyLimits SET	
						[AdminComments] 	= @sTemp,
						[IsDeleted]		= 0,
						[DownLoadMe]		= 1,
						[DateLastUpdated]	= @DateLastUpdated,
						[UpdateByUserID]	= @UpdateByUserID
				WHERE 		PolicyLimitsID		= @PolicyLimitsID
				--10.12.2005 BGS When Farmers UNDELETES a UNIT, need to update the Assignment close date to null
				UPDATE Assignments SET
					[AdminComments] 	= @sTemp,
					[StatusID] 		= (SELECT [StatusID] FROM Status WHERE [Status] = 'CLIENTReopen'), 		
					[CloseDate] 		= null,
					[DownLoadMe] 		= 1,
					[DateLastUpdated]	= @DateLastUpdated,
					[UpdateByUserID]	= @UpdateByUserID
				WHERE	[AssignmentsID] 	= @AssignmentsID
			END
			ELSE
			BEGIN
				--If its not deleted then this UNIT Assignment is an Error
				--Raise an Error
				SET @FarmersXMLAddUnit = 'Duplicate Unit Assignment to the same adjuster not allowed! UniqueID = ' + @FarmersUnitRowID
				RAISERROR (@FarmersXMLAddUnit, 16, 1)
			END
		END
		ELSE
		BEGIN
			--Otherwise need to raise error allowing this Unit to be added to exising Loss
			SET @FarmersXMLAddUnit = 'V2ECCarFarmers.clsLossXML01|AddToExistingLoss|' + Cast(@AssignmentsID As varchar(20))
			RAISERROR (@FarmersXMLAddUnit, 16, 1)
		END
	END
	ELSE
	BEGIN
		UPDATE Assignments SET	[AdminComments] 	= 'WEBCONTROL UPDATED Loss Report DATA ONLY!  probable duplicate report.',
				[LRFormat]		= @LRFormat,
				[LossReport] 		= Null, --Update this ntext field with pointer DO NOT DIRECTLY UPDATE !
				[DownLoadLossReport]	= @DownLoadLossReport,
				[UpLoadLossReport]	= @UpLoadLossReport,
				[DateLastUpdated]	= @DateLastUpdated,
				[UpdateByUserID]	= @UpdateByUserID
		WHERE AssignmentsID	=@AssignmentsID
		--Because SQL Server pukes on updates to ntext, text or image files,
		--you must set a pointer to the value you want to update and use WRITE TEXT.
		SELECT @ptrLossReport = TEXTPTR(LossReport) 
		FROM Assignments
		WHERE AssignmentsID	=@AssignmentsID
		WRITETEXT Assignments.LossReport @ptrLossReport @LossReportPrnData
		
		--Also Return this message so that WEBCONTROL is Informed of this action.
		RAISERROR ('WEBCONTROL UPDATED Loss Report DATA ONLY!  probable duplicate report.', 16, 1)
	END 
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[z_spuInsertAssignment]  TO [CFUSER]
GO

