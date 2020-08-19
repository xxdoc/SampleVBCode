SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spuEditECSFarmersXML01LossReport]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spuEditECSFarmersXML01LossReport]
GO

/******************************************************************
*
*    	Description:	Update Farmers Assignments LossReport WDDX XML
*			May include Inserts And OR Updates to...
*			Policy Limits, Assignments, FarmersXML01_LocationUpdateQue
*			Processed By WebControl (Getting WDDX TransForms from Biz Talk
*	
*	Author: 	Brad Skidmore
*	Date: 		5/22/2005
*	
******************************************************************/
CREATE PROCEDURE [dbo].[z_spuEditECSFarmersXML01LossReport]
	--*************BEGIN Merging Required Lookup Variables***********************
	--*
	--*
	--If Merging a Farmers unit into an Existing Assignment Record
	--Need to Get the Existing Assignment Record Via Lookup from 
	--ClientNumber, the ACID (AssignedTo) And the Class Name.
	@bMergeUnit			bit=0,
	@CLIENTNUM			varchar(20)=null,
	@ACIDAssignedTo			varchar(20)=null,
	@LRFormat			varchar(50)=null,
	--*
	--*
	--*************END Merging Required Lookup Variables***********************
	--*************BEGIN INSERT Policy Limits vars When Merging Unit to Existing Assignment Record***********
	@UnitLimit			varchar(20)='0',
	@UnitLimitOther			varchar(100)=' ',
	@UnitClass			varchar(1)='OTHER',
	@UnitCauseOfLoss		varchar(50)='UNKNOWN',
	--*************END INSERT Policy Limits vars When Merging Unit to Existing Assignment Record***********
	@UnitID				varchar(20)=null,
	@AssignmentsID			int=null,
	@LossReportPrnData		ntext=null,
--	For Mailing Address Update Synchro
	@MailAddress1 			varchar(250)=null,
	@MailAddress2 			varchar(250)=null,
--	For Property Address Update Synchro
	@PropAddress1			varchar(250)=null,
	@PropAddress2			varchar(250)=null,
--	Other Assignments fields Update synchro
	@PolicyNo			varchar(50)=null,
	@PolicyDescription		varchar(100)=null,
	@Insured			varchar(100)=null,
	@HomePhone			varchar(50)=null,
	@BusinessPhone			varchar(50)=null,
	@MortgageeName			varchar(100)=null,
	@AgentNo			varchar(50)=null,
	@Deductible 			Money=null,
--	For Vendor update Synchro
	@UpdUnitId			varchar(20)=null,
	@UpdClaimId			varchar(20)=null,
	@ContactRowId			varchar(20)=null,
	@PropertyRowId			varchar(20)=null,
	@PropertyDescription		varchar(50)='',
	@VehicleMake			varchar(30)='',
	@VehicleModel			varchar(30)='',
	@VehicleYear			varchar(4)='',
	@InspectedVIN			varchar(30)='',
	@LocationType			varchar(50)='',
	@LocationName			varchar(50)='',
	@LocationAddress		varchar(200)='',
	@LocationCity			varchar(50)='',
	@LocationCountry		varchar(50)='',
	@LocationState			varchar(50)='',
	@LocationPostalCode		varchar(30)='',
	@LocationPhoneNumber		varchar(50)='',
	@GMTUpdated			datetime=NULL,
	@UpdateUnitReport		varchar(8000)=Null
AS
DECLARE @UpdateByUserID 		Int
DECLARE @ptrLossReport			binary(16) --Pointer for ntext field update
--Address Vars
--Mailing
DECLARE @MailingAddress 		varchar(255)
DECLARE @MAStreet 			varchar(150)
DECLARE @MACity 			varchar(50)
DECLARE @MAState 			varchar(50)
DECLARE @MAZIP 				int
DECLARE @MAZIP4 			int
DECLARE @MAOtherPostCode 		varchar(20)
--Property
DECLARE @PropertyAddress 		varchar(255)
DECLARE @PAStreet 			varchar(150)
DECLARE @PACity 			varchar(50)
DECLARE @PAState 			varchar(50)
DECLARE @PAZIP 				int
DECLARE @PAZIP4 			int
DECLARE @PAOtherPostCode 		varchar(20)
--RTActivityLog INSERT
DECLARE	@ActDate			datetime
DECLARE @CurDate			datetime
DECLARE	@RTActivityLogID		int

SET @CurDate = GetDate()
SET @ActDate =	dbo.CleanFromOrToDate(@CurDate, @ActDate, 1)

--*************BEGIN Merging Unit to Exisiting Assignment Record***********************
--Policy Limits
DECLARE @PolicyLimitsList		varchar(8000)
--Used For Building Policy Limit List
DECLARE @COLUMN_DELIM 			varchar(1)
DECLARE @RECORD_DELIM 			varchar(1)
DECLARE @PLClassTypeID 			varchar(15)
DECLARE @PLLimitAmount 			varchar(15)
DECLARE @PLRCSaidProp 			varchar(15)
DECLARE @PLReserves 			varchar(15)
DECLARE @PLAdminComments 		varchar(50) 
DECLARE @DateLastUpdated 		datetime
DECLARE	@FarmersUnitRowID		varchar(20)
--Incase of Error
DECLARE @FarmersMergeUnitError		varchar(255)

--Get the UsersID For CFUSER
SET @UpdateByUserID = (SELECT [UsersID] FROM USERS WHERE USERNAME = 'CFUSER')
SET @DateLastUpdated = GetDate()

IF @bMergeUnit = 1 And @CLIENTNUM Is Not Null
BEGIN
	--Need to Get the Assignment Record ID that this Unit will be Merged to
	SET @AssignmentsID =	(
				SELECT 	Top 1 [AssignmentsID]
				FROM	Assignments
				WHERE	[CLIENTNUM] = @CLIENTNUM
				AND	[AdjusterSpecID] = 	(
								SELECT	TOP 1 [ClientCoAdjusterSpecID]
								FROM	ClientCoAdjusterSpec
								WHERE	[ACID] = @ACIDAssignedTo
								AND	[InactiveDate] Is Null
								AND	[ClientCompanyID] =	(
												SELECT	[ClientCompanyID]
												FROM	Company
												WHERE	[DBName] = 'Farmers'
												)
								ORDER BY [ActiveDate] Desc
								)
				AND	[LRFormat] = @LRFormat
				ORDER BY [DateLastUpdated] Desc
				)
	IF @AssignmentsID Is Null
	BEGIN
		--Raise an Error
		SET @FarmersMergeUnitError = ' Error Inserting UnitId: ' + @UnitId + '. '
		SET @FarmersMergeUnitError = @FarmersMergeUnitError + ' Attempt to Merge this Unit to an existing Assignment Record Failed! '
		RAISERROR (@FarmersMergeUnitError, 16, 1)
		RETURN 1 
	END

	--Build the Policy LImits List
	SET @COLUMN_DELIM = 	dbo.GetCOLUMN_DELIM()
	SET @RECORD_DELIM = 	dbo.GetRECORD_DELIM()
	SET @PLClassTypeID = 	(SELECT ClassTypeID FROM ClassType WHERE Class = @UnitClass)
	SET @PLLimitAmount =	@UnitLimit
	SET @PLRCSaidProp =	'0'
	SET @PLReserves =	'0'
	SET @PLAdminComments =	@UnitLimitOther--Must at least a Space in there!!!
	SET @FarmersUnitRowID = @UnitId --Must at least a Space in there!!!
	SET @PolicyLimitsList = '0' + @COLUMN_DELIM + 				--PolicyLimitsID	int,
				'0' + @COLUMN_DELIM + 				--AssignmentsID		int,
				'0' + @COLUMN_DELIM + 				--[ID]		 	int,
				'0' + @COLUMN_DELIM + 				--IDAssignments		int,
				@PLClassTypeID + @COLUMN_DELIM + 		--ClassTypeID 		int,
				@PLLimitAmount + @COLUMN_DELIM + 		--LimitAmount 		money,
				@PLRCSaidProp + @COLUMN_DELIM + 		--RCSaidProp 		money,
				@PLReserves + @COLUMN_DELIM + 			--Reserves 		money,
				'0' + @COLUMN_DELIM + 				--IsDeleted		bit,
				'0' + @COLUMN_DELIM + 				--DownLoadMe		bit,
				'0' + @COLUMN_DELIM + 				--UpLoadMe		bit,
				@PLAdminComments + @COLUMN_DELIM +		--AdminComments		Varchar(1000),
				Cast(GetDate() As varchar(50))+ @COLUMN_DELIM + -- DateLastUpdated 	datetime,
				Cast(@UpdateByUserID as Varchar(15)) + @COLUMN_DELIM +	--UpdateByUserID		int
				@FarmersUnitRowID + @RECORD_DELIM 		--FarmersUnitRowID	varchar(20)


	EXEC z_spuUpdatePolicyLimitsFromSplit
		@PolicyLimitsList	,--Policy Limits Delimited String to be parsed by SplitPolicyLimits UDF
		@AssignmentsID		,
		@DateLastUpdated 	,
		@UpdateByUserID 	

END
--*************END Merging Unit to Exisiting Assignment Record***********************

--If the UnitID is not null then Need to Return the Loss Report Data 
IF @UnitID Is Not Null 
BEGIN
	SELECT	[LossReport] As [LRData],
		[AssignmentsID]
	FROM	Assignments
	WHERE	[AssignmentsID] = 	(
					SELECT	TOP 1 [AssignmentsID]
					FROM 	PolicyLimits PL
					WHERE	PL.[FarmersUnitRowID] = @UnitID
					AND	[IsDeleted] = 0 
					ORDER BY [DateLastUpdated] DESC
					)
	RETURN -- BAIL this SP !!!!
END

--Because SQL Server pukes on updates to ntext, text or image files,
--you must set a pointer to the value you want to update and use WRITE TEXT.
SELECT 		@ptrLossReport 	= TEXTPTR(LossReport) 
FROM 		Assignments
WHERE 		[AssignmentsID]	= @AssignmentsID
WRITETEXT 	Assignments.[LossReport] @ptrLossReport @LossReportPrnData

-- Check for Synchronization required by this Update...
--	For Mailing Address Update Synchro
IF @MailAddress1 Is Not Null and @MailAddress2 Is Not Null
BEGIN
	--Clean Address info
	SET @MailingAddress 		= LTrim(RTrim(@MailAddress1)) + '    ' + LTrim(RTrim(@MailAddress2))
	EXEC z_spuCleanAndPopulateAddressInfo 
		@MailingAddress 	OUTPUT, 
		@MAStreet 		OUTPUT, 
		@MACity 		OUTPUT, 
		@MAState 		OUTPUT, 
		@MAZIP 			OUTPUT, 
		@MAZIP4 		OUTPUT,
		@MAOtherPostCode	OUTPUT
END

--	For Property Address Update Synchro
IF @PropAddress1 Is Not Null and @PropAddress2 Is Not Null
BEGIN
	SET @PropertyAddress 		= LTrim(RTrim(@PropAddress1)) + '   ' + LTrim(RTrim(@PropAddress2))
	EXEC z_spuCleanAndPopulateAddressInfo 
		@PropertyAddress 	OUTPUT, 
		@PAStreet 		OUTPUT,  
		@PACity 		OUTPUT, 
		@PAState 		OUTPUT, 
		@PAZIP 			OUTPUT, 
		@PAZIP4 		OUTPUT,
		@PAOtherPostCode	OUTPUT
END

--Now that the Trigger is RENABLED ! 
--flag this record for download... which will put a nice little
--record in the History table showing all the updates
UPDATE Assignments SET
	[MailingAddress] 	= (CASE WHEN @MailAddress1 Is Not Null and @MailAddress2 Is Not Null 	THEN @MailingAddress 	ELSE [MailingAddress] 	END),
	[MAStreet] 		= (CASE WHEN @MailAddress1 Is Not Null and @MailAddress2 Is Not Null 	THEN @MAStreet 		ELSE [MAStreet] 	END),
	[MACity] 		= (CASE WHEN @MailAddress1 Is Not Null and @MailAddress2 Is Not Null 	THEN @MACity 		ELSE [MACity] 		END),
	[MAState] 		= (CASE WHEN @MailAddress1 Is Not Null and @MailAddress2 Is Not Null 	THEN @MAState 		ELSE [MAState] 		END),
	[MAZIP] 		= (CASE WHEN @MailAddress1 Is Not Null and @MailAddress2 Is Not Null 	THEN @MAZIP 		ELSE [MAZIP] 		END),
	[MAZIP4] 		= (CASE WHEN @MailAddress1 Is Not Null and @MailAddress2 Is Not Null 	THEN @MAZIP4 		ELSE [MAZIP4] 		END),
	[MAOtherPostCode] 	= (CASE WHEN @MailAddress1 Is Not Null and @MailAddress2 Is Not Null 	THEN @MAOtherPostCode 	ELSE [MAOtherPostCode]	END),
	[PropertyAddress] 	= (CASE WHEN @PropAddress1 Is Not Null and @PropAddress2 Is Not Null 	THEN @PropertyAddress 	ELSE [PropertyAddress] 	END),
	[PAStreet] 		= (CASE WHEN @PropAddress1 Is Not Null and @PropAddress2 Is Not Null 	THEN @PAStreet 		ELSE [PAStreet] 	END),
	[PACity] 		= (CASE WHEN @PropAddress1 Is Not Null and @PropAddress2 Is Not Null 	THEN @PACity 		ELSE [PACity] 		END),
	[PAState] 		= (CASE WHEN @PropAddress1 Is Not Null and @PropAddress2 Is Not Null 	THEN @PAState 		ELSE [PAState] 		END),
	[PAZIP] 		= (CASE WHEN @PropAddress1 Is Not Null and @PropAddress2 Is Not Null 	THEN @PAZIP 		ELSE [PAZIP] 		END),
	[PAZIP4] 		= (CASE WHEN @PropAddress1 Is Not Null and @PropAddress2 Is Not Null 	THEN @PAZIP4 		ELSE [PAZIP4] 		END),
	[PAOtherPostCode] 	= (CASE WHEN @PropAddress1 Is Not Null and @PropAddress2 Is Not Null 	THEN @PAOtherPostCode 	ELSE [PAOtherPostCode] 	END),
	[PolicyNo] 		= (CASE WHEN @PolicyNo Is Not Null 					THEN @PolicyNo 		ELSE [PolicyNo] 	END),
	[PolicyDescription] 	= (CASE WHEN @PolicyDescription Is Not Null 				THEN @PolicyDescription ELSE [PolicyDescription]END),
	[Insured] 		= (CASE WHEN @Insured Is Not Null 					THEN @Insured 		ELSE [Insured] 		END),
	[HomePhone] 		= (CASE WHEN @HomePhone Is Not Null 					THEN @HomePhone 	ELSE [HomePhone]	END),
	[BusinessPhone] 	= (CASE WHEN @BusinessPhone Is Not Null 				THEN @BusinessPhone 	ELSE [BusinessPhone] 	END),
	[MortgageeName] 	= (CASE WHEN @MortgageeName Is Not Null 				THEN @MortgageeName 	ELSE [MortgageeName] 	END),
	[AgentNo] 		= (CASE WHEN @AgentNo Is Not Null 					THEN @AgentNo 		ELSE [AgentNo] 		END),
	[Deductible] 		= (CASE WHEN @Deductible Is Not Null 					THEN @Deductible 	ELSE [Deductible] 	END),
	[DownLoadMe] 		= 1,
	[DateLastUpdated] 	= GetDate(),
	[UpdateByUserID] 	= @UpdateByUserID
WHERE	[AssignmentsID]		= @AssignmentsID

/******************************************************************************************************
		Synchronize the Vendor Updates Que Table for Property (Automobile Location)
*******************************************************************************************************/
IF @UpdUnitId Is Not Null And @UpdClaimId is Not Null And @AssignmentsID Is Not Null
BEGIN
	UPDATE FarmersXML01_LocationUpdateQue SET
		[ContactRowId]		= @ContactRowId,	--varchar(20)=null
		[PropertyRowId]		= @PropertyRowId,	--varchar(20)=null
		[PropertyDescription]	= @PropertyDescription,	--varchar(50)=''
		[VehicleMake]		= @VehicleMake,		--varchar(30)=''
		[VehicleModel] 		= @VehicleModel,	--varchar(30)=''
		[VehicleYear]		= @VehicleYear,		--varchar(4)=''
		[InspectedVIN]		= @InspectedVIN,	--varchar(30)=''
		[LocationType]		= @LocationType,	--varchar(50)=''
		[LocationName]		= @LocationName,	--varchar(50)=''
		[LocationAddress]	= @LocationAddress,	--varchar(200)=''
		[LocationCity]		= @LocationCity,	--varchar(50)=''
		[LocationCountry]	= @LocationCountry,	--varchar(50)=''
		[LocationState]		= @LocationState,	--varchar(50)=''
		[LocationPostalCode]	= @LocationPostalCode,	--varchar(30)=''
		[LocationPhoneNumber]	= @LocationPhoneNumber,	--varchar(50)=''
		[GMTUpdated]		= @GMTUpdated,		--datetime=NULL
		[DateLastUpdated]	= GetDate(),		--datetime=NULL
		[UpdateByUserID]	= @UpdateByUserID	--Int
	WHERE	[AssignmentsID]		= @AssignmentsID	--Int
	AND	[UnitId]		= @UpdUnitId		--varchar(20)=null
	AND	[ClaimId]		= @UpdClaimId		--varchar(20)=null
END
/******************************************************************************************************
		INSERT Activity Entry concerning the update to this Loss Report
*******************************************************************************************************/
IF @UpdateUnitReport Is Not Null
BEGIN
	IF @bMergeUnit = 1
	BEGIN
		--If merging a unit only indicate the Added Unit Number and such
		-- Not every stinking Item in the merged unit!!!
		SET @UpdateUnitReport = 'Merged UNIT for Coverage: ' + @UnitClass 
	END
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
			 Null As [BillingCountID],
			 Null As [ID],
			 @AssignmentsID As [IDAssignments],
			 Null As [IDBillingCount],
			 0 As [ServiceTime],
			 @ActDate As [ActDate],
			 @UpdateUnitReport As [ActText],
			 @CurDate As [ActTime],
			 0 As [PageBreakAfter],
			 0 As [BlankPageAfter],
			 0 As [BlankRowsAfter],
			 0 As [IsMgrEntry],
			 0 As [IsDeleted],
			 1 As [DownLoadMe],
			 0 As [UpLoadMe],
			 'CLIENT_CO_UPDATE'As [AdminComments],
			 @CurDate As [DateLastUpdated],
			 @UpdateByUserID As [UpdateByUserID]
	--Need to Update the [ID] to be the Current Identity
	SET @RTActivityLogID = IDENT_CURRENT('RTActivityLog')
	UPDATE RTActivityLog SET [ID] = @RTActivityLogID
	WHERE 	[RTActivityLogID] = @RTActivityLogID
	AND  	[AssignmentsID] = @AssignmentsID
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[z_spuEditECSFarmersXML01LossReport]  TO [CFUSER]
GO

