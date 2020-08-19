SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spuInsertECSFarmersXML01Auto]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spuInsertECSFarmersXML01Auto]
GO

/******************************************************************
*
*    	Description:	Insert Eberls Farmers Assignments 
*			From XML01 Auto Loss Reports
*			Processed By WebControl
*	
*	Author: 	Brad Skidmore
*	Date: 		1/14/2004
*	
******************************************************************/


CREATE PROCEDURE [dbo].[z_spuInsertECSFarmersXML01Auto]
	@UserName			varchar(20),
	@LRFormatClassName		varchar(50),
	@LossReportPrnData		ntext,
	@CompanyCode			varchar(3)='ECS',
	@ClientCompanyCode		varchar(3)='FRE',
	@TypeOfACID			varchar(50)='FACT', --ACID = Adjuster Client ID'
	@DefualtFeeScheduleName		varchar(20)='AutoSched01',
	@AssignmentType			varchar(20)='Auto',
	@bAssignByZIP			bit=0,
	--XML01 Admin Loss Info Data type for AUTO
	@PolicyNumber 		varchar(100)=null,
	@Status			varchar(100)=null,
	@AgentPrimaryPhone 		varchar(100)=null,
	@HomePhoneNumber 		varchar(100)=null,
	@CellularPhoneNumber 		varchar(100)=null,
	@FRLienHolderName		varchar(100)=null,
	@FRCompanyCode 			varchar(100)=null,
	@FRPolicyDescription		varchar(100)=null,
	@NewBusinessDate 		varchar(100)=null,
	@RenewalDate 			varchar(100)=null,
	@CancellationDate 		varchar(100)=null,
	@NamedInsured 			varchar(100)=null,
	@MailAddress1 			varchar(250)=null,
	@MailAddress2 			varchar(250)=null,
	@PropAddress1			varchar(250)=null,
	@PropAddress2			varchar(250)=null,
	@VehicleDescription 		varchar(100)=null,
	@CompDed			varchar(100)=null,
	@VehicleLocation 		varchar(500)=null,
	--XML01 Current Loss Info
	@CatastropheCode		varchar(100)=null,
	@FRLossDate			varchar(100)=null,
	@AssignedTo			varchar(100)=null,
	@DateAssigned			varchar(100)=null,
	@DateClsd			varchar(100)=null,
	@ClaimNumber			varchar(100)=null,
	@UnitId			varchar(20)=' ',
	@ClaimId			varchar(20)=' ',
	@UnitLimit			varchar(20)='0',
	@UnitLimitOther			varchar(100)=' ',
	@UnitClass			varchar(1)='OTHER',
	@UnitCauseOfLoss		varchar(50)='UNKNOWN',
--****************************************************************
--*		BEGIN 	FarmersXML01_LocationUpdateQue For Auto Only
--*		5.10.2005 BGS Need to Insert Update Record for Location updates.
--*		This Record will exist FarmersXML01_LocationUpdateQue table
--*		Triggers on Assignment Updates will set SendToCRN bit column
--*		Every hour WebControl will query anything needing to be sent 
--*		to CRN.
--****************************************************************
	@ContactRowId			varchar(20),
	@PropertyRowId			varchar(20),
	@PropertyDescription		varchar(50),
	@VehicleMake			varchar(30),
	@VehicleModel			varchar(30),
	@VehicleYear			varchar(4),
	@InspectedVIN			varchar(30),
	@LocationType			varchar(50),
	@LocationName			varchar(50),
	@LocationAddress		varchar(200),
	@LocationCity			varchar(50),
	@LocationCountry		varchar(50),
	@LocationState			varchar(50),
	@LocationPostalCode		varchar(30),
	@LocationPhoneNumber		varchar(50),
	@GMTUpdated			datetime=NULL
--****************************************************************
--*		END 	FarmersXML01_LocationUpdateQue
--****************************************************************
AS
--Fields Updated from DB Lookups (z_spuInitCCAssignmentSpecs)
DECLARE @AssignmentsID			int
DECLARE	@CreatedBy			varchar(50)
DECLARE @CatCode			varchar(50)
DECLARE	@ACID				varchar(50)
DECLARE @FeeScheduleID 			int
DECLARE @AssignmentTypeID		int
DECLARE @UID				int
DECLARE @CompanyID 			int
DECLARE @CompanyName			varchar(30)
DECLARE @ClientCompanyID		int
DECLARE	@ClientCompanyName		varchar(30)
DECLARE @UNKOWNTypeOfLossID		int
DECLARE	@PropertyBaseFeeScheduleID	int
DECLARE @UACATIDProperty		int
DECLARE @UAUID				int
DECLARE @iAssignByZIP			int
DECLARE	@ClientCompanyCatSpecID		int--Assignments Table Field
DECLARE @AdjusterSpecID 		int--Assignments Table Field
DECLARE @AdjusterSpecIDDisplay 		int--Assignments Table Field

--Assignments Table Fields Updated From XML01 Parameters
DECLARE @SPVersion			int
DECLARE @CLIENTNUM 			varchar(20)
DECLARE @PolicyNo 			varchar(50)
DECLARE @PolicyDescription		varchar(100)
DECLARE @Insured 			varchar(100)
DECLARE @MailingAddress 		varchar(255)
DECLARE @MAStreet 			varchar(150)
DECLARE @MACity 			varchar(50)
DECLARE @MAState 			varchar(50)
DECLARE @MAZIP 				int
DECLARE @MAZIP4 			int
DECLARE @MAOtherPostCode 		varchar(20)
DECLARE @HomePhone 			varchar(50)
DECLARE @BusinessPhone			varchar(50)
DECLARE @PropertyAddress 		varchar(255)
DECLARE @PAStreet 			varchar(150)
DECLARE @PACity 			varchar(50)
DECLARE @PAState 			varchar(50)
DECLARE @PAZIP 				int
DECLARE @PAZIP4 			int
DECLARE @PAOtherPostCode 		varchar(20)
DECLARE @MortgageeName 			varchar(100) 
DECLARE @AgentNo 			varchar(50)
DECLARE @ReportedBy 			varchar(100)
DECLARE @ReportedByPhone 		varchar(50)
DECLARE @Deductible 			money
DECLARE @AppDedClassTypeIDOrder 	varchar(255)
DECLARE @LRFormat 			varchar(50)
--LossReport in ntext @LossReportPrnData parameter
--@DownLoadLossReport Use Default Value bit
--@UpLoadLossReport Use Default Value 	bit
DECLARE @StatusID 			int
DECLARE @TypeOfLossID 			int
DECLARE @CRNTypeOfLossID 		int
DECLARE @XactTypeOfLoss 		varchar(50)
DECLARE @SentToXact 			bit
DECLARE @LossDate 			datetime
DECLARE @AssignedDate 			datetime
-- @ReceivedDate Use Default Value	datetime
-- @ContactDate Use Default Value 	datetime
-- @InspectedDate Use Default Value	datetime
-- @CloseDate Use Default Value		datetime
-- @Reassigned Use Default Value 	bit
-- @DateReassigned Use Default Value 	datetime
-- @RAAdjusterSpecID Use Default Value 	int
-- @IsLocked 	Use Default Value	bit
-- @IsDeleted 	Use Default Value	bit
--@DownLoadMe Use Default Value		bit
--@UpLoadMe Use Default Value		bit
DECLARE @AdminComments 			varchar(1000)
DECLARE @MiscDelimSettings		varchar(8000)--MiscDelimString TEXT TYPE! Use Varchar 8000 for adding new record (This String will hold delim data of misc items applicable for each assignment)
--Policy Limits
DECLARE @PolicyLimitsList		varchar(8000)
--Used For Building Policy Limit List
DEClARE @COLUMN_DELIM 			varchar(1)
DECLARE @RECORD_DELIM 			varchar(1)
DECLARE @PLClassTypeID 			varchar(15)
DECLARE @PLLimitAmount 			varchar(15)
DECLARE @PLRCSaidProp 			varchar(15)
DECLARE @PLReserves 			varchar(15)
DECLARE @PLAdminComments 		varchar(50) 
DECLARE @DateLastUpdated 		datetime
DECLARE @UpdateByUserID 		int
DECLARE	@FarmersUnitRowID		varchar(20)
DECLARE @ThisAdjUsersID			int
DECLARE	@FarmersUAAdjError		varchar(200)
DECLARE	@FarmersUACatCodeError		varchar(200)
DECLARE	@FarmersUACauseOfLossError	varchar(200)


SET @CatCode = 	@CatastropheCode
SET @ACID =	@AssignedTo	

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

SET @HomePhone 			= LTrim(RTrim(@HomePhoneNumber))
SET @BusinessPhone		= LTrim(RTrim(@CellularPhoneNumber))
SET @PropertyAddress 		= LTrim(RTrim(@PropAddress1)) + '    ' + LTrim(RTrim(@PropAddress2)) 
EXEC z_spuCleanAndPopulateAddressInfo 
	@PropertyAddress 	OUTPUT, 
	@PAStreet 		OUTPUT,  
	@PACity 		OUTPUT, 
	@PAState 		OUTPUT, 
	@PAZIP 			OUTPUT, 
	@PAZIP4 		OUTPUT,
	@PAOtherPostCode	OUTPUT

--Check For Assign By Zip Flag
IF @bAssignByZIP =1
BEGIN
	--- use the Mailing address ZIP if the Property ZIPCODE IS Invalid
	IF (@PAZIP Is Null or @PAZIP = 0) AND (@MAZIP Is Not Null) AND (@MAZIP <> 0)
	BEGIN
		SET @iAssignByZIP = @MAZIP
	END
	ELSE
	BEGIN
		SET @iAssignByZIP = @PAZIP
	END
END
ELSE
BEGIN
	SET @iAssignByZIP = Null
END

--If the Cat code is Blank need to Reject this TransAction
IF Rtrim(@CatCode) = ''
BEGIN
	--Raise an Error
	SET @FarmersUACatCodeError = ' CatastropheCode is blank! '
	RAISERROR (@FarmersUACatCodeError, 16, 1)
	RETURN 1 
END
--If the ACID is Blank need to Reject this TransAction
IF Rtrim(@ACID) = ''
BEGIN
	--Raise an Error
	SET @FarmersUAAdjError = ' AssignedTo is blank! '
	RAISERROR (@FarmersUAAdjError, 16, 1)
	RETURN 1 
END
--Init Client Company Assignment Specs
EXEC z_spuInitCCAssignmentSpecs
	@UserName			,		
	@CompanyCode			,
	@ClientCompanyCode		,
	@CatCode			,
	@ACID				,	
	@TypeOfACID			,
	@DefualtFeeScheduleName		,
	@AssignmentType			,
	@iAssignByZIP			,
	--Output params
	@UID 				OUTPUT,
	@CompanyID			OUTPUT,
	@CompanyName			OUTPUT,
	@ClientCompanyID		OUTPUT,
	@ClientCompanyName		OUTPUT,
	@UNKOWNTypeOfLossID		OUTPUT,
	@PropertyBaseFeeScheduleID	OUTPUT,
	@UACATIDProperty		OUTPUT,
	@UAUID				OUTPUT,
	@ClientCompanyCatSpecID		OUTPUT,
	@AdjusterSpecID			OUTPUT,
	@AdjusterSpecIDDisplay		OUTPUT,
	@TypeOfLossID			OUTPUT,
	@FeeScheduleID			OUTPUT,
	@AssignmentTypeID		OUTPUT,
	@MiscDelimSettings		OUTPUT

--Since CRN is Providing a Cause Of Loss = Which Equates to TypeOfLoss.[TypeOfLoss]
--Get The Unknown Type Of Loss ID for this Client Company
SET @CRNTypeOfLossID = 	(
			SELECT 	TypeOfLossID 
			FROM	TypeOfLoss
			WHERE	[ClientCompanyID] = @ClientCompanyID
			AND 	[TypeOfLoss] = LTrim(RTrim(@UnitCauseOfLoss))
			AND 	[IsDeleted] = 0
			)

IF @CRNTypeOfLossID Is Null
BEGIN
	--Raise an Error
	SET @FarmersUACauseOfLossError = ' Cause Of Loss: ' + @UnitCauseOfLoss + ' is undefined! '
	RAISERROR (@FarmersUACauseOfLossError, 16, 1)
	RETURN 1 
END

SET @TypeOfLossID = @CRNTypeOfLossID

-- If The AdjusterSpecID is set to the Unassigned user t
--Farmers needs to be informed via Error message.
SELECT 	@ThisAdjUsersID = [USERSID]
FROM	ClientCoAdjusterSpec
WHERE	[ClientCoAdjusterSpecID] = @AdjusterSpecID

IF @ThisAdjUsersID = @UAUID
BEGIN
	--Raise an Error
	SET @FarmersUAAdjError = @ACID + ', User is not defined!'
	RAISERROR (@FarmersUAAdjError, 16, 1)
	RETURN 1 
END


--Populate the Assignments table Field Vars
SET @SPVersion			=(SELECT MAX(SPVersion)As SPVErsion FROM SoftwarePackage) 
SET @CLIENTNUM 			= LTrim(RTrim(@ClaimNumber))
SET @PolicyNo 			= LTrim(RTrim(@PolicyNumber))
SET @PolicyDescription		= LTrim(RTrim(@FRPolicyDescription))
IF LTrim(RTrim(@NamedInsured)) = ''
BEGIN
	SET @Insured 		= LTrim(RTrim(@NamedInsured))
END
ELSE
BEGIN
	SET @Insured 		= LTrim(RTrim(@NamedInsured))
END

SET @MortgageeName 		= LTrim(RTrim(@FRLienHolderName))
SET @AgentNo 			= LTrim(RTrim(@AgentPrimaryPhone))
-- SET @ReportedBy 		= LTrim(RTrim(@ali0049_ReportedBy))
-- SET @ReportedByPhone 		= LTrim(RTrim(@ali0050_RBPhone))
SET @Deductible 		= Convert(money,LTrim(RTrim(@CompDed)))
SET @AppDedClassTypeIDOrder 	= 	(	SELECT 	DefaultAppDedClassTypeIDOrder
						FROM 	FeeSchedule
						WHERE	FeeScheduleID = @FeeScheduleID
					)				

SET @LRFormat 			= @LRFormatClassName
--Set the Status to PENDING !!! for All Assignments that have yet to be downloaded!!!
SET @StatusID 			= 	(	SELECT 	StatusID 
						FROM 	Status
						WHERE 	Status = 'PENDING'
					)
--SET @TypeOfLossID 		Already Set
SET @XactTypeOfLoss 		= ''
SET @SentToXact 		=0
IF ISDate(LTrim(RTrim(@FRLossDate))) = 0
BEGIN
	SET @FRLossDate 	=Null
END
ELSE
BEGIN
	SET @LossDate 		=Convert(DateTime,LTrim(RTrim(@FRLossDate)))
END

IF ISDate(LTrim(RTrim(@DateAssigned))) = 0
BEGIN
	SET @AssignedDate	=Null
END
ELSE
BEGIN
	SET @AssignedDate 	=Convert(DateTime,LTrim(RTrim(@DateAssigned)))
END

-- SET @ReceivedDate 		=null
-- SET @ContactDate 		=null
-- SET @InspectedDate 		=null
-- SET @CloseDate 			=null
-- SET @Reassigned 		=0
-- SET @DateReassigned 		=null
-- SET @RAAdjusterSpecID 		=null
-- SET @IsLocked 			=0
-- SET @IsDeleted 			=0
-- SET @DownLoadMe 		=1
-- SET @UpLoadMe			=0
SET @AdminComments 		=null
--Build the Policy LImits List (Auto Does not have any limits in XML01 at this time)
SET @COLUMN_DELIM = 	dbo.GetCOLUMN_DELIM()
SET @RECORD_DELIM = 	dbo.GetRECORD_DELIM()
SET @PLClassTypeID = 	(SELECT ClassTypeID FROM ClassType WHERE Class = @UnitClass)
SET @PLLimitAmount =	@UnitLimit
SET @PLRCSaidProp =	'0'
SET @PLReserves =	'0'
SET @PLAdminComments =	@UnitLimitOther--Must at least a Space in there!!!
SET @FarmersUnitRowID = @UnitId --Must at least a Space in there!!!
SET @PolicyLimitsList = '0' + @COLUMN_DELIM + 				--PolicyLimitsID		int,
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
			Cast(@UID as Varchar(15)) + @COLUMN_DELIM +	--UpdateByUserID		int
			@FarmersUnitRowID + @RECORD_DELIM 		--FarmersUnitRowID	varchar(20)

SET @DateLastUpdated 		=GetDate()
SET @UpdateByUserID 		= @UID


--Finally, Insert the Assignment
EXEC z_spuInsertAssignment
	@ClientCompanyCode		, 
	default				, --@ID Use Default Value
	@AssignmentTypeID		,
	@ClientCompanyCatSpecID		,
	@AdjusterSpecID 		,
	@AdjusterSpecIDDisplay		,
	@SPVersion			,
	default				,--@IBNUM Use Default Value
	@CLIENTNUM 			,
	@PolicyNo 			,
	@PolicyDescription		,
	@Insured 			,

	@MailingAddress 		,
	@MAStreet 			,
	@MACity 			,
	@MAState 			,
	@MAZIP 				,
	@MAZIP4 			,
	@MAOtherPostCode 		,
	@HomePhone 			,
	@BusinessPhone			,
	@PropertyAddress 		,
	@PAStreet 			,
	@PACity 			,
	@PAState 			,
	@PAZIP 				,
	@PAZIP4 			,
	@PAOtherPostCode 		,
	@MortgageeName 			,
	@AgentNo 			,
	@ReportedBy 			,
	@ReportedByPhone 		,
	@Deductible 			,
	@AppDedClassTypeIDOrder 	,
	@LRFormat 			,
	@LossReportPrnData		,
	default				,--@DownLoadLossReport Use Default value
	default				,--@UpLoadLossReport Use Default Value
	@StatusID 			,
	@TypeOfLossID 			,
	@XactTypeOfLoss 		,
	@SentToXact 			,
	@LossDate 			,
	@AssignedDate 			,
	default 			,--@ReceivedDate Use Default value
	default 			,--@ContactDate Use Default value
	default 			,--@InspectedDate Use Default value
	default 			,--@CloseDate Use Default value
	default 			,--@Reassigned Use Default value
	default 			,--@DateReassigned Use Default value
	default 			,--@RAAdjusterSpecID Use Default value
	default 			,--@IsLocked Use Default value
	default 			,--@IsDeleted Use Default value
	default				,--@DownLoadMe Use Default value
	default				,--@UpLoadMe Use Default value
	@AdminComments 			,
	@MiscDelimSettings		,
	@PolicyLimitsList		,--Policy Limits
	@DateLastUpdated 		,
	@UpdateByUserID			,
	@AssignmentsID OUTPUT

--AFter Iserting record need to also insert FarmersXML01_LocationUpdateQue
--AFter Iserting record need to also insert FarmersXML01_LocationUpdateQue
--Actually need to check for previous Record... If it exists then
--Can't insert here !
DECLARE @IsThereAlready int
SET @IsThereAlready = (SELECT Top 1 [AssignmentsID] FROM FarmersXML01_ActivityUpdateQue WHERE [AssignmentsID] = @AssignmentsID  )

IF @FarmersUnitRowID Is Not Null And @PropertyRowId Is Not Null And @ContactRowID Is Not Null And @AssignmentsID Is Not Null And @IsThereAlready Is Null
BEGIN
	INSERT INTO FarmersXML01_LocationUpdateQue 	(
							[AssignmentsID],
							[UnitId],
							[ClaimId],
							[ContactRowId],
							[PropertyRowId],
							[PropertyDescription],
							[VehicleMake],
							[VehicleModel],
							[VehicleYear],
							[InspectedVIN],
							[LocationType],
							[LocationName],
							[LocationAddress],
							[LocationCity],
							[LocationCountry],
							[LocationState],
							[LocationPostalCode],
							[LocationPhoneNumber],
							[GMTUpdated],
--Default Items taken care of by Table on insert
-- 							[CRNAccepted],
-- 							[CRNAcceptedDate],
-- 							[CRNRejected],
-- 							[CRNRejectedDate],
-- 							[CRNRejectedReason],
-- 							[SendToCRN],
-- 							[SentToCRNDate],
							[DateLastUpdated],
							[UpdateByUserID]
					
							)
		SELECT			
					@AssignmentsID As [AssignmentsID],
					@UnitId As [UnitId],
					@ClaimId As [ClaimId],
					@ContactRowId As [ContactRowId],
					@PropertyRowId As [PropertyRowId],
					@PropertyDescription As [PropertyDescription],
					@VehicleMake As [VehicleMake],
					@VehicleModel As [VehicleModel],
					@VehicleYear As [VehicleYear],
					@InspectedVIN As [InspectedVIN],
					@LocationType As [LocationType],
					@LocationName As [LocationName],
					@LocationAddress As [LocationAddress],
					@LocationCity As [LocationCity],
					@LocationCountry As [LocationCountry],
					@LocationState As [LocationState],
					@LocationPostalCode As [LocationPostalCode],
					@LocationPhoneNumber As [LocationPhoneNumber],
					@GMTUpdated As [GMTUpdated],
--Default Items taken care of by Table on insert
-- 					default As [CRNAccepted],
-- 					default As [CRNAcceptedDate],
-- 					default As [CRNRejected],
-- 					default As [CRNRejectedDate],
-- 					default As [CRNRejectedReason],
-- 					default As [SendToCRN],
-- 					default As [SentToCRNDate],
					@DateLastUpdated As [DateLastUpdated],
					@UpdateByUserID As [UpdateByUserID]
	
END
--Need to insert CRN Expected Activities into Que So they can be sent to Farmers
--When they Occur
SET @CreatedBy = @AssignedTo
IF @FarmersUnitRowID Is Not Null And @AssignmentsID Is Not Null And @IsThereAlready Is Null
BEGIN
	--First Contact
	INSERT INTO FarmersXML01_ActivityUpdateQue 	(
							[AssignmentsID],
							[UnitId],
							[ClaimId],
							[Type],
							[Description],
							[Comment],
							[CreatedBy],
							[GMTCreated],
--Default Items taken care of by Table on insert
-- 							[CRNAccepted],
-- 							[CRNAcceptedDate],
-- 							[CRNRejected],
-- 							[CRNRejectedDate],
-- 							[CRNRejectedReason],
-- 							[SendToCRN],
-- 							[SentToCRNDate],
							[DateLastUpdated],
							[UpdateByUserID]
					
							)
		SELECT			@AssignmentsID As [AssignmentsID],
					@UnitId As [UnitId],
					@ClaimId As [ClaimId],
					'First Contact' As [Type],
					'Contacted Insured' As [Description],
					'' As [Comment],
					@CreatedBy As [CreatedBy],
					Null As [GMTCreated],
--Default Items taken care of by Table on insert
-- 					default As [CRNAccepted],
-- 					default As [CRNAcceptedDate],
-- 					default As [CRNRejected],
-- 					default As [CRNRejectedDate],
-- 					default As [CRNRejectedReason],
-- 					default As [SendToCRN],
-- 					default As [SentToCRNDate],
					@DateLastUpdated As [DateLastUpdated],
					@UpdateByUserID As [UpdateByUserID]
--Field Inspection
	INSERT INTO FarmersXML01_ActivityUpdateQue 	(
							[AssignmentsID],
							[UnitId],
							[ClaimId],
							[Type],
							[Description],
							[Comment],
							[CreatedBy],
							[GMTCreated],
--Default Items taken care of by Table on insert
-- 							[CRNAccepted],
-- 							[CRNAcceptedDate],
-- 							[CRNRejected],
-- 							[CRNRejectedDate],
-- 							[CRNRejectedReason],
-- 							[SendToCRN],
-- 							[SentToCRNDate],
							[DateLastUpdated],
							[UpdateByUserID]
					
							)
		SELECT			@AssignmentsID As [AssignmentsID],
					@UnitId As [UnitId],
					@ClaimId As [ClaimId],
					'Field Inspection' As [Type],
					'Inspection Completed' As [Description],
					'' As [Comment],
					@CreatedBy As [CreatedBy],
					Null As [GMTCreated],
--Default Items taken care of by Table on insert
-- 					default As [CRNAccepted],
-- 					default As [CRNAcceptedDate],
-- 					default As [CRNRejected],
-- 					default As [CRNRejectedDate],
-- 					default As [CRNRejectedReason],
-- 					default As [SendToCRN],
-- 					default As [SentToCRNDate],
					@DateLastUpdated As [DateLastUpdated],
					@UpdateByUserID As [UpdateByUserID]
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[z_spuInsertECSFarmersXML01Auto]  TO [CFUSER]
GO

