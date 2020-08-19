SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spuInsertECSFarmersCCMSProperty]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spuInsertECSFarmersCCMSProperty]
GO

/******************************************************************
*
*    	Description:	Insert Eberls Farmers Assignments 
*			From CCMS Property Loss Reports
*			Processed By WebControl
*	
*	Author: 	Brad Skidmore
*	Date: 		1/7/2004
*	
******************************************************************/


CREATE PROCEDURE [dbo].[z_spuInsertECSFarmersCCMSProperty]
	@UserName			varchar(20),
	@LRFormatClassName		varchar(50),
	@LossReportPrnData		ntext,
	@CompanyCode			varchar(3)='ECS',
	@ClientCompanyCode		varchar(3)='FRE',
	@TypeOfACID			varchar(50)='FACT', --ACID = Adjuster Client ID
	@DefualtFeeScheduleName		varchar(20)='Sched01',
	@AssignmentType			varchar(20)='Property',
	@bAssignByZIP			bit=0,
	--CCMS Admin Loss Info Data type
	@ali0004_DateTimePrinted 	varchar(100)=null,
	@ali0005_PrintedBy		varchar(100)=null,
	@ali0052_ReportedBy		varchar(100)=null,
	@ali0053_RBPhone		varchar(100)=null,
	@ali0054_PolicyNum		varchar(100)=null,
	@ali0055_SC			varchar(100)=null,
	@ali0056_AgentNum		varchar(100)=null,
	@ali0057_HomePhone		varchar(100)=null,
	@ali0058_BusPhone		varchar(100)=null,
	@ali0059_NewBusDate		varchar(100)=null,
	@ali0060_RenewalDate		varchar(100)=null,
	@ali0061_LastCancDate		varchar(100)=null,
	@ali0062_NamedInsured		varchar(100)=null,
	@ali0063_MailAddress1		varchar(250)=null,
	@ali0064_MailAddress2		varchar(250)=null,
	@ali0065_MainFileInsuredName	varchar(100)=null,
	@ali0066_MortgageHolder		varchar(100)=null,
	@ali0067_2ndMort		varchar(100)=null,
	@ali0068_CompCode		varchar(100)=null,
	@ali0069_PolicyDescription	varchar(100)=null,
	@ali0070_BldgLimit		varchar(100)=null,
	@ali0071_ContLimit		varchar(100)=null,
	@ali0072_Deductible1		varchar(100)=null,
	@ali0073_Deductible2		varchar(100)=null,
	@ali0074_Deductible3		varchar(100)=null,
	@ali0075_Deductible4		varchar(100)=null,
	@ali0076_AddlCoverage1		varchar(100)=null,
	@ali0077_AddlCoverage2		varchar(100)=null,
	@ali0078_AddlCoverage3		varchar(100)=null,
	@ali0079_AddlCoverage4		varchar(100)=null,
	@ali0080_LossLocAddress1	varchar(250)=null,
	@ali0081_LossLocAddress2	varchar(250)=null,
	--CCMS Current Loss Info
	@cli01_CAT			varchar(100)=null,
	@cli02_LossDate			varchar(100)=null,
	@cli03_Adjuster			varchar(100)=null,
	@cli04_DateAsgn			varchar(100)=null,
	@cli05_DateClsd			varchar(100)=null,
	@cli06_SALN			varchar(100)=null,
	@cli07_AdjusterOrigInfo		varchar(100)=null,
	@cli08_DateAsgnOrigInfo		varchar(100)=null,
	@cli09_DateClsdOrigInfo		varchar(100)=null,
	@cli10_PaymentsMadeThisClaim	varchar(100)=null
AS
--Fields Updated from DB Lookups (z_spuInitCCAssignmentSpecs)
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
DECLARE @AdjusterSpecIDDisplay		int--Assignments Table Field

--Assignments Table Fields Updated From CCMS Parameters
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

SET @CatCode = 	@cli01_CAT
SET @ACID =	@cli03_Adjuster	

--Clean Address info
SET @MailingAddress 		= LTrim(RTrim(@ali0063_MailAddress1)) + '    ' + LTrim(RTrim(@ali0064_MailAddress2))
EXEC z_spuCleanAndPopulateAddressInfo 
	@MailingAddress 	OUTPUT, 
	@MAStreet 		OUTPUT, 
	@MACity 		OUTPUT, 
	@MAState 		OUTPUT, 
	@MAZIP 			OUTPUT, 
	@MAZIP4 		OUTPUT,
	@MAOtherPostCode	OUTPUT

SET @HomePhone 			= LTrim(RTrim(@ali0057_HomePhone))
SET @BusinessPhone		= LTrim(RTrim(@ali0058_BusPhone))
SET @PropertyAddress 		= LTrim(RTrim(@ali0080_LossLocAddress1)) + '   ' + LTrim(RTrim(@ali0081_LossLocAddress2))
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

--Populate the Assignments table Field Vars
SET @SPVersion			=(SELECT MAX(SPVersion)As SPVErsion FROM SoftwarePackage) 
SET @CLIENTNUM 			= LTrim(RTrim(@cli06_SALN))
SET @PolicyNo 			= LTrim(RTrim(@ali0054_PolicyNum))
SET @PolicyDescription		= LTrim(RTrim(@ali0069_PolicyDescription))
IF LTrim(RTrim(@ali0065_MainFileInsuredName)) = ''
BEGIN
	SET @Insured 		= LTrim(RTrim(@ali0062_NamedInsured))
END
ELSE
BEGIN
	SET @Insured 		= LTrim(RTrim(@ali0065_MainFileInsuredName))
END

SET @MortgageeName 		= LTrim(RTrim(@ali0066_MortgageHolder))
SET @AgentNo 			= LTrim(RTrim(@ali0056_AgentNum))
SET @ReportedBy 		= LTrim(RTrim(@ali0052_ReportedBy))
SET @ReportedByPhone 		= LTrim(RTrim(@ali0053_RBPhone))
SET @Deductible 		= Convert(money,LTrim(RTrim(@ali0072_Deductible1)))
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
IF ISDate(LTrim(RTrim(@cli02_LossDate))) = 0
BEGIN
	SET @LossDate 		=Null
END
ELSE
BEGIN
	SET @LossDate 		=Convert(DateTime,LTrim(RTrim(@cli02_LossDate)))
END

IF ISDate(LTrim(RTrim(@cli04_DateAsgn))) = 0
BEGIN
	SET @AssignedDate	=Null
END
ELSE
BEGIN
	SET @AssignedDate 	=Convert(DateTime,LTrim(RTrim(@cli04_DateAsgn)))
END

-- SET @ReceivedDate 		=null
-- SET @ContactDate 		=null
-- SET @InspectedDate 		=null
-- SET @CloseDate 		=null
-- SET @Reassigned 		=0
-- SET @DateReassigned 		=null
-- SET @RAAdjusterSpecID 	=null
-- SET @IsLocked 		=0
-- SET @IsDeleted 		=0
SET @AdminComments 		=null
--Build the Policy LImits List
SET @COLUMN_DELIM = 	dbo.GetCOLUMN_DELIM()
SET @RECORD_DELIM = 	dbo.GetRECORD_DELIM()
SET @PLClassTypeID = 	(SELECT ClassTypeID FROM ClassType WHERE Class = 'A')
SET @PLLimitAmount =	@ali0070_BldgLimit
SET @PLRCSaidProp =	'0'
SET @PLReserves =	'0'
SET @PLAdminComments =	'CCMS Admin Loss Info ali0070_BldgLimit'--Must at least a Space in there!!!
SET @FarmersUnitRowID = ' '--Must at least a Space in there!!!
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
SET @PLClassTypeID = 	(SELECT ClassTypeID FROM ClassType WHERE Class = 'C')
SET @PLLimitAmount =	@ali0071_ContLimit
SET @PLRCSaidProp =	'0'
SET @PLReserves =	'0'
SET @PLAdminComments =	'CCMS Admin Loss Info ali0071_ContLimit'--Must at least a Space in there!!!
SET @PolicyLimitsList = @PolicyLimitsList + 
			'0' + @COLUMN_DELIM + 				--PolicyLimitsID		int,
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
	default				, --@ID Use Default value
	@AssignmentTypeID		,
	@ClientCompanyCatSpecID		,
	@AdjusterSpecID 		,
	@AdjusterSpecIDDisplay		,
	@SPVersion			,
	default				,--@IBNUM Use Default value
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
	@UpdateByUserID

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[z_spuInsertECSFarmersCCMSProperty]  TO [CFUSER]
GO

