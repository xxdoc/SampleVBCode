SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spuInsertECSFarmersASNProperty]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spuInsertECSFarmersASNProperty]
GO

/******************************************************************
*
*    	Description:	Insert Eberls Farmers Assignments 
*			From ASN Property Loss Reports
*			Processed By WebControl
*	
*	Author: 	Brad Skidmore
*	Date: 		1/26/2004
*	
******************************************************************/


CREATE PROCEDURE [dbo].[z_spuInsertECSFarmersASNProperty]
	@UserName			varchar(20),
	@LRFormatClassName		varchar(50),
	@LossReportPrnData		ntext,
	@CompanyCode			varchar(3)='ECS',
	@ClientCompanyCode		varchar(3)='FRE',
	@TypeOfACID			varchar(50)='CRID', --ACID = Adjuster Client ID
	@DefualtFeeScheduleName		varchar(20)='Sched01',
	@AssignmentType			varchar(20)='Property',
	@bAssignByZIP			bit=0,
	--ASN Admin Info Data type
	@AdminInfo_ADDRESS 		varchar(100)=null,
        @AdminInfo_Agent 		varchar(100)=null,
        @AdminInfo_AgentLR 		varchar(100)=null,
        @AdminInfo_CatCD 		varchar(100)=null,
        @AdminInfo_Company 		varchar(100)=null,
        @AdminInfo_CRAsgn 		varchar(100)=null,
        @AdminInfo_DateAsgn 		varchar(100)=null,
        @AdminInfo_DateRPTDByAgent 	varchar(100)=null,
        @AdminInfo_DateRPTDToAgent 	varchar(100)=null,
        @AdminInfo_District 		varchar(100)=null,
        @AdminInfo_HPhone 		varchar(100)=null,
        @AdminInfo_Insured 		varchar(100)=null,
        @AdminInfo_LoanNo 		varchar(100)=null,
        @AdminInfo_LossDate 		varchar(100)=null,
        @AdminInfo_LossLocationType 	varchar(100)=null,
        @AdminInfo_MTGAddress 		varchar(100)=null,
        @AdminInfo_MTGCode 		varchar(100)=null,
        @AdminInfo_MTGName 		varchar(100)=null,
        @AdminInfo_PolicyDesription 	varchar(100)=null,
        @AdminInfo_PolicyNumber 	varchar(100)=null,
        @AdminInfo_PropertyLocation 	varchar(100)=null,
        @AdminInfo_ReportedBy 		varchar(100)=null,
        @AdminInfo_ReportedByPhone 	varchar(100)=null,
        @AdminInfo_SALN 		varchar(100)=null,
        @AdminInfo_sT 			varchar(100)=null,
        @AdminInfo_State 		varchar(100)=null,
        @AdminInfo_TexasSuffix 		varchar(100)=null,
        @AdminInfo_TypeOfLoss 		varchar(100)=null,
        @AdminInfo_WEXT 		varchar(100)=null,
        @AdminInfo_WPhone 		varchar(100)=null,
        --Policy Info
       	@PolicyInfo_BalanceDue		varchar(100)=null,
        @PolicyInfo_BuildingLim 	varchar(100)=null,
        @PolicyInfo_ContentsLim 	varchar(100)=null,
        @PolicyInfo_Deductible 		varchar(100)=null,
	@PolicyInfo_Insured 		varchar(100)=null,
	@PolicyInfo_LastCancelDate 	varchar(100)=null,
        @PolicyInfo_MFRec 		varchar(100)=null,
        @PolicyInfo_NewBusReinDT 	varchar(100)=null,
        @PolicyInfo_RenewalDate 	varchar(100)=null
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
DECLARE @AdjusterSpecIDDisplay 		int--Assignments Table Field

--Assignments Table Fields Updated From ASN Parameters
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

SET @CatCode = 	@AdminInfo_CatCD
SET @ACID =	@AdminInfo_CRAsgn	

--Clean Address Info
SET @MailingAddress 		= LTrim(RTrim(@AdminInfo_ADDRESS)) 
EXEC z_spuCleanAndPopulateAddressInfo 
	@MailingAddress 	OUTPUT, 
	@MAStreet 		OUTPUT, 
	@MACity 		OUTPUT, 
	@MAState 		OUTPUT, 
	@MAZIP 			OUTPUT, 
	@MAZIP4 		OUTPUT,
	@MAOtherPostCode	OUTPUT

SET @HomePhone 			= LTrim(RTrim(@AdminInfo_HPhone))
SET @BusinessPhone		= LTrim(RTrim(@AdminInfo_WPhone))
SET @PropertyAddress 		= LTrim(RTrim(@AdminInfo_PropertyLocation)) 
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
SET @CLIENTNUM 			= LTrim(RTrim(@AdminInfo_SALN))
SET @PolicyNo 			= LTrim(RTrim(@AdminInfo_PolicyNumber))
SET @PolicyDescription		= LTrim(RTrim(@AdminInfo_PolicyDesription))
IF LTrim(RTrim(@AdminInfo_Insured)) = ''
BEGIN
	SET @Insured 		= LTrim(RTrim(@PolicyInfo_Insured))
END
ELSE
BEGIN
	SET @Insured 		= LTrim(RTrim(@AdminInfo_Insured))
END

SET @MortgageeName 		= LTrim(RTrim(@AdminInfo_MTGName))
SET @AgentNo 			= LTrim(RTrim(@AdminInfo_Agent))
SET @ReportedBy 		= LTrim(RTrim(@AdminInfo_ReportedBy))
SET @ReportedByPhone 		= LTrim(RTrim(@AdminInfo_ReportedByPhone))
SET @Deductible 		= Convert(money,LTrim(RTrim(@PolicyInfo_Deductible)))
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
IF ISDate(LTrim(RTrim(@AdminInfo_LossDate))) = 0
BEGIN
	SET @LossDate 		=Null
END
ELSE
BEGIN
	SET @LossDate 		=Convert(DateTime,LTrim(RTrim(@AdminInfo_LossDate)))
END

IF ISDate(LTrim(RTrim(@AdminInfo_DateAsgn))) = 0
BEGIN
	SET @AssignedDate	=Null
END
ELSE
BEGIN
	SET @AssignedDate 	=Convert(DateTime,LTrim(RTrim(@AdminInfo_DateAsgn)))
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
--Build the Policy LImits List
SET @COLUMN_DELIM = 	dbo.GetCOLUMN_DELIM()
SET @RECORD_DELIM = 	dbo.GetRECORD_DELIM()
SET @PLClassTypeID = 	(SELECT ClassTypeID FROM ClassType WHERE Class = 'A')
SET @PLLimitAmount =	@PolicyInfo_BuildingLim
SET @PLRCSaidProp =	'0'
SET @PLReserves =	'0'
SET @PLAdminComments =	'ASN PolicyInfo_BuildingLim'--Must at least a Space in there!!!
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
SET @PLLimitAmount =	@PolicyInfo_ContentsLim
SET @PLRCSaidProp =	'0'
SET @PLReserves =	'0'
SET @PLAdminComments =	'ASN PolicyInfo_ContentsLim'--Must at least a Space in there!!!
SET @FarmersUnitRowID = ' '--Must at least a Space in there!!!
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
	default				, --@ID
	@AssignmentTypeID		,
	@ClientCompanyCatSpecID		,
	@AdjusterSpecID 		,
	@AdjusterSpecIDDisplay		,
	@SPVersion			,
	default				,--@IBNUM
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

GRANT  EXECUTE  ON [dbo].[z_spuInsertECSFarmersASNProperty]  TO [CFUSER]
GO

