SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spsGetAssignmentsInfo]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spsGetAssignmentsInfo]
GO

/******************************************************************
*
*    	Description:	Get Assignments Info
*			
*	
*	Author: 	Brad Skidmore
*	Date: 		1/19/2004
*	Date		2/17/2004 Added DownLoad and Upload Loss Report Flags
******************************************************************/


CREATE PROCEDURE [dbo].[z_spsGetAssignmentsInfo]
	@bHideDeleted		bit =1, 
	@UID			int,
	@AssignmentsID		Int=null,
--Include (INNER JOIN) Dependant Tables
	@bShowAssignmentsInfo			bit=0,
	@bShowAssignmentTypeInfo		bit=0,
	@bShowCompanyInfo			bit=0,
	@bShowCatInfo				bit=0,
	@bShowClientCompanyInfo			bit=0,
	@bShowFeeScheduleInfo			bit=0,
	@bShowClientCompanyCatInfo		bit=0,
	@bShowClientCompanyCatSpecInfo		bit=0,
	@bShowAdjusterSpecUsersInfo		bit=0,
	@bShowAdjusterSpecInfo			bit=0,
	@bShowStatusInfo			bit=0,
	@bShowTypeOfLossInfo			bit=0,
	@bShowRAAdjusterSpecUsersInfo		bit=0,
	@bShowRAAdjusterSpecInfo		bit=0,
	@bShowBatchesInfo			bit=0,
--Sorting Grouping Searching
	@OrderBy		varchar(500)=null,
	@GroupBy		varchar(500)=null,
	@SearchBy		varchar(2000)=null,
--updating the Assignemtns Table
	@bDelete		bit=0,
	@bAdd			bit=0,
	@bEdit			bit=0,
--Update variables
	@ID 			int= NULL ,
	@AssignmentTypeID 	int= NULL ,
	@ClientCompanyCatSpecID	int= NULL ,
	@AdjusterSpecID		int= NULL ,
	@AdjusterSpecIDDisplay	int= NULL ,
	@SPVersion		int= NULL ,
	@IBNUM			varchar(20)= NULL ,
	@CLIENTNUM		varchar(20)= NULL ,
	@PolicyNo		varchar(50)= NULL ,
	@PolicyDescription	varchar(100)=null,
	@Insured 		varchar(100)=null,
	@MailingAddress 	varchar(255)=null,
	@MAStreet 		varchar(150)=null,
	@MACity			varchar(50)= NULL ,
	@MAState		varchar(50)= NULL ,
	@MAZIP			int= NULL ,
	@MAZIP4			int= NULL ,
	@MAOtherPostCode	varchar(20)= NULL ,
	@HomePhone 		varchar(50)=null,
	@BusinessPhone		varchar(50)=null,
	@PropertyAddress 	varchar(255)=null,
	@PAStreet 		varchar(150)=null,
	@PACity			varchar(50)= NULL ,
	@PAState 		varchar(50)=null,
	@PAZIP 			int=null,
	@PAZIP4 		int=null,
	@PAOtherPostCode 	varchar(20)=null,
	@MortgageeName 		varchar(100)=null,
	@AgentNo 		varchar(50)=null,
	@ReportedBy 		varchar(100)=null,
	@ReportedByPhone 	varchar(50)=null,
	@Deductible 		money=null,
	@AppDedClassTypeIDOrder varchar(255)=null,
	@LRFormat 		varchar(50)=null,
	@LossReport		ntext= NULL ,
	@DownLoadLossReport	bit=0,
	@UploadLossReport	bit=0,
	@StatusID		int= NULL ,
	@TypeOfLossID		int= NULL ,
	@XactTypeOfLoss		varchar(50)= NULL ,
	@SentToXact		bit= NULL ,
	@LossDate		datetime= NULL ,
	@AssignedDate		datetime= NULL ,
	@ReceivedDate		datetime= NULL ,
	@ContactDate		datetime= NULL ,
	@InspectedDate		datetime= NULL ,
	@CloseDate		datetime= NULL ,
	@Reassigned		bit= NULL ,
	@DateReassigned		datetime= NULL ,
	@RAAdjusterSpecID	int= NULL ,
	@IsLocked		bit= NULL ,
	@IsDeleted		bit= NULL ,
	@DownLoadMe		bit= NULL ,
	@UpLoadMe		bit= NULL ,
	@AdminComments		varchar(1000)= NULL ,
	@MiscDelimSettings	text=null,			--MiscDelimString TEXT TYPE!(This String will hold delim data of misc items applicable for each assignment)
	@DateLastUpdated	datetime= NULL, 
--ViewHistory
	@bViewHistory		bit=NULL,
	@FromDate		varchar(50)=Null,
	@ToDate			varchar(50)=Null,
--OtherSearchDateStringSpans
	@FromLossDate		varchar(50)='*',
	@ToLossDate		varchar(50)='*',
	@FromAssignedDate	varchar(50)='*',
	@ToAssignedDate		varchar(50)='*',
	@FromReceivedDate	varchar(50)='*',
	@ToReceivedDate		varchar(50)='*',
	@FromContactDate	varchar(50)='*',
	@ToContactDate		varchar(50)='*',
	@FromInspectedDate	varchar(50)='*',
	@ToInspectedDate	varchar(50)='*',
	@FromCloseDate		varchar(50)='*',
	@ToCloseDate		varchar(50)='*',
	@FromDateReassigned	varchar(50)='*',
	@ToDateReassigned	varchar(50)='*',
	@FromDateLastUpdated	varchar(50)='*',
	@ToDateLastUpdated	varchar(50)='*',
--OtherSearchNumSpans
--Mailing PropertyAdrresses
	@SearchMAAddress 	varchar(255)='*',
	@SearchMAStreet 	varchar(150)='*',
	@SearchMACity		varchar(50)='*',
	@SearchMAState		varchar(50)='*',
	@FromMAZIP		varchar(5)='*',
	@ToMAZIP		varchar(5)='*',
	@FromMAZIP4		varchar(4)='*',
	@ToMAZIP4		varchar(4)='*',
	@SearchPAAddress 	varchar(255)='*',
	@SearchPAStreet 	varchar(150)='*',
	@SearchPACity		varchar(50)='*',
	@SearchPAState 		varchar(50)='*',
	@FromPAZIP		varchar(5)='*',
	@ToPAZIP		varchar(5)='*',
	@FromPAZIP4		varchar(4)='*',
	@ToPAZIP4		varchar(4)='*',
--Report Specs
	@SelCompanyID		int=null,
	@SelClientCompanyID	int=null,
	@SelCatID		int=null,
	@SelClientCompanyUserID	int=null,
	@SelAssignmentTypeID	int=null,
	@FromDeductible		varchar(10)='*',
	@ToDeductible		varchar(10)='*',
--Policy Limits List
	@PolicyLimitsList	text=Null,--Contains String of Delim Policy Limits
--AssignByZipCode		
	@iAssignByZip		int=null, --When Adding an Assignment this flag will Assign by ZipCode 
--Download Or Upload All. If true All dependant table downloadme and uploadme flags will will be set 
--including file attachment download upload flags.
	@DownLoadAll		bit=null ,
	@UpLoadAll		bit=null 

AS
DECLARE @bResetDownLoadAll int
DECLARE @bResetUpLoadAll int
DECLARE @bFlagDownLoad int
DECLARE @bFlagUpLoad int
DECLARE @FEESCHEDULEMiscDelimSettings varchar(8000)
DECLARE @FeeScheduleID		int
DECLARE @SearchDates		varchar(2000)
DECLARE @SearchProperty		varchar(2000)
DECLARE	@SearchMailing		varchar(2000)
DECLARE @SearchReportSpecs	varchar(2000)
DECLARE @ClientCompanyCode	varchar(3) -- Used for building IBNumber when Inserting new assignment.
DECLARE @ptrLossReport		binary(16) --Pointer for ntext field update
DECLARE @ptrMiscDelimSettings	binary(16) --Pointer for ntext field update
DECLARE @sTempSQL 		varchar(1000)
DECLARE @sTempSQL2 		varchar(1000)
DECLARE @sSQL1 			varchar(8000)
DECLARE @sSQL2 			varchar(8000)
DECLARE @sSQL3 			varchar(8000)
DECLARE @sSQL4 			varchar(8000)
DECLARE @sSQL5 			varchar(8000)
DECLARE @sSQL6 			varchar(8000)
DECLARE @sSQL7 			varchar(8000)
DECLARE @sSQL8 			varchar(8000)
DECLARE @sSQL9 			varchar(8000)
DECLARE @sSQL10 		varchar(8000)
DECLARE @UAUID			int		--Holds the USERSID FOR 'UNASSINGED' User Name
DECLARE @ClientCompanyName	varchar(30)	--Holds the Name of the Client Company
DECLARE @ClientCompanyID	int		--Holds ID of the Client Company
DECLARE @CurDate 		DateTime
SET @CurDate = GetDate()
--See if the DateLastupdated is null
IF IsDate(@DateLastUpdated) =0
BEGIN
	SET @DateLastUpdated = GetDate()
END
--Clean Dates
SET @FromDate =			dbo.CleanFromOrToDate(@CurDate, @FromDate, 1)
SET @ToDate = 			dbo.CleanFromOrToDate(@CurDate, @ToDate, 0)
--Clean Date Strings
SET @FromLossDate =		dbo.CleanFromOrToDateString(@CurDate, @FromLossDate, 1)
SET @ToLossDate =		dbo.CleanFromOrToDateString(@CurDate, @ToLossDate, 0)
SET @FromAssignedDate =		dbo.CleanFromOrToDateString(@CurDate, @FromAssignedDate, 1)
SET @ToAssignedDate =		dbo.CleanFromOrToDateString(@CurDate, @ToAssignedDate, 0)
SET @FromReceivedDate =		dbo.CleanFromOrToDateString(@CurDate, @FromReceivedDate, 1)
SET @ToReceivedDate =		dbo.CleanFromOrToDateString(@CurDate, @ToReceivedDate, 0)
SET @FromContactDate =		dbo.CleanFromOrToDateString(@CurDate, @FromContactDate, 1)
SET @ToContactDate =		dbo.CleanFromOrToDateString(@CurDate, @ToContactDate, 0)
SET @FromInspectedDate =	dbo.CleanFromOrToDateString(@CurDate, @FromInspectedDate, 1)
SET @ToInspectedDate =		dbo.CleanFromOrToDateString(@CurDate, @ToInspectedDate, 0)
SET @FromCloseDate =		dbo.CleanFromOrToDateString(@CurDate, @FromCloseDate, 1)
SET @ToCloseDate =		dbo.CleanFromOrToDateString(@CurDate, @ToCloseDate, 0)
SET @FromDateReassigned =	dbo.CleanFromOrToDateString(@CurDate, @FromDateReassigned, 1)
SET @ToDateReassigned =		dbo.CleanFromOrToDateString(@CurDate, @ToDateReassigned, 0)
SET @FromDateLastUpdated =	dbo.CleanFromOrToDateString(@CurDate, @FromDateLastUpdated, 1)
SET @ToDateLastUpdated =	dbo.CleanFromOrToDateString(@CurDate, @ToDateLastUpdated, 0)

--If DownloadAll or UploadAll then need to be sure 
--the assignment DownLoadMe and Upload Me are Checked as well
IF @DownLoadAll = 1
BEGIN
	SET @bFlagDownLoad = 1
	SET @DownLoadLossReport = 1
	SET @DownLoadMe = 1
END
IF @DownLoadAll = 0
BEGIN
	IF (SELECT DownLoadAll FROM Assignments WHERE AssignmentsID = @AssignmentsID) = 1
	BEGIN
		SET @bResetDownLoadAll = 1
		SET @bFlagDownLoad = 0
		SET @DownLoadLossReport = 0
		SET @DownLoadMe = 0
	END
END

IF @UpLoadAll = 1
BEGIN
	SET @bFlagUpLoad = 1
	SET @UploadLossReport = 1
	SET @UploadMe = 1
END
IF @UpLoadAll = 0
BEGIN
	IF (SELECT UpLoadAll FROM Assignments WHERE AssignmentsID = @AssignmentsID) = 1
	BEGIN
		SET @bResetUpLoadAll = 1
		SET @bFlagUpLoad = 0
		SET @UploadLossReport = 0
		SET @UploadMe = 0
	END
END

IF @AssignmentsID Is Not Null And @bViewHistory is Null
BEGIN
	IF @bDelete=0 And @bAdd=0 And @bEdit=0
	BEGIN
		SELECT 
			[AssignmentsID] ,
			[ID],
			[AssignmentTypeID],
			[ClientCompanyCatSpecID],
			[AdjusterSpecID] ,
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
			[HomePhone] ,
			[BusinessPhone],
			[PropertyAddress],
			[PAStreet] ,
			[PACity] ,
			[PAState],
			[PAZIP] ,
			[PAZIP4],
			[PAOtherPostCode] ,
			[MortgageeName] ,

			[AgentNo] ,
			[ReportedBy],
			[ReportedByPhone],
			[Deductible] ,
			[AppDedClassTypeIDOrder],
			[LRFormat] ,
			[LossReport],
			[DownLoadLossReport],
			[UpLoadLossReport],
			[StatusID] ,
			[TypeOfLossID],
			[XactTypeOfLoss],
			[SentToXact] ,
			[LossDate] ,
			[AssignedDate],
			[ReceivedDate],
			[ContactDate],
			[InspectedDate],
			[CloseDate] ,
			[Reassigned] ,
			[DateReassigned],
			[RAAdjusterSpecID],
			[IsLocked],
			[IsDeleted],
			[DownLoadMe],
			[UpLoadMe] ,
			[DownLoadAll],
			[UpLoadAll] ,
			[AdminComments],
			[MiscDelimSettings],
			[DateLastUpdated],
			[UpdateByUserID] 
		FROM	Assignments
		WHERE 	[AssignmentsID] = @AssignmentsID
		ORDER BY [AssignedDate] 
	END
	ELSE
	BEGIN
		IF @bDelete =1
		BEGIN
			UPDATE Assignments SET 	[IsDeleted]=	(CASE 	WHEN [IsDeleted]= 0
									THEN 1
									ELSE 0
									END
								),
						[DownLoadMe] = 1,
						[DateLastUpdated]=@DateLastUpdated ,
						[UpdateByUserID]=@UID
			WHERE AssignmentsID = @AssignmentsID
		
		END
		ELSE 
		BEGIN
			IF @bAdd=1
			BEGIN
					--Get the Client CompanyID
					SET @ClientCompanyID =		(	
										SELECT 	ClientCompanyID	
										FROM	ClientCompanyCatSpec
										WHERE	ClientCompanyCatSpecID = @ClientCompanyCatSpecID
									)
					--Get the CLient Company CODE
					SET @ClientCompanyCode =	(	
										SELECT	Code
										FROM	Company
										WHERE 	CompanyID =@ClientCompanyID
									)
					--GET The Client Company Name
					SET @ClientCompanyName =	(
									SELECT 	[Name]
									FROM 	Company
									WHERE	CompanyID = @ClientCompanyID
									)
					--GET The PENDING Status ID
					SET @StatusID 			= 	(	SELECT 	StatusID 
											FROM 	Status
											WHERE 	Status = 'PENDING'
										)
					--Get the Software Package Version				
					SET @SPVersion =	(	
									SELECT 	MAX(SPVersion)As SPVErsion 
									FROM 	SoftwarePackage
								)
					--Get the UNASSIGNED USERNAME ID
					SET @UAUID = 	(	
								SELECT 	UsersID 
								FROM 	Users 
								WHERE 	UserName = 'UUNASSIGNED'
							)
					--Set the MiscDelim Settings from FeeSchedule
					SET @FeeScheduleID =	(
									SELECT 	FeeScheduleID
									FROM	ClientCompanyCat
									WHERE	ClientCompanyID = @ClientCompanyID
									AND	CATID = 	(
													SELECT	CATID
													FROM	ClientCompanyCatSpec
													WHERE	ClientCompanyCatSpecID =  @ClientCompanyCatSpecID
												)
								)
					SET @FEESCHEDULEMiscDelimSettings = dbo.GetFEESCHEDULEMiscDelimSettings(@FeeScheduleID)
					--If Assigning By ZIPCODE Reset the @AdjusterSpecID
					IF @iAssignByZip Is Not Null
					BEGIN
						EXEC z_spuInitAssignByZipcode
							@UID			,
							@UAUID			,
							@ClientCompanyID	,
							@ClientCompanyCatSpecID	,
							@ClientCompanyName	,
							@iAssignByZIP		,	
							
							--Output params
							@AdjusterSpecID		OUTPUT
					END
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
						@LossReport			,
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
						@FEESCHEDULEMiscDelimSettings	,--MiscDelimString TEXT TYPE!(This String will hold delim data of misc items applicable for each assignment)
						@PolicyLimitsList		,--Policy Limits
						@DateLastUpdated 		,
						@UID
					
			END
			ELSE 
			BEGIN
				IF @bEdit=1
				BEGIN
					Update Assignments Set 	
								[ID] = @ID,
								[AssignmentTypeID] = @AssignmentTypeID,
								[ClientCompanyCatSpecID] = @ClientCompanyCatSpecID,
								[AdjusterSpecID] = @AdjusterSpecID ,
								[AdjusterSpecIDDisplay] =@AdjusterSpecIDDisplay,
								[SPVersion] = @SPVersion,
								[IBNUM] = @IBNUM,
								[CLIENTNUM] = @CLIENTNUM,
								[PolicyNo] = @PolicyNo,
								[PolicyDescription] = @PolicyDescription,
								[Insured] = @Insured,
								[MailingAddress] = @MailingAddress,
								[MAStreet] = @MAStreet,
								[MACity] = @MACity,
								[MAState] = @MAState,
								[MAZIP] = @MAZIP,
								[MAZIP4] = @MAZIP4,
								[MAOtherPostCode] = @MAOtherPostCode,
								[HomePhone]  = @HomePhone,
								[BusinessPhone] = @BusinessPhone,
								[PropertyAddress] = @PropertyAddress,
								[PAStreet]  = @PAStreet,
								[PACity]  = @PACity,
								[PAState] = @PAState,
								[PAZIP]  = @PAZIP,
								[PAZIP4] = @PAZIP4,
								[PAOtherPostCode]  = @PAOtherPostCode,
								[MortgageeName]  = @MortgageeName,
								[AgentNo]  = @AgentNo,
								[ReportedBy] = @ReportedBy,
								[ReportedByPhone] = @ReportedByPhone,
								[Deductible]  = @Deductible,
								[AppDedClassTypeIDOrder] = @AppDedClassTypeIDOrder,
								[LRFormat]  = @LRFormat,
								[LossReport] = [LossReport], ----Update this ntext field with pointer DO NOT DIRECTLY UPDATE !
								[DownLoadLossReport] = @DownLoadLossReport,
								[UpLoadLossReport] = @UpLoadLossReport,
								[StatusID]  = @StatusID,
								[TypeOfLossID] = @TypeOfLossID,
								[XactTypeOfLoss] = @XactTypeOfLoss,
								[SentToXact] = @SentToXact ,
								[LossDate] = @LossDate ,
								[AssignedDate] = @AssignedDate,
								[ReceivedDate] = @ReceivedDate,
								[ContactDate] = @ContactDate,
								[InspectedDate] = @InspectedDate,
								[CloseDate]  = @CloseDate,
								[Reassigned]  = @Reassigned,
								[DateReassigned] = @DateReassigned,
								[RAAdjusterSpecID] = @RAAdjusterSpecID,
								[IsLocked] = @IsLocked,
								[IsDeleted] = @IsDeleted,
								[DownLoadMe] = @DownLoadMe,
								[UpLoadMe] = @UpLoadMe ,
								[DownLoadAll] = @DownLoadAll,
								[UpLoadAll] = @UpLoadAll ,
								[AdminComments] = @AdminComments,
								[MiscDelimSettings] = [MiscDelimSettings], ----Update this ntext field with pointer DO NOT DIRECTLY UPDATE !
								[DateLastUpdated] = @DateLastUpdated ,
								[UpdateByUserID] = @UID 
					WHERE AssignmentsID = @AssignmentsID
					--Because SQL Server pukes on updates to ntext, text or image files,
					--you must set a pointer to the value you want to update and use WRITE TEXT.
					--If Loss Report is set to '-1' means maintain the loss report 
					--that already exists in the DB
					IF Cast(@LossReport as varchar(2)) <> '-1'
					BEGIN 
						--V2ECcarFarmers.clsLossXML01
						-- Do NOT allow changes to Farmers Loss report clsLossXML01
						IF (SELECT Ass.[LRFormat] FROM Assignments Ass WHERE Ass.AssignmentsID = @AssignmentsID) <> 'V2ECcarFarmers.clsLossXML01'
						BEGIN
							SELECT 	@ptrLossReport = TEXTPTR(LossReport) 
							FROM 	Assignments
							WHERE 	AssignmentsID	=@AssignmentsID
							WRITETEXT Assignments.LossReport @ptrLossReport @LossReport
						END
					END
					IF Cast(@MiscDelimSettings as varchar(2)) <> '-1'
					BEGIN 
						SELECT 	@ptrMiscDelimSettings = TEXTPTR(MiscDelimSettings) 
						FROM 	Assignments
						WHERE 	AssignmentsID	=@AssignmentsID
						WRITETEXT Assignments.MiscDelimSettings @ptrMiscDelimSettings @MiscDelimSettings
					END
					--Also need to Update policy Limits
					IF @PolicyLimitsList Is Not Null
					BEGIN
						 
						EXEC z_spuUpdatePolicyLimitsFromSplit
							@PolicyLimitsList	,--Policy Limits Delimited String to be parsed by SplitPolicyLimits UDF
							@AssignmentsID		,
							@DateLastUpdated 	,
							@UID 	
					END
					--If @DownloadAll is Set then Need to Update Dependant Tables 
					--Need to update the DownloadMe Flags as well as the DownLoad Flags for Attached Files
					IF @DownLoadAll = 1 Or @bResetDownLoadAll = 1
					BEGIN
						--"RTIB"
						UPDATE 	RTIB SET 	
							[DownLoadMe] 		= @bFlagDownLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						--"RTIBFee"
						UPDATE 	RTIBFee SET 	
							[DownLoadMe] 		= @bFlagDownLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						--"IB"
						UPDATE 	IB SET 	
							[DownLoadMe] 		= @bFlagDownLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						--"IBFee"
						UPDATE 	IBFee SET 	
							[DownLoadMe] 		= @bFlagDownLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						--"BillingCount"
						UPDATE 	BillingCount SET 	
							[DownLoadMe] 		= @bFlagDownLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						--"PolicyLimits"
						UPDATE 	PolicyLimits SET 	
							[DownLoadMe] 		= @bFlagDownLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							--Don't update Admin comments here!
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						--"Package"
						UPDATE 	Package SET 	
							[DownLoadMe] 		= @bFlagDownLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						--"PackageItem"
						UPDATE 	PackageItem SET 	
							[DownLoadMe] 		= @bFlagDownLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						--"RTChecks"
						UPDATE 	RTChecks SET 	
							[DownLoadMe] 		= @bFlagDownLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						--"RTIndemnity"
						UPDATE 	RTIndemnity SET 	
							[DownLoadMe] 		= @bFlagDownLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						--"RTActivityLog"
						UPDATE 	RTActivityLog SET 	
							[DownLoadMe] 		= @bFlagDownLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						--"RTActivityLogInfo"
						UPDATE 	RTActivityLogInfo SET 	
							[DownLoadMe] 		= @bFlagDownLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						--"RTPhotoLogReport
						UPDATE RTPhotoReport SET
							[DownLoadMe] 		= @bFlagDownLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						--"RTPhotoLog"
						UPDATE 	RTPhotoLog SET 	
							[DownloadPhoto]		= @bFlagDownLoad,
							[DownloadPhotoThumb]	= @bFlagDownLoad,
							[DownloadPhotoHighRes]	= @bFlagDownLoad,
							[DownLoadMe] 		= @bFlagDownLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						--RTWSDiagram
						UPDATE RTWSDiagram SET
							[DownloadDiagramPhoto]	= @bFlagDownLoad,
							[DownLoadMe] 		= @bFlagDownLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						--"RTAttachments"
						UPDATE 	RTAttachments SET 	
							[DownloadAttachment]	= @bFlagDownLoad,
							[DownLoadMe] 		= @bFlagDownLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						--MiscReportParam
						UPDATE 	MiscReportParam SET 	
							[DownLoadMe] 		= @bFlagDownLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						UPDATE 	MiscReportParam01 SET 	
							[DownLoadMe] 		= @bFlagDownLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						UPDATE 	MiscReportParam02 SET 	
							[DownLoadMe] 		= @bFlagDownLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						UPDATE 	MiscReportParam03 SET 	
							[DownLoadMe] 		= @bFlagDownLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						UPDATE 	MiscReportParam04 SET 	
							[DownLoadMe] 		= @bFlagDownLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						UPDATE 	MiscReportParam05 SET 	
							[DownLoadMe] 		= @bFlagDownLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						UPDATE 	MiscReportParam06 SET 	
							[DownLoadMe] 		= @bFlagDownLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						UPDATE 	MiscReportParam07 SET 	
							[DownLoadMe] 		= @bFlagDownLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						UPDATE 	MiscReportParam08 SET 	
							[DownLoadMe] 		= @bFlagDownLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						UPDATE 	MiscReportParam09 SET 	
							[DownLoadMe] 		= @bFlagDownLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						UPDATE 	MiscReportParam10 SET 	
							[DownLoadMe] 		= @bFlagDownLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						UPDATE 	MiscReportParam11 SET 	
							[DownLoadMe] 		= @bFlagDownLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						UPDATE 	MiscReportParam12 SET 	
							[DownLoadMe] 		= @bFlagDownLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						UPDATE 	MiscReportParam13 SET 	
							[DownLoadMe] 		= @bFlagDownLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						UPDATE 	MiscReportParam14 SET 	
							[DownLoadMe] 		= @bFlagDownLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						UPDATE 	MiscReportParam15 SET 	
							[DownLoadMe] 		= @bFlagDownLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						UPDATE 	MiscReportParam16 SET 	
							[DownLoadMe] 		= @bFlagDownLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						UPDATE 	MiscReportParam17 SET 	
							[DownLoadMe] 		= @bFlagDownLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						UPDATE 	MiscReportParam18 SET 	
							[DownLoadMe] 		= @bFlagDownLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						UPDATE 	MiscReportParam19 SET 	
							[DownLoadMe] 		= @bFlagDownLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						UPDATE 	MiscReportParam20 SET 	
							[DownLoadMe] 		= @bFlagDownLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						UPDATE 	MiscReportParam21 SET 	
							[DownLoadMe] 		= @bFlagDownLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						UPDATE 	MiscReportParam22 SET 	
							[DownLoadMe] 		= @bFlagDownLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						UPDATE 	MiscReportParam23 SET 	
							[DownLoadMe] 		= @bFlagDownLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						UPDATE 	MiscReportParam24 SET 	
							[DownLoadMe] 		= @bFlagDownLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						UPDATE 	MiscReportParam25 SET 	
							[DownLoadMe] 		= @bFlagDownLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						UPDATE 	MiscReportParam26 SET 	
							[DownLoadMe] 		= @bFlagDownLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						UPDATE 	MiscReportParam27 SET 	
							[DownLoadMe] 		= @bFlagDownLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						UPDATE 	MiscReportParam28 SET 	
							[DownLoadMe] 		= @bFlagDownLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						UPDATE 	MiscReportParam29 SET 	
							[DownLoadMe] 		= @bFlagDownLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						UPDATE 	MiscReportParam30 SET 	
							[DownLoadMe] 		= @bFlagDownLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
					END

					--If @UpLoadAll is Set then Need to Update Dependant Tables 
					--Need to update the DownloadMe Flags as well as the DownLoad Flags for Attached Files
					IF @UpLoadAll = 1 Or @bResetUpLoadAll = 1
					BEGIN
						--"RTIB"
						UPDATE 	RTIB SET
							[DownLoadMe] 		= @bFlagUpLoad, 	
							[UpLoadMe] 		= @bFlagUpLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						--"RTIBFee"
						UPDATE 	RTIBFee SET
							[DownLoadMe] 		= @bFlagUpLoad, 	
							[UpLoadMe] 		= @bFlagUpLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						--"IB"
						UPDATE 	IB SET 	
							[DownLoadMe] 		= @bFlagUpLoad,
							[UpLoadMe] 		= @bFlagUpLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						--"IBFee"
						UPDATE 	IBFee SET 	
							[DownLoadMe] 		= @bFlagUpLoad,
							[UpLoadMe] 		= @bFlagUpLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						--"BillingCount"
						UPDATE 	BillingCount SET 
							[DownLoadMe] 		= @bFlagUpLoad,	
							[UpLoadMe] 		= @bFlagUpLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						--"PolicyLimits"
						UPDATE 	PolicyLimits SET 	
							[DownLoadMe] 		= @bFlagUpLoad,
							[UpLoadMe] 		= @bFlagUpLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						--"Package"
						UPDATE 	Package SET 	
							[DownLoadMe] 		= @bFlagUpLoad,
							[UpLoadMe] 		= @bFlagUpLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						--"PackageItem"
						UPDATE 	PackageItem SET 	
							[DownLoadMe] 		= @bFlagUpLoad,
							[UpLoadMe] 		= @bFlagUpLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						--"RTChecks"
						UPDATE 	RTChecks SET 	
							[DownLoadMe] 		= @bFlagUpLoad,
							[UpLoadMe] 		= @bFlagUpLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						--"RTIndemnity"
						UPDATE 	RTIndemnity SET 	
							[DownLoadMe] 		= @bFlagUpLoad,
							[UpLoadMe] 		= @bFlagUpLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						--"RTActivityLog"
						UPDATE 	RTActivityLog SET 	
							[DownLoadMe] 		= @bFlagUpLoad,
							[UpLoadMe] 		= @bFlagUpLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						--"RTActivityLogInfo"
						UPDATE 	RTActivityLogInfo SET 
							[DownLoadMe] 		= @bFlagUpLoad,	
							[UpLoadMe] 		= @bFlagUpLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						--"RTPhotoLogReport
						UPDATE RTPhotoReport SET
							[DownLoadMe] 		= @bFlagUpLoad,
							[UpLoadMe] 		= @bFlagUpLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						--"RTPhotoLog"
						UPDATE 	RTPhotoLog SET 	
							[DownLoadMe] 		= @bFlagUpLoad,
							[UpLoadPhoto]		= @bFlagUpLoad,
							[UpLoadPhotoThumb]	= @bFlagUpLoad,
							[UpLoadPhotoHighRes]	= @bFlagUpLoad,
							[UpLoadMe] 		= @bFlagUpLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						--RTWSDiagram
						UPDATE RTWSDiagram SET
							[DownLoadMe] 		= @bFlagUpLoad,
							[UploadDiagramPhoto]	= @bFlagUpLoad,
							[UpLoadMe] 		= @bFlagUpLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						--"RTAttachments"
						UPDATE 	RTAttachments SET 	
							[DownLoadMe] 		= @bFlagUpLoad,
							[UpLoadAttachment]	= @bFlagUpLoad,
							[UpLoadMe] 		= @bFlagUpLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						--MiscReportParam
						UPDATE 	MiscReportParam SET 	
							[DownLoadMe] 		= @bFlagUpLoad,
							[UpLoadMe] 		= @bFlagUpLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						--MiscReportParam
						UPDATE 	MiscReportParam01 SET 	
							[DownLoadMe] 		= @bFlagUpLoad,
							[UpLoadMe] 		= @bFlagUpLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						UPDATE 	MiscReportParam02 SET 	
							[DownLoadMe] 		= @bFlagUpLoad,
							[UpLoadMe] 		= @bFlagUpLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						UPDATE 	MiscReportParam03 SET 	
							[DownLoadMe] 		= @bFlagUpLoad,
							[UpLoadMe] 		= @bFlagUpLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						UPDATE 	MiscReportParam04 SET 	
							[DownLoadMe] 		= @bFlagUpLoad,
							[UpLoadMe] 		= @bFlagUpLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						UPDATE 	MiscReportParam05 SET 	
							[DownLoadMe] 		= @bFlagUpLoad,
							[UpLoadMe] 		= @bFlagUpLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						UPDATE 	MiscReportParam06 SET 	
							[DownLoadMe] 		= @bFlagUpLoad,
							[UpLoadMe] 		= @bFlagUpLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						UPDATE 	MiscReportParam07 SET 	
							[DownLoadMe] 		= @bFlagUpLoad,
							[UpLoadMe] 		= @bFlagUpLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						UPDATE 	MiscReportParam08 SET 	
							[DownLoadMe] 		= @bFlagUpLoad,
							[UpLoadMe] 		= @bFlagUpLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						UPDATE 	MiscReportParam09 SET 	
							[DownLoadMe] 		= @bFlagUpLoad,
							[UpLoadMe] 		= @bFlagUpLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						UPDATE 	MiscReportParam10 SET 	
							[DownLoadMe] 		= @bFlagUpLoad,
							[UpLoadMe] 		= @bFlagUpLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						UPDATE 	MiscReportParam11 SET 	
							[DownLoadMe] 		= @bFlagUpLoad,
							[UpLoadMe] 		= @bFlagUpLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						UPDATE 	MiscReportParam12 SET 	
							[DownLoadMe] 		= @bFlagUpLoad,
							[UpLoadMe] 		= @bFlagUpLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						UPDATE 	MiscReportParam13 SET 	
							[DownLoadMe] 		= @bFlagUpLoad,
							[UpLoadMe] 		= @bFlagUpLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						UPDATE 	MiscReportParam14 SET 	
							[DownLoadMe] 		= @bFlagUpLoad,
							[UpLoadMe] 		= @bFlagUpLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						UPDATE 	MiscReportParam15 SET 	
							[DownLoadMe] 		= @bFlagUpLoad,
							[UpLoadMe] 		= @bFlagUpLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						UPDATE 	MiscReportParam16 SET 	
							[DownLoadMe] 		= @bFlagUpLoad,
							[UpLoadMe] 		= @bFlagUpLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						UPDATE 	MiscReportParam17 SET 	
							[DownLoadMe] 		= @bFlagUpLoad,
							[UpLoadMe] 		= @bFlagUpLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						UPDATE 	MiscReportParam18 SET 	
							[DownLoadMe] 		= @bFlagUpLoad,
							[UpLoadMe] 		= @bFlagUpLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						UPDATE 	MiscReportParam19 SET 	
							[DownLoadMe] 		= @bFlagUpLoad,
							[UpLoadMe] 		= @bFlagUpLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						UPDATE 	MiscReportParam20 SET 	
							[DownLoadMe] 		= @bFlagUpLoad,
							[UpLoadMe] 		= @bFlagUpLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						UPDATE 	MiscReportParam21 SET 	
							[DownLoadMe] 		= @bFlagUpLoad,
							[UpLoadMe] 		= @bFlagUpLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						UPDATE 	MiscReportParam22 SET 	
							[DownLoadMe] 		= @bFlagUpLoad,
							[UpLoadMe] 		= @bFlagUpLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						UPDATE 	MiscReportParam23 SET 	
							[DownLoadMe] 		= @bFlagUpLoad,
							[UpLoadMe] 		= @bFlagUpLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						UPDATE 	MiscReportParam24 SET 	
							[DownLoadMe] 		= @bFlagUpLoad,
							[UpLoadMe] 		= @bFlagUpLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						UPDATE 	MiscReportParam25 SET 	
							[DownLoadMe] 		= @bFlagUpLoad,
							[UpLoadMe] 		= @bFlagUpLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						UPDATE 	MiscReportParam26 SET 	
							[DownLoadMe] 		= @bFlagUpLoad,
							[UpLoadMe] 		= @bFlagUpLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						UPDATE 	MiscReportParam27 SET 	
							[DownLoadMe] 		= @bFlagUpLoad,
							[UpLoadMe] 		= @bFlagUpLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						UPDATE 	MiscReportParam28 SET 	
							[DownLoadMe] 		= @bFlagUpLoad,
							[UpLoadMe] 		= @bFlagUpLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						UPDATE 	MiscReportParam29 SET 	
							[DownLoadMe] 		= @bFlagUpLoad,
							[UpLoadMe] 		= @bFlagUpLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
						UPDATE 	MiscReportParam30 SET 	
							[DownLoadMe] 		= @bFlagUpLoad,
							[UpLoadMe] 		= @bFlagUpLoad,
							[DateLastUpdated] 	= @DateLastUpdated ,
							[UpdateByUserID] 	= @UID 
						WHERE 	AssignmentsID 		= @AssignmentsID
					END

				END
			END
		END
	END
END
ELSE
BEGIN
	IF @GroupBy Is Not Null
	BEGIN
		SET @sSQL1 = 'SELECT RetAssignments.' + @GroupBy + ' '
	END
	ELSE
	BEGIN
		SET @sSQL1 = 'SELECT RetAssignments.* '
	END	
	SET @sSQL1 = @sSQL1 + 'FROM( ' 
	SET @sSQL1 = @sSQL1 + 'SELECT '
	IF @bViewHistory Is Not Null
	BEGIN
		SET @sSQL1 = @sSQL1 + '[AssignmentsHistoryID] , '
	END
	SET @sSQL1 = @sSQL1 + '[AssignmentsID] , '
	SET @sSQL1 = @sSQL1 + '[ID], '

	IF @bShowAssignmentTypeInfo = 1
	BEGIN
		SET @sSQL1 = @sSQL1 + '(SELECT	Type '
		SET @sSQL1 = @sSQL1 + 'FROM 	AssignmentType '
		SET @sSQL1 = @sSQL1 + 'WHERE 	AssignmentTypeID = A.[AssignmentTypeID]) As AssignmentTypeType,  '
		SET @sSQL1 = @sSQL1 + '(SELECT	Description '
		SET @sSQL1 = @sSQL1 + 'FROM 	AssignmentType '
		SET @sSQL1 = @sSQL1 + 'WHERE 	AssignmentTypeID = A.[AssignmentTypeID]) As AssignmentTypeDescription,  '
		SET @sSQL1 = @sSQL1 + '(SELECT	AdminComments '
		SET @sSQL1 = @sSQL1 + 'FROM 	AssignmentType '
		SET @sSQL1 = @sSQL1 + 'WHERE 	AssignmentTypeID = A.[AssignmentTypeID]) As AssignmentTypeAdminComments,  '
		SET @sSQL1 = @sSQL1 + '(SELECT	IsDeleted '
		SET @sSQL1 = @sSQL1 + 'FROM 	AssignmentType '
		SET @sSQL1 = @sSQL1 + 'WHERE 	AssignmentTypeID = A.[AssignmentTypeID]) As AssignmentTypeIsDeleted,  '
	END

	SET @sSQL1 = @sSQL1 + '[AssignmentTypeID], '
	
	IF @bShowCatInfo =1
	BEGIN
		SET @sTempSQL ='(	SELECT 	CATID
					FROM 	ClientCompanyCatSpec 
					WHERE	ClientCompanyCatSpecID = A.[ClientCompanyCatSpecID]
				)
				'
		SET @sSQL1 = @sSQL1 + '(SELECT	CATID '
		SET @sSQL1 = @sSQL1 + 'FROM 	CAT '
		SET @sSQL1 = @sSQL1 + 'WHERE 	CATID = ' + @sTempSQL + ')As CatID,  '

		SET @sSQL1 = @sSQL1 + '(SELECT	Name '
		SET @sSQL1 = @sSQL1 + 'FROM 	CAT '
		SET @sSQL1 = @sSQL1 + 'WHERE 	CATID = ' + @sTempSQL + ')As CatName,  '
		
		SET @sSQL1 = @sSQL1 + '(SELECT	Description '
		SET @sSQL1 = @sSQL1 + 'FROM 	CAT '
		SET @sSQL1 = @sSQL1 + 'WHERE 	CATID = ' + @sTempSQL + ')As CatDescription,  '

		SET @sSQL1 = @sSQL1 + '(SELECT	ActiveDate '
		SET @sSQL1 = @sSQL1 + 'FROM 	CAT '
		SET @sSQL1 = @sSQL1 + 'WHERE 	CATID = ' + @sTempSQL + ')As CatActiveDate,  '

		SET @sSQL1 = @sSQL1 + '(SELECT	InactiveDate '
		SET @sSQL1 = @sSQL1 + 'FROM 	CAT '
		SET @sSQL1 = @sSQL1 + 'WHERE 	CATID = ' + @sTempSQL + ')As CatInactiveDate,  '

	END

	IF @bShowCompanyInfo = 1
 	BEGIN
		SET @sTempSQL ='(	SELECT	IsClientOF
					FROM	Company
					WHERE	CompanyID =	(	SELECT 	ClientCompanyID
									FROM 	ClientCompanyCatSpec 
									WHERE	ClientCompanyCatSpecID = A.[ClientCompanyCatSpecID]
								)
				)
				'
		SET @sSQL1 = @sSQL1 + '(SELECT	CompanyID '
		SET @sSQL1 = @sSQL1 + 'FROM 	Company '
		SET @sSQL1 = @sSQL1 + 'WHERE 	CompanyID = ' + @sTempSQL + ')As CompanyID,  '

 		SET @sSQL1 = @sSQL1 + '(SELECT	Name '
		SET @sSQL1 = @sSQL1 + 'FROM 	Company '
		SET @sSQL1 = @sSQL1 + 'WHERE 	CompanyID = ' + @sTempSQL + ')As CompanyName,  '
		
		
		SET @sSQL1 = @sSQL1 + '(SELECT	DBName '
		SET @sSQL1 = @sSQL1 + 'FROM 	Company '
		SET @sSQL1 = @sSQL1 + 'WHERE 	CompanyID =' + @sTempSQL + ')As CompanyDBName,  '

		SET @sSQL1 = @sSQL1 + '(SELECT	Code '
		SET @sSQL1 = @sSQL1 + 'FROM 	Company '
		SET @sSQL1 = @sSQL1 + 'WHERE 	CompanyID =' + @sTempSQL + ')As CompanyCode,  '

		SET @sSQL1 = @sSQL1 + '(SELECT	CarrierPrefix '
		SET @sSQL1 = @sSQL1 + 'FROM 	Company '
		SET @sSQL1 = @sSQL1 + 'WHERE 	CompanyID =' + @sTempSQL + ')As CompanyCarrierPrefix,  '

		SET @sSQL1 = @sSQL1 + '(SELECT	Comments '
		SET @sSQL1 = @sSQL1 + 'FROM 	Company '
		SET @sSQL1 = @sSQL1 + 'WHERE 	CompanyID =' + @sTempSQL + ')As CompanyComments,  '

		SET @sSQL1 = @sSQL1 + '(SELECT	EnableSingleFile '
		SET @sSQL1 = @sSQL1 + 'FROM 	Company '
		SET @sSQL1 = @sSQL1 + 'WHERE 	CompanyID =' + @sTempSQL + ')As CompanyEnableSingleFile,  '

		SET @sSQL1 = @sSQL1 + '(SELECT	SingleFileEmail '
		SET @sSQL1 = @sSQL1 + 'FROM 	Company '
		SET @sSQL1 = @sSQL1 + 'WHERE 	CompanyID =' + @sTempSQL + ')As CompanySingleFileEmail,  '

		SET @sSQL1 = @sSQL1 + '(SELECT	Active '
		SET @sSQL1 = @sSQL1 + 'FROM 	Company '
		SET @sSQL1 = @sSQL1 + 'WHERE 	CompanyID =' + @sTempSQL + ')As CompanyActive,  '

		SET @sSQL1 = @sSQL1 + '(SELECT	LogoImageName '
		SET @sSQL1 = @sSQL1 + 'FROM 	Company '
		SET @sSQL1 = @sSQL1 + 'WHERE 	CompanyID =' + @sTempSQL + ')As CompanyLogoImageName,  '

 	END

	IF @bShowClientCompanyInfo =1
	BEGIN
		SET @sTempSQL ='(	SELECT 	ClientCompanyID
					FROM 	ClientCompanyCatSpec 
					WHERE	ClientCompanyCatSpecID = A.[ClientCompanyCatSpecID]
				)
				'
		SET @sSQL1 = @sSQL1 + '(SELECT	CompanyID '
		SET @sSQL1 = @sSQL1 + 'FROM 	Company '
		SET @sSQL1 = @sSQL1 + 'WHERE 	CompanyID =' + @sTempSQL + ')As ClientCompanyID,  '

		SET @sSQL1 = @sSQL1 + '(SELECT	Name '
		SET @sSQL1 = @sSQL1 + 'FROM 	Company '
		SET @sSQL1 = @sSQL1 + 'WHERE 	CompanyID =' + @sTempSQL + ')As ClientCompanyName,  '

		SET @sSQL1 = @sSQL1 + '(SELECT	DBName '
		SET @sSQL1 = @sSQL1 + 'FROM 	Company '
		SET @sSQL1 = @sSQL1 + 'WHERE 	CompanyID =' + @sTempSQL + ')As ClientCompanyDBName,  '

		SET @sSQL1 = @sSQL1 + '(SELECT	Code '
		SET @sSQL1 = @sSQL1 + 'FROM 	Company '
		SET @sSQL1 = @sSQL1 + 'WHERE 	CompanyID =' + @sTempSQL + ')As ClientCompanyCode,  '

		SET @sSQL1 = @sSQL1 + '(SELECT	CarrierPrefix '
		SET @sSQL1 = @sSQL1 + 'FROM 	Company '
		SET @sSQL1 = @sSQL1 + 'WHERE 	CompanyID =' + @sTempSQL + ')As ClientCompanyCarrierPrefix,  '

		SET @sSQL1 = @sSQL1 + '(SELECT	Comments '
		SET @sSQL1 = @sSQL1 + 'FROM 	Company '
		SET @sSQL1 = @sSQL1 + 'WHERE 	CompanyID =' + @sTempSQL + ')As ClientCompanyComments,  ' 

		SET @sSQL1 = @sSQL1 + '(SELECT	EnableSingleFile '
		SET @sSQL1 = @sSQL1 + 'FROM 	Company '
		SET @sSQL1 = @sSQL1 + 'WHERE 	CompanyID =' + @sTempSQL + ')As ClientCompanyEnableSingleFile,  '

		SET @sSQL1 = @sSQL1 + '(SELECT	SingleFileEmail '
		SET @sSQL1 = @sSQL1 + 'FROM 	Company '
		SET @sSQL1 = @sSQL1 + 'WHERE 	CompanyID =' + @sTempSQL + ')As ClientCompanySingleFileEmail,  '

		SET @sSQL1 = @sSQL1 + '(SELECT	Active '
		SET @sSQL1 = @sSQL1 + 'FROM 	Company '
		SET @sSQL1 = @sSQL1 + 'WHERE 	CompanyID =' + @sTempSQL + ')As ClientCompanyActive,  '

		SET @sSQL1 = @sSQL1 + '(SELECT	LogoImageName '
		SET @sSQL1 = @sSQL1 + 'FROM 	Company '
		SET @sSQL1 = @sSQL1 + 'WHERE 	CompanyID =' + @sTempSQL + ')As ClientCompanyLogoImageName,  '

	END
	SET @sSQL2 = ''
	IF @bShowFeeScheduleInfo =1
	BEGIN
		SET @sTempSQL ='(	SELECT 	FeeScheduleID
					FROM 	ClientCompanyCat 
					WHERE	ClientCompanyID = 	(	SELECT 	ClientCompanyID
										FROM 	ClientCompanyCatSpec 
										WHERE	ClientCompanyCatSpecID = A.[ClientCompanyCatSpecID]
									)
					AND	CATID =			(	SELECT 	CATID 
										FROM 	ClientCompanyCatSpec 
										WHERE	ClientCompanyCatSpecID = A.[ClientCompanyCatSpecID]
									)
				) 
				'
		SET @sSQL2 = @sSQL2 + '(SELECT	FeeScheduleID '
		SET @sSQL2 = @sSQL2 + 'FROM 	FeeSchedule '
		SET @sSQL2 = @sSQL2 + 'WHERE 	FeeScheduleID =' + @sTempSQL + ')As FeeScheduleFeeScheduleID,  '		

		SET @sSQL2 = @sSQL2 + '(SELECT	ScheduleName '
		SET @sSQL2 = @sSQL2 + 'FROM 	FeeSchedule '
		SET @sSQL2 = @sSQL2 + 'WHERE 	FeeScheduleID =' + @sTempSQL + ')As FeeScheduleScheduleName,  '

		SET @sSQL2 = @sSQL2 + '(SELECT	Description '
		SET @sSQL2 = @sSQL2 + 'FROM 	FeeSchedule '
		SET @sSQL2 = @sSQL2 + 'WHERE 	FeeScheduleID =' + @sTempSQL + ')As FeeScheduleDescription,  '

	END

	IF @bShowClientCompanyCatInfo = 1
	BEGIN
		SET @sTempSQL ='(	SELECT 	ClientCompanyID
					FROM 	ClientCompanyCatSpec 
					WHERE	ClientCompanyCatSpecID = A.[ClientCompanyCatSpecID]
				) 
				'
		SET @sTempSQL2 ='(	SELECT 	CATID 
					FROM 	ClientCompanyCatSpec 
					WHERE	ClientCompanyCatSpecID = A.[ClientCompanyCatSpecID]
				)
				'
		SET @sSQL2 = @sSQL2 + '(SELECT	SiteAddress '
		SET @sSQL2 = @sSQL2 + 'FROM 	ClientCompanyCat '
		SET @sSQL2 = @sSQL2 + 'WHERE 	ClientCompanyID =' + @sTempSQL + ' '
		SET @sSQL2 = @sSQL2 + 'AND 	CATID =' + @sTempSQL2 + ')As ClientCompanyCatSiteAddress,  '

		SET @sSQL2 = @sSQL2 + '(SELECT	SACity '
		SET @sSQL2 = @sSQL2 + 'FROM 	ClientCompanyCat '
		SET @sSQL2 = @sSQL2 + 'WHERE 	ClientCompanyID =' + @sTempSQL + ' '
		SET @sSQL2 = @sSQL2 + 'AND 	CATID =' + @sTempSQL2 + ')As ClientCompanyCatSACity,  '

		SET @sSQL2 = @sSQL2 + '(SELECT	SAState '
		SET @sSQL2 = @sSQL2 + 'FROM 	ClientCompanyCat '
		SET @sSQL2 = @sSQL2 + 'WHERE 	ClientCompanyID =' + @sTempSQL + ' '
		SET @sSQL2 = @sSQL2 + 'AND 	CATID =' + @sTempSQL2 + ')As ClientCompanyCatSAState,  '

		SET @sSQL2 = @sSQL2 + '(SELECT	SAZip '
		SET @sSQL2 = @sSQL2 + 'FROM 	ClientCompanyCat '
		SET @sSQL2 = @sSQL2 + 'WHERE 	ClientCompanyID =' + @sTempSQL + ' '
		SET @sSQL2 = @sSQL2 + 'AND 	CATID =' + @sTempSQL2 + ')As ClientCompanyCatSAZip,  '

		SET @sSQL2 = @sSQL2 + '(SELECT	SAZip4 '
		SET @sSQL2 = @sSQL2 + 'FROM 	ClientCompanyCat '
		SET @sSQL2 = @sSQL2 + 'WHERE 	ClientCompanyID =' + @sTempSQL + ' '
		SET @sSQL2 = @sSQL2 + 'AND 	CATID =' + @sTempSQL2 + ')As ClientCompanyCatSAZip4,  '

		SET @sSQL2 = @sSQL2 + '(SELECT	SAOtherPostCode '
		SET @sSQL2 = @sSQL2 + 'FROM 	ClientCompanyCat '
		SET @sSQL2 = @sSQL2 + 'WHERE 	ClientCompanyID =' + @sTempSQL + ' '
		SET @sSQL2 = @sSQL2 + 'AND 	CATID =' + @sTempSQL2 + ')As ClientCompanyCatSAOtherPostCode,  '

		SET @sSQL2 = @sSQL2 + '(SELECT	ActiveDate '
		SET @sSQL2 = @sSQL2 + 'FROM 	ClientCompanyCat '
		SET @sSQL2 = @sSQL2 + 'WHERE 	ClientCompanyID =' + @sTempSQL + ' '
		SET @sSQL2 = @sSQL2 + 'AND 	CATID =' + @sTempSQL2 + ')As ClientCompanyCatActiveDate,  '

		SET @sSQL2 = @sSQL2 + '(SELECT	InactiveDate '
		SET @sSQL2 = @sSQL2 + 'FROM 	ClientCompanyCat '
		SET @sSQL2 = @sSQL2 + 'WHERE 	ClientCompanyID =' + @sTempSQL + ' '
		SET @sSQL2 = @sSQL2 + 'AND 	CATID =' + @sTempSQL2 + ')As ClientCompanyCatInactiveDate,  '

	END
	SET @sSQL3 = ''
	IF @bShowClientCompanyCatSpecInfo =1
	BEGIN
		SET @sSQL3 = @sSQL3 + '(SELECT	ClientCompanyID '
		SET @sSQL3 = @sSQL3 + 'FROM 	ClientCompanyCatSpec '
		SET @sSQL3 = @sSQL3 + 'WHERE 	ClientCompanyCatSpecID = A.[ClientCompanyCatSpecID]) As ClientCompanyCatSpecClientCompanyID,  '
		SET @sSQL3 = @sSQL3 + '(SELECT	CATID '
		SET @sSQL3 = @sSQL3 + 'FROM 	ClientCompanyCatSpec '
		SET @sSQL3 = @sSQL3 + 'WHERE 	ClientCompanyCatSpecID = A.[ClientCompanyCatSpecID]) As ClientCompanyCatSpecCATID,  '
		SET @sSQL3 = @sSQL3 + '(SELECT	CatCode '
		SET @sSQL3 = @sSQL3 + 'FROM 	ClientCompanyCatSpec '
		SET @sSQL3 = @sSQL3 + 'WHERE 	ClientCompanyCatSpecID = A.[ClientCompanyCatSpecID]) As ClientCompanyCatSpecCatCode,  '
		SET @sSQL3 = @sSQL3 + '(SELECT	Comments '
		SET @sSQL3 = @sSQL3 + 'FROM 	ClientCompanyCatSpec '
		SET @sSQL3 = @sSQL3 + 'WHERE 	ClientCompanyCatSpecID = A.[ClientCompanyCatSpecID]) As ClientCompanyCatSpecComments,  '
		SET @sSQL3 = @sSQL3 + '(SELECT	ActiveDate '
		SET @sSQL3 = @sSQL3 + 'FROM 	ClientCompanyCatSpec '
		SET @sSQL3 = @sSQL3 + 'WHERE 	ClientCompanyCatSpecID = A.[ClientCompanyCatSpecID]) As ClientCompanyCatSpecActiveDate,  '
		SET @sSQL3 = @sSQL3 + '(SELECT	InactiveDate '
		SET @sSQL3 = @sSQL3 + 'FROM 	ClientCompanyCatSpec '
		SET @sSQL3 = @sSQL3 + 'WHERE 	ClientCompanyCatSpecID = A.[ClientCompanyCatSpecID]) As ClientCompanyCatSpecInactiveDate,  '
	END

	SET @sSQL3 = @sSQL3 + '[ClientCompanyCatSpecID], '

	IF @bShowAdjusterSpecUsersInfo = 1
	BEGIN
		SET @sTempSQL ='(	SELECT UsersID 
					FROM ClientCoAdjusterSpec 
					WHERE	ClientCoAdjusterSpecID = A.[AdjusterSpecID]
				)
				'
		SET @sSQL3 = @sSQL3 + '(SELECT	USERSID '
		SET @sSQL3 = @sSQL3 + 'FROM 	USERS '
		SET @sSQL3 = @sSQL3 + 'WHERE 	USERSID =' + @sTempSQL + ')As AdjusterSpecUsersID,  '
		
		SET @sSQL3 = @sSQL3 + '(SELECT	USERNAME '
		SET @sSQL3 = @sSQL3 + 'FROM 	USERS '
		SET @sSQL3 = @sSQL3 + 'WHERE 	USERSID =' + @sTempSQL + ')As AdjusterSpecUserName,  '

		SET @sSQL3 = @sSQL3 + '(SELECT	FirstName '
		SET @sSQL3 = @sSQL3 + 'FROM 	USERS '
		SET @sSQL3 = @sSQL3 + 'WHERE 	USERSID =' + @sTempSQL + ')As AdjusterSpecFirstName,  '

		SET @sSQL3 = @sSQL3 + '(SELECT	LastName '
		SET @sSQL3 = @sSQL3 + 'FROM 	USERS '
		SET @sSQL3 = @sSQL3 + 'WHERE 	USERSID =' + @sTempSQL + ')As AdjusterSpecLastName,  '

		SET @sSQL3 = @sSQL3 + '(SELECT	Email '
		SET @sSQL3 = @sSQL3 + 'FROM 	USERS '
		SET @sSQL3 = @sSQL3 + 'WHERE 	USERSID =' + @sTempSQL + ')As AdjusterSpecEmail,  '

		SET @sSQL3 = @sSQL3 + '(SELECT	ContactPhone '
		SET @sSQL3 = @sSQL3 + 'FROM 	USERS '
		SET @sSQL3 = @sSQL3 + 'WHERE 	USERSID =' + @sTempSQL + ')As AdjusterSpecContactPhone,  '

		SET @sSQL3 = @sSQL3 + '(SELECT	EmergencyPhone '
		SET @sSQL3 = @sSQL3 + 'FROM 	USERS '
		SET @sSQL3 = @sSQL3 + 'WHERE 	USERSID =' + @sTempSQL + ')As AdjusterSpecEmergencyPhone,  '

		SET @sSQL3 = @sSQL3 + '(SELECT	Active '
		SET @sSQL3 = @sSQL3 + 'FROM 	USERS '
		SET @sSQL3 = @sSQL3 + 'WHERE 	USERSID =' + @sTempSQL + ')As AdjusterSpecActive,  '

	END
	SET @sSQL4 = ''
	IF @bShowAdjusterSpecInfo = 1 
	BEGIN	
		SET @sSQL4 = @sSQL4 + '(SELECT	(	CASE 	WHEN 	ZipCode Is Null '
		SET @sSQL4 = @sSQL4 + '				THEN 	0 '
		SET @sSQL4 = @sSQL4 + '				ELSE	1 '
		SET @sSQL4 = @sSQL4 + '				END '
		SET @sSQL4 = @sSQL4 + '		) '
		SET @sSQL4 = @sSQL4 + 'FROM 	ClientCoAdjusterSpec '
		SET @sSQL4 = @sSQL4 + 'WHERE 	ClientCoAdjusterSpecID = A.[AdjusterSpecID]) As AdjusterSpecAssignByZipCode,  '
		SET @sSQL4 = @sSQL4 + '(SELECT	ACID '
		SET @sSQL4 = @sSQL4 + 'FROM 	ClientCoAdjusterSpec '
		SET @sSQL4 = @sSQL4 + 'WHERE 	ClientCoAdjusterSpecID = A.[AdjusterSpecID]) As AdjusterSpecACID,  '
		SET @sSQL4 = @sSQL4 + '(SELECT	ACIDDescription '
		SET @sSQL4 = @sSQL4 + 'FROM 	ClientCoAdjusterSpec '
		SET @sSQL4 = @sSQL4 + 'WHERE 	ClientCoAdjusterSpecID = A.[AdjusterSpecID]) As AdjusterSpecACIDDescription,  '
		SET @sSQL4 = @sSQL4 + '(SELECT	Comments '
		SET @sSQL4 = @sSQL4 + 'FROM 	ClientCoAdjusterSpec '
		SET @sSQL4 = @sSQL4 + 'WHERE 	ClientCoAdjusterSpecID = A.[AdjusterSpecID]) As AdjusterSpecComments,  '
		SET @sSQL4 = @sSQL4 + '(SELECT	ActiveDate '
		SET @sSQL4 = @sSQL4 + 'FROM 	ClientCoAdjusterSpec '
		SET @sSQL4 = @sSQL4 + 'WHERE 	ClientCoAdjusterSpecID = A.[AdjusterSpecID]) As AdjusterSpecActiveDate,  '
		SET @sSQL4 = @sSQL4 + '(SELECT	InactiveDate '
		SET @sSQL4 = @sSQL4 + 'FROM 	ClientCoAdjusterSpec '
		SET @sSQL4 = @sSQL4 + 'WHERE 	ClientCoAdjusterSpecID = A.[AdjusterSpecID]) As AdjusterSpecInactiveDate,  '
	END

	SET @sSQL4 = @sSQL4 + '[AdjusterSpecID] , '
	SET @sSQL4 = @sSQL4 + '(SELECT	ACID '
	SET @sSQL4 = @sSQL4 + 'FROM 	ClientCoAdjusterSpec '
	SET @sSQL4 = @sSQL4 + 'WHERE 	ClientCoAdjusterSpecID = A.[AdjusterSpecIDDisplay]) As AdjusterSpecACIDDisplay,  '
	SET @sSQL4 = @sSQL4 + '(SELECT	ACIDDescription '
	SET @sSQL4 = @sSQL4 + 'FROM 	ClientCoAdjusterSpec '
	SET @sSQL4 = @sSQL4 + 'WHERE 	ClientCoAdjusterSpecID = A.[AdjusterSpecIDDisplay]) As AdjusterSpecACIDDisplayDescription,  '
	SET @sSQL4 = @sSQL4 + '(SELECT	Comments '
	SET @sSQL4 = @sSQL4 + 'FROM 	ClientCoAdjusterSpec '
	SET @sSQL4 = @sSQL4 + 'WHERE 	ClientCoAdjusterSpecID = A.[AdjusterSpecIDDisplay]) As AdjusterSpecDisplayComments,  '
	SET @sSQL4 = @sSQL4 + '(SELECT	ActiveDate '
	SET @sSQL4 = @sSQL4 + 'FROM 	ClientCoAdjusterSpec '
	SET @sSQL4 = @sSQL4 + 'WHERE 	ClientCoAdjusterSpecID = A.[AdjusterSpecIDDisplay]) As AdjusterSpecDisplayActiveDate,  '
	SET @sSQL4 = @sSQL4 + '(SELECT	InactiveDate '
	SET @sSQL4 = @sSQL4 + 'FROM 	ClientCoAdjusterSpec '
	SET @sSQL4 = @sSQL4 + 'WHERE 	ClientCoAdjusterSpecID = A.[AdjusterSpecIDDisplay]) As AdjusterSpecDisplayInactiveDate,  '
	SET @sSQL4 = @sSQL4 + '[AdjusterSpecIDDisplay] , '
	SET @sSQL5 = ''
	IF @bShowAssignmentsInfo = 1
	BEGIN
		SET @sSQL5 = @sSQL5 + '[SPVersion], '
		SET @sSQL5 = @sSQL5 + '[IBNUM], '
		SET @sSQL5 = @sSQL5 + '[CLIENTNUM], '
		SET @sSQL5 = @sSQL5 + '[PolicyNo], '
		SET @sSQL5 = @sSQL5 + '[PolicyDescription], '
		SET @sSQL5 = @sSQL5 + '[Insured], '
		SET @sSQL5 = @sSQL5 + '[MailingAddress], '
		SET @sSQL5 = @sSQL5 + '[MAStreet], '
		SET @sSQL5 = @sSQL5 + '[MACity], '
		SET @sSQL5 = @sSQL5 + '[MAState], '
		SET @sSQL5 = @sSQL5 + '[MAZIP], '
		SET @sSQL5 = @sSQL5 + '[MAZIP4], '
		SET @sSQL5 = @sSQL5 + '[MAOtherPostCode], '
		SET @sSQL5 = @sSQL5 + '[HomePhone] , '
		SET @sSQL5 = @sSQL5 + '[BusinessPhone], '
		SET @sSQL5 = @sSQL5 + '[PropertyAddress], '
		SET @sSQL5 = @sSQL5 + '[PAStreet] , '
		SET @sSQL5 = @sSQL5 + '[PACity] , '
		SET @sSQL5 = @sSQL5 + '[PAState], '
		SET @sSQL5 = @sSQL5 + '[PAZIP] , '
		SET @sSQL5 = @sSQL5 + '[PAZIP4], '
		SET @sSQL5 = @sSQL5 + '[PAOtherPostCode] , '
		SET @sSQL5 = @sSQL5 + '[MortgageeName] , '
		SET @sSQL5 = @sSQL5 + '[AgentNo] , '
		SET @sSQL5 = @sSQL5 + '[ReportedBy], '
		SET @sSQL5 = @sSQL5 + '[ReportedByPhone], '
		SET @sSQL5 = @sSQL5 + '[Deductible] , '
		SET @sSQL5 = @sSQL5 + '[AppDedClassTypeIDOrder], '
		SET @sSQL5 = @sSQL5 + '[LRFormat] , '
		SET @sSQL5 = @sSQL5 + '(CASE WHEN [LRFormat] = ''V2ECcarFarmers.clsLossXML01'' THEN '''' ELSE [LossReport] END) As [LossReport], '
		SET @sSQL5 = @sSQL5 + '[DownLoadLossReport], '
		SET @sSQL5 = @sSQL5 + '[UpLoadLossReport], '
	END

	IF @bShowStatusInfo = 1 
	BEGIN
		SET @sSQL5 = @sSQL5 + '(SELECT	StatusAlias '
		SET @sSQL5 = @sSQL5 + 'FROM 	Status '
		SET @sSQL5 = @sSQL5 + 'WHERE 	StatusID = A.[StatusID]) As StatusStatusAlias,  '
		SET @sSQL5 = @sSQL5 + '(SELECT	Status '
		SET @sSQL5 = @sSQL5 + 'FROM 	Status '
		SET @sSQL5 = @sSQL5 + 'WHERE 	StatusID = A.[StatusID]) As StatusStatus,  '
		SET @sSQL5 = @sSQL5 + '(SELECT	Description '
		SET @sSQL5 = @sSQL5 + 'FROM 	Status '
		SET @sSQL5 = @sSQL5 + 'WHERE 	StatusID = A.[StatusID]) As StatusDescription,  '
		SET @sSQL5 = @sSQL5 + '(SELECT	AdminComments '
		SET @sSQL5 = @sSQL5 + 'FROM 	Status '
		SET @sSQL5 = @sSQL5 + 'WHERE 	StatusID = A.[StatusID]) As StatusAdminComments,  '
		SET @sSQL5 = @sSQL5 + '(SELECT	IsDeleted '
		SET @sSQL5 = @sSQL5 + 'FROM 	Status '
		SET @sSQL5 = @sSQL5 + 'WHERE 	StatusID = A.[StatusID]) As StatusIsDeleted,  '

		-- 9.23.2005 Also Include the Count Of Rejected PackageItems + Count of Items marked to Send
		SET @sSQL5 = @sSQL5 + '(SELECT	Count(PKGI.[IsClientCoReject])  As CountClientReject '
		SET @sSQL5 = @sSQL5 + 'FROM PackageItem PKGI '
		SET @sSQL5 = @sSQL5 + 'WHERE PKGI.[AssignmentsID] = A.[AssignmentsID] '
		SET @sSQL5 = @sSQL5 + 'AND PKGI.[IsDeleted] = 0 AND PKGI.[IsClientCoReject] = 1) As PKGIClientRejected, '
		
		SET @sSQL5 = @sSQL5 + '(SELECT	Count(PKGI.[IsClientCoDelete])  As CountClientDelete '
		SET @sSQL5 = @sSQL5 + 'FROM PackageItem PKGI '
		SET @sSQL5 = @sSQL5 + 'WHERE PKGI.[AssignmentsID] = A.[AssignmentsID] '
		SET @sSQL5 = @sSQL5 + 'AND PKGI.[IsDeleted] = 0 AND PKGI.[IsClientCoDelete] = 1) As PKGIClientDelete, '

		SET @sSQL5 = @sSQL5 + '(SELECT	Count(PKGI.[IsClientCoApprove])  As CountClientApprove '
		SET @sSQL5 = @sSQL5 + 'FROM PackageItem PKGI '
		SET @sSQL5 = @sSQL5 + 'WHERE PKGI.[AssignmentsID] = A.[AssignmentsID] '
		SET @sSQL5 = @sSQL5 + 'AND PKGI.[IsDeleted] = 0 AND PKGI.[IsClientCoApprove] = 1) As PKGIClientApprove, '

		SET @sSQL5 = @sSQL5 + '(SELECT	Count(PKGI.[SendMe]) As CountSendMe '
		SET @sSQL5 = @sSQL5 + 'FROM PackageItem PKGI '
		SET @sSQL5 = @sSQL5 + 'WHERE PKGI.[AssignmentsID] = A.[AssignmentsID] '
		SET @sSQL5 = @sSQL5 + 'AND PKGI.[IsDeleted] = 0 AND PKGI.[SendMe] = 1) As PKGISendMe, '
	END

	SET @sSQL5 = @sSQL5 + '[StatusID] , '

	IF @bShowTypeOfLossInfo = 1 
	BEGIN
		SET @sSQL5 = @sSQL5 + '(SELECT	TypeOfLoss '
		SET @sSQL5 = @sSQL5 + 'FROM 	TypeOfLoss '
		SET @sSQL5 = @sSQL5 + 'WHERE 	TypeOfLossID = A.[TypeOfLossID]) As TypeOfLoss,  '
		SET @sSQL5 = @sSQL5 + '(SELECT	Description '
		SET @sSQL5 = @sSQL5 + 'FROM 	TypeOfLoss '
		SET @sSQL5 = @sSQL5 + 'WHERE 	TypeOfLossID = A.[TypeOfLossID]) As TypeOfLossDescription,  '
		SET @sSQL5 = @sSQL5 + '(SELECT	Code '
		SET @sSQL5 = @sSQL5 + 'FROM 	TypeOfLoss '
		SET @sSQL5 = @sSQL5 + 'WHERE 	TypeOfLossID = A.[TypeOfLossID]) As TypeOfLossCode,  '
		SET @sSQL5 = @sSQL5 + '(SELECT	IsDeleted '
		SET @sSQL5 = @sSQL5 + 'FROM 	TypeOfLoss '
		SET @sSQL5 = @sSQL5 + 'WHERE 	TypeOfLossID = A.[TypeOfLossID]) As TypeOfLossIsDeleted,  '
	END 

	SET @sSQL5 = @sSQL5 + '[TypeOfLossID], '

	IF @bShowBatchesInfo = 1
	BEGIN
-- 		--From Batches Table (To be entered into Accounting System)
-- 		SET @sSQL5 = @sSQL5 + '(SELECT Sum(GrossLoss) '
-- 		SET @sSQL5 = @sSQL5 + 'FROM Batches '
-- 		SET @sSQL5 = @sSQL5 + 'WHERE AssignmentsID = A.[AssignmentsID] And Void = 0 And Copied = 2) As BatchesGrossLoss, '
-- 		SET @sSQL5 = @sSQL5 + '(SELECT Sum(TotalFee) '
-- 		SET @sSQL5 = @sSQL5 + 'FROM Batches '
-- 		SET @sSQL5 = @sSQL5 + 'WHERE AssignmentsID = A.[AssignmentsID] And Void = 0 And Copied = 2) As BatchesTotalFee, '
		--From IB Table
		SET @sSQL5 = @sSQL5 + '(SELECT Sum(IB13_cGrossLoss) '
		SET @sSQL5 = @sSQL5 + 'FROM IB '
		SET @sSQL5 = @sSQL5 + 'WHERE AssignmentsID = A.[AssignmentsID] And Void = 0 ) As BatchesGrossLoss, '
		SET @sSQL5 = @sSQL5 + '(SELECT Sum(IB33_cTotalAdjustingFee) '
		SET @sSQL5 = @sSQL5 + 'FROM IB '
		SET @sSQL5 = @sSQL5 + 'WHERE AssignmentsID = A.[AssignmentsID] And Void = 0 ) As BatchesTotalFee, '

	END

	SET @sSQL6 = ''
	IF @bShowAssignmentsInfo = 1
	BEGIN
		SET @sSQL6 = @sSQL6 + '[XactTypeOfLoss], '
		SET @sSQL6 = @sSQL6 + '[SentToXact] , '
		SET @sSQL6 = @sSQL6 + '[LossDate] , '
		SET @sSQL6 = @sSQL6 + '[AssignedDate], '
		SET @sSQL6 = @sSQL6 + '[ReceivedDate], '
		SET @sSQL6 = @sSQL6 + '[ContactDate], '
		SET @sSQL6 = @sSQL6 + '[InspectedDate], '
		SET @sSQL6 = @sSQL6 + '[CloseDate] , '
		SET @sSQL6 = @sSQL6 + '[Reassigned] , '
		SET @sSQL6 = @sSQL6 + '[DateReassigned], '
	END

	IF @bShowRAAdjusterSpecUsersInfo =1 
	BEGIN
		SET @sTempSQL ='(	SELECT UsersID 
					FROM ClientCoAdjusterSpec 
					WHERE	ClientCoAdjusterSpecID = A.[RAAdjusterSpecID]
				)
				'
		SET @sSQL6 = @sSQL6 + '(SELECT	USERSID '
		SET @sSQL6 = @sSQL6 + 'FROM 	USERS '
		SET @sSQL6 = @sSQL6 + 'WHERE 	USERSID =' + @sTempSQL + ')As RAAdjusterSpecUsersID,  '

		SET @sSQL6 = @sSQL6 + '(SELECT	USERNAME '
		SET @sSQL6 = @sSQL6 + 'FROM 	USERS '
		SET @sSQL6 = @sSQL6 + 'WHERE 	USERSID =' + @sTempSQL + ')As RAAdjusterSpecUserName,  '

		SET @sSQL6 = @sSQL6 + '(SELECT	FirstName '
		SET @sSQL6 = @sSQL6 + 'FROM 	USERS '
		SET @sSQL6 = @sSQL6 + 'WHERE 	USERSID =' + @sTempSQL + ')As RAAdjusterSpecFirstName,  '

		SET @sSQL6 = @sSQL6 + '(SELECT	LastName '
		SET @sSQL6 = @sSQL6 + 'FROM 	USERS '
		SET @sSQL6 = @sSQL6 + 'WHERE 	USERSID = ' + @sTempSQL + ')As RAAdjusterSpecLastName,  '

		SET @sSQL6 = @sSQL6 + '(SELECT	Email '
		SET @sSQL6 = @sSQL6 + 'FROM 	USERS '
		SET @sSQL6 = @sSQL6 + 'WHERE 	USERSID =' + @sTempSQL + ')As RAAdjusterSpecEmail,  '

		SET @sSQL6 = @sSQL6 + '(SELECT	ContactPhone '
		SET @sSQL6 = @sSQL6 + 'FROM 	USERS '
		SET @sSQL6 = @sSQL6 + 'WHERE 	USERSID =' + @sTempSQL + ')As RAAdjusterSpecContactPhone,  '

		SET @sSQL6 = @sSQL6 + '(SELECT	EmergencyPhone '
		SET @sSQL6 = @sSQL6 + 'FROM 	USERS '
		SET @sSQL6 = @sSQL6 + 'WHERE 	USERSID =' + @sTempSQL + ')As RAAdjusterSpecEmergencyPhone,  '

		SET @sSQL6 = @sSQL6 + '(SELECT	Active '
		SET @sSQL6 = @sSQL6 + 'FROM 	USERS '
		SET @sSQL6 = @sSQL6 + 'WHERE 	USERSID =' + @sTempSQL + ')As RAAdjusterSpecActive,  '

	END

	IF @bShowRAAdjusterSpecInfo = 1
	BEGIN
		SET @sSQL6 = @sSQL6 + '(SELECT	ACID '
		SET @sSQL6 = @sSQL6 + 'FROM 	ClientCoAdjusterSpec '
		SET @sSQL6 = @sSQL6 + 'WHERE 	ClientCoAdjusterSpecID = A.[RAAdjusterSpecID]) As RAAdjusterSpecACID,  '
		SET @sSQL6 = @sSQL6 + '(SELECT	ACIDDescription '
		SET @sSQL6 = @sSQL6 + 'FROM 	ClientCoAdjusterSpec '
		SET @sSQL6 = @sSQL6 + 'WHERE 	ClientCoAdjusterSpecID = A.[RAAdjusterSpecID]) As RAAdjusterSpecACIDDescription,  '
		SET @sSQL6 = @sSQL6 + '(SELECT	Comments '
		SET @sSQL6 = @sSQL6 + 'FROM 	ClientCoAdjusterSpec '
		SET @sSQL6 = @sSQL6 + 'WHERE 	ClientCoAdjusterSpecID = A.[RAAdjusterSpecID]) As RAAdjusterSpecComments,  '
		SET @sSQL6 = @sSQL6 + '(SELECT	ActiveDate '
		SET @sSQL6 = @sSQL6 + 'FROM 	ClientCoAdjusterSpec '
		SET @sSQL6 = @sSQL6 + 'WHERE 	ClientCoAdjusterSpecID = A.[RAAdjusterSpecID]) As RAAdjusterSpecActiveDate,  '
		SET @sSQL6 = @sSQL6 + '(SELECT	InactiveDate '
		SET @sSQL6 = @sSQL6 + 'FROM 	ClientCoAdjusterSpec '
		SET @sSQL6 = @sSQL6 + 'WHERE 	ClientCoAdjusterSpecID = A.[RAAdjusterSpecID]) As RAAdjusterSpecInactiveDate,  '
	END

	SET @sSQL6 = @sSQL6 + '[RAAdjusterSpecID], '
	SET @sSQL6 = @sSQL6 + '[IsLocked], '
	SET @sSQL6 = @sSQL6 + '[IsDeleted], '
	SET @sSQL6 = @sSQL6 + '[DownLoadMe], '
	SET @sSQL6 = @sSQL6 + '[UpLoadMe] , '
	SET @sSQL6 = @sSQL6 + '[DownLoadAll], '
	SET @sSQL6 = @sSQL6 + '[UpLoadAll] , '
	SET @sSQL6 = @sSQL6 + '[AdminComments], '
	SET @sSQL6 = @sSQL6 + '[MiscDelimSettings], '
	SET @sSQL6 = @sSQL6 + '[DateLastUpdated] , '
	SET @sSQL6 = @sSQL6 + '(SELECT	USERNAME '
	SET @sSQL6 = @sSQL6 + 'FROM 	USERS '
	SET @sSQL6 = @sSQL6 + 'WHERE 	USERSID = A.[UpdateByUserID]) As [UpdateByUserName],  '
	SET @sSQL6 = @sSQL6 + '[UpdateByUserID]  '

	--Set the Search Dates SQL String 
	SET @SearchDates = ''
	--Searching Dates
	--LossDate
	IF @FromLossDate <> '*'
	BEGIN
		SET @SearchDates = @SearchDates + 'AND [LossDate] >= Convert(datetime,''' + @FromLossDate + ''') '
	END
	IF @ToLossDate <> '*'
	BEGIN
		SET @SearchDates = @SearchDates + 'AND [LossDate] <= Convert(datetime,''' + @ToLossDate + ''') '
	END
	--AssignedDate
	IF @FromAssignedDate <> '*'
	BEGIN
		SET @SearchDates = @SearchDates + 'AND [AssignedDate] >= Convert(datetime,''' + @FromAssignedDate + ''') '
	END
	IF @ToAssignedDate <> '*'
	BEGIN
		SET @SearchDates = @SearchDates + 'AND [AssignedDate] <= Convert(datetime,''' + @ToAssignedDate + ''') '
	END
	--ReceivedDate
	IF @FromReceivedDate <> '*'
	BEGIN
		SET @SearchDates = @SearchDates + 'AND [ReceivedDate] >= Convert(datetime,''' + @FromReceivedDate + ''') '
	END
	IF @ToReceivedDate <> '*'
	BEGIN
		SET @SearchDates = @SearchDates + 'AND [ReceivedDate] <= Convert(datetime,''' + @ToReceivedDate + ''') '
	END
	--ContactDate
	IF @FromContactDate <> '*'
	BEGIN
		SET @SearchDates = @SearchDates + 'AND [ContactDate] >= Convert(datetime,''' + @FromContactDate + ''') '
	END
	IF @ToContactDate <> '*'
	BEGIN
		SET @SearchDates = @SearchDates + 'AND [ContactDate] <= Convert(datetime,''' + @ToContactDate + ''') '
	END
	--InspectedDate
	IF @FromInspectedDate <> '*'
	BEGIN
		SET @SearchDates = @SearchDates + 'AND [InspectedDate] >= Convert(datetime,''' + @FromInspectedDate + ''') '
	END
	IF @ToInspectedDate <> '*'
	BEGIN
		SET @SearchDates = @SearchDates + 'AND [InspectedDate] <= Convert(datetime,''' + @ToInspectedDate + ''') '
	END
	--CloseDate
	IF @FromCloseDate <> '*'
	BEGIN
		SET @SearchDates = @SearchDates + 'AND [CloseDate] >= Convert(datetime,''' + @FromCloseDate + ''') '
	END
	IF @ToCloseDate <> '*'
	BEGIN
		SET @SearchDates = @SearchDates + 'AND [CloseDate] <= Convert(datetime,''' + @ToCloseDate + ''') '
	END
	--DateReassigned
	IF @FromDateReassigned <> '*'
	BEGIN
		SET @SearchDates = @SearchDates + 'AND [DateReassigned] >= Convert(datetime,''' + @FromDateReassigned + ''') '
	END
	IF @ToDateReassigned <> '*'
	BEGIN
		SET @SearchDates = @SearchDates + 'AND [DateReassigned] <= Convert(datetime,''' + @ToDateReassigned + ''') '
	END
	--DateLastUpdated
	IF @FromDateLastUpdated <> '*'
	BEGIN
		SET @SearchDates = @SearchDates + 'AND [DateLastUpdated] >= Convert(datetime,''' + @FromDateLastUpdated + ''') '
	END
	IF @ToDateLastUpdated <> '*'
	BEGIN
		SET @SearchDates = @SearchDates + 'AND [DateLastUpdated] <= Convert(datetime,''' + @ToDateLastUpdated + ''') '
	END
	--Searching Mailing
	SET @SearchMailing = ''
	IF @SearchMAAddress <> '*'
	BEGIN
		SET @SearchMailing = @SearchMailing + 'AND [MailingAddress] Like ''' + Replace(@SearchMAAddress,'*','%') + ''' '
	END
	IF @SearchMAStreet <> '*'
	BEGIN
		SET @SearchMailing = @SearchMailing + 'AND [MAStreet] Like ''' + Replace(@SearchMAStreet,'*','%') + ''' '
	END
	IF @SearchMACity <> '*'
	BEGIN
		SET @SearchMailing = @SearchMailing + 'AND [MACity] Like ''' + Replace(@SearchMACity,'*','%') + ''' '
	END
	IF @SearchMAState <> '*'
	BEGIN
		SET @SearchMailing = @SearchMailing + 'AND [MAState] Like ''' + Replace(@SearchMAState,'*','%') + ''' '
	END
	IF @FromMAZIP <> '*'
	BEGIN
		IF @FromMAZIP = '' Or @ToMAZIP = ''
		BEGIN
			SET @SearchMailing = @SearchMailing + 'AND [MAZIP] Is Null '
		END
		ELSE
		BEGIN
			SET @SearchMailing = @SearchMailing + 'AND [MAZIP] >= Convert(int,''' + @FromMAZIP + ''') '
		END
		
	END
	IF @ToMAZIP <> '*'
	BEGIN
		IF @FromMAZIP = '' Or @ToMAZIP = ''
		BEGIN
			SET @SearchMailing = @SearchMailing + 'AND [MAZIP] Is Null '
		END
		ELSE
		BEGIN
			SET @SearchMailing = @SearchMailing + 'AND [MAZIP] <= Convert(int,''' + @ToMAZIP + ''') '
		END
	END
	IF @FromMAZIP4 <> '*'
	BEGIN
		IF @FromMAZIP4 = '' Or @ToMAZIP4 = ''
		BEGIN
			SET @SearchMailing = @SearchMailing + 'AND [MAZIP4] Is Null '
		END
		ELSE
		BEGIN
			SET @SearchMailing = @SearchMailing + 'AND [MAZIP4] >= Convert(int,''' + @FromMAZIP4 + ''') '
		END
	END
	IF @ToMAZIP4 <> '*'
	BEGIN
		IF @FromMAZIP4 = '' Or @ToMAZIP4 = ''
		BEGIN
			SET @SearchMailing = @SearchMailing + 'AND [MAZIP4] Is Null '
		END
		ELSE
		BEGIN
			SET @SearchMailing = @SearchMailing + 'AND [MAZIP4] <= Convert(int,''' + @ToMAZIP4 + ''') '
		END
		
	END

	--Searching Property
	SET @SearchProperty = ''
	IF @SearchPAAddress <> '*'
	BEGIN
		SET @SearchProperty = @SearchProperty + 'AND [PropertyAddress] Like ''' + Replace(@SearchPAAddress,'*','%') + ''' '
	END
	IF @SearchPAStreet <> '*'
	BEGIN
		SET @SearchProperty = @SearchProperty + 'AND [PAStreet] Like ''' + Replace(@SearchPAStreet,'*','%') + ''' '
	END
	IF @SearchPACity <> '*'
	BEGIN
		SET @SearchProperty = @SearchProperty + 'AND [PACity] Like ''' + Replace(@SearchPACity,'*','%') + ''' '
	END
	IF @SearchPAState <> '*'
	BEGIN
		SET @SearchProperty = @SearchProperty + 'AND [PAState] Like ''' + Replace(@SearchPAState,'*','%') + ''' '
	END
	IF @FromPAZIP <> '*'
	BEGIN
		IF @FromPAZIP = '' Or @ToPAZIP = ''
		BEGIN
			SET @SearchProperty = @SearchProperty + 'AND [PAZIP] Is Null '
		END
		ELSE
		BEGIN
			SET @SearchProperty = @SearchProperty + 'AND [PAZIP] >= Convert(int,''' + @FromPAZIP + ''') '
		END
		
	END
	IF @ToPAZIP <> '*'
	BEGIN
		IF @FromPAZIP = '' Or @ToPAZIP = ''
		BEGIN
			SET @SearchProperty = @SearchProperty + 'AND [PAZIP] Is Null '
		END
		ELSE
		BEGIN
			SET @SearchProperty = @SearchProperty + 'AND [PAZIP] <= Convert(int,''' + @ToPAZIP + ''') '
		END
	END
	IF @FromPAZIP4 <> '*'
	BEGIN
		IF @FromPAZIP4 = '' Or @ToPAZIP4 = ''
		BEGIN
			SET @SearchProperty = @SearchProperty + 'AND [PAZIP4] Is Null '
		END
		ELSE
		BEGIN
			SET @SearchProperty = @SearchProperty + 'AND [PAZIP4] >= Convert(int,''' + @FromPAZIP4 + ''') '
		END
	END
	IF @ToPAZIP4 <> '*'
	BEGIN
		IF @FromPAZIP4 = '' Or @ToPAZIP4 = ''
		BEGIN
			SET @SearchProperty = @SearchProperty + 'AND [PAZIP4] Is Null '
		END
		ELSE
		BEGIN
			SET @SearchProperty = @SearchProperty + 'AND [PAZIP4] <= Convert(int,''' + @ToPAZIP4 + ''') '
		END
		
	END
	
	--Searching Report Specs
	SET @SearchReportSpecs = ''
	IF @SelCompanyID Is Not Null And @bShowCompanyInfo = 1
	BEGIN
		SET @SearchReportSpecs = @SearchReportSpecs + 'AND [CompanyID] = ' + Cast(@SelCompanyID as varchar(10)) + ' '
	END
	IF @SelClientCompanyID Is Not Null And @bShowClientCompanyInfo = 1
	BEGIN
		SET @SearchReportSpecs = @SearchReportSpecs + 'AND [ClientCompanyID] = ' + Cast(@SelClientCompanyID as varchar(10)) + ' '
	END
	IF @SelCatID Is Not Null And @bShowCatInfo = 1
	BEGIN
		SET @SearchReportSpecs = @SearchReportSpecs + 'AND [CatID] = ' + Cast(@SelCatID as varchar(10)) + ' '
	END
	IF @SelClientCompanyUserID Is Not Null And @bShowAdjusterSpecUsersInfo = 1
	BEGIN
		IF @SelClientCompanyUserID <> -1
		BEGIN

			SET @SearchReportSpecs = @SearchReportSpecs + 'AND [AdjusterSpecUsersID] = ' + Cast(@SelClientCompanyUserID as varchar(10)) + ' '
		END 
	END
	IF @SelAssignmentTypeID Is Not Null
	BEGIN
		IF @SelAssignmentTypeID <> -1
		BEGIN
			SET @SearchReportSpecs = @SearchReportSpecs + 'AND [AssignmentTypeID] = ' + Cast(@SelAssignmentTypeID as varchar(10)) + ' '
		END
	END

	IF @bViewHistory Is Null
	BEGIN
		SET @sSQL6 = @sSQL6 + 'FROM Assignments A '
		SET @sSQL6 = @sSQL6 + ') RetAssignments '
		IF @bHideDeleted = 1
		BEGIN
			SET @sSQL6 = @sSQL6 + 'WHERE IsDeleted = 0 '
		END
		ELSE
		BEGIN
			SET @sSQL6 = @sSQL6 + 'WHERE IsDeleted Is Not Null '
		END
		--Search Dates
		IF @SearchDates <>  ''
		BEGIN
			SET @sSQL6 = @sSQL6 + @SearchDates + ' '
		END 
		--Search Property
		IF @SearchProperty <> ''
		BEGIN
			SET @sSQL6 = @sSQL6 + @SearchProperty + ' '
		END
		--Search Mailing 
		IF @SearchMailing <> ''
		BEGIN
			SET @sSQL6 = @sSQL6 + @SearchMailing + ' '
		END
		--Search Report Specs
		IF @SearchReportSpecs <> ''
		BEGIN
			SET @sSQL6 = @sSQL6 + @SearchReportSpecs + ' '
		END
		--Searching
		IF @SearchBy Is Not Null
		BEGIN
			SET @sSQL6 = @sSQL6 + @SearchBy + ' '
		END
	END
	ELSE
	BEGIN
		SET @sSQL6 = @sSQL6 + 'FROM AssignmentsHistory A '
		SET @sSQL6 = @sSQL6 + ') RetAssignments '
		SET @sSQL6 = @sSQL6 + 'WHERE [DateLastUpdated] >= Convert(datetime,''' + @FromDate + ''') '
		SET @sSQL6 = @sSQL6 + 'AND [DateLastUpdated] <= Convert(datetime,''' + @ToDate + ''') '
		IF @AssignmentsID <> -1
		BEGIN
			SET @sSQL6 = @sSQL6 + 'AND [AssignmentsID] = ' + Cast(@AssignmentsID as varchar(10)) + ' '
		END
		--Search Dates
		IF @SearchDates <>  ''
		BEGIN
			SET @sSQL6 = @sSQL6 + @SearchDates + ' '
		END 
		--Search Property
		IF @SearchProperty <> ''
		BEGIN
			SET @sSQL6 = @sSQL6 + @SearchProperty + ' '
		END
		--Search Mailing 
		IF @SearchMailing <> ''
		BEGIN
			SET @sSQL6 = @sSQL6 + @SearchMailing + ' '
		END
		--Search Report Specs
		IF @SearchReportSpecs <> ''
		BEGIN
			SET @sSQL6 = @sSQL6 + @SearchReportSpecs + ' '
		END
		--Searching
		IF @SearchBy Is Not Null
		BEGIN
			SET @sSQL6 = @sSQL6 + @SearchBy + ' '
		END
	END
	IF @bViewHistory Is Null
	BEGIN
		--Grouping
		IF @GroupBy Is Not Null
		BEGIN
			SET @sSQL6 = @sSQL6 + 'GROUP BY ' + @GroupBy + ' '			
		END
		---Sorting 
		IF @GroupBy Is Null
		BEGIN
			IF @OrderBy Is Null OR @OrderBy = 'Null'
			BEGIN
				IF @bShowAdjusterSpecUsersInfo =1 And @bShowRAAdjusterSpecInfo = 1
				BEGIN
					SET @sSQL6 = @sSQL6 + 'ORDER BY [AdjusterSpecACID], [AdjusterSpecLastName], [AssignedDate] '
				END
				ELSE
				BEGIN
					SET @sSQL6 = @sSQL6 + 'ORDER BY [AdjusterSpecACID], [AssignedDate] '
				END
					
			END
			ELSE
			BEGIN
				IF @bShowAdjusterSpecUsersInfo =1 And @bShowRAAdjusterSpecInfo = 1
				BEGIN
					SET @sSQL6 = @sSQL6 + 'ORDER BY [AdjusterSpecACID], [AdjusterSpecLastName], ' + @OrderBy + ' '
				END
				ELSE
				BEGIN
					SET @sSQL6 = @sSQL6 + 'ORDER BY [AdjusterSpecACID], ' + @OrderBy + ' '
				END
			END
		END
	END
	ELSE
	BEGIN
		--Grouping
		IF @GroupBy Is Not Null
		BEGIN
			SET @sSQL6 = @sSQL6 + 'GROUP BY ' + @GroupBy + ' '			
		END
		---Sorting
		IF @GroupBy Is Null
		BEGIN
			IF @OrderBy Is Null OR @OrderBy = 'Null'
			BEGIN
				IF @bShowAdjusterSpecUsersInfo =1 And @bShowRAAdjusterSpecInfo = 1
				BEGIN
					SET @sSQL6 = @sSQL6 + 'ORDER BY [AdjusterSpecACID], [AdjusterSpecLastName], [AssignedDate] '
				END
				ELSE
				BEGIN
					SET @sSQL6 = @sSQL6 + 'ORDER BY [AdjusterSpecACID], [AssignedDate] '
				END
					
			END
			ELSE
			BEGIN
				IF @bShowAdjusterSpecUsersInfo =1 And @bShowRAAdjusterSpecInfo = 1
				BEGIN
					SET @sSQL6 = @sSQL6 + 'ORDER BY [AdjusterSpecACID], [AdjusterSpecLastName], ' + @OrderBy + ' '
				END
				ELSE
				BEGIN
					SET @sSQL6 = @sSQL6 + 'ORDER BY [AdjusterSpecACID], ' + @OrderBy + ' '
				END
			END
		END
	END
	SET @sSQL7 =''
	SET @sSQL8 =''
	SET @sSQL9 =''
	SET @sSQL10 =''
--Debug
-- 	Print Len(@sSQL1) + Len(@sSQL2) + Len(@sSQL3) + Len(@sSQL4) + Len(@sSQL5) + Len(@sSQL6) + Len(@sSQL7) + Len(@sSQL8) + Len(@sSQL9) + Len(@sSQL10)
-- 	Print @sSQL1
-- 	Print @sSQL2
-- 	Print @sSQL3
-- 	Print @sSQL4
-- 	Print @sSQL5
-- 	Print @sSQL6
-- 	Print @sSQL7
-- 	Print @sSQL8
-- 	Print @sSQL9
-- 	Print @sSQL10
--End Debug
	EXEC(@sSQL1 + @sSQL2 + @sSQL3 + @sSQL4 + @sSQL5 + @sSQL6 + @sSQL7 + @sSQL8 + @sSQL9 + @sSQL10)
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[z_spsGetAssignmentsInfo]  TO [CFUSER]
GO

