SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spsGetBillingAssignmentsInfo]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spsGetBillingAssignmentsInfo]
GO

/******************************************************************
*
*    	Description:	Get Billing Assignments Info
*			
*	
*	Author: 	Brad Skidmore
*	Date: 		3/15/2005
******************************************************************/


CREATE PROCEDURE [dbo].[z_spsGetBillingAssignmentsInfo]
	@bHideDeleted		bit =1, 
	@UID			int,
	@BillAssignmentID	Int=null,
--Include (INNER JOIN) Dependant Tables
	@bShowAssignmentsInfo			bit=0,
	@bShowAssignmentTypeInfo		bit=0,
	@bShowCompanyInfo			bit=0,
	@bShowCatInfo				bit=0,
	@bShowClientCompanyInfo			bit=0,
	@bShowClientCompanyCatInfo		bit=0,
	@bShowClientCompanyCatSpecInfo		bit=0,
	@bShowAdjusterSpecUsersInfo		bit=0,
	@bShowAdjusterSpecInfo			bit=0,
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
	@AssignmentTypeID 	int=null,
	@ClientCompanyCatSpecID int=null,
	@AdjusterSpecID 	int=null,
	@IBNUM 			varchar(20)=null,
	@CLIENTNUM 		varchar(20)=null,
	@PolicyNo 		varchar(50)=null,
	@Insured 		varchar(100)=null,
	@LossLoc1 		varchar(50)=null,
	@LossLoc2 		varchar(50)=null,
	@LossLocCity 		varchar(50)=null,
	@LossLocState 		varchar(50)=null,
	@LossLocZipcode 	varchar(10)=null,
	@LossDate 		datetime=null,
	@CloseDate 		datetime=null,
	@IsLocked 		bit=0,
	@IsDeleted 		bit=0,
	@AdminComments 		varchar(1000)=null,
	@MiscDelimSettings 	text=null,
	@DateLastUpdated 	datetime=null,
--OtherSearchDateStringSpans
	@FromLossDate		varchar(50)='*',
	@ToLossDate		varchar(50)='*',
	@FromCloseDate		varchar(50)='*',
	@ToCloseDate		varchar(50)='*',
	@FromDateLastUpdated	varchar(50)='*',
	@ToDateLastUpdated	varchar(50)='*',
--Report Specs
	@SelCompanyID		int=null,
	@SelClientCompanyID	int=null,
	@SelCatID		int=null,
	@SelClientCompanyUserID	int=null,
	@SelAssignmentTypeID	int=null

AS
DECLARE @SearchDates		varchar(2000)
DECLARE @SearchReportSpecs	varchar(2000)
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
DECLARE @ClientCompanyName	varchar(30)	--Holds the Name of the Client Company
DECLARE @ClientCompanyID	int		--Holds ID of the Client Company
DECLARE @CurDate 		DateTime
SET @CurDate = GetDate()
--See if the DateLastupdated is null
IF IsDate(@DateLastUpdated) =0
BEGIN
	SET @DateLastUpdated = GetDate()
END
--Clean Date Strings
SET @FromLossDate =		dbo.CleanFromOrToDateString(@CurDate, @FromLossDate, 1)
SET @ToLossDate =		dbo.CleanFromOrToDateString(@CurDate, @ToLossDate, 0)
SET @FromCloseDate =		dbo.CleanFromOrToDateString(@CurDate, @FromCloseDate, 1)
SET @ToCloseDate =		dbo.CleanFromOrToDateString(@CurDate, @ToCloseDate, 0)
SET @FromDateLastUpdated =	dbo.CleanFromOrToDateString(@CurDate, @FromDateLastUpdated, 1)
SET @ToDateLastUpdated =	dbo.CleanFromOrToDateString(@CurDate, @ToDateLastUpdated, 0)

IF @BillAssignmentID Is Not Null 
BEGIN
	IF @bDelete=0 And @bAdd=0 And @bEdit=0
	BEGIN
		SELECT 
			[BillAssignmentID],
			[AssignmentTypeID],
			[ClientCompanyCatSpecID],
			[AdjusterSpecID],
			[IBNUM],
			[CLIENTNUM],
			[PolicyNo],
			[Insured],
			[LossLoc1],
			[LossLoc2],
			[LossLocCity],
			[LossLocState],
			[LossLocZipcode],
			[LossDate],
			[CloseDate],
			[IsLocked],
			[IsDeleted],
			[AdminComments],
			[MiscDelimSettings],
			[StatusID],
			[DateLastUpdated],
			[UpdateByUserID]
		FROM	BillAssignment
		WHERE 	[BillAssignmentID] = @BillAssignmentID
		ORDER BY [LossDate] 
	END

	ELSE
	BEGIN
		IF @bDelete =1
		BEGIN
			UPDATE BillAssignment SET [IsDeleted]=	(CASE 	WHEN [IsDeleted]= 0
									THEN 1
									ELSE 0
									END
								),
						[DateLastUpdated]=@DateLastUpdated ,
						[UpdateByUserID]=@UID
			WHERE BillAssignmentID = @BillAssignmentID
		
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
					--GET The Client Company Name
					SET @ClientCompanyName =	(
									SELECT 	[Name]
									FROM 	Company
									WHERE	CompanyID = @ClientCompanyID
									)
					EXEC z_spuInsertBillingAssignment
						@AssignmentTypeID		,--int,
						@ClientCompanyCatSpecID		,--int,
						@AdjusterSpecID 		,--int,
						@IBNUM 				,--varchar(20),
						@CLIENTNUM 			,--varchar(20),
						@PolicyNo 			,--varchar(50),
						@Insured 			,--varchar(100),
						@LossLoc1 			,--varchar(50),
						@LossLoc2 			,--varchar(50),
						@LossLocCity 			,--varchar(50),
						@LossLocState 			,--varchar(50),
						@LossLocZipcode 		,--varchar(10),
						@LossDate 			,--datetime=null,
						@CloseDate 			,--datetime=null,
						@IsLocked 			,--bit=0,
						@IsDeleted 			,--bit=0,
						@AdminComments 			,--varchar(1000)=null,
						@MiscDelimSettings		,--text,
						@DateLastUpdated 		,--datetime,
						@UID				--int
			END
			ELSE 
			BEGIN
				IF @bEdit=1
				BEGIN
					Update BillAssignment Set 	
						[AssignmentTypeID] = @AssignmentTypeID,
						[ClientCompanyCatSpecID] = @ClientCompanyCatSpecID,
						[AdjusterSpecID] = @AdjusterSpecID,
						[IBNUM] = [IBNUM], --IBNUM is autogenerated and can not be changed !
						[CLIENTNUM] = @CLIENTNUM,
						[PolicyNo] = @PolicyNo,
						[Insured] = @Insured,
						[LossLoc1] = @LossLoc1,
						[LossLoc2] = @LossLoc2,
						[LossLocCity] = @LossLocCity,
						[LossLocState] = @LossLocState,
						[LossLocZipcode] = @LossLocZipcode,
						[LossDate] = @LossDate,
						[CloseDate] = @CloseDate,
						[IsLocked] = @IsLocked,
						[IsDeleted] = @IsDeleted,
						[AdminComments] = @AdminComments,
						[MiscDelimSettings] = @MiscDelimSettings,
						[DateLastUpdated] = @DateLastUpdated,
						[UpdateByUserID] = @UID
					WHERE [BillAssignmentID] = @BillAssignmentID
					--Because SQL Server pukes on updates to ntext, text or image files,
					--you must set a pointer to the value you want to update and use WRITE TEXT.
					--If Loss Report is set to '-1' means maintain the loss report 
					--that already exists in the DB
					IF Cast(@MiscDelimSettings as varchar(2)) <> '-1'
					BEGIN 
						SELECT 	@ptrMiscDelimSettings = TEXTPTR(MiscDelimSettings) 
						FROM 	BillAssignment
						WHERE 	BillAssignmentID	=@BillAssignmentID
						WRITETEXT BillAssignment.MiscDelimSettings @ptrMiscDelimSettings @MiscDelimSettings
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
	SET @sSQL1 = @sSQL1 + '[BillAssignmentID] , '
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

	SET @sSQL5 = ''
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
	SET @sSQL5 = @sSQL5 + '[StatusID] , '
	IF @bShowAssignmentsInfo = 1
	BEGIN
		SET @sSQL5 = @sSQL5 + '[IBNUM], '
		SET @sSQL5 = @sSQL5 + '[CLIENTNUM], '
		SET @sSQL5 = @sSQL5 + '[PolicyNo], '
		SET @sSQL5 = @sSQL5 + '[Insured], '
		SET @sSQL5 = @sSQL5 + '[LossLoc1], '
		SET @sSQL5 = @sSQL5 + '[LossLoc2], '
		SET @sSQL5 = @sSQL5 + '[LossLocCity], '
		SET @sSQL5 = @sSQL5 + '[LossLocState], '
		SET @sSQL5 = @sSQL5 + '[LossLocZipcode], '
	END

	IF @bShowBatchesInfo = 1
	BEGIN
		--From Batches Table (To be entered into Accounting System)
		SET @sSQL5 = @sSQL5 + '(SELECT Sum(GrossLoss) '
		SET @sSQL5 = @sSQL5 + 'FROM Batches '
		SET @sSQL5 = @sSQL5 + 'WHERE BillAssignmentID = A.[BillAssignmentID] And Void = 0 And Copied = 2) As BatchesGrossLoss, '
		SET @sSQL5 = @sSQL5 + '(SELECT Sum(TotalFee) '
		SET @sSQL5 = @sSQL5 + 'FROM Batches '
		SET @sSQL5 = @sSQL5 + 'WHERE BillAssignmentID = A.[BillAssignmentID] And Void = 0 And Copied = 2) As BatchesTotalFee, '
	END

	SET @sSQL6 = ''
	IF @bShowAssignmentsInfo = 1
	BEGIN
		SET @sSQL6 = @sSQL6 + '[LossDate] , '
		SET @sSQL6 = @sSQL6 + '[CloseDate] , '
	END
	SET @sSQL6 = @sSQL6 + '[IsLocked], '
	SET @sSQL6 = @sSQL6 + '[IsDeleted], '
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
	--CloseDate
	IF @FromCloseDate <> '*'
	BEGIN
		SET @SearchDates = @SearchDates + 'AND [CloseDate] >= Convert(datetime,''' + @FromCloseDate + ''') '
	END
	IF @ToCloseDate <> '*'
	BEGIN
		SET @SearchDates = @SearchDates + 'AND [CloseDate] <= Convert(datetime,''' + @ToCloseDate + ''') '
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

	SET @sSQL6 = @sSQL6 + 'FROM BillAssignment A '
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
			IF @bShowAdjusterSpecUsersInfo =1 
			BEGIN
				SET @sSQL6 = @sSQL6 + 'ORDER BY [AdjusterSpecACID], [AdjusterSpecLastName], [LossDate] '
			END
			ELSE
			BEGIN
				SET @sSQL6 = @sSQL6 + 'ORDER BY [AdjusterSpecACID], [LossDate] '
			END
				
		END
		ELSE
		BEGIN
			IF @bShowAdjusterSpecUsersInfo =1 
			BEGIN
				SET @sSQL6 = @sSQL6 + 'ORDER BY [AdjusterSpecACID], [AdjusterSpecLastName], ' + @OrderBy + ' '
			END
			ELSE
			BEGIN
				SET @sSQL6 = @sSQL6 + 'ORDER BY [AdjusterSpecACID], ' + @OrderBy + ' '
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

GRANT  EXECUTE  ON [dbo].[z_spsGetBillingAssignmentsInfo]  TO [CFUSER]
GO

