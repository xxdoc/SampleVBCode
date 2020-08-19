SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spuInitCCAssignmentSpecs]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spuInitCCAssignmentSpecs]
GO

/******************************************************************
*
*    	Description:	Initialize Company / Client Company Assignment Specs
*			Use this SP for All Client Company Loss SP 
*	Author: 	Brad Skidmore
*	Date: 		1/14/2004
*	
*	Date		2/12/2004	Changed the Params for z_spuInitCCAssignAdjToZipCode
*			4.27.2005	BGS Adding UNASSIGNEDAUTO for Automobile Assignments
******************************************************************/


CREATE PROCEDURE [dbo].[z_spuInitCCAssignmentSpecs]
	--Required params
	@UserName			varchar(20),		
	@CompanyCode			varchar(3),
	@ClientCompanyCode		varchar(3),
	@CatCode			varchar(50),
	@ACID				varchar(50),	
	@TypeOfACID			varchar(20),
	@DefualtFeeScheduleName		varchar(20),
	@AssignmentType			varchar(20),
	@iAssignByZIP			int=Null,
	--Output params
	@UID 				int=null OUTPUT,
	@CompanyID			int=null OUTPUT,
	@CompanyName			varchar(30)=null OUTPUT,
	@ClientCompanyID		int=null OUTPUT,
	@ClientCompanyName		varchar(30)=null OUTPUT,
	@UNKOWNTypeOfLossID		int=null OUTPUT,
	@PropertyBaseFeeScheduleID	int=null OUTPUT,
	@UACATIDProperty		int=null OUTPUT,
	@UAUID				int=null OUTPUT,
	@ClientCompanyCatSpecID		int=null OUTPUT,
	@AdjusterSpecID			int=null OUTPUT,
	@AdjusterSpecIDDisplay		int=null OUTPUT,
	@TypeOfLossID			int=null OUTPUT,
	@FeeScheduleID			int=null OUTPUT,
	@AssignmentTypeID		int=null OUTPUT,
	@MiscDelimSettings		varchar(8000)=null OUTPUT
	

AS
DECLARE @TempID int
DECLARE @sTempACID varchar(50)
DECLARE @sTempACIDDescription varchar(150)
DECLARE @sTempGetDate datetime
DECLARE @TempUACatName varchar(20)

--Set the Fields 
--Get the UserID for the User doing the DB Updates on WebControl
SET @UID = 	(
		SELECT 	UsersID 
		FROM 	Users 
		WHERE 	UserName = @UserName
		)
--Get Company ID
SET @CompanyID = 	(	
			SELECT 	CompanyID
			FROM 	Company
			WHERE	Code = @CompanyCode
			)
SET @CompanyName = 	(
			SELECT 	[Name]
			FROM 	Company
			WHERE	CompanyID = @CompanyID
			)
--Get Client Company
SET @ClientCompanyID = 	(
		 	SELECT 	CompanyID
			FROM 	Company
			WHERE 	Code = @ClientCompanyCode
			AND 	IsClientOf = @CompanyID
			)
SET @ClientCompanyName =	(
				SELECT 	[Name]
				FROM 	Company
				WHERE	CompanyID = @ClientCompanyID
				)
--Get The Unknown Type Of Loss ID for this Client Company
SET @UNKOWNTypeOfLossID = 	(
				SELECT 	TypeOfLossID 
				FROM	TypeOfLoss
				WHERE	ClientCompanyID = @ClientCompanyID
				AND 	TypeOfLoss = 'UNKNOWN'
				)
IF @UNKOWNTypeOfLossID Is Null  
BEGIN
	INSERT INTO TypeOfLoss
			(
			[ClientCompanyID],
			[TypeOfLoss],
			[Description],
			[Code],
			[IsDeleted],
			[DateLastUpdated],
			[UpdateByUserID]
			)
	SELECT		@ClientCompanyID As [ClientCompanyID],
			'UNKNOWN' As [TypeOfLoss],
			'UNKNOWN' As [Description],
			'NA' As [Code],
			0 As [IsDeleted],
			GetDate() As [DateLastUpdated],
			@UID As [UpdateByUserID]
	SET @UNKOWNTypeOfLossID = (Select IDENT_CURRENT('TypeOfLoss') as TypeOfLossID)
END	

--Get Property Base Fee Schedule ID
SET @PropertyBaseFeeScheduleID = 	(
						SELECT 	FeeScheduleID
						FROM	FeeSchedule
						WHERE	ClientCompanyID = @ClientCompanyID
						AND 	ScheduleName = 'Sched01'
					)


IF @PropertyBaseFeeScheduleID Is Null 
BEGIN
	--If there is no Base Fee Schedule for Property Then Add it.
	INSERT INTO FeeSchedule
				(
				[ClientCompanyID],
				[ScheduleName],
				[Description],
				[NumOfLevels],
				[NumOfFeeTypes],
				[FeeServiceHourlyRate],
				[TaxPercent],
				[InitialOptions],
				[Options],
				[DefaultAppDedClassTypeIDOrder],
				[IsDeleted],
				[DateLastUpdated],
				[UpdateByUserID]
				)
	SELECT			@ClientCompanyID As [ClientCompanyID],
				@DefualtFeeScheduleName As [ScheduleName],
				@ClientCompanyName + ' Base Fee Schedule' As [Description],
				0 As [NumOfLevels],
				0 As [NumOfFeeTypes],
				0 As [FeeServiceHourlyRate],
				0 As [TaxPercent],
				'' As [InitialOptions],
				'' As [Options],
				'1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,23,22,24,25,26' As [DefaultAppDedClassTypeIDOrder],
				0 As [IsDeleted],
				GetDate() As [DateLastUpdated],
				@UID As [UpdateByUserID]
	SET @PropertyBaseFeeScheduleID = (Select IDENT_CURRENT('FeeSchedule') as FeeScheduleID)
END


SET @AssignmentTypeID = 	(	SELECT 	AssignmentTypeID 
					FROM 	AssignmentType
					WHERE	Type = @AssignmentType
				)

--Get The Property UNASSIGNED CAT ID
IF @AssignmentType = 'Property'
BEGIN
	SET @TempUACatName = 'UNASSIGNED'
END
ELSE
BEGIN
	IF  @AssignmentType = 'Auto'
	BEGIN
		SET @TempUACatName = 'UNASSIGNEDAUTO'
	END
END

SET @UACATIDProperty = 	(	
			SELECT	CATID	
			FROM	CAT
			WHERE	Name = @TempUACatName
			)

IF @UACATIDProperty Is Null
BEGIN
	--If there is not CAT Un Assigned Property for this Company then Create it!
	INSERT INTO CAT	(	
			[CompanyID] ,
			[Name] ,
			[Description],
			[ActiveDate],
			[InactiveDate],
			[DateLastUpdated],
			[UpdateByUserID]
			)
	SELECT		@CompanyID As [CompanyID],
			@TempUACatName As [Name],
			'Claims yet to be assigned to CAT' As [Description],
			GetDate() As [ActiveDate],
			Null As [InactiveDate],
			GetDate() As [DateLastUpdated],
			@UID As [UpdateByUserID]
	SET @UACATIDProperty = (Select IDENT_CURRENT('CAT') as UACATIDProperty)

END
--Check to See if the ClientCO Cat For Unassigned was set up
SET @TempID =	(
		SELECT 	ClientCompanyID
		FROM 	ClientCompanyCAT
		WHERE	CATID = @UACATIDProperty
		AND ClientCompanyID = @ClientCompanyID
		)
IF @TempID Is Null
BEGIN
	--If No Client CO Cat for Unassigned Create it
	INSERT INTO ClientCompanyCat	(
					[ClientCompanyID],
					[CATID],
					[TypeOfLossID],
					[FeeScheduleID],
					[SiteAddress],
					[SACity],
					[SAState],
					[SAZip],
					[SAZip4],
					[SAOtherPostCode],
					[ActiveDate],
					[InactiveDate],
					[DateLastUpdated],
					[UpdateByUserID]
					)
	SELECT				@ClientCompanyID AS [ClientCompanyID],
					@UACATIDProperty As [CATID],
					@UNKOWNTypeOfLossID As [TypeOfLossID],
					@PropertyBaseFeeScheduleID As [FeeScheduleID],
					null As [SiteAddress],
					null As [SACity],
					null As [SAState],
					null As [SAZip],
					null As [SAZip4],
					null As [SAOtherPostCode],
					GetDate() As [ActiveDate],
					Null As [InactiveDate],
					GetDate() As [DateLastUpdated],
					@UID As [UpdateByUserID]
END

SET @UAUID = 	(	
		SELECT 	UsersID 
		FROM 	Users 
		WHERE 	UserName = 'UUNASSIGNED'
		)
IF @UAUID Is Null
BEGIN
	--If there is No UNASSIGNED User Set Up Create it
	INSERT INTO Users	(
				[UserName] ,
				[PassWord] ,
				[FirstName],
				[LastName] ,
				[SSN] ,
				[Email],
				[ContactPhone],
				[EmergencyPhone],
				[Address] ,
				[City] ,
				[State],
				[Zip] ,
				[ZIP4],
				[OtherPostCode],
				[Active] ,
				[ActiveDate],
				[InactiveDate],
				[SecurityLevel] ,
				[Comments] ,
				[DateLastUpdated],
				[UpdateByUserID] 
				)
	SELECT			'UUNASSIGNED' As [UserName] ,
				' z51223-27i l71223-E5v i02223-32y k81223-s4f m12223- 1l a91223-a3d f41223-18y g61223-Z6x' As [PassWord] ,
				'UNASSIGNED' As [FirstName],
				'UNASSIGNED' As [LastName] ,
				null As [SSN] ,
				null As [Email],
				null As [ContactPhone],
				null As [EmergencyPhone],
				null As [Address] ,
				null As [City] ,
				null As [State],
				null As [Zip] ,
				null As [ZIP4],
				null As [OtherPostCode],
				1 As [Active] ,
				GetDate() As [ActiveDate],
				null As [InactiveDate],
				1 As [SecurityLevel] ,
				'WEBCONTROL UPDATED Added UNASSIGNED User as default account for Unassigned Adjusters' As [Comments] ,
				GetDate() As [DateLastUpdated],
				@UID As [UpdateByUserID] 
	SET @UAUID = (Select IDENT_CURRENT('Users') as UAUID)
END
--Set the Client Company Cat Spec to the Latest ID for this CLient Company Cat Code
SET @ClientCompanyCatSpecID = 	(
				SELECT 	Top 1 ClientCompanyCatSpecID 
				FROM	ClientCompanyCatSpec
				WHERE	ClientCompanyID = @ClientCompanyID
				AND 	CatCode = @CatCode
				AND	AssignmentTypeID = @AssignmentTypeID
				AND	InactiveDate Is Null
				Order By ActiveDate Desc
				)
IF @ClientCompanyCatSpecID Is Null 
--If there is not a current Spec (Means there never was a spec assigned for this Cat or 
--there was but is currently inactive) for this Cat Code,  Add it to the  Unassigned Property Cat Spec 
--for this CLient Company.
BEGIN
	
	INSERT INTO ClientCompanyCatSpec	(
						[ClientCompanyID],
						[CATID],
						[CatCode],
						[Comments],
						[ActiveDate],
						[InactiveDate],
						[DateLastUpdated],
						[UpdateByUserID],
						[AssignmentTypeID]
						)
	SELECT					@ClientCompanyID As [ClientCompanyID],
						@UACATIDProperty As [CATID],
						@CatCode As [CatCode],
						'WEBCONTROL UPDATED UNASSIGNED' As [Comments],
						GetDate() As [ActiveDate],
						Null As [InactiveDate],
						GetDate() As [DateLastUpdated],
						@UID As [UpdateByUserID],
						@AssignmentTypeID As [AssignmentTypeID]
	SET @ClientCompanyCatSpecID = (Select IDENT_CURRENT('ClientCompanyCatSpec') as ClientCompanyCatSpecID)
END	

--Get and Set the CLient Company Variables: TypeOfLossID and FeeScheduleID, from the
-- Client Company Cat Spec Table...
DECLARE RSClientCo Cursor for
SELECT 	(	
		SELECT 	TypeOfLossId
		FROM	ClientCompanyCat
		WHERE	ClientCompanyID = CCCS.ClientCompanyID
		AND	CATID = CCCS.CATID
	) As TypeOfLossID,
	(
		SELECT 	FeeScheduleID
		FROM	ClientCompanyCat
		WHERE	ClientCompanyID = CCCS.ClientCompanyID
		AND	CATID = CCCS.CATID
	)As FeeScheduleID
FROM 	ClientCompanyCatSpec CCCS
WHERE 	ClientCompanyCatSpecID = @ClientCompanyCatSpecID
OPEN RSClientCo

Fetch Next From RSClientCo into @TypeOfLossID, @FeeScheduleID

Close RSClientCo
Deallocate RSClientCo

--If Assigning by ZIPCODE need to Check for Users Already Assigned to Specific ZIPCODES for a Specific Cat
--As Well need to Account for the Same ZipCode For a Specific Cat being Assigned to More than one Adjuster.
--This means the Assignement being processed will be Assigned to the adjuster with the Least Number Of Assignment for that Zip Code.
--As Well, IF USER has Already been set up for the ZipCode on the Assignment Being Processed, Need to Assign that to the
--UNASSIGNED Account.  When an Assignment is Assigned to the UNASSIGNED Account,  The Admin will have to Manually Assign the 
--ZIP Code Assigned to the UNASSIGNED Account to an Adjuster Assigned to the same CAT.
IF @iAssignByZIP is Not Null
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
ELSE --NOT ASSIGNING BY ZIPCODE BUT CLIENT SUPPLIED ACID (Adjuster Client ID)
BEGIN
	--Set the Client Company Adjuster Spec to the Latest ID for this Client Company ACID (Adjuster Client ID)
	SET @AdjusterSpecID = 	(	
				SELECT 	Top 1 ClientCoAdjusterSpecID
				FROM 	ClientCoAdjusterSpec
				WHERE 	ClientCompanyID = @ClientCompanyID	
				AND 	ACID = @ACID 
				AND	InactiveDate Is Null
				Order By ActiveDate Desc
				)
	--If not found then Set it to UNASSIGNED account
	IF @AdjusterSpecID Is Null 
	BEGIN
		SET @sTempACIDDescription = @ClientCompanyName + ' ' + @TypeOfACID
		SET @sTempGetDate = GetDate()
		EXEC z_spsGetClientCoAdjusterSpecZIPInfo 	
		/*@bHideDeleted */				0,
		/*@UID */					@UID, 
		/*@ClientCoAdjusterSpecID */			1,
		/*@OrderBy */					Null,
		/*@GroupBy */					Null,
		/*@SearchBy */					Null,
		/*@bDelete */					0, 
		/*@bAdd */					1, 
		/*@bEdit */					0,	
		/*@ClientCompanyID	*/			@ClientCompanyID,
		/*@UsersID*/					@UAUID,
		/*@ACID*/					@ACID,
		/*@ACIDDescription*/				@sTempACIDDescription,
		/*@Comments	*/				'WEBCONTROL UPDATED',
		/*@ClientCompanyCatSpecID*/			null,
		/*@ZipCode*/					null,
		/*@ZipCodeCount*/				null,
		/*@ActiveDate */				@sTempGetDate,
		/*@InactiveDate */				null
		SET @AdjusterSpecID = (Select IDENT_CURRENT('ClientCoAdjusterSpec') as AdjusterSpecID)
		
	END  
END

-- Update the @AdjusterSpecIDDisplay if it is Null
--This ID is used for those carriers that like to process a claim under on ID
-- And then have reports (Like the IB) returned to them displaying another ID also associated with the Adjuster.
IF @AdjusterSpecIDDisplay Is Null 
BEGIN
	SET @AdjusterSpecIDDisplay = @AdjusterSpecID
END


--Set the MiscDelim Settings from FeeSchedule
SET @MiscDelimSettings = dbo.GetFEESCHEDULEMiscDelimSettings(@FeeScheduleID)

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[z_spuInitCCAssignmentSpecs]  TO [CFUSER]
GO

