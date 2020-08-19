SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spsGetFeeScheduleInfo]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spsGetFeeScheduleInfo]
GO

/******************************************************************
*
*    	Description:	FeeScheduleInfo Info
*			
*	
*	Author: 	Brad Skidmore
*	Date: 		09/17/2003
*	Date		1/6/2004 Add DefaultAppDedClassTypeIDOrder
******************************************************************/


CREATE PROCEDURE [dbo].[z_spsGetFeeScheduleInfo]
	@bHideDeleted		bit =1, 
	@UID			int,
	@FeeScheduleID		Int=null,
	@OrderBy		varchar(500)=null,
	@GroupBy		varchar(500)=null,
	@SearchBy		varchar(1000)=null,
--updating the FeeSchedule Table
	@bDelete		bit=0,
	@bAdd			bit=0,
	@bEdit			bit=0,
--Update variables
	@ClientCompanyID 	int=null,
	@ScheduleName 		varchar(20)=null,
	@Description 		varchar(100)=null,
	@NumOfLevels 		int=null,
	@NumOfFeeTypes 		int=null,
	@FeeServiceHourlyRate 	money =null,
	@TaxPercent 		decimal(18, 4) =null,
	@InitialOptions		varchar(3000)=null,
	@Options 	 	varchar(3000)=null,
	@DefaultAppDedClassTypeIDOrder Varchar(255)=null,
	@IsDeleted 		bit =null,
	@DateLastUpdated 	datetime =null,
--ViewHistory
	@bViewHistory		bit=NULL,
	@FromDate		varchar(50)=Null,
	@ToDate			varchar(50)=Null
	
AS
DECLARE @sSQL Varchar(8000)
DECLARE @CurDate 		DateTime
--FeeSchedule Levels Vars
DECLARE @LevelCount int
DECLARE @FeeTypeCount int

SET @CurDate = GetDate()
--See if the DateLastupdated is null
IF IsDate(@DateLastUpdated) =0
BEGIN
	SET @DateLastUpdated = GetDate()
END
--Clean Dates
SET @FromDate =			dbo.CleanFromOrToDate(@CurDate, @FromDate, 1)
SET @ToDate = 			dbo.CleanFromOrToDate(@CurDate, @ToDate, 0)

IF @FeeScheduleID Is Not Null And @bViewHistory is Null
BEGIN
	IF @bDelete=0 And @bAdd=0 And @bEdit=0
	BEGIN
		SELECT 
			[FeeScheduleID],
			[ClientCompanyID] ,
			[ScheduleName],
			[Description] ,
			[NumOfLevels],
			[NumOfFeeTypes],
			[FeeServiceHourlyRate],
			[TaxPercent] ,
			[InitialOptions],
			[Options],
			[DefaultAppDedClassTypeIDOrder],
			[IsDeleted] ,
			[DateLastUpdated],
			[UpdateByUserID] 
		FROM	FeeSchedule 
		WHERE 	[FeeScheduleID] = @FeeScheduleid
	END
	ELSE
	BEGIN
		IF @bDelete =1
		BEGIN
			UPDATE FeeSchedule SET 	[IsDeleted]=	(CASE 	WHEN [IsDeleted]= 0
									THEN 1
									ELSE 0
									END
								),
						[DateLastUpdated]=@DateLastUpdated ,
						[UpdateByUserID]=@UID
			WHERE FeeScheduleID = @FeeScheduleID
		
		END
		ELSE 
		BEGIN
			IF @bAdd=1
			BEGIN
				INSERT INTO FeeSchedule(
							[ClientCompanyID] ,
							[ScheduleName],
							[Description] ,
							[NumOfLevels],
							[NumOfFeeTypes],
							[FeeServiceHourlyRate],
							[TaxPercent] ,
							[InitialOptions],
							[Options],
							[DefaultAppDedClassTypeIDOrder],
							[IsDeleted] ,
							[DateLastUpdated],
							[UpdateByUserID] 
						)
				SELECT 			@ClientCompanyID As [ClientCompanyID] ,
							@ScheduleName As [ScheduleName],
							@Description As [Description] ,
							@NumOfLevels As [NumOfLevels],
							@NumOfFeeTypes As [NumOfFeeTypes],
							@FeeServiceHourlyRate As [FeeServiceHourlyRate],
							@TaxPercent As [TaxPercent] ,
							@InitialOptions As [InitialOptions],
							@Options As [Options],
							@DefaultAppDedClassTypeIDOrder As [DefaultAppDedClassTypeIDOrder],
							@IsDeleted As [IsDeleted] ,
							@DateLastUpdated As [DateLastUpdated],
							@UID As [UpdateByUserID] 
				SET @LevelCount = 0
				SET @FeeTypeCount = 0  
				SET @FeeScheduleID = (SELECT IDENT_CURRENT('FeeSchedule') As FeeScheduleID)
				--After Insert Insert the number of Levels for the new FeeSchedule ID
				While @LevelCount < @NumOfLevels
					BEGIN
						INSERT INTO FeeScheduleLevels
							(
								[FeeScheduleID],
								[LevelNum],
								[LevelMax],
								[LevelPctApp],
								[LevelMin],
								[IsDeleted],
								[DateLastUpdated],
								[UpdateByUserID]
							)
						SELECT 
								@FeeScheduleID As [FeeScheduleID],
								@LevelCount + 1 As [LevelNum],
								0 As [LevelMax],
								0 As [LevelPctApp],
								0 As [LevelMin],
								0 As [IsDeleted],
								@DateLastUpdated As [DateLastUpdated],
								@UID As [UpdateByUserID]
						SET @LevelCount = @LevelCount + 1
					END
				--After Insert Insert the number of Fee Types for the new FeeSchedule ID
				While @FeeTypeCount < @NumOfFeeTypes
					BEGIN
						INSERT INTO FeeScheduleFeeTypes
							(
								[FeeScheduleID],
								[TypeNum],
								[Name],
								[Description],
								[FeeAmount],
								[IsDeleted],
								[DateLastUpdated],
								[UpdateByUserID]
							)
						SELECT 
								@FeeScheduleID As [FeeScheduleID],
								@FeeTypeCount + 1 As [TypeNum],
								'' As [Name],
								'' As [Description],
								0 As [FeeAmount],
								0 As [IsDeleted],
								@DateLastUpdated As [DateLastUpdated],
								@UID As [UpdateByUserID]
						SET @FeeTypeCount = @FeeTypeCount + 1
					END
			END
			ELSE 
			BEGIN
				IF @bEdit=1
				BEGIN
					Update FeeSchedule Set 	
								[ClientCompanyID] = @ClientCompanyID  ,
								[ScheduleName] = @ScheduleName ,
								[Description] = @Description  ,
								[NumOfLevels] = @NumOfLevels,
								[NumOfFeeTypes] = @NumOfFeeTypes,
								[FeeServiceHourlyRate] = @FeeServiceHourlyRate  ,
								[TaxPercent] = @TaxPercent   ,
								[InitialOptions] = @InitialOptions,
								[Options] = @Options,	
								[DefaultAppDedClassTypeIDOrder] = @DefaultAppDedClassTypeIDOrder,
								[IsDeleted] = @IsDeleted   ,
								[DateLastUpdated] = @DateLastUpdated  ,
								[UpdateByUserID] = @UID   
					WHERE 	FeeScheduleID = @FeeScheduleID
					--After Edit Edit the Number OF Levels 
					SET @LevelCount = (SELECT Max(LevelNum) FROM FeeScheduleLevels WHERE FeeScheduleID = @FeeScheduleID)
					IF @LevelCount Is Null 
					BEGIN
						SET @LevelCount = 0
					END
					IF @LevelCount < @NumOfLevels
					BEGIN
						While @LevelCount < @NumOfLevels
						BEGIN
							INSERT INTO FeeScheduleLevels
								(
									[FeeScheduleID],
									[LevelNum],
									[LevelMax],
									[LevelPctApp],
									[LevelMin],
									[IsDeleted],
									[DateLastUpdated],
									[UpdateByUserID]
								)
							SELECT 
									@FeeScheduleID As [FeeScheduleID],
									@LevelCount + 1 As [LevelNum],
									0 As [LevelMax],
									0 As [LevelPctApp],
									0 As [LevelMin],
									0 As [IsDeleted],
									@DateLastUpdated As [DateLastUpdated],
									@UID As [UpdateByUserID]
							SET @LevelCount = @LevelCount + 1 
						END
					END
					ELSE IF @LevelCount > @NumOfLevels
					BEGIN
						While @NumOfLevels < @LevelCount
						BEGIN
							DELETE FROM FeeScheduleLevels
							WHERE 	FeeScheduleID = @FeeScheduleID
							AND 	LevelNum = @LevelCount
							SET @LevelCount = @LevelCount -1
						END
					END
					--After Edit Edit the Number Fee Types 
					SET @FeeTypeCount = (SELECT Max(TypeNum) FROM FeeScheduleFeeTypes WHERE FeeScheduleID = @FeeScheduleID)
					IF @FeeTypeCount Is Null 
					BEGIN
						SET @FeeTypeCount = 0
					END
					IF @FeeTypeCount < @NumOfFeeTypes
					BEGIN
						While @FeeTypeCount < @NumOfFeeTypes
						BEGIN
							INSERT INTO FeeScheduleFeeTypes
								(
									[FeeScheduleID],
									[TypeNum],
									[Name],
									[Description],
									[FeeAmount],
									[IsDeleted],
									[DateLastUpdated],
									[UpdateByUserID]
								)
							SELECT 
									@FeeScheduleID As [FeeScheduleID],
									@FeeTypeCount + 1 As [TypeNum],
									'' As [Name],
									'' As [Description],
									0 As [FeeAmount],
									0 As [IsDeleted],
									@DateLastUpdated As [DateLastUpdated],
									@UID As [UpdateByUserID]
							SET @FeeTypeCount = @FeeTypeCount + 1 
						END
					END
					ELSE IF @FeeTypeCount > @NumOfFeeTypes
					BEGIN
						While @NumOfFeeTypes < @FeeTypeCount
						BEGIN
							DELETE FROM FeeScheduleFeeTypes
							WHERE 	FeeScheduleID = @FeeScheduleID
							AND 	TypeNum = @FeeTypeCount
							SET @FeeTypeCount = @FeeTypeCount -1
						END
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
		SET @sSQL = 'SELECT RetFeeSchedule.' + @GroupBy + ' '
	END
	ELSE
	BEGIN
		SET @sSQL = 'SELECT RetFeeSchedule.* '
	END	
	SET @sSQL = @sSQL + 'FROM( ' 
	SET @sSQL = @sSQL + 'SELECT '
	SET @sSQL = @sSQL + '[FeeScheduleID], '
	SET @sSQL = @sSQL + '(SELECT 	NAME '
	SET @sSQL = @sSQL + 'FROM		Company '
	SET @sSQL = @sSQL + 'WHERE CompanyID = F.[ClientCompanyID]) As ClientCompanyName, '
	SET @sSQL = @sSQL + '[ClientCompanyID] , '
	SET @sSQL = @sSQL + '[ScheduleName], '
	SET @sSQL = @sSQL + '[Description] , '
	SET @sSQL = @sSQL + '[NumOfLevels], '
	SET @sSQL = @sSQL + '[NumOfFeeTypes], '
	SET @sSQL = @sSQL + '[FeeServiceHourlyRate], '
	SET @sSQL = @sSQL + '[TaxPercent] , '
	SET @sSQL = @sSQL + '[InitialOptions], '
	SET @sSQL = @sSQL + '[Options], '
	SET @sSQL = @sSQL + '[DefaultAppDedClassTypeIDOrder], '
	SET @sSQL = @sSQL + '[IsDeleted] , '
	SET @sSQL = @sSQL + '[DateLastUpdated] , '
	SET @sSQL = @sSQL + '(SELECT	USERNAME '
	SET @sSQL = @sSQL + 'FROM 		USERS '
	SET @sSQL = @sSQL + 'WHERE 	USERSID = F.[UpdateByUserID]) As [UpdateByUserName],  '
	SET @sSQL = @sSQL + '[UpdateByUserID]  '

	IF @bViewHistory Is Null
	BEGIN
		SET @sSQL = @sSQL + 'FROM FeeSchedule F '
		SET @sSQL = @sSQL + ') RetFeeSchedule '
		IF @bHideDeleted = 1
		BEGIN
			SET @sSQL = @sSQL + 'WHERE IsDeleted = 0 '
		END
		ELSE
		BEGIN
			SET @sSQL = @sSQL + 'WHERE IsDeleted Is Not Null '
		END
		--Searching
		IF @SearchBy Is Not Null
		BEGIN
			SET @sSQL = @sSQL + @SearchBy + ' '
		END
	END
	ELSE
	BEGIN
		SET @sSQL = @sSQL + 'FROM FeeScheduleHistory F '
		SET @sSQL = @sSQL + ') RetFeeSchedule '
		SET @sSQL = @sSQL + 'WHERE [DateLastUpdated] >= Convert(datetime,''' + @FromDate + ''') '
		SET @sSQL = @sSQL + 'AND [DateLastUpdated] <= Convert(datetime,''' + @ToDate + ''') '
		IF @FeeScheduleid <> -1
		BEGIN
			SET @sSQL = @sSQL + 'AND [FeeScheduleID] = ' + Cast(@FeeScheduleid as varchar(10)) + ' '
		END
		--Searching
		IF @SearchBy Is Not Null
		BEGIN
			SET @sSQL = @sSQL + @SearchBy + ' '
		END
	END
	IF @bViewHistory Is Null
	BEGIN
		--Grouping
		IF @GroupBy Is Not Null
		BEGIN
			SET @sSQL = @sSQL + 'GROUP BY ' + @GroupBy + ' '			
		END
		---Sorting 
		IF @GroupBy Is Null
		BEGIN
			IF @OrderBy Is Null OR @OrderBy = 'Null'
			BEGIN
				SET @sSQL = @sSQL + 'ORDER BY [IsDeleted] Asc, [ScheduleName] '
			END
			ELSE
			BEGIN
				SET @sSQL = @sSQL + 'ORDER BY  [IsDeleted] Asc, ' + @OrderBy + ' '

			END
		END
	END
	ELSE
	BEGIN
		--Grouping
		IF @GroupBy Is Not Null
		BEGIN
			SET @sSQL = @sSQL + 'GROUP BY ' + @GroupBy + ' '			
		END
		---Sorting
		IF @GroupBy Is Null
		BEGIN
			SET @sSQL = @sSQL + 'ORDER BY  ' + @OrderBy + ' '
		END
	END
	EXEC(@sSQL)
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[z_spsGetFeeScheduleInfo]  TO [CFUSER]
GO

