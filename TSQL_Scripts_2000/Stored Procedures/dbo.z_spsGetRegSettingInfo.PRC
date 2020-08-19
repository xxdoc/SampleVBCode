SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spsGetRegSettingInfo]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spsGetRegSettingInfo]
GO

/******************************************************************
*
*    	Description:	RegSetting Info
*			
*	
*	Author: 	Brad Skidmore
*	Date: 		11/17/2003
*
******************************************************************/


CREATE PROCEDURE [dbo].[z_spsGetRegSettingInfo]
	@bHideDeleted		bit =1, 
	@UID			int,
	@RegSettingID		Int=null,
	@ClientCompanyID	int=null,
	@SoftwarePackageID	int=null,
	@OrderBy		varchar(500)=null,
	@GroupBy		varchar(500)=null,
	@SearchBy		varchar(1000)=null,
--updating the Table
	@bDelete		bit=0,
	@bAdd			bit=0,
	@bEdit			bit=0,
--Update variables
	@RegNameBase		varchar(50)=null,
	@RegName		varchar(50)=null,
	@Description		varchar(500)=null,
	@Version		int=null,
	@SPVersionBase		int=null,
	@SPVersion		int=null,
	@VersionDate		datetime=null,
	@SectionLevel01		varchar(50)=null,
	@SectionLevel02		varchar(50)=null,
	@SectionLevel03		varchar(50)=null,
	@SectionLevel04		varchar(50)=null,
	@SectionLevel05		varchar(50)=null,
	@SPName			varchar(50)=null,
	@DateLastUpdated 	datetime =null,
--ViewHistory
	@bViewHistory		bit=NULL,
	@FromVersionDate	varchar(50)='*',
	@ToVersionDate		varchar(50)='*',
	@FromDate		varchar(50)=Null,
	@ToDate			varchar(50)=Null
	
AS
DECLARE @sSQL Varchar(8000)
DECLARE @iCount int
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
--Clean String Dates
SET @FromVersionDate = 		dbo.CleanFromOrToDateString(@CurDate, @FromVersionDate, 1)
SET @ToVersionDate = 		dbo.CleanFromOrToDateString(@CurDate, @ToVersionDate, 0)

IF @RegSettingID Is Not Null And @bViewHistory is Null
BEGIN
	IF @bDelete=0 And @bAdd=0 And @bEdit=0
	BEGIN
		SELECT 
			[RegSettingID],
			[RegNameBase],
			[RegName] ,
			[Description],
			[Version] ,
			[SPVersionBase],
			[SPVersion] ,
			[VersionDate],
			[SectionLevel01],
			[SectionLevel02],
			[SectionLevel03],
			[SectionLevel04],
			[SectionLevel05],
			[SPName],
			[IsDeleted],
			[DateLastUpdated],
			[UpdateByUserID]
		FROM	RegSetting
		WHERE 	[RegSettingID] = @RegSettingID
	END
	ELSE
	BEGIN
/****************************************
*		***DELETE***
****************************************/
		IF @bDelete =1
		BEGIN
			IF @SoftwarePackageID Is Null 
			BEGIN
				UPDATE RegSetting SET 	[IsDeleted]=	(CASE 	WHEN [IsDeleted]= 0
									THEN 1
									ELSE 0
									END
								),
						[DateLastUpdated]=@DateLastUpdated ,
						[UpdateByUserID]=@UID
				WHERE RegSettingID = @RegSettingID
			END
			ELSE
			BEGIN---See Which Child Table Removing From ...
				IF @SoftwarePackageID is Not Null And @ClientCompanyID Is Not Null
				BEGIN--User Profile 
					UPDATE SoftwarePackageRegSetting SET 	[IsDeleted] = 1, 
										[DateLastUpdated]= @DateLastUpdated,
										[UpdateByUserID] = @UID 
					FROM 	SoftwarePackageRegSetting SPR  
					WHERE 	SPR.SoftwarePackageID = @SoftwarePackageID 
					AND 	SPR.RegSettingID =@RegSettingID 
				END
			END
		END
		ELSE 
		BEGIN
/****************************************
*		***ADD***
****************************************/
			IF @bAdd=1
			BEGIN
				IF @SoftwarePackageID Is Null 
				BEGIN
					INSERT INTO RegSetting(
							[RegNameBase],
							[RegName] ,
							[Description],
							[Version] ,
							[SPVersionBase],
							[SPVersion] ,
							[VersionDate],
							[SectionLevel01],	
							[SectionLevel02],	
							[SectionLevel03],	
							[SectionLevel04],	
							[SectionLevel05],	
							[SPName],
							[IsDeleted],
							[DateLastUpdated],
							[UpdateByUserID]
							)
					SELECT 			
								@RegNameBase As [RegNameBase],
								@RegName As [RegName] ,
								@Description As [Description],
								@Version As [Version] ,
								@SPVersionBase As [SPVersionBase],
								@SPVersion As [SPVersion] ,
								@VersionDate As [VersionDate],
								@SectionLevel01	As [SectionLevel01],
								@SectionLevel02	As [SectionLevel02],	
								@SectionLevel03	As [SectionLevel03],	
								@SectionLevel04	As [SectionLevel04],	
								@SectionLevel05	As [SectionLevel05],
								@SPName As [SPName],
								0 As [IsDeleted] ,
								@DateLastUpdated As [DateLastUpdated],
								@UID As [UpdateByUserID] 
				END
				ELSE
				BEGIN---See Which Child Table Adding to ...
					IF @SoftwarePackageID is Not Null And @ClientCompanyID Is Not Null
					BEGIN--User Assgn To CAT
						SET @iCount = (	SELECT Count([RegSettingID]) 
								FROM SoftwarePackageRegSetting
								WHERE [RegSettingID] = @RegSettingID
								AND  [SoftwarePackageID] = @SoftwarePackageID
								)
						IF @iCount = 0 
						BEGIN--Since its not in there yet Insert it
							INSERT INTO SoftwarePackageRegSetting (
												[RegSettingID] ,
												[SoftWarePackageID] ,
												[IsDeleted],
												[DateLastUpdated] ,
												[UpdateByUserID] 
											)
							SELECT					@RegSettingID As [RegSettingID], 
												@SoftWarePackageID As [SoftWarePackageID] ,
												0 As [IsDeleted] ,
												@DateLastUpdated As [DateLastUpdated], 
												@UID As [UpdateByUserID]
						END
						ELSE	
						BEGIN--Its already there make sure it is flagged as Active
							UPDATE SoftwarePackageRegSetting SET 	[IsDeleted] = 0,
												[DateLastUpdated] = @DateLastUpdated,
												[UpdateByUserID] = @UID
							FROM SoftwarePackageRegSetting SPR 
							WHERE 	SPR.SoftwarePackageID = @SoftwarePackageID 
							AND 	SPR.RegSettingID =@RegSettingID  	
						END
					END 
				END
			END
			ELSE 
			BEGIN
				IF @bEdit=1
				BEGIN
					Update RegSetting Set 	
									[RegNameBase] = @RegNameBase,
									[RegName] = @RegName ,
									[Description] = @Description,
									[Version] = @Version ,
									[SPVersionBase] = @SPVersionBase,
									[SPVersion] = @SPVersion ,
									[VersionDate] = @VersionDate,
									[SectionLevel01] = @SectionLevel01,
									[SectionLevel02] = @SectionLevel02,
									[SectionLevel03] = @SectionLevel03,
									[SectionLevel04] = @SectionLevel04,
									[SectionLevel05] = @SectionLevel05,
									[SPName] = @SPName,
									[IsDeleted] = [IsDeleted],
									[DateLastUpdated] = @DateLastUpdated  ,
									[UpdateByUserID] = @UID   
					WHERE 	RegSettingID = @RegSettingID
				END
			END
		END
	END
END
ELSE
BEGIN
	IF @GroupBy Is Not Null
	BEGIN
		SET @sSQL = 'SELECT RetRegSetting.' + @GroupBy + ' '
	END
	ELSE
	BEGIN
		SET @sSQL = 'SELECT RetRegSetting.* '
	END	
	SET @sSQL = @sSQL + 'FROM( ' 
	SET @sSQL = @sSQL + 'SELECT '
	SET @sSQL = @sSQL + 'R.[RegSettingID], '
	SET @sSQL = @sSQL + 'R.[RegNameBase], '
	SET @sSQL = @sSQL + 'R.[RegName] , '
	SET @sSQL = @sSQL + 'R.[Description], '
	SET @sSQL = @sSQL + 'R.[Version] , '
	SET @sSQL = @sSQL + 'R.[SPVersionBase], '
	SET @sSQL = @sSQL + 'R.[SPVersion] , '
	SET @sSQL = @sSQL + 'R.[VersionDate], '
	SET @sSQL = @sSQL + 'R.[SectionLevel01], '
	SET @sSQL = @sSQL + 'R.[SectionLevel02], '
	SET @sSQL = @sSQL + 'R.[SectionLevel03], '
	SET @sSQL = @sSQL + 'R.[SectionLevel04], '
	SET @sSQL = @sSQL + 'R.[SectionLevel05], '
	SET @sSQL = @sSQL + 'R.[SPName], '
	SET @sSQL = @sSQL + 'R.[IsDeleted] , '
--*******************************************************************************
--*			Begin Application SoftwarePackage List			*
--*										*
--VVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV
	IF @ClientCompanyID Is Not Null And @bViewHistory Is Null
	BEGIN
		DECLARE @SPSoftWarePackageID int
		DECLARE @SPClientCompanyID int
		DECLARE @SPPackageName varchar(50)
		DEClARE @SPDescription varchar(500)
		IF @bHideDeleted =1 
		BEGIN
			Declare CurSP Cursor for
			SELECT 
				SP.[SoftWarePackageID],
				SP.[ClientCompanyID],
				SP.[PackageName],
				SP.[Description]
			FROM 	SoftwarePackage SP
			WHERE 	SP.IsDeleted = 0
			AND 	SP.[ClientCompanyID] = @ClientCompanyID
			ORDER BY SP.[PackageName]
		END
		ELSE
		BEGIN
			Declare CurSP Cursor for
			SELECT 
				SP.[SoftWarePackageID],
				SP.[ClientCompanyID],
				SP.[PackageName],
				SP.[Description]
			FROM 	SoftwarePackage SP 
			WHERE 	SP.[ClientCompanyID] = @ClientCompanyID
			ORDER BY SP.[PackageName]
		END
		Open CurSP
		Fetch Next From CurSP into 			
									@SPSoftWarePackageID ,
									@SPClientCompanyID,
									@SPPackageName ,
									@SPDescription
									
		While @@FEtch_Status = 0
		BEGIN
			SET @sSQL = @sSQL + '(SELECT Count(SPR.SoftWarePackageID) as CountOfSoftWarePackageID '	
			SET @sSQL = @sSQL + 'FROM SoftwarePackageRegSetting SPR '
			SET @sSQL = @sSQL + 'WHERE SPR.RegSettingID = R.[RegSettingID] '
			SET @sSQL = @sSQL + 'AND SPR.SoftWarePackageID = ' + Cast( @SPSoftWarePackageID as Varchar(10)) + ' '
			SET @sSQL = @sSQL + 'AND SPR.IsDeleted = 0 ' 
			SET @sSQL = @sSQL + 'GROUP BY SPR.SoftWarePackageID '
			SET @sSQL = @sSQL + ') As [SoftwarePackage' + Cast( @SPSoftWarePackageID as Varchar(10)) + '_' + Cast( @SPClientCompanyID as Varchar(10)) + '], '
			Fetch Next From CurSP into 			
									@SPSoftWarePackageID ,
									@SPClientCompanyID,
									@SPPackageName ,
									@SPDescription
		END
		Close CurSP
		Deallocate CurSP
		
	END
--^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
--*			End Application SoftwarePackage List			*
--*										*
--*******************************************************************************
	SET @sSQL = @sSQL + 'R.[DateLastUpdated] , '
	SET @sSQL = @sSQL + '(SELECT	USERNAME '
	SET @sSQL = @sSQL + 'FROM 		USERS '
	SET @sSQL = @sSQL + 'WHERE 	USERSID = R.[UpdateByUserID]) As [UpdateByUserName],  '
	SET @sSQL = @sSQL + 'R.[UpdateByUserID]  '

	IF @bViewHistory Is Null
	BEGIN
		SET @sSQL = @sSQL + 'FROM RegSetting R '
		IF @ClientCompanyID Is not Null And @SoftwarePackageID Is not Null
		Begin
			SET @sSQL = @sSQL + 'INNER JOIN SoftwarePackageRegSetting SPR '
			SET @sSQL = @sSQL + 'ON R.RegSettingID = SPR.RegSettingID '
			SET @sSQL = @sSQL + 'INNER JOIN SoftwarePackage SP ON SPR.SoftwarePackageID = SP.SoftwarePackageID '
			SET @sSQL = @sSQL + 'AND SP.ClientCompanyID = ' + Cast(@ClientCompanyID  as varchar(10)) + ' '
			SET @sSQL = @sSQL + 'AND SP.SoftwarePackageID = ' + Cast(@SoftwarePackageID  as varchar(10)) + ' '
			SET @sSQL = @sSQL + 'AND SPR.IsDeleted = 0 '
			SET @sSQL = @sSQL + 'AND SP.IsDeleted = 0 '
		END
		SET @sSQL = @sSQL + ') RetRegSetting '
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
		--Searching Dates
		IF @FromVersionDate <> '*'
		BEGIN
			SET @sSQL = @sSQL + 'AND [VersionDate] >= Convert(datetime,''' + @FromVersionDate + ''') '
		END
		IF @ToVersionDate <> '*'
		BEGIN
			SET @sSQL = @sSQL + 'AND [VersionDate] <= Convert(datetime,''' + @ToVersionDate + ''') '
		END
	END	
	ELSE
	BEGIN
		SET @sSQL = @sSQL + 'FROM RegSettingHistory R '
		SET @sSQL = @sSQL + ') RetRegSetting '
		SET @sSQL = @sSQL + 'WHERE [DateLastUpdated] >= Convert(datetime,''' + @FromDate + ''') '
		SET @sSQL = @sSQL + 'AND [DateLastUpdated] <= Convert(datetime,''' + @ToDate + ''') '
		IF @RegSettingID <> -1
		BEGIN
			SET @sSQL = @sSQL + 'AND [RegSettingID] = ' + Cast(@RegSettingID as varchar(10)) + ' '
		END
		--Searching
		IF @SearchBy Is Not Null
		BEGIN
			SET @sSQL = @sSQL + @SearchBy + ' '
		END
		--Searching Dates
		IF @FromVersionDate <> '*'
		BEGIN
			SET @sSQL = @sSQL + 'AND [VersionDate] >= Convert(datetime,''' + @FromVersionDate + ''') '
		END
		IF @ToVersionDate <> '*'
		BEGIN
			SET @sSQL = @sSQL + 'AND [VersionDate] <= Convert(datetime,''' + @ToVersionDate + ''') '
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
				SET @sSQL = @sSQL + 'ORDER BY [IsDeleted] Asc, [RegNameBase] '
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

GRANT  EXECUTE  ON [dbo].[z_spsGetRegSettingInfo]  TO [CFUSER]
GO

