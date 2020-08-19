SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spsGetApplicationInfo]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spsGetApplicationInfo]
GO

/******************************************************************
*
*    	Description:	Application Info
*			
*	
*	Author: 	Brad Skidmore
*	Date: 		11/13/2003
*
******************************************************************/


CREATE PROCEDURE [dbo].[z_spsGetApplicationInfo]
	@bHideDeleted		bit =1, 
	@UID			int,
	@ApplicationID		int=null,
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
	@AppNameBase		varchar(50)=null,
	@AppName		varchar(50)=null,
	@Description		varchar(500)=null,
	@Version		int=null,
	@MajorVS		int=null,
	@MinorVS		int=null,
	@RevisionVS		int=null,
	@SPVersionBase		int=null,
	@SPVersion		int=null,
	@VersionDate		datetime=null,
	@ProjectName		varchar(50)=null,
	@ClassName		varchar(50)=null,
	@SectionLevel01		varchar(50)=null,
	@SectionLevel02		varchar(50)=null,
	@SectionLevel03		varchar(50)=null,
	@SectionLevel04		varchar(50)=null,
	@SectionLevel05		varchar(50)=null,
	@InstallFileLocation	varchar(255)=null,
	@SPName			varchar(50)=null,
	@SELF_REG		bit=null,
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

IF @ApplicationID Is Not Null And @bViewHistory is Null
BEGIN
	IF @bDelete=0 And @bAdd=0 And @bEdit=0
	BEGIN
		SELECT 
			[ApplicationID],
			[AppNameBase],
			[AppName] ,
			[Description],
			[Version] ,
			[MajorVS] ,
			[MinorVS] ,
			[RevisionVS],
			[SPVersionBase],
			[SPVersion] ,
			[VersionDate],
			[ProjectName],
			[ClassName] ,
			[SectionLevel01],
			[SectionLevel02],
			[SectionLevel03],
			[SectionLevel04],
			[SectionLevel05],
			[InstallFileLocation],
			[SPName],
			[SELF_REG],
			[IsDeleted],
			[DateLastUpdated],
			[UpdateByUserID]
		FROM	Application
		WHERE 	[ApplicationID] = @ApplicationID
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
				UPDATE Application SET 	[IsDeleted]=	(CASE 	WHEN [IsDeleted]= 0
										THEN 1
										ELSE 0
										END
									),
							[DateLastUpdated]=@DateLastUpdated ,
							[UpdateByUserID]=@UID
				WHERE ApplicationID = @ApplicationID
			END
			ELSE
			BEGIN---See Which Child Table Removing From ...
				IF @SoftwarePackageID is Not Null And @ClientCompanyID Is Not Null
				BEGIN--User Profile 
					UPDATE SoftwarePackageApplication SET 	[IsDeleted] = 1, 
										[DateLastUpdated]= @DateLastUpdated,
										[UpdateByUserID] = @UID 
					FROM 	SoftwarePackageApplication SPA  
					WHERE 	SPA.SoftwarePackageID = @SoftwarePackageID 
					AND 	SPA.ApplicationID =@ApplicationID 
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
					INSERT INTO Application(
								[AppNameBase],
								[AppName] ,
								[Description],
								[Version] ,
								[MajorVS] ,
								[MinorVS] ,
								[RevisionVS],
								[SPVersionBase],
								[SPVersion] ,
								[VersionDate],
								[ProjectName],
								[ClassName] ,
								[SectionLevel01],	
								[SectionLevel02],	
								[SectionLevel03],	
								[SectionLevel04],	
								[SectionLevel05],	
								[InstallFileLocation],
								[SPName],
								[SELF_REG],
								[IsDeleted],
								[DateLastUpdated],
								[UpdateByUserID]
							)
					SELECT 			
								@AppNameBase As [AppNameBase],
								@AppName As [AppName] ,
								@Description As [Description],
								@Version As [Version] ,
								@MajorVS As [MajorVS] ,
								@MinorVS As [MinorVS] ,
								@RevisionVS As [RevisionVS],
								@SPVersionBase As [SPVersionBase],
								@SPVersion As [SPVersion] ,
								@VersionDate As [VersionDate],
								@ProjectName As [ProjectName],
								@ClassName As [ClassName] ,
								@SectionLevel01	As [SectionLevel01],
								@SectionLevel02	As [SectionLevel02],	
								@SectionLevel03	As [SectionLevel03],	
								@SectionLevel04	As [SectionLevel04],	
								@SectionLevel05	As [SectionLevel05],	
								@InstallFileLocation As [InstallFileLocation],
								@SPName As [SPName],
								@SELF_REG As [SELF_REG],
								0 As [IsDeleted] ,
								@DateLastUpdated As [DateLastUpdated],
								@UID As [UpdateByUserID] 
				END
				ELSE
				BEGIN---See Which Child Table Adding to ...
					IF @SoftwarePackageID is Not Null And @ClientCompanyID Is Not Null
					BEGIN--User Assgn To CAT
						SET @iCount = (	SELECT Count([ApplicationID]) 
								FROM SoftwarePackageApplication
								WHERE [ApplicationID] = @ApplicationID
								AND  [SoftwarePackageID] = @SoftwarePackageID
								)
						IF @iCount = 0 
						BEGIN--Since its not in there yet Insert it
							INSERT INTO SoftwarePackageApplication (
												[ApplicationID] ,
												[SoftWarePackageID] ,
												[IsDeleted],
												[DateLastUpdated] ,
												[UpdateByUserID] 
											)
							SELECT					@ApplicationID As [ApplicationID], 
												@SoftWarePackageID As [SoftWarePackageID] ,
												0 As [IsDeleted] ,
												@DateLastUpdated As [DateLastUpdated], 
												@UID As [UpdateByUserID]
						END
						ELSE	
						BEGIN--Its already there make sure it is flagged as Active
							UPDATE SoftwarePackageApplication SET 	[IsDeleted] = 0,
												[DateLastUpdated] = @DateLastUpdated,
												[UpdateByUserID] = @UID
							FROM SoftwarePackageApplication SPA 
							WHERE 	SPA.SoftwarePackageID = @SoftwarePackageID 
							AND 	SPA.ApplicationID =@ApplicationID  	
						END
					END 
				END
			END
			ELSE 
			BEGIN
				IF @bEdit=1
				BEGIN
					Update Application Set 	
									[AppNameBase] = @AppNameBase,
									[AppName] = @AppName ,
									[Description] = @Description,
									[Version] = @Version ,
									[MajorVS] = @MajorVS ,
									[MinorVS] = @MinorVS ,
									[RevisionVS] = @RevisionVS,
									[SPVersionBase] = @SPVersionBase,
									[SPVersion] = @SPVersion ,
									[VersionDate] = @VersionDate,
									[ProjectName] = @ProjectName,
									[ClassName] = @ClassName ,
									[SectionLevel01] = @SectionLevel01,
									[SectionLevel02] = @SectionLevel02,
									[SectionLevel03] = @SectionLevel03,
									[SectionLevel04] = @SectionLevel04,
									[SectionLevel05] = @SectionLevel05,
									[InstallFileLocation] = @InstallFileLocation,
									[SPName] = @SPName,
									[SELF_REG] = @SELF_REG,
									[IsDeleted] = [IsDeleted],
									[DateLastUpdated] = @DateLastUpdated  ,
									[UpdateByUserID] = @UID   
					WHERE 	ApplicationID = @ApplicationID
				END
			END
		END
	END
END
ELSE
BEGIN
	IF @GroupBy Is Not Null
	BEGIN
		SET @sSQL = 'SELECT RetApplication.' + @GroupBy + ' '
	END
	ELSE
	BEGIN
		SET @sSQL = 'SELECT RetApplication.* '
	END	
	SET @sSQL = @sSQL + 'FROM( ' 
	SET @sSQL = @sSQL + 'SELECT '
	SET @sSQL = @sSQL + 'A.[ApplicationID], '
	SET @sSQL = @sSQL + 'A.[AppNameBase], '
	SET @sSQL = @sSQL + 'A.[AppName] , '
	SET @sSQL = @sSQL + 'A.[Description], '
	SET @sSQL = @sSQL + 'A.[Version] , '
	SET @sSQL = @sSQL + 'A.[MajorVS] , '
	SET @sSQL = @sSQL + 'A.[MinorVS] , '
	SET @sSQL = @sSQL + 'A.[RevisionVS], '
	SET @sSQL = @sSQL + 'A.[SPVersionBase], '
	SET @sSQL = @sSQL + 'A.[SPVersion] , '
	SET @sSQL = @sSQL + 'A.[VersionDate], '
	SET @sSQL = @sSQL + 'A.[ProjectName], '
	SET @sSQL = @sSQL + 'A.[ClassName] , '
	SET @sSQL = @sSQL + 'A.[SectionLevel01], '
	SET @sSQL = @sSQL + 'A.[SectionLevel02], '
	SET @sSQL = @sSQL + 'A.[SectionLevel03], '
	SET @sSQL = @sSQL + 'A.[SectionLevel04], '
	SET @sSQL = @sSQL + 'A.[SectionLevel05], '
	SET @sSQL = @sSQL + 'A.[InstallFileLocation], '
	SET @sSQL = @sSQL + 'A.[SPName], '
	SET @sSQL = @sSQL + 'A.[SELF_REG], '
	SET @sSQL = @sSQL + 'A.[IsDeleted] , '
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
			SET @sSQL = @sSQL + '(SELECT Count(SPA.SoftWarePackageID) as CountOfSoftWarePackageID '	
			SET @sSQL = @sSQL + 'FROM SoftwarePackageApplication SPA '
			SET @sSQL = @sSQL + 'WHERE SPA.ApplicationID = A.[ApplicationID] '
			SET @sSQL = @sSQL + 'AND SPA.SoftWarePackageID = ' + Cast( @SPSoftWarePackageID as Varchar(10)) + ' '
			SET @sSQL = @sSQL + 'AND SPA.IsDeleted = 0 ' 
			SET @sSQL = @sSQL + 'GROUP BY SPA.SoftWarePackageID '
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
	SET @sSQL = @sSQL + 'A.[DateLastUpdated] , '
	SET @sSQL = @sSQL + '(SELECT	USERNAME '
	SET @sSQL = @sSQL + 'FROM 		USERS '
	SET @sSQL = @sSQL + 'WHERE 	USERSID = A.[UpdateByUserID]) As [UpdateByUserName],  '
	SET @sSQL = @sSQL + 'A.[UpdateByUserID]  '

	IF @bViewHistory Is Null
	BEGIN
		SET @sSQL = @sSQL + 'FROM Application A '
		IF @ClientCompanyID Is not Null And @SoftwarePackageID Is not Null
		Begin
			SET @sSQL = @sSQL + 'INNER JOIN SoftwarePackageApplication SPA '
			SET @sSQL = @sSQL + 'ON A.ApplicationID = SPA.ApplicationID '
			SET @sSQL = @sSQL + 'INNER JOIN SoftwarePackage SP ON SPA.SoftwarePackageID = SP.SoftwarePackageID '
			SET @sSQL = @sSQL + 'AND SP.ClientCompanyID = ' + Cast(@ClientCompanyID  as varchar(10)) + ' '
			SET @sSQL = @sSQL + 'AND SP.SoftwarePackageID = ' + Cast(@SoftwarePackageID  as varchar(10)) + ' '
			SET @sSQL = @sSQL + 'AND SPA.IsDeleted = 0 '
			SET @sSQL = @sSQL + 'AND SP.IsDeleted = 0 '

		END
		SET @sSQL = @sSQL + ') RetApplication '
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
		SET @sSQL = @sSQL + 'FROM ApplicationHistory A '
		SET @sSQL = @sSQL + ') RetApplication '
		SET @sSQL = @sSQL + 'WHERE [DateLastUpdated] >= Convert(datetime,''' + @FromDate + ''') '
		SET @sSQL = @sSQL + 'AND [DateLastUpdated] <= Convert(datetime,''' + @ToDate + ''') '
		IF @ApplicationID <> -1
		BEGIN
			SET @sSQL = @sSQL + 'AND [ApplicationID] = ' + Cast(@ApplicationID as varchar(10)) + ' '
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
				SET @sSQL = @sSQL + 'ORDER BY [IsDeleted] Asc, [AppNameBase] '
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
	print(@sSQL)
	EXEC(@sSQL)
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[z_spsGetApplicationInfo]  TO [CFUSER]
GO

