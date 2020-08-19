SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spsGetSoftwarePackageInfo]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spsGetSoftwarePackageInfo]
GO

/******************************************************************
*
*    	Description:	Software Package Info
*			
*	
*	Author: 	Brad Skidmore
*	Date: 		11/11/2003
*
******************************************************************/


CREATE PROCEDURE [dbo].[z_spsGetSoftwarePackageInfo]
	@bHideDeleted		bit =1, 
	@UID			int,
	@SoftwarePackageID	Int=null,
	@OrderBy		varchar(500)=null,
	@GroupBy		varchar(500)=null,
	@SearchBy		varchar(1000)=null,
--updating the TypeOfLoss Table
	@bDelete		bit=0,
	@bAdd			bit=0,
	@bEdit			bit=0,
--Update variables
	@ClientCompanyID	int=null ,
	@CATID			int=null,
	@PackageName 		varchar(50)=null ,
	@Description		varchar(500)=null ,
	@SPVersion		int=null,
	@VersionDate		datetime =null,
	@DateLastUpdated 	datetime =null,
--ViewHistory
	@bViewHistory		bit=NULL,
	@FromVersionDate	varchar(50)='*',
	@ToVersionDate		varchar(50)='*',
	@FromDate		varchar(50)=Null,
	@ToDate			varchar(50)=Null
	
AS
DECLARE @sSQL Varchar(8000)
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

IF @SoftwarePackageID Is Not Null And @bViewHistory Is Null
BEGIN
	IF @bDelete=0 And @bAdd=0 And @bEdit=0
	BEGIN
		SELECT 
			[SoftwarePackageID],
			[ClientCompanyID] ,
			[CATID] ,
			[PackageName] ,
			[Description],
			[SPVersion],
			[VersionDate],
			[IsDeleted]
			[DateLastUpdated],
			[UpdateByUserID]
		FROM	SoftwarePackage
		WHERE 	[SoftwarePackageID] = @SoftwarePackageID
	END
	ELSE
	BEGIN
		IF @bDelete =1
		BEGIN
			UPDATE SoftwarePackage SET 	[IsDeleted]=	(CASE 	WHEN [IsDeleted]= 0
									THEN 1
									ELSE 0
									END
								),
						[DateLastUpdated]=@DateLastUpdated ,
						[UpdateByUserID]=@UID
			WHERE SoftwarePackageID = @SoftwarePackageID
		
		END
		ELSE 
		BEGIN
			IF @bAdd=1
			BEGIN
				INSERT INTO SoftwarePackage(
							[ClientCompanyID] ,
							[CATID] ,
							[PackageName] ,
							[Description],
							[SPVersion],
							[VersionDate],
							[IsDeleted],
							[DateLastUpdated],
							[UpdateByUserID]
						)
				SELECT 			
							@ClientCompanyID As [ClientCompanyID] ,
							@CATID As [CATID] ,
							@PackageName As [PackageName] ,
							@Description As [Description],
							@SPVersion As [SPVersion],
							@VersionDate As [VersionDate],
							0 As [IsDeleted] ,
							@DateLastUpdated As [DateLastUpdated],
							@UID As [UpdateByUserID] 
					
			END
			ELSE 
			BEGIN
				IF @bEdit=1
				BEGIN
					Update SoftwarePackage Set 	
									[ClientCompanyID] = @ClientCompanyID  ,
									[CATID] = @CATID  ,
									[PackageName] = @PackageName  ,
									[Description] = @Description ,
									[SPVersion] = @SPVersion,
									[VersionDate] = @VersionDate,
									[IsDeleted] = [IsDeleted],
									[DateLastUpdated] = @DateLastUpdated  ,
									[UpdateByUserID] = @UID   
					WHERE 	SoftwarePackageID = @SoftwarePackageID
				END
			END
		END
	END
END
ELSE
BEGIN
	IF @GroupBy Is Not Null
	BEGIN
		SET @sSQL = 'SELECT RetSoftwarePackage.' + @GroupBy + ' '
	END
	ELSE
	BEGIN
		SET @sSQL = 'SELECT RetSoftwarePackage.* '
	END	
	SET @sSQL = @sSQL + 'FROM( ' 
	SET @sSQL = @sSQL + 'SELECT '
	SET @sSQL = @sSQL + '[SoftwarePackageID], '
	SET @sSQL = @sSQL + '(SELECT 	NAME '
	SET @sSQL = @sSQL + 'FROM		Company '
	SET @sSQL = @sSQL + 'WHERE CompanyID = S.[ClientCompanyID]) As ClientCompanyName, '
	SET @sSQL = @sSQL + '[ClientCompanyID], '
	SET @sSQL = @sSQL + '(SELECT 	NAME '
	SET @sSQL = @sSQL + 'FROM	Cat '
	SET @sSQL = @sSQL + 'WHERE CatID = S.[CATID]) As CatName, '
	SET @sSQL = @sSQL + '[CATID], '
	SET @sSQL = @sSQL + '[PackageName], '
	SET @sSQL = @sSQL + '[Description], '
	SET @sSQL = @sSQL + '[SPVersion], '
	SET @sSQL = @sSQL + '[VersionDate], '
	SET @sSQL = @sSQL + '[IsDeleted] , '
	SET @sSQL = @sSQL + '[DateLastUpdated] , '
	SET @sSQL = @sSQL + '(SELECT	USERNAME '
	SET @sSQL = @sSQL + 'FROM 		USERS '
	SET @sSQL = @sSQL + 'WHERE 	USERSID = S.[UpdateByUserID]) As [UpdateByUserName],  '
	SET @sSQL = @sSQL + '[UpdateByUserID]  '

	IF @bViewHistory Is Null
	BEGIN
		SET @sSQL = @sSQL + 'FROM SoftwarePackage S '
		SET @sSQL = @sSQL + ') RetSoftwarePackage '
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
		SET @sSQL = @sSQL + 'FROM SoftwarePackageHistory S '
		SET @sSQL = @sSQL + ') RetSoftwarePackage '
		SET @sSQL = @sSQL + 'WHERE [DateLastUpdated] >= Convert(datetime,''' + @FromDate + ''') '
		SET @sSQL = @sSQL + 'AND [DateLastUpdated] <= Convert(datetime,''' + @ToDate + ''') '
		IF @SoftwarePackageID <> -1
		BEGIN
			SET @sSQL = @sSQL + 'AND [SoftwarePackageID] = ' + Cast(@SoftwarePackageID as varchar(10)) + ' '
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
				SET @sSQL = @sSQL + 'ORDER BY [IsDeleted] Asc, [CatName] '
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
			IF @OrderBy Is Null OR @OrderBy = 'Null'
			BEGIN
				SET @sSQL = @sSQL + 'ORDER BY [IsDeleted] Asc, [CatName] '
			END
			ELSE
			BEGIN
				SET @sSQL = @sSQL + 'ORDER BY  ' + @OrderBy + ' '
			END
		END
	END
	--Print (@ssql)
	EXEC(@sSQL)
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[z_spsGetSoftwarePackageInfo]  TO [CFUSER]
GO

