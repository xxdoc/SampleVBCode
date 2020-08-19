SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spsGetDocumentInfo]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spsGetDocumentInfo]
GO

/******************************************************************
*
*    	Description:	Document Info
*			
*	
*	Author: 	Brad Skidmore
*	Date: 		11/17/2003
*
******************************************************************/


CREATE PROCEDURE [dbo].[z_spsGetDocumentInfo]
	@bHideDeleted		bit =1, 
	@UID			int,
	@DocumentID		Int=null,
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
	@DocNameBase		varchar(50)=null,
	@DocName		varchar(50)=null,
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
	@InstallFileLocation	varchar(255)=null,
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

IF @DocumentID Is Not Null And @bViewHistory is Null
BEGIN
	IF @bDelete=0 And @bAdd=0 And @bEdit=0
	BEGIN
		SELECT 
			[DocumentID],
			[DocNameBase],
			[DocName] ,
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
			[InstallFileLocation],
			[SPName],
			[IsDeleted],
			[DateLastUpdated],
			[UpdateByUserID]
		FROM	Document
		WHERE 	[DocumentID] = @DocumentID
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
				UPDATE Document SET 	[IsDeleted]=	(CASE 	WHEN [IsDeleted]= 0
									THEN 1
									ELSE 0
									END
								),
						[DateLastUpdated]=@DateLastUpdated ,
						[UpdateByUserID]=@UID
				WHERE DocumentID = @DocumentID
			END
			ELSE
			BEGIN---See Which Child Table Removing From ...
				IF @SoftwarePackageID is Not Null And @ClientCompanyID Is Not Null
				BEGIN--User Profile 
					UPDATE SoftwarePackageDocument SET 	[IsDeleted] = 1, 
										[DateLastUpdated]= @DateLastUpdated,
										[UpdateByUserID] = @UID 
					FROM 	SoftwarePackageDocument SPD  
					WHERE 	SPD.SoftwarePackageID = @SoftwarePackageID 
					AND 	SPD.DocumentID =@DocumentID 
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
					INSERT INTO Document(
							[DocNameBase],
							[DocName] ,
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
							[InstallFileLocation],
							[SPName],
							[IsDeleted],
							[DateLastUpdated],
							[UpdateByUserID]
							)
					SELECT 			
								@DocNameBase As [DocNameBase],
								@DocName As [DocName] ,
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
								@InstallFileLocation As [InstallFileLocation],
								@SPName As [SPName],
								0 As [IsDeleted] ,
								@DateLastUpdated As [DateLastUpdated],
								@UID As [UpdateByUserID] 
				END
				ELSE
				BEGIN---See Which Child Table Adding to ...
					IF @SoftwarePackageID is Not Null And @ClientCompanyID Is Not Null
					BEGIN--User Assgn To CAT
						SET @iCount = (	SELECT Count([DocumentID]) 
								FROM SoftwarePackageDocument
								WHERE [DocumentID] = @DocumentID
								AND  [SoftwarePackageID] = @SoftwarePackageID
								)
						IF @iCount = 0 
						BEGIN--Since its not in there yet Insert it
							INSERT INTO SoftwarePackageDocument (
												[DocumentID] ,
												[SoftWarePackageID] ,
												[IsDeleted],
												[DateLastUpdated] ,
												[UpdateByUserID] 
											)
							SELECT					@DocumentID As [DocumentID], 
												@SoftWarePackageID As [SoftWarePackageID] ,
												0 As [IsDeleted] ,
												@DateLastUpdated As [DateLastUpdated], 
												@UID As [UpdateByUserID]
						END
						ELSE	
						BEGIN--Its already there make sure it is flagged as Active
							UPDATE SoftwarePackageDocument SET 	[IsDeleted] = 0,
												[DateLastUpdated] = @DateLastUpdated,
												[UpdateByUserID] = @UID
							FROM SoftwarePackageDocument SPD 
							WHERE 	SPD.SoftwarePackageID = @SoftwarePackageID 
							AND 	SPD.DocumentID =@DocumentID  	
						END
					END 
				END
			END
			ELSE 
			BEGIN
				IF @bEdit=1
				BEGIN
					Update Document Set 	
									[DocNameBase] = @DocNameBase,
									[DocName] = @DocName ,
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
									[InstallFileLocation] = @InstallFileLocation,
									[SPName] = @SPName,
									[IsDeleted] = [IsDeleted],
									[DateLastUpdated] = @DateLastUpdated  ,
									[UpdateByUserID] = @UID   
					WHERE 	DocumentID = @DocumentID
				END
			END
		END
	END
END
ELSE
BEGIN
	IF @GroupBy Is Not Null
	BEGIN
		SET @sSQL = 'SELECT RetDocument.' + @GroupBy + ' '
	END
	ELSE
	BEGIN
		SET @sSQL = 'SELECT RetDocument.* '
	END	
	SET @sSQL = @sSQL + 'FROM( ' 
	SET @sSQL = @sSQL + 'SELECT '
	SET @sSQL = @sSQL + 'D.[DocumentID], '
	SET @sSQL = @sSQL + 'D.[DocNameBase], '
	SET @sSQL = @sSQL + 'D.[DocName] , '
	SET @sSQL = @sSQL + 'D.[Description], '
	SET @sSQL = @sSQL + 'D.[Version] , '
	SET @sSQL = @sSQL + 'D.[SPVersionBase], '
	SET @sSQL = @sSQL + 'D.[SPVersion] , '
	SET @sSQL = @sSQL + 'D.[VersionDate], '
	SET @sSQL = @sSQL + 'D.[SectionLevel01], '
	SET @sSQL = @sSQL + 'D.[SectionLevel02], '
	SET @sSQL = @sSQL + 'D.[SectionLevel03], '
	SET @sSQL = @sSQL + 'D.[SectionLevel04], '
	SET @sSQL = @sSQL + 'D.[SectionLevel05], '
	SET @sSQL = @sSQL + 'D.[InstallFileLocation], '
	SET @sSQL = @sSQL + 'D.[SPName], '
	SET @sSQL = @sSQL + 'D.[IsDeleted] , '
--*******************************************************************************
--*			Begin Document SoftwarePackage List			*
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
			SET @sSQL = @sSQL + '(SELECT Count(SPD.SoftWarePackageID) as CountOfSoftWarePackageID '	
			SET @sSQL = @sSQL + 'FROM SoftwarePackageDocument SPD '
			SET @sSQL = @sSQL + 'WHERE SPD.DocumentID = D.[DocumentID] '
			SET @sSQL = @sSQL + 'AND SPD.SoftWarePackageID = ' + Cast( @SPSoftWarePackageID as Varchar(10)) + ' '
			SET @sSQL = @sSQL + 'AND SPD.IsDeleted = 0 ' 
			SET @sSQL = @sSQL + 'GROUP BY SPD.SoftWarePackageID '
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
--*			End Document SoftwarePackage List			*
--*										*
--*******************************************************************************
	SET @sSQL = @sSQL + 'D.[DateLastUpdated] , '
	SET @sSQL = @sSQL + '(SELECT	USERNAME '
	SET @sSQL = @sSQL + 'FROM 		USERS '
	SET @sSQL = @sSQL + 'WHERE 	USERSID = D.[UpdateByUserID]) As [UpdateByUserName],  '
	SET @sSQL = @sSQL + 'D.[UpdateByUserID]  '

	IF @bViewHistory Is Null
	BEGIN
		SET @sSQL = @sSQL + 'FROM Document D '
		IF @ClientCompanyID Is not Null And @SoftwarePackageID Is not Null
		Begin
			SET @sSQL = @sSQL + 'INNER JOIN SoftwarePackageDocument SPD '
			SET @sSQL = @sSQL + 'ON D.DocumentID = SPD.DocumentID '
			SET @sSQL = @sSQL + 'INNER JOIN SoftwarePackage SP ON SPD.SoftwarePackageID = SP.SoftwarePackageID '
			SET @sSQL = @sSQL + 'AND SP.ClientCompanyID = ' + Cast(@ClientCompanyID  as varchar(10)) + ' '
			SET @sSQL = @sSQL + 'AND SP.SoftwarePackageID = ' + Cast(@SoftwarePackageID  as varchar(10)) + ' '
			SET @sSQL = @sSQL + 'AND SPD.IsDeleted = 0 '
			SET @sSQL = @sSQL + 'AND SP.IsDeleted = 0 '
		END
		SET @sSQL = @sSQL + ') RetDocument '
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
		SET @sSQL = @sSQL + 'FROM DocumentHistory D '
		SET @sSQL = @sSQL + ') RetDocument '
		SET @sSQL = @sSQL + 'WHERE [DateLastUpdated] >= Convert(datetime,''' + @FromDate + ''') '
		SET @sSQL = @sSQL + 'AND [DateLastUpdated] <= Convert(datetime,''' + @ToDate + ''') '
		IF @DocumentID <> -1
		BEGIN
			SET @sSQL = @sSQL + 'AND [DocumentID] = ' + Cast(@DocumentID as varchar(10)) + ' '
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
				SET @sSQL = @sSQL + 'ORDER BY [IsDeleted] Asc, [DocNameBase] '
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

GRANT  EXECUTE  ON [dbo].[z_spsGetDocumentInfo]  TO [CFUSER]
GO

