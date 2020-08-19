SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spsGetCompanyInfo]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spsGetCompanyInfo]
GO

/******************************************************************
*
*    	Description:	Company Info
*			
*	
*	Author: 	Brad Skidmore
*	Date: 		09/8/2003
*	Date:		10/1/2003 Added LogoImageName BGS
*	Date:		12/17/2003 Added CarrierPrefix BGS
*	Date:		12/18/2003 Added DBName BGS
*	Date		7/19/2005 Added Packing Destinations variables Like from @EnableSinglePDF to @B2BDir
******************************************************************/
CREATE PROCEDURE [dbo].[z_spsGetCompanyInfo]
	@bHideDeleted		bit =1, 
	@UID			int,
	@Companyid		Int=null,
	@OrderBy		varchar(500)=null,
	@GroupBy		varchar(500)=null,
	@SearchBy		varchar(1000)=null,
--updating the Company Table
	@bDelete		bit=0,
	@bAdd			bit=0,
	@bEdit			bit=0,
--Update variables
	@Name 			varchar(30)=null,
	@DBName			varchar(50)=null,
	@Code			varchar(3)=null,
	@CarrierPrefix		varchar(50)=null,
	@Comments		varchar(150)=null,
	@IsClientOf		int=NULL,
	@EnableSingleFile	bit=NULL,
	@SingleFileEmail	varchar(50)=Null,
	@PDFJpegQuality		int=null,
	@LogoImageName	varchar(50)=Null,
	@DateLastUpdated	datetime=Null,
--ViewHistory
	@bViewHistory		bit=NULL,
	@FromDate		varchar(50)=Null,
	@ToDate			varchar(50)=Null,
--Add Packing Destinations
	@EnableSinglePDF	bit=0,
	@SinglePDFEmail		varchar(255)='',
	@SinglePDFEmailCC	varchar(255)='',
	@SinglePDFEmailBCC	varchar(255)='',
	@SingleFileEmailCC	varchar(255)='',
	@SingleFileEmailBCC	varchar(255)='',
	@EmailDocsOnly		varchar(255)='',
	@EmailDocsOnlyCC	varchar(255)='',
	@EmailDocsOnlyBCC	varchar(255)='',
	@EmailPhotosOnly	varchar(255)='',
	@EmailPhotosOnlyCC	varchar(255)='',
	@EmailPhotosOnlyBCC	varchar(255)='',
	@FTPSingleFileUrl	varchar(255)='',
	@FTPSingleFileUserName	varchar(50)='',
	@FTPSingleFilePassword	varchar(50)='',
	@HttpPostSingleFileUrl	varchar(255)='',
	@HttpPostSingleFileUserName	varchar(50)='',
	@HttpPostSingleFilePassword	varchar(50)='',
	@B2BDir			varchar(255)=''
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

IF @Companyid Is Not Null And @bViewHistory is Null
BEGIN
	IF @bDelete=0 And @bAdd=0 And @bEdit=0
	BEGIN
		SELECT 
			[CompanyID],
			[Name] ,
			[DBName],
			[Code] ,
			[CarrierPrefix],
			[Comments] ,
			[IsClientOf] ,
			[EnableSingleFile] ,
			[SingleFileEmail] ,
			[PDFJpegQuality],
			[Active] ,
			[LogoImageName],
			[DateLastUpdated] ,
			[UpdateByUserID],
			[EnableSinglePDF],
			[SinglePDFEmail],
			[SinglePDFEmailCC],
			[SinglePDFEmailBCC],
			[SingleFileEmailCC],
			[SingleFileEmailBCC],
			[EmailDocsOnly],
			[EmailDocsOnlyCC],
			[EmailDocsOnlyBCC],
			[EmailPhotosOnly],
			[EmailPhotosOnlyCC],
			[EmailPhotosOnlyBCC],
			[FTPSingleFileUrl],
			[FTPSingleFileUserName],
			[FTPSingleFilePassword],
			[HttpPostSingleFileUrl],
			[HttpPostSingleFileUserName],
			[HttpPostSingleFilePassword],
			[B2BDir]
		FROM	Company
		WHERE 	[CompanyID] = @Companyid
		AND 	[Active] =1
		ORDER BY [Active] Desc, [Name]
	END
	ELSE
	BEGIN
		IF @bDelete =1
		BEGIN
			UPDATE Company SET 	[Active]=	(CASE 	WHEN [Active]= 0
									THEN 1
									ELSE 0
									END
								),
						[DateLastUpdated]=@DateLastUpdated ,
						[UpdateByUserID]=@UID
			WHERE CompanyID = @Companyid
		
		END
		ELSE 
		BEGIN
			IF @bAdd=1
			BEGIN
				INSERT INTO Company(
							[Name] ,
							[DBName],
							[Code] ,
							[CarrierPrefix],
							[Comments],
							[IsClientOf],
							[EnableSingleFile],
							[SingleFileEmail] ,
							[PDFJpegQuality],
							[Active] ,
							[LogoImageName],
							[DateLastUpdated],
							[UpdateByUserID],
							[EnableSinglePDF],
							[SinglePDFEmail],
							[SinglePDFEmailCC],
							[SinglePDFEmailBCC],
							[SingleFileEmailCC],
							[SingleFileEmailBCC],
							[EmailDocsOnly],
							[EmailDocsOnlyCC],
							[EmailDocsOnlyBCC],
							[EmailPhotosOnly],
							[EmailPhotosOnlyCC],
							[EmailPhotosOnlyBCC],
							[FTPSingleFileUrl],
							[FTPSingleFileUserName],
							[FTPSingleFilePassword],
							[HttpPostSingleFileUrl],
							[HttpPostSingleFileUserName],
							[HttpPostSingleFilePassword],
							[B2BDir]
						)
				SELECT 		@Name As [Name],
						@DBName as [DBName],
						@Code As [Code],
						@CarrierPrefix As [CarrierPrefix],
						@Comments As [Comments],
						@IsClientOf As [IsClientOf],
						@EnableSingleFile As [EnableSingleFile],
						@SingleFileEmail As [SingleFileEmail] ,
						@PDFJpegQuality As [PDFJpegQuality],
						1 As [Active] ,
						@LogoImageName As [LogoImageName],
						@DateLastUpdated As [DateLastUpdated],
						@UID As [UpdateByUserID],
						@EnableSinglePDF As [EnableSinglePDF],
						@SinglePDFEmail As [SinglePDFEmail],
						@SinglePDFEmailCC As [SinglePDFEmailCC],
						@SinglePDFEmailBCC As [SinglePDFEmailBCC],
						@SingleFileEmailCC As [SingleFileEmailCC],
						@SingleFileEmailBCC As [SingleFileEmailBCC],
						@EmailDocsOnly As [EmailDocsOnly],
						@EmailDocsOnlyCC As [EmailDocsOnlyCC],
						@EmailDocsOnlyBCC As [EmailDocsOnlyBCC],
						@EmailPhotosOnly As [EmailPhotosOnly],
						@EmailPhotosOnlyCC As [EmailPhotosOnlyCC],
						@EmailPhotosOnlyBCC As [EmailPhotosOnlyBCC],
						@FTPSingleFileUrl As [FTPSingleFileUrl],
						@FTPSingleFileUserName As [FTPSingleFileUserName],
						@FTPSingleFilePassword As [FTPSingleFilePassword],
						@HttpPostSingleFileUrl As [HttpPostSingleFileUrl],
						@HttpPostSingleFileUserName As [HttpPostSingleFileUserName],
						@HttpPostSingleFilePassword As [HttpPostSingleFilePassword],
						@B2BDir As [B2BDir]							
					
			END
			ELSE 
			BEGIN
				IF @bEdit=1
				BEGIN
					Update Company Set 	[Name] = @Name ,
								[DBName] = @DBName,
								[Code] = @Code,
								[CarrierPrefix] = @CarrierPrefix,
								[Comments] = @Comments,
								[IsClientOf] = @IsClientOf,
								[EnableSingleFile] = @EnableSingleFile,
								[SingleFileEmail] = @SingleFileEmail,
								[PDFJpegQuality] = @PDFJpegQuality,
								[LogoImageName] = @LogoImageName,
								[DateLastUpdated] = @DateLastUpdated ,
								[UpdateByUserID] = @UID,
								[EnableSinglePDF] = @EnableSinglePDF,
								[SinglePDFEmail] = @SinglePDFEmail,
								[SinglePDFEmailCC] = @SinglePDFEmailCC,
								[SinglePDFEmailBCC] = @SinglePDFEmailBCC,
								[SingleFileEmailCC] = @SingleFileEmailCC,
								[SingleFileEmailBCC] = @SingleFileEmailBCC,
								[EmailDocsOnly] = @EmailDocsOnly,
								[EmailDocsOnlyCC] = @EmailDocsOnlyCC,
								[EmailDocsOnlyBCC] = @EmailDocsOnlyBCC,
								[EmailPhotosOnly] = @EmailPhotosOnly,
								[EmailPhotosOnlyCC] = @EmailPhotosOnlyCC,
								[EmailPhotosOnlyBCC] = @EmailPhotosOnlyBCC,
								[FTPSingleFileUrl] = @FTPSingleFileUrl,
								[FTPSingleFileUserName] = @FTPSingleFileUserName,
								[FTPSingleFilePassword] = @FTPSingleFilePassword,
								[HttpPostSingleFileUrl] = @HttpPostSingleFileUrl,
								[HttpPostSingleFileUserName] = @HttpPostSingleFileUserName,
								[HttpPostSingleFilePassword] = @HttpPostSingleFilePassword,
								[B2BDir] = @B2BDir
					WHERE CompanyID = @Companyid
				END
			END
		END
	END
END
ELSE
BEGIN
	IF @GroupBy Is Not Null
	BEGIN
		SET @sSQL = 'SELECT RetCompany.' + @GroupBy + ' '
	END
	ELSE
	BEGIN
		SET @sSQL = 'SELECT RetCompany.* '
	END	
	SET @sSQL = @sSQL + 'FROM( ' 
	SET @sSQL = @sSQL + 'SELECT '
	SET @sSQL = @sSQL + '[CompanyID], '
	SET @sSQL = @sSQL + '[Name] , '
	SET @sSQL = @sSQL + '[DBName] , '
	SET @sSQL = @sSQL + '[Code] , '
	SET @sSQL = @sSQL + '[CarrierPrefix] , '
	SET @sSQL = @sSQL + '[Comments] , '
	SET @sSQL = @sSQL + '[IsClientOf] , '
	SET @sSQL = @sSQL + '(SELECT 	NAME '
	SET @sSQL = @sSQL + 'FROM		Company '
	SET @sSQL = @sSQL + 'WHERE CompanyID = C.[IsClientOf]) As IsClientOfName, '
	SET @sSQL = @sSQL + '[EnableSingleFile] , '
	SET @sSQL = @sSQL + '[SingleFileEmail] , '
	SET @sSQL = @sSQL + '[PDFJpegQuality], '
	SET @sSQL = @sSQL + '[Active] , '
	SET @sSQL = @sSQL + '[LogoImageName], '
	SET @sSQL = @sSQL + '[DateLastUpdated] , '
	SET @sSQL = @sSQL + '(SELECT	USERNAME '
	SET @sSQL = @sSQL + 'FROM 		USERS '
	SET @sSQL = @sSQL + 'WHERE 	USERSID = C.[UpdateByUserID]) As [UpdateByUserName],  '
	SET @sSQL = @sSQL + '[UpdateByUserID],  '
	SET @sSQL = @sSQL + '[EnableSinglePDF], '
	SET @sSQL = @sSQL + '[SinglePDFEmail], '
	SET @sSQL = @sSQL + '[SinglePDFEmailCC], '
	SET @sSQL = @sSQL + '[SinglePDFEmailBCC], '
	SET @sSQL = @sSQL + '[SingleFileEmailCC], '
	SET @sSQL = @sSQL + '[SingleFileEmailBCC], '
	SET @sSQL = @sSQL + '[EmailDocsOnly], '
	SET @sSQL = @sSQL + '[EmailDocsOnlyCC], '
	SET @sSQL = @sSQL + '[EmailDocsOnlyBCC], '
	SET @sSQL = @sSQL + '[EmailPhotosOnly], '
	SET @sSQL = @sSQL + '[EmailPhotosOnlyCC], '
	SET @sSQL = @sSQL + '[EmailPhotosOnlyBCC], '
	SET @sSQL = @sSQL + '[FTPSingleFileUrl], '
	SET @sSQL = @sSQL + '[FTPSingleFileUserName], '
	SET @sSQL = @sSQL + '[FTPSingleFilePassword], '
	SET @sSQL = @sSQL + '[HttpPostSingleFileUrl], '
	SET @sSQL = @sSQL + '[HttpPostSingleFileUserName], '
	SET @sSQL = @sSQL + '[HttpPostSingleFilePassword], '
	SET @sSQL = @sSQL + '[B2BDir] '
	
	IF @bViewHistory Is Null
	BEGIN
		SET @sSQL = @sSQL + 'FROM Company C '
		SET @sSQL = @sSQL + ') RetCompany '
		IF @bHideDeleted = 1
		BEGIN
			SET @sSQL = @sSQL + 'WHERE Active = 1 '
		END
		ELSE
		BEGIN
			SET @sSQL = @sSQL + 'WHERE Active Is Not Null '
		END
		--Searching
		IF @SearchBy Is Not Null
		BEGIN
			SET @sSQL = @sSQL + @SearchBy + ' '
		END
	END
	ELSE
	BEGIN
		SET @sSQL = @sSQL + 'FROM CompanyHistory C '
		SET @sSQL = @sSQL + ') RetCompany '
		SET @sSQL = @sSQL + 'WHERE [DateLastUpdated] >= Convert(datetime,''' + @FromDate + ''') '
		SET @sSQL = @sSQL + 'AND [DateLastUpdated] <= Convert(datetime,''' + @ToDate + ''') '
		IF @Companyid <> -1
		BEGIN
			SET @sSQL = @sSQL + 'AND [CompanyID] = ' + Cast(@Companyid as varchar(10)) + ' '
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
				SET @sSQL = @sSQL + 'ORDER BY [Active] Desc, [Name] '
			END
			ELSE
			BEGIN
				SET @sSQL = @sSQL + 'ORDER BY  [Active] Desc, ' + @OrderBy + ' '
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

GRANT  EXECUTE  ON [dbo].[z_spsGetCompanyInfo]  TO [CFUSER]
GO

