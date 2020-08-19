SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spsGetClientCompanyCatInfo]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spsGetClientCompanyCatInfo]
GO

/******************************************************************
*
*    	Description:	Client Company Cat Info
*			
*	
*	Author: 	Brad Skidmore
*	Date: 		09/16/2003
*	Date		10/2/2003 change name from z_spsGetCompanyCatInfo to z_spsGetClientCompanyCatInfo
*			1/26/2004 Added BillingCode for Batches processing
*	Date		7/19/2005 Added Packing Destinations variables Like from @EnableSinglePDF to @B2BDir
******************************************************************/
CREATE PROCEDURE [dbo].[z_spsGetClientCompanyCatInfo]
	@bHideDeleted		bit =1, 
	@UID			int,
	@ClientCompanyID		int=null,
	@Catid			Int=null,
	@OrderBy		varchar(500)=null,
	@GroupBy		varchar(500)=null,
	@SearchBy		varchar(1000)=null,
--updating the ClientCompanyCat Table
	@bDelete		bit=0,
	@bAdd			bit=0,
	@bEdit			bit=0,
--Update variables
	@BillingCode		varchar(50)=null,
	@TypeOfLossID		int=null,
	@FeeScheduleID		int=null,
	@SiteAddress		varchar(200)=null,
	@SACity			varchar(50)=null,
	@SAState		varchar(50)=null,
	@SAZip			int=NULL ,
	@SAZip4			int=NULL ,
	@SAOtherPostCode	varchar(20)=null,
	@ActiveDate		datetime=Null,
	@InactiveDate		datetime=Null,
	@DateLastUpdated	datetime=Null,
--ViewHistory
	@bViewHistory		bit=NULL,
	@FromActiveDate		varchar(50)='*',
	@ToActiveDate		varchar(50)='*',
	@FromInactiveDate	varchar(50)='*',
	@ToInactiveDate		varchar(50)='*',
	@FromDate		varchar(50)=Null,
	@ToDate			varchar(50)=Null,
--Assign By ZipCode Default for this Client Company Cat
	@AssignByZipDefault	bit =0,
--Add Packing Destinations
	@EnableSinglePDF	bit=0,
	@SinglePDFEmail		varchar(255)='',
	@SinglePDFEmailCC	varchar(255)='',
	@SinglePDFEmailBCC	varchar(255)='',
	@EnableSingleFile	bit=0,
	@SingleFileEmail	varchar(50)='',
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
--Clean String Dates
SET @FromActiveDate = 		dbo.CleanFromOrToDateString(@CurDate, @FromActiveDate, 1)
SET @ToActiveDate = 		dbo.CleanFromOrToDateString(@CurDate, @ToActiveDate, 0)
SET @FromInactiveDate = 	dbo.CleanFromOrToDateString(@CurDate, @FromInactiveDate, 1)
SET @ToInactiveDate = 		dbo.CleanFromOrToDateString(@CurDate, @ToInactiveDate, 0)

IF @ClientCompanyID Is Not Null And @Catid Is Not Null And @bViewHistory is Null
BEGIN
	IF @bDelete=0 And @bAdd=0 And @bEdit=0
	BEGIN
		SELECT 
			[ClientCompanyID] ,
			[CATID] ,
			[BillingCode],
			[TypeOfLossID] ,
			[FeeScheduleID] ,
			[SiteAddress] ,
			[SACity] ,
			[SAState],
			[SAZip] ,
			[SAZip4] ,
			[SAOtherPostCode] ,
			[ActiveDate] ,
			[InactiveDate] ,
			[AssignByZipDefault],
			[DateLastUpdated],
			[UpdateByUserID],
			[EnableSinglePDF],
			[SinglePDFEmail],
			[SinglePDFEmailCC],
			[SinglePDFEmailBCC],
			[EnableSingleFile],
			[SingleFileEmail],
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
		FROM	ClientCompanyCat
		WHERE 	[ClientCompanyID] = @ClientCompanyID
		AND	[CatID] = @Catid
	END
	ELSE
	BEGIN
		IF @bDelete =1
		BEGIN
			UPDATE ClientCompanyCat SET 	[InactiveDate]=	(CASE 	WHEN [InactiveDate]Is Null
									THEN Convert(datetime,dbo.FormatDate(GetDate(),'mm/dd/yyy'))
									ELSE Null
									END
								),
							[DateLastUpdated]=@DateLastUpdated ,
							[UpdateByUserID]=@UID
			WHERE 	[ClientCompanyID] = @ClientCompanyID
			AND	[CatID] = @Catid
		
		END
		ELSE 
		BEGIN
			IF @bAdd=1
			BEGIN
				INSERT INTO ClientCompanyCat(
							[ClientCompanyID] ,
							[CATID] ,
							[BillingCode],
							[TypeOfLossID] ,
							[FeeScheduleID] ,
							[SiteAddress] ,
							[SACity] ,
							[SAState],
							[SAZip] ,
							[SAZip4] ,
							[SAOtherPostCode] ,
							[ActiveDate] ,
							[InactiveDate] ,
							[AssignByZipDefault],
							[DateLastUpdated],
							[UpdateByUserID],
 							[EnableSinglePDF],
							[SinglePDFEmail],
							[SinglePDFEmailCC],
							[SinglePDFEmailBCC],
							[EnableSingleFile],
							[SingleFileEmail],
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
				SELECT 			@ClientCompanyID As [ClientCompanyID],
							@Catid As [CATID] ,
							@BillingCode As [BillingCode],
							@TypeOfLossID As [TypeOfLossID] ,
							@FeeScheduleID As [FeeScheduleID] ,
							@SiteAddress As [SiteAddress] ,
							@SACity	As [SACity] ,
							@SAState As [SAState],
							@SAZip As [SAZip] ,
							@SAZip4 As [SAZip4] ,
							@SAOtherPostCode As [SAOtherPostCode] ,
							@ActiveDate As [ActiveDate] ,
							@InactiveDate As [InactiveDate] ,
							@AssignByZipDefault As [AssignByZipDefault],
							@DateLastUpdated As [DateLastUpdated],
							@UID As [UpdateByUserID],
							@EnableSinglePDF As [EnableSinglePDF],
							@SinglePDFEmail As [SinglePDFEmail],
							@SinglePDFEmailCC As [SinglePDFEmailCC],
							@SinglePDFEmailBCC As [SinglePDFEmailBCC],
							@EnableSingleFile As [EnableSingleFile],
							@SingleFileEmail As [SingleFileEmail],
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
					Update ClientCompanyCat Set 	
								[ClientCompanyID] = @ClientCompanyID ,
								[CATID] = @Catid  ,
								[BillingCode] = @BillingCode,
								[TypeOfLossID] = @TypeOfLossID  ,
								[FeeScheduleID] = @FeeScheduleID  ,
								[SiteAddress] = @SiteAddress  ,
								[SACity] = @SACity ,
								[SAState] = @SAState ,
								[SAZip] = @SAZip ,
								[SAZip4] = @SAZip4 ,
								[SAOtherPostCode] = @SAOtherPostCode ,
								[ActiveDate] = @ActiveDate ,
								[InactiveDate] = @InactiveDate ,
								[AssignByZipDefault] = @AssignByZipDefault,
								[DateLastUpdated] = @DateLastUpdated ,
								[UpdateByUserID] = @UID,
								[EnableSinglePDF] = @EnableSinglePDF,
								[SinglePDFEmail] = @SinglePDFEmail,
								[SinglePDFEmailCC] = @SinglePDFEmailCC,
								[SinglePDFEmailBCC] = @SinglePDFEmailBCC,
								[EnableSingleFile] = @EnableSingleFile,
								[SingleFileEmail] = @SingleFileEmail,
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
					WHERE 	[ClientCompanyID] = @ClientCompanyID
					AND	[CatID] = @Catid
				END
			END
		END
	END
END
ELSE
BEGIN
	IF @GroupBy Is Not Null
	BEGIN
		SET @sSQL = 'SELECT RetClientCompanyCat.' + @GroupBy + ' '
	END
	ELSE
	BEGIN
		SET @sSQL = 'SELECT RetClientCompanyCat.* '
	END	
	SET @sSQL = @sSQL + 'FROM( ' 
	SET @sSQL = @sSQL + 'SELECT '
	SET @sSQL = @sSQL + '(SELECT 	NAME '
	SET @sSQL = @sSQL + 'FROM	Company '
	SET @sSQL = @sSQL + 'WHERE CompanyID = C.[ClientCompanyID]) As ClientCompanyName, '
	SET @sSQL = @sSQL + '(SELECT 	IsClientOF '
	SET @sSQL = @sSQL + 'FROM Company '
	SET @sSQL = @sSQL + 'WHERE CompanyID = C.[ClientCompanyID]) As IsClientOF, '
	SET @sSQL = @sSQL + '[ClientCompanyID], '
	SET @sSQL = @sSQL + '(SELECT 	NAME '
	SET @sSQL = @sSQL + 'FROM	Cat '
	SET @sSQL = @sSQL + 'WHERE CatID = C.[CATID]) As CatName, '
	SET @sSQL = @sSQL + '(SELECT 	DESCRIPTION '
	SET @sSQL = @sSQL + 'FROM Cat '
	SET @sSQL = @sSQL + 'WHERE CatID = C.[CATID]) As CatDescription, '
	SET @sSQL = @sSQL + '(SELECT 	Type '
	SET @sSQL = @sSQL + 'FROM	AssignmentType '
	SET @sSQL = @sSQL + 'WHERE 	AssignmentTypeID = (SELECT AssignmentTypeid FROM CAT WHERE CATID =C.[CATID])) As CatAssignmentType, '
	SET @sSQL = @sSQL + '[CATID], '
	SET @sSQL = @sSQL + '[BillingCode], '
	SET @sSQL = @sSQL + '(SELECT	TypeOfLoss + ''('' + Code + '')'' '
	SET @sSQL = @sSQL + 'FROM 	TypeOfLoss '
	SET @sSQL = @sSQL + 'WHERE 	TypeOfLossID = C.[TypeOfLossID]) As [TypeOfLoss],  '
	SET @sSQL = @sSQL + '[TypeOfLossID], '
	SET @sSQL = @sSQL + '(SELECT	ScheduleName  '
	SET @sSQL = @sSQL + 'FROM 	FeeSchedule '
	SET @sSQL = @sSQL + 'WHERE 	FeeScheduleID = C.[FeeScheduleID]) As [FeeSchedule],  '
	SET @sSQL = @sSQL + '[FeeScheduleID], '
	SET @sSQL = @sSQL + '[SiteAddress], '
	SET @sSQL = @sSQL + '[SACity], '
	SET @sSQL = @sSQL + '[SAState], '
	SET @sSQL = @sSQL + '[SAZip], '
	SET @sSQL = @sSQL + '[SAZip4], '
	SET @sSQL = @sSQL + '[SAOtherPostCode], '
	SET @sSQL = @sSQL + '[ActiveDate] , '
	SET @sSQL = @sSQL + '[InactiveDate] , '
	SET @sSQL = @sSQL + '[AssignByZipDefault], '
	SET @sSQL = @sSQL + '[DateLastUpdated] , '
	SET @sSQL = @sSQL + '(SELECT	USERNAME '
	SET @sSQL = @sSQL + 'FROM 	USERS '
	SET @sSQL = @sSQL + 'WHERE 	USERSID = C.[UpdateByUserID]) As [UpdateByUserName],  '
	SET @sSQL = @sSQL + '[UpdateByUserID], '
	SET @sSQL = @sSQL + '[EnableSinglePDF], '
	SET @sSQL = @sSQL + '[SinglePDFEmail], '
	SET @sSQL = @sSQL + '[SinglePDFEmailCC], '
	SET @sSQL = @sSQL + '[SinglePDFEmailBCC], '
	SET @sSQL = @sSQL + '[EnableSingleFile], '
	SET @sSQL = @sSQL + '[SingleFileEmail], '
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
		SET @sSQL = @sSQL + 'FROM ClientCompanyCat C '
		SET @sSQL = @sSQL + ') RetClientCompanyCat '
		IF @bHideDeleted = 1
		BEGIN
			SET @sSQL = @sSQL + 'WHERE [InactiveDate] Is Null '
		END
		ELSE
		BEGIN
			SET @sSQL = @sSQL + 'WHERE [ActiveDate] Is Not Null '
		END
		--Searching
		IF @SearchBy Is Not Null
		BEGIN
			SET @sSQL = @sSQL + @SearchBy + ' '
			--Searching Dates
			IF @FromActiveDate <> '*'
			BEGIN
				SET @sSQL = @sSQL + 'AND [ActiveDate] >= Convert(datetime,''' + @FromActiveDate + ''') '
			END
			IF @ToActiveDate <> '*'
			BEGIN
				SET @sSQL = @sSQL + 'AND [ActiveDate] <= Convert(datetime,''' + @ToActiveDate + ''') '
			END
			IF @FromInactiveDate <> '*'
			BEGIN
				SET @sSQL = @sSQL + 'AND [InactiveDate] >= Convert(datetime,''' + @FromInactiveDate + ''') '
			END
			IF @ToInactiveDate <> '*'
			BEGIN
				SET @sSQL = @sSQL + 'AND [InactiveDate] <= Convert(datetime,''' + @ToInactiveDate + ''') '
			END
		END
	END
	ELSE
	BEGIN
		SET @sSQL = @sSQL + 'FROM ClientCompanyCatHistory C '
		SET @sSQL = @sSQL + ') RetClientCompanyCat '
		SET @sSQL = @sSQL + 'WHERE [DateLastUpdated] >= Convert(datetime,''' + @FromDate + ''') '
		SET @sSQL = @sSQL + 'AND [DateLastUpdated] <= Convert(datetime,''' + @ToDate + ''') '
		IF @ClientCompanyID <> -1 And @Catid <> -1
		BEGIN
			SET @sSQL = @sSQL + 'AND [ClientCompanyID] = ' + Cast(@ClientCompanyID as varchar(10)) + ' '
			SET @sSQL = @sSQL + 'AND [CatID] = ' + Cast(@Catid as varchar(10)) + ' '
		END
		--Searching
		IF @SearchBy Is Not Null
		BEGIN
			SET @sSQL = @sSQL + @SearchBy + ' '
			--Searching Dates
			IF @FromActiveDate <> '*'
			BEGIN
				SET @sSQL = @sSQL + 'AND [ActiveDate] >= Convert(datetime,''' + @FromActiveDate + ''') '
			END
			IF @ToActiveDate <> '*'
			BEGIN
				SET @sSQL = @sSQL + 'AND [ActiveDate] <= Convert(datetime,''' + @ToActiveDate + ''') '
			END
			IF @FromInactiveDate <> '*'
			BEGIN
				SET @sSQL = @sSQL + 'AND [InactiveDate] >= Convert(datetime,''' + @FromInactiveDate + ''') '
			END
			IF @ToInactiveDate <> '*'
			BEGIN
				SET @sSQL = @sSQL + 'AND [InactiveDate] <= Convert(datetime,''' + @ToInactiveDate + ''') '
			END
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
				SET @sSQL = @sSQL + 'ORDER BY [InactiveDate] Asc, [ClientCompanyName] '
			END
			ELSE
			BEGIN
				SET @sSQL = @sSQL + 'ORDER BY  [InactiveDate] Asc, ' + @OrderBy + ' '
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

GRANT  EXECUTE  ON [dbo].[z_spsGetClientCompanyCatInfo]  TO [CFUSER]
GO

