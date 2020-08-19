SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spsGetClientCoAdjusterSpecInfo]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spsGetClientCoAdjusterSpecInfo]
GO

/******************************************************************
*
*    	Description:	Client Co Adjuster Spec Info
*			
*	
*	Author: 	Brad Skidmore
*	Date: 		10/28/2003
*	Date		1/29/2004	Only Include where ZipCode is Null
*					Since Zip Code ACIDs will be handled in
*					In another SP.					
******************************************************************/

CREATE PROCEDURE [dbo].[z_spsGetClientCoAdjusterSpecInfo]
	@bHideDeleted		bit =1, 
	@UID			int,
	@ClientCoAdjusterSpecID	int=null,
	@OrderBy		varchar(500)=null,
	@GroupBy		varchar(500)=null,
	@SearchBy		varchar(1000)=null,
--updating the CompanyCat Table
	@bDelete		bit=0,
	@bAdd			bit=0,
	@bEdit			bit=0,
--Update variables
	@ClientCompanyID	int=null,
	@UsersID		int=null,
	@ACID			varchar(50)=null,
	@ACIDDescription	varchar(150)=null,
	@Comments		varchar(255)=null,
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
--	IncludeZipCode
	@bIncludeZipCode	bit=0
	
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

IF @ClientCoAdjusterSpecID Is Not Null And @bViewHistory is Null
BEGIN
	IF @bDelete=0 And @bAdd=0 And @bEdit=0
	BEGIN
		IF @bIncludeZipCode = 0
		BEGIN
			SELECT 
				[ClientCoAdjusterSpecID],
				[ClientCompanyID],
				[UsersID],
				[ACID],
				[ACIDDescription],
				[Comments],
				[ClientCompanyCatSpecID],
				[ZipCode],
				[ZipCodeCount],
				[ActiveDate],
				[InactiveDate],
				[DateLastUpdated],
				[UpdateByUserID]
			FROM	ClientCoAdjusterSpec
			WHERE 	[ClientCoAdjusterSpecID] = @ClientCoAdjusterSpecID
			AND 	[ZipCode] Is Null
		END
		ELSE
		BEGIN
			SELECT 
				[ClientCoAdjusterSpecID],
				[ClientCompanyID],
				[UsersID],
				[ACID],
				[ACIDDescription],
				[Comments],
				[ClientCompanyCatSpecID],
				[ZipCode],
				[ZipCodeCount],
				[ActiveDate],
				[InactiveDate],
				[DateLastUpdated],
				[UpdateByUserID]
			FROM	ClientCoAdjusterSpec
			WHERE 	[ClientCoAdjusterSpecID] = @ClientCoAdjusterSpecID
		END
		
	END
	ELSE
	BEGIN
		IF @bDelete =1
		BEGIN
			UPDATE ClientCoAdjusterSpec SET [InactiveDate]=	(CASE 	WHEN [InactiveDate]Is Null
									THEN Convert(datetime,dbo.FormatDate(GetDate(),'mm/dd/yyy'))
									ELSE Null
									END
								),
						[DateLastUpdated]=@DateLastUpdated ,
						[UpdateByUserID]=@UID
			WHERE 	[ClientCoAdjusterSpecID] = @ClientCoAdjusterSpecID
			AND 	[ZipCode] Is Null
		END
		ELSE 
		BEGIN
			IF @bAdd=1
			BEGIN
				INSERT INTO ClientCoAdjusterSpec(
							[ClientCompanyID],
							[UsersID],
							[ACID],
							[ACIDDescription],
							[Comments],
							[ClientCompanyCatSpecID],
							[ZipCode],
							[ZipCodeCount],
							[ActiveDate],
							[InactiveDate],
							[DateLastUpdated],
							[UpdateByUserID]
						)
				SELECT 			
							@ClientCompanyID As [ClientCompanyID],
							@UsersID As [UsersID],
							@ACID As [ACID],
							@ACIDDescription As [ACIDDescription],
							@Comments As [Comments],
							Null As [ClientCompanyCatSpecID],
							Null As [ZipCode],
							Null As [ZipCodeCount],
							@ActiveDate As [ActiveDate] ,
							@InactiveDate As [InactiveDate] ,
							@DateLastUpdated As [DateLastUpdated],
							@UID As [UpdateByUserID] 
			END
			ELSE 
			BEGIN
				IF @bEdit=1
				BEGIN
					Update ClientCoAdjusterSpec Set 	
							[ClientCompanyID] = @ClientCompanyID,
							[UsersID] = @UsersID,
							[ACID] = @ACID,
							[ACIDDescription] = @ACIDDescription,
							[Comments] = @Comments ,
							[ClientCompanyCatSpecID] = Null,
							[ZipCode] = Null,
							[ZipCodeCount] = Null,
							[ActiveDate] = @ActiveDate ,
							[InactiveDate] = @InactiveDate ,
							[DateLastUpdated] = @DateLastUpdated ,
							[UpdateByUserID] = @UID
					WHERE 		[ClientCoAdjusterSpecID] = @ClientCoAdjusterSpecID
					AND 		[ZipCode] Is Null
				END
			END
		END
	END
END
ELSE
BEGIN
	IF @GroupBy Is Not Null
	BEGIN
		SET @sSQL = 'SELECT RetClientCoAdjusterSpec.' + @GroupBy + ' '
	END
	ELSE
	BEGIN
		SET @sSQL = 'SELECT RetClientCoAdjusterSpec.* '
	END	
	SET @sSQL = @sSQL + 'FROM( ' 
	SET @sSQL = @sSQL + 'SELECT '
	SET @sSQL = @sSQL + 'ClientCoAdjusterSpecID, '
	SET @sSQL = @sSQL + '(SELECT 	NAME '
	SET @sSQL = @sSQL + 'FROM	Company '
	SET @sSQL = @sSQL + 'WHERE CompanyID = C.[ClientCompanyID]) As ClientCompanyName, '
	SET @sSQL = @sSQL + '(SELECT 	IsClientOF '
	SET @sSQL = @sSQL + 'FROM Company '
	SET @sSQL = @sSQL + 'WHERE CompanyID = C.[ClientCompanyID]) As IsClientOF, '
	SET @sSQL = @sSQL + '[ClientCompanyID], '
	SET @sSQL = @sSQL + '(SELECT	USERNAME '
	SET @sSQL = @sSQL + 'FROM 	USERS '
	SET @sSQL = @sSQL + 'WHERE 	USERSID = C.[UsersID]) As [UserName],  '
	SET @sSQL = @sSQL + '(SELECT	LastName + '', '' + FirstName '
	SET @sSQL = @sSQL + 'FROM 	USERS '
	SET @sSQL = @sSQL + 'WHERE 	USERSID = C.[UsersID]) As [LFName],  '
	SET @sSQL = @sSQL + '[UsersID], '
	SET @sSQL = @sSQL + '[ACID], '
	SET @sSQL = @sSQL + '[ACIDDescription], '
	SET @sSQL = @sSQL + '[Comments], '
	SET @sSQL = @sSQL + '[ClientCompanyCatSpecID], '
	SET @sSQL = @sSQL + '[ZipCode], '
	SET @sSQL = @sSQL + '[ZipCodeCount], '
	SET @sSQL = @sSQL + '[ActiveDate] , '
	SET @sSQL = @sSQL + '[InactiveDate] , '
	SET @sSQL = @sSQL + '[DateLastUpdated] , '
	SET @sSQL = @sSQL + '(SELECT	USERNAME '
	SET @sSQL = @sSQL + 'FROM 	USERS '
	SET @sSQL = @sSQL + 'WHERE 	USERSID = C.[UpdateByUserID]) As [UpdateByUserName],  '
	SET @sSQL = @sSQL + '[UpdateByUserID] '

	IF @bViewHistory Is Null
	BEGIN
		--Ignore ZipCode
		IF @bIncludeZipCode = 0
		BEGIN
			SET @sSQL = @sSQL + 'FROM ClientCoAdjusterSpec C WHERE [ZipCode] Is Null '
		END
		ELSE
		BEGIN
			SET @sSQL = @sSQL + 'FROM ClientCoAdjusterSpec C '
		END
		SET @sSQL = @sSQL + ') RetClientCoAdjusterSpec '
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
		--Ignore ZipCode
		IF @bIncludeZipCode = 0
		BEGIN
			SET @sSQL = @sSQL + 'FROM ClientCoAdjusterSpecHistory C WHERE [ZipCode] Is Null '
		END
		ELSE
		BEGIN
			SET @sSQL = @sSQL + 'FROM ClientCoAdjusterSpecHistory C '
		END
		
		SET @sSQL = @sSQL + ') RetClientCoAdjusterSpec '
		SET @sSQL = @sSQL + 'WHERE [DateLastUpdated] >= Convert(datetime,''' + @FromDate + ''') '
		SET @sSQL = @sSQL + 'AND [DateLastUpdated] <= Convert(datetime,''' + @ToDate + ''') '
		IF @ClientCoAdjusterSpecID <> -1 
		BEGIN
			SET @sSQL = @sSQL + 'AND [ClientCoAdjusterSpecID] = ' + Cast(@ClientCoAdjusterSpecID as varchar(10)) + ' '
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
				SET @sSQL = @sSQL + 'ORDER BY [InactiveDate] Asc, [LFName] '
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

GRANT  EXECUTE  ON [dbo].[z_spsGetClientCoAdjusterSpecInfo]  TO [CFUSER]
GO

