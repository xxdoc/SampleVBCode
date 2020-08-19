SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spsGetClientCoAdjusterSpecZIPInfo]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spsGetClientCoAdjusterSpecZIPInfo]
GO

/******************************************************************
*
*    	Description:	Client Co Adjuster Spec ZIP Info
*			
*	
*	Author: 	Brad Skidmore
*	Date: 		1/29/2004
******************************************************************/

CREATE PROCEDURE [dbo].[z_spsGetClientCoAdjusterSpecZIPInfo]
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
	@ClientCompanyCatSpecID int=null,
	@ZipCode		int=null,
	@ZipCodeCount		int=null,
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
	@ToDate			varchar(50)=Null
	
AS
DECLARE @sSQL Varchar(8000)
DECLARE @CurDate 		DateTime
DECLARE @sSQLTemp2 varchar(8000)
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
		AND 	[ZipCode] Is Not Null
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
			AND 	[ZipCode] Is Not Null
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
							@ClientCompanyCatSpecID As [ClientCompanyCatSpecID],
							@ZipCode As [ZipCode],
							@ZipCodeCount As [ZipCodeCount],
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
							[ClientCompanyCatSpecID] = @ClientCompanyCatSpecID,
							[ZipCode] = @ZipCode,
							[ZipCodeCount] = @ZipCodeCount,
							[ActiveDate] = @ActiveDate ,
							[InactiveDate] = @InactiveDate ,
							[DateLastUpdated] = @DateLastUpdated ,
							[UpdateByUserID] = @UID
					WHERE 		[ClientCoAdjusterSpecID] = @ClientCoAdjusterSpecID
					AND 		[ZipCode] Is Not Null
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
	SET @sSQL = @sSQL + '(SELECT	Name '
	SET @sSQL = @sSQL + 'FROM 	Cat '
	SET @sSQL = @sSQL + 'WHERE 	CATID =	(	SELECT 	CATID
								FROM 	ClientCompanyCatSpec 
								WHERE	ClientCompanyCatSpecID = C.[ClientCompanyCatSpecID]
							)
	
	)As CATName,  '
	SET @sSQL = @sSQL + '(SELECT 	CATID '
	SET @sSQL = @sSQL + 'FROM	ClientCompanyCatSpec '
	SET @sSQL = @sSQL + 'WHERE ClientCompanyCatSpecID = C.[ClientCompanyCatSpecID]) As CATID, '
	SET @sSQL = @sSQL + '(SELECT 	CatCode '
	SET @sSQL = @sSQL + 'FROM	ClientCompanyCatSpec '
	SET @sSQL = @sSQL + 'WHERE ClientCompanyCatSpecID = C.[ClientCompanyCatSpecID]) As CatCode, '
	SET @sSQL = @sSQL + '[ClientCompanyCatSpecID], '
	SET @sSQL = @sSQL + '[ZipCode], '
	SET @sSQL = @sSQL + '[ZipCodeCount], '
	SET @sSQLTemp2 = ''
	SET @sSQLTemp2 = @sSQLTemp2  + 'SELECT 	TOP 1 ClientCompanyCatSpecID '
	SET @sSQLTemp2 = @sSQLTemp2  + 'FROM 	ClientCoAdjusterSpec '
	SET @sSQLTemp2 = @sSQLTemp2  + 'WHERE 	UsersID = (SELECT UsersID FROM 	Users WHERE UserName = ''UUNASSIGNED'') '
	SET @sSQLTemp2 = @sSQLTemp2  + 'ORDER BY ClientCompanyCatSpecID Desc '
	SET @sSQL = @sSQL + '(SELECT 	city '
	SET @sSQL = @sSQL + 'FROM	ZIPS '
	SET @sSQL = @sSQL + 'WHERE 	Zip = Cast(C.[ZipCode] As VarChar(53))) As PACity, '
	SET @sSQL = @sSQL + '(SELECT 	State '
	SET @sSQL = @sSQL + 'FROM	ZIPS '
	SET @sSQL = @sSQL + 'WHERE 	Zip = Cast(C.[ZipCode] As VarChar(53))) As PAState, '
	SET @sSQL = @sSQL + '(SELECT 	COUNT(PAZIP) As CountOFPAZIP '
	SET @sSQL = @sSQL + 'FROM	Assignments '
	SET @sSQL = @sSQL + 'WHERE 	ClientCompanyCatSpecID = C.[ClientCompanyCatSpecID] '
	SET @sSQL = @sSQL + 'AND 	AdjusterSpecID = C.[ClientCoAdjusterSpecID] '
	SET @sSQL = @sSQL + 'AND 	PAZIP = C.[ZipCode] '
	SET @sSQL = @sSQL + 'AND 	IsDeleted = 0 ) As CountOFPAZIP, '
	SET @sSQL = @sSQL + '(SELECT 	COUNT(PAZIP) As CountOFPAZIP '
	SET @sSQL = @sSQL + 'FROM	Assignments '
	SET @sSQL = @sSQL + 'WHERE 	ClientCompanyCatSpecID = C.[ClientCompanyCatSpecID] '
	SET @sSQL = @sSQL + 'AND 	AdjusterSpecID = C.[ClientCoAdjusterSpecID] '
	SET @sSQL = @sSQL + 'AND 	PAZIP = C.[ZipCode] '
	SET @sSQL = @sSQL + 'AND	StatusID =(SELECT StatusID FROM Status WHERE Status = ''PENDING'')) As CountOFPAZIPPending, '
	SET @sSQL = @sSQL + '(SELECT 	COUNT(AssignmentsID) As TotalAssignments '
	SET @sSQL = @sSQL + 'FROM	Assignments '
	SET @sSQL = @sSQL + 'WHERE 	ClientCompanyCatSpecID IN	(	SELECT 	ClientCompanyCatSpecID
										FROM	ClientCompanyCatSpec

										WHERE 	ClientCompanyID = C.[ClientCompanyID]
										AND	CATID	= 	(	SELECT	CATID 
														FROM	ClientCompanyCatSpec
														WHERE ClientCompanyCatSpecID = C.[ClientCompanyCatSpecID]
													)
									) 
				'
	SET @sSQL = @sSQL + 'AND 	AdjusterSpecID IN	(	SELECT 	ClientCoAdjusterSpecID 
									FROM 	ClientCoAdjusterSpec 
									WHERE 	ClientCompanyID = C.[ClientCompanyID]
									AND	UsersID = C.[UsersID]
									AND	ClientCompanyCatSpecID IN	(	SELECT 	ClientCompanyCatSpecID
															FROM	ClientCompanyCatSpec
															WHERE 	ClientCompanyID = C.[ClientCompanyID]
															AND	CATID	= 	(	SELECT	CATID 
																			FROM	ClientCompanyCatSpec
																			WHERE ClientCompanyCatSpecID = C.[ClientCompanyCatSpecID]
																		)
														)
								)
				) As TotalAssignments, '
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
		SET @sSQL = @sSQL + 'FROM ClientCoAdjusterSpec C WHERE [ZipCode] Is Not Null '
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
		SET @sSQL = @sSQL + 'FROM ClientCoAdjusterSpecHistory C WHERE [ZipCode] Is Not Null '
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
				SET @sSQL = @sSQL + 'ORDER BY [InactiveDate] Asc, [CountOFPAZIP] Desc '
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
	--print len(@ssql)
	--print @ssql
	EXEC(@sSQL)
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[z_spsGetClientCoAdjusterSpecZIPInfo]  TO [CFUSER]
GO

