SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spsGetClientCompanyCatSpecInfo]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spsGetClientCompanyCatSpecInfo]
GO

/******************************************************************
*
*    	Description:	Client Company Cat Spec Info
*			
*	
*	Author: 	Brad Skidmore
*	Date: 		09/23/2003
*
******************************************************************/

CREATE PROCEDURE [dbo].[z_spsGetClientCompanyCatSpecInfo]
	@bHideDeleted		bit =1, 
	@UID			int,
	@ClientCompanyCatSpecID	int=null,
	@OrderBy		varchar(500)=null,
	@GroupBy		varchar(500)=null,
	@SearchBy		varchar(1000)=null,
--updating the CompanyCat Table
	@bDelete		bit=0,
	@bAdd			bit=0,
	@bEdit			bit=0,
--Update variables
	@ClientCompanyID		int=null,
	@CATID		int=null,
	@CatCode		varchar(50)=null,
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
	@ToDate			varchar(50)=Null
	
AS
DECLARE @sSQL Varchar(8000)
DECLARE @CurDate 		DateTime
DECLARE @AssignmentTypeID	int
SET @CurDate = GetDate()
--Get the AssignmentTypeID from the CAT Table !!!
SET @AssignmentTypeID = (
			SELECT 	[AssignmentTypeID]
			FROM	CAT
			WHERE	CATID = @CATID
			)
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

IF @ClientCompanyCatSpecID Is Not Null And @bViewHistory is Null
BEGIN
	IF @bDelete=0 And @bAdd=0 And @bEdit=0
	BEGIN
		SELECT 
			[ClientCompanyCatSpecID] ,
			[ClientCompanyID],
			[CATID] ,
			[CatCode],
			[Comments],
			[ActiveDate],
			[InactiveDate],
			[DateLastUpdated],
			[UpdateByUserID],
			[AssignmentTypeID] 
		FROM	ClientCompanyCatSpec
		WHERE 	[ClientCompanyCatSpecID] = @ClientCompanyCatSpecID
	END
	ELSE
	BEGIN
		IF @bDelete =1
		BEGIN
			UPDATE ClientCompanyCatSpec SET [InactiveDate]=	(CASE 	WHEN [InactiveDate]Is Null
									THEN Convert(datetime,dbo.FormatDate(GetDate(),'mm/dd/yyy'))
									ELSE Null
									END
								),
						[DateLastUpdated]=@DateLastUpdated ,
						[UpdateByUserID]=@UID
			WHERE 	[ClientCompanyCatSpecID] = @ClientCompanyCatSpecID
		END
		ELSE 
		BEGIN
			IF @bAdd=1
			BEGIN
				INSERT INTO ClientCompanyCatSpec(
							[ClientCompanyID],
							[CATID] ,
							[CatCode],
							[Comments],
							[ActiveDate],
							[InactiveDate],
							[DateLastUpdated],
							[UpdateByUserID],
							[AssignmentTypeID] 
						)
				SELECT 			@ClientCompanyID As [ClientCompanyID],
							@Catid As [CATID] ,
							@CatCode As [CatCode],
							@Comments As [Comments],
							@ActiveDate As [ActiveDate] ,
							@InactiveDate As [InactiveDate] ,
							@DateLastUpdated As [DateLastUpdated],
							@UID As [UpdateByUserID],
							@AssignmentTypeID As [AssignmentTypeID]
							 
			END
			ELSE 
			BEGIN
				IF @bEdit=1
				BEGIN
					Update ClientCompanyCatSpec Set 	
									[ClientCompanyID] = @ClientCompanyID ,
									[CATID] = @Catid  ,
									[CatCode] = @CatCode ,
									[Comments] = @Comments ,
									[ActiveDate] = @ActiveDate ,
									[InactiveDate] = @InactiveDate ,

									[DateLastUpdated] = @DateLastUpdated ,
									[UpdateByUserID] = @UID,
									[AssignmentTypeID] = @AssignmentTypeID
					WHERE 	[ClientCompanyCatSpecID] = @ClientCompanyCatSpecID
				END
			END
		END
	END
END
ELSE
BEGIN
	IF @GroupBy Is Not Null
	BEGIN
		SET @sSQL = 'SELECT RetClientCompanyCatSpec.' + @GroupBy + ' '
	END
	ELSE
	BEGIN
		SET @sSQL = 'SELECT RetClientCompanyCatSpec.* '
	END	
	SET @sSQL = @sSQL + 'FROM( ' 
	SET @sSQL = @sSQL + 'SELECT '
	SET @sSQL = @sSQL + 'ClientCompanyCatSpecID, '
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
	SET @sSQL = @sSQL + '[CATID], '
	SET @sSQL = @sSQL + '[CatCode], '
	SET @sSQL = @sSQL + '[Comments], '
	SET @sSQL = @sSQL + '[ActiveDate] , '
	SET @sSQL = @sSQL + '[InactiveDate] , '
	SET @sSQL = @sSQL + '[DateLastUpdated] , '
	SET @sSQL = @sSQL + '(SELECT	USERNAME '
	SET @sSQL = @sSQL + 'FROM 	USERS '
	SET @sSQL = @sSQL + 'WHERE 	USERSID = C.[UpdateByUserID]) As [UpdateByUserName],  '
	SET @sSQL = @sSQL + '[UpdateByUserID], '
	SET @sSQL = @sSQL + '(SELECT	[Type] '
	SET @sSQL = @sSQL + 'FROM 	AssignmentType '
	SET @sSQL = @sSQL + 'WHERE 	[AssignmentTypeID] = C.[AssignmentTypeID]) As [AssignmentTypeType],  '
	SET @sSQL = @sSQL + '[AssignmentTypeID] '

	IF @bViewHistory Is Null
	BEGIN
		SET @sSQL = @sSQL + 'FROM ClientCompanyCatSpec C '
		SET @sSQL = @sSQL + ') RetClientCompanyCatSpec '
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
		SET @sSQL = @sSQL + 'FROM ClientCompanyCatSpecHistory C '
		SET @sSQL = @sSQL + ') RetClientCompanyCatSpec '
		SET @sSQL = @sSQL + 'WHERE [DateLastUpdated] >= Convert(datetime,''' + @FromDate + ''') '
		SET @sSQL = @sSQL + 'AND [DateLastUpdated] <= Convert(datetime,''' + @ToDate + ''') '
		IF @ClientCompanyCatSpecID <> -1 
		BEGIN
			SET @sSQL = @sSQL + 'AND [ClientCompanyCatSpecID] = ' + Cast(@ClientCompanyCatSpecID as varchar(10)) + ' '
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

GRANT  EXECUTE  ON [dbo].[z_spsGetClientCompanyCatSpecInfo]  TO [CFUSER]
GO

