SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spsGetCatInfo]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spsGetCatInfo]
GO

/******************************************************************
*
*    	Description:	Cat Info
*			
*	
*	Author: 	Brad Skidmore
*	Date: 		09/11/2003
*	Date:		10/2/2003 Added CompanyID BGS
******************************************************************/


CREATE PROCEDURE [dbo].[z_spsGetCatInfo]
	@bHideDeleted		bit =1, 
	@UID			int,
	@Catid			Int=null,
--Current Selected Company
	@selCompany		Int=null,
	@OrderBy		varchar(500)=null,
	@GroupBy		varchar(500)=null,
	@SearchBy		varchar(1000)=null,
--updating the Cat Table
	@bDelete		bit=0,
	@bAdd			bit=0,
	@bEdit			bit=0,
--Update variables
	@AssignmentTypeID	int=null,
	@Name 			varchar(20)=null,
	@Description		varchar(255)=null,
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

IF @Catid Is Not Null And @bViewHistory is Null
BEGIN
	IF @bDelete=0 And @bAdd=0 And @bEdit=0
	BEGIN
		SELECT 
			[CatID],
			[CompanyID],
			[AssignmentTypeID],
			[Name] ,
			[Description],
			[ActiveDate],
			[InactiveDate] ,
			[DateLastUpdated] ,
			[UpdateByUserID]
		FROM	Cat
		WHERE 	[CatID] = @Catid
		AND 	[CompanyID] = @selCompany
	END
	ELSE
	BEGIN
		IF @bDelete =1
		BEGIN
			UPDATE Cat SET 	[InactiveDate]=	(CASE 	WHEN [InactiveDate]Is Null
									THEN Convert(datetime,dbo.FormatDate(GetDate(),'mm/dd/yyy'))
									ELSE Null
									END
								),
						[DateLastUpdated]=@DateLastUpdated ,
						[UpdateByUserID]=@UID
			WHERE CatID = @Catid
		
		END
		ELSE 
		BEGIN
			IF @bAdd=1
			BEGIN
				INSERT INTO Cat(
							[CompanyID],
							[AssignmentTypeID],
							[Name] ,
							[Description],
							[ActiveDate],
							[InactiveDate] ,
							[DateLastUpdated] ,
							[UpdateByUserID] 
						)
				SELECT 			
							@selCompany As [CompanyID],
							@AssignmentTypeID As [AssignmentTypeID],
							@Name As [Name],
							@Description As [Description],
							@ActiveDate As [ActiveDate],
							@InactiveDate As [InactiveDate] ,
							@DateLastUpdated As [DateLastUpdated],
							@UID As [UpdateByUserID] 
			END
			ELSE 
			BEGIN
				IF @bEdit=1
				BEGIN
					Update Cat Set 	
							[CompanyID] = @selCompany,
							[AssignmentTypeID] = @AssignmentTypeID,
							[Name] = @Name ,
							[Description] = @Description,
							[ActiveDate] = @ActiveDate,
							[InactiveDate] = @InactiveDate,
							[DateLastUpdated] = @DateLastUpdated ,
							[UpdateByUserID] = @UID 
					WHERE CatID = @Catid
				END
			END
		END
	END
END

ELSE
BEGIN
	IF @selCompany IS Null
	BEGIN
		SET @selCompany = 0
	END
	IF @GroupBy Is Not Null
	BEGIN
		SET @sSQL = 'SELECT RetCat.' + @GroupBy + ' '
	END
	ELSE
	BEGIN
		SET @sSQL = 'SELECT RetCat.* '
	END	
	SET @sSQL = @sSQL + 'FROM( ' 
	SET @sSQL = @sSQL + 'SELECT '
	SET @sSQL = @sSQL + '[CatID], '
	SET @sSQL = @sSQL + '[CompanyID], '
	SET @sSQL = @sSQL + '[AssignmentTypeID], '
	SET @sSQL = @sSQL + '(SELECT	Type '
	SET @sSQL = @sSQL + 'FROM 	AssignmentType '
	SET @sSQL = @sSQL + 'WHERE 	AssignmentTypeID = C.[AssignmentTypeID]) As [AssignmentTypeType],  '
	SET @sSQL = @sSQL + '(SELECT	Description '
	SET @sSQL = @sSQL + 'FROM 	AssignmentType '
	SET @sSQL = @sSQL + 'WHERE 	AssignmentTypeID = C.[AssignmentTypeID]) As [AssignmentTypeDescription],  '
	SET @sSQL = @sSQL + '[Name] , '
	SET @sSQL = @sSQL + '[Description] , '
	SET @sSQL = @sSQL + '[ActiveDate] , '
	SET @sSQL = @sSQL + '[InactiveDate] , '
	SET @sSQL = @sSQL + '[DateLastUpdated] , '
	SET @sSQL = @sSQL + '(SELECT	USERNAME '
	SET @sSQL = @sSQL + 'FROM 	USERS '
	SET @sSQL = @sSQL + 'WHERE 	USERSID = C.[UpdateByUserID]) As [UpdateByUserName],  '
	SET @sSQL = @sSQL + '[UpdateByUserID] '

	IF @bViewHistory Is Null
	BEGIN
		SET @sSQL = @sSQL + 'FROM Cat C '
		SET @sSQL = @sSQL + 'WHERE [CompanyID] = ' + Cast(@selCompany as varchar(20)) + ' '
		SET @sSQL = @sSQL + ') RetCat '
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
		SET @sSQL = @sSQL + 'FROM CatHistory C '
		SET @sSQL = @sSQL + 'WHERE [CompanyID] = ' + Cast(@selCompany as varchar(20)) + ' '
		SET @sSQL = @sSQL + ') RetCat '
		SET @sSQL = @sSQL + 'WHERE [DateLastUpdated] >= Convert(datetime,''' + @FromDate + ''') '
		SET @sSQL = @sSQL + 'AND [DateLastUpdated] <= Convert(datetime,''' + @ToDate + ''') '
		IF @Catid <> -1
		BEGIN
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
				SET @sSQL = @sSQL + 'ORDER BY [InactiveDate] Asc, [Name] '
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

GRANT  EXECUTE  ON [dbo].[z_spsGetCatInfo]  TO [CFUSER]
GO

