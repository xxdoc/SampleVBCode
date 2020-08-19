SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spsGetClassOfLossInfo]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spsGetClassOfLossInfo]
GO

/******************************************************************
*
*    	Description:	Class Of Loss Info
*			
*	
*	Author: 	Brad Skidmore
*	Date: 		11/5/2003
*
******************************************************************/


CREATE PROCEDURE [dbo].[z_spsGetClassOfLossInfo]
	@bHideDeleted		bit =1, 
	@UID			int,
	@ClassOfLossID		Int=null,
	@OrderBy		varchar(500)=null,
	@GroupBy		varchar(500)=null,
	@SearchBy		varchar(1000)=null,
--updating the ClassOfLoss Table
	@bDelete		bit=0,
	@bAdd			bit=0,
	@bEdit			bit=0,
--Update variables
	@ClientCompanyID	int=null ,
	@ClassTypeID		int=null,
	@IsSubSetOFClassOfLossID	int=null,
	@Code			varchar(50)=null,
	@Description 		varchar(150)=null ,
	@IsDeleted		bit=null,
	@DateLastUpdated 	datetime =null,
--ViewHistory
	@bViewHistory		bit=NULL,
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

IF @ClassOfLossID Is Not Null And @bViewHistory is Null
BEGIN
	IF @bDelete=0 And @bAdd=0 And @bEdit=0
	BEGIN
		SELECT 
			[ClassOfLossID] ,
			[ClientCompanyID],
			[ClassTypeID],
			[IsSubSetOFClassOfLossID],
			[Code],
			[Description] ,
			[IsDeleted] ,
			[DateLastUpdated],
			[UpdateByUserID]
		FROM	ClassOfLoss
		WHERE 	[ClassOfLossID] = @ClassOfLossID
	END
	ELSE
	BEGIN
		IF @bDelete =1
		BEGIN
			UPDATE ClassOfLoss SET 	[IsDeleted]=	(CASE 	WHEN [IsDeleted]= 0
									THEN 1
									ELSE 0
									END
								),
						[DateLastUpdated]=@DateLastUpdated ,
						[UpdateByUserID]=@UID
			WHERE ClassOfLossID = @ClassOfLossID
		
		END
		ELSE 
		BEGIN
			IF @bAdd=1
			BEGIN
				INSERT INTO ClassOfLoss(
							[ClientCompanyID],
							[ClassTypeID],
							[IsSubSetOFClassOfLossID],
							[Code],
							[Description] ,
							[IsDeleted] ,
							[DateLastUpdated],
							[UpdateByUserID]
						)
				SELECT 			
							@ClientCompanyID As [ClientCompanyID] ,
							@ClassTypeID AS [ClassTypeID],
							@IsSubSetOFClassOfLossID As [IsSubSetOFClassOfLossID],
							@Code As [Code],
							@Description As [Description],
							@IsDeleted As [IsDeleted],
							@DateLastUpdated As [DateLastUpdated],
							@UID As [UpdateByUserID] 
					
			END
			ELSE 
			BEGIN
				IF @bEdit=1
				BEGIN
					Update ClassOfLoss Set 	
								[ClientCompanyID] = @ClientCompanyID,
								[ClassTypeID] = @ClassTypeID,
								[IsSubSetOFClassOfLossID] = @IsSubSetOFClassOfLossID,
								[Code] = @Code,
								[Description] = @Description ,
								[IsDeleted] = @IsDeleted   ,
								[DateLastUpdated] = @DateLastUpdated  ,
								[UpdateByUserID] = @UID   
					WHERE 	ClassOfLossID = @ClassOfLossID
				END
			END
		END
	END
END
ELSE
BEGIN
	IF @GroupBy Is Not Null
	BEGIN
		SET @sSQL = 'SELECT RetClassOfLoss.' + @GroupBy + ' '
	END
	ELSE
	BEGIN
		SET @sSQL = 'SELECT RetClassOfLoss.* '
	END	
	SET @sSQL = @sSQL + 'FROM( ' 
	SET @sSQL = @sSQL + 'SELECT '
	SET @sSQL = @sSQL + '[ClassOfLossID] , '
	SET @sSQL = @sSQL + '(SELECT 	NAME '
	SET @sSQL = @sSQL + 'FROM	Company '
	SET @sSQL = @sSQL + 'WHERE CompanyID = COL.[ClientCompanyID]) As ClientCompanyName, '
	SET @sSQL = @sSQL + '[ClientCompanyID], '
	SET @sSQL = @sSQL + '(SELECT 	Class '
	SET @sSQL = @sSQL + 'FROM	ClassType '
	SET @sSQL = @sSQL + 'WHERE ClassTypeID = COL.[ClassTypeID]) As ClassTypeClass, '
	SET @sSQL = @sSQL + '(SELECT 	Description '
	SET @sSQL = @sSQL + 'FROM	ClassType '
	SET @sSQL = @sSQL + 'WHERE ClassTypeID = COL.[ClassTypeID]) As ClassTypeDescription, '
	SET @sSQL = @sSQL + '[ClassTypeID], '
	SET @sSQL = @sSQL + '(SELECT 	Class '
	SET @sSQL = @sSQL + 'FROM	ClassType '
	SET @sSQL = @sSQL + 'WHERE ClassTypeID = (SELECT ClassTypeID FROM ClassOfLoss WHERE ClassOfLossID = COL.[IsSubSetOFClassOfLossID])) As IsSubSetOFClassOfLossClass, '
	SET @sSQL = @sSQL + '(SELECT 	Code '
	SET @sSQL = @sSQL + 'FROM	ClassOfLoss '
	SET @sSQL = @sSQL + 'WHERE ClassOfLossID = COL.[IsSubSetOFClassOfLossID]) As IsSubSetOFClassOfLossCode, '
	SET @sSQL = @sSQL + '(SELECT 	Description '
	SET @sSQL = @sSQL + 'FROM	ClassOfLoss '
	SET @sSQL = @sSQL + 'WHERE ClassOfLossID = COL.[IsSubSetOFClassOfLossID]) As IsSubSetOFClassOfLossDescription, '
	SET @sSQL = @sSQL + '[IsSubSetOFClassOfLossID], '
	SET @sSQL = @sSQL + '[Code], '
	SET @sSQL = @sSQL + '[Description] , '
	SET @sSQL = @sSQL + '[IsDeleted] , '
	SET @sSQL = @sSQL + '[DateLastUpdated] , '
	SET @sSQL = @sSQL + '(SELECT	USERNAME '
	SET @sSQL = @sSQL + 'FROM 		USERS '
	SET @sSQL = @sSQL + 'WHERE 	USERSID = COL.[UpdateByUserID]) As [UpdateByUserName],  '
	SET @sSQL = @sSQL + '[UpdateByUserID]  '

	IF @bViewHistory Is Null
	BEGIN
		SET @sSQL = @sSQL + 'FROM ClassOfLoss COL '
		SET @sSQL = @sSQL + ') RetClassOfLoss '
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
	END
	ELSE
	BEGIN
		SET @sSQL = @sSQL + 'FROM ClassOfLossHistory COL '
		SET @sSQL = @sSQL + ') RetClassOfLoss '
		SET @sSQL = @sSQL + 'WHERE [DateLastUpdated] >= Convert(datetime,''' + @FromDate + ''') '
		SET @sSQL = @sSQL + 'AND [DateLastUpdated] <= Convert(datetime,''' + @ToDate + ''') '
		IF @ClassOfLossID <> -1
		BEGIN
			SET @sSQL = @sSQL + 'AND [ClassOfLossID] = ' + Cast(@ClassOfLossID as varchar(10)) + ' '
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
				SET @sSQL = @sSQL + 'ORDER BY [IsDeleted] Asc, [ClassTypeClass] '
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

GRANT  EXECUTE  ON [dbo].[z_spsGetClassOfLossInfo]  TO [CFUSER]
GO

