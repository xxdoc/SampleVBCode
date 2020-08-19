SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spsGetPolicyLimitsInfo]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spsGetPolicyLimitsInfo]
GO

/******************************************************************
*
*    	Description:	PolicyLimits Info
*			
*	
*	Author: 	Brad Skidmore
*	Date: 		4/15/2004
*
******************************************************************/
CREATE PROCEDURE [dbo].[z_spsGetPolicyLimitsInfo]
	@bHideDeleted		bit =1, 
	@UID			int,
	@PolicyLimitsID		Int=null,
	@OrderBy		varchar(500)=null,
	@GroupBy		varchar(500)=null,
	@SearchBy		varchar(1000)=null,
--updating the TypeOfLoss Table
	@bDelete		bit=0,
	@bAdd			bit=0,
	@bEdit			bit=0,
--Update variables
	@AssignmentsID	 	int=null,
	@ID			int=null,
	@IDAssignments		int=null,
	@ClassTypeID		int=null,
	@LimitAmount		money=null,
	@RCSaidProp		money=null,
	@Reserves		money=null,
	@IsDeleted		bit=null,
	@DownLoadMe		bit=null,
	@UpLoadMe		bit=null,
	@AdminComments		varchar(255)=null,
	@DateLastUpdated	datetime=null,
--ViewHistory
	@bViewHistory		bit=NULL,
	@FromDate		varchar(50)=Null,
	@ToDate			varchar(50)=Null,
--IncludeFlags
	@bTableFieldsOnly	bit=0,
--Farmers
	@FarmersUnitRowID	varchar(20)=''
	
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

IF @PolicyLimitsID Is Not Null And @bViewHistory is Null
BEGIN
	IF @bDelete=0 And @bAdd=0 And @bEdit=0
	BEGIN
		SELECT 
			[PolicyLimitsID],
			[AssignmentsID] ,
			[ID],
			[IDAssignments],
			[ClassTypeID] ,
			[LimitAmount] ,
			[RCSaidProp] ,
			[Reserves] ,
			[IsDeleted] ,
			[DownLoadMe],
			[UpLoadMe],
			[AdminComments],
			[DateLastUpdated],
			[UpdateByUserID],
			[FarmersUnitRowID]
		FROM	PolicyLimits
		WHERE 	[PolicyLimitsID] = @PolicyLimitsID
	END
	ELSE
	BEGIN
		IF @bDelete =1
		BEGIN
			UPDATE PolicyLimits SET 	[IsDeleted]=	(CASE 	WHEN [IsDeleted]= 0
										THEN 1
										ELSE 0
										END
									),
							[DateLastUpdated]=@DateLastUpdated ,
							[UpdateByUserID]=@UID
			WHERE PolicyLimitsID = @PolicyLimitsID
		
		END
		ELSE 
		BEGIN
			IF @bAdd=1
			BEGIN
				INSERT INTO PolicyLimits(
							[AssignmentsID] ,
							[ID],
							[IDAssignments],
							[ClassTypeID] ,
							[LimitAmount] ,
							[RCSaidProp] ,
							[Reserves] ,
							[IsDeleted] ,
							[DownLoadMe],
							[UpLoadMe],
							[AdminComments],
							[DateLastUpdated],
							[UpdateByUserID],
							[FarmersUnitRowID]
						)
				SELECT 			@AssignmentsID AS [AssignmentsID] ,
							@ID As [ID],
							@IDAssignments As [IDAssignments],
							@ClassTypeID As [ClassTypeID] ,
							@LimitAmount As [LimitAmount] ,
							@RCSaidProp As [RCSaidProp] ,
							@Reserves As [Reserves] ,
							@IsDeleted As [IsDeleted] ,
							@DownLoadMe As [DownLoadMe],
							@UpLoadMe As [UpLoadMe],
							@AdminComments As [AdminComments],
							@DateLastUpdated As [DateLastUpdated],
							@UID As [UpdateByUserID],
							@FarmersUnitRowID As [FarmersUnitRowID] 
					
			END
			ELSE 
			BEGIN
				IF @bEdit=1
				BEGIN
					Update PolicyLimits Set 	[AssignmentsID] = @AssignmentsID ,
									[ID] = @ID,
									[IDAssignments] = @IDAssignments,
									[ClassTypeID] = @ClassTypeID ,
									[LimitAmount] = @LimitAmount ,
									[RCSaidProp] =  @RCSaidProp,
									[Reserves] =  @Reserves,
									[IsDeleted] =  @IsDeleted,
									[DownLoadMe] = @DownLoadMe,
									[UpLoadMe] = @UpLoadMe,
									[AdminComments] = @AdminComments,
									[DateLastUpdated] = @DateLastUpdated  ,
									[UpdateByUserID] = @UID,
									[FarmersUnitRowID] = @FarmersUnitRowID   
					WHERE 	PolicyLimitsID = @PolicyLimitsID
				END
			END
		END
	END
END
ELSE
BEGIN
	IF @GroupBy Is Not Null
	BEGIN
		SET @sSQL = 'SELECT RetPolicyLimits.' + @GroupBy + ' '
	END
	ELSE
	BEGIN
		SET @sSQL = 'SELECT RetPolicyLimits.* '
	END	
	SET @sSQL = @sSQL + 'FROM( ' 
	SET @sSQL = @sSQL + 'SELECT '
	SET @sSQL = @sSQL + '[PolicyLimitsID], '
	SET @sSQL = @sSQL + '[AssignmentsID] , '
	SET @sSQL = @sSQL + '[ID], '
	SET @sSQL = @sSQL + '[IDAssignments], '
	SET @sSQL = @sSQL + '[ClassTypeID] , '
	IF @bTableFieldsOnly = 0
	BEGIN
		SET @sSQL = @sSQL + '(SELECT	Class '
		SET @sSQL = @sSQL + 'FROM 	ClassType '
		SET @sSQL = @sSQL + 'WHERE 	ClassTypeID = S.[ClassTypeID]) As [ClassTypeClass],  '
		SET @sSQL = @sSQL + '(SELECT	Description '
		SET @sSQL = @sSQL + 'FROM 	ClassType '
		SET @sSQL = @sSQL + 'WHERE 	ClassTypeID = S.[ClassTypeID]) As [ClassTypeDescription],  '
	END
	SET @sSQL = @sSQL + '[LimitAmount] , '
	SET @sSQL = @sSQL + '[RCSaidProp] , '
	SET @sSQL = @sSQL + '[Reserves] , '
	SET @sSQL = @sSQL + '[IsDeleted] , '
	SET @sSQL = @sSQL + '[DownLoadMe], '
	SET @sSQL = @sSQL + '[UpLoadMe], '
	SET @sSQL = @sSQL + '[AdminComments], '
	SET @sSQL = @sSQL + '[DateLastUpdated] , '
	IF @bTableFieldsOnly = 0
	BEGIN
		SET @sSQL = @sSQL + '(SELECT	USERNAME '
		SET @sSQL = @sSQL + 'FROM 	USERS '
		SET @sSQL = @sSQL + 'WHERE 	USERSID = S.[UpdateByUserID]) As [UpdateByUserName],  '
	END
	SET @sSQL = @sSQL + '[UpdateByUserID],  '
	SET @sSQL = @sSQL + '[FarmersUnitRowID] '

	IF @bViewHistory Is Null
	BEGIN
		SET @sSQL = @sSQL + 'FROM PolicyLimits S '
		SET @sSQL = @sSQL + ') RetPolicyLimits '
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
		SET @sSQL = @sSQL + 'FROM PolicyLimitsHistory S '
		SET @sSQL = @sSQL + ') RetPolicyLimits '
		SET @sSQL = @sSQL + 'WHERE [DateLastUpdated] >= Convert(datetime,''' + @FromDate + ''') '
		SET @sSQL = @sSQL + 'AND [DateLastUpdated] <= Convert(datetime,''' + @ToDate + ''') '
		IF @PolicyLimitsID <> -1
		BEGIN
			SET @sSQL = @sSQL + 'AND [PolicyLimitsID] = ' + Cast(@PolicyLimitsID as varchar(10)) + ' '
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
				IF @bTableFieldsOnly = 0
				BEGIN
					SET @sSQL = @sSQL + 'ORDER BY [IsDeleted] Asc, [ClassTypeClass] '
				END
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

GRANT  EXECUTE  ON [dbo].[z_spsGetPolicyLimitsInfo]  TO [CFUSER]
GO

