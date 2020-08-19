SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spsGetRTChecksInfo]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spsGetRTChecksInfo]
GO

/******************************************************************
*
*    	Description:	IB Info
*			
*	
*	Author: 	Brad Skidmore
*	Date: 		5/10/2004
*
******************************************************************/

CREATE PROCEDURE [dbo].[z_spsGetRTChecksInfo]
	@bHideDeleted		bit =1, 
	@UID			int,
	@RTChecksID		Int=null,
	@OrderBy		varchar(500)=null,
	@GroupBy		varchar(500)=null,
	@SearchBy		varchar(1000)=null,
--updating the TypeOfLoss Table
	@bDelete		bit=0,
	@bAdd			bit=0,
	@bEdit			bit=0,
--Update variables
	@AssignmentsID	 			int=null,
	@BillingCountID				int=null,
	@ID					int=null,
	@IDAssignments				int=null,
	@IDBillingCount				int=null,
	@CheckNum				int=null,
	@RT42_ClassOfLossID			int=null,
	@RT43_TypeOfLossID			int=null,
	@RT50_sInsuredPayeeName			varchar(100)=null,
	@RT51_sPayeeNames			varchar(100)=null,
	@RT52_sAddress				varchar(255)=null,
	@RT53_cAmountOfCheck			numeric(10, 2)=null,
	@RT54_CompanyCatSpecID			int=null,
	@tempCHeckName				varchar(20)=null,
	@PrintOnIB				bit=null,
	@IsDeleted				bit=null,
	@DownLoadMe				bit=null,
	@UpLoadMe				bit=null,
	@AdminComments				varchar(1000)=null,
	@DateLastUpdated			datetime=null,
--ViewHistory
	@bViewHistory		bit=NULL,
	@FromDate		varchar(50)=Null,
	@ToDate			varchar(50)=Null,
--IncludeFlags
	@bTableFieldsOnly	bit=0
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

IF @RTChecksID Is Not Null And @bViewHistory is Null
BEGIN
	IF @bDelete=0 And @bAdd=0 And @bEdit=0
	BEGIN
		SELECT 
			[RTChecksID] ,
			[AssignmentsID],
			[BillingCountID],
			[ID] ,
			[IDAssignments],
			[IDBillingCount],
			[CheckNum],
			[RT42_ClassOfLossID],
			[RT43_TypeOfLossID],
			[RT50_sInsuredPayeeName],
			[RT51_sPayeeNames],
			[RT52_sAddress] ,
			[RT53_cAmountOfCheck],
			[RT54_CompanyCatSpecID],
			[tempCHeckName] ,
			[PrintOnIB] ,
			[IsDeleted] ,
			[DownLoadMe],
			[UpLoadMe] ,
			[AdminComments],
			[DateLastUpdated],
			[UpdateByUserID] 
		FROM	RTChecks
		WHERE 	[RTChecksID] = @RTChecksID
	END
	ELSE
	BEGIN
		IF @bDelete =1
		BEGIN
			UPDATE RTChecks SET 	[IsDeleted]=	(CASE 	WHEN [IsDeleted]= 0
										THEN 1
										ELSE 0
										END
									),
							[DateLastUpdated]=@DateLastUpdated ,
							[UpdateByUserID]=@UID
			WHERE RTChecksID = @RTChecksID
		
		END
		ELSE 
		BEGIN
			IF @bAdd=1
			BEGIN
				INSERT INTO RTChecks(
							[AssignmentsID],
							[BillingCountID],
							[ID] ,
							[IDAssignments],
							[IDBillingCount],
							[CheckNum],
							[RT42_ClassOfLossID],
							[RT43_TypeOfLossID],
							[RT50_sInsuredPayeeName],
							[RT51_sPayeeNames],
							[RT52_sAddress] ,
							[RT53_cAmountOfCheck],
							[RT54_CompanyCatSpecID],
							[tempCHeckName] ,
							[PrintOnIB] ,
							[IsDeleted] ,
							[DownLoadMe],
							[UpLoadMe] ,
							[AdminComments],
							[DateLastUpdated],
							[UpdateByUserID] 
						)
				SELECT 			@AssignmentsID As [AssignmentsID],
							@BillingCountID As [BillingCountID],
							@ID As [ID] ,
							@IDAssignments As [IDAssignments],
							@IDBillingCount As [IDBillingCount],
							@CheckNum As [CheckNum],
							@RT42_ClassOfLossID As [RT42_ClassOfLossID],
							@RT43_TypeOfLossID As [RT43_TypeOfLossID],
							@RT50_sInsuredPayeeName As [RT50_sInsuredPayeeName],
							@RT51_sPayeeNames As [RT51_sPayeeNames],
							@RT52_sAddress As [RT52_sAddress] ,
							@RT53_cAmountOfCheck As [RT53_cAmountOfCheck],
							@RT54_CompanyCatSpecID As [RT54_CompanyCatSpecID],
							@tempCHeckName As [tempCHeckName] ,
							@PrintOnIB As [PrintOnIB] ,
							@IsDeleted As [IsDeleted] ,
							@DownLoadMe As [DownLoadMe],
							@UpLoadMe As [UpLoadMe] ,
							@AdminComments As [AdminComments],
							@DateLastUpdated As [DateLastUpdated], 
							@UID As [UpdateByUserID] 
				--Need to update ID if it is null
				Update RTChecks Set 	[ID] = [RTCHecksID],
							[IDAssignments] = [AssignmentsID],
							[IDBillingCount] = [BillingCountID]
				WHERE	[ID] Is Null
			END
			ELSE 
			BEGIN
				IF @bEdit=1
				BEGIN
					Update RTChecks Set 	
							[AssignmentsID] = @AssignmentsID,
							[BillingCountID] = @BillingCountID,
							[ID] = @ID ,
							[IDAssignments] = @IDAssignments,
							[IDBillingCount] = @IDBillingCount,
							[CheckNum] = @CheckNum,
							[RT42_ClassOfLossID] = @RT42_ClassOfLossID,
							[RT43_TypeOfLossID] = @RT43_TypeOfLossID,
							[RT50_sInsuredPayeeName] = @RT50_sInsuredPayeeName,
							[RT51_sPayeeNames] = @RT51_sPayeeNames,
							[RT52_sAddress] = @RT52_sAddress ,
							[RT53_cAmountOfCheck] = @RT53_cAmountOfCheck,
							[RT54_CompanyCatSpecID] = @RT54_CompanyCatSpecID,
							[tempCHeckName] = @tempCHeckName ,
							[PrintOnIB]  = @PrintOnIB,
							[IsDeleted]  = @IsDeleted,
							[DownLoadMe] = @DownLoadMe,
							[UpLoadMe]  = @UpLoadMe,
							[AdminComments] = @AdminComments,
							[DateLastUpdated] = @DateLastUpdated,
							[UpdateByUserID] = @UID
					WHERE 	RTChecksID = @RTChecksID
				END
			END
		END
	END
END
ELSE
BEGIN
	IF @GroupBy Is Not Null
	BEGIN
		SET @sSQL = 'SELECT RetRTChecks.' + @GroupBy + ' '
	END
	ELSE
	BEGIN
		SET @sSQL = 'SELECT RetRTChecks.* '
	END	
	SET @sSQL = @sSQL + 'FROM( ' 
	SET @sSQL = @sSQL + 'SELECT '
	IF @bViewHistory Is Not Null
	BEGIN
		SET @sSQL = @sSQL + '[RTChecksHistoryID], '
	END
	SET @sSQL = @sSQL + '[RTChecksID], '
	SET @sSQL = @sSQL + '[AssignmentsID], '
	SET @sSQL = @sSQL + '[BillingCountID], '
	SET @sSQL = @sSQL + '[ID] , '
	SET @sSQL = @sSQL + '[IDAssignments] , '
	SET @sSQL = @sSQL + '[IDBillingCount], '
	SET @sSQL = @sSQL + '[CheckNum], '
	SET @sSQL = @sSQL + '[RT42_ClassOfLossID], '
	SET @sSQL = @sSQL + '[RT43_TypeOfLossID], '
	SET @sSQL = @sSQL + '[RT50_sInsuredPayeeName], '
	SET @sSQL = @sSQL + '[RT51_sPayeeNames], '
	SET @sSQL = @sSQL + '[RT52_sAddress] , '
	SET @sSQL = @sSQL + '[RT53_cAmountOfCheck], '
	SET @sSQL = @sSQL + '[RT54_CompanyCatSpecID], '
	SET @sSQL = @sSQL + '[tempCHeckName] , '
	SET @sSQL = @sSQL + '[PrintOnIB] , '
	SET @sSQL = @sSQL + '[IsDeleted] , '
	SET @sSQL = @sSQL + '[DownLoadMe], '
	SET @sSQL = @sSQL + '[UpLoadMe] , '
	SET @sSQL = @sSQL + '[AdminComments], '
	SET @sSQL = @sSQL + '[DateLastUpdated] , ' 
	IF @bTableFieldsOnly = 0
	BEGIN
		SET @sSQL = @sSQL + '(SELECT	USERNAME '
		SET @sSQL = @sSQL + 'FROM 	USERS '
		SET @sSQL = @sSQL + 'WHERE 	USERSID = S.[UpdateByUserID]) As [UpdateByUserName],  '
	END
	SET @sSQL = @sSQL + '[UpdateByUserID]  '

	IF @bViewHistory Is Null
	BEGIN
		SET @sSQL = @sSQL + 'FROM RTChecks S '
		SET @sSQL = @sSQL + ') RetRTChecks '
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
		SET @sSQL = @sSQL + 'FROM RTChecksHistory S '
		SET @sSQL = @sSQL + ') RetRTChecks '
		SET @sSQL = @sSQL + 'WHERE [RTChecksID] Is Not Null '
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
					SET @sSQL = @sSQL + 'ORDER BY [IsDeleted] Asc '
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
			IF @OrderBy Is Null OR @OrderBy = 'Null'
			BEGIN
				IF @bTableFieldsOnly = 0
				BEGIN
					SET @sSQL = @sSQL + 'ORDER BY [IsDeleted] Asc '
				END
			END
			ELSE
			BEGIN
				SET @sSQL = @sSQL + 'ORDER BY  [IsDeleted] Asc, ' + @OrderBy + ' '
			END
		END
	END
	--print 'This is the sql= ' + @sSQL
	EXEC(@sSQL)
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[z_spsGetRTChecksInfo]  TO [CFUSER]
GO

