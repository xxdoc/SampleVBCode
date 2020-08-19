SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spsGetRTIBInfo]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spsGetRTIBInfo]
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
CREATE PROCEDURE [dbo].[z_spsGetRTIBInfo]
	@bHideDeleted		bit =1, 
	@UID			int,
	@AssignmentsID		Int=null,
	@OrderBy		varchar(500)=null,
	@GroupBy		varchar(500)=null,
	@SearchBy		varchar(1000)=null,
--updating the TypeOfLoss Table
	@bDelete		bit=0,
	@bAdd			bit=0,
	@bEdit			bit=0,
--Update variables
	@BillingCountID				int=null,
	@IDAssignments				int=null,
	@IDBillingCount				int=null,
	@RT00_lssn				int=null,
	@RT01_sSubToCarrier			varchar(100)=null,
	@RT02_sIBNumber				varchar(20)=null,
	@RT05_sLocation				varchar(50)=null,
	@RT05a_sState				varchar(10)=null,
	@RT06_dtDateClosed			datetime=null,
	@RT07_sAdjusterName			varchar(50)=null,
	@RT09_sSALN				varchar(20)=null,
	@RT10_sInsuredName			varchar(100)=null,
	@RT11_sLossLocation			varchar(254)=null,
	@RT12_dtDateOfLoss			varchar(20)=null,
	@RT13_cGrossLoss			money=null,
	@RT14_cDepreciation			money=null,
	@RT14a_sSupplement			int=null,
	@RT14b_sRebilled			int=null,
	@RT15_cDeductible			money=null,
	@RT15a_cLessExcessLimits		money=null,
	@RT15b_sExcessLimDesc			varchar(254)=null,
	@RT15c_cLessMiscellaneous		money=null,
	@RT15d_cMiscellaneousDesc		varchar(254)=null,
	@RT16_cNetClaim				money=null,
	@RT17_cServiceFee			money=null,
	@RT17a_cMiscServiceFee			money=null,
	@RT18_sServiceFeeComment		varchar(254)=null,
	@RT18a_sMiscServiceFeeComment		varchar(50)=null,
	@RT25_cServiceFeeSubTotal		money=null,
	@RT29a_sMiscExpenseFeeComment		varchar(50)=null,
	@RT29b_cMiscExpenseFee			money=null,
	@RT30_cTotalExpenses			money=null,
	@RT31_dTaxPercent			numeric(8, 3)=null,
	@RT32_cTaxAmount			money=null,
	@RT33_cTotalAdjustingFee		money=null,
	@RT33a_sAccountCode			varchar(50)=null,
	@Void					bit=null,
	@FeeByTime				bit=null,
	@UseActivityTime			bit=null,
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

IF @AssignmentsID Is Not Null And @bViewHistory is Null
BEGIN
	IF @bDelete=0 And @bAdd=0 And @bEdit=0
	BEGIN
		SELECT 
			[AssignmentsID],
			[BillingCountID],
			[IDAssignments] ,
			[IDBillingCount],
			[RT00_lssn],
			[RT01_sSubToCarrier] ,
			[RT02_sIBNumber] ,
			[RT05_sLocation],
			[RT05a_sState] ,
			[RT06_dtDateClosed] ,
			[RT07_sAdjusterName],
			[RT09_sSALN] ,
			[RT10_sInsuredName] ,
			[RT11_sLossLocation],
			[RT12_dtDateOfLoss],
			[RT13_cGrossLoss],
			[RT14_cDepreciation] ,
			[RT14a_sSupplement] ,
			[RT14b_sRebilled] ,
			[RT15_cDeductible],
			[RT15a_cLessExcessLimits],
			[RT15b_sExcessLimDesc],
			[RT15c_cLessMiscellaneous],
			[RT15d_cMiscellaneousDesc],
			[RT16_cNetClaim] ,
			[RT17_cServiceFee],
			[RT17a_cMiscServiceFee],
			[RT18_sServiceFeeComment],
			[RT18a_sMiscServiceFeeComment],
			[RT25_cServiceFeeSubTotal],
			[RT29a_sMiscExpenseFeeComment],
			[RT29b_cMiscExpenseFee],
			[RT30_cTotalExpenses] ,
			[RT31_dTaxPercent],
			[RT32_cTaxAmount] ,
			[RT33_cTotalAdjustingFee],
			[RT33a_sAccountCode] ,
			[Void],
			[FeeByTime] ,
			[UseActivityTime],
			[DownLoadMe],
			[UpLoadMe] ,
			[AdminComments],
			[DateLastUpdated],
			[UpdateByUserID] 
		FROM	RTIB
		WHERE 	[AssignmentsID] = @AssignmentsID
	END
	ELSE
	BEGIN
		IF @bDelete =1
		BEGIN
			UPDATE RTIB SET 	[VOID]=	(CASE 	WHEN [VOID]= 0
										THEN 1
										ELSE 0
										END
									),
							[DateLastUpdated]=@DateLastUpdated ,
							[UpdateByUserID]=@UID
			WHERE AssignmentsID = @AssignmentsID
		
		END
		ELSE 
		BEGIN
			IF @bAdd=1
			BEGIN
				INSERT INTO RTIB(
							[AssignmentsID],
							[BillingCountID],
							[IDAssignments] ,
							[IDBillingCount],
							[RT00_lssn],
							[RT01_sSubToCarrier] ,
							[RT02_sIBNumber] ,
							[RT05_sLocation],
							[RT05a_sState] ,
							[RT06_dtDateClosed] ,
							[RT07_sAdjusterName],
							[RT09_sSALN] ,
							[RT10_sInsuredName] ,
							[RT11_sLossLocation],
							[RT12_dtDateOfLoss],
							[RT13_cGrossLoss],
							[RT14_cDepreciation] ,
							[RT14a_sSupplement] ,
							[RT14b_sRebilled] ,
							[RT15_cDeductible],
							[RT15a_cLessExcessLimits],
							[RT15b_sExcessLimDesc],
							[RT15c_cLessMiscellaneous],
							[RT15d_cMiscellaneousDesc],
							[RT16_cNetClaim] ,
							[RT17_cServiceFee],
							[RT17a_cMiscServiceFee],
							[RT18_sServiceFeeComment],
							[RT18a_sMiscServiceFeeComment],
							[RT25_cServiceFeeSubTotal],
							[RT29a_sMiscExpenseFeeComment],
							[RT29b_cMiscExpenseFee],
							[RT30_cTotalExpenses] ,
							[RT31_dTaxPercent],
							[RT32_cTaxAmount] ,
							[RT33_cTotalAdjustingFee],
							[RT33a_sAccountCode] ,
							[Void],
							[FeeByTime] ,
							[UseActivityTime],
							[DownLoadMe],
							[UpLoadMe] ,
							[AdminComments],
							[DateLastUpdated],
							[UpdateByUserID] 
						)
				SELECT 			@AssignmentsID AS [AssignmentsID],
							@BillingCountID AS [BillingCountID],
							@IDAssignments AS [IDAssignments] ,
							@IDBillingCount AS [IDBillingCount],
							@RT00_lssn AS [RT00_lssn],
							@RT01_sSubToCarrier AS [RT01_sSubToCarrier] ,
							@RT02_sIBNumber AS [RT02_sIBNumber] ,
							@RT05_sLocation AS [RT05_sLocation],
							@RT05a_sState AS [RT05a_sState] ,
							@RT06_dtDateClosed AS [RT06_dtDateClosed] ,
							@RT07_sAdjusterName AS [RT07_sAdjusterName],
							@RT09_sSALN AS [RT09_sSALN] ,
							@RT10_sInsuredName AS [RT10_sInsuredName] ,
							@RT11_sLossLocation AS [RT11_sLossLocation],
							@RT12_dtDateOfLoss AS [RT12_dtDateOfLoss],
							@RT13_cGrossLoss AS [RT13_cGrossLoss],
							@RT14_cDepreciation AS [RT14_cDepreciation] ,
							@RT14a_sSupplement AS [RT14a_sSupplement] ,
							@RT14b_sRebilled AS [RT14b_sRebilled] ,
							@RT15_cDeductible AS [RT15_cDeductible],
							@RT15a_cLessExcessLimits AS [RT15a_cLessExcessLimits],
							@RT15b_sExcessLimDesc AS [RT15b_sExcessLimDesc],
							@RT15c_cLessMiscellaneous AS [RT15c_cLessMiscellaneous],
							@RT15d_cMiscellaneousDesc AS [RT15d_cMiscellaneousDesc],
							@RT16_cNetClaim AS [RT16_cNetClaim] ,
							@RT17_cServiceFee AS [RT17_cServiceFee],
							@RT17a_cMiscServiceFee AS [RT17a_cMiscServiceFee],
							@RT18_sServiceFeeComment AS [RT18_sServiceFeeComment],
							@RT18a_sMiscServiceFeeComment AS [RT18a_sMiscServiceFeeComment],
							@RT25_cServiceFeeSubTotal AS [RT25_cServiceFeeSubTotal],
							@RT29a_sMiscExpenseFeeComment AS [RT29a_sMiscExpenseFeeComment],
							@RT29b_cMiscExpenseFee AS [RT29b_cMiscExpenseFee],
							@RT30_cTotalExpenses AS [RT30_cTotalExpenses] ,
							@RT31_dTaxPercent AS [RT31_dTaxPercent],
							@RT32_cTaxAmount AS [RT32_cTaxAmount] ,
							@RT33_cTotalAdjustingFee AS [RT33_cTotalAdjustingFee],
							@RT33a_sAccountCode AS [RT33a_sAccountCode] ,
							@Void AS [Void],
							@FeeByTime AS [FeeByTime] ,
							@UseActivityTime AS [UseActivityTime],
							@DownLoadMe As [DownLoadMe],
							@UpLoadMe As [UpLoadMe],
							@AdminComments As [AdminComments],
							@DateLastUpdated As [DateLastUpdated],
							@UID As [UpdateByUserID] 
					
			END
			ELSE 
			BEGIN
				IF @bEdit=1
				BEGIN
					Update RTIB Set 	
							[AssignmentsID] = @AssignmentsID,
							[BillingCountID] = @BillingCountID,
							[IDAssignments] = @IDAssignments ,
							[IDBillingCount] = @IDBillingCount,
							[RT00_lssn] = @RT00_lssn,
							[RT01_sSubToCarrier] = @RT01_sSubToCarrier ,
							[RT02_sIBNumber] = @RT02_sIBNumber ,
							[RT05_sLocation] = @RT05_sLocation,
							[RT05a_sState] =  @RT05a_sState,
							[RT06_dtDateClosed] = @RT06_dtDateClosed ,
							[RT07_sAdjusterName] = @RT07_sAdjusterName,
							[RT09_sSALN] = @RT09_sSALN ,
							[RT10_sInsuredName] = @RT10_sInsuredName ,
							[RT11_sLossLocation] = @RT11_sLossLocation,
							[RT12_dtDateOfLoss] = @RT12_dtDateOfLoss,
							[RT13_cGrossLoss] = @RT13_cGrossLoss,
							[RT14_cDepreciation] = @RT14_cDepreciation ,
							[RT14a_sSupplement] = @RT14a_sSupplement ,
							[RT14b_sRebilled] = @RT14b_sRebilled ,
							[RT15_cDeductible] = @RT15_cDeductible,
							[RT15a_cLessExcessLimits] = @RT15a_cLessExcessLimits,
							[RT15b_sExcessLimDesc] = @RT15b_sExcessLimDesc ,
							[RT15c_cLessMiscellaneous] = @RT15c_cLessMiscellaneous,
							[RT15d_cMiscellaneousDesc] = @RT15d_cMiscellaneousDesc ,
							[RT16_cNetClaim] = @RT16_cNetClaim ,
							[RT17_cServiceFee] = @RT17_cServiceFee,
							[RT17a_cMiscServiceFee] = @RT17a_cMiscServiceFee,
							[RT18_sServiceFeeComment] = @RT18_sServiceFeeComment,
							[RT18a_sMiscServiceFeeComment] = @RT18a_sMiscServiceFeeComment,
							[RT25_cServiceFeeSubTotal] = @RT25_cServiceFeeSubTotal,
							[RT29a_sMiscExpenseFeeComment] = @RT29a_sMiscExpenseFeeComment,
							[RT29b_cMiscExpenseFee] = @RT29b_cMiscExpenseFee,
							[RT30_cTotalExpenses] = @RT30_cTotalExpenses ,
							[RT31_dTaxPercent] = @RT31_dTaxPercent,
							[RT32_cTaxAmount] = @RT32_cTaxAmount ,
							[RT33_cTotalAdjustingFee] = @RT33_cTotalAdjustingFee,
							[RT33a_sAccountCode] = @RT33a_sAccountCode ,
							[Void] = @Void,
							[FeeByTime] = @FeeByTime ,
							[UseActivityTime] = @UseActivityTime,
							[DownLoadMe] = @DownLoadMe,
							[UpLoadMe] = @UpLoadMe,
							[AdminComments] = @AdminComments,
							[DateLastUpdated] = @DateLastUpdated  ,
							[UpdateByUserID] = @UID   
					WHERE 	AssignmentsID = @AssignmentsID
				END
			END
		END
	END
END
ELSE
BEGIN
	IF @GroupBy Is Not Null
	BEGIN
		SET @sSQL = 'SELECT RetRTIB.' + @GroupBy + ' '
	END
	ELSE
	BEGIN
		SET @sSQL = 'SELECT RetRTIB.* '
	END	
	SET @sSQL = @sSQL + 'FROM( ' 
	SET @sSQL = @sSQL + 'SELECT '
	SET @sSQL = @sSQL + '[AssignmentsID], '
	SET @sSQL = @sSQL + '[BillingCountID], '
	SET @sSQL = @sSQL + '[IDAssignments] , '
	SET @sSQL = @sSQL + '[IDBillingCount], '
	SET @sSQL = @sSQL + '[RT00_lssn], '
	SET @sSQL = @sSQL + '[RT01_sSubToCarrier] , '
	SET @sSQL = @sSQL + '[RT02_sIBNumber] , '
	SET @sSQL = @sSQL + '[RT05_sLocation], '
	SET @sSQL = @sSQL + '[RT05a_sState] , '
	SET @sSQL = @sSQL + '[RT06_dtDateClosed] , '
	SET @sSQL = @sSQL + '[RT07_sAdjusterName], '
	SET @sSQL = @sSQL + '[RT09_sSALN] , '
	SET @sSQL = @sSQL + '[RT10_sInsuredName] , '
	SET @sSQL = @sSQL + '[RT11_sLossLocation], '
	SET @sSQL = @sSQL + '[RT12_dtDateOfLoss], '
	SET @sSQL = @sSQL + '[RT13_cGrossLoss], '
	SET @sSQL = @sSQL + '[RT14_cDepreciation] , '
	SET @sSQL = @sSQL + '[RT14a_sSupplement] , '
	SET @sSQL = @sSQL + '[RT14b_sRebilled] , '
	SET @sSQL = @sSQL + '[RT15_cDeductible], '
	SET @sSQL = @sSQL + '[RT15a_cLessExcessLimits], '
	SET @sSQL = @sSQL + '[RT15b_sExcessLimDesc], '
	SET @sSQL = @sSQL + '[RT15c_cLessMiscellaneous], '
	SET @sSQL = @sSQL + '[RT15d_cMiscellaneousDesc], '
	SET @sSQL = @sSQL + '[RT16_cNetClaim] , '
	SET @sSQL = @sSQL + '[RT17_cServiceFee], '
	SET @sSQL = @sSQL + '[RT17a_cMiscServiceFee], '
	SET @sSQL = @sSQL + '[RT18_sServiceFeeComment], '
	SET @sSQL = @sSQL + '[RT18a_sMiscServiceFeeComment], '
	SET @sSQL = @sSQL + '[RT25_cServiceFeeSubTotal], '
	SET @sSQL = @sSQL + '[RT29a_sMiscExpenseFeeComment], '
	SET @sSQL = @sSQL + '[RT29b_cMiscExpenseFee], '
	SET @sSQL = @sSQL + '[RT30_cTotalExpenses] , '
	SET @sSQL = @sSQL + '[RT31_dTaxPercent], '
	SET @sSQL = @sSQL + '[RT32_cTaxAmount] , '
	SET @sSQL = @sSQL + '[RT33_cTotalAdjustingFee], '
	SET @sSQL = @sSQL + '[RT33a_sAccountCode] , '
	SET @sSQL = @sSQL + '[Void], '
	SET @sSQL = @sSQL + '[FeeByTime] , '
	SET @sSQL = @sSQL + '[UseActivityTime], '
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
		SET @sSQL = @sSQL + 'FROM RTIB S '
		SET @sSQL = @sSQL + ') RetRTIB '
		IF @bHideDeleted = 1
		BEGIN
			SET @sSQL = @sSQL + 'WHERE VOID = 0 '
		END
		ELSE
		BEGIN
			SET @sSQL = @sSQL + 'WHERE VOID Is Not Null '
		END
		--Searching
		IF @SearchBy Is Not Null
		BEGIN
			SET @sSQL = @sSQL + @SearchBy + ' '
		END
	END
	ELSE
	BEGIN
		SET @sSQL = @sSQL + 'FROM RTIBHistory S '
		SET @sSQL = @sSQL + ') RetRTIB '
		SET @sSQL = @sSQL + 'WHERE [DateLastUpdated] >= Convert(datetime,''' + @FromDate + ''') '
		SET @sSQL = @sSQL + 'AND [DateLastUpdated] <= Convert(datetime,''' + @ToDate + ''') '
		IF @AssignmentsID <> -1
		BEGIN
			SET @sSQL = @sSQL + 'AND [AssignmentsID] = ' + Cast(@AssignmentsID as varchar(10)) + ' '
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
					SET @sSQL = @sSQL + 'ORDER BY [VOID] Asc '
				END
			END
			ELSE
			BEGIN
				SET @sSQL = @sSQL + 'ORDER BY  [VOID] Asc, ' + @OrderBy + ' '
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
	print @sSQL
	EXEC(@sSQL)
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[z_spsGetRTIBInfo]  TO [CFUSER]
GO

