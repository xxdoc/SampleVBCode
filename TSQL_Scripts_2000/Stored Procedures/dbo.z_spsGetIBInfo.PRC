SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spsGetIBInfo]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spsGetIBInfo]
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
CREATE PROCEDURE [dbo].[z_spsGetIBInfo]
	@bHideDeleted		bit =1, 
	@UID			int,
	@IBID			Int=null,
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
	@IB00_lssn				int=null,
	@IB01_sSubToCarrier			varchar(100)=null,
	@IB02_sIBNumber				varchar(20)=null,
	@IB05_sLocation				varchar(50)=null,
	@IB05a_sState				varchar(10)=null,
	@IB06_dtDateClosed			datetime=null,
	@IB07_sAdjusterName			varchar(50)=null,
	@IB09_sSALN				varchar(20)=null,
	@IB10_sInsuredName			varchar(100)=null,
	@IB11_sLossLocation			varchar(254)=null,
	@IB12_dtDateOfLoss			varchar(20)=null,
	@IB13_cGrossLoss			money=null,
	@IB14_cDepreciation			money=null,
	@IB14a_sSupplement			int=null,
	@IB14b_sRebilled			int=null,
	@IB15_cDeductible			money=null,
	@IB15a_cLessExcessLimits		money=null,
	@IB15b_sExcessLimDesc			varchar(254)=null,
	@IB15c_cLessMiscellaneous		money=null,
	@IB15d_cMiscellaneousDesc		varchar(254)=null,
	@IB16_cNetClaim				money=null,
	@IB17_cServiceFee			money=null,
	@IB17a_cMiscServiceFee			money=null,
	@IB18_sServiceFeeComment		varchar(254)=null,
	@IB18a_sMiscServiceFeeComment		varchar(50)=null,
	@IB25_cServiceFeeSubTotal		money=null,
	@IB29a_sMiscExpenseFeeComment		varchar(50)=null,
	@IB29b_cMiscExpenseFee			money=null,
	@IB30_cTotalExpenses			money=null,
	@IB31_dTaxPercent			numeric(8, 3)=null,
	@IB32_cTaxAmount			money=null,
	@IB33_cTotalAdjustingFee		money=null,
	@IB33a_sAccountCode			varchar(50)=null,
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

IF @IBID Is Not Null And @bViewHistory is Null
BEGIN
	IF @bDelete=0 And @bAdd=0 And @bEdit=0
	BEGIN
		SELECT 
			[IBID] ,
			[AssignmentsID],
			[BillingCountID],
			[ID] ,
			[IDAssignments] ,
			[IDBillingCount],
			[IB00_lssn],
			[IB01_sSubToCarrier] ,
			[IB02_sIBNumber] ,
			[IB05_sLocation],
			[IB05a_sState] ,
			[IB06_dtDateClosed] ,
			[IB07_sAdjusterName],
			[IB09_sSALN] ,
			[IB10_sInsuredName] ,
			[IB11_sLossLocation],
			[IB12_dtDateOfLoss],
			[IB13_cGrossLoss],
			[IB14_cDepreciation] ,
			[IB14a_sSupplement] ,
			[IB14b_sRebilled] ,
			[IB15_cDeductible],
			[IB15a_cLessExcessLimits],
			[IB15b_sExcessLimDesc],
			[IB15c_cLessMiscellaneous],
			[IB15d_cMiscellaneousDesc],
			[IB16_cNetClaim] ,
			[IB17_cServiceFee],
			[IB17a_cMiscServiceFee],
			[IB18_sServiceFeeComment],
			[IB18a_sMiscServiceFeeComment],
			[IB25_cServiceFeeSubTotal],
			[IB29a_sMiscExpenseFeeComment],
			[IB29b_cMiscExpenseFee],
			[IB30_cTotalExpenses] ,
			[IB31_dTaxPercent],
			[IB32_cTaxAmount] ,
			[IB33_cTotalAdjustingFee],
			[IB33a_sAccountCode] ,
			[Void],
			[FeeByTime] ,
			[UseActivityTime],
			[DownLoadMe],
			[UpLoadMe] ,
			[AdminComments],
			[DateLastUpdated],
			[UpdateByUserID] 
		FROM	IB
		WHERE 	[IBID] = @IBID
	END
	ELSE
	BEGIN
		IF @bDelete =1
		BEGIN
			UPDATE IB SET 	[VOID]=	(CASE 	WHEN [VOID]= 0
										THEN 1
										ELSE 0
										END
									),
							[DateLastUpdated]=@DateLastUpdated ,
							[UpdateByUserID]=@UID
			WHERE IBID = @IBID
		
		END
		ELSE 
		BEGIN
			IF @bAdd=1
			BEGIN
				INSERT INTO IB(
							[AssignmentsID],
							[BillingCountID],
							[ID] ,
							[IDAssignments] ,
							[IDBillingCount],
							[IB00_lssn],
							[IB01_sSubToCarrier] ,
							[IB02_sIBNumber] ,
							[IB05_sLocation],
							[IB05a_sState] ,
							[IB06_dtDateClosed] ,
							[IB07_sAdjusterName],
							[IB09_sSALN] ,
							[IB10_sInsuredName] ,
							[IB11_sLossLocation],
							[IB12_dtDateOfLoss],
							[IB13_cGrossLoss],
							[IB14_cDepreciation] ,
							[IB14a_sSupplement] ,
							[IB14b_sRebilled] ,
							[IB15_cDeductible],
							[IB15a_cLessExcessLimits],
							[IB15b_sExcessLimDesc],
							[IB15c_cLessMiscellaneous],
							[IB15d_cMiscellaneousDesc],
							[IB16_cNetClaim] ,
							[IB17_cServiceFee],
							[IB17a_cMiscServiceFee],
							[IB18_sServiceFeeComment],
							[IB18a_sMiscServiceFeeComment],
							[IB25_cServiceFeeSubTotal],
							[IB29a_sMiscExpenseFeeComment],
							[IB29b_cMiscExpenseFee],
							[IB30_cTotalExpenses] ,
							[IB31_dTaxPercent],
							[IB32_cTaxAmount] ,
							[IB33_cTotalAdjustingFee],
							[IB33a_sAccountCode] ,
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
							@ID AS [ID] ,
							@IDAssignments AS [IDAssignments] ,
							@IDBillingCount AS [IDBillingCount],
							@IB00_lssn AS [IB00_lssn],
							@IB01_sSubToCarrier AS [IB01_sSubToCarrier] ,
							@IB02_sIBNumber AS [IB02_sIBNumber] ,
							@IB05_sLocation AS [IB05_sLocation],
							@IB05a_sState AS [IB05a_sState] ,
							@IB06_dtDateClosed AS [IB06_dtDateClosed] ,
							@IB07_sAdjusterName AS [IB07_sAdjusterName],
							@IB09_sSALN AS [IB09_sSALN] ,
							@IB10_sInsuredName AS [IB10_sInsuredName] ,
							@IB11_sLossLocation AS [IB11_sLossLocation],
							@IB12_dtDateOfLoss AS [IB12_dtDateOfLoss],
							@IB13_cGrossLoss AS [IB13_cGrossLoss],
							@IB14_cDepreciation AS [IB14_cDepreciation] ,
							@IB14a_sSupplement AS [IB14a_sSupplement] ,
							@IB14b_sRebilled AS [IB14b_sRebilled] ,
							@IB15_cDeductible AS [IB15_cDeductible],
							@IB15a_cLessExcessLimits AS [IB15a_cLessExcessLimits],
							@IB15b_sExcessLimDesc AS [IB15b_sExcessLimDesc],
							@IB15c_cLessMiscellaneous AS [IB15c_cLessMiscellaneous],
							@IB15d_cMiscellaneousDesc AS [IB15d_cMiscellaneousDesc],
							@IB16_cNetClaim AS [IB16_cNetClaim] ,
							@IB17_cServiceFee AS [IB17_cServiceFee],
							@IB17a_cMiscServiceFee AS [IB17a_cMiscServiceFee],
							@IB18_sServiceFeeComment AS [IB18_sServiceFeeComment],
							@IB18a_sMiscServiceFeeComment AS [IB18a_sMiscServiceFeeComment],
							@IB25_cServiceFeeSubTotal AS [IB25_cServiceFeeSubTotal],
							@IB29a_sMiscExpenseFeeComment AS [IB29a_sMiscExpenseFeeComment],
							@IB29b_cMiscExpenseFee AS [IB29b_cMiscExpenseFee],
							@IB30_cTotalExpenses AS [IB30_cTotalExpenses] ,
							@IB31_dTaxPercent AS [IB31_dTaxPercent],
							@IB32_cTaxAmount AS [IB32_cTaxAmount] ,
							@IB33_cTotalAdjustingFee AS [IB33_cTotalAdjustingFee],
							@IB33a_sAccountCode AS [IB33a_sAccountCode] ,
							@Void AS [Void],
							@FeeByTime AS [FeeByTime] ,
							@UseActivityTime AS [UseActivityTime],
							@DownLoadMe As [DownLoadMe],
							@UpLoadMe As [UpLoadMe],
							@AdminComments As [AdminComments],
							@DateLastUpdated As [DateLastUpdated],
							@UID As [UpdateByUserID] 
				--Need to update ID if it is null
				Update IB Set	[ID] = [IBID],
						[IDAssignments] = [AssignmentsID],
						[IDBillingCount] = [BillingCountID]
				WHERE	[ID] Is Null
			END
			ELSE 
			BEGIN
				IF @bEdit=1
				BEGIN
					Update IB Set 	
							[AssignmentsID] = @AssignmentsID,
							[BillingCountID] = @BillingCountID,
							[ID] =  @ID,
							[IDAssignments] = @IDAssignments ,
							[IDBillingCount] = @IDBillingCount,
							[IB00_lssn] = @IB00_lssn,
							[IB01_sSubToCarrier] = @IB01_sSubToCarrier ,
							[IB02_sIBNumber] = @IB02_sIBNumber ,
							[IB05_sLocation] = @IB05_sLocation,
							[IB05a_sState] =  @IB05a_sState,
							[IB06_dtDateClosed] = @IB06_dtDateClosed ,
							[IB07_sAdjusterName] = @IB07_sAdjusterName,
							[IB09_sSALN] = @IB09_sSALN ,
							[IB10_sInsuredName] = @IB10_sInsuredName ,
							[IB11_sLossLocation] = @IB11_sLossLocation,
							[IB12_dtDateOfLoss] = @IB12_dtDateOfLoss,
							[IB13_cGrossLoss] = @IB13_cGrossLoss,
							[IB14_cDepreciation] = @IB14_cDepreciation ,
							[IB14a_sSupplement] = @IB14a_sSupplement ,
							[IB14b_sRebilled] = @IB14b_sRebilled ,
							[IB15_cDeductible] = @IB15_cDeductible,
							[IB15a_cLessExcessLimits] = @IB15a_cLessExcessLimits,
							[IB15b_sExcessLimDesc] = @IB15b_sExcessLimDesc ,
							[IB15c_cLessMiscellaneous] = @IB15c_cLessMiscellaneous,
							[IB15d_cMiscellaneousDesc] = @IB15d_cMiscellaneousDesc ,
							[IB16_cNetClaim] = @IB16_cNetClaim ,
							[IB17_cServiceFee] = @IB17_cServiceFee,
							[IB17a_cMiscServiceFee] = @IB17a_cMiscServiceFee,
							[IB18_sServiceFeeComment] = @IB18_sServiceFeeComment,
							[IB18a_sMiscServiceFeeComment] = @IB18a_sMiscServiceFeeComment,
							[IB25_cServiceFeeSubTotal] = @IB25_cServiceFeeSubTotal,
							[IB29a_sMiscExpenseFeeComment] = @IB29a_sMiscExpenseFeeComment,
							[IB29b_cMiscExpenseFee] = @IB29b_cMiscExpenseFee,
							[IB30_cTotalExpenses] = @IB30_cTotalExpenses ,
							[IB31_dTaxPercent] = @IB31_dTaxPercent,
							[IB32_cTaxAmount] = @IB32_cTaxAmount ,
							[IB33_cTotalAdjustingFee] = @IB33_cTotalAdjustingFee,
							[IB33a_sAccountCode] = @IB33a_sAccountCode ,
							[Void] = @Void,
							[FeeByTime] = @FeeByTime ,
							[UseActivityTime] = @UseActivityTime,
							[DownLoadMe] = @DownLoadMe,
							[UpLoadMe] = @UpLoadMe,
							[AdminComments] = @AdminComments,
							[DateLastUpdated] = @DateLastUpdated  ,
							[UpdateByUserID] = @UID   
					WHERE 	IBID = @IBID
				END
			END
		END
	END
END
ELSE
BEGIN
	IF @GroupBy Is Not Null
	BEGIN
		SET @sSQL = 'SELECT RetIB.' + @GroupBy + ' '
	END
	ELSE
	BEGIN
		SET @sSQL = 'SELECT RetIB.* '
	END	
	SET @sSQL = @sSQL + 'FROM( ' 
	SET @sSQL = @sSQL + 'SELECT '
	IF @bViewHistory Is Not Null
	BEGIN
		SET @sSQL = @sSQL + '[IBHistoryID], '
	END
	SET @sSQL = @sSQL + '[IBID], '
	SET @sSQL = @sSQL + '[AssignmentsID], '
	SET @sSQL = @sSQL + '[BillingCountID], '
	SET @sSQL = @sSQL + '[ID] , '
	SET @sSQL = @sSQL + '[IDAssignments] , '
	SET @sSQL = @sSQL + '[IDBillingCount], '
	SET @sSQL = @sSQL + '[IB00_lssn], '
	SET @sSQL = @sSQL + '[IB01_sSubToCarrier] , '
	SET @sSQL = @sSQL + '[IB02_sIBNumber] , '
	SET @sSQL = @sSQL + '[IB05_sLocation], '
	SET @sSQL = @sSQL + '[IB05a_sState] , '
	SET @sSQL = @sSQL + '[IB06_dtDateClosed] , '
	SET @sSQL = @sSQL + '[IB07_sAdjusterName], '
	SET @sSQL = @sSQL + '[IB09_sSALN] , '
	SET @sSQL = @sSQL + '[IB10_sInsuredName] , '
	SET @sSQL = @sSQL + '[IB11_sLossLocation], '
	SET @sSQL = @sSQL + '[IB12_dtDateOfLoss], '
	SET @sSQL = @sSQL + '[IB13_cGrossLoss], '
	SET @sSQL = @sSQL + '[IB14_cDepreciation] , '
	SET @sSQL = @sSQL + '[IB14a_sSupplement] , '
	SET @sSQL = @sSQL + '[IB14b_sRebilled] , '
	SET @sSQL = @sSQL + '[IB15_cDeductible], '
	SET @sSQL = @sSQL + '[IB15a_cLessExcessLimits], '
	SET @sSQL = @sSQL + '[IB15b_sExcessLimDesc], '
	SET @sSQL = @sSQL + '[IB15c_cLessMiscellaneous], '
	SET @sSQL = @sSQL + '[IB15d_cMiscellaneousDesc], '
	SET @sSQL = @sSQL + '[IB16_cNetClaim] , '
	SET @sSQL = @sSQL + '[IB17_cServiceFee], '
	SET @sSQL = @sSQL + '[IB17a_cMiscServiceFee], '
	SET @sSQL = @sSQL + '[IB18_sServiceFeeComment], '
	SET @sSQL = @sSQL + '[IB18a_sMiscServiceFeeComment], '
	SET @sSQL = @sSQL + '[IB25_cServiceFeeSubTotal], '
	SET @sSQL = @sSQL + '[IB29a_sMiscExpenseFeeComment], '
	SET @sSQL = @sSQL + '[IB29b_cMiscExpenseFee], '
	SET @sSQL = @sSQL + '[IB30_cTotalExpenses] , '
	SET @sSQL = @sSQL + '[IB31_dTaxPercent], '
	SET @sSQL = @sSQL + '[IB32_cTaxAmount] , '
	SET @sSQL = @sSQL + '[IB33_cTotalAdjustingFee], '
	SET @sSQL = @sSQL + '[IB33a_sAccountCode] , '
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
		SET @sSQL = @sSQL + 'FROM IB S '
		SET @sSQL = @sSQL + ') RetIB '
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
		SET @sSQL = @sSQL + 'FROM IBHistory S '
		SET @sSQL = @sSQL + ') RetIB '
		SET @sSQL = @sSQL + 'WHERE [IBID] Is Not Null '
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
	print 'This is the sql= ' + @sSQL
	EXEC(@sSQL)
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[z_spsGetIBInfo]  TO [CFUSER]
GO

