SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spDelClientCoAdjusterSpecInfo]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spDelClientCoAdjusterSpecInfo]
GO

/******************************************************************
*
*    	Description:	Delete Client Co Adjuster Spec Info
*			
*	
*	Author: 	Brad Skidmore
*	Date: 		10/28/2003
*	Date		2/17/2004 Added Date Span For Delete From History Table
******************************************************************/


CREATE PROCEDURE [dbo].[z_spDelClientCoAdjusterSpecInfo]
	@ClientCoAdjusterSpecID	int,
	@bRemoveHistoryOnly	int=0,
	@dtBeginHistoryDate	datetime=null,
	@dtEndHistoryDate	datetime=null
AS
IF @bRemoveHistoryOnly = 1 
BEGIN
	--Format the Begin and End History Dates
	IF IsDate(@dtBeginHistoryDate) = 0 
	BEGIN
		SET @dtBeginHistoryDate = Convert(datetime,dbo.FormatDate(GetDate(),'mm/dd/yyy'))
	END
	ELSE
	BEGIN
		SET @dtBeginHistoryDate = Convert(datetime,dbo.FormatDate(@dtBeginHistoryDate,'mm/dd/yyy'))
	END
	
	IF IsDate(@dtEndHistoryDate) = 0 
	BEGIN
		SET @dtEndHistoryDate = GetDate()
	END
	ELSE
	BEGIN
		SET @dtEndHistoryDate = Convert(datetime,dbo.FormatDate(@dtEndHistoryDate,'mm/dd/yyy')) + ' 23:59:59.999'
	END

	IF @ClientCoAdjusterSpecID = -1
	BEGIN
		DELETE FROM ClientCoAdjusterSpecHistory
		WHERE ZipCode Is null
		AND 	DateLastUpdated >= @dtBeginHistoryDate
		AND 	DateLastUpdated <= @dtEndHistoryDate	
	END
	ELSE
	BEGIN
		DELETE FROM ClientCoAdjusterSpecHistory
		WHERE ClientCoAdjusterSpecID = @ClientCoAdjusterSpecID
		And ZipCode Is Null
		AND 	DateLastUpdated >= @dtBeginHistoryDate
		AND 	DateLastUpdated <= @dtEndHistoryDate	
	END
END
ELSE
BEGIN
	DELETE FROM ClientCoAdjusterSpec
	WHERE 	ClientCoAdjusterSpecID = @ClientCoAdjusterSpecID
	And ZipCode Is Null
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[z_spDelClientCoAdjusterSpecInfo]  TO [CFUSER]
GO

