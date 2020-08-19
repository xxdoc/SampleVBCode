SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spsBuildHistoryItem]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spsBuildHistoryItem]
GO

/******************************************************************
*
*    	Description:	Build Dynamically From
			supplied historyTableName, ItemNAme,
			And Identities
*			
*	
*	Author: 	Brad Skidmore
*	Date: 		10/15/2003
*
******************************************************************/

CREATE PROCEDURE [dbo].[z_spsBuildHistoryItem]
	@bHideDeleted		bit =1, 
	@UID			int,
	@FromDate		varchar(50)=Null,
	@ToDate			varchar(50)=Null,
	@HistoryTableName	varchar(50)=null,
	@HistoryItemName	varchar(50)=null,
	@ID1Name 		varchar(50)=null,
	@ID1			int=0,
	@ID2Name		varchar(50)=null,
	@ID2			int=0,
	@ID3Name		varchar(50)=null,
	@ID3			int=0
AS
DECLARE @sSQL varchar(8000)
SET @sSQL = ''
--check history date span
IF IsDate(@FromDate) = 0 
	SET @FromDate = Convert(datetime,dbo.FormatDate(GetDate(),'mm/dd/yyy'))
ELSE
	SET @FromDate = Convert(datetime,dbo.FormatDate(@FromDate,'mm/dd/yyy'))
IF IsDate(@ToDate) = 0
	SET @ToDate = GetDate()
ELSE
	SET @ToDate = Convert(datetime,dbo.FormatDate(@ToDate,'mm/dd/yyy')) + ' 23:59:59'

--Build The History Item 
IF @HistoryTableName is Not Null And @HistoryItemName Is Not Null
BEGIN
	SET @sSQL = @sSQL + 'SELECT BuildHistory.* FROM( '
	SET @sSQL = @sSQL + 'SELECT 	BHis.' + @HistoryTableName + 'ID, '
	SET @sSQL = @sSQL + '		BHis.' + @HistoryItemName + ', '
	SET @sSQL = @sSQL + '		BHis.DateLastUpdated, '
	SET @sSQL = @sSQL + '(		SELECT 	UserName '
	SET @sSQL = @sSQL + '		FROM	Users '
	SET @sSQL = @sSQL + '		WHERE 	UsersID = BHis.UpdateByUserID '
	SET @sSQL = @sSQL + ') 		As UpdateByUserName '
	SET @sSQL = @sSQL + 'FROM ' + @HistoryTableName + ' BHis '
	IF @ID1Name Is Not Null 
	BEGIN
		SET @sSQL = @sSQL + 'WHERE BHis.' + @ID1Name + ' = ' + Cast(@ID1 As Varchar(10)) + ' '
	END
	IF @ID1Name Is Not Null And @ID2Name Is Not Null 
	BEGIN
		SET @sSQL = @sSQL + 'AND BHis.' + @ID2Name + ' = ' + Cast(@ID2 As Varchar(10)) + ' '
	END
	IF @ID1Name Is Not Null And @ID2Name Is Not Null And @ID3Name Is Not Null 
	BEGIN
		SET @sSQL = @sSQL + 'AND BHis.' + @ID3Name + ' = ' + Cast(@ID3 As Varchar(10)) + ' '
	END
	--Date Span
	SET @sSQL = @sSQL + 'AND BHis.DateLastUpdated >= Convert(datetime,''' + @FromDate + ''') '
	SET @sSQL = @sSQL + 'AND BHis.DateLastUpdated <= Convert(datetime,''' + @ToDate + ''') '
	SET @sSQL = @sSQL + ') BuildHistory '
	--Sort
	SET @sSQL = @sSQL + 'ORDER BY BuildHistory.DateLastUpdated '

	--Debug
	Print Len(@sSQL)
	Print @sSQL

	--Execute here
	EXEC(@sSQL)
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[z_spsBuildHistoryItem]  TO [CFUSER]
GO

