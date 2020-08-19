SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spGetIISLog]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[spGetIISLog]
GO

/******************************************************************
*
*    	Description:	Return HTTP or FTP Log Info 
*	
*	Author: 		Brad Skidmore
*	Date: 		06/04/2003
*
******************************************************************/

CREATE PROCEDURE [dbo].[spGetIISLog] 
	@DateStart	Varchar(20),
	@DateEnd	Varchar(20),
	@SearchText	Varchar(50),
	@SearchIP	Varchar(20),
	@LogName	Varchar(20),
	@SortBy		Varchar(20),
	@SearchService	Varchar(20)
AS

Declare @sSQL 	Varchar(8000)

/*Build the Like Statements*/
/*Search text*/
If @SearchText <> ''
BEGIN
	Set @SearchText = '%' + @SearchText + '%'
END
ELSE
BEGIN
	Set @SearchText = '%'
END
/* Search IP */
If @SearchIP <> ''
BEGIN
	Set @SearchIP = '%' + @SearchIP + '%'
END
ELSE
BEGIN
	Set @SearchIP = '%'
END

/* Search Service */
If @SearchService <> ''
BEGIN
	Set @SearchService = '%' + @SearchService + '%'
END
ELSE
BEGIN
	Set @SearchService = '%'
END

/* Build Order BY */
IF @SortBy = 'Service' Set @SortBy = 'service ' 
ELSE IF @SortBy = 'User IP' Set @SortBy = 'clienthost ' 
ELSE IF @SortBy ='Proc Time' Set @SortBy = 'processingtime ' 
ELSE IF @SortBy = 'Bytes Recvd' Set @SortBy = 'bytesrecvd ' 
ELSE IF @SortBy = 'Bytes Sent' Set @SortBy = 'bytessent ' 
ELSE IF @SortBy = 'Svc Status' Set @SortBy = 'servicestatus ' 
ELSE IF @SortBy = 'Win Status' Set @SortBy = 'win32status ' 
ELSE IF @SortBy = 'Operation' Set @SortBy = 'operation ' 
ELSE IF @SortBy = 'Target' Set @SortBy = 'target ' 
ELSE IF @SortBy = 'Param' Set @SortBy = 'parameters ' 
ELSE Set @SortBy = 'logtime '



/* Build The SQL */

Set @sSQL = 'SELECT  service, '
Set @sSQL = @sSQL + 'clienthost, '
Set @sSQL = @sSQL + 'logtime, '
Set @sSQL = @sSQL + 'processingtime, '
Set @sSQL = @sSQL + 'bytesrecvd, '
Set @sSQL = @sSQL + 'bytessent, '
Set @sSQL = @sSQL + 'servicestatus, '
Set @sSQL = @sSQL + 'win32status, '
Set @sSQL = @sSQL + 'operation, '
IF @LogName ='HTTPLog'
BEGIN
	Set @sSQL = @sSQL + 'Replace(target,''<Rejected-By-UrlScan>'',''(Rejected-By-UrlScan)'') As target, '
	Set @sSQL = @sSQL + 'parameters '
END
ELSE
BEGIN
	Set @sSQL = @sSQL + 'target '
END
Set @sSQL = @sSQL + 'FROM	' + @LOGName + ' '
Set @sSQL = @sSQL + 'WHERE 	logtime >= Convert(DateTime,''' + @DateStart + ''') '
Set @sSQL = @sSQL + 'AND 	logtime <= Convert(DateTime,''' + @DateEnd + ''') '
Set @sSQL = @sSQL + 'AND 	target Like ''' + @SearchText + ''' '
Set @sSQL = @sSQL + 'AND 	clienthost Like ''' + @SearchIP + ''' '
Set @sSQL = @sSQL + 'AND	service Like ''' + @SearchService + ''' '
Set @sSQL = @sSQL + 'Order By ' + @SortBy + ' '

/* Execute it !*/
Exec (@sSQL)

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[spGetIISLog]  TO [CFUSER]
GO

