SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spsGetFTPIP]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spsGetFTPIP]
GO

/******************************************************************
*
*    	Description:	Get FTP IP ADDRESS
*			
*	
*	Author: 	Brad Skidmore
*	Date: 		4/20/2005
*
******************************************************************/
CREATE PROCEDURE [dbo].[z_spsGetFTPIP]
	@UID			int,
	@SearchTarget		VarChar(2)='',
	@LogTime		DateTime
AS
BEGIN
	SELECT TOP 1 [ClientHost] As [IPAddress] 
	FROM 	FTPLog
	WHERE 	[Target] Like '%' + @SearchTarget + '%'
	AND [LogTime] >= @LogTime
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[z_spsGetFTPIP]  TO [CFUSER]
GO

