SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spsGetLastUserSynch]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spsGetLastUserSynch]
GO

/******************************************************************
*
*    	Description:	Get Last SuccessFull User FTP SYNCH
*			
*	
*	Author: 	Brad Skidmore
*	Date: 		10/22/2005
*
******************************************************************/
CREATE PROCEDURE [dbo].[z_spsGetLastUserSynch]
	@UserName 	Varchar(50),
	@DaysToLookBack int=7
AS
BEGIN
	DECLARE @sUserNameSynch As Varchar(100)
	SET @sUserNameSynch = '%/USER_FOLDERS/' + @UserName + '/synch_%'
	SELECT TOP 1 	[ClientHost] As [IPAddress],
		 	[LogTime]
	FROM 	FTPLog
	WHERE [LogTime] >= (GetDate() - @DaysToLookBack)
	AND target like @sUserNameSynch
order by logtime desc
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[z_spsGetLastUserSynch]  TO [CFUSER]
GO

