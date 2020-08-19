SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spsFTPServiceGroup]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spsFTPServiceGroup]
GO

/******************************************************************
*
*    	Description:	Get FTP Service Group
*			
*	
*	Author: 	Brad Skidmore
*	Date: 		4/26/2005
*
******************************************************************/
CREATE PROCEDURE [dbo].[z_spsFTPServiceGroup]
	@DaysBack			int=10
AS
BEGIN
	SELECT 	[Service]
	FROM 	Ftplog
	WHERE [LogTime] >= getDate()-@DaysBack
	Group BY [Service]
	Order By [Service]
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[z_spsFTPServiceGroup]  TO [CFUSER]
GO

