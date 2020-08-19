SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spsHttpServiceGroup]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spsHttpServiceGroup]
GO

/******************************************************************
*
*    	Description:	Get Http Service Group
*			
*	
*	Author: 	Brad Skidmore
*	Date: 		4/26/2005
*
******************************************************************/
CREATE PROCEDURE [dbo].[z_spsHttpServiceGroup]
	@DaysBack			int=10
AS
BEGIN
	SELECT 	[Service]
	FROM 	Httplog
	WHERE [LogTime] >= getDate()-@DaysBack
	Group BY [Service]
	Order By [Service]
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[z_spsHttpServiceGroup]  TO [CFUSER]
GO

