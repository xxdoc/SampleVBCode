SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spsGetFirstLastECSADJ]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[spsGetFirstLastECSADJ]
GO

/******************************************************************
*
*    	Description:	Selects Get First and Last name from ECSADJ
*	
*	Author: 		Dan Warner
*	Date: 		06/10/2003
*
******************************************************************/


CREATE PROCEDURE [dbo].[spsGetFirstLastECSADJ] 
	@LinkName varchar(20),
	@auto varchar(3)='no'

AS
DECLARE @sSQL AS Varchar(1500)
SET @sSQL = ''
SET @sSQL = @sSQL + 'SELECT '
SET @sSQL = @sSQL + 'ECSADJ.* '
SET @sSQL = @sSQL + 'FROM	OpenQuery(' + @LinkName + ', ''SELECT 	'
SET @sSQL = @sSQL + 'ECSADJ.FIRST_NAME, '
SET @sSQL = @sSQL + 'ECSADJ.LAST_NAME, '
SET @sSQL = @sSQL + 'ECSADJ.UID, ecsadj.AUTO '
SET @sSQL = @sSQL + 'FROM 	ECSADJ'') '
SET @sSQL = @sSQL + 'as ECSADJ '
IF @auto='yes'
BEGIN
	SET @sSQL = @sSQL + ' where ECSADJ.[AUTO] =''Yes'' '
END
SET @sSQL = @sSQL + 'ORDER BY Last_name '

Exec (@sSQL)

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[spsGetFirstLastECSADJ]  TO [CFUSER]
GO

