SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spuCreateUIDECSADJ]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[spuCreateUIDECSADJ]
GO

/******************************************************************
*
*    	Description:	Creates UID if none exists
*	
*	Author: 	Brad Skidmore
*	Date: 		10/31/2003
*
******************************************************************/
CREATE PROCEDURE [dbo].[spuCreateUIDECSADJ] 
	@LinkName		 varchar(20)
AS
DECLARE @sSQL AS Varchar(8000)
DECLARE @MaxUID as int
DECLARE @LastFirst as Varchar(50)

IF @LinkName = 'CLAIMDEV'
BEGIN
	SET @MaxUID = (SELECT Max(UID) as MyMaxUID FROM openquery(CLAIMDEV,'SELECT UID from ecsadj ') )
	SET @LastFirst = (SELECT * FROM openquery(CLAIMDEV,'SELECT TOP 1 LastFirst from ecsadj Order By LastFirst where UID = 0 or UID is Null'))
END
ELSE IF @LinkName = 'CLAIMPRODV1'
BEGIN
	SET @MaxUID = (SELECT Max(UID) as MyMaxUID FROM openquery(CLAIMPRODV1,'SELECT UID from ecsadj ') )
	SET @LastFirst = (SELECT * FROM openquery(CLAIMPRODV1,'SELECT TOP 1 LastFirst from ecsadj Order By LastFirst where UID = 0 or UID is Null'))
END

--Increment the UID by 1
SET @MaxUID = @MaxUID + 1

Set @sSQL = 'UPDATE openquery(' + @LinkName + ',''Select TOP 1 UID, LastFirst FROM ECSADJ Order By LastFirst where UID = 0 or UID is Null'') '
Set @sSQL = @sSQL + 'SET UID = ' + Cast(@MaxUID as varchar(10)) + ' '
Set @sSQL = @sSQL + ' WHERE LastFirst = ''' + @LastFirst + ''' ' 
Exec (@sSQL)

--Then return which Record UID was updated
SET @sSQL = ''
SET @sSQL = @sSQL + 'SELECT '
SET @sSQL = @sSQL + 'ECSADJ.* '
SET @sSQL = @sSQL + 'FROM	OpenQuery(' + @LinkName + ', ''SELECT 	'
SET @sSQL = @sSQL + 'ECSADJ.FIRST_NAME, '
SET @sSQL = @sSQL + 'ECSADJ.LAST_NAME, '
SET @sSQL = @sSQL + 'ECSADJ.UID '
SET @sSQL = @sSQL + 'FROM 	ECSADJ '
SET @sSQL = @sSQL + 'WHERE UID = ' + cast(@MaxUID as varchar(10)) + ''') '
SET @sSQL = @sSQL + 'as ECSADJ '
Exec (@sSQL)

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[spuCreateUIDECSADJ]  TO [CFUSER]
GO

