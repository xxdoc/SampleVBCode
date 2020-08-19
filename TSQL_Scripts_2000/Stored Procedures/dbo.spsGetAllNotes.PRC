SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spsGetAllNotes]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[spsGetAllNotes]
GO

/******************************************************************
*
*    	Description:	Selects All data From Notes.dbf 
*	
*	Author: 	Dan Warner
*	Date: 		06/23/2003
*
******************************************************************/

CREATE PROCEDURE [dbo].[spsGetAllNotes] 
	@LinkName varchar(20),
	@RECNUM VARCHAR(10)

AS
DECLARE @sSQL AS Varchar(1500)

Set @sSQL ='SELECT '
Set @sSQL =@sSQL + '* FROM openquery(' + @LinkName + ',''SELECT ECSADJ.*, Notes.* '
Set @sSQL =@sSQL + 'FROM ECSADJ INNER JOIN Notes ON ECSADJ.UID = Notes.ADJID WHERE NOTES.ADJID = ' + @RECNUM + ''') As NOTES '
Set @sSQL =@sSQL + 'ORDER BY NOTES.NoteID DESC'

Exec (@sSQL)

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[spsGetAllNotes]  TO [CFUSER]
GO

