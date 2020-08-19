SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spdDeleteNote]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[spdDeleteNote]
GO

/******************************************************************
*
*    	Description:	Deletes a note from Notes.dbf 
*	@RECNUM 	= passed in NoteID of Note to delete
*	
*	Author: 	Dan Warner
*	Date: 		06/25/2003
*
******************************************************************/

CREATE PROCEDURE [dbo].[spdDeleteNote] 
	@LinkName varchar(20),
	@RECNUM VARCHAR(10)

AS
DECLARE @sSQL AS Varchar(1500)

Set @sSQL ='DELETE '
Set @sSQL =@sSQL + 'FROM openquery(' + @LinkName + ',''SELECT * '
Set @sSQL =@sSQL + 'FROM NOTES WHERE NOTES.NOTEID = ' + @RECNUM + ''')'

Exec (@sSQL)

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[spdDeleteNote]  TO [CFUSER]
GO

