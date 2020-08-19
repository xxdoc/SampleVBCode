SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spsGetNOTEID]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[spsGetNOTEID]
GO

/******************************************************************
*
*    	Description:	Selects All data From Notes.dbf So we can 
			get a unique ID for the NOTEID field. 
*	
*	Author: 	Dan Warner
*	Date: 		06/24/2003
*
******************************************************************/

CREATE PROCEDURE [dbo].[spsGetNOTEID] 
	@LinkName varchar(20)
AS
DECLARE @sSQL AS Varchar(1500)

Set @sSQL ='SELECT '
Set @sSQL =@sSQL + '* FROM openquery(' + @LinkName + ',''SELECT Notes.NOTEID '
Set @sSQL =@sSQL + 'FROM Notes '') As NOTES '
Set @sSQL =@sSQL + 'ORDER BY NOTES.NoteID ASC'

Exec (@sSQL)

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[spsGetNOTEID]  TO [CFUSER]
GO

