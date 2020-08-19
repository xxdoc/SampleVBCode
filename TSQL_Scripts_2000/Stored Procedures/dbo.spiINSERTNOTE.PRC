SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spiINSERTNOTE]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[spiINSERTNOTE]
GO

CREATE PROCEDURE [dbo].[spiINSERTNOTE]
	@LinkName varchar(20),
	@NOTEID Varchar(10),
	@RECNUM varchar(10),
	@TEXTBODY Varchar(254)
	

AS
DECLARE @sSQL AS Varchar(1500)

Set @sSQL ='INSERT '
Set @sSQL =@sSQL + 'openquery(' + @LinkName + ',''Select NOTEID, ADJID, NOTE_BODY, NOTE_DATE FROM NOTES'') '
Set @sSQL =@sSQL + 'VALUES (' + @NOTEID + ', ' + @RECNUM + ', ''' + @TEXTBODY +  ''', ''' + Cast(GetDate() As Varchar(30)) + ''' )'

Exec (@sSQL)

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[spiINSERTNOTE]  TO [CFUSER]
GO

