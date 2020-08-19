SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spsWindPoolQuery]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[spsWindPoolQuery]
GO

CREATE PROCEDURE [dbo].[spsWindPoolQuery] 
	@LinkName varchar(20),
	@LastFirst Varchar(50)

AS

DECLARE @sSQL AS Varchar(1000)

Set @sSQL ='Select WindPool.* '
Set @sSQL = @sSQL + 'FROM OpenQuery(' + @LinkName + ',''Select * from WindPool WHERE WindPool.Adjuster_n = "'+ @lastfirst + '" '') ' 
Set @sSQL = @sSQL + 'As WindPool '

Exec (@sSQL)

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[spsWindPoolQuery]  TO [CFUSER]
GO

