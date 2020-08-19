SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spsGetRSSpecial]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[spsGetRSSpecial]
GO

CREATE PROCEDURE [dbo].[spsGetRSSpecial] 
	@LinkName varchar(20),
	@CurEffADJSite Varchar(20)
AS
DECLARE @sSQL AS Varchar(1000)

Set @sSQL ='Select SPECIAL.* '
Set @sSQL = @sSQL + 'FROM OpenQuery(' + @LinkName + ',''Select * from Special WHERE special.adj_cat ="'+ @CurEffADJSite + '" '') ' 
Set @sSQL = @sSQL + 'As SPECIAL '

Exec (@sSQL)

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[spsGetRSSpecial]  TO [CFUSER]
GO

