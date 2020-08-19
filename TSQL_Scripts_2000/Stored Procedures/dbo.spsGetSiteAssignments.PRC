SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spsGetSiteAssignments]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[spsGetSiteAssignments]
GO

CREATE PROCEDURE [dbo].[spsGetSiteAssignments] 
	@LinkName varchar(20),
	@CurADJ Varchar(20)
AS
DECLARE @sSQL AS Varchar(1000)

Set @sSQL ='Select SPECIAL.* '
Set @sSQL = @sSQL + 'FROM OpenQuery(' + @LinkName + ',''Select * from Special WHERE adjuster_n="'+ @CurADJ + '" '') ' 
Set @sSQL = @sSQL + 'As SPECIAL '

Exec (@sSQL)

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[spsGetSiteAssignments]  TO [CFUSER]
GO

