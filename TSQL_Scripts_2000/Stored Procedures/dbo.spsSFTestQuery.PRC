SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spsSFTestQuery]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[spsSFTestQuery]
GO

CREATE PROCEDURE [dbo].[spsSFTestQuery] 
	@LinkName varchar(20),
	@First Varchar(50),
	@Last  Varchar(50)
AS
DECLARE @sSQL AS Varchar(1000)

Set @sSQL ='Select SFTEST.* '
Set @sSQL = @sSQL + 'FROM OpenQuery(' + @LinkName + ',''Select * from SFTEST'
Set @sSQL = @sSQL + ' WHERE SFTEST.Last_name ="'+ @last + '" AND SFTEST.First_name = "' + @FIRST + '" '') ' 
Set @sSQL = @sSQL + 'As SFTEST '

Exec (@sSQL)

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[spsSFTestQuery]  TO [CFUSER]
GO

