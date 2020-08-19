SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spuUpdateECSADJ]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[spuUpdateECSADJ]
GO

/******************************************************************
*
*    	Description:	Updates ECSADJ from the web
*	
*	Author: 	Dan Warner
*	Date: 		06/11/2003
*
******************************************************************/
CREATE PROCEDURE [dbo].[spuUpdateECSADJ] 
	@LinkName		 varchar(20),
	@NICKNAME 		VARCHAR(20),
	@SPOUSENAME 	VARCHAR(20),
	@LASTCONTACT 	VARCHAR(20),
	@HOMEPHONE 	VARCHAR(20),
	@BUSINESSPHONE 	VARCHAR(20),
	@PAGER 		VARCHAR(20),
	@EMAIL 		VARCHAR(50),
	@AVAILABLE 	VARCHAR(5),
	@PRIMARYCONTACT VARCHAR(20),
	@RECNUM		VARCHAR(12),
	@SFIND		VARCHAR(20),
	@SPANISH	VARCHAR(20)
	
AS
DECLARE @sSQL AS Varchar(2000)

Set @sSQL = 'UPDATE openquery(' + @LinkName + ',''Select NICKNAME, SPOUSENAME, SPANISH, LAST_CONTA,  HOME_PHONE, BUSN_PHONE, PAGER, E_MAIL, AVAILABLE, PRIMARY_CO, SPECIAL_FI, UID  from ECSADJ WHERE UID = ' + @RECNUM + ''') '
Set @sSQL = @sSQL + 'SET NICKNAME = ''' + @NICKNAME + ''', '
Set @sSQL = @sSQL + 'SPOUSENAME = ''' + @SPOUSENAME + ''', '
Set @sSQL = @sSQL + 'SPANISH = ''' + @SPANISH + ''', '
Set @sSQL = @sSQL + 'LAST_CONTA = ''' + @LASTCONTACT + ''', '
Set @sSQL = @sSQL + 'HOME_PHONE = ''' + @HOMEPHONE + ''', '
Set @sSQL = @sSQL + 'BUSN_PHONE = ''' + @BUSINESSPHONE + ''', '
Set @sSQL = @sSQL + 'PAGER = ''' + @PAGER + ''', '
Set @sSQL = @sSQL + 'E_MAIL = ''' + @EMAIL + ''', '
Set @sSQL = @sSQL + 'AVAILABLE = ''' + @AVAILABLE + ''', '
Set @sSQL = @sSQL + 'PRIMARY_CO = ''' + @PRIMARYCONTACT + ''', '
Set @sSQL = @sSQL + 'SPECIAL_FI = ''' + @SFIND + ''' '
Set @sSQL = @sSQL + ' WHERE UID = ' + @RECNUM + ' ' 

Exec (@sSQL)

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[spuUpdateECSADJ]  TO [CFUSER]
GO

