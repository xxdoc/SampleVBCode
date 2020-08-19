SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spsGetEvaluation]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[spsGetEvaluation]
GO

/******************************************************************
*
*    	Description:	Selects Chosen Evaluations From ECSSTATUS.MDB 
*	
*	Author: 	Dan Warner
*	Date: 		07/15/2003
*
******************************************************************/

CREATE PROCEDURE [dbo].[spsGetEvaluation] 
	@LAST VARCHAR(50),
	@FIRST VARCHAR(50)

AS
DECLARE @sSQL AS Varchar(1500)

Set @sSQL ='SELECT * '
Set @sSQL = @sSQL + 'FROM OpenQuery(MDB_ECSSTATUS,''Select * from AdjusterEvaluations  WHERE '
Set @sSql = @sSQL + 'AdjFirstName ='''+ @FIRST + ''' AND AdjLastName = ''' + @LAST +''' '') ' 


Exec (@sSQL)

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[spsGetEvaluation]  TO [CFUSER]
GO

