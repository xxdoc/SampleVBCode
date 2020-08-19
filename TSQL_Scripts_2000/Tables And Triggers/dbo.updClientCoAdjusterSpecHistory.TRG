SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[updClientCoAdjusterSpecHistory]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[updClientCoAdjusterSpecHistory]
GO


CREATE TRIGGER updClientCoAdjusterSpecHistory
ON dbo.ClientCoAdjusterSpec
FOR UPDATE
AS
INSERT INTO ClientCoAdjusterSpecHistory
	SELECT del.* 
	FROM DELETED del

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

