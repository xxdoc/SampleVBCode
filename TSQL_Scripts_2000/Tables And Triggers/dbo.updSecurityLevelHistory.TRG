SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[updSecurityLevelHistory]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[updSecurityLevelHistory]
GO

CREATE TRIGGER updSecurityLevelHistory
ON dbo.SecurityLevel
FOR UPDATE
AS
INSERT INTO SecurityLevelHistory
	SELECT del.* 
	FROM DELETED del
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

