SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[updAdjusterUsersUpdatesHistory]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[updAdjusterUsersUpdatesHistory]
GO

CREATE TRIGGER updAdjusterUsersUpdatesHistory
ON dbo.AdjusterUsersUpdates
FOR UPDATE
AS
INSERT INTO AdjusterUsersUpdatesHistory
	SELECT del.* 
	FROM DELETED del

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

