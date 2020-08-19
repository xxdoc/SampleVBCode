SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[updAssignmentTypeHistory]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[updAssignmentTypeHistory]
GO

CREATE TRIGGER updAssignmentTypeHistory
ON dbo.AssignmentType
FOR UPDATE
AS
INSERT INTO AssignmentTypeHistory
	SELECT del.* 
	FROM DELETED del

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

