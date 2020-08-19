SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[updClientCompanyCatSpecHistory]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[updClientCompanyCatSpecHistory]
GO

CREATE TRIGGER updClientCompanyCatSpecHistory
ON dbo.ClientCompanyCatSpec
FOR UPDATE
AS
INSERT INTO ClientCompanyCatSpecHistory
	SELECT del.* 
	FROM DELETED del

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

