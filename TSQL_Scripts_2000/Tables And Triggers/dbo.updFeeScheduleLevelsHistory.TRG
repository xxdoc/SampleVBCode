SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[updFeeScheduleLevelsHistory]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[updFeeScheduleLevelsHistory]
GO

CREATE TRIGGER updFeeScheduleLevelsHistory
ON dbo.FeeScheduleLevels
FOR UPDATE
AS
INSERT INTO FeeScheduleLevelsHistory
	SELECT del.* 
	FROM DELETED del

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

