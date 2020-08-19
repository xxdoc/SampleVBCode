SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[updTypeOfLossHistory]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[updTypeOfLossHistory]
GO

CREATE TRIGGER updTypeOfLossHistory
ON dbo.TypeOfLoss
FOR UPDATE
AS
INSERT INTO TypeOfLossHistory
	SELECT del.* 
	FROM DELETED del

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

