SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[updClassOfLossHistory]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[updClassOfLossHistory]
GO

CREATE TRIGGER updClassOfLossHistory
ON dbo.ClassOfLoss
FOR UPDATE
AS
INSERT INTO ClassOfLossHistory
	SELECT del.* 
	FROM DELETED del

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

