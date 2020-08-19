SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[updRTActivityLogInfoHistory]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[updRTActivityLogInfoHistory]
GO

CREATE TRIGGER updRTActivityLogInfoHistory
ON dbo.RTActivityLogInfo
FOR UPDATE
AS
INSERT INTO RTActivityLogInfoHistory
	SELECT del.* 
	FROM DELETED del

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

