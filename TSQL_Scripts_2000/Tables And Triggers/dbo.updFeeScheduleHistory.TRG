SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[updFeeScheduleHistory]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[updFeeScheduleHistory]
GO

CREATE TRIGGER updFeeScheduleHistory
ON dbo.FeeSchedule
FOR UPDATE
AS
INSERT INTO FeeScheduleHistory
	SELECT del.* 
	FROM DELETED del

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

