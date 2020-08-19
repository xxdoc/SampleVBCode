SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[updFeeScheduleFeeTypesHistory]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[updFeeScheduleFeeTypesHistory]
GO

CREATE TRIGGER updFeeScheduleFeeTypesHistory
ON dbo.FeeScheduleFeeTypes
FOR UPDATE
AS
INSERT INTO FeeScheduleFeeTypesHistory
	SELECT del.* 
	FROM DELETED del

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

