SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[updRTIBFeeHistory]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[updRTIBFeeHistory]
GO

CREATE TRIGGER updRTIBFeeHistory
ON dbo.RTIBFee
FOR UPDATE
AS
INSERT INTO RTIBFeeHistory
	SELECT del.* 
	FROM DELETED del

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

