SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[updRTIndemnityHistory]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[updRTIndemnityHistory]
GO

CREATE TRIGGER updRTIndemnityHistory
ON dbo.RTIndemnity
FOR UPDATE
AS
INSERT INTO RTIndemnityHistory
	SELECT del.* 
	FROM DELETED del

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

