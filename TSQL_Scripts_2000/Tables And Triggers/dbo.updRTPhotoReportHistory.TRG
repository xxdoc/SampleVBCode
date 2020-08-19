SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[updRTPhotoReportHistory]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[updRTPhotoReportHistory]
GO

CREATE TRIGGER updRTPhotoReportHistory
ON dbo.RTPhotoReport
FOR UPDATE
AS
INSERT INTO RTPhotoReportHistory
	SELECT del.* 
	FROM DELETED del

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

