SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[updPackageItemHistory]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[updPackageItemHistory]
GO

CREATE TRIGGER updPackageItemHistory
ON dbo.PackageItem
FOR UPDATE
AS
INSERT INTO PackageItemHistory
	SELECT del.* 
	FROM DELETED del

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

