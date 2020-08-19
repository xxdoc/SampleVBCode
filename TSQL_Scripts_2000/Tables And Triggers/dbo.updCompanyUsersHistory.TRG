SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[updCompanyUsersHistory]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[updCompanyUsersHistory]
GO

CREATE TRIGGER updCompanyUsersHistory
ON dbo.CompanyUsers
FOR UPDATE
AS
INSERT INTO CompanyUsersHistory
	SELECT del.* 
	FROM DELETED del

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

