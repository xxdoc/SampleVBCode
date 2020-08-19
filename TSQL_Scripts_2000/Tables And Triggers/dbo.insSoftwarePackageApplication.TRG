SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[insSoftwarePackageApplication]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[insSoftwarePackageApplication]
GO

CREATE TRIGGER insSoftwarePackageApplication
ON dbo.SoftwarePackageApplication 
AFTER INSERT
AS

	UPDATE Application set 	Application.SPVersion = SP.SPVersion,
			    	Application.SPVersionBase =	(CASE 	WHEN Application.SPVersionBase IS Null
								  	THEN SP.SPVersion
									ELSE Application.SPVersionBase
									END
								)
						
	FROM  SoftwarePackage SP  	INNER JOIN INSERTED INS On SP.SoftwarePackageID = INS.SoftwarePackageID 
					INNER JOIN Application  ON INS.ApplicationID = Application.ApplicationID
	WHERE INS.ApplicationID = Application.ApplicationID

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

