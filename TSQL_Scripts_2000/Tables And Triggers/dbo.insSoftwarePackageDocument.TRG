SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[insSoftwarePackageDocument]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[insSoftwarePackageDocument]
GO

CREATE TRIGGER insSoftwarePackageDocument
ON dbo.SoftwarePackageDocument 
AFTER INSERT
AS

	UPDATE Document set 	Document.SPVersion = SP.SPVersion,
			    	Document.SPVersionBase =	(CASE 	WHEN Document.SPVersionBase IS Null
								  	THEN SP.SPVersion
									ELSE Document.SPVersionBase
									END
								)
						
	FROM  SoftwarePackage SP  	INNER JOIN INSERTED INS On SP.SoftwarePackageID = INS.SoftwarePackageID 
					INNER JOIN Document  ON INS.DocumentID = Document.DocumentID
	WHERE INS.DocumentID = Document.DocumentID

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

