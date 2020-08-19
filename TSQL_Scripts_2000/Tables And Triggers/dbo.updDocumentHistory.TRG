SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[updDocumentHistory]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[updDocumentHistory]
GO

CREATE TRIGGER updDocumentHistory
ON dbo.Document
AFTER UPDATE
AS
--See if there are any Version Changes
DECLARE @CountVersion int
DECLARE @CountVersionDate int

SET @CountVersion = 	(
				SELECT Count(DEL.Version) As CountOFVersion
				FROM DELETED DEL INNER JOIN INSERTED INS
				ON DEL.DocumentID = INS.DocumentID
				WHERE (DEL.Version <> INS.Version)
				AND DEL.SPVersion Is Not Null	
			)
SET @CountVersionDate = 	(
				SELECT Count(DEL.VersionDate) As CountOFVersionDate
				FROM DELETED DEL INNER JOIN INSERTED INS
				ON DEL.DocumentID = INS.DocumentID
				WHERE (DEL.VersionDate <> INS.VersionDate)
				AND INS.VersionDate Is Not Null 
				AND DEL.SPVersion Is Not Null	
				)

IF @CountVersion > 0
BEGIN
	INSERT INTO DocumentHistory
		SELECT DEL.* 
		FROM DELETED DEL INNER JOIN INSERTED INS
		ON DEL.DocumentID = INS.DocumentID
		WHERE (DEL.Version <> INS.Version )
		AND DEL.SPVersion Is Not Null	

	--Then update the Softwarepackage Table which will in turn Update All dependant Tables
	--with new SPVersion.
	UPDATE SoftwarePackage SET SoftwarePackage.SPVersion = (SELECT MAX(SoftwarePackage.SPVersion) +1  FROM SoftwarePackage),
					SoftwarePackage.VersionDate = GetDate()
	FROM SoftwarePackage 
	WHERE SoftwarePackage.SoftwarePackageID IN 	(
									SELECT SPD.SoftwarePackageID
									FROM SoftwarePackageDocument SPD INNER JOIN (INSERTED INS INNER JOIN DELETED DEL ON INS.DocumentID = DEL.DocumentID) ON SPD.DocumentID = INS.DocumentID								
									WHERE DEL.Version <> INS.Version
									AND SPD.IsDeleted =0
								)
	--As Well Update this records SPName, VersionDate and SPVersionBase
	UPDATE Document SET 	Document.SPName = Document.DocName + '_V' + cast(Document.Version As VarChar(10)) + '.exe',
					Document.VersionDate = GetDate(),
					Document.SPVersionBase = Document.SPVersion
	FROM Document  INNER JOIN INSERTED INS On Document.DocumentID = INS.DocumentID INNER JOIN DELETED DEL On INS.DocumentID = DEL.DocumentID
END

ELSE
	IF @CountVersionDate > 0
	BEGIN
		--Update the software package VersionDate
		UPDATE SoftwarePackage SET 	SoftwarePackage.VersionDate = getdate()
		FROM SoftwarePackage 
		WHERE SoftwarePackage.SoftwarePackageID IN 	(
										SELECT SPD.SoftwarePackageID
										FROM SoftwarePackageDocument SPD INNER JOIN (INSERTED INS INNER JOIN DELETED DEL ON INS.DocumentID = DEL.DocumentID) ON SPD.DocumentID = INS.DocumentID								
										WHERE (DEL.VersionDate <> INS.VersionDate)
										AND INS.VersionDate Is Not Null
										AND DEL.SPVersion Is Not Null	
										AND SPD.IsDeleted =0
									)
		--Use Server DATE Time for Version Date on Updates
		UPDATE Document SET 	Document.VersionDate = GetDate()
		FROM Document  INNER JOIN INSERTED INS On Document.DocumentID = INS.DocumentID INNER JOIN DELETED DEL On INS.DocumentID = DEL.DocumentID
		
		--Besure that the Version Date on Development matches the Version Date On Production
		UPDATE DevWebV2..Document SET  
					DevWebV2..Document.[VersionDate] = ProdWebV2..Document.[VersionDate],
					DevWebV2..Document.[DatelastUpdated] = ProdWebV2..Document.[DatelastUpdated] 
		FROM DevWebV2..Document 
			INNER JOIN INSERTED INS On DevWebV2..Document.[DocName] = INS.[DocName] 
			INNER JOIN DELETED DEL On INS.[DocumentID] = DEL.[DocumentID] 
			INNER JOIN ProdWebV2..Document On ProdWebV2..Document.[DocName] = DEL.[DocName]
	END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

