SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[updApplicationHistory]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[updApplicationHistory]
GO

CREATE TRIGGER updApplicationHistory
ON dbo.Application
AFTER UPDATE
AS
--See if there are any Version Changes
DECLARE @CountVersion int
DECLARE @CountVersionDate int

SET @CountVersion = 	(
				SELECT Count(DEL.Version) As CountOFVersion
				FROM DELETED DEL INNER JOIN INSERTED INS
				ON DEL.ApplicationID = INS.ApplicationID
				WHERE (DEL.Version <> INS.Version)
				AND DEL.SPVersion Is Not Null	
			)

SET @CountVersionDate = 	(
				SELECT Count(DEL.VersionDate) As CountOFVersionDate
				FROM DELETED DEL INNER JOIN INSERTED INS
				ON DEL.ApplicationID = INS.ApplicationID
				WHERE (DEL.VersionDate <> INS.VersionDate)
				AND INS.VersionDate Is Not Null
				AND DEL.SPVersion Is Not Null	
				)

IF @CountVersion > 0
BEGIN
	INSERT INTO ApplicationHistory
		SELECT DEL.* 
		FROM DELETED DEL INNER JOIN INSERTED INS
		ON DEL.ApplicationID = INS.ApplicationID
		WHERE (DEL.Version <> INS.Version)
		AND DEL.SPVersion Is Not Null	

	--Then update the Softwarepackage Table which will in turn Update All dependant Tables
	--with new SPVersion.
	UPDATE SoftwarePackage SET SoftwarePackage.SPVersion =  (SELECT MAX(SoftwarePackage.SPVersion) +1  FROM SoftwarePackage),
					SoftwarePackage.VersionDate = getdate()
	FROM SoftwarePackage 
	WHERE SoftwarePackage.SoftwarePackageID IN 	(
									SELECT SPA.SoftwarePackageID
									FROM SoftwarePackageApplication SPA INNER JOIN (INSERTED INS INNER JOIN DELETED DEL ON INS.ApplicationID = DEL.ApplicationID) ON SPA.ApplicationID = INS.ApplicationID								
									WHERE (DEL.Version <> INS.Version)
									AND SPA.IsDeleted =0
								)
	--As Well Update this records SPName, VersionDate and SPVersionBase
	UPDATE Application SET 	Application.SPName = Application.AppName + '_V' + cast(Application.Version As VarChar(10)) + '.exe',
					Application.VersionDate = GetDate(),
					Application.SPVersionBase = Application.SPVersion
	FROM Application  INNER JOIN INSERTED INS On Application.ApplicationID = INS.ApplicationID INNER JOIN DELETED DEL On INS.ApplicationID = DEL.ApplicationID
END

ELSE
	IF @CountVersionDate > 0
	BEGIN
		--Update the software package VersionDate
		UPDATE SoftwarePackage SET 	SoftwarePackage.VersionDate = getdate()
		FROM SoftwarePackage 
		WHERE SoftwarePackage.SoftwarePackageID IN 	(
										SELECT SPA.SoftwarePackageID
										FROM SoftwarePackageApplication SPA INNER JOIN (INSERTED INS INNER JOIN DELETED DEL ON INS.ApplicationID = DEL.ApplicationID) ON SPA.ApplicationID = INS.ApplicationID								
										WHERE (DEL.VersionDate <> INS.VersionDate)
										AND INS.VersionDate Is Not Null
										AND DEL.SPVersion Is Not Null	
										AND SPA.IsDeleted =0
									)
		--Use Server DATE Time for Version Date on Updates
		UPDATE Application SET 	Application.VersionDate = GetDate()
		FROM Application  INNER JOIN INSERTED INS On Application.ApplicationID = INS.ApplicationID INNER JOIN DELETED DEL On INS.ApplicationID = DEL.ApplicationID

		--Besure that the Version Date on Development matches the Version Date On Production
		UPDATE DevWebV2..Application SET  
					DevWebV2..Application.[VersionDate] = ProdWebV2..Application.[VersionDate],
					DevWebV2..Application.[DatelastUpdated] = ProdWebV2..Application.[DatelastUpdated] 
		FROM DevWebV2..Application 
			INNER JOIN INSERTED INS On DevWebV2..Application.[AppName] = INS.[AppName] 
			INNER JOIN DELETED DEL On INS.[ApplicationID] = DEL.[ApplicationID] 
			INNER JOIN ProdWebV2..Application On ProdWebV2..Application.[AppName] = DEL.[AppName]
	END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

