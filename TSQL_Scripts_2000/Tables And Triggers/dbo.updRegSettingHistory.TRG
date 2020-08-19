SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[updRegSettingHistory]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[updRegSettingHistory]
GO

CREATE TRIGGER updRegSettingHistory
ON dbo.RegSetting
AFTER UPDATE
AS
--See if there are any Version Changes
DECLARE @CountVersion int
DECLARE @CountVersionDate int

SET @CountVersion = 	(
				SELECT Count(DEL.Version) As CountOFVersion
				FROM DELETED DEL INNER JOIN INSERTED INS
				ON DEL.RegSettingID = INS.RegSettingID
				WHERE (DEL.Version <> INS.Version)
				AND DEL.SPVersion Is Not Null	
			)
SET @CountVersionDate = 	(
				SELECT Count(DEL.VersionDate) As CountOFVersionDate
				FROM DELETED DEL INNER JOIN INSERTED INS
				ON DEL.RegSettingID = INS.RegSettingID
				WHERE (DEL.VersionDate <> INS.VersionDate)
				AND INS.VersionDate Is Not Null
				AND DEL.SPVersion Is Not Null	
				)

IF @CountVersion > 0
BEGIN
	INSERT INTO RegSettingHistory
		SELECT DEL.* 
		FROM DELETED DEL INNER JOIN INSERTED INS
		ON DEL.RegSettingID = INS.RegSettingID
		WHERE (DEL.Version <> INS.Version )
		AND DEL.SPVersion Is Not Null	

	--Then update the Softwarepackage Table which will in turn Update All dependant Tables
	--with new SPVersion.
	UPDATE SoftwarePackage SET SoftwarePackage.SPVersion =  (SELECT MAX(SoftwarePackage.SPVersion) +1  FROM SoftwarePackage),
					SoftwarePackage.VersionDate = getdate()
	FROM SoftwarePackage 
	WHERE SoftwarePackage.SoftwarePackageID IN 	(
									SELECT SPR.SoftwarePackageID
									FROM SoftwarePackageRegSetting SPR INNER JOIN (INSERTED INS INNER JOIN DELETED DEL ON INS.RegSettingID = DEL.RegSettingID) ON SPR.RegSettingID = INS.RegSettingID								
									WHERE DEL.Version <> INS.Version
									AND SPR.IsDeleted =0
								)
	--As Well Update this records SPName, VersionDate and SPVersionBase
	UPDATE RegSetting SET 	RegSetting.SPName = RegSetting.RegName + '_V' + cast(RegSetting.Version As VarChar(10)) + '.exe',
					RegSetting.VersionDate = GetDate(),
					RegSetting.SPVersionBase = RegSetting.SPVersion
	FROM RegSetting  INNER JOIN INSERTED INS On RegSetting.RegSettingID = INS.RegSettingID INNER JOIN DELETED DEL On INS.RegSettingID = DEL.RegSettingID
	
END

ELSE
	IF @CountVersionDate > 0
	BEGIN
		--Update the software package VersionDate
		UPDATE SoftwarePackage SET 	SoftwarePackage.VersionDate = getdate()
		FROM SoftwarePackage 
		WHERE SoftwarePackage.SoftwarePackageID IN 	(
										SELECT SPR.SoftwarePackageID
										FROM SoftwarePackageRegSetting SPR INNER JOIN (INSERTED INS INNER JOIN DELETED DEL ON INS.RegSettingID = DEL.RegSettingID) ON SPR.RegSettingID = INS.RegSettingID								
										WHERE (DEL.VersionDate <> INS.VersionDate)
										AND INS.VersionDate Is Not Null
										AND DEL.SPVersion Is Not Null	
										AND SPR.IsDeleted =0
									)
		--Use Server DATE Time for Version Date on Updates
		UPDATE RegSetting SET RegSetting.VersionDate = GetDate()
		FROM RegSetting  INNER JOIN INSERTED INS On RegSetting.RegSettingID = INS.RegSettingID INNER JOIN DELETED DEL On INS.RegSettingID = DEL.RegSettingID

		--Besure that the Version Date on Development matches the Version Date On Production
		UPDATE DevWebV2..RegSetting SET  
					DevWebV2..RegSetting.[VersionDate] = ProdWebV2..RegSetting.[VersionDate],
					DevWebV2..RegSetting.[DatelastUpdated] = ProdWebV2..RegSetting.[DatelastUpdated] 
		FROM DevWebV2..RegSetting 
			INNER JOIN INSERTED INS On DevWebV2..RegSetting.[RegName] = INS.[RegName] 
			INNER JOIN DELETED DEL On INS.[RegSettingID] = DEL.[RegSettingID] 
			INNER JOIN ProdWebV2..RegSetting On ProdWebV2..RegSetting.[RegName] = DEL.[RegName]
	END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

