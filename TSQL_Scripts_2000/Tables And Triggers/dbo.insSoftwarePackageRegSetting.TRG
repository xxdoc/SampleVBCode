SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[insSoftwarePackageRegSetting]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[insSoftwarePackageRegSetting]
GO

CREATE TRIGGER insSoftwarePackageRegSetting
ON dbo.SoftwarePackageRegSetting 
AFTER INSERT
AS

	UPDATE RegSetting set RegSetting.SPVersion = SP.SPVersion,
				RegSetting.SPVersionBase =	(CASE 	WHEN RegSetting.SPVersionBase IS Null
								  	THEN SP.SPVersion
									ELSE RegSetting.SPVersionBase
									END
								)		
	FROM  SoftwarePackage SP  	INNER JOIN INSERTED INS On SP.SoftwarePackageID = INS.SoftwarePackageID 
					INNER JOIN RegSetting  ON INS.RegSettingID = RegSetting.RegSettingID
	WHERE INS.RegsettingID = RegSetting.RegSettingID

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

