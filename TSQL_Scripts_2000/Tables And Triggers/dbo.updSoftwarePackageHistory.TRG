SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[updSoftwarePackageHistory]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[updSoftwarePackageHistory]
GO



CREATE TRIGGER updSoftwarePackageHistory
ON dbo.SoftwarePackage
AFTER UPDATE
AS
--Update History table
INSERT INTO  SoftwarePackageHistory
	SELECT del.* 
	FROM DELETED del

--See if there are any Version Changes
DECLARE @CountSPVersion int

SET @CountSPVersion = (
				SELECT Count(DEL.SPVersion) As CountOFSPVersion
				FROM DELETED DEL INNER JOIN INSERTED INS
				ON DEL.SoftwarePackageID = INS.SoftwarePackageID
				WHERE DEL.SPVersion <> INS.SPVersion
			)
IF @CountSPVersion > 0
BEGIN

	--RegSetting...
	UPDATE RegSetting SET 	RegSetting.SPVersion = INS.SPVersion,
					RegSetting.SPVersionBase =	(CASE 	WHEN RegSetting.SPVersionBase IS Null Or RegSetting.SPVersionBase > INS.SPVersion
								  	THEN INS.SPVersion
									ELSE RegSetting.SPVersionBase
									END
								)
	FROM INSERTED INS INNER JOIN  SoftwarePackageRegSetting SPR ON INS.SoftwarePackageID = SPR.SoftwarePackageID
	
	--Also update all dependant tables with new SPVersion
	--Documents...
	UPDATE Document SET 	Document.SPVersion = INS.SPVersion,
					Document.SPVersionBase =	(CASE 	WHEN Document.SPVersionBase IS Null Or Document.SPVersionBase >  INS.SPVersion
								  	THEN INS.SPVersion
									ELSE Document.SPVersionBase
									END
								)
	FROM INSERTED INS INNER JOIN  SoftwarePackageDocument SPD ON INS.SoftwarePackageID = SPD.SoftwarePackageID
	
	
	--Application...
	UPDATE Application SET 	Application.SPVersion = INS.SPVersion,
					Application.SPVersionBase =	(CASE 	WHEN Application.SPVersionBase IS Null OR Application.SPVersionBase > INS.SPVersion
								  	THEN INS.SPVersion
									ELSE Application.SPVersionBase
									END
								)
	FROM INSERTED INS INNER JOIN  SoftwarePackageApplication SPA ON INS.SoftwarePackageID = SPA.SoftwarePackageID
	
	
	--Update SPVersion in Software package
	UPDATE SoftwarePackage Set SoftwarePackage.SPVersion = INS.SPVersion
	FROM INSERTED INS
	
	
END




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

