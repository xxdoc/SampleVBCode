SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[insClientCompanyCat]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[insClientCompanyCat]
GO

CREATE TRIGGER insClientCompanyCat
ON dbo.ClientCompanyCat 
AFTER INSERT
AS
--After Adding Company Cats Update the Software Package Table 
insert into SoftwarePackage (
	[ClientCompanyID],
	[CatID],
	[PackageName] ,
	[Description] ,
	[SPVersion] ,
	[VersionDate],
	[IsDeleted],
	[DateLastUpdated] ,
	[UpdateByUserID] 
	)
SELECT
	INS.[ClientCompanyID],
	INS.[CatID],
	'CAT_SP_' + Cat.Name As [PackageName],
	'Software Package for ' + C.Name + ' For CAT ' + CAT.Name As [Description],
	(CASE	WHEN (SELECT MAX(SPVersion) FROM SoftwarePackage) IS Null
		THEN 1
		ELSE (SELECT MAX(SPVersion) FROM SoftwarePackage)
		END
	) as [SPVersion] ,
	GetDate() as [VersionDate],
	0 As [IsDeleted],
	GetDate() As [DateLastUpdated] ,
	INS.[UpdateByUserID]
FROM 	INSERTED INS 	INNER JOIN CAT ON INS.CATID = CAT.CATID 
			INNER JOIN Company C ON INS.ClientCompanyID = C.CompanyID
WHERE C.IsClientOF Is Not Null

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

