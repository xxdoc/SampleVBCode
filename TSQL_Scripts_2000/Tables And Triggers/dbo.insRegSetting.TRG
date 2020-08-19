SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[insRegSetting]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[insRegSetting]
GO

CREATE TRIGGER insRegSetting
 ON dbo.RegSetting 
INSTEAD OF INSERT
AS

--As Well Update this records SPName and VersionDate And SPVersion and SPVersionBase
	INSERT INTO  RegSetting (
					[RegNameBase],
					[RegName],
					[Description],
					[Version] ,
					[SPVersionBase] ,
					[SPVersion] ,
					[VersionDate],
					[SectionLevel01],
					[SectionLevel02],
					[SectionLevel03],
					[SectionLevel04],
					[SectionLevel05],
					[SPName] ,
					[IsDeleted],
					[DateLastUpdated] ,
					[UpdateByUserID] 
				)
	SELECT 		
					INS.[RegNameBase],
					INS.[RegName],
					INS.[Description],
					INS.[Version],
					(SELECT Max(SPVersion) FROM SoftwarePackage) As [SPVersionBase] ,
					(SELECT Max(SPVersion) FROM SoftwarePackage) As [SPVersion] ,
					GetDate() as [VersionDate],
					INS.[SectionLevel01],
					INS.[SectionLevel02],
					INS.[SectionLevel03],
					INS.[SectionLevel04],
					INS.[SectionLevel05],
					INS.RegName + '_V' + cast(INS.Version As VarChar(10)) + '.exe' As [SPName] ,
					INS.[IsDeleted],
					Getdate() as [DateLastUpdated] ,
					INS.[UpdateByUserID] 	
	FROM INSERTED INS

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

