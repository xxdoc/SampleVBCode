SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[insApplication]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[insApplication]
GO

CREATE TRIGGER insApplication
ON dbo.Application 
INSTEAD OF INSERT
AS

--As Well Update this records SPName and VersionDate And SPVersion and SPVersionBase
	INSERT INTO  Application (
					[AppNameBase],
					[AppName],
					[Description],
					[Version] ,
					[MajorVS],
					[MinorVS],
					[RevisionVS],
					[SPVersionBase],
					[SPVersion] ,
					[VersionDate],
					[ProjectName] ,
					[ClassName] ,
					[SectionLevel01],
					[SectionLevel02],
					[SectionLevel03],
					[SectionLevel04],
					[SectionLevel05],
					[InstallFileLocation],
					[SPName] ,
					[IsDeleted],
					[DateLastUpdated],
					[UpdateByUserID] 
				)
	SELECT 		
					INS.[AppNameBase],
					INS.[AppName],
					INS.[Description],
					INS.[Version],
					INS.[MajorVS],
					INS.[MinorVS],
					INS.[RevisionVS],
					(SELECT Max(SPVersion) FROM SoftwarePackage) As [SPVersionBase] ,
					(SELECT Max(SPVersion) FROM SoftwarePackage) As [SPVersion] ,
					GetDate() as [VersionDate],
					INS.[ProjectName] ,
					INS.[ClassName] ,
					INS.[SectionLevel01],
					INS.[SectionLevel02],
					INS.[SectionLevel03],
					INS.[SectionLevel04],
					INS.[SectionLevel05],
					INS.[InstallFileLocation],
					INS.[AppName] + '_V' + cast(INS.[Version] As VarChar(10)) + '.exe' As [SPName] ,
					INS.[IsDeleted],
					Getdate() as [DateLastUpdated] ,
					INS.[UpdateByUserID] 	
	FROM INSERTED INS

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

