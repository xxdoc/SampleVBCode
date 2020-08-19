SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spsGetSPApplication]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spsGetSPApplication]
GO

/******************************************************************
*
*    	Description:	Get Software Package Application
*			Will Return the SPVERSION Associated records 
			From the Software Application Table
			For the apropriate Client Company and Cat
*	
*	Author: 	Brad Skidmore
*	Date: 		5/13/2004
*
******************************************************************/


CREATE PROCEDURE [dbo].[z_spsGetSPApplication]
	@ClientCompanyID	int,
	@CATID			int,
	@SPVERSION		int,
	@bHasProjectName	bit=0,
	@bGetFromHistory	bit=0,
	@SectionLevel01		varchar(50)=null,
	@SectionLevel02		varchar(50)=null,
	@SectionLevel03		varchar(50)=null,
	@SectionLevel04		varchar(50)=null,
	@SectionLevel05		varchar(50)=null
AS
IF @bGetFromHistory = 0
BEGIN
	IF @bHasProjectName = 0
	BEGIN
		SELECT A.* 
		FROM	SoftwarePackage SP
				INNER JOIN 	SoftwarePackageApplication SPA 
						ON SPA.SoftWarePackageID = SP.SoftWarePackageID
				INNER JOIN	Application A
						ON A.ApplicationID = SPA.ApplicationID
		WHERE	A.IsDeleted = 0
		AND	@SPVERSION >= A.SPVersionBase
		AND	@SPVERSION <= A.SPVersion
		AND	SPA.IsDeleted = 0 
		AND 	SP.CATID = @CATID
		AND 	SP.CLientCompanyID = @ClientCompanyID
		AND 	SectionLevel01 Like (CASE WHEN @SectionLevel01 Is Null THEN '%' ELSE  @SectionLevel01 END)
		AND 	SectionLevel02 Like (CASE WHEN @SectionLevel02 Is Null THEN '%' ELSE  @SectionLevel02 END) 
		AND 	SectionLevel03 Like (CASE WHEN @SectionLevel03 Is Null THEN '%' ELSE  @SectionLevel03 END)
		AND 	SectionLevel04 Like (CASE WHEN @SectionLevel04 Is Null THEN '%' ELSE  @SectionLevel04 END) 
		AND 	SectionLevel05 Like (CASE WHEN @SectionLevel05 Is Null THEN '%' ELSE  @SectionLevel05 END) 
		Order by A.[Description], A.[ProjectName], A.[ClassName] 
	END
	ELSE
	BEGIN
		SELECT A.* 
		FROM	SoftwarePackage SP
				INNER JOIN 	SoftwarePackageApplication SPA 
						ON SPA.SoftWarePackageID = SP.SoftWarePackageID
				INNER JOIN	Application A
						ON A.ApplicationID = SPA.ApplicationID
		WHERE	A.IsDeleted = 0
		AND	@SPVERSION >= A.SPVersionBase
		AND	@SPVERSION <= A.SPVersion
		AND	SPA.IsDeleted = 0 
		AND 	SP.CATID = @CATID
		AND 	SP.CLientCompanyID = @ClientCompanyID
		AND 	SectionLevel01 Like (CASE WHEN @SectionLevel01 Is Null THEN '%' ELSE  @SectionLevel01 END)
		AND 	SectionLevel02 Like (CASE WHEN @SectionLevel02 Is Null THEN '%' ELSE  @SectionLevel02 END) 
		AND 	SectionLevel03 Like (CASE WHEN @SectionLevel03 Is Null THEN '%' ELSE  @SectionLevel03 END)
		AND 	SectionLevel04 Like (CASE WHEN @SectionLevel04 Is Null THEN '%' ELSE  @SectionLevel04 END) 
		AND 	SectionLevel05 Like (CASE WHEN @SectionLevel05 Is Null THEN '%' ELSE  @SectionLevel05 END) 
		AND	A.ProjectName Is Not Null
		AND	A.ProjectName <> ''
		Order by A.[Description], A.[ProjectName], A.[ClassName]
	END
END
ELSE
BEGIN
	IF @bHasProjectName = 0
	BEGIN
		SELECT A.*  
		FROM	SoftwarePackage SP
				INNER JOIN 	SoftwarePackageApplication SPA 
						ON SPA.SoftWarePackageID = SP.SoftWarePackageID
				INNER JOIN	ApplicationHistory A
						ON A.ApplicationID = SPA.ApplicationID
		WHERE	A.IsDeleted = 0
		AND	@SPVERSION >= A.SPVersionBase
		AND	@SPVERSION <= A.SPVersion
		AND	SPA.IsDeleted = 0 
		AND 	SP.CATID = @CATID
		AND 	SP.CLientCompanyID = @ClientCompanyID
		AND 	SectionLevel01 Like (CASE WHEN @SectionLevel01 Is Null THEN '%' ELSE  @SectionLevel01 END)
		AND 	SectionLevel02 Like (CASE WHEN @SectionLevel02 Is Null THEN '%' ELSE  @SectionLevel02 END) 
		AND 	SectionLevel03 Like (CASE WHEN @SectionLevel03 Is Null THEN '%' ELSE  @SectionLevel03 END)
		AND 	SectionLevel04 Like (CASE WHEN @SectionLevel04 Is Null THEN '%' ELSE  @SectionLevel04 END) 
		AND 	SectionLevel05 Like (CASE WHEN @SectionLevel05 Is Null THEN '%' ELSE  @SectionLevel05 END) 
		Order by A.[Description], A.[ProjectName], A.[ClassName]
	END
	ELSE
	BEGIN
		SELECT A.*  
		FROM	SoftwarePackage SP
				INNER JOIN 	SoftwarePackageApplication SPA 
						ON SPA.SoftWarePackageID = SP.SoftWarePackageID
				INNER JOIN	ApplicationHistory A
						ON A.ApplicationID = SPA.ApplicationID
		WHERE	A.IsDeleted = 0
		AND	@SPVERSION >= A.SPVersionBase
		AND	@SPVERSION <= A.SPVersion
		AND	SPA.IsDeleted = 0 
		AND 	SP.CATID = @CATID
		AND 	SP.CLientCompanyID = @ClientCompanyID
		AND 	SectionLevel01 Like (CASE WHEN @SectionLevel01 Is Null THEN '%' ELSE  @SectionLevel01 END)
		AND 	SectionLevel02 Like (CASE WHEN @SectionLevel02 Is Null THEN '%' ELSE  @SectionLevel02 END) 
		AND 	SectionLevel03 Like (CASE WHEN @SectionLevel03 Is Null THEN '%' ELSE  @SectionLevel03 END)
		AND 	SectionLevel04 Like (CASE WHEN @SectionLevel04 Is Null THEN '%' ELSE  @SectionLevel04 END) 
		AND 	SectionLevel05 Like (CASE WHEN @SectionLevel05 Is Null THEN '%' ELSE  @SectionLevel05 END) 
		AND	A.ProjectName Is Not Null
		AND	A.ProjectName <> ''
		Order by A.[Description], A.[ProjectName], A.[ClassName]
	END
END	

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[z_spsGetSPApplication]  TO [CFUSER]
GO

