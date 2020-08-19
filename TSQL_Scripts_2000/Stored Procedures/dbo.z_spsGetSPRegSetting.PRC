SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spsGetSPRegSetting]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spsGetSPRegSetting]
GO

/******************************************************************
*
*    	Description:	Get Software Package Regsetting
*			Will Return the SPVERSION Associated records 
			From the Software RegSetting Table
			For the apropriate Client Company and Cat
*	
*	Author: 	Brad Skidmore
*	Date: 		5/13/2004
*
******************************************************************/


CREATE PROCEDURE [dbo].[z_spsGetSPRegSetting]
	@ClientCompanyID	int,
	@CATID			int,
	@SPVERSION		int,
	@bGetFromHistory	bit=0,
	@SectionLevel01		varchar(50)=null,
	@SectionLevel02		varchar(50)=null,
	@SectionLevel03		varchar(50)=null,
	@SectionLevel04		varchar(50)=null,
	@SectionLevel05		varchar(50)=null
AS
IF @bGetFromHistory = 0
BEGIN
	SELECT R.* 
	FROM	SoftwarePackage SP
			INNER JOIN 	SoftwarePackageRegSetting SPR 
					ON SPR.SoftWarePackageID = SP.SoftWarePackageID
			INNER JOIN	RegSetting R
					ON R.RegSettingID = SPR.RegSettingID
	WHERE	R.IsDeleted = 0
	AND	@SPVERSION >= R.SPVersionBase
	AND	@SPVERSION <= R.SPVersion
	AND	SPR.IsDeleted = 0 
	AND 	SP.CATID = @CATID
	AND 	SP.CLientCompanyID = @ClientCompanyID
	AND 	SectionLevel01 Like (CASE WHEN @SectionLevel01 Is Null THEN '%' ELSE  @SectionLevel01 END)
	AND 	SectionLevel02 Like (CASE WHEN @SectionLevel02 Is Null THEN '%' ELSE  @SectionLevel02 END) 
	AND 	SectionLevel03 Like (CASE WHEN @SectionLevel03 Is Null THEN '%' ELSE  @SectionLevel03 END)
	AND 	SectionLevel04 Like (CASE WHEN @SectionLevel04 Is Null THEN '%' ELSE  @SectionLevel04 END) 
	AND 	SectionLevel05 Like (CASE WHEN @SectionLevel05 Is Null THEN '%' ELSE  @SectionLevel05 END)
	Order by R.[Description]
END
ELSE
BEGIN
	SELECT R.*  
	FROM	SoftwarePackage SP
			INNER JOIN 	SoftwarePackageRegSetting SPR 
					ON SPR.SoftWarePackageID = SP.SoftWarePackageID
			INNER JOIN	RegSettingHistory R
					ON R.RegSettingID = SPR.RegSettingID
	WHERE	R.IsDeleted = 0
	AND	@SPVERSION >= R.SPVersionBase
	AND	@SPVERSION <= R.SPVersion
	AND	SPR.IsDeleted = 0 
	AND 	SP.CATID = @CATID
	AND 	SP.CLientCompanyID = @ClientCompanyID
	AND 	SectionLevel01 Like (CASE WHEN @SectionLevel01 Is Null THEN '%' ELSE  @SectionLevel01 END)
	AND 	SectionLevel02 Like (CASE WHEN @SectionLevel02 Is Null THEN '%' ELSE  @SectionLevel02 END) 
	AND 	SectionLevel03 Like (CASE WHEN @SectionLevel03 Is Null THEN '%' ELSE  @SectionLevel03 END)
	AND 	SectionLevel04 Like (CASE WHEN @SectionLevel04 Is Null THEN '%' ELSE  @SectionLevel04 END) 
	AND 	SectionLevel05 Like (CASE WHEN @SectionLevel05 Is Null THEN '%' ELSE  @SectionLevel05 END)
	Order by R.[Description]
END	

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[z_spsGetSPRegSetting]  TO [CFUSER]
GO

