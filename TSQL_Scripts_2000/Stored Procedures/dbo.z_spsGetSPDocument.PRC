SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spsGetSPDocument]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spsGetSPDocument]
GO

/******************************************************************
*
*    	Description:	Get Software Package Document
*			Will Return the SPVERSION Associated records 
			From the Software Document Table
			For the apropriate Client Company and Cat
*	
*	Author: 	Brad Skidmore
*	Date: 		5/13/2004
*
******************************************************************/


CREATE PROCEDURE [dbo].[z_spsGetSPDocument]
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
	SELECT D.* 
	FROM	SoftwarePackage SP
			INNER JOIN 	SoftwarePackageDocument SPD 
					ON SPD.SoftWarePackageID = SP.SoftWarePackageID
			INNER JOIN	Document D
					ON D.DocumentID = SPD.DocumentID
	WHERE	D.IsDeleted = 0
	AND	@SPVERSION >= D.SPVersionBase
	AND	@SPVERSION <= D.SPVersion
	AND	SPD.IsDeleted = 0 
	AND 	SP.CATID = @CATID
	AND 	SP.CLientCompanyID = @ClientCompanyID
	AND 	SectionLevel01 Like (CASE WHEN @SectionLevel01 Is Null THEN '%' ELSE  @SectionLevel01 END)
	AND 	SectionLevel02 Like (CASE WHEN @SectionLevel02 Is Null THEN '%' ELSE  @SectionLevel02 END) 
	AND 	SectionLevel03 Like (CASE WHEN @SectionLevel03 Is Null THEN '%' ELSE  @SectionLevel03 END)
	AND 	SectionLevel04 Like (CASE WHEN @SectionLevel04 Is Null THEN '%' ELSE  @SectionLevel04 END) 
	AND 	SectionLevel05 Like (CASE WHEN @SectionLevel05 Is Null THEN '%' ELSE  @SectionLevel05 END)
	Order by D.[Description]
END
ELSE
BEGIN
	SELECT D.*  
	FROM	SoftwarePackage SP
			INNER JOIN 	SoftwarePackageDocument SPD 
					ON SPD.SoftWarePackageID = SP.SoftWarePackageID
			INNER JOIN	DocumentHistory D
					ON D.DocumentID = SPD.DocumentID
	WHERE	D.IsDeleted = 0
	AND	@SPVERSION >= D.SPVersionBase
	AND	@SPVERSION <= D.SPVersion
	AND	SPD.IsDeleted = 0 
	AND 	SP.CATID = @CATID
	AND 	SP.CLientCompanyID = @ClientCompanyID
	AND 	SectionLevel01 Like (CASE WHEN @SectionLevel01 Is Null THEN '%' ELSE  @SectionLevel01 END)
	AND 	SectionLevel02 Like (CASE WHEN @SectionLevel02 Is Null THEN '%' ELSE  @SectionLevel02 END) 
	AND 	SectionLevel03 Like (CASE WHEN @SectionLevel03 Is Null THEN '%' ELSE  @SectionLevel03 END)
	AND 	SectionLevel04 Like (CASE WHEN @SectionLevel04 Is Null THEN '%' ELSE  @SectionLevel04 END) 
	AND 	SectionLevel05 Like (CASE WHEN @SectionLevel05 Is Null THEN '%' ELSE  @SectionLevel05 END)
	Order by D.[Description]
END	

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[z_spsGetSPDocument]  TO [CFUSER]
GO

