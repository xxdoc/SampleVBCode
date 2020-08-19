SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spuSetPKGISendMeFlag]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spuSetPKGISendMeFlag]
GO

/******************************************************************
*
*    	Description:	Set PackageItem SendMe Flag
*			
*	
*	Author: 	Brad Skidmore
*	Date: 		10/31/2005
******************************************************************/

CREATE PROCEDURE [dbo].[z_spuSetPKGISendMeFlag]
	@UID			int,
	@PackageItemID		int,
	@SendMe			bit
AS
	UPDATE PackageItem SET 	
				[Downloadme] = 1,
				[SendMe] = @SendMe,
				[AdminComments] = (CASE WHEN @SendMe = 1 THEN 'Item marked to Send by Admin or Manager.' ELSE 'Item UN marked to Send by Admin or Manager.' END),
				DateLastUpdated = GetDate(),
				UpdateByUserID = @UID
	WHERE	PackageItemID = @PackageItemID

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[z_spuSetPKGISendMeFlag]  TO [CFUSER]
GO

