SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spuSetPKGIisDeletedFlag]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spuSetPKGIisDeletedFlag]
GO

/******************************************************************
*
*    	Description:	Set PackageItem IsDeleted Flag
*			
*	
*	Author: 	Brad Skidmore
*	Date: 		10/31/2005
******************************************************************/

CREATE PROCEDURE [dbo].[z_spuSetPKGIisDeletedFlag]
	@UID			int,
	@PackageItemID		int,
	@IsDeleted		bit,
	@SortOrder		int=1
AS
	UPDATE PackageItem SET 
				[Downloadme] = 1,
				[IsDeleted] = @IsDeleted,
				[SortOrder] = @SortOrder,
				[AdminComments] = (CASE WHEN @IsDeleted = 1 THEN 'Item was deleted out of the package by a Manager or Admin.' ELSE 'Item was UN deleted by a Manager or Admin.' END),
				DateLastUpdated = GetDate(),
				UpdateByUserID = @UID
	WHERE	PackageItemID = @PackageItemID

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[z_spuSetPKGIisDeletedFlag]  TO [CFUSER]
GO

