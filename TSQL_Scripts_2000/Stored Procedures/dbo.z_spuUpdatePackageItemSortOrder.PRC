SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spuUpdatePackageItemSortOrder]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spuUpdatePackageItemSortOrder]
GO

/******************************************************************
*
*    	Description:	Update package Item Sort Order
*			
*	
*	Author: 	Brad Skidmore
*	Date: 		8.29.2005
******************************************************************/
CREATE PROCEDURE [dbo].[z_spuUpdatePackageItemSortOrder]
	@UID			int,
	@PackageItemID		int,
	@SortOrder		int
AS	
	UPDATE PackageItem SET
		[SortOrder]		= @SortOrder,
		[DownLoadMe] 		= 1,
		[DateLastUpdated] 	= GetDate(),
		[UpdateByUserID] 	= @UID
	WHERE	[PackageItemID] 	= @PackageItemID

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[z_spuUpdatePackageItemSortOrder]  TO [CFUSER]
GO

