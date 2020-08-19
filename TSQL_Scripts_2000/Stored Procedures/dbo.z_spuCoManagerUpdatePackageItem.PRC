SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spuCoManagerUpdatePackageItem]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spuCoManagerUpdatePackageItem]
GO

/******************************************************************
*
*    	Description:	Co Manager Update Package Items 
*			Approval, Comments Only
*	
*	Author: 	Brad Skidmore
*	Date: 		7/25/2005
******************************************************************/
CREATE PROCEDURE [dbo].[z_spuCoManagerUpdatePackageItem]
	@UID			int,
	@PackageID 		int,
	@PackageItemID 		int,
	@sIsCoApprove 		varchar(50)='NO_CHANGE',
	@sCoApproveDate 	varchar(50)='NO_CHANGE',
	@sCoApproveDesc 	varchar(100)='NO_CHANGE',
	@sAdminComments 	varchar(1000)='NO_CHANGE',
	@sSendMe  		varchar(50)='NO_CHANGE',
	@sDownLoadMe 		varchar(50)='NO_CHANGE',
	@sUpLoadMe 		varchar(50)='NO_CHANGE'
AS	
	--Vars
	DECLARE @DateLastUpdated DateTime
	SET @DateLastUpdated = GetDate()
	IF IsDate(@sCoApproveDate) = 0 
	BEGIN
		SET @sCoApproveDate = Convert(Varchar(50), @DateLastUpdated)
	END
	
	--Mark items in packageItems as Sent when appropriate
	UPDATE PackageItem SET
		[IsCoApprove] 		=(CASE WHEN @sIsCoApprove =	'NO_CHANGE'	THEN [IsCoApprove] 	ELSE Convert(bit, @sIsCoApprove)	END),
		[CoApproveDate] 	=(CASE WHEN @sCoApproveDate =	'NO_CHANGE'	THEN [CoApproveDate] 	ELSE Convert(Datetime, @sCoApproveDate)	END),
		[CoApproveDesc] 	=(CASE WHEN @sCoApproveDesc =	'NO_CHANGE'	THEN [CoApproveDesc] 	ELSE @sCoApproveDesc	 		END),
		[SendMe] 		=(CASE WHEN @sSendMe =		'NO_CHANGE'	THEN [SendMe] 		ELSE Convert(bit, @sSendMe) 		END),
		[DownLoadMe] 		=(CASE WHEN @sDownLoadMe =	'NO_CHANGE'	THEN [DownLoadMe] 	ELSE Convert(bit, @sDownLoadMe) 	END),
		[UpLoadMe] 		=(CASE WHEN @sUpLoadMe =	'NO_CHANGE'	THEN [UpLoadMe] 	ELSE Convert(bit, @sUpLoadMe) 		END),
		[AdminComments] 	=(CASE WHEN @sAdminComments =	'NO_CHANGE'	THEN [AdminComments] 	ELSE @sAdminComments 			END),
		[DateLastUpdated] 	= @DateLastUpdated,
		[UpdateByUserID] 	= @UID
	WHERE	[PackageID]		= @PackageID
	AND	[PackageItemID]		= @PackageItemID

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[z_spuCoManagerUpdatePackageItem]  TO [CFUSER]
GO

