SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spDelSoftwarePackageInfo]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spDelSoftwarePackageInfo]
GO

/******************************************************************
*
*    	Description:	Delete Software Package Info
*			
*	
*	Author: 	Brad Skidmore
*	Date: 		11/11/2003
*	Date		2/17/2004 Added Date Span For Delete From History Table
******************************************************************/
CREATE PROCEDURE [dbo].[z_spDelSoftwarePackageInfo]
	@ClientCompanyID	int,
	@SoftwarePackageID	Int,
	@bRemoveHistoryOnly 	int=0,
	@dtBeginHistoryDate	datetime=null,
	@dtEndHistoryDate	datetime=null
AS
IF @bRemoveHistoryOnly = 1 
BEGIN
	--Format the Begin and End History Dates
	IF IsDate(@dtBeginHistoryDate) = 0 
	BEGIN
		SET @dtBeginHistoryDate = Convert(datetime,dbo.FormatDate(GetDate(),'mm/dd/yyy'))
	END
	ELSE
	BEGIN
		SET @dtBeginHistoryDate = Convert(datetime,dbo.FormatDate(@dtBeginHistoryDate,'mm/dd/yyy'))
	END
	
	IF IsDate(@dtEndHistoryDate) = 0 
	BEGIN
		SET @dtEndHistoryDate = GetDate()
	END
	ELSE
	BEGIN
		SET @dtEndHistoryDate = Convert(datetime,dbo.FormatDate(@dtEndHistoryDate,'mm/dd/yyy')) + ' 23:59:59.999'
	END

	IF @ClientCompanyID =-1 And  @SoftwarePackageID =-1
	BEGIN
		DELETE FROM SoftwarePackageHistory	
	END
	ELSE
	BEGIN
		IF @SoftwarePackageID =-1
		BEGIN
			DELETE FROM SoftwarePackageHistory	
			WHERE 	ClientCompanyID = @ClientCompanyID
			AND 	DateLastUpdated >= @dtBeginHistoryDate
			AND 	DateLastUpdated <= @dtEndHistoryDate	
		END
		ELSE
		BEGIN
			DELETE FROM SoftwarePackageHistory	
			WHERE 	ClientCompanyID = @ClientCompanyID
			AND	SoftwarePackageID = @SoftwarePackageID
			AND 	DateLastUpdated >= @dtBeginHistoryDate
			AND 	DateLastUpdated <= @dtEndHistoryDate	
		END
	END
END
ELSE
BEGIN
	--Don't Delete from Software just mark as deleted
-- 	DELETE FROM 	SoftwarePackage
-- 	WHERE 	ClientCompanyID = @ClientCompanyID
-- 	AND	SoftwarePackageID = @SoftwarePackageID
	--Also Return this message so that WEBCONTROL is Informed of this action.
	RAISERROR ('ERROR FLAG UPDATE IsDeleted INSTEAD OF DELETE.', 16, 1) 
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[z_spDelSoftwarePackageInfo]  TO [CFUSER]
GO

