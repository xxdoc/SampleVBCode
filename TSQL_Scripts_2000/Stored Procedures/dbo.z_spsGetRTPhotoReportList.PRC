SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spsGetRTPhotoReportList]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spsGetRTPhotoReportList]
GO

/******************************************************************
*
*    	Description:	Get RTPhotoReport List
*			
*	
*	Author: 	Brad Skidmore
*	Date: 		2/18/2005
*
******************************************************************/
CREATE PROCEDURE [dbo].[z_spsGetRTPhotoReportList]
	@UID			int,
	@AssignmentsID		Int=null
AS
BEGIN
	SELECT 
		[RTPhotoReportID],
		[AssignmentsID] ,
		[ID],
		[IDAssignments],
		[Name],
		[Description],
		[Number],
		[IsDeleted],
		[DownLoadMe],
		[UpLoadMe],
		[AdminComments],
		[DateLastUpdated],
		[UpdateByUserID]
	FROM 	RTPhotoReport
	WHERE 	AssignmentsID = @AssignmentsID
	ORDER BY [Number]
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[z_spsGetRTPhotoReportList]  TO [CFUSER]
GO

