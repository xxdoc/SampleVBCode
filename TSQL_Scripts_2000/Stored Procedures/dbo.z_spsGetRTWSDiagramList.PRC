SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spsGetRTWSDiagramList]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spsGetRTWSDiagramList]
GO

/******************************************************************
*
*    	Description:	Get RTWS Diagram List
*			
*	
*	Author: 	Brad Skidmore
*	Date: 		2/18/2005
*
******************************************************************/
CREATE PROCEDURE [dbo].[z_spsGetRTWSDiagramList]
	@UID			int,
	@AssignmentsID		Int=null
AS
BEGIN
	SELECT 
		[RTWSDiagramID],
		[AssignmentsID],
		[ID],
		[IDAssignments],
		[Name],
		[Description],
		[Number],
		[DiagramPhotoName],
		[DownloadDiagramPhoto],
		[UploadDiagramPhoto],
		[DiagramXML],
		[IsDeleted],
		[DownLoadMe],
		[UpLoadMe],
		[AdminComments],
		[DateLastUpdated],
		[UpdateByUserID] 
	FROM 	RTWSDiagram
	WHERE 	AssignmentsID = @AssignmentsID
	ORDER BY [Number]
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[z_spsGetRTWSDiagramList]  TO [CFUSER]
GO

