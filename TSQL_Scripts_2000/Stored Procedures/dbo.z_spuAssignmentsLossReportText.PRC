SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spuAssignmentsLossReportText]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spuAssignmentsLossReportText]
GO

/******************************************************************
*
*    	Description:	Update Blob Assignments Loss Report TEXT
*			Update the TEXT ONLY without affecting Triggers
*	
*	Author: 	Brad Skidmore
*	Date: 		4/22/2004
******************************************************************/
CREATE PROCEDURE [dbo].[z_spuAssignmentsLossReportText]
	@AssignmentsID	int,
	@LossReport	ntext
AS

DECLARE @ptrLossReport		binary(16) --Pointer for ntext field update
--V2ECcarFarmers.clsLossXML01
-- Do NOT allow changes to Farmers Loss report clsLossXML01
IF (SELECT Ass.[LRFormat] FROM Assignments Ass WHERE Ass.AssignmentsID = @AssignmentsID) <> 'V2ECcarFarmers.clsLossXML01'
BEGIN
	SELECT 	@ptrLossReport = TEXTPTR(LossReport) 
	FROM 	Assignments
	WHERE 	AssignmentsID	=@AssignmentsID
	WRITETEXT Assignments.LossReport @ptrLossReport @LossReport
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[z_spuAssignmentsLossReportText]  TO [CFUSER]
GO

