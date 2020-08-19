SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spDelAssignmentsInfo]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spDelAssignmentsInfo]
GO

/******************************************************************
*
*    	Description:	Delete Assignments Info
*			
*	
*	Author: 	Brad Skidmore
*	Date: 		1/20/2004
*	Date		2/17/2004 Added Date Span For Delete From History Table
******************************************************************/
CREATE PROCEDURE [dbo].[z_spDelAssignmentsInfo]
	@AssignmentTypeID	int=null,
	@ClientCompanyCatSpecID	int=null,
	@AdjusterSpecID		int=null,
	@AssignmentsID		Int=null,
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

	IF @AssignmentTypeID =-1 And  @ClientCompanyCatSpecID =-1 And @AdjusterSpecID = -1 And @AssignmentsID = -1
	BEGIN
		DELETE FROM AssignmentsHistory	
		WHERE 	DateLastUpdated >= @dtBeginHistoryDate
		AND 	DateLastUpdated <= @dtEndHistoryDate	
	END
	ELSE
	BEGIN
		IF @AssignmentsID = -1
		BEGIN
			DELETE FROM AssignmentsHistory	
			WHERE 	DateLastUpdated >= @dtBeginHistoryDate
			AND 	DateLastUpdated <= @dtEndHistoryDate	
		END
		ELSE IF @AssignmentTypeID >0 And  @ClientCompanyCatSpecID =-1 And @AdjusterSpecID = -1 And @AssignmentsID = -1
		BEGIN
			DELETE FROM AssignmentsHistory	
			WHERE 	AssignmentTypeID = @AssignmentTypeID
			AND 	DateLastUpdated >= @dtBeginHistoryDate
			AND 	DateLastUpdated <= @dtEndHistoryDate	
		END
		ELSE IF @AssignmentTypeID >0 And  @ClientCompanyCatSpecID >0 And @AdjusterSpecID = -1 And @AssignmentsID = -1
		BEGIN
			DELETE FROM AssignmentsHistory	
			WHERE 	AssignmentTypeID = @AssignmentTypeID
			AND 	ClientCompanyCatSpecID = @ClientCompanyCatSpecID
			AND 	DateLastUpdated >= @dtBeginHistoryDate
			AND 	DateLastUpdated <= @dtEndHistoryDate	
		END
		ELSE IF @AssignmentTypeID >0 And  @ClientCompanyCatSpecID >0 And @AdjusterSpecID >0 And @AssignmentsID = -1
		BEGIN
			DELETE FROM AssignmentsHistory	
			WHERE 	AssignmentTypeID = @AssignmentTypeID
			AND 	ClientCompanyCatSpecID = @ClientCompanyCatSpecID
			AND	AdjusterSpecID = @AdjusterSpecID
			AND 	DateLastUpdated >= @dtBeginHistoryDate
			AND 	DateLastUpdated <= @dtEndHistoryDate	
		END
		ELSE IF @AssignmentTypeID >0 And  @ClientCompanyCatSpecID >0 And @AdjusterSpecID >0 And @AssignmentsID >0
		BEGIN
			DELETE FROM AssignmentsHistory	
			WHERE 	AssignmentTypeID = @AssignmentTypeID
			AND 	ClientCompanyCatSpecID = @ClientCompanyCatSpecID
			AND	AdjusterSpecID = @AdjusterSpecID
			AND 	AssignmentsID = @AssignmentsID
			AND 	DateLastUpdated >= @dtBeginHistoryDate
			AND 	DateLastUpdated <= @dtEndHistoryDate
		END	
		ELSE IF @AssignmentTypeID Is Null And  @ClientCompanyCatSpecID Is Null And @AdjusterSpecID Is Null And @AssignmentsID >0
		BEGIN
			DELETE FROM AssignmentsHistory	
			WHERE 	AssignmentsID = @AssignmentsID
			AND 	DateLastUpdated >= @dtBeginHistoryDate
			AND 	DateLastUpdated <= @dtEndHistoryDate
		END	
	END
END
ELSE
BEGIN
	IF @AssignmentTypeID =-1 And  @ClientCompanyCatSpecID =-1 And @AdjusterSpecID = -1 And @AssignmentsID = -1
	BEGIN
		DELETE FROM Assignments
	END
	ELSE
	BEGIN
		IF @AssignmentTypeID >0 And  @ClientCompanyCatSpecID =-1 And @AdjusterSpecID = -1 And @AssignmentsID = -1
		BEGIN
			DELETE FROM Assignments
			WHERE 	AssignmentTypeID = @AssignmentTypeID
		END
		ELSE IF @AssignmentTypeID >0 And  @ClientCompanyCatSpecID >0 And @AdjusterSpecID = -1 And @AssignmentsID = -1
		BEGIN
			DELETE FROM Assignments
			WHERE 	AssignmentTypeID = @AssignmentTypeID
			AND 	ClientCompanyCatSpecID = @ClientCompanyCatSpecID
		END
		ELSE IF @AssignmentTypeID >0 And  @ClientCompanyCatSpecID >0 And @AdjusterSpecID >0 And @AssignmentsID = -1
		BEGIN
			DELETE FROM Assignments
			WHERE 	AssignmentTypeID = @AssignmentTypeID
			AND 	ClientCompanyCatSpecID = @ClientCompanyCatSpecID
			AND	AdjusterSpecID = @AdjusterSpecID
		END
		ELSE IF @AssignmentTypeID >0 And  @ClientCompanyCatSpecID >0 And @AdjusterSpecID >0 And @AssignmentsID >0
		BEGIN
			DELETE FROM Assignments
			WHERE 	AssignmentTypeID = @AssignmentTypeID
			AND 	ClientCompanyCatSpecID = @ClientCompanyCatSpecID
			AND	AdjusterSpecID = @AdjusterSpecID
			AND 	AssignmentsID = @AssignmentsID
		END
		ELSE IF @AssignmentTypeID Is Null And  @ClientCompanyCatSpecID Is Null And @AdjusterSpecID Is Null And @AssignmentsID >0
		BEGIN
			DELETE FROM Assignments
			WHERE AssignmentsID = @AssignmentsID
		END
	END
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[z_spDelAssignmentsInfo]  TO [CFUSER]
GO

