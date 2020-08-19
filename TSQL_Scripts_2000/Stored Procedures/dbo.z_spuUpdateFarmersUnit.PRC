SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spuUpdateFarmersUnit]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spuUpdateFarmersUnit]
GO

/******************************************************************
*
*    	Description:	Update Policy Limits to accommodate IA_CRN_CLOSE transaction.
*			
*	
*	Author: 	Brad Skidmore
*	Date: 		5/06/2005
*	
******************************************************************/


CREATE PROCEDURE [dbo].[z_spuUpdateFarmersUnit]
	@UserName		varchar(20),
	@FarmersUnitRowID	varchar(20),
	@IsDeleted		bit,
	@AdminComments		varchar(255)

AS

DECLARE @PolicyLimitsID 		int
DECLARE @UID				int
DECLARE	@AssignmentsID			int
DECLARE @AssignmentOpenUnits		int
DECLARE @AssignmentClosedUnits		int
DECLARE @CurDate			DateTime
DECLARE	@CloseDate			DateTime
DECLARE @StatusID			int

--Get Midnight Date
SET @CurDate = GetDate()
SET @CloseDate = @CurDate
SET @CloseDate = dbo.CleanFromOrToDate(@CurDate, @CloseDate, 1)

SET @UID = 	(
		SELECT 	UsersID 
		FROM 	Users 
		WHERE 	UserName = @UserName
		)

--Becuase Farmers is ONLY GIVING US A UNIQUE ID when they send us a Close Transaction 
--and the unit could be assigned to More than one of our adjusters during the Assignment and Reassignment process...
--- Need to see what the 
IF @IsDeleted = 1
BEGIN
	--Set some Variables
	SELECT 	Top 1 	@PolicyLimitsID = [PolicyLimitsID],
			@AssignmentsID = [AssignmentsID]
	FROM 	PolicyLimits
	WHERE 	[FarmersUnitRowID] = @FarmersUnitRowID
	AND [IsDeleted] = 0  -- Need to set the latest one that is not yet deleted.
	Order By [DateLastUpdated] Desc
END
IF @IsDeleted = 0
BEGIN
	--Set some Variables
	SELECT 	Top 1 	@PolicyLimitsID = [PolicyLimitsID],
			@AssignmentsID = [AssignmentsID]
	FROM 	PolicyLimits
	WHERE 	[FarmersUnitRowID] = @FarmersUnitRowID
	AND [IsDeleted] = 1 -- Need to set the latest one that is Deleted.
	Order By [DateLastUpdated] Desc
END

IF @PolicyLimitsID Is Null
BEGIN
	RAISERROR ('No Records Affected', 16, 1)
	RETURN 1 
END
ELSE
BEGIN
	UPDATE PolicyLimits SET	
		[IsDeleted]		= @IsDeleted,
		[AdminComments] 	= @AdminComments,
		[DownLoadMe]		= 1,
		[DateLastUpdated]	= GetDate(),
		[UpdateByUserID]	= @UID
	WHERE 		[PolicyLimitsID]	= @PolicyLimitsID
	
	-- If Farmers is closing a unit... Need to see if All other units associated
	-- with the same assignmentsID are also closed.  If they Are all closed, then
	-- need to update the Assigments Table ... to Close the Entire Assignment 
	IF @IsDeleted = 1
	BEGIN
		--Get the Count of Open Units for the AssignmentID
		SELECT 	@AssignmentOpenUnits = Count([AssignmentsID])
		FROM	PolicyLimits
		WHERE	[AssignmentsID] = @AssignmentsID
		AND 	[IsDeleted] = 0
		
		IF @AssignmentOpenUnits = 0
		BEGIN
			--Get the Count of all the Closed Units
			SELECT @AssignmentClosedUnits = Count([AssignmentsID])
			FROM	PolicyLimits
			WHERE	[AssignmentsID] = @AssignmentsID
			AND 	[IsDeleted] = 1
			
			IF @AdminComments Is null
			BEGIN
				SET @AdminComments = ''
			END
			
			IF @AdminComments Like '%CRN Reassigned to%'
			BEGIN
				SELECT 	@StatusID = [StatusID]
				FROM	Status
				WHERE	[Status] = 'CLIENTReassigned'
			END
			ELSE
			BEGIN
				SET @AdminComments = 'All ' + convert(varchar(10), @AssignmentClosedUnits) + ' Unit(s) Closed by client.'
				SELECT 	@StatusID = [StatusID]
				FROM	Status
				WHERE	[Status] = 'CLOSED'
			END
			--If the count >0 then Update Assignments to Close this Assignment
			UPDATE	Assignments SET
				[StatusID]		= @StatusID,
				[CloseDate] 		= @CloseDate,
				[AdminComments] 	= @AdminComments,
				[DownLoadMe]		= 1,
				[DateLastUpdated]	= GetDate(),
				[UpdateByUserID]	= @UID
			WHERE 	[AssignmentsID]		= @AssignmentsID
		END
	END
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[z_spuUpdateFarmersUnit]  TO [CFUSER]
GO

