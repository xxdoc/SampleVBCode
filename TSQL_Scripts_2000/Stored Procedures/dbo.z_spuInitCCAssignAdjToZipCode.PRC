SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spuInitCCAssignAdjToZipCode]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spuInitCCAssignAdjToZipCode]
GO

/******************************************************************
*
*    	Description:	Assign Adjuster to ZipCode
*			
*	Author: 	Brad Skidmore
*	Date: 		2/5/2004
*	
******************************************************************/

CREATE PROCEDURE [dbo].[z_spuInitCCAssignAdjToZipCode]
--Requird Params
@UID				int,
@ClientCompanyCatSpecID		int,
@iAssignByZIP			int,
--OutPut Params
@iZIP_ClientCoAdjusterSpecID	int=null OUTPUT

AS
DECLARE @iMyMaxAssignmentsAllowed int
DECLARE @iMyNumCurrentAssignments int
DECLARE @iMyUsersID int
DECLARE @iMyClientCompanyID int

	--This will Set the OutPut Param
	SELECT 	Top 1 
		@iZIP_ClientCoAdjusterSpecID = CCAS.ClientCoAdjusterSpecID
	FROM 	ClientCoAdjusterSpec CCAS INNER JOIN Adjuster ADJ ON CCAS.ClientCompanyID = ADJ.CompanyID AND CCAS.UsersID = ADJ.UsersID
	WHERE	CCAS.ClientCompanyCatSpecID = @ClientCompanyCatSpecID
	AND 	CCAS.ZipCode = @iAssignByZIP
	AND	CCAS.InactiveDate Is Null
	--ALWAYS INCLUDE Adjusters that Have a MaxOpenAssgn = 0 
	--ONLY INCLUDE Adjusters WHERE their Total OPEN Assignments is < the MaxOpenAssgn
	AND	(	(	CASE 	WHEN ADJ.MaxOpenAssgn > 0
					THEN
						(
							SELECT 	COUNT(AssignmentsID) As TotalOpenAssignments 
							FROM	Assignments 
							WHERE 	ClientCompanyCatSpecID IN	(	SELECT 	ClientCompanyCatSpecID
													FROM	ClientCompanyCatSpec
													WHERE 	ClientCompanyID = CCAS.[ClientCompanyID]
													AND	CATID	= 	(	SELECT	CATID 
																	FROM	ClientCompanyCatSpec
																	WHERE ClientCompanyCatSpecID = CCAS.[ClientCompanyCatSpecID]
																)
												) 
							AND 	AdjusterSpecID IN	(	SELECT 	ClientCoAdjusterSpecID 
												FROM 	ClientCoAdjusterSpec 
												WHERE 	ClientCompanyID = CCAS.[ClientCompanyID]
												AND	UsersID = CCAS.[UsersID]
											)
							AND 	StatusID IN 		(	SELECT 	StatusID
												FROM 	Status
												WHERE 	Status.Status NOT IN ('CLOSED', 'DELETED', 'REASSIGNED')
											)
							AND	IsDeleted = 0
						) 
					ELSE -1
					END
			) < ADJ.MaxOpenAssgn
		)
	Order By  	(
				SELECT 	COUNT(AssignmentsID) As TotalOpenAssignments 
				FROM	Assignments 
				WHERE 	ClientCompanyCatSpecID IN	(	SELECT 	ClientCompanyCatSpecID
										FROM	ClientCompanyCatSpec
										WHERE 	ClientCompanyID = CCAS.[ClientCompanyID]
										AND	CATID	= 	(	SELECT	CATID 
														FROM	ClientCompanyCatSpec
														WHERE ClientCompanyCatSpecID = CCAS.[ClientCompanyCatSpecID]
													)
									) 
				AND 	AdjusterSpecID IN	(	SELECT 	ClientCoAdjusterSpecID 
									FROM 	ClientCoAdjusterSpec 
									WHERE 	ClientCompanyID = CCAS.[ClientCompanyID]
									AND	UsersID = CCAS.[UsersID]
								)
				AND 	StatusID IN 		(	SELECT 	StatusID
									FROM 	Status
									WHERE 	Status.Status NOT IN ('CLOSED', 'DELETED', 'REASSIGNED')
								)
				AND	IsDeleted = 0
			) Asc, 
			ActiveDate Asc
	
	IF @iZIP_ClientCoAdjusterSpecID Is Not Null
	BEGIN
		--First Populate these common Vars
		SET @iMyUsersID = 	(
					SELECT 	UsersID
					FROM	ClientCoAdjusterSpec
					WHERE	ClientCoAdjusterSpecID = @iZIP_ClientCoAdjusterSpecID
					)
		SET @iMyClientCompanyID =	(
						SELECT 	ClientCompanyID
						FROM	ClientCoAdjusterSpec
						WHERE	ClientCoAdjusterSpecID = @iZIP_ClientCoAdjusterSpecID
						)
		--Get the Maximum number of Assignments Allowed for the Selected adjuster
		--Check the Assignments Table for the Selected Adjuster
		SET @iMyMaxAssignmentsAllowed =	(
						SELECT 	[MaxOpenAssgn]
						FROM 	Adjuster
						WHERE	[CompanyID] = @iMyClientCompanyID	
						AND 	[UsersID] = @iMyUsersID
						)
		--If the max allowed is Greater than 0 then Check this max against 
		--the number of Assignments not in closed status
		
		IF @iMyMaxAssignmentsAllowed > 0 
		BEGIN
			SET @iMyNumCurrentAssignments = (
								SELECT 	COUNT(AssignmentsID) As TotalOpenAssignments 
								FROM	Assignments 
								WHERE 	ClientCompanyCatSpecID IN	(	SELECT 	ClientCompanyCatSpecID
														FROM	ClientCompanyCatSpec
														WHERE 	ClientCompanyID = @iMyClientCompanyID
														AND	CATID	=	(	
																		SELECT	CATID 
																		FROM	ClientCompanyCatSpec
																		WHERE ClientCompanyCatSpecID = @ClientCompanyCatSpecID
																	)
													) 
								AND 	AdjusterSpecID IN	(	SELECT 	ClientCoAdjusterSpecID 
													FROM 	ClientCoAdjusterSpec 
													WHERE 	ClientCompanyID = @iMyClientCompanyID
													AND	UsersID = @iMyUsersID
												)
								AND 	StatusID IN 		(	SELECT 	StatusID
													FROM 	Status
													WHERE 	Status.Status NOT IN ('CLOSED', 'DELETED', 'REASSIGNED')
												)
								AND	IsDeleted = 0
							)
			--If the total number of of current assignments is >= to the MAx Allowed 
			-- need to set @iZIP_ClientCoAdjusterSpecID to null 
			IF @iMyNumCurrentAssignments >= @iMyMaxAssignmentsAllowed
			BEGIN
				SET @iZIP_ClientCoAdjusterSpecID = Null
			END
		END
	END
	
	--IF @iZIP_ClientCoAdjusterSpecID Is Null Then That means Nobody IS Available to
	-- Take the Assignment.  Need to Reactive the UNASSIGNED Adjuster to Take this Assignment
	IF  @iZIP_ClientCoAdjusterSpecID Is Null
	BEGIN
		--Set this to the Unassigned USER for this ZIPCODE 
		SELECT 	@iZIP_ClientCoAdjusterSpecID = CCAS.ClientCoAdjusterSpecID
		FROM 	ClientCoAdjusterSpec CCAS 
		WHERE	CCAS.ClientCompanyCatSpecID = @ClientCompanyCatSpecID
		AND 	CCAS.ZipCode = @iAssignByZIP
		AND 	CCAS.UsersID = 	(	SELECT	UsersID
						FROM 	Users
						WHERE	UserName = 'UUNASSIGNED'
					)
		--Reset the Inactive Date to Null
		UPDATE 	ClientCoAdjusterSpec SET 	InactiveDate = Null,
							DateLastUpdated = GetDate(),
							UpdateByUserID = @UID	
		WHERE	ClientCoAdjusterSpecID = @iZIP_ClientCoAdjusterSpecID
		
	END
	--This will Also Return  The same as a Record Set
	SELECT @iZIP_ClientCoAdjusterSpecID As AssignToAdjusterSpecID

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[z_spuInitCCAssignAdjToZipCode]  TO [CFUSER]
GO

