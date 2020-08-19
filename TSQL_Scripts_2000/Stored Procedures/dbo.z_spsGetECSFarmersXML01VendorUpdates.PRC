SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spsGetECSFarmersXML01VendorUpdates]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spsGetECSFarmersXML01VendorUpdates]
GO

/******************************************************************
*
*    	Description:	Returns Vendor Updates Recordset(s) for
*			Activtities and Property updates.
*			Initially the first call to this SP should pass in null @AssignmentsID
*			to get the list of Active Units that have Updates Flaged
*	Author: 	Brad Skidmore
*	Date: 		5/19/2005
*
******************************************************************/

CREATE PROCEDURE [dbo].[z_spsGetECSFarmersXML01VendorUpdates]
	@AssignmentsID		int=null,
	@UnitId			varchar(20)=null,
	@ClaimId		varchar(20)=null,
	@SetSendToCRN		bit=null	
AS
--Check the Get Active Units Flag
-- If true this procedure will only return active units that have Updates flagged
--Active units = Rows under PolicyLimits Table where the 
--PolicyLimits.[AssignmentsID] = FarmersXML01_LocationUpdateQue.[AssignmentsID] and
--PolicyLimits.[IsDeleted] = 0  and 
--FarmersXML01_LocationUpdateQue.[SendToCRN] = 1 and
IF @AssignmentsID Is Null
BEGIN
	SELECT 	PL.[AssignmentsID],
		PL.[FarmersUnitRowID] As [UnitId]
	FROM 	PolicyLimits PL
	--8.12.2005 This is an impossible flag to set
	--Basically this is what Farmers Expects... If an adjuster makes an update after Farmers
	--Closed the unit...  They don't want any updates after 30 days from the date of close
	-- So If you check for the Deleted Flag then don't send anything to avoid getting an rejected result
	-- response back from Farmers .. but there was an update... so then you have the problem
	-- of having updates that were done to an item after Farmers Closed it and 30 days.  This can get hairy
	-- because there are a zillion instances where farmers will close something then unclose it reclose it
	-- because they are feeling naughty that day... Don't believe me , thats your personal problem and ignorance. 
	-- so the timing is everything on this.  There is no way in hell we are going to sit around
	-- and wait for Farmers to unclose something, because there is no gaurantee they will,
	-- so that we can then send an update.  So I opt to get a Rejected update message from Farmers
	-- for the outgoing update after farmers closes something 30 days ago. 
	-- if we happen to send them an update to something they should have unclosed so we can update them...
	-- Farmers can figure out that they should have unclosed the item before they reject it because it was after 
	-- the 30 days they closed the sob.
	--WHERE	PL.[IsDeleted] = 0 
	WHERE	(
			PL.[AssignmentsID] IN	(
						SELECT	[AssignmentsID]
						FROM	FarmersXML01_LocationUpdateQue
						WHERE	[SendToCRN] = 1
						)
		OR	PL.[AssignmentsID] IN	(
						SELECT	[AssignmentsID]
						FROM	FarmersXML01_ActivityUpdateQue
						WHERE	[SendToCRN] = 1
						)
		)
END
ELSE
BEGIN
	IF @SetSendToCRN Is Null
	BEGIN
		--- Return the Property Update flagged for this AssignmentsID
		SELECT 
			[AssignmentsID],
			[UnitId],
			[ClaimId],
			[ContactRowId],
			[PropertyRowId],
			[PropertyDescription],
			[VehicleMake],
			[VehicleModel],
			[VehicleYear],
			[InspectedVIN],
			[LocationType],
			[LocationName],
			[LocationAddress],
			[LocationCity],
			[LocationCountry],
			[LocationState],
			[LocationPostalCode],
			[LocationPhoneNumber],
			[GMTUpdated],
			[CRNAccepted],
			[CRNAcceptedDate],
			[CRNRejected],
			[CRNRejectedDate],
			[CRNRejectedReason],
			[SendToCRN],
			[SentToCRNDate],
			[DateLastUpdated],
			[UpdateByUserID]
		FROM	FarmersXML01_LocationUpdateQue
		WHERE	[SendToCRN] = 1
		AND 	[AssignmentsID] = @AssignmentsID
		AND	[UnitId] = @UnitId
		
		-- Return the Activity update(s) Flagged for this AssignmentsID
		SELECT
			[FarmersXML01_ActivityUpdateQueID],
			[AssignmentsID],
			[UnitId],
			[ClaimId],
			[Type],
			[Description],
			[Comment],
			[CreatedBy],
			[CompletedDate],
			[GMTCreated],
			[CRNAccepted],
			[CRNAcceptedDate],
			[CRNRejected],
			[CRNRejectedDate],
			[CRNRejectedReason],
			[SendToCRN],
			[SentToCRNDate],
			[DateLastUpdated],
			[UpdateByUserID]
		FROM	FarmersXML01_ActivityUpdateQue
		WHERE	[SendToCRN] = 1
		AND 	[AssignmentsID] = @AssignmentsID
		AND	[UnitId] = @UnitId
		ORDER BY [DateLastUpdated]	
	END
	ELSE
	BEGIN
		--Un flag the Send to Crn for Property
		UPDATE 	FarmersXML01_LocationUpdateQue SET
			[SendToCRN] 		= @SetSendToCRN ,
			[SentToCRNDate]		= GetUtcDate(),
			[DateLastUpdated]	= GetDate(),
			[UpdateByUserID]	= [UpdateByUserID]
		WHERE	[AssignmentsID] = @AssignmentsID
		AND	[UnitId] = @UnitId
		AND 	[ClaimID] = @ClaimId

		-- and then unflag the Send to CRN Flag for Activities
		UPDATE 	FarmersXML01_ActivityUpdateQue SET
			[SendToCRN] 		= @SetSendToCRN ,
			[SentToCRNDate]		= GetUtcDate(),
			[DateLastUpdated]	= GetDate(),
			[UpdateByUserID]	= [UpdateByUserID]
		WHERE	[AssignmentsID] = @AssignmentsID
		AND	[UnitId] = @UnitId
		AND 	[ClaimID] = @ClaimId
	END
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[z_spsGetECSFarmersXML01VendorUpdates]  TO [CFUSER]
GO

