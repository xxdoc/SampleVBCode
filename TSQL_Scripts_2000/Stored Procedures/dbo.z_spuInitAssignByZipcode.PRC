SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spuInitAssignByZipcode]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spuInitAssignByZipcode]
GO

/******************************************************************
*
*    	Description:	Initialize Assign By ZipCode
*			Use this SP for All Client Company Loss SP 
*	Author: 	Brad Skidmore
*	Date: 		/21/2004
*	
*	Date		
******************************************************************/

CREATE PROCEDURE [dbo].[z_spuInitAssignByZipcode]
	--Required params
	@UID				int,
	@UAUID				int,
	@ClientCompanyID		int,
	@ClientCompanyCatSpecID		int,
	@ClientCompanyName		varchar(30),
	@iAssignByZIP			int,	
	--Output params
	@AdjusterSpecID			int=null OUTPUT

	

AS
DECLARE @iZIP_ClientCoAdjusterSpecID 	int
DECLARE @iZIP_UsersID 			int
DECLARE	@iZip_CountOFPAZIP 		int
DECLARE @sTempACID 			varchar(50)
DECLARE @sTempACIDDescription 		varchar(150)
DECLARE @sTempGetDate 			datetime
--If Assigning by ZIPCODE need to Check for Users Already Assigned to Specific ZIPCODES for a Specific Cat
--As Well need to Account for the Same ZipCode For a Specific Cat being Assigned to More than one Adjuster.
--This means the Assignement being processed will be Assigned to the adjuster with the Least Number Of Assignment for that Zip Code.
--As Well, IF USER has Already been set up for the ZipCode on the Assignment Being Processed, Need to Assign that to the
--UNASSIGNED Account.  When an Assignment is Assigned to the UNASSIGNED Account,  The Admin will have to Manually Assign the 
--ZIP Code Assigned to the UNASSIGNED Account to an Adjuster Assigned to the same CAT.

--- First Check to See If this ZIP Code IS Already Assigned 
SELECT 	Top 1 
	@iZIP_ClientCoAdjusterSpecID = ClientCoAdjusterSpecID,
	@iZIP_UsersID = UsersID
FROM 	ClientCoAdjusterSpec
WHERE	ClientCompanyCatSpecID = @ClientCompanyCatSpecID
AND 	ZipCode = @iAssignByZIP
AND	InactiveDate Is Null	
Order By ZipCodeCount Asc, ActiveDate Desc 
	
-- If the  @iZIP_ClientCoAdjusterSpecID = @UAUID (UUNASSIGNED USERID)
--That means no user is currenlty Assigned to the ZipCode for the
--assignment being processed, or their Inactive Dates have been Set,
-- But the UUNASSIGNED User has been set up and it's Inactive date is not Yet Set For the Zip Code.

--(The UUNASSIGNED Users's Inactive date Will be Set for a ZIPCODE Once that ZipCode is Assigned to one or more Adjusters)
--**Note*** It is Possible that the UUNASSIGNED User Inactive Date Could be Turned
--Off If an Admin Determines it necessary to Keep Some Assignments under that ZIP in "UNASSIGNED STATUS" Temporarily.
--This would allow for Manual management of the Rate of Assignments going to Adjusters.
--The Admin will see the Zipcode under both the Assigned and UnAssigned Columns on the Admin Screen.
--Once the Admin Is ready to turn off the UNASSIGNED, the admin would just Select the ZIP CODE
-- FROM the UNASSIGNED COLUMN and ASSIGN it to either the Same Adjusters already Assigned to that ZIP
-- or to Other Adjusters.
IF @iZIP_UsersID = @UAUID 
BEGIN
	--Update the @AdjusterSpecID
	SET @AdjusterSpecID = @iZIP_ClientCoAdjusterSpecID
END
--IF the @iZIP_ClientCoAdjusterSpecID IS Null...
--That means no user is currenlty Assigned to the ZipCode for the
--assignment being processed. As Well, The Un Assigned User is 
--either Not Yet Set Up for this ZIP CODE or the Inactive date was Turned On for the UNASSIGNED USer
ELSE IF @iZIP_ClientCoAdjusterSpecID Is Null 
BEGIN
	SELECT 	@iZIP_ClientCoAdjusterSpecID = ClientCoAdjusterSpecID,
		@iZIP_UsersID = UsersID
	FROM 	ClientCoAdjusterSpec
	WHERE	ClientCompanyCatSpecID = @ClientCompanyCatSpecID
	AND 	ZipCode = @iAssignByZIP
	AND	InactiveDate Is Not Null
	AND 	UsersID = @UAUID 
	
	--IF the UUNASSIGNED User is already in the Table
	--Then just need to turn off the inactive Date and Update the @AdjusterSpecID
	IF @iZIP_ClientCoAdjusterSpecID Is Not Null
	BEGIN
		UPDATE ClientCoAdjusterSpec SET 
						[InactiveDate] = Null,
						[DateLastUpdated] = GetDate(),
						[UpdateByUserID] = @UID
		WHERE ClientCompanyCatSpecID = @iZIP_ClientCoAdjusterSpecID
		SET @AdjusterSpecID = @iZIP_ClientCoAdjusterSpecID
	END
	ELSE
	BEGIN
		--IF the UUNASSIGNED USER Is not in there for this ZIPCODE
		-- NEED to ADD It in there.

		SET @sTempACID = 'UUNASSIGNED_' + Convert(varchar(20),@ClientCompanyCatSpecID) + '_' + Convert(varchar(20),@iAssignByZIP) + '_0'
		SET @sTempACIDDescription = @ClientCompanyName + ' ASSIGN BY ZIPCODE'
		SET @sTempGetDate = GetDate()

		EXEC z_spsGetClientCoAdjusterSpecZIPInfo 	
		/*@bHideDeleted */				0,
		/*@UID */					@UID, 
		/*@ClientCoAdjusterSpecID */			1,
		/*@OrderBy */					Null,
		/*@GroupBy */					Null,
		/*@SearchBy */					Null,
		/*@bDelete */					0, 
		/*@bAdd */					1, 
		/*@bEdit */					0,	
		/*@ClientCompanyID	*/			@ClientCompanyID,
		/*@UsersID*/					@UAUID,
		/*@ACID*/					@sTempACID,
		/*@ACIDDescription*/				@sTempACIDDescription,
		/*@Comments	*/				'WEBCONTROL UPDATED',
		/*@ClientCompanyCatSpecID*/			@ClientCompanyCatSpecID,
		/*@ZipCode*/					@iAssignByZIP, 
		/*@ZipCodeCount*/				0,--The ZipCodeCount WILL ALWAYS BE 0 for the UUNASSIGNED USER !!! 
		/*@ActiveDate */				@sTempGetDate,
		/*@InactiveDate */				null
		--UPDATE the @AdjusterSpecID
		SET @AdjusterSpecID = (Select IDENT_CURRENT('ClientCoAdjusterSpec') as AdjusterSpecID)
	END
END 
ELSE
BEGIN
	--IF We Get Here it means there is one or more adjusters assigned to this ZIPCODE for this CAT.
	--Need UPDATE the @AdjusterSpecID with the Adjuster who has the least amount of assignments.
	EXEC z_spuInitCCAssignAdjToZipCode
		@UID					,
		@ClientCompanyCatSpecID			,
		@iAssignByZIP				,
		@iZIP_ClientCoAdjusterSpecID		OUTPUT	
		
	--UPDATE the @AdjusterSpecID
	SET @AdjusterSpecID = @iZIP_ClientCoAdjusterSpecID
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[z_spuInitAssignByZipcode]  TO [CFUSER]
GO

