SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[insMiscReportParam05]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[insMiscReportParam05]
GO

/******************************************************************
*
*    	Description:	Allows for snagging Items to do upates elsewhere 
*			and then allow the actual Update to take place.
*			
*	Author: 	Brad Skidmore
*	Date: 		5/20/2005 at 6:38 PM 
*
******************************************************************/
CREATE TRIGGER insMiscReportParam05
ON dbo.MiscReportParam05
FOR INSERT
AS
--5.19.2005 BGS  Farmers XML01 Need to check for certain updates Like InspectedVIN !!!
-- When these changes occur need to update the FarmersXML01_LocationUpdateQue with the updated Item(s)
-- and then flag them to be sent off.
--Only do this if one Record assignment at a time is being updated
IF (SELECT COUNT(INS.[ID]) FROM INSERTED INS) = 1
BEGIN
	--- Ignore updates from CFUSER !!!
	IF 	(SELECT INS.[ClassName] FROM INSERTED INS) = 'ECRptFarmers_arActivityLog02.clsRptDB'
		AND
		(SELECT INS.[ParamName] FROM INSERTED INS) = 'f_InspectedVIN'
		AND
		(SELECT	[UserName] FROM USERS WHERE [UsersID] = (SELECT INS.[UpdateByUserID] FROM INSERTED INS)) <> 'CFUSER'
	BEGIN
		--Inspected Vin Param value
		DECLARE @INSParamValue as varchar(2000)
		--ID vars
		DECLARE @AssignmentsID as int
		DECLARE @UID as int
		--- Property Que Vars
		DECLARE @InspectedVIN  as varchar(30)
		DECLARE @GMTUpdated as datetime

		--First Set all the Inserted and Deleted Vars so they
		--Can be compared!
		--ID vars
		SET @AssignmentsID = 	(SELECT INS.[AssignmentsID] FROM INSERTED INS)
		SET @UID =		(SELECT INS.[UpdateByUserID] FROM INSERTED INS)	
		--Inserted
		SET @INSParamValue = 	(SELECT INS.[ParamValue] FROM INSERTED INS)

		--If there are any differences need to update entire Property Location
		IF @INSParamValue <> ''
		BEGIN
			--- Build the Property Field
			SET @InspectedVIN =  	Convert(varchar(30),@INSParamValue)
			SET @GMTUpdated = 	GetUTCDate()
			UPDATE 	FarmersXML01_LocationUpdateQue SET
				[InspectedVIN] = @InspectedVIN,
				[GMTUpdated] = @GMTUpdated,
				[SendToCRN] = 1,
				[DateLastUpdated] = GetDate(),
				[UpdateByUserID] = @UID				
			WHERE	[AssignmentsID] = @AssignmentsID
		END
	END
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

