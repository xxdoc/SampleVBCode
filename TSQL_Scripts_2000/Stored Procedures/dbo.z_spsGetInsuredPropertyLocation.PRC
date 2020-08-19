SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spsGetInsuredPropertyLocation]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spsGetInsuredPropertyLocation]
GO

/******************************************************************
*
*    	Description:	Insured Property Location
*			Report request for Katrina Hurc 10.15.2005
*		
*	Author: 	Brad Skidmore
*	Date: 		10/15/2005
******************************************************************/

CREATE PROCEDURE [dbo].[z_spsGetInsuredPropertyLocation]
	@AssignmentTypeID		int,
	@ClientCompanyID		int,
	@CatID				int,
	@USERSID			int=null
	
AS
IF @USERSID is Null
BEGIN
	SELECT 	
		RetProd.[Adjuster] As [f001_Adjuster],
		RetProd.[ContactPhone] As [f002_Phone],
		RetProd.[CLIENTNUM] As [f003a_Claim No],
		RetProd.[StatusAlias] As [f003b_Status],
		RetProd.[Insured] As [f004_Insured],
		RetProd.[PAStreet] As [f005_Prop Street],
		RetProd.[PACity] As [f006_Prop City],
		RetProd.[PAState] As [f007_Prop State],
		RetProd.[PAZip] As [f008_Prop Zip],
		RetProd.[MAStreet] As [f009_Mail Street],
		RetProd.[MACity] As [f010_Mail City],
		RetProd.[MAState] As [f011_Mail State],
		RetProd.[MAZip] As [f012_Mail Zip],
		RetProd.[UserName] As [z_Hidden_UserName]
	FROM
	(
	SELECT	
		U.[UserName],
		U.[Lastname] + ', ' + U.[FirstName] As [Adjuster],	
		U.[ContactPhone],
		--U.[EmergencyPhone],
		A.[CLIENTNUM],
		S.[StatusAlias],
		A.[Insured],
		A.[PAStreet],
		A.[PACity],
		A.[PAState],
		A.[PAZip],
		A.[MAStreet],
		A.[MACity],
		A.[MAState],
		A.[MAZip]
	FROM Assignments A
		Inner Join ClientCompanyCatSpec CCCS ON A.[ClientCompanyCatSpecID] = CCCS.[ClientCompanyCatSpecID]
		Inner Join ClientCoAdjusterSpec CCAS ON A.[AdjusterSpecID] = CCAS.[ClientCoAdjusterSpecID]
		Inner Join Users U On CCAS.[USersID] = U.[UsersID]
		Inner Join Status S On S.[StatusID] = A.[StatusID]
		Left Join PolicyLimits P ON A.[AssignmentsID] = P.[AssignmentsID] And P.[FarmersUnitRowID] <> ''
	WHERE A.[IsDeleted] = 0
	AND A.[AssignmentTypeID] = @AssignmentTypeID 
	AND CCCS.[ClientCompanyID] = @ClientCompanyID
	AND CCCS.[CatID] = @CatID
	Group By 	U.[Lastname] + ', ' + U.[FirstName],
			U.[UserName], 
			U.[ContactPhone], 
			U.[EmergencyPhone],
			A.[CLIENTNUM],
			S.[StatusAlias],
			A.[Insured],
			A.[PAStreet],
			A.[PACity],
			A.[PAState],
			A.[PAZip],
			A.[MAStreet],
			A.[MACity],
			A.[MAState],
			A.[MAZip]
	) As RetProd
	Order By RetProd.[Adjuster], RetProd.[CLIENTNUM]
END
ELSE
BEGIN
			
	SELECT 	
		--RetProd.[Adjuster] As [f001_Adjuster],
		--RetProd.[ContactPhone] As [f002_Phone],
		RetProd.[CLIENTNUM] As [f003a_Claim No],
		RetProd.[StatusAlias] As [f003b_Status],
		RetProd.[Insured] As [f004_Insured],
		RetProd.[PAStreet] As [f005_Prop Street],
		RetProd.[PACity] As [f006_Prop City],
		RetProd.[PAState] As [f007_Prop State],
		RetProd.[PAZip] As [f008_Prop Zip],
		RetProd.[MAStreet] As [f009_Mail Street],
		RetProd.[MACity] As [f010_Mail City],
		RetProd.[MAState] As [f011_Mail State],
		RetProd.[MAZip] As [f012_Mail Zip],
		RetProd.[UserName] As [z_Hidden_UserName]
	FROM
	(
	SELECT	U.[UserName],
		U.[Lastname] + ', ' + U.[FirstName] As [Adjuster],	
		U.[ContactPhone],
		--U.[EmergencyPhone],
		A.[CLIENTNUM],
		S.[StatusAlias],
		A.[Insured],
		A.[PAStreet],
		A.[PACity],
		A.[PAState],
		A.[PAZip],
		A.[MAStreet],
		A.[MACity],
		A.[MAState],
		A.[MAZip]
	FROM Assignments A
		Inner Join ClientCompanyCatSpec CCCS ON A.[ClientCompanyCatSpecID] = CCCS.[ClientCompanyCatSpecID]
		Inner Join ClientCoAdjusterSpec CCAS ON A.[AdjusterSpecID] = CCAS.[ClientCoAdjusterSpecID]
		Inner Join Users U On CCAS.[USersID] = U.[UsersID]
		Inner Join Status S On S.[StatusID] = A.[StatusID]
		Left Join PolicyLimits P ON A.[AssignmentsID] = P.[AssignmentsID] And P.[FarmersUnitRowID] <> ''
	WHERE A.[IsDeleted] = 0
	AND A.[AssignmentTypeID] = @AssignmentTypeID 
	AND CCAS.[UsersID] = @USERSID
	AND CCCS.[ClientCompanyID] = @ClientCompanyID
	AND CCCS.[CatID] = @CatID
	Group By 	U.[Lastname] + ', ' + U.[FirstName], 
			U.[UserName],
			U.[ContactPhone], 
			U.[EmergencyPhone],
			A.[CLIENTNUM],
			S.[StatusAlias],
			A.[Insured],
			A.[PAStreet],
			A.[PACity],
			A.[PAState],
			A.[PAZip],
			A.[MAStreet],
			A.[MACity],
			A.[MAState],
			A.[MAZip]
	) As RetProd
	Order By RetProd.[CLIENTNUM]
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[z_spsGetInsuredPropertyLocation]  TO [CFUSER]
GO

