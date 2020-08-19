SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spsGetInsuredContactInfo]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spsGetInsuredContactInfo]
GO

/******************************************************************
*
*    	Description:	Insured Contact Report
*			Report request for Katrina Hurc 10.04.2005
*		
*	Author: 	Brad Skidmore
*	Date: 		10/04/2005
******************************************************************/

CREATE PROCEDURE [dbo].[z_spsGetInsuredContactInfo]
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
		RetProd.[HomePhone] As [f005_Home Phone],
		RetProd.[BusinessPhone] As [f006_Work Phone],
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
		A.[HomePhone],
		A.[BusinessPhone]
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
			A.[HomePhone],
			A.[BusinessPhone]
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
		RetProd.[HomePhone] As [f005_Home Phone],
		RetProd.[BusinessPhone] As [f006_Work Phone],
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
		A.[HomePhone],
		A.[BusinessPhone]
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
			A.[HomePhone],
			A.[BusinessPhone]
	) As RetProd
	Order By RetProd.[CLIENTNUM]
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[z_spsGetInsuredContactInfo]  TO [CFUSER]
GO

