SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spsGetProdTotal]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spsGetProdTotal]
GO

/******************************************************************
*
*    	Description:	Prod Totals
*			Report request for Katrina Hurc 9.15.2005
*	
*	Author: 	Brad Skidmore
*	Date: 		09/15/2005
******************************************************************/

CREATE PROCEDURE [dbo].[z_spsGetProdTotal]
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
		RetProd.[TotalClaims] As [f003_Total Claims],
		RetProd.[Closed] As [f004_Closed],
		RetProd.[CRNClosed] As [f005_CRN Closed],
		RetProd.[OPEN] As [f006_Open],
		RetProd.[Contacted] As [f007_Contacted],
		Round((Cast(RetProd.[Contacted] as decimal(5,2)) / Cast(RetProd.[TotalClaims] as decimal(5,2))),2)  As [f008_Contacted Pct.],
		RetProd.[Inspected]  As [f009_Inspected],
		Round((Cast(RetProd.[Inspected] as decimal(5,2)) / Cast(RetProd.[TotalClaims] as decimal(5,2))),2) As [f010_Inspected Pct.]
		
	FROM
	(
	SELECT	U.[Lastname] + ', ' + U.[FirstName] As [Adjuster],	
		U.[ContactPhone],
		--U.[EmergencyPhone],
		Count(A.AssignmentsID) As [TotalClaims],
		SUM((CASE WHEN A.[CloseDate] Is not Null Then 1 Else 0 END )) As [Closed],
		SUM((CASE WHEN P.[IsDeleted] =1 Then 1 Else 0 END )) As [CRNClosed],
		SUM((CASE WHEN A.[CloseDate] Is Null Then 1 Else 0 END )) As [OPEN],
		SUM((CASE WHEN A.[ContactDate] Is not Null Then 1 Else 0 END )) As [Contacted],
		SUM((CASE WHEN A.[InspectedDate] Is not Null Then 1 Else 0 END )) As [Inspected]
	FROM Assignments A
		Inner Join ClientCompanyCatSpec CCCS ON A.[ClientCompanyCatSpecID] = CCCS.[ClientCompanyCatSpecID]
		Inner Join ClientCoAdjusterSpec CCAS ON A.[AdjusterSpecID] = CCAS.[ClientCoAdjusterSpecID]
		Inner Join Users U On CCAS.[USersID] = U.[UsersID]
		Left Join PolicyLimits P ON A.[AssignmentsID] = P.[AssignmentsID] And P.[FarmersUnitRowID] <> ''
	WHERE A.[IsDeleted] = 0
	AND A.[AssignmentTypeID] = @AssignmentTypeID 
	AND CCCS.[ClientCompanyID] = @ClientCompanyID
	AND CCCS.[CatID] = @CatID
	Group By U.[Lastname] + ', ' + U.[FirstName], U.[ContactPhone], U.[EmergencyPhone]
	) As RetProd
	Order By Adjuster
END
ELSE
BEGIN
			
	SELECT 	
		RetProd.[Adjuster] As [f001_Adjuster],
		RetProd.[ContactPhone] As [f002_Phone],
		RetProd.[TotalClaims] As [f003_Total Claims],
		RetProd.[Closed] As [f004_Closed],
		RetProd.[CRNClosed] As [f005_CRN Closed],
		RetProd.[OPEN] As [f006_Open],
		RetProd.[Contacted] As [f007_Contacted],
		Round((Cast(RetProd.[Contacted] as decimal(5,2)) / Cast(RetProd.[TotalClaims] as decimal(5,2))),2)  As [f008_Contacted Pct.],
		RetProd.[Inspected] As [f009_Inspected],
		Round((Cast(RetProd.[Inspected] as decimal(5,2)) / Cast(RetProd.[TotalClaims] as decimal(5,2))),2) As [f010_Inspected Pct.]
		
	FROM
	(
	SELECT	U.[Lastname] + ', ' + U.[FirstName] As [Adjuster],	
		U.[ContactPhone],
		--U.[EmergencyPhone],
		Count(A.AssignmentsID) As [TotalClaims],
		SUM((CASE WHEN A.[CloseDate] Is not Null Then 1 Else 0 END )) As [Closed],
		SUM((CASE WHEN P.[IsDeleted] =1 Then 1 Else 0 END )) As [CRNClosed],
		SUM((CASE WHEN A.[CloseDate] Is Null Then 1 Else 0 END )) As [OPEN],
		SUM((CASE WHEN A.[ContactDate] Is not Null Then 1 Else 0 END )) As [Contacted],
		SUM((CASE WHEN A.[InspectedDate] Is not Null Then 1 Else 0 END )) As [Inspected]
	FROM Assignments A
		Inner Join ClientCompanyCatSpec CCCS ON A.[ClientCompanyCatSpecID] = CCCS.[ClientCompanyCatSpecID]
		Inner Join ClientCoAdjusterSpec CCAS ON A.[AdjusterSpecID] = CCAS.[ClientCoAdjusterSpecID]
		Inner Join Users U On CCAS.[USersID] = U.[UsersID]
		Left Join PolicyLimits P ON A.[AssignmentsID] = P.[AssignmentsID] And P.[FarmersUnitRowID] <> ''
	WHERE A.[IsDeleted] = 0
	AND A.[AssignmentTypeID] = @AssignmentTypeID 
	AND CCAS.[UsersID] = @USERSID
	AND CCCS.[ClientCompanyID] = @ClientCompanyID
	AND CCCS.[CatID] = @CatID
	Group By U.[Lastname] + ', ' + U.[FirstName], U.[ContactPhone], U.[EmergencyPhone]
	) As RetProd
	Order By Adjuster
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[z_spsGetProdTotal]  TO [CFUSER]
GO

