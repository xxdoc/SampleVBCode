SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spsGetPackageToProcess]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[spsGetPackageToProcess]
GO

/******************************************************************
*
*    	Description:	Get Package to Process
*			Will get the Package in line
*			To be processed
*	
*	Author: 	Brad Skidmore
*	Date: 		06/14/2005
*
******************************************************************/
CREATE PROCEDURE [dbo].[spsGetPackageToProcess] 
AS
--6.14.2005 BGS Check for packges to be processed!
SELECT TOP 1	
	CO.[Name] As [ClientCoName],
	U.[UserName] As [AdjUserName],
	PKG.[AssignmentsID],
	PKG.[PackageID],
	CO.[CarrierPrefix] + CO.[DBName] + '.clsLists' As [CarListClassName],
	PKG.[PackageStatus]
FROM	Package PKG
	INNER JOIN Assignments A ON A.[AssignmentsID] = PKG.[AssignmentsID]
	INNER JOIN ClientCompanyCatSpec CCCS ON CCCS.[ClientCompanyCatSpecID] = A.[ClientCompanyCatSpecID]
	INNER JOIN ClientCoAdjusterSpec CCAS ON CCAS.[ClientCoAdjusterSpecID] = A.[AdjusterSpecID]
	INNER JOIN Users U ON U.UsersID = CCAS.[UsersID]
	INNER JOIN Company CO ON CO.[CompanyID] = CCCS.[ClientCompanyID]
	INNER JOIN ClientCompanyCat CCCat ON CCCat.[ClientCompanyID] = CCCS.[ClientCompanyID] And CCCat.[CATID] = CCCS.[CATID]
WHERE	PKG.[SendMe] = 1
AND	CO.[EnableSingleFile] = 1 	--If this is False No Client Cats will be processed
AND	CCCat.[EnableSingleFile] = 1	--If this is False and the Above is True then Just this Client Cat Will not be processed
ORDER BY PKG.[DateLastUpdated] 

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[spsGetPackageToProcess]  TO [CFUSER]
GO

