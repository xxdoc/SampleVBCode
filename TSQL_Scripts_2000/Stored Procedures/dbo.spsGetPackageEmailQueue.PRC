SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spsGetPackageEmailQueue]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[spsGetPackageEmailQueue]
GO

/******************************************************************
*
*    	Description:	Get Packges Qued for Email...
*			To be processed
*	
*	Author: 	Brad Skidmore
*	Date: 		07/28/2005
*			10/04/2005  Added DateDontSendUntil
			Do not include any records that have
			the DateDontSendUntil >  the current date
******************************************************************/
CREATE PROCEDURE [dbo].[spsGetPackageEmailQueue]
	@PackageEmailQueueID 		int=null,
	@bDelPackageEmailQueueID	bit=0
AS

IF @PackageEmailQueueID Is Null
BEGIN
	SELECT TOP 1	
		CO.[Name] 		As [ClientCoName],
		U.[UserName] 		As [AdjUserName],
		PKG.[AssignmentsID],
		PKG.[PackageID],
		CO.[CarrierPrefix] + CO.[DBName] + '.clsLists' As [CarListClassName],
		PKG.[PackageStatus],
		PEQ.[PackageEmailQueueID],
		PEQ.[PackageID]		As [PEQPackageID],
		PEQ.[AssignmentsID]	As [PEQAssignmentsID],
		PEQ.[PackageStatus]	As [PEQPackageStatus],
		PEQ.[EmailTo]		As [PEQEmailTo],
		PEQ.[EmailCC]		As [PEQEmailCC],
		PEQ.[EmailBCC]		As [PEQEmailBCC],
		PEQ.[EmailSubject]	As [PEQEmailSubject],
		PEQ.[EmailMess]		As [PEQEmailMess],
		PEQ.[PackageItemIDList]	As [PEQPackageItemIDList],
		PEQ.[DateLastUpdated]	As [PEQDateLastUpdated],
		PEQ.[UpdateByUserID]	As [PEQUpdateByUserID],
		(LEN(PEQ.[PackageItemIDList])) As [PrioritySort], --The smaller packages get to go first, then date time entered
		PEQ.[DateDontSendUntil] As [PEQDateDontSendUntil] --When items are entered and this date is not specifically set it will default to the current date
	FROM	PackageEmailQueue PEQ
		INNER JOIN Package PKG ON PKG.[PackageID] = PEQ.[PackageID]
		INNER JOIN Assignments A ON A.[AssignmentsID] = PKG.[AssignmentsID]
		INNER JOIN ClientCompanyCatSpec CCCS ON CCCS.[ClientCompanyCatSpecID] = A.[ClientCompanyCatSpecID]
		INNER JOIN ClientCoAdjusterSpec CCAS ON CCAS.[ClientCoAdjusterSpecID] = A.[AdjusterSpecID]
		INNER JOIN Users U ON U.UsersID = CCAS.[UsersID]
		INNER JOIN Company CO ON CO.[CompanyID] = CCCS.[ClientCompanyID]
		INNER JOIN ClientCompanyCat CCCat ON CCCat.[ClientCompanyID] = CCCS.[ClientCompanyID] And CCCat.[CATID] = CCCS.[CATID]
		WHERE GetDate() >= PEQ.[DateDontSendUntil] --If the current date is later or = to the DateDontSenduntil then it will be included in the results
	ORDER BY [PrioritySort], PEQ.[DateLastUpdated] 
END
ELSE
BEGIN
	IF @bDelPackageEmailQueueID = 1
	BEGIN
		DELETE 	FROM PackageEmailQueue 
		WHERE  	[PackageEmailQueueID] = @PackageEmailQueueID
	END
	ELSE
	BEGIN
		SELECT 	
			CO.[Name] 		As [ClientCoName],
			U.[UserName] 		As [AdjUserName],
			PKG.[AssignmentsID],
			PKG.[PackageID],
			CO.[CarrierPrefix] + CO.[DBName] + '.clsLists' As [CarListClassName],
			PKG.[PackageStatus],
			PEQ.[PackageEmailQueueID],
			PEQ.[PackageID]		As [PEQPackageID],
			PEQ.[AssignmentsID]	As [PEQAssignmentsID],
			PEQ.[PackageStatus]	As [PEQPackageStatus],
			PEQ.[EmailTo]		As [PEQEmailTo],
			PEQ.[EmailCC]		As [PEQEmailCC],
			PEQ.[EmailBCC]		As [PEQEmailBCC],
			PEQ.[EmailSubject]	As [PEQEmailSubject],
			PEQ.[EmailMess]		As [PEQEmailMess],
			PEQ.[PackageItemIDList]	As [PEQPackageItemIDList],
			PEQ.[DateLastUpdated]	As [PEQDateLastUpdated],
			PEQ.[UpdateByUserID]	As [PEQUpdateByUserID],
			PEQ.[DateDontSendUntil] As [PEQDateDontSendUntil]
		FROM	PackageEmailQueue PEQ
			INNER JOIN Package PKG ON PKG.[PackageID] = PEQ.[PackageID]
			INNER JOIN Assignments A ON A.[AssignmentsID] = PKG.[AssignmentsID]
			INNER JOIN ClientCompanyCatSpec CCCS ON CCCS.[ClientCompanyCatSpecID] = A.[ClientCompanyCatSpecID]
			INNER JOIN ClientCoAdjusterSpec CCAS ON CCAS.[ClientCoAdjusterSpecID] = A.[AdjusterSpecID]
			INNER JOIN Users U ON U.UsersID = CCAS.[UsersID]
			INNER JOIN Company CO ON CO.[CompanyID] = CCCS.[ClientCompanyID]
			INNER JOIN ClientCompanyCat CCCat ON CCCat.[ClientCompanyID] = CCCS.[ClientCompanyID] And CCCat.[CATID] = CCCS.[CATID]
		WHERE   PEQ.[PackageEmailQueueID] = @PackageEmailQueueID
		ORDER BY PEQ.[DateLastUpdated] 
	END
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[spsGetPackageEmailQueue]  TO [CFUSER]
GO

