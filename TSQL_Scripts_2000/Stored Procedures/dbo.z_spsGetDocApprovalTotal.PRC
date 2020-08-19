SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spsGetDocApprovalTotal]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spsGetDocApprovalTotal]
GO

/******************************************************************
*
*    	Description:	Client Document Approval Totals
*			Report request for Katrina Hurc 9.15.2005
*	
*	Author: 	Brad Skidmore
*	Date: 		09/17/2005
******************************************************************/

CREATE PROCEDURE [dbo].[z_spsGetDocApprovalTotal]
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
		RetProd.[TotalDocs] As [f003_Total Docs],
		RetProd.[TotalIsDeleted] As [f004_Deleted],
		--RetProd.[TotalAttachments] As [f005_Total Attachments],
		RetProd.[TotalCoApprove] As [f006_Co Approved],
		RetProd.[TotalSent] As [f007_Sent to Client],
-- 		(
-- 		CASE 	WHEN RetProd.[TotalNumSentToClientApprove] > 0 
-- 			THEN TotalDaysFromSentToClientApprove / TotalNumSentToClientApprove
-- 			ELSE 0
-- 			END
-- 		) As [f007_Avg Days Until Client Approval],
		RetProd.[TotalClientCoReject] As [f008_Client Rejecting],
		RetProd.[TotalClientCoDelete] As [f009_Client Deleted],
		RetProd.[TotalClientCoApprove] As [f010_Client Approved]
		--RetProd.[TotalSendMe] As [f011_Need to be sent],
-- 		(
-- 		CASE	WHEN (RetProd.[TotalDocs] - RetProd.[TotalIsDeleted] - RetProd.[TotalSent]) < RetProd.[TotalSendMe]
-- 			THEN RetProd.[TotalSendMe] - (RetProd.[TotalDocs] - RetProd.[TotalIsDeleted] - RetProd.[TotalSent])
-- 			ELSE 0
-- 			END
-- 		) As [f012_Resending],
		
		
	FROM
	(SELECT
		U.[Lastname] + ', ' + U.[FirstName] As [Adjuster],	
		U.[ContactPhone],
		Count(PKGI.[PackageItemID]) As [TotalDocs],
		SUM((CASE WHEN PKGI.[IsDeleted] =1 Then 1 Else 0 END )) As [TotalIsDeleted],
		Count(PKGI.[RTAttachmentsID]) As [TotalAttachments],
		SUM((CASE WHEN PKGI.[IsCoApprove] =1 Then 1 Else 0 END )) As [TotalCoApprove],
		SUM((CASE WHEN PKGI.[IsClientCoReject] =1 Then 1 Else 0 END )) As [TotalClientCoReject],
		SUM((CASE WHEN PKGI.[IsClientCoDelete] =1 Then 1 Else 0 END )) As [TotalClientCoDelete],
		SUM((CASE WHEN PKGI.[IsClientCoApprove] =1 Then 1 Else 0 END )) As [TotalClientCoApprove],
		SUM((CASE WHEN PKGI.[SendMe] =1 Then 1 Else 0 END )) As [TotalSendMe],
		SUM((CASE WHEN PKGI.[SentDate] Is not Null Then 1 Else 0 END )) As [TotalSent]
-- 		SUM	(
-- 				(
-- 				CASE 	WHEN PKGI.[SentDate] Is not Null And PKGI.[ClientCoApproveDate] Is not Null  
-- 					THEN 	(
-- 						CASE	WHEN PKGI.[ClientCoRejectDate] Is Not Null
-- 							THEN DATEDIFF(d, PKGI.[ClientCoRejectDate], PKGI.[ClientCoApproveDate]) 
-- 							ELSE DATEDIFF(d, PKGI.[SentDate], PKGI.[ClientCoApproveDate]) 
-- 							END
-- 						)
-- 					ELSE 0 
-- 					END 
-- 				)
-- 			) As [TotalDaysFromSentToClientApprove],
-- 		SUM	(
-- 				(
-- 				CASE 	WHEN PKGI.[SentDate] Is not Null And PKGI.[ClientCoApproveDate] Is not Null 
-- 					THEN 1	
-- 					ELSE 0 
-- 					END 
-- 				)
-- 			) As [TotalNumSentToClientApprove]
		
	FROM Assignments A
		Inner Join ClientCompanyCatSpec CCCS ON A.[ClientCompanyCatSpecID] = CCCS.[ClientCompanyCatSpecID]
		Inner Join ClientCoAdjusterSpec CCAS ON A.[AdjusterSpecID] = CCAS.[ClientCoAdjusterSpecID]
		Inner Join Users U On CCAS.[USersID] = U.[UsersID]
		Left join PackageItem PKGI ON PKGI.[AssignmentsID] = A.[AssignmentsID]
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
		RetProd.[TotalDocs] As [f003_Total Docs],
		RetProd.[TotalIsDeleted] As [f004_Deleted],
		--RetProd.[TotalAttachments] As [f005_Total Attachments],
		RetProd.[TotalCoApprove] As [f006_Co Approved],
		RetProd.[TotalSent] As [f007_Sent to Client],
-- 		(
-- 		CASE 	WHEN RetProd.[TotalNumSentToClientApprove] > 0 
-- 			THEN TotalDaysFromSentToClientApprove / TotalNumSentToClientApprove
-- 			ELSE 0
-- 			END
-- 		) As [f007_Avg Days Until Client Approval],
		RetProd.[TotalClientCoReject] As [f008_Client Rejecting],
		RetProd.[TotalClientCoDelete] As [f009_Client Deleted],
		RetProd.[TotalClientCoApprove] As [f010_Client Approved]
		--RetProd.[TotalSendMe] As [f011_Need to be sent],
-- 		(
-- 		CASE	WHEN (RetProd.[TotalDocs] - RetProd.[TotalIsDeleted] - RetProd.[TotalSent]) < RetProd.[TotalSendMe]
-- 			THEN RetProd.[TotalSendMe] - (RetProd.[TotalDocs] - RetProd.[TotalIsDeleted] - RetProd.[TotalSent])
-- 			ELSE 0
-- 			END
-- 		) As [f012_Resending],
		
	FROM
	(
	SELECT	
		U.[Lastname] + ', ' + U.[FirstName] As [Adjuster],	
		U.[ContactPhone],
		Count(PKGI.[PackageItemID]) As [TotalDocs],
		SUM((CASE WHEN PKGI.[IsDeleted] =1 Then 1 Else 0 END )) As [TotalIsDeleted],
		Count(PKGI.[RTAttachmentsID]) As [TotalAttachments],
		SUM((CASE WHEN PKGI.[IsCoApprove] =1 Then 1 Else 0 END )) As [TotalCoApprove],
		SUM((CASE WHEN PKGI.[IsClientCoReject] =1 Then 1 Else 0 END )) As [TotalClientCoReject],
		SUM((CASE WHEN PKGI.[IsClientCoDelete] =1 Then 1 Else 0 END )) As [TotalClientCoDelete],
		SUM((CASE WHEN PKGI.[IsClientCoApprove] =1 Then 1 Else 0 END )) As [TotalClientCoApprove],
		SUM((CASE WHEN PKGI.[SendMe] =1 Then 1 Else 0 END )) As [TotalSendMe],
		SUM((CASE WHEN PKGI.[SentDate] Is not Null Then 1 Else 0 END )) As [TotalSent]
-- 		SUM	(
-- 				(
-- 				CASE 	WHEN PKGI.[SentDate] Is not Null And PKGI.[ClientCoApproveDate] Is not Null  
-- 					THEN 	(
-- 						CASE	WHEN PKGI.[ClientCoRejectDate] Is Not Null
-- 							THEN DATEDIFF(d, PKGI.[ClientCoRejectDate], PKGI.[ClientCoApproveDate]) 
-- 							ELSE DATEDIFF(d, PKGI.[SentDate], PKGI.[ClientCoApproveDate]) 
-- 							END
-- 						)
-- 					ELSE 0 
-- 					END 
-- 				)
-- 			) As [TotalDaysFromSentToClientApprove],
-- 		SUM	(
-- 				(
-- 				CASE 	WHEN PKGI.[SentDate] Is not Null And PKGI.[ClientCoApproveDate] Is not Null 
-- 					THEN 1	
-- 					ELSE 0 
-- 					END 
-- 				)
-- 			) As [TotalNumSentToClientApprove]
	FROM Assignments A
		Inner Join ClientCompanyCatSpec CCCS ON A.[ClientCompanyCatSpecID] = CCCS.[ClientCompanyCatSpecID]
		Inner Join ClientCoAdjusterSpec CCAS ON A.[AdjusterSpecID] = CCAS.[ClientCoAdjusterSpecID]
		Inner Join Users U On CCAS.[USersID] = U.[UsersID]
		Left join PackageItem PKGI ON PKGI.[AssignmentsID] = A.[AssignmentsID]
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

GRANT  EXECUTE  ON [dbo].[z_spsGetDocApprovalTotal]  TO [CFUSER]
GO

