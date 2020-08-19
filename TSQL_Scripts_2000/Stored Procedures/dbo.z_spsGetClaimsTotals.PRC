SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spsGetClaimsTotals]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spsGetClaimsTotals]
GO

/******************************************************************
*
*    	Description:	Claims Totals
*			Summary Totals for Assignments, Or BillAssignment Table
*	
*	Author: 	Brad Skidmore
*	Date: 		03/18/2005
******************************************************************/

CREATE PROCEDURE [dbo].[z_spsGetClaimsTotals]
	@bGetBillAssignment		bit=0,
	@AssignmentTypeID		int,
	@ClientCompanyID		int,
	@CatID				int,
	@USERSID			int=null
AS

IF @bGetBillAssignment = 0
BEGIN
	IF @USERSID is Null
	BEGIN
		SELECT 	RetStatTotal.[CountOfStatusItem],
			RetStatTotal.[StatusID],
			RetStatTotal.[Status],
			RetStatTotal.[StatusAlias],
			RetStatTotal.[Description],
			RetStatTotal.[TotalDocs],
			RetStatTotal.[TotalIsDeleted],
			RetStatTotal.[TotalAttachments],
			RetStatTotal.[TotalCoApprove],
			RetStatTotal.[TotalClientCoReject],
			RetStatTotal.[TotalClientCoDelete],
			RetStatTotal.[TotalClientCoApprove],
			RetStatTotal.[TotalSendMe],
			RetStatTotal.[TotalSent]
		FROM (
			SELECT	RetStat.[CountOfStatusItem],
				RetStat.[StatusID],
				RetStat.[Status],
				RetStat.[StatusAlias],
				RetStat.[Description],
				Count(PKGI.[PackageItemID]) As [TotalDocs],
				SUM((CASE WHEN PKGI.[IsDeleted] =1 Then 1 Else 0 END )) As [TotalIsDeleted],
				Count(PKGI.[RTAttachmentsID]) As [TotalAttachments],
				SUM((CASE WHEN PKGI.[IsCoApprove] =1 Then 1 Else 0 END )) As [TotalCoApprove],
				SUM((CASE WHEN PKGI.[IsClientCoReject] =1 Then 1 Else 0 END )) As [TotalClientCoReject],
				SUM((CASE WHEN PKGI.[IsClientCoDelete] =1 Then 1 Else 0 END )) As [TotalClientCoDelete],
				SUM((CASE WHEN PKGI.[IsClientCoApprove] =1 Then 1 Else 0 END )) As [TotalClientCoApprove],
				SUM((CASE WHEN PKGI.[SendMe] =1 And PKGI.[IsDeleted]= 0 Then 1 Else 0 END )) As [TotalSendMe],
				SUM((CASE WHEN PKGI.[SentDate] Is not Null Then 1 Else 0 END )) As [TotalSent]
			FROM (
				SELECT  
					COUNT(A.[StatusID]) As CountOfStatusItem,
					S.[StatusID],
					S.[Status],
					S.[StatusAlias],
					S.[Description]
				FROM Status S
					Inner Join Assignments A ON S.[StatusID] = A.[StatusID]
					Inner Join ClientCompanyCatSpec CCCS ON A.[ClientCompanyCatSpecID] = CCCS.[ClientCompanyCatSpecID]
				WHERE S.[IsDeleted] = 0
				AND A.[IsDeleted] = 0
				AND A.[AssignmentTypeID] = @AssignmentTypeID
				AND A.[CLIENTNUM] Not Like '%TEST%'
				AND (A.[Reassigned] Is Null Or A.[Reassigned] = 0)
				AND CCCS.[ClientCompanyID] = @ClientCompanyID
				AND CCCS.[CatID] = @CatID
				Group By S.[Status], S.[StatusAlias], S.[Description], S.[StatusID]
			) As RetStat
				Left Join Assignments A2 ON A2.[StatusID] = RetStat.[StatusID]
				Left Join ClientCompanyCatSpec CCCS2 ON A2.[ClientCompanyCatSpecID] = CCCS2.[ClientCompanyCatSpecID]
				Left join PackageItem PKGI ON PKGI.[AssignmentsID] = A2.[AssignmentsID]
			WHERE A2.[IsDeleted] = 0
			AND A2.[AssignmentTypeID] = @AssignmentTypeID 
			AND A2.[CLIENTNUM] Not Like '%TEST%'
			AND (A2.[Reassigned] Is Null Or A2.[Reassigned] = 0)
			AND CCCS2.[ClientCompanyID] = @ClientCompanyID
			AND CCCS2.[CatID] = @CatID
			Group By 	RetStat.[CountOfStatusItem],
					RetStat.[StatusID],
					RetStat.[Status],
					RetStat.[StatusAlias],
					RetStat.[Description]
		) As RetStatTotal
		Group By 	RetStatTotal.[CountOfStatusItem], 
				RetStatTotal.[Status], 
				RetStatTotal.[StatusAlias], 
				RetStatTotal.[Description], 
				RetStatTotal.[StatusID],
				RetStatTotal.[TotalDocs],
				RetStatTotal.[TotalIsDeleted],
				RetStatTotal.[TotalAttachments],
				RetStatTotal.[TotalCoApprove],
				RetStatTotal.[TotalClientCoReject],
				RetStatTotal.[TotalClientCoDelete],
				RetStatTotal.[TotalClientCoApprove],
				RetStatTotal.[TotalSendMe],
				RetStatTotal.[TotalSent]
		Order by RetStatTotal.[CountOfStatusItem] DESC
	END
	ELSE
	BEGIN
		SELECT 	RetStatTotal.[CountOfStatusItem],
			RetStatTotal.[StatusID],
			RetStatTotal.[Status],
			RetStatTotal.[StatusAlias],
			RetStatTotal.[Description],
			RetStatTotal.[TotalDocs],
			RetStatTotal.[TotalIsDeleted],
			RetStatTotal.[TotalAttachments],
			RetStatTotal.[TotalCoApprove],
			RetStatTotal.[TotalClientCoReject],
			RetStatTotal.[TotalClientCoDelete],
			RetStatTotal.[TotalClientCoApprove],
			RetStatTotal.[TotalSendMe],
			RetStatTotal.[TotalSent]
		FROM (
			SELECT	RetStat.[CountOfStatusItem],
				RetStat.[StatusID],
				RetStat.[Status],
				RetStat.[StatusAlias],
				RetStat.[Description],
				Count(PKGI.[PackageItemID]) As [TotalDocs],
				SUM((CASE WHEN PKGI.[IsDeleted] =1 Then 1 Else 0 END )) As [TotalIsDeleted],
				Count(PKGI.[RTAttachmentsID]) As [TotalAttachments],
				SUM((CASE WHEN PKGI.[IsCoApprove] =1 Then 1 Else 0 END )) As [TotalCoApprove],
				SUM((CASE WHEN PKGI.[IsClientCoReject] =1 Then 1 Else 0 END )) As [TotalClientCoReject],
				SUM((CASE WHEN PKGI.[IsClientCoDelete] =1 Then 1 Else 0 END )) As [TotalClientCoDelete],
				SUM((CASE WHEN PKGI.[IsClientCoApprove] =1 Then 1 Else 0 END )) As [TotalClientCoApprove],
				SUM((CASE WHEN PKGI.[SendMe] =1 And PKGI.[IsDeleted]= 0 Then 1 Else 0 END )) As [TotalSendMe],
				SUM((CASE WHEN PKGI.[SentDate] Is not Null Then 1 Else 0 END )) As [TotalSent]
			FROM (
				SELECT  
					COUNT(A.[StatusID]) As CountOfStatusItem,
					S.[StatusID],
					S.[Status],
					S.[StatusAlias],
					S.[Description]
				FROM Status S
					Inner Join Assignments A ON S.[StatusID] = A.[StatusID]
					Inner Join ClientCompanyCatSpec CCCS ON A.[ClientCompanyCatSpecID] = CCCS.[ClientCompanyCatSpecID]
					Inner Join ClientCoAdjusterSpec CCAS ON A.[AdjusterSpecID] = CCAS.[ClientCoAdjusterSpecID]
				WHERE S.[IsDeleted] = 0
				AND A.[IsDeleted] = 0
				AND A.[AssignmentTypeID] = @AssignmentTypeID
				AND A.[CLIENTNUM] Not Like '%TEST%'
				AND (A.[Reassigned] Is Null Or A.[Reassigned] = 0)
				AND CCCS.[ClientCompanyID] = @ClientCompanyID
				AND CCCS.[CatID] = @CatID
				AND CCAS.[ClientCompanyID] = @ClientCompanyID
				AND CCAS.[UsersID] = @USERSID
				Group By S.[Status], S.[StatusAlias], S.[Description], S.[StatusID]
			) As RetStat
				Left Join Assignments A2 ON A2.[StatusID] = RetStat.[StatusID]
				Left Join ClientCompanyCatSpec CCCS2 ON A2.[ClientCompanyCatSpecID] = CCCS2.[ClientCompanyCatSpecID]
				Left Join ClientCoAdjusterSpec CCAS2 ON A2.[AdjusterSpecID] = CCAS2.[ClientCoAdjusterSpecID]
				Left join PackageItem PKGI ON PKGI.[AssignmentsID] = A2.[AssignmentsID]
			WHERE A2.[IsDeleted] = 0
			AND A2.[AssignmentTypeID] = @AssignmentTypeID 
			AND A2.[CLIENTNUM] Not Like '%TEST%'
			AND (A2.[Reassigned] Is Null Or A2.[Reassigned] = 0)
			AND CCCS2.[ClientCompanyID] = @ClientCompanyID
			AND CCCS2.[CatID] = @CatID
			AND CCAS2.[ClientCompanyID] = @ClientCompanyID
			AND CCAS2.[UsersID] = @USERSID
			Group By 	RetStat.[CountOfStatusItem],
					RetStat.[StatusID],
					RetStat.[Status],
					RetStat.[StatusAlias],
					RetStat.[Description]
		) As RetStatTotal
		Group By 	RetStatTotal.[CountOfStatusItem], 
				RetStatTotal.[Status], 
				RetStatTotal.[StatusAlias], 
				RetStatTotal.[Description], 
				RetStatTotal.[StatusID],
				RetStatTotal.[TotalDocs],
				RetStatTotal.[TotalIsDeleted],
				RetStatTotal.[TotalAttachments],
				RetStatTotal.[TotalCoApprove],
				RetStatTotal.[TotalClientCoReject],
				RetStatTotal.[TotalClientCoDelete],
				RetStatTotal.[TotalClientCoApprove],
				RetStatTotal.[TotalSendMe],
				RetStatTotal.[TotalSent]
		Order by RetStatTotal.[CountOfStatusItem] DESC
	END
END
ELSE
BEGIN
	IF @USERSID is Null
	BEGIN
		SELECT 
			COUNT(A.[StatusID]) As CountOfStatusItem,
			S.[StatusID],
			S.[Status],
			S.[StatusAlias],
			S.[Description]
		FROM Status S
			Inner Join BillAssignment A ON S.[StatusID] = A.[StatusID]
			Inner Join ClientCompanyCatSpec CCCS ON A.[ClientCompanyCatSpecID] = CCCS.[ClientCompanyCatSpecID]
		WHERE S.[IsDeleted] = 0
		AND A.[IsDeleted] = 0
		AND A.[AssignmentTypeID] = @AssignmentTypeID
		AND A.[CLIENTNUM] Not Like '%TEST%'
		AND CCCS.[ClientCompanyID] = @ClientCompanyID
		AND CCCS.[CatID] = @CatID
		Group By S.[Status], S.[StatusAlias], S.[Description], S.[StatusID]
		Order By CountOfStatusItem DESC
	END
	ELSE
	BEGIN
		SELECT 
			COUNT(A.[StatusID]) As CountOfStatusItem,
			S.[StatusID],
			S.[Status],
			S.[StatusAlias],
			S.[Description]
		FROM Status S
			Inner Join BillAssignment A ON S.[StatusID] = A.[StatusID]
			Inner Join ClientCompanyCatSpec CCCS ON A.[ClientCompanyCatSpecID] = CCCS.[ClientCompanyCatSpecID]
			Inner Join ClientCoAdjusterSpec CCAS ON A.[AdjusterSpecID] = CCAS.[ClientCoAdjusterSpecID]
		WHERE S.[IsDeleted] = 0
		AND A.[IsDeleted] = 0
		AND A.[AssignmentTypeID] = @AssignmentTypeID
		AND A.[AdjusterSpecID] IN 	(
							SELECT 	[ClientCoAdjusterSpecID]
							FROM	ClientCoAdjusterSpec
							WHERE	[ClientCompanyID] = @ClientCompanyID
							AND	[UsersID] = @USERSID 
						)
		AND A.[CLIENTNUM] Not Like '%TEST%'
		AND CCCS.[ClientCompanyID] = @ClientCompanyID
		AND CCCS.[CatID] = @CatID
		AND CCAS.[ClientCompanyID] = @ClientCompanyID
		AND CCAS.[UsersID] = @USERSID
		Group By S.[Status], S.[StatusAlias], S.[Description], S.[StatusID]
		Order By CountOfStatusItem DESC
	END
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[z_spsGetClaimsTotals]  TO [CFUSER]
GO

