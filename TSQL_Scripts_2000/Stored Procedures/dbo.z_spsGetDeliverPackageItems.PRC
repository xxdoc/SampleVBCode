SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spsGetDeliverPackageItems]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spsGetDeliverPackageItems]
GO

/******************************************************************
*
*    	Description:	Get Items to be Delivered via Email, FTP, Or B2B
*			
*	
*	Author: 	Brad Skidmore
*	Date: 		7/13/2005
******************************************************************/
CREATE PROCEDURE [dbo].[z_spsGetDeliverPackageItems]
	@AssignmentsID		int,
	@bVerifyIntegrity 	bit=0,
	@OrderByPackageItemID	bit=0 --V2ECcarFarmers.clsLossXML01 Needs to Sort by PackageItemID
AS	
	IF @bVerifyIntegrity = 0
	BEGIN
		IF @OrderByPackageItemID = 1
		BEGIN
			--Need to get the Package itmes for this Assingnment
			--V2ECcarFarmers.clsLossXML01 Needs to Sort by PackageItemID
			SELECT
				PI.[PackageItemID], 
				PI.[PackageID],
				PI.[AssignmentsID], 
				PI.[ID], 
				PI.[IDPackage], 
				PI.[IDAssignments], 
				PI.[ReportFormat], 
				PI.[RTAttachmentsID], 
				PI.[IDRTAttachments], 
				PI.[Number], 
				PI.[AttachmentName], 
				PI.[SortOrder], 
				PI.[Name], 
				PI.[Description], 
				PI.[IsCoApprove], 
				PI.[CoApproveDate], 
				PI.[CoApproveDesc], 
				PI.[IsClientCoReject], 
				PI.[ClientCoRejectDate], 
				PI.[ClientCoRejectDesc], 
				PI.[IsClientCoDelete], 
				PI.[ClientCoDeleteDate], 
				PI.[ClientCoDeleteDesc], 
				PI.[IsClientCoApprove], 
				PI.[ClientCoApproveDate], 
				PI.[ClientCoApproveDesc], 
				PI.[PackageItemGUID], 
				PI.[SendMe], 
				PI.[SentDate], 
				PI.[IsDeleted], 
				PI.[DownLoadMe], 
				PI.[UpLoadMe], 
				PI.[AdminComments], 
				PI.[DateLastUpdated], 
				PI.[UpdateByUserID],
				P.[AdminComments] As [PackageAdminComments],
				-- the CCCat.[EnableSingleFile] must be true before this SP will even be processed
				--Both the Company [EnableSingleFile] and Client Co Cat [EnableSingleFile] flag 
				-- Must be true before any of the below destinations will be processed.
				-- If the SingleFileEmail to CC and Bcc are '' then they will not be processed
				-- even though the [EnableSingleFile] is set to true.
				CCCat.[EnableSingleFile],  
				CCCat.[SingleFileEmail],
				CCCat.[SingleFileEmailCC],
				CCCat.[SingleFileEmailBCC],
				CCCat.[EnableSinglePDF],
				CCCat.[SinglePDFEmail],
				CCCat.[SinglePDFEmailCC],
				CCCat.[SinglePDFEmailBCC],
				CCCat.[EmailDocsOnly],
				CCCat.[EmailDocsOnlyCC],
				CCCat.[EmailDocsOnlyBCC],
				CCCat.[EmailPhotosOnly],
				CCCat.[EmailPhotosOnlyCC],
				CCCat.[EmailPhotosOnlyBCC],
				CCCat.[FTPSingleFileUrl],
				CCCat.[FTPSingleFileUserName],
				CCCat.[FTPSingleFilePassword],
				CCCat.[HttpPostSingleFileUrl],
				CCCat.[HttpPostSingleFileUserName],
				CCCat.[HttpPostSingleFilePassword],
				CCCat.[B2BDir],
				A.[MAStreet], 
				A.[MACity], 
				A.[MAState], 
				A.[MAZIP], 
				A.[PAStreet], 
				A.[PACity], 
				A.[PAState], 
				A.[PAZIP], 
				A.[LRFormat], 
				A.[LossReport],
				A.[CLIENTNUM],
				A.[IBNUM],
				A.[Insured],
				CAT.[Name] As [CatName], 
				C.[Name] As [CoName],
				C.[Code] As [CoCode],
				CCAS.[ACID] As [ACIDDisplay],
				U.[UserName] As [AdjUserName]
			FROM PackageItem PI 
				INNER JOIN Assignments A ON A.[AssignmentsID] = PI.[AssignmentsID]
				INNER JOIN ClientCoAdjusterSpec CCAS ON CCAS.[ClientCoAdjusterSpecID] = A.[AdjusterSpecIDDisplay]
				INNER JOIN USERS U ON U.[UsersID] =  CCAS.[UsersID]
				INNER JOIN ClientCompanyCatSpec CCCS ON CCCS.[ClientCompanyCatSpecID] = A.[ClientCompanyCatSpecID] 
				INNER JOIN CAT ON CAT.[CATID] = CCCS.[CATID]  
				INNER JOIN Company C ON C.[CompanyID] = CCCS.[ClientCompanyID]
				INNER JOIN ClientCompanyCat CCCat ON CCCat.[ClientCompanyID] = CCCS.[ClientCompanyID] And CCCat.[CATID] = CCCS.[CATID]
				INNER JOIN Package P ON P.[PackageID] = PI.[PackageID] 
			WHERE    PI.[AssignmentsID] = @AssignmentsID
			ORDER BY PI.[PackageItemID]
		END
		ELSE
		BEGIN
			--Need to get the Package itmes for this Assingnment
			SELECT
				PI.[PackageItemID], 
				PI.[PackageID],
				PI.[AssignmentsID], 
				PI.[ID], 
				PI.[IDPackage], 
				PI.[IDAssignments], 
				PI.[ReportFormat], 
				PI.[RTAttachmentsID], 
				PI.[IDRTAttachments], 
				PI.[Number], 
				PI.[AttachmentName], 
				PI.[SortOrder], 
				PI.[Name], 
				PI.[Description], 
				PI.[IsCoApprove], 
				PI.[CoApproveDate], 
				PI.[CoApproveDesc], 
				PI.[IsClientCoReject], 
				PI.[ClientCoRejectDate], 
				PI.[ClientCoRejectDesc], 
				PI.[IsClientCoDelete], 
				PI.[ClientCoDeleteDate], 
				PI.[ClientCoDeleteDesc], 
				PI.[IsClientCoApprove], 
				PI.[ClientCoApproveDate], 
				PI.[ClientCoApproveDesc], 
				PI.[PackageItemGUID], 
				PI.[SendMe], 
				PI.[SentDate], 
				PI.[IsDeleted], 
				PI.[DownLoadMe], 
				PI.[UpLoadMe], 
				PI.[AdminComments], 
				PI.[DateLastUpdated], 
				PI.[UpdateByUserID],
				P.[AdminComments] As [PackageAdminComments],
				-- the CCCat.[EnableSingleFile] must be true before this SP will even be processed
				--Both the Company [EnableSingleFile] and Client Co Cat [EnableSingleFile] flag 
				-- Must be true before any of the below destinations will be processed.
				-- If the SingleFileEmail to CC and Bcc are '' then they will not be processed
				-- even though the [EnableSingleFile] is set to true.
				CCCat.[EnableSingleFile],  
				CCCat.[SingleFileEmail],
				CCCat.[SingleFileEmailCC],
				CCCat.[SingleFileEmailBCC],
				CCCat.[EnableSinglePDF],
				CCCat.[SinglePDFEmail],
				CCCat.[SinglePDFEmailCC],
				CCCat.[SinglePDFEmailBCC],
				CCCat.[EmailDocsOnly],
				CCCat.[EmailDocsOnlyCC],
				CCCat.[EmailDocsOnlyBCC],
				CCCat.[EmailPhotosOnly],
				CCCat.[EmailPhotosOnlyCC],
				CCCat.[EmailPhotosOnlyBCC],
				CCCat.[FTPSingleFileUrl],
				CCCat.[FTPSingleFileUserName],
				CCCat.[FTPSingleFilePassword],
				CCCat.[HttpPostSingleFileUrl],
				CCCat.[HttpPostSingleFileUserName],
				CCCat.[HttpPostSingleFilePassword],
				CCCat.[B2BDir],
				A.[MAStreet], 
				A.[MACity], 
				A.[MAState], 
				A.[MAZIP], 
				A.[PAStreet], 
				A.[PACity], 
				A.[PAState], 
				A.[PAZIP], 
				A.[LRFormat], 
				A.[LossReport],
				A.[CLIENTNUM],
				A.[IBNUM],
				A.[Insured],
				CAT.[Name] As [CatName], 
				C.[Name] As [CoName],
				C.[Code] As [CoCode],
				CCAS.[ACID] As [ACIDDisplay],
				U.[UserName] As [AdjUserName]
			FROM PackageItem PI 
				INNER JOIN Assignments A ON A.[AssignmentsID] = PI.[AssignmentsID]
				INNER JOIN ClientCoAdjusterSpec CCAS ON CCAS.[ClientCoAdjusterSpecID] = A.[AdjusterSpecIDDisplay]
				INNER JOIN USERS U ON U.[UsersID] =  CCAS.[UsersID]
				INNER JOIN ClientCompanyCatSpec CCCS ON CCCS.[ClientCompanyCatSpecID] = A.[ClientCompanyCatSpecID] 
				INNER JOIN CAT ON CAT.[CATID] = CCCS.[CATID]  
				INNER JOIN Company C ON C.[CompanyID] = CCCS.[ClientCompanyID]
				INNER JOIN ClientCompanyCat CCCat ON CCCat.[ClientCompanyID] = CCCS.[ClientCompanyID] And CCCat.[CATID] = CCCS.[CATID]
				INNER JOIN Package P ON P.[PackageID] = PI.[PackageID] 
			WHERE    PI.[AssignmentsID] = @AssignmentsID
			ORDER BY PI.[SortOrder]
		END 
	END
	ELSE
	BEGIN
		--Need to get the Package itmes for this Assingnment
		SELECT
			PI.[PackageItemID], 
			PI.[PackageID],
			PI.[AssignmentsID], 
			PI.[ID], 
			PI.[IDPackage], 
			PI.[IDAssignments], 
			PI.[ReportFormat], 
			PI.[RTAttachmentsID], 
			PI.[IDRTAttachments], 
			PI.[Number], 
			PI.[AttachmentName], 
			PI.[SortOrder], 
			PI.[Name], 
			PI.[Description], 
			PI.[IsCoApprove], 
			PI.[CoApproveDate], 
			PI.[CoApproveDesc], 
			PI.[IsClientCoReject], 
			PI.[ClientCoRejectDate], 
			PI.[ClientCoRejectDesc], 
			PI.[IsClientCoDelete], 
			PI.[ClientCoDeleteDate], 
			PI.[ClientCoDeleteDesc], 
			PI.[IsClientCoApprove], 
			PI.[ClientCoApproveDate], 
			PI.[ClientCoApproveDesc], 
			PI.[PackageItemGUID], 
			PI.[SendMe], 
			PI.[SentDate], 
			PI.[IsDeleted], 
			PI.[DownLoadMe], 
			PI.[UpLoadMe], 
			PI.[AdminComments], 
			PI.[DateLastUpdated], 
			PI.[UpdateByUserID]
		FROM PackageItem PI 
		WHERE   PI.[AssignmentsID] = @AssignmentsID
		AND	PI.[IsDeleted] = 0
		ORDER BY PI.[SortOrder] 
	END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[z_spsGetDeliverPackageItems]  TO [CFUSER]
GO

