SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spsGetPackageItems]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spsGetPackageItems]
GO

/******************************************************************
*
*    	Description:	Get Package Items
*			
*	
*	Author: 	Brad Skidmore
*	Date: 		4/14/2005
*
******************************************************************/
CREATE PROCEDURE [dbo].[z_spsGetPackageItems]
	@UID			int,
	@PackageID		Int=null,
	@bGetHistory		bit=0
AS
BEGIN
	IF @bGetHistory = 0
	BEGIN
		SELECT 
			[PackageItemID],
			[PackageID],
			[AssignmentsID],
			[ID],
			[IDPackage],
			[IDAssignments],
			[ReportFormat],
			[RTAttachmentsID],
			[IDRTAttachments],
			[Number],
			[AttachmentName],
			[SortOrder],
			[Name],
			[Description],
			[IsCoApprove],
			[CoApproveDate],
			[CoApproveDesc],
			[IsClientCoReject],
			[ClientCoRejectDate],
			[ClientCoRejectDesc],
			[IsClientCoDelete],
			[ClientCoDeleteDate],
			[ClientCoDeleteDesc],
			[IsClientCoApprove],
			[ClientCoApproveDate],
			[ClientCoApproveDesc],
			[PackageItemGUID],
			[SendMe],
			[SentDate],
			[IsDeleted],
			[DownLoadMe],
			[UpLoadMe],
			[AdminComments],
			dbo.FormatDate([DateLastUpdated],'MM/dd/yyy HHHH:nn') As [DateLastUpdated],
			[UpdateByUserID],
			(
			SELECT Top 1 [USERNAME]
			FROM	Users
			WHERE	[UsersID] = PKI.[UpdateByUserID]
			) As UpdateByUserName
		FROM 	PackageItem PKI
		WHERE 	[PackageID] = @PackageID
		ORDER BY [SortOrder]
	END
	IF @bGetHistory = 1
	BEGIN
		SELECT 
			[PackageItemHistoryID],
			[PackageItemID],
			[PackageID],
			[AssignmentsID],
			[ID],
			[IDPackage],
			[IDAssignments],
			[ReportFormat],
			[RTAttachmentsID],
			[IDRTAttachments],
			[Number],
			[AttachmentName],
			[SortOrder],
			[Name],
			[Description],
			[IsCoApprove],
			[CoApproveDate],
			[CoApproveDesc],
			[IsClientCoReject],
			[ClientCoRejectDate],
			[ClientCoRejectDesc],
			[IsClientCoDelete],
			[ClientCoDeleteDate],
			[ClientCoDeleteDesc],
			[IsClientCoApprove],
			[ClientCoApproveDate],
			[ClientCoApproveDesc],
			[PackageItemGUID],
			[SendMe],
			[SentDate],
			[IsDeleted],
			[DownLoadMe],
			[UpLoadMe],
			[AdminComments],
			dbo.FormatDate([DateLastUpdated],'MM/dd/yyy HHHH:nn') As [DateLastUpdated],
			[UpdateByUserID],
			(
			SELECT Top 1 [USERNAME]
			FROM	Users
			WHERE	[UsersID] = PKIH.[UpdateByUserID]
			) As UpdateByUserName
		FROM 	PackageItemHistory PKIH
		WHERE 	[PackageID] = @PackageID
		ORDER BY [DateLastUpdated] Desc
	END
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[z_spsGetPackageItems]  TO [CFUSER]
GO

