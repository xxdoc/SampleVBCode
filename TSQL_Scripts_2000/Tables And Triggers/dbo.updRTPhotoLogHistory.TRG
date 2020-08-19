SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[updRTPhotoLogHistory]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[updRTPhotoLogHistory]
GO

--Because of the Restriction on Triggers when Ntext fields ...
/*
In a DELETE, INSERT, or UPDATE trigger, SQL Server does not allow text, ntext, or image column references 
in the inserted and deleted tables if the compatibility level is equal to 70. The text, ntext, and image values 
in the inserted and deleted tables cannot be accessed. To retrieve the new value in either an INSERT or UPDATE 
trigger, join the inserted table with the original update table. When the compatibility level is 65 or lower, 
null values are returned for inserted or deleted text, ntext, or image columns that allow null values; zero-length 
strings are returned if the columns are not nullable. 
If the compatibility level is 80 or higher, SQL Server allows the update of text, ntext, or image columns through 
the INSTEAD OF trigger on tables or views.
*/
--Use Instead OF to get around the above restriction
CREATE TRIGGER updRTPhotoLogHistory
ON dbo.RTPhotoLog
INSTEAD OF UPDATE
AS
INSERT INTO RTPhotoLogHistory
	SELECT del.* 
	FROM DELETED del

-- Now that the History table was updated first...
--Allow the original update to process...
Update RTPhotoLog SET
[RTPhotoReportID]=		INS.RTPhotoReportID,
[AssignmentsID]=		INS.AssignmentsID,
[BillingCountID]=		INS.BillingCountID,
[ID]=				INS.ID,
[IDRTPhotoReport]=		INS.IDRTPhotoReport,
[IDAssignments]=		INS.IDAssignments,
[IDBillingCount]=		INS.IDBillingCount,
[PhotoDate]=			INS.PhotoDate,
[SortOrder]=			INS.SortOrder,
[Description]=			INS.Description,
[PhotoName]=			INS.PhotoName,
[Photo]=			INS.Photo,
[DownloadPhoto]=		INS.DownloadPhoto,
[UpLoadPhoto]=			INS.UpLoadPhoto,
[PhotoThumb]=			INS.PhotoThumb,
[DownloadPhotoThumb]=		INS.DownloadPhotoThumb,
[UpLoadPhotoThumb]=		INS.UpLoadPhotoThumb,
[PhotoHighRes]=		INS.PhotoHighRes,
[DownloadPhotoHighRes]=		INS.DownloadPhotoHighRes,
[UploadPhotoHighRes]=		INS.UploadPhotoHighRes,
[IsDeleted]=			INS.IsDeleted,
[DownLoadMe]=			INS.DownLoadMe,
[UpLoadMe]=			INS.UpLoadMe,
[AdminComments]=		INS.AdminComments,
[DateLastUpdated]=		INS.DateLastUpdated,
[UpdateByUserID]=		INS.UpdateByUserID
FROM RTPhotoLog RTPL INNER JOIN INSERTED INS ON RTPL.RTPhotoLogID = INS.RTPhotoLogID

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

