SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[insAdjusterUsersSoftwareHistory]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[insAdjusterUsersSoftwareHistory]
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
CREATE TRIGGER insAdjusterUsersSoftwareHistory
ON dbo.AdjusterUsersSoftware
INSTEAD OF INSERT
AS

INSERT INTO AdjusterUsersSoftware (
					[UsersID] ,
					[VersionInfo],
					[LicenseDaysLeft] ,
					[ResetLicense],
					[IBPrefix] ,
					[ResetIBPrefix] ,
					[SingleFileSendAuthority] ,
					[DateLastUpdated] ,
					[UpdateByUserID] 
				)
	SELECT
		INS.[UsersID] ,
		INS.[VersionInfo],
		INS.[LicenseDaysLeft],
		INS.[ResetLicense],
		dbo.VerifyNotDupIBPrefix(INS.[IBPrefix], INS.[UsersID]) AS [IBPrefix],
		INS.[ResetIBPrefix],
		INS.[SingleFileSendAuthority],	
		INS.[DateLastUpdated],
		INS.[UpdateByUserID]
	FROM  INSERTED INS

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

