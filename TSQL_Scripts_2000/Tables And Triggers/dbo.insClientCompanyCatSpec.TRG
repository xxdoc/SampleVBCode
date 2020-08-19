SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[insClientCompanyCatSpec]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[insClientCompanyCatSpec]
GO

CREATE TRIGGER insClientCompanyCatSpec
ON dbo.ClientCompanyCatSpec 
AFTER INSERT
AS
--After Adding Cat code Update the Inactive date of Duplicate Cat code when applicable
UPDATE ClientCompanyCatSpec SET
			[InactiveDate] = (CASE 	WHEN CCS.[ClientCompanyCatSpecID] <> INS.[ClientCompanyCatSpecID] 
					  	THEN INS.[ActiveDate]
						ELSE CCS.[InactiveDate]
						END
					),
			[DateLastUpdated] = INS.[DateLastUpdated],
			[UpdateByUserID] = INS.[UpdateByUserID]
FROM 	ClientCompanyCatSpec CCS INNER JOIN INSERTED INS ON CCS.ClientCompanyID = INS.ClientCompanyID 
			AND CCS.CatCode = INS.CatCode
			AND CCS.AssignmentTypeID = INS.AssignmentTypeID
WHERE CCS.InactiveDate Is Null

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

