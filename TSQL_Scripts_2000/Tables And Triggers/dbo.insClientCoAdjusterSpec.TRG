SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[insClientCoAdjusterSpec]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[insClientCoAdjusterSpec]
GO

CREATE TRIGGER insClientCoAdjusterSpec
ON dbo.ClientCoAdjusterSpec 
AFTER INSERT
AS
--After Adding ACID (Adjuster Client Company identification) Update the Inactive date of Duplicate ACID when applicable
UPDATE ClientCoAdjusterSpec SET
			[InactiveDate] = (CASE 	WHEN CCS.[ClientCoAdjusterSpecID] <> INS.[ClientCoAdjusterSpecID] 
					  	THEN INS.[ActiveDate]
						ELSE CCS.[InactiveDate]
						END
					),
			[DateLastUpdated] = INS.[DateLastUpdated],
			[UpdateByUserID] = INS.[UpdateByUserID]
FROM 	ClientCoAdjusterSpec CCS INNER JOIN INSERTED INS ON CCS.ClientCompanyID = INS.ClientCompanyID 
			AND CCS.ACID = INS.ACID
WHERE CCS.InactiveDate Is Null

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

