SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[updClientCompanyCatHistory]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[updClientCompanyCatHistory]
GO

CREATE TRIGGER updClientCompanyCatHistory
ON dbo.ClientCompanyCat
FOR UPDATE
AS
INSERT INTO ClientCompanyCatHistory
	SELECT del.* 
	FROM DELETED del

--When the INactive dates ONLY !!! Change for the Client Company Cat....
--need to apply the changes as appropriate to the ClientCompanyCatSPecs associated with the ClientCompanyCat table
--InactiveDate
UPDATE ClientCompanyCatSpec set 
	ClientCompanyCatSpec.inactivedate = 	(	CASE 	WHEN cccs.inactivedate is null 
							THEN INS.inactivedate 
							ELSE (	CASE 	WHEN cccs.inactivedate > INS.inactivedate
									THEN INS.inactivedate
									ELSE cccs.inactivedate
									END
								)
							END
						),
	ClientCompanyCatSpec.DateLastUpdated = INS.DateLastUpdated,
	ClientCompanyCatSpec.UpdateByUserID = INS.UpdateByUserID
FROM  INSERTED INS 
	INNER JOIN DELETED DEL ON 
			INS.CatID = DEL.CatID 
			AND INS.ClientCompanyID = DEL.ClientCompanyID 
			AND (INS.inactivedate <> DEL.inactivedate  Or (DEL.inactivedate Is null And INS.inactivedate Is Not Null) Or (DEL.inactivedate Is Not null And INS.inactivedate Is Null)) 
	INNER JOIN ClientCompanyCatSpec cccs ON 
			cccs.catid = INS.catid
			AND cccs.ClientCompanyID = INS.ClientCompanyID

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

