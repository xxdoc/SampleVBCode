SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[updCATHistory]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[updCATHistory]
GO

CREATE TRIGGER updCATHistory
ON dbo.CAT
FOR UPDATE
AS
INSERT INTO CATHistory
	SELECT del.* 
	FROM DELETED del


---When the Active and INactive dates Change need to apply the changes 
--as appropriate to the ClientCompanyCats associated with the CAT table

--InactiveDate
UPDATE ClientCompanyCat set ClientCompanyCat.inactivedate = (	CASE 	WHEN cc.inactivedate is null 
								THEN INS.inactivedate 
								ELSE (	CASE 	WHEN cc.inactivedate > INS.inactivedate
										THEN INS.inactivedate
										ELSE cc.inactivedate
										END
									)
								END
							),
			ClientCompanyCat.DateLastUpdated = INS.DateLastUpdated,
			ClientCompanyCat.UpdateByUserID = INS.UpdateByUserID
FROM  INSERTED INS 
	INNER JOIN DELETED DEL ON INS.CatID = DEL.CatID AND (INS.inactivedate <> DEL.inactivedate  Or (DEL.inactivedate Is null And INS.inactivedate Is Not Null) Or (DEL.inactivedate Is Not null And INS.inactivedate Is Null)) 
	INNER JOIN ClientCompanyCat cc ON cc.catid = INS.catid

--ActiveDate
UPDATE ClientCompanyCat set ClientCompanyCat.activedate = (	CASE 	WHEN cc.activedate is null 
								THEN INS.activedate 
								ELSE(	CASE	WHEN  cc.activedate < INS.activedate
										THEN INS.activedate
										ELSE cc.activedate
										END
									)
								END
							),
			ClientCompanyCat.DateLastUpdated = INS.DateLastUpdated,
			ClientCompanyCat.UpdateByUserID = INS.UpdateByUserID
FROM  INSERTED INS 
	INNER JOIN DELETED DEL ON INS.CatID = DEL.CatID AND (INS.activedate <> DEL.activedate  Or (DEL.activedate Is null And INS.activedate Is Not Null) Or (DEL.activedate Is Not null And INS.activedate Is Null)) 
	INNER JOIN ClientCompanyCat cc ON cc.catid = INS.catid

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

