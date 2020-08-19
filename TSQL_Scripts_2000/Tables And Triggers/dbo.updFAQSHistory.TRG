SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[updFAQSHistory]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[updFAQSHistory]
GO

CREATE TRIGGER updFAQSHistory
ON dbo.FAQS
--Use "Instead OF UPDATE" to get around TEXT MEMO BLOB Restrictions
INSTEAD OF UPDATE
AS
INSERT INTO FAQSHistory
	SELECT del.* 
	FROM DELETED del
-- Now that the History table was updated first...
--Allow the original update to process...

Update FAQS SET
	[Question] = INS.Question,
	[Answer] = INS.Answer,
	[IsDeleted] = INS.IsDeleted,
	[DateLastUpdated] = INS.DateLastUpdated,
	[UpdateByUserID] =INS.UpdateByUserID
FROM FAQS INNER JOIN INSERTED INS ON FAQS.FAQSID = INS.FAQSID

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

