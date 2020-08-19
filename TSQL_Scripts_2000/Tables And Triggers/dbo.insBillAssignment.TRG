SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[insBillAssignment]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[insBillAssignment]
GO

CREATE TRIGGER insBillAssignment
ON dbo.BillAssignment
AFTER INSERT
AS
BEGIN
	--Make the IBNUM
	UPDATE BillAssignment
	SET 
		BillAssignment.[IBNUM] = 'ETS' + Cast(BillAssignment.[BillAssignmentID] As VarChar(20)),
		BillAssignment.[StatusID] = (SELECT [StatusID] FROM Status WHERE [Status] = 'NEW')
	FROM INSERTED INS Inner Join BillAssignment On BillAssignment.[BillAssignmentID] = INS.[BillAssignmentID]

	--Insert the First Billing in BillBillingCount
	INSERT INTO BillBillingCount 
	(
		[BillAssignmentID],
		[Rebill],
		[Supplement],
		[AdminComments],
		[DateLastUpdated],
		[UpdateByUserID]
	)
	SELECT 
		INS.[BillAssignmentID] As [BillAssignmentID],
		0 As [Rebill],
		0 As [Supplement],
		'' As [AdminComments],
		INS.[DateLastUpdated] As [DateLastUpdated],
		INS.[UpdateByUserID] As [UpdateByUserID]
	FROM INSERTED INS Inner Join BillAssignment On BillAssignment.[BillAssignmentID] = INS.[BillAssignmentID]
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

