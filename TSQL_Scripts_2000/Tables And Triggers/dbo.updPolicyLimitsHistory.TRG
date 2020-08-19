SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[updPolicyLimitsHistory]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[updPolicyLimitsHistory]
GO

CREATE TRIGGER updPolicyLimitsHistory
ON dbo.PolicyLimits
INSTEAD OF UPDATE
AS
INSERT INTO PolicyLimitsHistory
	SELECT del.* 
	FROM DELETED del

-- Now that the History table was updated first...
--Allow the original update to process...

Update PolicyLimits SET
	[AssignmentsID]= 	INS.[AssignmentsID],
	[ID]= 			INS.[ID],
	[IDAssignments]= 	INS.[IDAssignments],
	[ClassTypeID]= 		INS.[ClassTypeID],
	[LimitAmount]= 		INS.[LimitAmount],
	[RCSaidProp]= 		INS.[RCSaidProp],
	[Reserves]= 		INS.[Reserves],
				--Account for Adjuster Changing things at same time Farmers Sends Close Transaction
	[IsDeleted]= 		(Case 	
					WHEN (DEL.[FarmersUnitRowID] <> '' And DEL.[AdminComments] = 'Unit Closed by client.' And INS.[AdminComments] <> 'Unit Closed by client.' And INS.[AdminComments] <> 'UNDELETE UNIT') 
					THEN DEL.[IsDeleted] 
					ELSE INS.[IsDeleted] 
					END
				) 	,
				--Account for Adjuster Changing things at same time Farmers Sends Close Transaction
	[DownLoadMe]= 		(Case 	
					WHEN (DEL.[FarmersUnitRowID] <> '' And DEL.[AdminComments] = 'Unit Closed by client.' And INS.[AdminComments] <> 'Unit Closed by client.' And INS.[AdminComments] <> 'UNDELETE UNIT') 
					THEN DEL.[DownLoadMe] 
					ELSE INS.[DownLoadMe] 
					END
				) 	,
	[UpLoadMe]= 		INS.[UpLoadMe],
				--Account for Adjuster Changing things at same time Farmers Sends Close Transaction
	[AdminComments]= 	(Case 	
					WHEN (DEL.[FarmersUnitRowID] <> '' And DEL.[AdminComments] = 'Unit Closed by client.' And INS.[AdminComments] <> 'Unit Closed by client.' And INS.[AdminComments] <> 'UNDELETE UNIT') 
					THEN DEL.[AdminComments] 
					ELSE INS.[AdminComments] 
					END
				) 	,
	[DateLastUpdated]= 	INS.[DateLastUpdated],
	[UpdateByUserID]= 	INS.[UpdateByUserID],
	[FarmersUnitRowID]= 	(Case 	
					WHEN (INS.[FarmersUnitRowID] = '' And DEL.[FarmersUnitRowID] <> '') 
					THEN DEL.[FarmersUnitRowID] 
					ELSE INS.[FarmersUnitRowID] 
					END
				) 	
FROM PolicyLimits PL 
	INNER JOIN INSERTED INS ON PL.PolicyLimitsID = INS.PolicyLimitsID
	INNER JOIN DELETED DEL ON PL.PolicyLimitsID = DEL.PolicyLimitsID

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

