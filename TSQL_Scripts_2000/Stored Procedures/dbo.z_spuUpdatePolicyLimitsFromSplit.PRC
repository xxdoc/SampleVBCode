SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spuUpdatePolicyLimitsFromSplit]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spuUpdatePolicyLimitsFromSplit]
GO

/******************************************************************
*
*    	Description:	Update Policy Limits Table From Split Delim String 
*	
*	Author: 	Brad Skidmore
*	Date: 		4/21/2004
*	
******************************************************************/

CREATE PROCEDURE [dbo].[z_spuUpdatePolicyLimitsFromSplit]
	@PolicyLimitsList		text=Null,--Policy Limits Delimited String to be parsed by SplitPolicyLimits UDF
	@AssignmentsID			int=null,
	@DateLastUpdated 		datetime,
	@UpdateByUserID 		int
AS
DECLARE @CurDate 		DateTime
SET @CurDate = GetDate()
--See if the DateLastupdated is null
IF IsDate(@DateLastUpdated) =0
BEGIN
	SET @DateLastUpdated = GetDate()
END
IF DATALENGTH(@PolicyLimitsList) > 0
BEGIN
	--1. Add New Records
	INSERT INTO PolicyLimits	(
						[AssignmentsID] ,
						[ID] ,
						[IDAssignments] ,
						[ClassTypeID] ,
						[LimitAmount] ,
						[RCSaidProp] ,
						[Reserves] ,
						[IsDeleted] ,
						[DownLoadMe] ,
						[UpLoadMe] ,
						[AdminComments] ,
						[DateLastUpdated] ,
						[UpdateByUserID],
						[FarmersUnitRowID] 
					)
	SELECT				
						@AssignmentsID As [AssignmentsID] ,
						PL.[ID] AS [ID] ,
						PL.IDAssignments As [IDAssignments] ,
						PL.ClassTypeID As [ClassTypeID] ,
						PL.LimitAmount As [LimitAmount] ,
						PL.RCSaidProp As [RCSaidProp] ,
						PL.Reserves As [Reserves] ,
						0 As [IsDeleted] ,
						1 As [DownLoadMe] ,
						0 As [UpLoadMe] ,
						PL.AdminComments As [AdminComments] ,
						@DateLastUpdated As [DateLastUpdated] ,
						@UpdateByUserID As [UpdateByUserID],
						PL.FarmersUnitRowID As [FarmersUnitRowID] 
					
	FROM dbo.SplitPolicyLimits(@PolicyLimitsList, default, default) As PL
	WHERE	PL.PolicyLimitsID = 0 

	--2. Update Existing Records
	UPDATE PolicyLimits 	SET	
					[AssignmentsID] 	=PL.AssignmentsID,
					[ID]			=PL.[ID],
					[IDAssignments] 	=PL.IDAssignments,
					[ClassTypeID] 		=PL.ClassTypeID,
					[LimitAmount] 		=PL.LimitAmount,
					[RCSaidProp] 		=PL.RCSaidProp,
					[Reserves] 		=PL.Reserves,
					[IsDeleted] 		=PL.IsDeleted,
					[DownLoadMe] 		=PL.DownLoadMe,
					[UpLoadMe] 		=PL.UpLoadMe,
					[AdminComments]		=PL.AdminComments,
					[DateLastUpdated] 	=@DateLastUpdated,
					[UpdateByUserID] 	=@UpdateByUserID,
					[FarmersUnitRowID]	=PL.FarmersUnitRowID

	FROM dbo.SplitPolicyLimits(@PolicyLimitsList, default, default) As PL	
	INNER JOIN PolicyLimits A ON A.PolicyLimitsID = PL.PolicyLimitsID
	AND DATEDIFF(s, A.DateLastUpdated, PL.DateLastUpdated) > 1

	-- Also Need to update the ID to Main ID if it is 0
	-- This will be the ID in EasyClaim Applet (Macth the server)
	UPDATE 	PolicyLimits SET [ID] = PolicyLimitsID
	WHERE 	[ID] = 0
	AND 	[AssignmentsID] = @AssignmentsID

	UPDATE 	PolicyLimits SET IDAssignments = AssignmentsID
	WHERE 	IDAssignments = 0
	AND 	[AssignmentsID] = @AssignmentsID
	
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[z_spuUpdatePolicyLimitsFromSplit]  TO [CFUSER]
GO

