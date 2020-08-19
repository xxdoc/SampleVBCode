SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spuScrubIBFee]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[spuScrubIBFee]
GO

Create Procedure spuScrubIBFee
As

DECLARE @AssignmentsID as int
DECLARE @IBID int

--1. Get any IB assignmentsId that has IBFee Items that need to be
--Populated from the RTIBFee table 
Declare MyIBFeeScrub Cursor for
	SELECT 	RetIB2.AssignmentsID,
		RetIB2.IBID --,
-- 	RetIB2.NUMFeeTypes,
-- 		(	
-- 		SELECT 	Count(RTIBFee.AssignmentsID) 
-- 		FROM 	RTIBFee 
-- 		WHERE RTIBFee.AssignmentsID = RetIB2.AssignmentsID 
-- 		) As NumRTFeeTypes
	FROM
		(
		SELECT 	(SELECT TOP 1 RetIB.IBID) As IBID,
			(SELECT TOP 1 RetIB.AssignmentsID) As AssignmentsID,
			Count(RetIB.IBID) As NUMFeeTypes
		FROM
			(
			SELECT 	IB1.IBID, IB1.AssignmentsID
			FROM 	IB IB1 
			LEFT JOIN IBFee ON IB1.[IBID] = IBFee.[IBID]
			WHERE 	IB1.[DateLastUpdated] >= DateAdd(d,-1, getDate())
			AND 	IB1.[datelastupdated] <= DateAdd(mi,-1, getDate())
			) RetIB
		Group by RetIB.IBID, RetIB.AssignmentsID
		) RetIB2
	WHERE RetIB2.NUMFeeTypes < 	(	
					SELECT 	Count(RTIBFee.AssignmentsID) 
					FROM 	RTIBFee 
					WHERE RTIBFee.AssignmentsID = RetIB2.AssignmentsID 
					) 
Open MyIBFeeScrub

Fetch Next From MyIBFeeScrub into @AssignmentsID, @IBID

WHile @@FEtch_Status = 0
	Begin
		INSERT INTO IBFee
			(
			[AssignmentsID],
			[IBID] ,
			[ID] ,
			[IDAssignments],
			[IDIB] ,
			[FeeScheduleFeeTypesID],
			[NumberOfItems],
			[Amount],
			[Comment],
			[DownLoadMe] ,
			[UpLoadMe] ,
			[AdminComments],
			[DateLastUpdated],
			[UpdateByUserID]
			)
		SELECT 
			[AssignmentsID] ,
			@IBID As [IBID] ,
			Null As [ID],
			[IDAssignments],
			@IBID As [IDIB],
			[FeeScheduleFeeTypesID],
			[NumberOfItems],
			[Amount] ,
			[Comment] ,
			1 As [DownLoadMe],
			[UpLoadMe] ,
			[AdminComments],
			[DateLastUpdated],
			[UpdateByUserID] 
			FROM RTIBFee
			WHERE [AssignmentsID] = @AssignmentsID
			AND  [FeeScheduleFeeTypesID] Not In 	(	
								SELECT 	[FeeScheduleFeeTypesID] 
								FROM 	IBFEE 
								WHERE 	[AssignmentsID] = @AssignmentsID
								AND 	[IBID] = @IBID
								)

		UPDATE IBFee SET 
			[ID] =  [IBFEEID] 
		WHERE [ID] Is Null 
		AND [AssignmentsID] = @AssignmentsID 
						
		Fetch Next From MyIBFeeScrub into @AssignmentsID, @IBID
	End

Close MyIBFeeScrub
Deallocate MyIBFeeScrub

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[spuScrubIBFee]  TO [CFUSER]
GO

