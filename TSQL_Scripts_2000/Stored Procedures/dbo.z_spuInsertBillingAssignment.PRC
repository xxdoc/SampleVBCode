SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spuInsertBillingAssignment]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spuInsertBillingAssignment]
GO

/******************************************************************
*
*    	Description:	Insert Billing Assignments 
*	
*	Author: 	Brad Skidmore
*	Date: 		3/15/2005
*	
******************************************************************/

CREATE PROCEDURE [dbo].[z_spuInsertBillingAssignment]
	@AssignmentTypeID		int,
	@ClientCompanyCatSpecID		int,
	@AdjusterSpecID 		int,
	@IBNUM 				varchar(20),
	@CLIENTNUM 			varchar(20),
	@PolicyNo 			varchar(50),
	@Insured 			varchar(100),
	@LossLoc1 			varchar(50),
	@LossLoc2 			varchar(50),
	@LossLocCity 			varchar(50),
	@LossLocState 			varchar(50),
	@LossLocZipcode 		varchar(10),
	@LossDate 			datetime=null,
	@CloseDate 			datetime=null,
	@IsLocked 			bit=0,
	@IsDeleted 			bit=0,
	@AdminComments 			varchar(1000)=null,
	@MiscDelimSettings		text,
	@DateLastUpdated 		datetime,
	@UpdateByUserID 		int
AS
--Assignments ID
DECLARE @BillAssignmentID 			int

--First check to see if the CLIENTNUM IS already in Assignments
-- Assigned to the same adjuster.  If it is then Just update the Loss Report with this new one.
-- When the Loss Report is updated, the trigger on Assignemts will update it's history table.
SET @BillAssignmentID = 	(	SELECT 	TOP 1 BillAssignmentID 
				FROM 	BillAssignment 
				WHERE  	[CLIENTNUM] = @CLIENTNUM
				AND 	[AdjusterSpecID] IN	(
									SELECT 	[ClientCoAdjusterSpecID]
									FROM 	ClientCoAdjusterSpec
									WHERE	[USERSID]	IN	(
													SELECT 	[USERSID]
													FROM	ClientCoAdjusterSpec
													WHERE	[ClientCoAdjusterSpecID] = @AdjusterSpecID
												)
								) 
			)

IF @BillAssignmentID IS NULL
BEGIN
	INSERT INTO BillAssignment 	(
					[AssignmentTypeID],
					[ClientCompanyCatSpecID],
					[AdjusterSpecID],
					[IBNUM],
					[CLIENTNUM],
					[PolicyNo],
					[Insured],
					[LossLoc1],
					[LossLoc2],
					[LossLocCity],
					[LossLocState],
					[LossLocZipcode],
					[LossDate],
					[CloseDate],
					[IsLocked],
					[IsDeleted],
					[AdminComments],
					[MiscDelimSettings],
					[DateLastUpdated],
					[UpdateByUserID]
					)
		SELECT			@AssignmentTypeID As [AssignmentTypeID],
					@ClientCompanyCatSpecID As [ClientCompanyCatSpecID],
					@AdjusterSpecID As [AdjusterSpecID],
					@IBNUM As [IBNUM],
					@CLIENTNUM As [CLIENTNUM],
					@PolicyNo As [PolicyNo],
					@Insured As [Insured],
					@LossLoc1 As [LossLoc1],
					@LossLoc2 As [LossLoc2],
					@LossLocCity As [LossLocCity],
					@LossLocState As [LossLocState],
					@LossLocZipcode As [LossLocZipcode],
					@LossDate As [LossDate],
					@CloseDate As [CloseDate],
					@IsLocked As [IsLocked],
					@IsDeleted As [IsDeleted],
					@AdminComments As [AdminComments],
					@MiscDelimSettings As [MiscDelimSettings],
					@DateLastUpdated As [DateLastUpdated],
					@UpdateByUserID As [UpdateByUserID]	
END
ELSE
BEGIN
	UPDATE BillAssignment SET	[AdminComments] 	= 'Duplicate Insert Aborted.',
					[DateLastUpdated]	= @DateLastUpdated,
					[UpdateByUserID]	= @UpdateByUserID
	WHERE BillAssignmentID		=@BillAssignmentID
	--Also Return this message so that WEBCONTROL is Informed of this action.
	DECLARE @sMess as varchar(50)
	SET @sMess = 'Duplicate Insert Aborted. For ' + @CLIENTNUM
	RAISERROR ( @sMess, 16, 1) 
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[z_spuInsertBillingAssignment]  TO [CFUSER]
GO

