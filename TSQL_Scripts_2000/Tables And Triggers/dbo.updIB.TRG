SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[updIB]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[updIB]
GO

CREATE TRIGGER updIB
ON dbo.IB
AFTER UPDATE
AS
--1 After Updating an IB Insert Applicable record into Batches
--This is the Fee Bill that will be sent to Billings
--Only insert into batches if the ID of IB is > 0
--And the Count of inserted is 1.  This Trigger only handles single record Updates on the IB table Not Multi Updates
--This means the record has successfully been Synched with Client
IF (SELECT COUNT(INS.[ID]) FROM INSERTED INS) = 1
BEGIN
	IF (SELECT INS.[ID] FROM INSERTED INS) > 0
	BEGIN
		DECLARE @CurDate 	DateTime
		DECLARE @MYDATE		DateTime
		DECLARE @MYIBNUM 	VarChar(50)
		DECLARE @VBCRLF		VarChar(2)
		SET @CurDate = GetDate()
		SET @VBCRLF = dbo.GetVBCRLF()
		SET @MYDATE = @CurDate
		SET @MYDATE = dbo.CleanFromOrToDateString(@CurDate, @MYDATE, 1)
		SET @MYIBNUM = ( 
				SELECT
					(CASE	WHEN 	INS.[IB14a_sSupplement] > 0 And INS.[IB14b_sRebilled] > 0
						THEN	INS.[IB02_sIBNumber] + 'S' + cast(INS.[IB14a_sSupplement] As varchar(4)) + 'R' + cast(INS.[IB14b_sRebilled] As varchar(4))
						ELSE	(CASE	WHEN 	INS.[IB14a_sSupplement] > 0
								THEN	INS.[IB02_sIBNumber] + 'S' + cast(INS.[IB14a_sSupplement] As varchar(4))
								ELSE	(CASE	WHEN	INS.[IB14b_sRebilled] > 0
										THEN	INS.[IB02_sIBNumber] + 'R' + cast(INS.[IB14b_sRebilled] As varchar(4))
										ELSE	INS.[IB02_sIBNumber]
										END
									)
								END
							)
						END
					) AS [MYIBNUM]
				FROM INSERTED INS
				)		
		INSERT INTO Batches (
			
			[AssignmentsID],		-- [int] NOT NULL ,
			[ClientCompanyCatSpecID],	-- [int] NOT NULL ,
			[ssn],				-- [numeric](9, 0) NULL ,
			[ibnumber],			-- [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
			[date],				-- [datetime] NULL ,
			[EnteredDate],			-- [datetime] NULL ,
			[adj_name],			-- [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
			[adjuster_n],			-- [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
			[claimnumber],			-- [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
			[insuredname],			-- [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
			[loss_loc],			-- [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
			[losscity],			-- [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
			[lossstate],			-- [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
			[dateofloss],			-- [datetime] NULL ,
			[dateclosed],			-- [datetime] NULL ,
			[grossloss],			-- [decimal](20, 5) NULL ,
			[totalservice],			-- [decimal](20, 5) NULL ,
			[administrative],			-- [decimal](20, 5) NULL ,
			[misccharge],			-- [decimal](20, 5) NULL ,
			[taxestotal],			-- [decimal](20, 5) NULL ,
			[totalfee],			-- [decimal](20, 5) NULL ,
			[catsite],			-- [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
			[Void],				-- [bit] NOT NULL ,
			[billingdup],			-- [bit] NULL ,
			[ecupdated],			-- [bit] NULL ,
			[copied],			-- [int] NULL ,
			[duplicate],			-- [bit] NULL ,
			[Comments],			-- [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
			[Reassigned],			-- [int] NULL ,
			[DateLastUpdated],		-- [datetime] NOT NULL ,
			[UpdateByUserID]		-- [int] NOT NULL 
			)
		SELECT
			INS.[AssignmentsID] 				As [AssignmentsID],		--  [int] NOT NULL ,
			(
			SELECT 	[ClientCompanyCatSpecID]
			FROM	Assignments
			WHERE	[AssignmentsID] = INS.[AssignmentsID]
			) 						As [ClientCompanyCatSpecID],	--  [int] NOT NULL ,
			INS.[IB00_lssn]					As [ssn],			-- [numeric](9, 0) NULL ,
			@MYIBNUM					As [ibnumber],			-- [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
			@MYDATE						As [date],			-- [datetime] NULL ,
			Null						As [EnteredDate],		-- [datetime] NULL ,
			''						As [adj_name],			-- [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
			INS.[IB07_sAdjusterName]			As [adjuster_n],		-- [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
			INS.[IB09_sSALN]				As [claimnumber],		-- [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
			INS.[IB10_sInsuredName]				As [insuredname],		-- [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
			REPLACE(INS.[IB11_sLossLocation],@VBCRLF,'    ')	As [loss_loc],			-- [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
			(
			SELECT 	[PACITY]
			FROM 	Assignments
			WHERE	[AssignmentsID] = INS.[AssignmentsID]
			)
									As [losscity],			-- [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
			(
			SELECT 	[PASTATE]
			FROM 	Assignments
			WHERE	[AssignmentsID] = INS.[AssignmentsID]
			)
									As [lossstate],			-- [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
			INS.[IB12_dtDateOfLoss]				As [dateofloss],		-- [datetime] NULL ,
			INS.[IB06_dtDateClosed]				As [dateclosed],		-- [datetime] NULL ,
			INS.[IB13_cGrossLoss]				As [grossloss],			-- [decimal](20, 5) NULL ,
			INS.[IB25_cServiceFeeSubTotal]			As [totalservice],		-- [decimal](20, 5) NULL ,
			0						As [administrative],		-- [decimal](20, 5) NULL ,
			INS.[IB30_cTotalExpenses]			As [misccharge],		-- [decimal](20, 5) NULL ,
			INS.[IB32_cTaxAmount]				As [taxestotal],		-- [decimal](20, 5) NULL ,
			INS.[IB33_cTotalAdjustingFee]			As [totalfee],			-- [decimal](20, 5) NULL ,
			(
		    	SELECT [BillingCode] 
		    	FROM ClientCompanyCat 
		    	WHERE ClientCompanyID = 
		                ( 
		                SELECT   [ClientCompanyID] 
		                FROM     ClientCompanyCatSpec 
		                WHERE    [ClientCompanyCatSpecID] = 	(
									SELECT 	[ClientCompanyCatSpecID]
									FROM	Assignments
									WHERE	[AssignmentsID] = INS.[AssignmentsID]
									) 		
		                )
		    	AND [CATID] = 	(
		                	SELECT   [CatID] 
		                	FROM     ClientCompanyCatSpec 
		                	WHERE    [ClientCompanyCatSpecID] = 	(
										SELECT 	[ClientCompanyCatSpecID]
										FROM	Assignments
										WHERE	[AssignmentsID] = INS.[AssignmentsID]
										) 	
		                	)
		    	) 						As [catsite],			-- [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
			INS.[Void]					As [Void],			-- [bit] NOT NULL ,
			0						As [billingdup],		-- [bit] NULL ,
			0						As [ecupdated],			-- [bit] NULL ,
			0						As [copied],			-- [int] NULL ,
			0						As [duplicate],			-- [bit] NULL ,
			Left(INS.[Comments],100)			As [Comments],			-- [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
			0						As [Reassigned],		-- [int] NULL ,
			GetDate()					As [DateLastUpdated],		-- [datetime] NOT NULL ,
			INS.[UpdateByUserID]				As [UpdateByUserID]		-- [int] NOT NULL 
		FROM 	INSERTED INS
		WHERE 	@MYIBNUM NOT IN	(
						SELECT 	[ibnumber]
						FROM	BATCHES
					)
	END
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

