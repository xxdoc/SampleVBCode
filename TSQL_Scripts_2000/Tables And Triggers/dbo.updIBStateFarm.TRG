SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[updIBStateFarm]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[updIBStateFarm]
GO

CREATE TRIGGER updIBStateFarm
ON dbo.IBStateFarm
AFTER UPDATE
AS
--1 After Updating an IB Insert Applicable record into Batches
--This is the Fee Bill that will be sent to Billings
--Only insert into batches if the CloseDate is not null
--And the Count of inserted is 1.  This Trigger only handles single record Updates on the IB table Not Multi Updates
IF (SELECT COUNT(INS.[IBStateFarmID]) FROM INSERTED INS) = 1
BEGIN
	IF (SELECT INS.[CloseDate] FROM INSERTED INS) Is Not Null
	BEGIN
		DECLARE @CurDate 	DateTime
		DECLARE @MYDATE		DateTime
		DECLARE @MYIBNUM 	VarChar(50)
		SET @CurDate = GetDate()
		SET @MYDATE = @CurDate
		SET @MYDATE = dbo.CleanFromOrToDateString(@CurDate, @MYDATE, 1)
		SET @MYIBNUM = ( 
				SELECT
					(CASE	WHEN 	INS.[Supplement] > 0 And INS.[Rebilled] > 0
						THEN	INS.[IBNumber] + 'S' + cast(INS.[Supplement] As varchar(4)) + 'R' + cast(INS.[Rebilled] As varchar(4))
						ELSE	(CASE	WHEN 	INS.[Supplement] > 0
								THEN	INS.[IBNumber] + 'S' + cast(INS.[Supplement] As varchar(4))
								ELSE	(CASE	WHEN	INS.[Rebilled] > 0
										THEN	INS.[IBNumber] + 'R' + cast(INS.[Rebilled] As varchar(4))
										ELSE	INS.[IBNumber]
										END
									)
								END
							)
						END
					) AS [MYIBNUM]
				FROM INSERTED INS
				)
		
		INSERT INTO Batches (
			[AssignmentsID],		-- [int] NULL ,
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
			[administrative],		-- [decimal](20, 5) NULL ,
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
			[UpdateByUserID],		-- [int] NOT NULL ,
			[BillAssignmentID]		-- [int] NULL 
			)
		SELECT
			-- Since this is a State Farm Ebill Make sure the AssignmentsID is NULL !!!!
			Null 						As [AssignmentsID],		--  [int] NOT NULL ,
			CCCS.[ClientCompanyCatSpecID]			As [ClientCompanyCatSpecID],	--  [int] NOT NULL ,
			INS.[lssn]					As [ssn],			-- [numeric](9, 0) NULL ,
			@MYIBNUM					As [ibnumber],			-- [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
			@MYDATE						As [date],			-- [datetime] NULL ,
			Null						As [EnteredDate],		-- [datetime] NULL ,
			''						As [adj_name],			-- [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
			(	
				Users.[LastName] + ', ' + Users.[FirstName]
			)						As [adjuster_n],		-- [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
			BASS.[CLIENTNUM]				As [claimnumber],		-- [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
			INS.[Insured] 					As [insuredname],		-- [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,


			(
				INS.[LossLoc1] + ' ' + INS.[LossLoc2] + '    ' + INS.[LossLocCity] + ', ' + INS.[LossLocState] + ' ' + INS.[LossLocZipcode] 
			)						As [loss_loc],			-- [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
			INS.[LossLocCity]				As [losscity],			-- [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
			INS.[LossLocState]				As [lossstate],			-- [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
			INS.[LossDate]					As [dateofloss],		-- [datetime] NULL ,
			INS.[CloseDate]					As [dateclosed],		-- [datetime] NULL ,
			INS.[GrossLoss]					As [grossloss],			-- [decimal](20, 5) NULL ,
			INS.[ServiceFeeTotal]				As [totalservice],		-- [decimal](20, 5) NULL ,
			0						As [administrative],		-- [decimal](20, 5) NULL ,
			(
				INS.[ExpensePagerPhone] + INS.[ExpenseOther]
			)						As [misccharge],		-- [decimal](20, 5) NULL ,
			INS.[TaxesTotal]				As [taxestotal],		-- [decimal](20, 5) NULL ,
			INS.[TotalFee]					As [totalfee],			-- [decimal](20, 5) NULL ,
			CCC.[BillingCode]				As [catsite],			-- [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
			INS.[Void]					As [Void],			-- [bit] NOT NULL ,
			0						As [billingdup],		-- [bit] NULL ,
			0						As [ecupdated],			-- [bit] NULL ,
			0						As [copied],			-- [int] NULL ,
			0						As [duplicate],			-- [bit] NULL ,
			Left(INS.[Comments],100)			As [Comments],			-- [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
			0						As [Reassigned],		-- [int] NULL ,
			GetDate()					As [DateLastUpdated],		-- [datetime] NOT NULL ,
			INS.[UpdateByUserID]				As [UpdateByUserID],		-- [int] NOT NULL 
			INS.[BillAssignmentID]			As [BillAssignmentID]	--[int]  NULL
		FROM 	INSERTED INS
			Inner Join BillAssignment BASS On BASS.[BillAssignmentID] = INS.[BillAssignmentID]
			Inner Join AssignmentType AssType On AssType.[AssignmentTypeID] = BASS.[AssignmentTypeID]
			Inner Join ClientCompanyCatSpec CCCS On CCCS.[ClientCompanyCatSpecID] = BASS.[ClientCompanyCatSpecID]
			Inner Join ClientCompanyCat CCC On CCC.[ClientCompanyID] = CCCS.[ClientCompanyID] And CCC.[CATID] = CCCS.[CATID]
			Inner Join ClientCoAdjusterSpec CCAS On CCAS.[ClientCoAdjusterSpecID] = BASS.[AdjusterSpecID]
			Inner Join Users ON Users.[UsersID] = CCAS.[UsersID]
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

