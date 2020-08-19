SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[insBillBillingCount]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[insBillBillingCount]
GO

CREATE TRIGGER insBillBillingCount
ON dbo.BillBillingCount
AFTER INSERT
AS
IF (SELECT COUNT(INS.[BillBillingCountID]) FROM INSERTED INS) = 1
BEGIN
	DECLARE @sClientCompanyDBName varchar(30)
	SET @sClientCompanyDBName = 	(
					SELECT	[DBName]
					FROM	Company C
					WHERE 	[IsClientOF] = 	(
								SELECT 	[CompanyID]
								FROM	Company
								WHERE 	[DBName] = 'Eberls'
								)
					AND	[CompanyID] =	(
								SELECT 	CCCS.[ClientCompanyID]	
								FROM	INSERTED INS
								Inner Join BillAssignment BA ON INS.[BillAssignmentID] = BA.[BillAssignmentID]
								Inner Join ClientCompanyCatSpec CCCS ON BA.[ClientCompanyCatSpecID] = CCCS.[ClientCompanyCatSpecID]
								)
					)
	IF @sClientCompanyDBName = 'StateFarm' 
	BEGIN
		--Insert the First IB into IBStateFarm For Eberls Company
		INSERT INTO IBStateFarm 
		(
			[BillAssignmentID],
			[BillBillingCountID],
			[lssn],
			[IBNumber],
			[PolicyNo],
			[Insured],
			[LossLoc1],
			[LossLoc2],
			[LossLocCity],
			[LossLocState],
			[LossLocZipcode],
			[LossDate],
			[CloseDate],
			[GrossLoss],
			[Supplement],
			[SupplementExplain],
			[AdditionalLoss],
			[Rebilled],
			[OrigIBIBNumber],
			[OrigIBTotalFee],
			[RebillExplain],
			[MultiClaimBldgUnitNum],
			[ClientCompanyCatSpecID],
			[SeverityCode],
			[ServiceFeeBase],
			[ServiceFeeCovAExterior],
			[ServiceFeeCovAFraming],
			[ServiceFeeCovAInterior],
			[ServiceFeeCovB],
			[ServiceFeeALE],
			[OutBuildCount],
			[OutBuildPerItemCharge],
			[ServiceFeeOutBuildings],
			[ServiceFeeSteepCharge],
			[ServiceFeeTwoStory],
			[ServiceFeeMoreThan50Squares],
			[ServiceFeeWoodSlateTileConRoof],
			[ServiceFeeAdditionalDamage],
			[ServiceFeeRopeAndHarness],
			[ServiceFeeMisc],
			[MiscFeesExplain],
			[ServiceFeeTotal],
			[ExpensePagerPhoneExplain],
			[ExpensePagerPhone],
			[ExpenseOtherExplain],
			[ExpenseOther],
			[SumTotalServiceFeeAndExpense],
			[TaxPercent],
			[TaxesTotal],
			[TotalFee],
			[Void],
			[Comments],
			[AdminComments],
			[DateLastUpdated],
			[UpdateByUserID]
		)
		SELECT 
			INS.[BillAssignmentID] As [BillAssignmentID],		--  [int] NOT NULL ,
			INS.[BillBillingCountID] As [BillBillingCountID],	--  [int] NOT NULL ,
			Cast(USERS.[SSN] As Numeric(9,0)) As [lssn],	--  [numeric](9, 0) NULL ,
			BASS.[IBNUM] As [IBNumber],			--  [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL  ,
			BASS.[PolicyNo] As [PolicyNo],			--  [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL  ,
			BASS.[Insured] As [Insured],			--  [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL  ,
			BASS.[LossLoc1] As [LossLoc1],			--  [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL  ,
			BASS.[LossLoc2] As [LossLoc2],			--  [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL  ,
			BASS.[LossLocCity] As [LossLocCity],		--  [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL  ,
			BASS.[LossLocState] As [LossLocState],		--  [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL  ,
			BASS.[LossLocZipcode] As [LossLocZipcode],	--  [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL  ,
			BASS.[LossDate] As [LossDate],			--  [datetime] NULL ,
			--Leave the CloseDate Null on the IB since this is 
			--what will trigger the insert into Batches for Ebill processing
			Null As [CloseDate],				--  [datetime] NULL ,
			0 As [GrossLoss],				--  [money] NOT NULL  ,
			INS.[Supplement] As [Supplement],		--  [int] NOT NULL ,
			'' As [SupplementExplain],			--  [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL  ,
			0 As [AdditionalLoss],				--  [money] NOT NULL  ,
			INS.[Rebill] As [Rebilled],			--  [int] NOT NULL ,
			'' As [OrigIBIBNumber],				--  [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
			0 As [OrigIBTotalFee],				--  [money] NULL ,
			'' As [RebillExplain],				--  [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
			'' As [MultiClaimBldgUnitNum],			--  [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
			BASS.[ClientCompanyCatSpecID] As [ClientCompanyCatSpecID],		--  [int] NOT NULL ,
			1 As [SeverityCode],				--  [int] NOT NULL ,
			0 As [ServiceFeeBase],				--  [money] NOT NULL ,
			0 As [ServiceFeeCovAExterior] ,			--  [money] NOT NULL ,
			0 As [ServiceFeeCovAFraming],			--  [money] NOT NULL ,
			0 As [ServiceFeeCovAInterior],			--  [money] NOT NULL ,
			0 As [ServiceFeeCovB],				--  [money] NOT NULL ,
			0 As [ServiceFeeALE],				--  [money] NOT NULL ,
			0 As [OutBuildCount],				--  [int] NOT NULL ,
			0 As [OutBuildPerItemCharge] ,			--  [money] NOT NULL ,
			0 As [ServiceFeeOutBuildings],			--  [money] NOT NULL ,
			0 As [ServiceFeeSteepCharge],			--  [money] NOT NULL ,
			0 As [ServiceFeeTwoStory] ,			--  [money] NOT NULL ,
			0 As [ServiceFeeMoreThan50Squares],		--  [money] NOT NULL ,
			0 As [ServiceFeeWoodSlateTileConRoof],		--  [money] NOT NULL ,
			0 As [ServiceFeeAdditionalDamage],		--  [money] NOT NULL ,	
			0 As [ServiceFeeRopeAndHarness],		--  [money] NOT NULL ,
			0 As [ServiceFeeMisc],				--  [money] NOT NULL ,
			'' As [MiscFeesExplain] ,			--  [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
			0 As [ServiceFeeTotal] ,			--  [money] NOT NULL ,
			'' As [ExpensePagerPhoneExplain],		--  [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
			0 As [ExpensePagerPhone] ,			--  [money] NOT NULL ,
			'' As [ExpenseOtherExplain],			--  [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
			0 As [ExpenseOther] ,				--  [money] NOT NULL ,
			0 As [SumTotalServiceFeeAndExpense] ,	--  [money] NOT NULL ,
			8.250 As [TaxPercent],				--  [numeric](8, 3) NOT NULL ,
			0 As [TaxesTotal],				--  [money] NOT NULL ,
			0 As [TotalFee],				--  [money] NOT NULL ,
			0 As [Void],					--  [bit] NOT NULL ,
			'' As [Comments],				--  [varchar] (150) COLLATE SQL_Latin1_General_CP1_CI_AS  NULL ,
			'' As [AdminComments],				--  [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
			INS.[DateLastUpdated] As [DateLastUpdated],	--  [datetime] NOT NULL ,
			INS.[UpdateByUserID] As [UpdateByUserID]	--  [int] NOT NULL 
		FROM INSERTED INS 
			Inner Join BillAssignment BASS On BASS.[BillAssignmentID] = INS.[BillAssignmentID]
			Inner Join AssignmentType AssType On AssType.[AssignmentTypeID] = BASS.[AssignmentTypeID]
			Inner Join ClientCompanyCatSpec CCCS On CCCS.[ClientCompanyCatSpecID] = BASS.[ClientCompanyCatSpecID]
			Inner Join ClientCoAdjusterSpec CCAS On CCAS.[ClientCoAdjusterSpecID] = BASS.[AdjusterSpecID]
			Inner Join Users ON Users.[UsersID] = CCAS.[UsersID]
	END
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

