SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spsList_PolicyLimits]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spsList_PolicyLimits]
GO

/******************************************************************
*
*    	Description:	Creates List of Policy Limits From A RecordSet Of Policy Limits
*	Instructions: 	Fill in the Sepecified Areas for When Adding Columns
			Or changing Column Names Where you see
			--*********************************
			--BEGIN %
			--END %
			--*********************************
	Notes:		Default Column Delim is 'Þ' Asc Code 222
			Default Record Delim is '¶' Asc Code 182
*	Author: 	Brad Skidmore
*	Date: 		5/4/2004
*
******************************************************************/

CREATE Procedure [dbo].[z_spsList_PolicyLimits]
	@bHideDeleted		bit=1, 
	@UID			int,
	@iPolicyLimitsID	Int=null,
	@OrderBy		varchar(500)=null,
	@GroupBy		varchar(500)=null,
	@SearchBy		varchar(1000)=null,
	@COLUMN_DELIM		varchar(1)='*',
	@RECORD_DELIM		varchar(1)='*',
	@myUniqueID		uniqueidentifier=null OUTPUT,
	@bShowWithNoWork	bit=1,	--Will just return the results without populating work table
	@bInitID		bit=0,	--If set, Main ID will be set to 0 Only do this when Copying Records
--BEGIN Text Parse vars for Table Columns
	@PolicyLimitsID		int=null,
	@AssignmentsID		int=null,
	@ID		 	int=null,
	@IDAssignments		int=null,
	@ClassTypeID 		int=null,
	@LimitAmount 		money=null,
	@RCSaidProp 		money=null,
	@Reserves 		money=null,
	@IsDeleted		bit=null,
	@DownLoadMe		bit=null,
	@UpLoadMe		bit=null,
	@AdminComments		Varchar(255)=null,
	@DateLastUpdated 	datetime=null,
	@UpdateByUserID		int=null,
	@FarmersUnitRowID	varchar(20)=null
--END Text Parse vars for Table Columns
AS
BEGIN
	--Check the DELIM CHARS
	IF @COLUMN_DELIM = '*'
	BEGIN
		SET @COLUMN_DELIM = dbo.GetCOLUMN_DELIM()
	END
	IF @RECORD_DELIM = '*'
	BEGIN
		SET @RECORD_DELIM = dbo.GetRECORD_DELIM()
	END

	-- @sTempData is used to convert different data types 
	-- to then be added to Text List
	-- ** Note** If the Data type is varchar, char, Text or ntext, 
	-- just use that var in the UPDATETEXT line without having to convert to Varchar.
	DECLARE @sTempData	varchar(8000)

	--Must create a temp table to help build the Text List
	--1. Pointer for Text field
	DECLARE @ptrList_Text		binary(16) 	
	CREATE TABLE #List_Text
	(
		ListText	Text
	) 
	--2.Initialize the List Text Table
	INSERT INTO	#List_Text
	SELECT '' As ListText
	--3.Get the List Pointer
	SELECT 	@ptrList_Text = TEXTPTR(ListText) 
	FROM 	#List_Text


	--BEGIN Create Table 
	CREATE TABLE  #tblPolicyLimits 
	(	
		PolicyLimitsID		int,
		AssignmentsID		int,
		[ID]		 	int,
		IDAssignments		int,
		ClassTypeID 		int,
		LimitAmount 		money,
		RCSaidProp 		money,
		Reserves 		money,
		IsDeleted		bit,
		DownLoadMe		bit,
		UpLoadMe		bit,
		AdminComments		Varchar(255),
		DateLastUpdated 	datetime,
		UpdateByUserID		int,
		FarmersUnitRowID	varchar(20)
	)
	--END Create Table

	--BEGIN Populate the Table
	INSERT INTO 	#tblPolicyLimits
		 	EXEC z_spsGetPolicyLimitsInfo
			/*@bHideDeleted		bit =1*/		@bHideDeleted, 
			/*@UID			int*/			@UID,
			/*@PolicyLimitsID	Int=null*/		@iPolicyLimitsID,
			/*@OrderBy		varchar(500)=null*/	@OrderBy,
			/*@GroupBy		varchar(500)=null*/	@GroupBy,
			/*@SearchBy		varchar(1000)=null*/	@SearchBy,
			/*@bDelete		bit=0*/			default,
			/*@bAdd			bit=0*/			default,
			/*@bEdit		bit=0*/			default,
			/*@AssignmentsID	int=null*/		default,
			/*@ID			int=null*/		default,
			/*@IDAssignments	int=null*/		default,
			/*@ClassTypeID		int=null*/		default,
			/*@LimitAmount		money=null*/		default,
			/*@RCSaidProp		money=null*/		default,
			/*@Reserves		money=null*/		default,
			/*@IsDeleted		bit=null*/		default,
			/*@DownLoadMe		bit=null*/		default,
			/*@UpLoadMe		bit=null*/		default,
			/*@AdminComments	varchar(255)=null*/	default,
			/*@DateLastUpdated	datetime=null*/		default,
			/*@bViewHistory		bit=NULL*/		default,
			/*@FromDate		varchar(50)=Null*/	default,
			/*@ToDate		varchar(50)=Null*/	default,
			/*@bTableFieldsOnly	bit=1*/			1,
			/*@FarmersUnitRowID	varchar(2)*/		default
	--END Populate the Table
	
	
	--BEGIN Table Cursor
	Declare RSPolicyLimits Cursor for
		SELECT * FROM #tblPolicyLimits	 	
	Open RSPolicyLimits

	Fetch Next From RSPolicyLimits into 
						@PolicyLimitsID,
						@AssignmentsID,
						@ID,
						@IDAssignments,
						@ClassTypeID,
						@LimitAmount,
						@RCSaidProp,
						@Reserves,
						@IsDeleted,
						@DownLoadMe,
						@UpLoadMe,
						@AdminComments,
						@DateLastUpdated,
						@UpdateByUserID,
						@FarmersUnitRowID

	WHILE @@FEtch_Status = 0
	BEGIN
		IF @bInitID = 1
		BEGIN
			SET @PolicyLimitsID = 0
			SET @AssignmentsID = 0
			SET @ID = 0
			SET @IDAssignments = 0
		END
		--PolicyLimitsID	int,
		SET @sTempData = Cast(@PolicyLimitsID 	as varchar(10)) 	+ @COLUMN_DELIM 	
		UPDATETEXT #List_Text.ListText @ptrList_Text null null @sTempData
		--AssignmentsID		int,
		SET @sTempData = Cast(@AssignmentsID 	as varchar(10)) 	+ @COLUMN_DELIM  	
		UPDATETEXT #List_Text.ListText @ptrList_Text null null @sTempData
		--[ID]		 	int,
		SET @sTempData =Cast(@ID 		as varchar(10)) 	+ @COLUMN_DELIM  	
		UPDATETEXT #List_Text.ListText @ptrList_Text null null @sTempData
		--IDAssignments		int,
		SET @sTempData =Cast(@IDAssignments 	as varchar(10)) 	+ @COLUMN_DELIM  	
		UPDATETEXT #List_Text.ListText @ptrList_Text null null @sTempData
		--ClassTypeID 		int,
		SET @sTempData =Cast(@ClassTypeID 	as varchar(10)) 	+ @COLUMN_DELIM  	
		UPDATETEXT #List_Text.ListText @ptrList_Text null null @sTempData
		--LimitAmount 		money,
		SET @sTempData =Cast(@LimitAmount 	as varchar(20)) 	+ @COLUMN_DELIM  	
		UPDATETEXT #List_Text.ListText @ptrList_Text null null @sTempData
		--RCSaidProp 		money,
		SET @sTempData =Cast(@RCSaidProp 	as varchar(20)) 	+ @COLUMN_DELIM  	
		UPDATETEXT #List_Text.ListText @ptrList_Text null null @sTempData
		--Reserves 		money,
		SET @sTempData =Cast(@Reserves 		as varchar(20)) 	+ @COLUMN_DELIM  	
		UPDATETEXT #List_Text.ListText @ptrList_Text null null @sTempData
		--IsDeleted		bit,
		SET @sTempData =Cast(@IsDeleted 	as varchar(1)) 		+ @COLUMN_DELIM  	
		UPDATETEXT #List_Text.ListText @ptrList_Text null null @sTempData
		--DownLoadMe		bit,
		SET @sTempData =Cast(@DownLoadMe 	as varchar(1)) 		+ @COLUMN_DELIM  	
		UPDATETEXT #List_Text.ListText @ptrList_Text null null @sTempData
		--UpLoadMe		bit,
		SET @sTempData =Cast(@UpLoadMe 		as varchar(1)) 		+ @COLUMN_DELIM  	
		UPDATETEXT #List_Text.ListText @ptrList_Text null null @sTempData
		--AdminComments		Varchar(255),
		SET @sTempData =Cast(@AdminComments 	as varchar(255)) 	+ @COLUMN_DELIM 	
		UPDATETEXT #List_Text.ListText @ptrList_Text null null @sTempData
		-- DateLastUpdated 	datetime,
		SET @sTempData =Cast(@DateLastUpdated 	as varchar(50)) 	+ @COLUMN_DELIM  	
		UPDATETEXT #List_Text.ListText @ptrList_Text null null @sTempData
		--UpdateByUserID	int
		SET @sTempData =Cast(@UpdateByUserID 	as varchar(10)) 	+ @COLUMN_DELIM 	
		UPDATETEXT #List_Text.ListText @ptrList_Text null null @sTempData
		--FarmersUnitRowID
		SET @sTempData =Cast(@FarmersUnitRowID 	as varchar(50)) 	+ @RECORD_DELIM 	
		UPDATETEXT #List_Text.ListText @ptrList_Text null null @sTempData
						
		Fetch Next From RSPolicyLimits into 
							@PolicyLimitsID,
							@AssignmentsID,
							@ID,
							@IDAssignments,
							@ClassTypeID,
							@LimitAmount,
							@RCSaidProp,
							@Reserves,
							@IsDeleted,
							@DownLoadMe,
							@UpLoadMe,
							@AdminComments,
							@DateLastUpdated,
							@UpdateByUserID,
							@FarmersUnitRowID
	END

	Close RSPolicyLimits
	Deallocate RSPolicyLimits
	Drop Table #tblPolicyLimits
	--END Table Cursor
	
	IF @bShowWithNoWork = 1
	BEGIN
		SELECT 	ListText
		FROM	#List_Text
	END
	ELSE
	BEGIN
		--Need to Populate Work table
		SET @myUniqueID = NEWID()
		INSERT INTO 	Work_ListText 
		SELECT  	@myUniqueID 			As Work_ListTextID, 
				#List_Text.ListText		As ListText
		FROM #List_Text
	END
	

	Drop Table #List_Text
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[z_spsList_PolicyLimits]  TO [CFUSER]
GO

