
/**********************************************************************************************
*       Jira:           2018.09.07 TAP2-46 Optimization Dynamic SQL to SQL SP Iteration 2
*                       https://microstartap3.atlassian.net/browse/TAP2-46
*    	Description:	TAP Get Company Search Data Billing [HTML]
*		Returns:        Create pseudo Temp Table @pTapGUID
*                       Insert Rows of MAX LEN 7000	chars into @pTapGUID.[CompSearchDataBillingContentHtml]
*       Note:           Caller must Drop Table @pTapGUID as cleanup                            
*	
*	Author: 	Brad Skidmore
*	Date: 		10/09/2018
*   Date:       10/18/2018 Modify ShipState size for 2018.09.18 TAP2-55 Widen the BillState and ShipState fields in the USER table
*   Date:       10/23/2018 Bug with apostrophes in search field
**********************************************************************************************/
CREATE PROCEDURE [dbo].[tap_spsGetCompanySearchDataBillingHtml] (
@pTAPAppVersion varchar(50)='TAP2.5',   --Versioning of HTML output
@pTapGUID nvarchar(80)=null,            --The GUID [without dashes] Table Name for creation of a pseudo temp table that will contain html rows.
@pQuerySearchText nvarchar(500),        --User entered Query Search Text
@pCompanyCriteria nvarchar(500),        --Could be a single CompanyID OR a comma separated list of Bigint CompanyID
@pPermExec nvarchar(20),                --Request.Cookies("PermExec")
@pOrderBy varchar(50)='[NAME]',         --Used to generate the correct sort order: SELECT ROW_NUMBER() OVER(ORDER BY ' + Replace(@pOrderBy, '''', '''''') + ') AS [mySort].
@pTopLimit bigint=20,                   --Limit the Results to the @pTopLimit amount
@pDebugOn bit=0                         --Debugging?  SET to 1 IF NOT SET to 0  Will return a robust set of queries in order to interrogate/reconcile the sp results.
)
AS
BEGIN
    --Begin Validate Params:
    BEGIN
        IF NULLIF(@pTAPAppVersion, '') Is Null
        BEGIN
            SET @pTAPAppVersion = 'TAP2.5'
        END
        
        IF NULLIF(@pTapGUID, '') Is Null
        BEGIN
            SET @pTapGUID = 'tap_spsGetCompanySearchDataBillingHtml_' + Cast(newid() as nvarchar(36))
        END
        
        IF NULLIF(@pQuerySearchText, '') Is Null
        BEGIN
            SET @pQuerySearchText = '--INVALID_SEARCH_TEXT--'
        END
        
        IF NULLIF(@pCompanyCriteria, '') Is Null
        BEGIN
            SET @pCompanyCriteria = ''
        END
        
        IF NULLIF(@pPermExec, '') Is Null
        BEGIN
            SET @pPermExec = ''
        END
        
        IF NULLIF(@pTopLimit, 0) Is Null
        BEGIN
            SET @pTopLimit = 20
        END
        
        IF NULLIF(@pOrderBy, '') Is Null
        BEGIN
            SET @pOrderBy = '[NAME]'
        END
    END     
    --Debug
    IF @pDebugOn = 1
    BEGIN
        --List Current Params
        SELECT  @pTAPAppVersion As [@pTAPAppVersion]
            , @pTapGUID As [@pTapGUID]
            , @pQuerySearchText As [@pQuerySearchText]
            , @pCompanyCriteria As [@pCompanyCriteria]
            , @pPermExec As [@pPermExec]
            , @pOrderBy As [@pOrderBy] 
            , @pDebugOn As [@pDebugOn] 
    END
    --Debug
   
    --2018.08.29 only include the required fields
    -- [myID]
    --, [mySort]
    --, [ID]
    --, [NAME]
    --, [ShipAddress1]
    --, [ShipCity]
    --, [ShipState]
    --, [ShipZip]
    --, [CustType]
    --, [PermCreditHold]
    --, [OptPrePay]
    --, [Company]
    --BEGIN ALL KNOWN SORT COLUMNS
    
    --ENDALL KNOWN SORT COLUMNS
    
    CREATE TABLE #myUSERS 
	(
		[myID] bigint IDENTITY(1,1) NOT NULL
        , [mySort] bigint NOT NULL
        , [ID] nvarchar(20) NOT NULL
        , [Name] nvarchar(100) NULL
        , [ShipAddress1] nvarchar(100) NULL
        , [ShipCity] nvarchar(50) NULL
        , [ShipState] nvarchar(20) NULL
        , [ShipZip] nvarchar(12) NULL
        , [CustType] nvarchar(20) NULL
        , [PermCreditHold] nvarchar(3) NULL
        , [OptPrepay] nvarchar(3) NULL
        , [Company] bigint NULL
        --, [Enabler] nvarchar(3) NULL
        --, [Username] nvarchar(20) NULL
        --, [Password] nvarchar(20) NULL
        --, [Region] nvarchar(50) NULL
        --, [Email] nvarchar(100) NULL
        --, [DynID] nvarchar(50) NULL
        --, [ContractStart] datetime NULL
        --, [ContractEnd] datetime NULL
        --, [InactiveDate] datetime NULL
        --, [BillLatLon] nvarchar(100) NULL
        --, [ShipLatLon] nvarchar(100) NULL
        --, [BillName] nvarchar(100) NULL
        --, [BillName2] nvarchar(100) NULL
        --, [BillAddress1] nvarchar(100) NULL
        --, [BillAddress2] nvarchar(100) NULL
        --, [BillAddress3] nvarchar(100) NULL
        --, [BillCity] nvarchar(50) NULL
        --, [BillZip] nvarchar(12) NULL
        --, [BillState] nvarchar(20) NULL
        --, [BillContact] nvarchar(100) NULL
        --, [BillContact2] nvarchar(100) NULL
        --, [BillPhone] nvarchar(30) NULL
        --, [BillPhone2] nvarchar(30) NULL
        --, [BillFax] nvarchar(30) NULL
        --, [BillEmail] nvarchar(200) NULL
        --, [BillEmail2] nvarchar(200) NULL
        --, [BillCountry] nvarchar(200) NULL
        --, [ShipName] nvarchar(100) NULL
        --, [ShipName2] nvarchar(100) NULL
        --, [ShipAddress2] nvarchar(100) NULL
        --, [ShipAddress3] nvarchar(100) NULL
        --, [ShipContact] nvarchar(100) NULL
        --, [ShipContact2] nvarchar(100) NULL
        --, [ShipPhone] nvarchar(30) NULL
        --, [ShipPhone2] nvarchar(30) NULL
        --, [ShipFax] nvarchar(30) NULL
        --, [ShipEmail] nvarchar(200) NULL
        --, [ShipEmail2] nvarchar(200) NULL
        --, [ShipCountry] nvarchar(200) NULL
        --, [ShipNotes] text NULL
        --, [HBInventory] bigint NOT NULL  DEFAULT ((0))
        --, [SBInventory] bigint NOT NULL  DEFAULT ((0))
        --, [PermExec] nvarchar(3) NULL
        --, [PermTrafficWest] nvarchar(3) NULL
        --, [PermTrafficEast] nvarchar(3) NULL
        --, [PermTrafficMidwest] nvarchar(3) NULL
        --, [PermTrafficInternational] nvarchar(3) NULL
        --, [PermTrafficAckWest] nvarchar(3) NULL
        --, [PermTrafficAckEast] nvarchar(3) NULL
        --, [PermTrafficAckMidwest] nvarchar(3) NULL
        --, [PermTrafficAckInternational] nvarchar(3) NULL
        --, [PermBilling] nvarchar(3) NULL
        --, [PermSettings] nvarchar(3) NULL
        --, [PermHistory] nvarchar(3) NULL
        --, [PermTrafficManager] nvarchar(3) NULL
        --, [PermBillingManager] nvarchar(3) NULL
        --, [PermHistorySearch] nvarchar(3) NULL
        --, [PermDistEmpties] nvarchar(3) NOT NULL
        --, [PermTrafficBrewSum] nvarchar(3) NULL
        --, [PermOpenCollars] nvarchar(3) NULL
        --, [OptPostPay] nvarchar(3) NULL
        --, [OptRetrieval] nvarchar(3) NULL
        --, [OptRetrievalFee] money NOT NULL  DEFAULT((0))
        --, [OptInactivity] nvarchar(3) NULL
        --, [OptInactivityFee] money NOT NULL  DEFAULT((0))
        --, [OptInactivityFee2] SMALLINT NULL
        --, [OptInactivityFeeDays] SMALLINT NULL
        --, [OptForeign] nvarchar(3) NULL
        --, [OptForeignFee] money NOT NULL  DEFAULT((0))
        --, [OptPallet] nvarchar(3) NULL
        --, [OptPalletFee] money NOT NULL  DEFAULT((0))
        --, [OptExpedite] nvarchar(3) NULL
        --, [OptExpediteFee1] money NOT NULL  DEFAULT((0))
        --, [OptExpediteFeeDays1] SMALLINT NULL
        --, [OptExpediteFee2] money NOT NULL  DEFAULT((0))
        --, [OptExpediteFeeDays2] SMALLINT NULL
        --, [OptTakePay] nvarchar(3) NULL
        --, [OptTakePayPercent] FLOAT NOT NULL  DEFAULT ((0))
        --, [OptTakePayFeeDays] nvarchar(20) NULL
        --, [OptTakePayFeeDays1] nvarchar(3) NULL
        --, [OptTakePayFeeDays2] nvarchar(3) NULL
        --, [OptTakePayFeeDays3] nvarchar(3) NULL
        --, [OptTakePayFeeDays4] nvarchar(3) NULL
        --, [OptTakePayFeeDays5] nvarchar(3) NULL
        --, [OptTakePayFeeDays6] nvarchar(3) NULL
        --, [OptTakePayFeeDays7] nvarchar(3) NULL
        --, [OptTakePayFeeDays8] nvarchar(3) NULL
        --, [OptTakePayFeeDays9] nvarchar(3) NULL
        --, [OptTakePayFeeDays10] nvarchar(3) NULL
        --, [OptTakePayFeeDays11] nvarchar(3) NULL
        --, [OptTakePayFeeDays12] nvarchar(3) NULL
        --, [OptTakePayFeeHalfBbl] money NOT NULL  DEFAULT((0))
        --, [OptTakePayFeeSixthBbl] money NOT NULL  DEFAULT((0))
        --, [OptTakePayFeeHalfBbl2] SMALLINT NULL
        --, [OptTakePayFeeSixthBbl2] SMALLINT NULL
        --, [OptDistDeposit] nvarchar(3) NULL
        --, [OptDistDepositFee] money NOT NULL  DEFAULT((0))
        --, [OptKMPI] nvarchar(3) NULL
        --, [OptUse] nvarchar(3) NULL
        --, [OptUseFee] money NULL
        --, [OptCollar] nvarchar(3) NULL
        --, [OptCollarFee] money NULL  DEFAULT((0))
        --, [OptPopupReminder] nvarchar(3) NULL
        --, [OptRebillDeposit] nvarchar(3) NULL
        --, [OptCustomBillRules] nvarchar(3) NULL
        --, [TempComments] nvarchar(500) NULL
        --, [AddedBy] nvarchar(20) NULL
        --, [AddedDate] datetime NULL
        --, [DateTake] SMALLINT NULL
        --, [DatePut] SMALLINT NULL
        --, [BillingInvLead] nvarchar(20) NULL
        --, [OptBudgetRevFill] money NULL
        --, [Pointer] nvarchar(20) NULL
        --, [SupressStatement] nvarchar(3) NULL
        --, [OptBillPaper] nvarchar(3) NULL
        --, [OptBillEmail] nvarchar(3) NULL
        --, [OptSelfDist] nvarchar(3) NULL
        --, [OptSelfDistParent] nvarchar(20) NULL
        --, [BillingRep] nvarchar(40) NULL
        --, [TrafficRep] nvarchar(40) NULL
        --, [ActiveRep] tinyint NULL
        --, [ContactFormAdmin] tinyint NULL
        --, [OptCollarPlateFee] money NOT NULL  DEFAULT ((0))
        --, [OptCollarPlate] nvarchar(3) NULL
        --, [optproductselect] nvarchar(3) NULL
        --, [optbadpalletletter] nvarchar(3) NULL
        --, [PermTrafficWine] nvarchar(3) NULL
        --, [PermTrafficAckWine] nvarchar(3) NULL
        --, [permhistadv] nvarchar(3) NULL
        --, [permsource] nvarchar(20) NULL
        --, [OptBadPalletMethod] nvarchar(10) NULL
        --, [RentalCustomer] BIT NULL  DEFAULT ((0))
        --, [SalesCustomer] BIT NULL
        --, [SalesRep] nvarchar(40) NULL
        --, [ReceivingType] nvarchar(15) NULL DEFAULT ('business dock')
        --, [FreightViewAPIKey] nvarchar(50) NULL
        --, [TAP2APIKey] nvarchar(50) NULL
        --, [TAP2APITranslateFlag] BIT NOT NULL  DEFAULT ((0))
	)

    --Insert into #myUSERS
    --2018.08.29 only include the required fields
    -- [myID]
    --, [mySort]
    --, [ID]
    --, [NAME]
    --, [ShipAddress1]
    --, [ShipCity]
    --, [ShipState]
    --, [ShipZip]
    --, [CustType]
    --, [PermCreditHold]
    --, [OptPrePay]
    --, [Company]
    DECLARE @mySQL As varchar(8000)
    SET @mySQL = 'INSERT INTO #myUSERS([mySort], [ID], [NAME], [ShipAddress1], [ShipCity], [ShipState], [ShipZip], [CustType], [PermCreditHold], [OptPrePay], [Company] ) '
    SET @mySQL = @mySQL + 'SELECT Top ' + Replace(Cast(@pTopLimit as varchar(20)), '-', '') + ' ROW_NUMBER() OVER(ORDER BY ' + Replace(@pOrderBy, '''', '''''') + ') AS [mySort], [ID], Replace([Name], ''"'', '''') as [Name], [ShipAddress1], [ShipCity], [ShipState], [ShipZip], [CustType], [PermCreditHold], [OptPrePay], [Company] '
    SET @mySQL = @mySQL + 'FROM [dbo].[USERS] '
    
    --Debug
    IF @pDebugOn = 1
    BEGIN
        Select @mySQL as [@mySQL_PART001]
    END
    --Debug
    
    --Build the Where Clause for Different Content Types
    SET @mySQL = @mySQL + 'WHERE 1=1 AND '
    --Company Criteria
    IF (@pCompanyCriteria <> '-1') AND (Len(RTrim(LTrim(@pCompanyCriteria))) > 0)
    BEGIN
        IF Charindex(',', @pCompanyCriteria) > 0
        BEGIN
            SET @pCompanyCriteria = 'Company IN (' + Replace(@pCompanyCriteria, '''', '''''') + ') AND '
        END
		ELSE
        BEGIN
            SET @pCompanyCriteria = 'Company = ' + Replace(@pCompanyCriteria, '''', '''''') + ' AND '
        END
    END
    
    --Query Search Text
    DECLARE @CleanQuerySearchText As nvarchar(500)
    SET @CleanQuerySearchText = Replace(@pQuerySearchText, '''', '''''')
    SET @CleanQuerySearchText = Replace(@CleanQuerySearchText, ' ', '%%')
    SET @mySQL = @mySQL + '(' 
    SET @mySQL = @mySQL + '[ShipPhone] LIKE ''%%' + @CleanQuerySearchText + '%%'' OR '
    SET @mySQL = @mySQL + '[BillPhone] LIKE ''%%' + @CleanQuerySearchText + '%%'' OR '
    SET @mySQL = @mySQL + '[BillAddress2] LIKE ''%%' + @CleanQuerySearchText + '%%'' OR '
    SET @mySQL = @mySQL + '[BillAddress1] LIKE ''%%' + @CleanQuerySearchText + '%%'' OR ' 
	SET @mySQL = @mySQL + '[ShipAddress2] LIKE ''%%' + @CleanQuerySearchText + '%%'' OR '
	SET @mySQL = @mySQL + '[ShipAddress1] LIKE ''%%' + @CleanQuerySearchText + '%%'' OR '
	SET @mySQL = @mySQL + '[ShipEmail] LIKE ''%%' + @CleanQuerySearchText + '%%'' OR '
	SET @mySQL = @mySQL + '[BillEmail] LIKE ''%%' + @CleanQuerySearchText + '%%'' OR ' 
	SET @mySQL = @mySQL + '[Email] LIKE ''%%' + @CleanQuerySearchText + '%%'' OR '
	SET @mySQL = @mySQL + '[ID] LIKE ''%%' + @CleanQuerySearchText + '%%'' OR '
	SET @mySQL = @mySQL + '[Username] LIKE ''%%' + @CleanQuerySearchText + '%%'' OR '
	SET @mySQL = @mySQL + '[NAME] LIKE ''%%' + @CleanQuerySearchText + '%%'' '
	SET @mySQL = @mySQL + ') '
	
	IF lower(RTrim(LTrim(@pPermExec))) <> 'yes'
	BEGIN
	    SET @mySQL = @mySQL + 'AND CustType != ''Employee'' '
	END
    
    --Debug
    IF @pDebugOn = 1
    BEGIN
        Select @mySQL as [@mySQL_PART002_Final]
    END
        
    --INSERT INTO TEMP
    BEGIN TRANSACTION
        EXEC(@mySQL)
    COMMIT;
    
    IF @pDebugOn = 1
    BEGIN
        --Select
        --2018.08.29 only include the required fields
        -- [myID]
        --, [mySort]
        --, [ID]
        --, [NAME]
        --, [ShipAddress1]
        --, [ShipCity]
        --, [ShipState]
        --, [ShipZip]
        --, [CustType]
        --, [PermCreditHold]
        --, [OptPrePay]
        --, [Company]
        --BEGIN ALL KNOWN SORT COLUMNS

        --ENDALL KNOWN SORT COLUMNS

        --SELECT [myID], [ID], [ParentOrderID], [OrderID], [Company], [Status], [MoveType], [Brewer], [BrewerName], [BrewerRevenue], [Distributor], [DistributorName], [DistributorRevenue], [Carrier], [CarrierFee], [CarrierQuote], [BadPallets], [ForeignKegs], [Pallets], [HalfBbl], [SixthBbl], [HalfBblDistributor], [SixthBblDistributor], [PalletsDistributor], [TentPallets], [TentHalfBbl], [TentSixthBbl], [DateShipped], [DateVerified], [DateAck], [DateReceived], [DateReported], [DateBilled], [DatePosted], [DateETA], [QuoteShipExpense1], [QuoteShipName1], [QuoteSelect1], [QuoteQuoteNum1], [QuoteShipExpense2], [QuoteShipName2], [QuoteSelect2], [QuoteQuoteNum2], [QuoteShipExpense3], [QuoteShipName3], [QuoteSelect3], [QuoteQuoteNum3], [ActualShipExpense], [RequireCustomer], [IgnoreExp], [ProNumber], [BilledBy], [BilledBatchNum], [LocalMove], [MassUploadDate], [MassUploadKey], [IgnoreError], [Comments], [BrewerDynID], [DistributorDynID], [Verified], [Source], [SourceName], [Destination], [DestinationName], [ShipCostActual], [DepositCheck], [OptExpediteFee], [OptCustomDeposit], [MoveTypeComment], [OptZeroDeposit], [newkegs], [revrelease], [HalfBblProduct], [SixthBblProduct], [NoBill], [NoBillReason], [NoBillAllow], [locktrx], [locktrxdate], [locktrxreason], [PalletHeight], [Weight], [APIShipID], [OutCount]
        SELECT [myID]
            , [mySort]
            , [ID]
            , [NAME]
            , [ShipAddress1]
            , [ShipCity]
            , [ShipState]
            , [ShipZip]
            , [CustType]
            , [PermCreditHold]
            , [OptPrePay]
            , [Company]
        FROM #myUSERS
        Order By [mySort]
    END
    
    --Create The Temp HTML Output Table
    DECLARE @myHtmlOutPut as varchar(80)
    SET @myHtmlOutPut = @pTapGUID
    SET @mySQL = 'CREATE TABLE ' + @myHtmlOutPut + ' ' 
	SET @mySQL = @mySQL + '( '
    SET @mySQL = @mySQL + '[myID] bigint IDENTITY(1,1) NOT NULL '
    SET @mySQL = @mySQL + ', [mySort] bigint NOT NULL '
    SET @mySQL = @mySQL + ', [CompSearchDataBillingContentHtml] varchar(7000) NULL '
    SET @mySQL = @mySQL + ') '
    
    BEGIN TRANSACTION
        EXEC(@mySQL)
    COMMIT;
    
    CREATE TABLE #myHtmlOutPut
    (     
        [myID] bigint IDENTITY(1,1) NOT NULL
        , [mySort] bigint NOT NULL
        , [CompSearchDataBillingContentHtml] varchar(7000) NULL 
    ) 
       
    
    --------------------------------------------------------------------------------------------------------------------------
    --BEGIN CURSOR 
    --------------------------------------------------------------------------------------------------------------------------
        --@HTML Current ROW
        DECLARE @HtmlRow as varchar(2000)
        DECLARE @mySort as bigint
        DECLARE @CompSearchDataBillingContentHtmlInsert as varchar(7000)
        --Declare Required Vars from #myUSERS to call dbo.tap_fnGetDDWNfrmtHtml
        DECLARE @iCompany As bigint
        DECLARE @ID nvarchar(20)                     
        DECLARE @Name nvarchar(100)                   
        DECLARE @ShipAddress1 nvarchar(100)           
        DECLARE @ShipCity nvarchar(50)                
        DECLARE @ShipState nvarchar(20)                
        DECLARE @ShipZip nvarchar(12)                 
        DECLARE @CustType nvarchar(20)
        --Derived from PermCreditHold=yes OR OptPrePay=yes
        DECLARE @ExtraHTML nvarchar(150)                      
        Declare CurUSERS Cursor for
            SELECT  (CASE WHEN NULLIF(U.[Company], 0) Is Not Null Then U.[Company] Else Cast(0 As bigint) END) As [Company]
            , (CASE WHEN NULLIF(U.[ID], '') Is Not Null Then U.[ID] Else Cast('' As nvarchar(20)) END) AS [ID]
            , (CASE WHEN NULLIF(U.[NAME], '') Is Not Null Then U.[NAME] Else Cast('' As nvarchar(100)) END) As [NAME]
            , (CASE WHEN NULLIF(U.[ShipAddress1], '') Is Not Null Then U.[ShipAddress1] Else Cast('' As nvarchar(100)) END) As [ShipAddress1]
            , (CASE WHEN NULLIF(U.[ShipCity], '') Is Not Null Then U.[ShipCity] Else Cast('' As nvarchar(50)) END) As [ShipCity]
            , (CASE WHEN NULLIF(U.[ShipState], '') Is Not Null Then U.[ShipState] Else Cast('' As nvarchar(20)) END) As [ShipState]
            , (CASE WHEN NULLIF(U.[ShipZip], '') Is Not Null Then U.[ShipZip] Else Cast('' As nvarchar(12)) END) As [ShipZip]
            , (CASE WHEN NULLIF(U.[CustType], '') Is Not Null Then U.[CustType] Else Cast('' As nvarchar(20)) END) As [CustType]
            , (
                ( 
                    CASE WHEN (Lower(CASE WHEN NULLIF(U.[PermCreditHold], '') Is Not Null Then U.[PermCreditHold] Else '' END)) = 'yes' 
                    THEN 
                        '&nbsp;&nbsp;&nbsp;<font color=red>*** CREDIT HOLD</font>'
                    ELSE 
                        '' 
                    END
                )
                +
                ( 
                    CASE WHEN (Lower(CASE WHEN NULLIF(U.[OptPrePay], '') Is Not Null Then U.[OptPrePay] Else '' END)) = 'yes' 
                    THEN 
                        '&nbsp;&nbsp;&nbsp;<font color=red>*** PREPAY CUSTOMER</font>' 
                    ELSE 
                        '' 
                    END
                ) 
            ) As [ExtraHTML]
            
            FROM #myUSERS U
            Order By U.[mySort]
    --------------------------------------------------------------------------------------------------------------------------
    --END CURSOR 
    --------------------------------------------------------------------------------------------------------------------------
    --Init some vars
    BEGIN
        SET @mySort = 0
        SET @CompSearchDataBillingContentHtmlInsert = ''
    END
    Open CurUSERS
        Fetch Next From CurUSERS into
            @iCompany
            , @ID                     
            , @Name                   
            , @ShipAddress1           
            , @ShipCity                
            , @ShipState                
            , @ShipZip                 
            , @CustType
            , @ExtraHTML               
		While @@FEtch_Status = 0		
		BEGIN
		    --------------------------------------------------------------------------------------------------------------------------
            --BEGIN dbo.tap_fnGetDDWNfrmtHtml
            --------------------------------------------------------------------------------------------------------------------------
		    SET	@HtmlRow = dbo.tap_fnGetDDWNfrmtHtml(@iCompany
                , @ID
                , @Name
                , @ShipAddress1
                , @ShipCity
                , @ShipState
                , @ShipZip
                , @CustType
                , @ExtraHTML)
            --------------------------------------------------------------------------------------------------------------------------
            --END dbo.tap_fnGetDDWNfrmtHtml
            --------------------------------------------------------------------------------------------------------------------------
            
            ------------------------------------------------------------------------------------------------------BUILD the HTML INSERT
            IF Len(@CompSearchDataBillingContentHtmlInsert) + Len(@HtmlRow) <= 7000
            BEGIN
                SET @CompSearchDataBillingContentHtmlInsert = @CompSearchDataBillingContentHtmlInsert + @HtmlRow
            END
            ELSE
            BEGIN
                SET @mySort = @mySort + 1
                IF @pDebugOn = 1
                BEGIN
                    SELECT @mySort As [@mySort] , @CompSearchDataBillingContentHtmlInsert As [@CompSearchDataBillingContentHtmlInsert]
                END
                BEGIN TRANSACTION
                    INSERT INTO #myHtmlOutPut ([mySort], [CompSearchDataBillingContentHtml]) Values(@mySort, @CompSearchDataBillingContentHtmlInsert)
                COMMIT;
                --RESET
                SET @CompSearchDataBillingContentHtmlInsert = @HtmlRow
            END
            
            Fetch Next From CurUSERS into
                @iCompany
                , @ID                     
                , @Name                   
                , @ShipAddress1           
                , @ShipCity                
                , @ShipState                
                , @ShipZip                 
                , @CustType
                , @ExtraHTML           
		END
		
		Close CurUSERS
		Deallocate CurUSERS
		
		--Final Insert
		IF Len(@CompSearchDataBillingContentHtmlInsert) > 0
		BEGIN
            SET @mySort = @mySort + 1
            IF @pDebugOn = 1
            BEGIN
                SELECT @mySort As [@mySort] , @CompSearchDataBillingContentHtmlInsert As [@CompSearchDataBillingContentHtmlInsert]
            END
            BEGIN TRANSACTION
                INSERT INTO #myHtmlOutPut ([mySort], [CompSearchDataBillingContentHtml]) Values(@mySort, @CompSearchDataBillingContentHtmlInsert)
            COMMIT;
		END
        
        --Finally Output the HTML
        SET @mySQL = 'INSERT INTO dbo.[' + Replace(@myHtmlOutPut, '''', '''''') + '] ([mySort], [CompSearchDataBillingContentHtml]) Select [mySort], [CompSearchDataBillingContentHtml] From #myHtmlOutPut Order By [mySort] '
        
        IF @pDebugOn = 1
        BEGIN
            SELECT @mySQL As [@mySQL_INSERT INTO_GUID]
        END
         
        BEGIN TRANSACTION
            EXEC(@mySQL)
        COMMIT;
    
    --Cleanup
    Drop Table #myHtmlOutPut
    Drop Table #myUSERS
END
GO
