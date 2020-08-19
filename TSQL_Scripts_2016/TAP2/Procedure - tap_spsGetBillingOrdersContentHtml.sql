
/******************************************************************
*       Jira:           2018.08.21 TAP2-31 Billing Optimization
*                       https://microstartap3.atlassian.net/browse/TAP2-31
*    	Description:	TAP Get Billing Orders Content [HTML]
*		Returns:        Create pseudo Temp Table @pTapGUID
*                       Insert Rows of MAX LEN 7000	chars into @pTapGUID.[BOContentHtml]
*       Note:           Caller must Drop Table @pTapGUID as cleanup                            
*	
*	Author: 	Brad Skidmore
*	Date: 		8/24/2018
******************************************************************/
CREATE PROCEDURE [dbo].[tap_spsGetBillingOrdersContentHtml] (
@pTAPAppVersion          varchar(50)='TAP2.5',      --Versioning of HTMLL output
@pTapGUID                nvarchar(80)=null,         --The GUID [without dashes] Table Name for creation of a pseudo temp table that will contain html rows.
@pContentType            varchar(50)='default',     --The Name of the scenario/Content Type [For building the main Where Clause].
@pIDLimitBrewer          nvarchar(50)=null,         --When @pContentType = 'IDLimit' AND WHERE [Brewer] = ''@pIDLimitBrewer'' AND [Status] = ''Completed''.
@pNameLimitBrewerName    nvarchar(200)=null,        --When @pContentType = 'IDLimit' AND WHERE [BrewerName] like ''%%@pNameLimitBrewerName%%'' AND [Status] = ''Completed''.
@pMassUploadKey          nvarchar(20)=null,         --When @pContentType = 'MassUpload' AND WHERE [MassUploadKey] = ''' + Cast(RTrim(LTrim(Replace(@pMassUploadKey, '''',''''''))) As varchar(200)) + ''' AND Status = ''Completed''.
@pCompany                bigint=null,               --When @pCompany > 0 THEN main where clause Adds:  AND [Company] = ' + Replace(Cast(@pCompany As varchar(20)), '''', '''''').
@pRegion                 nvarchar(50)='',           --{strOnClickPassit4Region} = @pRegion HTML Row onClick="window.parent.passit4(''{strID}'', ''{strOnClickPassit4Region}'', ''{strOnClickPassit4Tab}'');Eventually Passed in to function manageshipment(ID, Region, Tab, Force) in home.asp.
@pTab                    nvarchar(50)='',           --{strOnClickPassit4Tab} = @pTab HTML Row onClick="window.parent.passit4(''{strID}'', ''{strOnClickPassit4Region}'', ''{strOnClickPassit4Tab}'');Eventually Passed in to function manageshipment(ID, Region, Tab, Force) in home.asp.
@pOrderBy                varchar(50)='[ID] DESC',   --Used to generate the correct sort order: SELECT ROW_NUMBER() OVER(ORDER BY ' + Replace(@pOrderBy, '''', '''''') + ') AS [mySort].
@pDebugOn                bit=0                      --Debugging?  SET to 1 IF NOT SET to 0  Will return a robust set of queries in order to interrogate/reconcile the sp results.
)
AS
BEGIN
    DECLARE @myContentType varchar(50)
    DECLARE @mySQL varchar(8000)
    DECLARE @mySelectSQL varchar(2000)
    
    --Begin Validate Params:
    BEGIN
        IF NULLIF(@pTAPAppVersion, '') Is Null
        BEGIN
            SET @pTAPAppVersion = 'TAP2.5'
        END
        
        IF NULLIF(@pTapGUID, '') Is Null
        BEGIN
            SET @pTapGUID = 'myHtmlOutPut_' + Cast(newid() as varchar(36))
        END
        
        IF NULLIF(@pContentType, '') Is Null
        BEGIN
            SET @pContentType = 'default'
        END
        SET @myContentType = Lower(RTrim(LTrim(@pContentType)))
        
        IF NULLIF(@pCompany, -1) Is Null
        BEGIN
            SET @pCompany = -1
        END
        
        IF NULLIF(@pRegion, '') Is Null
        BEGIN
            SET @pRegion = ''
        END
        
        IF NULLIF(@pTab, '') Is Null
        BEGIN
            SET @pTab = ''
        END
        
        IF NULLIF(@pOrderBy, '') Is Null
        BEGIN
            SET @pOrderBy = '[ID] DESC'
        END
    END     
    --Debug
    IF @pDebugOn = 1
    BEGIN
        --List Current Params
        SELECT 
            @pTAPAppVersion As [@pTAPAppVersion]
            , @pTapGUID As [@pTapGUID]
            , @pContentType As [@pContentType]
            , @pIDLimitBrewer As [@pIDLimitBrewer]
            , @pNameLimitBrewerName As [@pNameLimitBrewerName]
            , @pMassUploadKey As [@pMassUploadKey]
            , @pCompany As [@pCompany]
            , @pRegion As [@pRegion]
            , @pTab As [@pTab]
            , @pOrderBy As [@pOrderBy]
            , @pDebugOn As [@pDebugOn]
       --List All Possible Content Types
        Select 1 As [ID], 'NoBill' As [@pContentType], 'WHERE [Status] = ''Completed'' AND [NoBill] = ''yes''' As [Description]
            Union Select  2 As [ID], 'CustTypeBrewers' As [@pContentType], 'WHERE [MoveType] = ''Distributor to Brewer'' AND [Status] = ''Completed''' As [Description] 
            Union Select  3 [ID], 'CustTypeDistributors' As [@pContentType], 'WHERE [MoveType] = ''Brewer to Distributor'' AND [Status] = ''Completed''' As [Description] 
            Union Select  4 As [ID], 'Rebill' As [@pContentType], 'WHERE [Status] = ''Re-bill''' As [Description]
            Union Select  5 As [ID], 'Hold' As [@pContentType], 'WHERE [Status] = ''Hold''' As [Description]
            Union Select  6 As [ID], 'SelfInput' As [@pContentType], 'WHERE [Status] = ''Completed - Brewer Input''' As [Description]
            Union Select  7 As [ID], 'CompletedAll' As [@pContentType], 'WHERE [Status] = ''Completed''' As [Description]
            Union Select  8 As [ID], 'CompletedAllTraffic' As [@pContentType], 'WHERE [Status] = ''Completed'' AND [MassUploadKey] Is NULL' As [Description]
            Union Select  9 As [ID], 'PreBillAll' As [@pContentType], 'WHERE [Status] = ''Pre-Billed''' As [Description]
            Union Select  10 As [ID], 'IDLimit' As [@pContentType], 'WHERE [Brewer] = ''@pIDLimitBrewer'' AND [Status] = ''Completed''	NameLimit	{OR}	WHERE [BrewerName] like ''%%@pNameLimitBrewerName%%'' AND [Status] = ''Completed''' As [Description]
            Union Select  11 As [ID], 'Test' As [@pContentType], 'WHERE [BrewerName] like ''%%test%%brew%%''' As [Description]
            Union Select  12 As [ID], 'MassUpload' As [@pContentType], 'WHERE [MassUploadKey] = ''@pMassUploadKey'' AND Status = ''Completed''' As [Description]
            Union Select  13 As [ID], 'Default' As [@pContentType], 'WHERE [MoveType] <> ''Brewer to Distributor'' AND [MoveType] <> ''Distributor to Brewer'' AND [Status] = ''Completed''' As [Description]
    END
    --Debug
    
    --2018.08.29 only include the required fields
    -- [myID]
    -- , [mySort]
    --BEGIN ALL KNOWN SORT COLUMNS
    --, [DateAck]
    --, [BrewerName]
    --, [DistributorName]
    --, [DateShipped] --Already in main columns below
    --, [HalfBbl] --Already in main columns below
    --, [SixthBbl] --Already in main columns below
    --ENDALL KNOWN SORT COLUMNS
    -- , [ID]
    -- , [DateReceived]
    -- , [Company]
    -- , [DateShipped]
    -- , [OrderID]
    -- , [HalfBbl]
    -- , [SixthBbl]
    -- , [SourceName]
    -- , [DestinationName]
    -- , [NoBillReason]
    -- , [OutCount]
    CREATE TABLE #mySHIPMENTS 
	(
		[myID] bigint IDENTITY(1,1) NOT NULL
		, [mySort] bigint NOT NULL
		, [DateAck] datetime NULL
        , [BrewerName] nvarchar(200) NULL
        , [DistributorName] nvarchar(200) NULL
		, [ID] bigint NOT NULL
        , [DateReceived] datetime NULL
        , [Company] bigint NULL
        , [DateShipped] datetime NULL
        , [OrderID] nvarchar(50) NULL
        , [HalfBbl] bigint NOT NULL DEFAULT ((0))
		, [SixthBbl] bigint NOT NULL DEFAULT ((0))
		, [SourceName] nvarchar(200) NULL
		, [DestinationName] nvarchar(200) NULL
		, [NoBillReason] nvarchar(300) NULL
		, [OutCount] bigint NULL
--		, [ParentOrderID] bigint NULL
--		, [Status] nvarchar(40) NULL
--		, [MoveType] nvarchar(100) NULL
--		, [Brewer] nvarchar(50) NULL
--		, [BrewerRevenue] money NOT NULL DEFAULT ((0))
--		, [Distributor] nvarchar(50) NULL
--		, [DistributorRevenue] money NOT NULL DEFAULT ((0))
--		, [Carrier] nvarchar(100) NULL
--		, [CarrierFee] money NULL
--		, [CarrierQuote] nvarchar(30) NULL
--		, [BadPallets] bigint NOT NULL DEFAULT ((0))
--		, [ForeignKegs] bigint NOT NULL DEFAULT ((0))
--		, [Pallets] bigint NOT NULL DEFAULT ((0))
--		, [HalfBblDistributor] bigint NOT NULL DEFAULT ((0))
--		, [SixthBblDistributor] bigint NOT NULL DEFAULT ((0))
--		, [PalletsDistributor] bigint NOT NULL DEFAULT ((0))
--		, [TentPallets] bigint NULL
--		, [TentHalfBbl] bigint NULL
--		, [TentSixthBbl] bigint NULL
--		, [DateVerified] datetime NULL
--		, [DateReported] datetime NULL
--		, [DateBilled] datetime NULL
--		, [DatePosted] datetime NULL
--		, [DateETA] datetime NULL
--		, [QuoteShipExpense1] money NOT NULL DEFAULT ((0))
--		, [QuoteShipName1] nvarchar(100) NULL
--		, [QuoteSelect1] nvarchar(3) NULL
--		, [QuoteQuoteNum1] nvarchar(30) NULL
--		, [QuoteShipExpense2] money NOT NULL DEFAULT ((0))
--		, [QuoteShipName2] nvarchar(100) NULL
--		, [QuoteSelect2] nvarchar(3) NULL
--		, [QuoteQuoteNum2] nvarchar(30) NULL
--		, [QuoteShipExpense3] money NOT NULL DEFAULT ((0))
--		, [QuoteShipName3] nvarchar(100) NULL
--		, [QuoteSelect3] nvarchar(3) NULL
--		, [QuoteQuoteNum3] nvarchar(30) NULL
--		, [ActualShipExpense] money NULL DEFAULT ((0))
--		, [RequireCustomer] nvarchar(3) NULL
--		, [IgnoreExp] SMALLINT NULL
--		, [ProNumber] nvarchar(40) NULL
--		, [BilledBy] bigint NULL
--		, [BilledBatchNum] nvarchar(50) NULL
--		, [LocalMove] nvarchar(3) NULL
--		, [MassUploadDate] datetime NULL
--		, [MassUploadKey] nvarchar(20) NULL
--		, [IgnoreError] SMALLINT NULL
--		, [Comments] text NULL
--		, [BrewerDynID] nvarchar(50) NULL
--		, [DistributorDynID] nvarchar(50) NULL
--		, [Verified] nvarchar(3) NULL
--		, [Source] nvarchar(50) NULL
--		, [Destination] nvarchar(50) NULL
--		, [ShipCostActual] money NULL
--		, [DepositCheck] nvarchar(3) NULL
--		, [OptExpediteFee] nvarchar(3) NULL
--		, [OptCustomDeposit] money NULL
--		, [MoveTypeComment] nvarchar(100) NULL
--		, [OptZeroDeposit] nvarchar(3) NULL
--		, [newkegs] nvarchar(3) NULL
--		, [revrelease] nvarchar(3) NULL
--		, [HalfBblProduct] nvarchar(50) NULL
--		, [SixthBblProduct] nvarchar(50) NULL
--		, [NoBill] nvarchar(3) NULL
--		, [NoBillAllow] nvarchar(3) NULL
--		, [locktrx] nvarchar(20) NULL
--		, [locktrxdate] datetime NULL
--		, [locktrxreason] nvarchar(300) NULL
--		, [PalletHeight] INT NULL DEFAULT ((83))
--		, [Weight] INT NULL DEFAULT ((0))
--		, [APIShipID] nvarchar(50) NULL
	)

    --Insert into #mySHIPMENTS
    --2018.08.29 only include the required fields
    -- [myID]
    -- , [mySort]
    --BEGIN ALL KNOWN SORT COLUMNS
    --, [DateAck]
    --, [BrewerName]
    --, [DistributorName]
    --, [DateShipped] --Already in main columns below
    --, [HalfBbl] --Already in main columns below
    --, [SixthBbl] --Already in main columns below
    --ENDALL KNOWN SORT COLUMNS
    -- , [ID]
    -- , [DateReceived]
    -- , [Company]
    -- , [DateShipped]
    -- , [OrderID]
    -- , [HalfBbl]
    -- , [SixthBbl]
    -- , [SourceName]
    -- , [DestinationName]
    -- , [NoBillReason]
    -- , [OutCount]
    --SET @mySQL = 'INSERT INTO #mySHIPMENTS([mySort], [ID], [ParentOrderID], [OrderID], [Company], [Status], [MoveType], [Brewer], [BrewerName], [BrewerRevenue], [Distributor], [DistributorName], [DistributorRevenue], [Carrier], [CarrierFee], [CarrierQuote], [BadPallets], [ForeignKegs], [Pallets], [HalfBbl], [SixthBbl], [HalfBblDistributor], [SixthBblDistributor], [PalletsDistributor], [TentPallets], [TentHalfBbl], [TentSixthBbl], [DateShipped], [DateVerified], [DateAck], [DateReceived], [DateReported], [DateBilled], [DatePosted], [DateETA], [QuoteShipExpense1], [QuoteShipName1], [QuoteSelect1], [QuoteQuoteNum1], [QuoteShipExpense2], [QuoteShipName2], [QuoteSelect2], [QuoteQuoteNum2], [QuoteShipExpense3], [QuoteShipName3], [QuoteSelect3], [QuoteQuoteNum3], [ActualShipExpense], [RequireCustomer], [IgnoreExp], [ProNumber], [BilledBy], [BilledBatchNum], [LocalMove], [MassUploadDate], [MassUploadKey], [IgnoreError], [Comments], [BrewerDynID], [DistributorDynID], [Verified], [Source], [SourceName], [Destination], [DestinationName], [ShipCostActual], [DepositCheck], [OptExpediteFee], [OptCustomDeposit], [MoveTypeComment], [OptZeroDeposit], [newkegs], [revrelease], [HalfBblProduct], [SixthBblProduct], [NoBill], [NoBillReason], [NoBillAllow], [locktrx], [locktrxdate], [locktrxreason], [PalletHeight], [Weight], [APIShipID], [OutCount]) '
    SET @mySQL = 'INSERT INTO #mySHIPMENTS([mySort], [DateAck], [BrewerName], [DistributorName], [ID], [DateReceived], [Company], [DateShipped], [OrderID], [HalfBbl], [SixthBbl], [SourceName], [DestinationName], [NoBillReason], [OutCount] ) '

    --SET @mySelectSQL = 'SELECT ROW_NUMBER() OVER(ORDER BY ' + Replace(@pOrderBy, '''', '''''') + ') AS [mySort], [ID], [ParentOrderID], [OrderID], [Company], [Status], [MoveType], [Brewer], [BrewerName], [BrewerRevenue], [Distributor], [DistributorName], [DistributorRevenue], [Carrier], [CarrierFee], [CarrierQuote], [BadPallets], [ForeignKegs], [Pallets], [HalfBbl], [SixthBbl], [HalfBblDistributor], [SixthBblDistributor], [PalletsDistributor], [TentPallets], [TentHalfBbl], [TentSixthBbl], [DateShipped], [DateVerified], [DateAck], [DateReceived], [DateReported], [DateBilled], [DatePosted], [DateETA], [QuoteShipExpense1], [QuoteShipName1], [QuoteSelect1], [QuoteQuoteNum1], [QuoteShipExpense2], [QuoteShipName2], [QuoteSelect2], [QuoteQuoteNum2], [QuoteShipExpense3], [QuoteShipName3], [QuoteSelect3], [QuoteQuoteNum3], [ActualShipExpense], [RequireCustomer], [IgnoreExp], [ProNumber], [BilledBy], [BilledBatchNum], [LocalMove], [MassUploadDate], [MassUploadKey], [IgnoreError], [Comments], [BrewerDynID], [DistributorDynID], [Verified], [Source], [SourceName], [Destination], [DestinationName], [ShipCostActual], [DepositCheck], [OptExpediteFee], [OptCustomDeposit], [MoveTypeComment], [OptZeroDeposit], [newkegs], [revrelease], [HalfBblProduct], [SixthBblProduct], [NoBill], [NoBillReason], [NoBillAllow], [locktrx], [locktrxdate], [locktrxreason], [PalletHeight], [Weight], [APIShipID] '
    SET @mySelectSQL = 'SELECT ROW_NUMBER() OVER(ORDER BY ' + Replace(@pOrderBy, '''', '''''') + ') AS [mySort], [DateAck], [BrewerName], [DistributorName], [ID], [DateReceived], [Company], [DateShipped], [OrderID], [HalfBbl], [SixthBbl], [SourceName], [DestinationName], [NoBillReason] '
    IF @myContentType = Lower('PreBillAll')
    BEGIN
        SET @mySQL = @mySQL + @mySelectSQL + ', Cast(0 As bigint) As [OutCount] ' 
    END
    ELSE
    BEGIN
        SET @mySQL = @mySQL + @mySelectSQL + ', Cast((SELECT Count([ID]) FROM [dbo].[SECLOG] WHERE [dbo].[SECLOG].[Description] = ''Shipment Updated to Acknowledged by Brewer'' AND [ShipID] = SHIPMENTS.ID) As bigint) As [OutCount] ' 
    END
    SET @mySQL = @mySQL + 'FROM [dbo].[SHIPMENTS] '
    
    --Debug
    IF @pDebugOn = 1
    BEGIN
        Select @mySQL as [@mySQLPART001]
    END
    --Debug
 
    --Build the Where Clause for Different Content Types
    IF @myContentType = Lower('NoBill')
    BEGIN
        SET @mySQL = @mySQL + 'WHERE [Status] = ''Completed'' AND [NoBill] = ''yes'' '
    END
    ELSE IF @myContentType = Lower('CustTypeBrewers')
    BEGIN
        SET @mySQL = @mySQL + 'WHERE [MoveType] = ''Distributor to Brewer'' AND [Status] = ''Completed'' '
    END   
    ELSE IF @myContentType = Lower('CustTypeDistributors')
    BEGIN
        SET @mySQL = @mySQL + 'WHERE [MoveType] = ''Brewer to Distributor'' AND [Status] = ''Completed'' '
    END
    ELSE IF @myContentType = Lower('Rebill')
    BEGIN
        SET @mySQL = @mySQL + 'WHERE [Status] = ''Re-bill'' '
    END
    ELSE IF @myContentType = Lower('Hold')
    BEGIN
        SET @mySQL = @mySQL + 'WHERE [Status] = ''Hold'' '
    END
    ELSE IF @myContentType = Lower('SelfInput')
    BEGIN
        SET @mySQL = @mySQL + 'WHERE [Status] = ''Completed - Brewer Input'' '
    END
    ELSE IF @myContentType = Lower('CompletedAll')
    BEGIN
        SET @mySQL = @mySQL + 'WHERE [Status] = ''Completed'' '
    END
    ELSE IF @myContentType = Lower('CompletedAllTraffic')
    BEGIN
        SET @mySQL = @mySQL + 'WHERE [Status] = ''Completed'' AND [MassUploadKey] Is NULL '
    END
    ELSE IF @myContentType = Lower('PreBillAll')
    BEGIN
        SET @mySQL = @mySQL + 'WHERE [Status] = ''Pre-Billed'' '
    END
    ELSE IF @myContentType = Lower('IDLimit') And @pIDLimitBrewer Is Not Null
    BEGIN
        SET @mySQL = @mySQL + 'WHERE [Brewer] = ''' + Cast(RTrim(LTrim(Replace(@pIDLimitBrewer, '''',''''''))) As varchar(50)) + ''' AND [Status] = ''Completed'' '
    END
    ELSE IF @myContentType = Lower('IDLimit') And @pIDLimitBrewer Is Null
    BEGIN
        SET @mySQL = @mySQL + 'WHERE [Brewer] = '''' AND [Status] = ''Completed'' '
    END
    ELSE IF @myContentType = Lower('NameLimit') And @pNameLimitBrewerName Is Not Null
    BEGIN
        SET @mySQL = @mySQL + 'WHERE [BrewerName] like ''%%' + Cast(RTrim(LTrim(Replace(@pNameLimitBrewerName, '''',''''''))) As varchar(200)) + '%%'' AND [Status] = ''Completed'' '
    END
    ELSE IF @myContentType = Lower('NameLimit') And @pNameLimitBrewerName Is Null
    BEGIN
        SET @mySQL = @mySQL + 'WHERE [BrewerName] like ''%%%%'' AND [Status] = ''Completed'' '
    END
    ELSE IF @myContentType = Lower('Test')
    BEGIN
        SET @mySQL = @mySQL + 'WHERE [BrewerName] like ''%%test%%brew%%'' '
    END
    ELSE IF @myContentType = Lower('MassUpload') And @pMassUploadKey Is Not Null
    BEGIN
        SET @mySQL = @mySQL + 'WHERE [MassUploadKey] = ''' + Cast(RTrim(LTrim(Replace(@pMassUploadKey, '''',''''''))) As varchar(200)) + ''' AND Status = ''Completed'' '
    END
    ELSE IF @myContentType = Lower('MassUpload') And @pMassUploadKey Is Null
    BEGIN
        SET @mySQL = @mySQL + 'WHERE [MassUploadKey] = '''' AND Status = ''Completed'' '
    END
    ELSE 
    BEGIN
        --Default 
        SET @mySQL = @mySQL + 'WHERE [MoveType] <> ''Brewer to Distributor'' AND [MoveType] <> ''Distributor to Brewer'' AND [Status] = ''Completed'' '
    END
    
    --Debug
    IF @pDebugOn = 1
    BEGIN
        Select @mySQL as [@mySQLPART002]
    END
    --Debug

    IF NULLIF(@pCompany, -1) IS NOT NULL
    BEGIN
        IF @pCompany > 0
        BEGIN
            SET @mySQL = @mySQL + ' AND [Company] = ' + Replace(Cast(@pCompany As varchar(20)), '''', '''''') + ' '
        END
    END
    --Order By - moved to ROW_NUMBER() OVER ^^^^ABOVE^^^^
    --SET @mySQL = @mySQL + ' ORDER BY ' + Replace(@pOrderBy, '''', '''''') + ' '
    
    --Debug
    IF @pDebugOn = 1
    BEGIN
        Select @mySQL as [@mySQLPART003Final]
    END
    --Debug
    
    --INSERT INTO TEMP
    BEGIN TRANSACTION
        EXEC(@mySQL)
    COMMIT;
    
    IF @pDebugOn = 1
    BEGIN
        --Select
        --2018.08.29 only include the required fields
        -- [myID]
        -- , [mySort]
        --BEGIN ALL KNOWN SORT COLUMNS
        --, [DateAck]
        --, [BrewerName]
        --, [DistributorName]
        --, [DateShipped] --Already in main columns below
        --, [HalfBbl] --Already in main columns below
        --, [SixthBbl] --Already in main columns below
        --ENDALL KNOWN SORT COLUMNS
        -- , [ID]
        -- , [DateReceived]
        -- , [Company]
        -- , [DateShipped]
        -- , [OrderID]
        -- , [HalfBbl]
        -- , [SixthBbl]
        -- , [SourceName]
        -- , [DestinationName]
        -- , [NoBillReason]
        -- , [OutCount]
        --SELECT [myID], [ID], [ParentOrderID], [OrderID], [Company], [Status], [MoveType], [Brewer], [BrewerName], [BrewerRevenue], [Distributor], [DistributorName], [DistributorRevenue], [Carrier], [CarrierFee], [CarrierQuote], [BadPallets], [ForeignKegs], [Pallets], [HalfBbl], [SixthBbl], [HalfBblDistributor], [SixthBblDistributor], [PalletsDistributor], [TentPallets], [TentHalfBbl], [TentSixthBbl], [DateShipped], [DateVerified], [DateAck], [DateReceived], [DateReported], [DateBilled], [DatePosted], [DateETA], [QuoteShipExpense1], [QuoteShipName1], [QuoteSelect1], [QuoteQuoteNum1], [QuoteShipExpense2], [QuoteShipName2], [QuoteSelect2], [QuoteQuoteNum2], [QuoteShipExpense3], [QuoteShipName3], [QuoteSelect3], [QuoteQuoteNum3], [ActualShipExpense], [RequireCustomer], [IgnoreExp], [ProNumber], [BilledBy], [BilledBatchNum], [LocalMove], [MassUploadDate], [MassUploadKey], [IgnoreError], [Comments], [BrewerDynID], [DistributorDynID], [Verified], [Source], [SourceName], [Destination], [DestinationName], [ShipCostActual], [DepositCheck], [OptExpediteFee], [OptCustomDeposit], [MoveTypeComment], [OptZeroDeposit], [newkegs], [revrelease], [HalfBblProduct], [SixthBblProduct], [NoBill], [NoBillReason], [NoBillAllow], [locktrx], [locktrxdate], [locktrxreason], [PalletHeight], [Weight], [APIShipID], [OutCount]
        SELECT [myID]
        , [mySort]
        , [DateAck]
        , [BrewerName]
        , [DistributorName]
        , [ID]
        , [DateReceived]
        , [Company]
        , [DateShipped]
        , [OrderID]
        , [HalfBbl]
        , [SixthBbl]
        , [SourceName]
        , [DestinationName]
        , [NoBillReason]
        , [OutCount]
        FROM #mySHIPMENTS
        Order By [mySort]
    END
    
    --Create The Temp HTML Output Table
    DECLARE @myHtmlOutPut as varchar(80)
    SET @myHtmlOutPut = @pTapGUID
    SET @mySQL = 'CREATE TABLE ' + @myHtmlOutPut + ' ' 
	SET @mySQL = @mySQL + '( '
    SET @mySQL = @mySQL + '[myID] bigint IDENTITY(1,1) NOT NULL '
    SET @mySQL = @mySQL + ', [mySort] bigint NOT NULL '
    SET @mySQL = @mySQL + ', [BOContentHtml] varchar(7000) NULL '
    SET @mySQL = @mySQL + ') '
    
    BEGIN TRANSACTION
        EXEC(@mySQL)
    COMMIT;
    
    CREATE TABLE #myHtmlOutPut
    (     
        [myID] bigint IDENTITY(1,1) NOT NULL
        , [mySort] bigint NOT NULL
        , [BOContentHtml] varchar(7000) NULL 
    ) 
    
    --Build Template HTML String
    DECLARE @Htmplt as varchar(3000)
    DECLARE @defaultTemplate as varchar(3000)
    SET @defaultTemplate = ''
    SET @defaultTemplate = @defaultTemplate + '<tr onMouseOver="this.bgColor=''#FFFF99'';" onMouseOut="this.bgColor=''#f2f2f2'';"> '
        SET @defaultTemplate = @defaultTemplate + '<td><input type="checkbox" id="checkbox_{strCheckboxRowIndex}" name="billitem" value="{strID}"></td> '
        SET @defaultTemplate = @defaultTemplate + '<td style="text-align: left; cursor: pointer;" nowrap onClick="window.parent.passit4(''{strID}'', ''{strOnClickPassit4Region}'', ''{strOnClickPassit4Tab}'');" onmouseover="parent.ajax_showTooltip(''billing-preview.asp?ID={strID}'', this);" onmouseout="parent.ajax_hideTooltip();"> '
            SET @defaultTemplate = @defaultTemplate + '{strReturnNameNameDateShipped} '
        SET @defaultTemplate = @defaultTemplate + '</td> '
        SET @defaultTemplate = @defaultTemplate + '<td>{strFormatNumber}</td> '
        SET @defaultTemplate = @defaultTemplate + '<td class="{strClass}" id="{strCheckboxRowIndex}">{strDateReceived}</td> '
        SET @defaultTemplate = @defaultTemplate + '<td>{strSourceName}</td> '
        SET @defaultTemplate = @defaultTemplate + '<td>{strDestinationName}</td> '
        SET @defaultTemplate = @defaultTemplate + '<td>{strOutCountHtml}</td> '
        SET @defaultTemplate = @defaultTemplate + '<td><img src="/img/calendar.gif" border="0" alt="Invoice Summary" onClick="window.parent.passit7(''{strID}'');"></td> '
        SET @defaultTemplate = @defaultTemplate + '<td>{strIDBOLCheck}</td> '
    SET @defaultTemplate = @defaultTemplate + '</tr> '
    SET @defaultTemplate = @defaultTemplate + '{strNoBillReasonHtml}' 
    
    IF @pTAPAppVersion = 'TAP2.5'
    BEGIN
        SET @Htmplt = @defaultTemplate
    END
    ELSE
    BEGIN
        SET @Htmplt = @defaultTemplate
    END
    --------------------------------------------------------------------------------------------------------------------------
    --BEGIN CURSOR GET the BO records where the OrderID is being used as the ParentID within BREWERORDERSCOLLARS
    --------------------------------------------------------------------------------------------------------------------------
        --@HTML Current ROW
            DECLARE @HtmlRow as varchar(3000)
            DECLARE @mySort as bigint
            DECLARE @BOContentHtmlInsert as varchar(7000)
        --Declare HTML Replace Vars
            DECLARE @iCheckboxRowIndex as int
            DECLARE @strMonth as varchar(2)
            DECLARE @strYear as varchar(2)
            DECLARE @strID as varchar(20)
            DECLARE @strOnClickPassit4Region as varchar(50)		    
            DECLARE @strOnClickPassit4Tab as varchar(50)	    
            DECLARE @strReturnNameNameDateShipped as varchar(50)
            DECLARE @strFormatNumber as varchar(50)
            DECLARE @strClass as varchar(20)
            DECLARE @strDateReceived as varchar(50)
            DECLARE @strSourceName as varchar(200)
            DECLARE @strDestinationName as varchar(200)
            DECLARE @strOutCountHtml as varchar(1000)
            DECLARE @strIDBOLCheck as varchar(1000)
            DECLARE @strNoBillReasonHtml as varchar(1000)
        --Declare Required Vars from #mySHIPMENTS
            DECLARE @ID as bigint
            DECLARE @DateReceived as datetime
            DECLARE @Company as bigint
            DECLARE @DateShipped as datetime
            DECLARE @OrderID as nvarchar(50)
            DECLARE @HalfBbl as bigint
            DECLARE @SixthBbl as bigint
            DECLARE @SourceName as nvarchar(200)
            DECLARE @DestinationName as nvarchar(200)
            DECLARE @NoBillReason as nvarchar(300)
            DECLARE @OutCount as bigint
        Declare CurSHIPMENTS Cursor for
            SELECT  S.[ID]
                , S.[DateReceived]
                , S.[Company]
                , S.[DateShipped]
                , S.[OrderID]
                , S.[HalfBbl]
                , S.[SixthBbl]
                , S.[SourceName]
                , S.[DestinationName]
                , S.[NoBillReason]
                , S.[OutCount]
            FROM #mySHIPMENTS S
            Order By S.[mySort]
    --------------------------------------------------------------------------------------------------------------------------
    --END CURSOR GET the BO records where the OrderID is being used as the ParentID within BREWERORDERSCOLLARS
    --------------------------------------------------------------------------------------------------------------------------
    --Init some vars
    BEGIN
        SET @iCheckboxRowIndex = 0
        SET @strMonth = ''
        SET @strYear = ''
        SET @mySort = 0
        SET @BOContentHtmlInsert = ''
    END
    Open CurSHIPMENTS
        Fetch Next From CurSHIPMENTS into
            @ID
            , @DateReceived
            , @Company
            , @DateShipped
            , @OrderID
            , @HalfBbl
            , @SixthBbl
            , @SourceName
            , @DestinationName
            , @NoBillReason
            , @OutCount
		While @@FEtch_Status = 0		
		BEGIN
		    --------------------------------------------------------------------------------------------------------------------------
            --BEGIN Replace the Template
            --------------------------------------------------------------------------------------------------------------------------
		    --Replace the Template
		    SET @HtmlRow = @Htmplt
            --1. {strCheckboxRowIndex}
                SET @HtmlRow = Replace(@HtmlRow, '{strCheckboxRowIndex}', Cast(@iCheckboxRowIndex as varchar(20)))
            --2. {strID}
                SET @strID = Cast(@ID as varchar(20))
                SET @HtmlRow = Replace(@HtmlRow, '{strID}', @strID)
            --3. {strOnClickPassit4Region}
                SET @strOnClickPassit4Region = Replace(@pRegion, '''', '''''')
                SET @HtmlRow = Replace(@HtmlRow, '{strOnClickPassit4Region}', @strOnClickPassit4Region)
            --4. {strOnClickPassit4Tab}
                SET @strOnClickPassit4Tab = Replace(@pTab, '''', '''''')
                SET @HtmlRow = Replace(@HtmlRow, '{strOnClickPassit4Tab}', @strOnClickPassit4Tab)
            --5. {strReturnNameNameDateShipped}
                SET @strReturnNameNameDateShipped = (SELECT TOP 1 [Name] From dbo.[COMPANIES] WHERE [ID] = @Company) + '-S-'
                IF NULLIF(@DateShipped, '') Is Not Null
                BEGIN
                    SET @strMonth = (SELECT RIGHT('0' + RTRIM(MONTH(@DateShipped)), 2))
                    SET @strYear = (SELECT RIGHT(RTRIM(YEAR(@DateShipped)), 2))
                END
                ELSE
                BEGIN
                    SET @strMonth = ''
                    SET @strYear = ''
                END
                SET @strReturnNameNameDateShipped = @strReturnNameNameDateShipped + @strMonth + @strYear + '-' + @OrderID
                SET @HtmlRow = Replace(@HtmlRow, '{strReturnNameNameDateShipped}', @strReturnNameNameDateShipped)
            --6. {strFormatNumber}
                SET @strFormatNumber = Cast(@HalfBbl as varchar(20)) + ' / ' + Cast(@SixthBbl as varchar(20))
                SET @HtmlRow = Replace(@HtmlRow, '{strFormatNumber}', @strFormatNumber)
            --7. {strClass}
                IF NULLIF(@DateReceived, '') Is Not Null
                BEGIN
                    SET @strDateReceived = (SELECT LTRIM(STR(MONTH(@DateReceived))) + '/' + LTRIM(STR(DAY(@DateReceived))) + '/' + STR(YEAR(@DateReceived), 4))
                    IF MONTH(GETDATE()) = MONTH(@DateReceived) 
                    BEGIN
                        SET @strClass = 'CurMonth'
                    END
                    ELSE
                    BEGIN
                        SET @strClass = 'NotCurMonth'
                    END
                END
                ELSE
                BEGIN
                    SET @strDateReceived = ''
                    SET @strClass = 'NotCurMonth'
                END
                SET @HtmlRow = Replace(@HtmlRow, '{strClass}', @strClass)
            --8. {strDateReceived}
                SET @HtmlRow = Replace(@HtmlRow, '{strDateReceived}', @strDateReceived)
            --9. {strSourceName}
                SET @strSourceName = @SourceName
                SET @HtmlRow = Replace(@HtmlRow, '{strSourceName}', @strSourceName)
            --10. {strDestinationName}
                SET @strDestinationName = @DestinationName
                SET @HtmlRow = Replace(@HtmlRow, '{strDestinationName}', @strDestinationName)
            --11. {strOutCountHtml}
                IF @OutCount > 0
                BEGIN
                    SET @strOutCountHtml = '<img src="/img/check.gif" alt="Verified Online by Brewer">'
                END
                ELSE
                BEGIN
                    SET @strOutCountHtml = '&nbsp;'
                END
                SET @HtmlRow = Replace(@HtmlRow, '{strOutCountHtml}', @strOutCountHtml)
            --12. {strIDBOLCheck}
                SET @strIDBOLCheck = (SELECT TOP 1 Cast([ID] As varchar(20)) As [IDBOLCheck] FROM FILES WHERE [OrderID] = @ID ORDER BY [FileDate])
                IF NULLIF(@strIDBOLCheck, '') Is Not Null
                BEGIN
                    SET @strIDBOLCheck = '<a href="reportShipmentDocs.asp?Download=yes&ID=' + Replace(@strIDBOLCheck, '''', '''''') +  '"><img src="/img/download.gif" border="0"></a>'
                END
                ELSE
                BEGIN
                    SET @strIDBOLCheck = '<img src=/img/minus.gif border=0 alt="Bill of Lading">'
                END
                SET @HtmlRow = Replace(@HtmlRow, '{strIDBOLCheck}', @strIDBOLCheck)
            --13. {strNoBillReasonHtml}
                SET @strNoBillReasonHtml = ''
                IF NULLIF(@NoBillReason, '') Is Not Null
                BEGIN
                    IF Len(@NoBillReason) > 0
                    BEGIN
                        SET @strNoBillReasonHtml = @strNoBillReasonHtml + '<tr> ' 
                        SET @strNoBillReasonHtml = @strNoBillReasonHtml + '<td>&nbsp;</td> '
                        SET @strNoBillReasonHtml = @strNoBillReasonHtml + '<td colspan="8" style="text-align: left;"><font color="red">' + Replace(@NoBillReason, '''', '''''') + '</font></td> '
                        SET @strNoBillReasonHtml = @strNoBillReasonHtml + '</tr> '
                    END
                END
                SET @HtmlRow = Replace(@HtmlRow, '{strNoBillReasonHtml}', @strNoBillReasonHtml)
            --------------------------------------------------------------------------------------------------------------------------
            --END Replace the Template
            --------------------------------------------------------------------------------------------------------------------------
            
            ------------------------------------------------------------------------------------------------------BUILD the HTML INSERT
            IF Len(@BOContentHtmlInsert) + Len(@HtmlRow) <= 7000
            BEGIN
                SET @BOContentHtmlInsert = @BOContentHtmlInsert + @HtmlRow
            END
            ELSE
            BEGIN
                SET @mySort = @mySort + 1
                IF @pDebugOn = 1
                BEGIN
                    SELECT @mySort As [@mySort] , @BOContentHtmlInsert As [@BOContentHtmlInsert]
                END
                BEGIN TRANSACTION
                    INSERT INTO #myHtmlOutPut ([mySort], [BOContentHtml]) Values(@mySort, @BOContentHtmlInsert)
                COMMIT;
                --RESET
                SET @BOContentHtmlInsert = @HtmlRow
            END
            
            --Increment
            BEGIN
                SET @iCheckboxRowIndex = @iCheckboxRowIndex + 1
            END
            Fetch Next From CurSHIPMENTS into
                @ID
                , @DateReceived
                , @Company
                , @DateShipped
                , @OrderID
                , @HalfBbl
                , @SixthBbl
                , @SourceName
                , @DestinationName
                , @NoBillReason
                , @OutCount
		END
		
		Close CurSHIPMENTS
		Deallocate CurSHIPMENTS
		
		--Final Insert
		IF Len(@BOContentHtmlInsert) > 0
		BEGIN
            SET @mySort = @mySort + 1
            IF @pDebugOn = 1
            BEGIN
                SELECT @mySort As [@mySort] , @BOContentHtmlInsert As [@BOContentHtmlInsert]
            END
            BEGIN TRANSACTION
                INSERT INTO #myHtmlOutPut ([mySort], [BOContentHtml]) Values(@mySort, @BOContentHtmlInsert)
            COMMIT;
		END
        
        --Finally Output the HTML
        SET @mySQL = 'INSERT INTO dbo.[' + Replace(@myHtmlOutPut, '''', '''''') + '] ([mySort], [BOContentHtml]) Select [mySort], [BOContentHtml] From #myHtmlOutPut Order By [mySort] '
         IF @pDebugOn = 1
        BEGIN
            SELECT @mySQL As [@mySQL_INSERT INTO_GUID]
        END 
        BEGIN TRANSACTION
            EXEC(@mySQL)
        COMMIT;
    
    --Cleanup
    Drop Table #myHtmlOutPut
    Drop Table #mySHIPMENTS
END
GO
