
/**********************************************************************************************
*       Jira:           2018.09.07 TAP2-45 Upload Shipment Ack [brewPortal.asp]-ImportFile
*                       https://microstartap3.atlassian.net/browse/TAP2-45
*    	Description:	TAP Insert Into tmpMassUploadShipAck
*		Returns:        Creates Create pseudo Temp Table @pTapGUID
*                       Insert Rows of MAX LEN 7000	chars into @pTapGUID.[ShipAckContentHtml]
*       Note:                                     
*	
*	Author: 	Brad Skidmore
*	Date: 		9/26/2018
**********************************************************************************************/
CREATE PROCEDURE [dbo].[tap_spInstmpMassUploadShipAck] 
@pTAPAppVersion varchar(50)='TAP2.5',   --1. Versioning of HTML output
@pTapGUID nvarchar(80)=null,            --2. The GUID [without dashes] Table Name for creation of a pseudo temp table that will contain html rows.
@pMassUploadShipAckKey nvarchar(40),    --3. e.g. eTPpbAzLTWFpV8U-XFd4TlGuxR62rCgyL77L prefix is tmp directory files: eTPpbAzLTWFpV8U-shipment-ack.txt
@pOrderIDListForInsert varchar(2000),   --4. CSV List of SHIPMENTS.[OrderID] '123456789', '987654321', '555555555' 
@pSelectRowsForInsert varchar(8000),    --5. CSV ['OrderID', HalfBbl, SixthBbl, Pallets, BadPallets, ForeignKegs, 'DateReceived']
@pInsertPreviewONLY bit=0,              --6. WHEN @pInsertPreviewONLY = 1 then DO NOT UPDATE SHIPMENTS! DO NOT Create Display. ONLY insert a batch of rows into tmpMassUploadShipAck with @pMassUploadShipAckKey
@pPreviewONLY bit=0,                    --7. WHEN @pPreviewONLY = 1 then Display what was successfully inserted into tmpMassUploadShipAck for @pMassUploadShipAckKey
@pDoUpdateShipAck bit=0,                --8. WHEN @pDoUpdateShipAck = 1 then That means the Inserts for Preview and Preview all past and SHIPMENTS Can Be updated from what is in tmpMassUploadShipAck for @pMassUploadShipAckKey
@pAdjUser nvarchar(20),                 --9. The User Logged on to Tap2
@pDebugOn bit=0                         --10. Debugging?  SET to 1 IF NOT SET to 0  Will return a robust set of queries in order to interrogate/reconcile the sp results.
AS
BEGIN
    --------------------------------------------------------------------------------------------------
    --BEGIN Validate Params:                                                     BEGIN Validate Params
    -------------------------------------------------------------------------------------------------- 
    BEGIN
        IF NULLIF(@pTAPAppVersion, '') Is Null
        BEGIN
            SET @pTAPAppVersion = 'TAP2.5'
        END
        
        IF NULLIF(@pTapGUID, '') Is Null
        BEGIN
            SET @pTapGUID = 'tap_spInstmpMassUploadShipAck_' + Cast(newid() as nvarchar(36))
        END
        
        IF NULLIF(@pMassUploadShipAckKey, '') Is Null
        BEGIN
            SET @pMassUploadShipAckKey = @pTapGUID
        END
        
        IF NULLIF(Cast(@pOrderIDListForInsert as varchar(8000)), '') Is Null
        BEGIN
            SET @pOrderIDListForInsert = ''
        END
        
        IF NULLIF(Cast(@pSelectRowsForInsert as varchar(8000)), '') Is Null
        BEGIN
            SET @pSelectRowsForInsert = ''
        END
        
        IF NULLIF(@pInsertPreviewONLY, null) Is Null
        BEGIN
            SET @pInsertPreviewONLY = 0
        END
        
        IF NULLIF(@pPreviewONLY, null) Is Null
        BEGIN
            SET @pPreviewONLY = 0
        END
        
        IF NULLIF(@pDoUpdateShipAck, null) Is Null
        BEGIN
            SET @pDoUpdateShipAck = 0
        END
        
        IF NULLIF(@pAdjUser, '') Is Null
        BEGIN
            SET @pAdjUser = 'Unknown'
        END
        
        IF NULLIF(@pDebugOn, null) Is Null
        BEGIN
            SET @pDebugOn = 0
        END
    END
    --------------------------------------------------------------------------------------------------
    --END Validate Params:                                                         END Validate Params
    -------------------------------------------------------------------------------------------------- 
    
    --------------------------------------------------------------------------------------------------
    --BEGIN Required fields                                                      BEGIN Required fields
    --------------------------------------------------------------------------------------------------
    --**************************
    -- [myID]
    -- , [mySort]
    --**************************
    -- , [ID] 
    -- , [OrderID]
    -- , [Brewer]
    -- , [BrewerName]
    -- , [Distributor]
    -- , [DistributorName]
    -- , [HalfBbl]
    -- , [TentHalfBbl]
    -- , [SixthBbl]
    -- , [TentSixthBbl]
    -- , [Pallets]
    -- , [TentPallets]
    -- , [BadPallets]
    -- , [ForeignKegs]
    -- , [DateShipped]
    -- , [DateReceived]
    -- , [MassUploadShipAckKey]
    DECLARE @ID bigint   
    DECLARE @myID bigint
    DECLARE @mySort bigint
    DECLARE @ShipID bigint                 	--SHIPMENTS.[ID]    
    DECLARE @OrderID nvarchar(50)           --SHIPMENTS.[OrderID]
    DECLARE @Brewer nvarchar(50)            --SHIPMENTS.[Brewer]
    DECLARE @BrewerName nvarchar(200)       --SHIPMENTS.[BrewerName]
    DECLARE @Distributor nvarchar(200)      --SHIPMENTS.[Distributor]
    DECLARE @DistributorName nvarchar(200)  --SHIPMENTS.[DistributorName]
    DECLARE @HalfBbl bigint                 --SHIPMENTS.[HalfBbl]
    DECLARE @TentHalfBbl bigint             --SHIPMENTS.[TentHalfBbl]
    DECLARE @SixthBbl bigint                --SHIPMENTS.[SixthBbl]
    DECLARE @TentSixthBbl bigint            --SHIPMENTS.[TentSixthBbl]
    DECLARE @Pallets bigint                 --SHIPMENTS.[Pallets]
    DECLARE @TentPallets bigint             --SHIPMENTS.[TentPallets]
    DECLARE @BadPallets bigint              --SHIPMENTS.[BadPallets]
    DECLARE @ForeignKegs bigint             --SHIPMENTS.[ForeignKegs]
    DECLARE @DateShipped datetime           --SHIPMENTS.[DateShipped]
    DECLARE @DateReceived datetime          --SHIPMENTS.[DateReceived]
    DECLARE @MassUploadShipAckKey nvarchar(40)
    
    DECLARE @mySQL varchar(8000)
    --------------------------------------------------------------------------------------------------
    --END Required fields                                                          END Required fields
    -------------------------------------------------------------------------------------------------- 
     
    --Debug
    IF @pDebugOn = 1
    BEGIN
        --List Current Params
        SELECT 
            @pTAPAppVersion As [@pTAPAppVersion]
            , @pTapGUID As [@pTapGUID]
            , @pMassUploadShipAckKey As [@pMassUploadShipAckKey]
            , @pOrderIDListForInsert As [@pOrderIDListForInsert]
            , @pSelectRowsForInsert As [@pSelectRowsForInsert]
            , @pInsertPreviewONLY AS [@pInsertPreviewONLY]
            , @pPreviewONLY As [@pPreviewONLY]
            , @pDoUpdateShipAck As [@pDoUpdateShipAck]
            , @pAdjUser As [@pAdjUser]
            , @pDebugOn As [@pDebugOn]
    END
    --Debug
    	
	----------------------------------------------------------------------------------------------------------------------------------------------------------------
	--InsertPreviewONLY -- BEGIN                                                                                                          InsertPreviewONLY -- BEGIN
	----------------------------------------------------------------------------------------------------------------------------------------------------------------
	IF @pInsertPreviewONLY = 1 AND @pPreviewONLY = 0 AND @pDoUpdateShipAck = 0
	BEGIN
		--2018.09.18 used for Looking up Tentative values and Names for pre insertion validation
		IF OBJECT_ID('tempdb..[#myOrderID]') IS NULL
		BEGIN
		    CREATE TABLE #myOrderID	
		    (
		        [ID] bigint IDENTITY(1,1) NOT NULL
		        , [ShipID] bigint NOT NULL
		        , [OrderID] nvarchar(50) NULL
		        , [Processed] bit NULL
            )
		END
        IF OBJECT_ID('tempdb..[#mySHIPLU]') IS NULL
        BEGIN
            CREATE TABLE #mySHIPLU 
            (
                 [ShipID]                  	bigint NOT NULL 
                , [OrderID]             	nvarchar(50) NULL
                , [Brewer]              	nvarchar(50) NULL
                , [BrewerName]          	nvarchar(200) NULL
                , [Distributor]         	nvarchar(200) NULL
                , [DistributorName]     	nvarchar(200) NULL
                , [HalfBbl]             	bigint NULL DEFAULT ((0))
                , [TentHalfBbl]         	bigint NULL
                , [SixthBbl]            	bigint NULL  DEFAULT ((0))
                , [TentSixthBbl]        	bigint NULL
                , [Pallets]             	bigint NULL  DEFAULT ((0))
                , [TentPallets]         	bigint NULL
                , [DateShipped]         	datetime NULL
            )
        END
        IF OBJECT_ID('tempdb..[#mySHIPACKS]') IS NULL
        BEGIN
            CREATE TABLE #mySHIPACKS 
            (
                [myID]                      bigint IDENTITY(1,1) NOT NULL
                , [mySort]                  bigint NOT NULL
                , [ShipID]                  bigint NULL 
                , [OrderID]             	nvarchar(50) NULL
                , [Brewer]              	nvarchar(50) NULL
                , [BrewerName]          	nvarchar(200) NULL
                , [Distributor]         	nvarchar(200) NULL
                , [DistributorName]     	nvarchar(200) NULL
                , [HalfBbl]             	bigint NULL  DEFAULT ((0))
                , [TentHalfBbl]         	bigint NULL  DEFAULT ((0))
                , [SixthBbl]            	bigint NULL  DEFAULT ((0))
                , [TentSixthBbl]        	bigint NULL  DEFAULT ((0))
                , [Pallets]             	bigint NULL  DEFAULT ((0))
                , [TentPallets]         	bigint NULL  DEFAULT ((0))
                , [BadPallets]          	bigint NULL  DEFAULT ((0))
                , [ForeignKegs]         	bigint NULL  DEFAULT ((0))
                , [DateShipped]         	datetime NULL
                , [DateReceived]        	datetime NULL
                , [MassUploadShipAckKey]	nvarchar(40) NULL 
            )
	    END
	    ------------------------------------------------------------------------------------------------------
	    --BEGIN Insert into #myOrderID                                            BEGIN Insert into #myOrderID
	    ------------------------------------------------------------------------------------------------------
	    BEGIN
            SET @mySQL = 'INSERT INTO #myOrderID([ShipID], [OrderID]) '
            SET @mySQL = @mySQL + 'SELECT [ID], [OrderID] '
            SET @mySQL = @mySQL + 'FROM [dbo].[SHIPMENTS] '
            SET @mySQL = @mySQL + 'WHERE [OrderID] IN (' + @pOrderIDListForInsert + ') '            
            
            IF @pDebugOn = 1
            BEGIN
                Select @mySQL As [@mySQL]
            END
            
            BEGIN TRANSACTION INSERTmyOrderID
                EXEC(@mySQL)
            COMMIT TRANSACTION INSERTmyOrderID
        END                                      
        ------------------------------------------------------------------------------------------------------
	    --END Insert into #myOrderID                                                END Insert into #myOrderID
	    ------------------------------------------------------------------------------------------------------
        
        --Insert into the Lookup Table
        BEGIN TRANSACTION INSERTmySHIPLU
            INSERT INTO #mySHIPLU([ShipID], [OrderID], [Brewer], [BrewerName], [Distributor], [DistributorName], [HalfBbl], [TentHalfBbl], [SixthBbl], [TentSixthBbl], [Pallets], [TentPallets], [DateShipped]) 
            SELECT [ID] As [ShipID], [OrderID], [Brewer], [BrewerName], [Distributor], [DistributorName], [HalfBbl], [TentHalfBbl], [SixthBbl], [TentSixthBbl], [Pallets], [TentPallets], [DateShipped] 
            FROM [dbo].[SHIPMENTS] 
            WHERE [ID] In (SELECT [ShipID] FROM #myOrderID WHERE [Processed] Is Null)
            Order By [ID] Desc
        COMMIT TRANSACTION INSERTmySHIPLU
        
        --Insert into #mySHIPACKS
        --2018.09.14 only include the required fields
        --**************************
        -- [myID]
        -- , [mySort]
        --**************************
        -- , [ShipID] 
        -- , [OrderID]
        -- , [Brewer]
        -- , [BrewerName]
        -- , [Distributor]
        -- , [DistributorName]
        -- , [HalfBbl]
        -- , [TentHalfBbl]
        -- , [SixthBbl]
        -- , [TentSixthBbl]
        -- , [Pallets]
        -- , [TentPallets]
        -- , [BadPallets]
        -- , [ForeignKegs]
        -- , [DateShipped]
        -- , [DateReceived]
        -- , [MassUploadShipAckKey]
        
        ------------------------------------------------------------------------------------------------------
        --BEGIN Insert into #mySHIPACKS                                          BEGIN Insert into #mySHIPACKS
	    ------------------------------------------------------------------------------------------------------
	    BEGIN
            DECLARE @mySQLINSERTSHIPACKS as varchar(500)
            --SET @mySQLINSERTSHIPACKS = 'INSERT INTO #mySHIPACKS([mySort], [ShipID], [OrderID], [Brewer], [BrewerName], [Distributor], [DistributorName], [HalfBbl], [TentHalfBbl], [SixthBbl], [TentSixthBbl], [Pallets], [TentPallets], [BadPallets], [ForeignKegs], [DateShipped], [DateReceived], [MassUploadShipAckKey]) '
            SET @mySQLINSERTSHIPACKS = 'INSERT INTO #mySHIPACKS([mySort], [OrderID], [HalfBbl], [SixthBbl], [Pallets], [BadPallets], [ForeignKegs], [DateReceived]) '
            SET @mySQL = @mySQLINSERTSHIPACKS
            SET @mySQL = @mySQL + @pSelectRowsForInsert
            --INSERT INTO TEMP
            BEGIN TRANSACTION INSERTmySHIPACKS
            EXEC(@mySQL)
            COMMIT TRANSACTION INSERTmySHIPACKS
        END
        ------------------------------------------------------------------------------------------------------
        --END Insert into #mySHIPACKS                                              END Insert into #mySHIPACKS
	    ------------------------------------------------------------------------------------------------------
	    
	    ------------------------------------------------------------------------------------------------------
        --BEGIN UPDATE #mySHIPACKS                                                    BEGIN UPDATE #mySHIPACKS
	    ------------------------------------------------------------------------------------------------------
	    BEGIN
            BEGIN TRANSACTION UPDATEmySHIPACKS
                --First make sure the Key is updated 
                UPDATE #mySHIPACKS SET [MassUploadShipAckKey] = @pMassUploadShipAckKey
                
                UPDATE #mySHIPACKS SET [ShipID] = tRet.[ShipID]
                    , [Brewer] = tRet.[Brewer]
                    , [BrewerName] = tRet.[BrewerName]
                    , [Distributor] = tRet.[Distributor]
                    , [DistributorName] = tRet.[DistributorName]
                    , [TentHalfBbl] = tRet.[TentHalfBbl]
                    , [TentSixthBbl] = tRet.[TentSixthBbl]
                    , [TentPallets] = tRet.[TentPallets]
                    , [DateShipped] = tRet.[DateShipped]
                FROM (
                        SELECT [OrderID]
                            , [ShipID] 
                            , [Brewer]
                            , [BrewerName]
                            , [Distributor]
                            , [DistributorName]
                            , [TentHalfBbl]
                            , [TentSixthBbl]
                            , [TentPallets]
                            , [DateShipped]
                        FROM #mySHIPLU
                        WHERE [ShipID] In (SELECT [ShipID] FROM #myOrderID WHERE [Processed] Is Null)
                    ) tRet
                WHERE #mySHIPACKS.[OrderID] = tRet.[OrderID]
            COMMIT TRANSACTION UPDATEmySHIPACKS
        END
        ------------------------------------------------------------------------------------------------------
        --END UPDATE #mySHIPACKS                                                        END UPDATE #mySHIPACKS
	    ------------------------------------------------------------------------------------------------------
        
        BEGIN TRANSACTION UPDATEmyOrderIDProcessed
        UPDATE #myOrderID SET [Processed] = 1 WHERE [Processed] Is Null
        COMMIT TRANSACTION UPDATEmyOrderIDProcessed
        
        --INSERT Into tmpMassUploadShipAck
        --------------------------------------------------------------------------------------------------------------------------
        --BEGIN CURSOR GET CurPreviewSHIPAcks
        --------------------------------------------------------------------------------------------------------------------------
        Declare CurPreviewSHIPAcks Cursor for
            SELECT  [myID]
                , [mySort]
                , [ShipID]
                , [OrderID]
                , [Brewer]
                , [BrewerName]
                , [Distributor]
                , [DistributorName]
                , [HalfBbl]
                , [TentHalfBbl]
                , [SixthBbl]
                , [TentSixthBbl]
                , [Pallets]
                , [TentPallets]
                , [BadPallets]
                , [ForeignKegs]
                , [DateShipped]
                , [DateReceived]
                , [MassUploadShipAckKey] 
            FROM #mySHIPACKS
            Order By [mySort]
        --------------------------------------------------------------------------------------------------------------------------
        --END CURSOR GET CurPreviewSHIPAcks
        --------------------------------------------------------------------------------------------------------------------------
        --Init some vars
        BEGIN
            DECLARE @Error nvarchar(1500)
            DECLARE @Warning nvarchar(1500)
            DECLARE @BrewerIsSame nvarchar(50)
        END
        Open CurPreviewSHIPAcks
            Fetch Next From CurPreviewSHIPAcks into
                @myID 
                , @mySort 
                , @ShipID            	    
                , @OrderID
                , @Brewer
                , @BrewerName
                , @Distributor
                , @DistributorName
                , @HalfBbl
                , @TentHalfBbl
                , @SixthBbl
                , @TentSixthBbl
                , @Pallets
                , @TentPallets
                , @BadPallets
                , @ForeignKegs
                , @DateShipped
                , @DateReceived
                , @MassUploadShipAckKey
                
            While @@FEtch_Status = 0		
            BEGIN
                SET @Error = ''
                SET @Warning = ''
                IF NULLIF(@BrewerIsSame, '') Is Null
                BEGIN
                    SET @BrewerIsSame = @Brewer
                END
                --------------------------------------------------------------------------------------------------------------------------
                --BEGIN Validate @ShipID vs @OrderID  And the Looked up Values @Brewer, @BrewerName, @Distributor, @DistributorName
                --------------------------------------------------------------------------------------------------------------------------
                --@ShipID
                IF @ShipID Is Null
                BEGIN
                    SET @Error = @Error + '[OrderID] Not Found!</br>'
                END
                --@Brewer
                IF @Brewer Is Not Null
                BEGIN 
                    IF LTrim(RTrim(@Brewer)) = ''
                    BEGIN
                        SET @Brewer = 'Blank'
                        SET @Error = @Error + '[Brewer] Is Blank</br>'
                    END
                    IF @BrewerIsSame <> @Brewer
                    BEGIN
                        --2018.09.21 the [OrderID] supplied in the CSV is not associated with same Brewer as the very first [OrderID]
                        --So nothing can be imported until ALL the [OrderId] are verified.
                        SET @Error = @Error + 'Invalid [Brewer]vs.[OrderID]!</br> ALL [OrderID] from CSV import must be from the same [Brewer]!'
                    END 
                END
                Else 
                BEGIN
                    SET @Brewer = 'Null'
                    SET @Error = @Error + '[Brewer] Is Null</br>'
                END
                --@BrewerName
                IF @BrewerName Is Not Null
                BEGIN 
                    IF LTrim(RTrim(@BrewerName)) = ''
                    BEGIN
                        SET @BrewerName = 'Blank'
                        SET @Error = @Error + '[BrewerName] Is Blank</br>'
                    END
                END
                Else 
                BEGIN
                    SET @BrewerName = 'Null'
                    SET @Error = @Error + '[BrewerName] Is Null</br>'
                END
                --@Distributor
                IF @Distributor Is Not Null
                BEGIN 
                    IF LTrim(RTrim(@Distributor)) = ''
                    BEGIN
                        SET @Distributor = 'Blank'
                        SET @Error = @Error + '[Distributor] Is Blank</br>'
                    END
                END
                Else 
                BEGIN
                    SET @Distributor = 'Null'
                    SET @Error = @Error + '[Distributor] Is Null</br>'
                END
                --@DistributorName
                IF @DistributorName Is Not Null
                BEGIN 
                    IF LTrim(RTrim(@DistributorName)) = ''
                    BEGIN
                        SET @DistributorName = 'Blank'
                        SET @Error = @Error + '[DistributorName] Is Blank</br>'
                    END
                END
                Else 
                BEGIN
                    SET @DistributorName = 'Null'
                    SET @Error = @Error + '[DistributorName] Is Null</br>'
                END
                --@HalfBbl, @TentHalfBbl
                IF @HalfBbl Is Not Null And @TentHalfBbl Is Not Null 
                BEGIN
                    IF @HalfBbl <> @TentHalfBbl
                    BEGIN
                        SET @Warning = @Warning + '[HalfBbl]: [ ' + Cast(@HalfBbl as nvarchar(20)) + ' ] Initial Value: [ ' + Cast(@TentHalfBbl as nvarchar(20)) + ' ]</br>'
                    END
                END
                --@SixthBbl, @TentSixthBbl
                IF @SixthBbl Is Not Null And @TentSixthBbl Is Not Null 
                BEGIN
                    IF @SixthBbl <> @TentSixthBbl
                    BEGIN
                        SET @Warning = @Warning + '[SixthBbl]: [ ' + Cast(@SixthBbl as nvarchar(20)) + ' ] Initial Value: [ ' + Cast(@TentSixthBbl as nvarchar(20)) + ' ]</br>'
                    END
                END
                --@Pallets, @TentPallets
                IF @Pallets Is Not Null And @TentPallets Is Not Null
                BEGIN
                    IF @Pallets <> @TentPallets
                    BEGIN
                        SET @Warning = @Warning + '[Pallets]: [ ' + Cast(@Pallets as nvarchar(20)) + ' ] Initial Value: [ ' + Cast(@TentPallets as nvarchar(20)) + ' ]</br>'
                    END
                END
                --@BadPallets, @ForeignKegs No validation yet.
                --@DateShipped, @DateReceived
                IF @DateShipped Is Not Null AND @DateReceived Is Not Null
                BEGIN
                    IF @DateShipped > @DateReceived
                    BEGIN
                        SET @Error = @Error + '[DateReceived]: [ ' + Cast(@DateReceived as nvarchar(50)) + ' ] Is Before [DateShipped]: [ ' + Cast(@DateShipped as nvarchar(50)) + ' ]</br>'
                    END
                    IF @DateShipped = @DateReceived 
                    BEGIN
                        SET @Warning = @Warning + '[DateReceived: [ ' + Cast(@DateReceived as nvarchar(50)) + ' ] Same Day As [DateShipped]: [ ' + Cast(@DateShipped as nvarchar(50)) + ' ]</br>'
                    END
                END
                ELSE
                BEGIN
                    IF @DateShipped Is Null
                    BEGIN
                        SET @Error = @Error + '[DateShipped] Is Null</br>'
                    END
                    IF @DateReceived Is Null
                    BEGIN
                        SET @Error = @Error + '[DateReceived] Is Null</br>'
                    END
                END
                --------------------------------------------------------------------------------------------------------------------------
                --END 
                --------------------------------------------------------------------------------------------------------------------------
                
                --INSERT Results into tmpMassUploadShipAck
                INSERT INTO [dbo].[tmpMassUploadShipAck]([mySort], [ShipID], [OrderID], [Brewer], [BrewerName], [Distributor], [DistributorName], [HalfBbl], [TentHalfBbl], [SixthBbl], [TentSixthBbl], [Pallets], [TentPallets], [BadPallets], [ForeignKegs], [DateShipped], [DateReceived], [MassUploadShipAckKey], [Error], [Warning]) 
                SELECT
                @mySort As [mySort]
                , (CASE WHEN @ShipID Is Null THEN -1 ELSE @ShipID END) As [ShipID]
                , @OrderID As [OrderID]
                , @Brewer As [Brewer]
                , @BrewerName As [BrewerName]
                , @Distributor As [Distributor]
                , @DistributorName As [DistributorName]
                , @HalfBbl As [HalfBbl]
                , @TentHalfBbl As [TentHalfBbl]
                , @SixthBbl As [SixthBbl]
                , @TentSixthBbl As [TentSixthBbl]
                , @Pallets As [Pallets]
                , @TentPallets As [TentPallets]
                , @BadPallets As [BadPallets]
                , @ForeignKegs As [ForeignKegs]
                , @DateShipped As [DateShipped]
                , @DateReceived As [DateReceived]
                , @MassUploadShipAckKey As [MassUploadShipAckKey]
                , @Error As [Error]
                , @Warning As [Warning]
                --------------------------------------------------------------------------------------------------------------------------
                Fetch Next From CurPreviewSHIPAcks into
                    @myID 
                    , @mySort 
                    , @ShipID            	    
                    , @OrderID
                    , @Brewer
                    , @BrewerName
                    , @Distributor
                    , @DistributorName
                    , @HalfBbl
                    , @TentHalfBbl
                    , @SixthBbl
                    , @TentSixthBbl
                    , @Pallets
                    , @TentPallets
                    , @BadPallets
                    , @ForeignKegs
                    , @DateShipped
                    , @DateReceived
                    , @MassUploadShipAckKey
                   
            END
            
        Close CurPreviewSHIPAcks
        Deallocate CurPreviewSHIPAcks
        
        Drop Table #mySHIPACKS
        Drop Table #mySHIPLU
        Drop Table #myOrderID
        --Exit SP
        RETURN 1
	END
	----------------------------------------------------------------------------------------------------------------------------------------------------------------
	--InsertPreviewONLY -- END                                                                                                              InsertPreviewONLY -- END
	----------------------------------------------------------------------------------------------------------------------------------------------------------------
	
    ----------------------------------------------------------------------------------------------------------------------------------------------------------------
	--PreviewONLY -- BEGIN                                                                                                                      PreviewONLY -- BEGIN
	----------------------------------------------------------------------------------------------------------------------------------------------------------------
    IF @pPreviewONLY = 1 AND @pInsertPreviewONLY = 0 AND @pDoUpdateShipAck = 0
    BEGIN
        --Create The Temp HTML Output Table
        DECLARE @myHtmlOutPut as varchar(80)
        SET @myHtmlOutPut = @pTapGUID
        SET @mySQL = 'CREATE TABLE ' + @myHtmlOutPut + ' ' 
        SET @mySQL = @mySQL + '( '
        SET @mySQL = @mySQL + '[myID] bigint IDENTITY(1,1) NOT NULL '
        SET @mySQL = @mySQL + ', [mySort] bigint NOT NULL '
        SET @mySQL = @mySQL + ', [ShipAckContentHtml] varchar(8000) NULL '
        SET @mySQL = @mySQL + ', [Error] nvarchar(1500) NULL '
        SET @mySQL = @mySQL + ', [Warning] nvarchar(1500) NULL '
        SET @mySQL = @mySQL + ') '
        
        BEGIN TRANSACTION CREATETABLEmyHtmlOutPutpTapGUID
            EXEC(@mySQL)
        COMMIT TRANSACTION CREATETABLEmyHtmlOutPutpTapGUID
        
        CREATE TABLE #myHtmlOutPut
        (     
            [myID] bigint IDENTITY(1,1) NOT NULL
            , [mySort] bigint NOT NULL
            , [ShipAckContentHtml] varchar(8000) NULL
            , [Error] nvarchar(1500) NULL
            , [Warning] nvarchar(1500) NULL
        ) 
        
        --Build Template HTML String
        DECLARE @Htmplt as varchar(8000)
        DECLARE @defaultTemplate as varchar(8000)
        DECLARE @tdStylePrefix varchar(200)
        SET @tdStylePrefix = '<td style="white-space:nowrap; color:'
        SET @defaultTemplate = ''
        SET @defaultTemplate = @defaultTemplate + '<tr> '
            SET @defaultTemplate = @defaultTemplate + @tdStylePrefix + '{strMySort-color}">{strMySort}.</td> '
            SET @defaultTemplate = @defaultTemplate + @tdStylePrefix + '{strOrderID-color}">{strOrderID}</td> '
            SET @defaultTemplate = @defaultTemplate + @tdStylePrefix + '{strBrewer-color}">{strBrewer}</td> '
            SET @defaultTemplate = @defaultTemplate + @tdStylePrefix + '{strBrewerName-color}">{strBrewerName}</td> '
            SET @defaultTemplate = @defaultTemplate + @tdStylePrefix + '{strDistributor-color}">{strDistributor}</td> '
            SET @defaultTemplate = @defaultTemplate + @tdStylePrefix + '{strDistributorName-color}">{strDistributorName}</td> '
            SET @defaultTemplate = @defaultTemplate + @tdStylePrefix + '{strHalfBbl-color}">{strHalfBbl}</td> '
            SET @defaultTemplate = @defaultTemplate + @tdStylePrefix + '{strTentHalfBbl-color}">{strTentHalfBbl}</td> '
            SET @defaultTemplate = @defaultTemplate + @tdStylePrefix + '{strSixthBbl-color}">{strSixthBbl}</td> '
            SET @defaultTemplate = @defaultTemplate + @tdStylePrefix + '{strTentSixthBbl-color}">{strTentSixthBbl}</td> '
            SET @defaultTemplate = @defaultTemplate + @tdStylePrefix + '{strPallets-color}">{strPallets}</td> '
            SET @defaultTemplate = @defaultTemplate + @tdStylePrefix + '{strTentPallets-color}">{strTentPallets}</td> '
            SET @defaultTemplate = @defaultTemplate + @tdStylePrefix + '{strBadPallets-color}">{strBadPallets}</td> '
            SET @defaultTemplate = @defaultTemplate + @tdStylePrefix + '{strForeignKegs-color}">{strForeignKegs}</td> '
            SET @defaultTemplate = @defaultTemplate + @tdStylePrefix + '{strDateShipped-color}">{strDateShipped}</td> '
            SET @defaultTemplate = @defaultTemplate + @tdStylePrefix + '{strDateReceived-color}">{strDateReceived}</td> '
            SET @defaultTemplate = @defaultTemplate + @tdStylePrefix + '{strErrorWarning-color} border-style:{strErrorWarning-border-style}; " title="{strErrorWarning-title}">{strErrorWarning}</td> '
        SET @defaultTemplate = @defaultTemplate + '</tr> '
        
        IF @pTAPAppVersion = 'TAP2.5'
        BEGIN
            SET @Htmplt = @defaultTemplate
        END
        ELSE
        BEGIN
            SET @Htmplt = @defaultTemplate
        END
        --------------------------------------------------------------------------------------------------------------------------
        --BEGIN CURSOR GET CurSHIPAcks 
        --------------------------------------------------------------------------------------------------------------------------
            --@HTML Current ROW
                DECLARE @HtmlRow as varchar(3000)
                DECLARE @mySortHtmlRow as bigint
                DECLARE @ShipAckContentHtml as varchar(7000)
                DECLARE @ErrorSHIPAcks as nvarchar(1500)
                DECLARE @WarningSHIPAcks  as nvarchar(1500)
            --Declare HTML Replace Vars
                
            Declare CurSHIPAcks Cursor for
                SELECT [ID], [mySort], [ShipID], [OrderID], [Brewer], [BrewerName], [Distributor], [DistributorName], [HalfBbl], [TentHalfBbl], [SixthBbl], [TentSixthBbl], [Pallets], [TentPallets], [BadPallets], [ForeignKegs], [DateShipped], [DateReceived], [MassUploadShipAckKey], [Error], [Warning] 
                FROM [dbo].[tmpMassUploadShipAck]
                WHERE [MassUploadShipAckKey] = @pMassUploadShipAckKey
                Order By [mySort]
        --------------------------------------------------------------------------------------------------------------------------
        --END CURSOR GET CurSHIPAcks 
        --------------------------------------------------------------------------------------------------------------------------
        --Init some vars
        BEGIN
            SET @mySortHtmlRow = 0
            SET @ShipAckContentHtml = ''
            SET @ErrorSHIPAcks = ''
            SET @WarningSHIPAcks = ''
        END
        Open CurSHIPAcks
            Fetch Next From CurSHIPAcks into
                @ID
                , @mySort
                , @ShipID
                , @OrderID
                , @Brewer
                , @BrewerName
                , @Distributor
                , @DistributorName
                , @HalfBbl
                , @TentHalfBbl
                , @SixthBbl
                , @TentSixthBbl
                , @Pallets
                , @TentPallets
                , @BadPallets
                , @ForeignKegs
                , @DateShipped
                , @DateReceived
                , @MassUploadShipAckKey
                , @ErrorSHIPAcks
                , @WarningSHIPAcks
            While @@FEtch_Status = 0		
            BEGIN
                BEGIN
                --------------------------------------------------------------------------------------------------------------------------
                --BEGIN Replace the Template                                                                    BEGIN Replace the Template
                --------------------------------------------------------------------------------------------------------------------------
                    --Replace the Template
                    SET @HtmlRow = @Htmplt
                --1.  {strMySort} {strMySort-color}
                    SET @HtmlRow = Replace(@HtmlRow, '{strMySort}', Cast(COALESCE(@mySort, '') as varchar(20)))
                    IF Len(COALESCE(@ErrorSHIPAcks, '')) > 0
                    BEGIN
                        SET @HtmlRow = Replace(@HtmlRow, '{strMySort-color}', 'red;')
                    END
                    ELSE IF Len(COALESCE(@WarningSHIPAcks, '')) > 0
                    BEGIN
                        SET @HtmlRow = Replace(@HtmlRow, '{strMySort-color}', 'green;')
                    END
                    ELSE
                    BEGIN
                        SET @HtmlRow = Replace(@HtmlRow, '{strMySort-color}', 'black;')
                    END
                --2.  {strOrderID} {strOrderID-color}
                    SET @HtmlRow = Replace(@HtmlRow, '{strOrderID}', COALESCE(@OrderID, ''))
                    IF CHARINDEX('[OrderID]', COALESCE(@ErrorSHIPAcks, '')) > 0
                    BEGIN
                        SET @HtmlRow = Replace(@HtmlRow, '{strOrderID-color}', 'red;')
                    END
                    ELSE IF CHARINDEX('[OrderID]', COALESCE(@WarningSHIPAcks, '')) > 0
                    BEGIN
                        SET @HtmlRow = Replace(@HtmlRow, '{strOrderID-color}', 'green;')
                    END
                    ELSE
                    BEGIN
                        SET @HtmlRow = Replace(@HtmlRow, '{strOrderID-color}', 'black;')
                    END
                --3.  {strBrewer} {strBrewer-color}
                    SET @HtmlRow = Replace(@HtmlRow, '{strBrewer}', COALESCE(@Brewer, ''))
                    IF CHARINDEX('[Brewer]', COALESCE(@ErrorSHIPAcks, '')) > 0
                    BEGIN
                        SET @HtmlRow = Replace(@HtmlRow, '{strBrewer-color}', 'red;')
                    END
                    ELSE IF CHARINDEX('[Brewer]', COALESCE(@WarningSHIPAcks, '')) > 0
                    BEGIN
                        SET @HtmlRow = Replace(@HtmlRow, '{strBrewer-color}', 'green;')
                    END
                    ELSE
                    BEGIN
                        SET @HtmlRow = Replace(@HtmlRow, '{strBrewer-color}', 'black;')
                    END
                --4.  {strBrewerName} {strBrewerName-color}
                    SET @HtmlRow = Replace(@HtmlRow, '{strBrewerName}', COALESCE(@BrewerName, ''))
                    IF CHARINDEX('[BrewerName]', COALESCE(@ErrorSHIPAcks, '')) > 0
                    BEGIN
                        SET @HtmlRow = Replace(@HtmlRow, '{strBrewerName-color}', 'red;')
                    END
                    ELSE IF CHARINDEX('[BrewerName]', COALESCE(@WarningSHIPAcks, '')) > 0
                    BEGIN
                        SET @HtmlRow = Replace(@HtmlRow, '{strBrewerName-color}', 'green;')
                    END
                    ELSE
                    BEGIN
                        SET @HtmlRow = Replace(@HtmlRow, '{strBrewerName-color}', 'black;')
                    END
                --5.  {strDistributor} {strDistributor-color}
                    SET @HtmlRow = Replace(@HtmlRow, '{strDistributor}', COALESCE(@Distributor, ''))
                    IF CHARINDEX('[Distributor]', COALESCE(@ErrorSHIPAcks, '')) > 0
                    BEGIN
                        SET @HtmlRow = Replace(@HtmlRow, '{strDistributor-color}', 'red;')
                    END
                    ELSE IF CHARINDEX('[Distributor]', COALESCE(@WarningSHIPAcks, '')) > 0
                    BEGIN
                        SET @HtmlRow = Replace(@HtmlRow, '{strDistributor-color}', 'green;')
                    END
                    ELSE
                    BEGIN
                        SET @HtmlRow = Replace(@HtmlRow, '{strDistributor-color}', 'black;')
                    END
                --6.  {strDistributorName} {strDistributorName-color}
                    SET @HtmlRow = Replace(@HtmlRow, '{strDistributorName}', COALESCE(@DistributorName, ''))
                    IF CHARINDEX('[DistributorName]', COALESCE(@ErrorSHIPAcks, '')) > 0
                    BEGIN
                        SET @HtmlRow = Replace(@HtmlRow, '{strDistributorName-color}', 'red;')
                    END
                    ELSE IF CHARINDEX('[DistributorName]', COALESCE(@WarningSHIPAcks, '')) > 0
                    BEGIN
                        SET @HtmlRow = Replace(@HtmlRow, '{strDistributorName-color}', 'green;')
                    END
                    ELSE
                    BEGIN
                        SET @HtmlRow = Replace(@HtmlRow, '{strDistributorName-color}', 'black;')
                    END
                --7.  {strHalfBbl} {strHalfBbl-color}
                    SET @HtmlRow = Replace(@HtmlRow, '{strHalfBbl}', Cast(COALESCE(@HalfBbl, '') as nvarchar(20)))
                    IF CHARINDEX('[HalfBbl]', COALESCE(@ErrorSHIPAcks, '')) > 0
                    BEGIN
                        SET @HtmlRow = Replace(@HtmlRow, '{strHalfBbl-color}', 'red;')
                    END
                    ELSE IF CHARINDEX('[HalfBbl]', COALESCE(@WarningSHIPAcks, '')) > 0
                    BEGIN
                        SET @HtmlRow = Replace(@HtmlRow, '{strHalfBbl-color}', 'green;')
                    END
                    ELSE
                    BEGIN
                        SET @HtmlRow = Replace(@HtmlRow, '{strHalfBbl-color}', 'black;')
                    END
                --8.  {strTentHalfBbl} {strTentHalfBbl-color}
                    SET @HtmlRow = Replace(@HtmlRow, '{strTentHalfBbl}', Cast(COALESCE(@TentHalfBbl, '') as nvarchar(20)))
                    IF CHARINDEX('[TentHalfBbl]', COALESCE(@ErrorSHIPAcks, '')) > 0
                    BEGIN
                        SET @HtmlRow = Replace(@HtmlRow, '{strTentHalfBbl-color}', 'red;')
                    END
                    ELSE IF CHARINDEX('[TentHalfBbl]', COALESCE(@WarningSHIPAcks, '')) > 0
                    BEGIN
                        SET @HtmlRow = Replace(@HtmlRow, '{strTentHalfBbl-color}', 'green;')
                    END
                    ELSE
                    BEGIN
                        SET @HtmlRow = Replace(@HtmlRow, '{strTentHalfBbl-color}', 'black;')
                    END
                --9.  {strSixthBbl} {strSixthBbl-color}
                    SET @HtmlRow = Replace(@HtmlRow, '{strSixthBbl}', Cast(COALESCE(@SixthBbl, '') as nvarchar(20)))
                    IF CHARINDEX('[SixthBbl]', COALESCE(@ErrorSHIPAcks, '')) > 0
                    BEGIN
                        SET @HtmlRow = Replace(@HtmlRow, '{strSixthBbl-color}', 'red;')
                    END
                    ELSE IF CHARINDEX('[SixthBbl]', COALESCE(@WarningSHIPAcks, '')) > 0
                    BEGIN
                        SET @HtmlRow = Replace(@HtmlRow, '{strSixthBbl-color}', 'green;')
                    END
                    ELSE
                    BEGIN
                        SET @HtmlRow = Replace(@HtmlRow, '{strSixthBbl-color}', 'black;')
                    END
                --10. {strTentSixthBbl} {strTentSixthBbl-color}
                    SET @HtmlRow = Replace(@HtmlRow, '{strTentSixthBbl}', Cast(COALESCE(@TentSixthBbl, '') as nvarchar(20)))
                    IF CHARINDEX('[TentSixthBbl]', COALESCE(@ErrorSHIPAcks, '')) > 0
                    BEGIN
                        SET @HtmlRow = Replace(@HtmlRow, '{strTentSixthBbl-color}', 'red;')
                    END
                    ELSE IF CHARINDEX('[TentSixthBbl]', COALESCE(@WarningSHIPAcks, '')) > 0
                    BEGIN
                        SET @HtmlRow = Replace(@HtmlRow, '{strTentSixthBbl-color}', 'green;')
                    END
                    ELSE
                    BEGIN
                        SET @HtmlRow = Replace(@HtmlRow, '{strTentSixthBbl-color}', 'black;')
                    END
                --11. {strPallets} {strPallets-color}
                    SET @HtmlRow = Replace(@HtmlRow, '{strPallets}', Cast(COALESCE(@Pallets, '') as nvarchar(20)))
                    IF CHARINDEX('[Pallets]', COALESCE(@ErrorSHIPAcks, '')) > 0
                    BEGIN
                        SET @HtmlRow = Replace(@HtmlRow, '{strPallets-color}', 'red;')
                    END
                    ELSE IF CHARINDEX('[Pallets]', COALESCE(@WarningSHIPAcks, '')) > 0
                    BEGIN
                        SET @HtmlRow = Replace(@HtmlRow, '{strPallets-color}', 'green;')
                    END
                    ELSE
                    BEGIN
                        SET @HtmlRow = Replace(@HtmlRow, '{strPallets-color}', 'black;')
                    END
                --12. {strTentPallets} {strTentPallets-color}
                    SET @HtmlRow = Replace(@HtmlRow, '{strTentPallets}', Cast(COALESCE(@TentPallets, '') as nvarchar(20)))
                    IF CHARINDEX('[TentPallets]', COALESCE(@ErrorSHIPAcks, '')) > 0
                    BEGIN
                        SET @HtmlRow = Replace(@HtmlRow, '{strTentPallets-color}', 'red;')
                    END
                    ELSE IF CHARINDEX('[TentPallets]', COALESCE(@WarningSHIPAcks, '')) > 0
                    BEGIN
                        SET @HtmlRow = Replace(@HtmlRow, '{strTentPallets-color}', 'green;')
                    END
                    ELSE
                    BEGIN
                        SET @HtmlRow = Replace(@HtmlRow, '{strTentPallets-color}', 'black;')
                    END
                --13. {strBadPallets} {strBadPallets-color}
                    SET @HtmlRow = Replace(@HtmlRow, '{strBadPallets}', Cast(COALESCE(@BadPallets, '') as nvarchar(20)))
                    IF CHARINDEX('[BadPallets]', COALESCE(@ErrorSHIPAcks, '')) > 0
                    BEGIN
                        SET @HtmlRow = Replace(@HtmlRow, '{strBadPallets-color}', 'red;')
                    END
                    ELSE IF CHARINDEX('[BadPallets]', COALESCE(@WarningSHIPAcks, '')) > 0
                    BEGIN
                        SET @HtmlRow = Replace(@HtmlRow, '{strBadPallets-color}', 'green;')
                    END
                    ELSE
                    BEGIN
                        SET @HtmlRow = Replace(@HtmlRow, '{strBadPallets-color}', 'black;')
                    END
                --14. {strForeignKegs} {strForeignKegs-color}
                    SET @HtmlRow = Replace(@HtmlRow, '{strForeignKegs}', Cast(COALESCE(@ForeignKegs, '') as nvarchar(20)))
                    IF CHARINDEX('[ForeignKegs]', COALESCE(@ErrorSHIPAcks, '')) > 0
                    BEGIN
                        SET @HtmlRow = Replace(@HtmlRow, '{strForeignKegs-color}', 'red;')
                    END
                    ELSE IF CHARINDEX('[ForeignKegs]', COALESCE(@WarningSHIPAcks, '')) > 0
                    BEGIN
                        SET @HtmlRow = Replace(@HtmlRow, '{strForeignKegs-color}', 'green;')
                    END
                    ELSE
                    BEGIN
                        SET @HtmlRow = Replace(@HtmlRow, '{strForeignKegs-color}', 'black;')
                    END
                --15. {strDateShipped} {strDateShipped-color}
                    SET @HtmlRow = Replace(@HtmlRow, '{strDateShipped}', Cast(COALESCE(@DateShipped, '') as nvarchar(20)))
                    IF CHARINDEX('[DateShipped]', COALESCE(@ErrorSHIPAcks, '')) > 0
                    BEGIN
                        SET @HtmlRow = Replace(@HtmlRow, '{strDateShipped-color}', 'red;')
                    END
                    ELSE IF CHARINDEX('[DateShipped]', COALESCE(@WarningSHIPAcks, '')) > 0
                    BEGIN
                        SET @HtmlRow = Replace(@HtmlRow, '{strDateShipped-color}', 'green;')
                    END
                    ELSE
                    BEGIN
                        SET @HtmlRow = Replace(@HtmlRow, '{strDateShipped-color}', 'black;')
                    END
                --16. {strDateReceived} {strDateReceived-color}
                    SET @HtmlRow = Replace(@HtmlRow, '{strDateReceived}', Cast(COALESCE(@DateReceived, '') as nvarchar(50)))
                    IF CHARINDEX('[DateReceived]', COALESCE(@ErrorSHIPAcks, '')) > 0
                    BEGIN
                        SET @HtmlRow = Replace(@HtmlRow, '{strDateReceived-color}', 'red;')
                    END
                    ELSE IF CHARINDEX('[DateReceived]', COALESCE(@WarningSHIPAcks, '')) > 0
                    BEGIN
                        SET @HtmlRow = Replace(@HtmlRow, '{strDateReceived-color}', 'green;')
                    END
                    ELSE
                    BEGIN
                        SET @HtmlRow = Replace(@HtmlRow, '{strDateReceived-color}', 'black;')
                    END
                --17. {strErrorWarning} {strErrorWarning-color} {strErrorWarning-border-style} {strErrorWarning-title}
                    IF Len(COALESCE(@ErrorSHIPAcks, '')) > 0
                    BEGIN
                        SET @HtmlRow = Replace(@HtmlRow, '{strErrorWarning}', @ErrorSHIPAcks)
                        SET @HtmlRow = Replace(@HtmlRow, '{strErrorWarning-color}', 'red;')
                        SET @HtmlRow = Replace(@HtmlRow, '{strErrorWarning-border-style}', 'outset;')
                        SET @HtmlRow = Replace(@HtmlRow, '{strErrorWarning-title}', Replace(@ErrorSHIPAcks, '</br>', Cast(Char(10) + Char(13) as varchar(5)) ) )
                    END
                    ELSE IF Len(COALESCE(@WarningSHIPAcks, '')) > 0
                    BEGIN
                        SET @HtmlRow = Replace(@HtmlRow, '{strErrorWarning}', @WarningSHIPAcks)
                        SET @HtmlRow = Replace(@HtmlRow, '{strErrorWarning-color}', 'green;')
                        SET @HtmlRow = Replace(@HtmlRow, '{strErrorWarning-border-style}', 'outset;')
                        SET @HtmlRow = Replace(@HtmlRow, '{strErrorWarning-title}', Replace(@WarningSHIPAcks, '</br>', Cast(Char(10) + Char(13) as varchar(5)) ) )
                    END
                    ELSE
                    BEGIN
                        SET @HtmlRow = Replace(@HtmlRow, '{strErrorWarning}', '<img title="No Errors." src="/img/check.gif" border="0">')
                        SET @HtmlRow = Replace(@HtmlRow, '{strErrorWarning-color}', 'black;')
                        SET @HtmlRow = Replace(@HtmlRow, '{strErrorWarning-border-style}', 'none;')
                        SET @HtmlRow = Replace(@HtmlRow, '{strErrorWarning-title}', 'No Errors')
                    END
                END
                --------------------------------------------------------------------------------------------------------------------------
                --END Replace the Template                                                                        END Replace the Template
                --------------------------------------------------------------------------------------------------------------------------
                
                ------------------------------------------------------------------------------------------------------BUILD the HTML INSERT
                IF Len(@ShipAckContentHtml) + Len(@HtmlRow) <= 7000
                BEGIN
                    SET @ShipAckContentHtml = @ShipAckContentHtml + @HtmlRow
                END
                ELSE
                BEGIN
                    SET @mySortHtmlRow = @mySortHtmlRow + 1
                    IF @pDebugOn = 1
                    BEGIN
                        SELECT @mySortHtmlRow As [@mySort] , @ShipAckContentHtml As [@ShipAckContentHtml], @ErrorSHIPAcks As [@Error], @WarningSHIPAcks As [@Warning]
                    END
                    BEGIN TRANSACTION
                        INSERT INTO #myHtmlOutPut ([mySort], [ShipAckContentHtml], [Error], [Warning]) Values(@mySortHtmlRow, @ShipAckContentHtml, @ErrorSHIPAcks, @WarningSHIPAcks)
                    COMMIT;
                    --RESET
                    SET @ShipAckContentHtml = @HtmlRow
                END
                
                Fetch Next From CurSHIPAcks into
                    @ID
                    , @mySort
                    , @ShipID
                    , @OrderID
                    , @Brewer
                    , @BrewerName
                    , @Distributor
                    , @DistributorName
                    , @HalfBbl
                    , @TentHalfBbl
                    , @SixthBbl
                    , @TentSixthBbl
                    , @Pallets
                    , @TentPallets
                    , @BadPallets
                    , @ForeignKegs
                    , @DateShipped
                    , @DateReceived
                    , @MassUploadShipAckKey
                    , @ErrorSHIPAcks
                    , @WarningSHIPAcks
            END
        Close CurSHIPAcks
        Deallocate CurSHIPAcks
            
        --Final Insert
        IF Len(@ShipAckContentHtml) > 0
        BEGIN
            SET @mySortHtmlRow = @mySortHtmlRow + 1
            IF @pDebugOn = 1
            BEGIN
                SELECT @mySortHtmlRow As [@mySort] , @ShipAckContentHtml As [@ShipAckContentHtml], @ErrorSHIPAcks As [@Error], @WarningSHIPAcks As [@Warning]
            END
            BEGIN TRANSACTION INSERTmyHtmlOutPut1 
                INSERT INTO #myHtmlOutPut ([mySort], [ShipAckContentHtml], [Error], [Warning]) Values(@mySortHtmlRow, @ShipAckContentHtml, @ErrorSHIPAcks, @WarningSHIPAcks)
            COMMIT TRANSACTION INSERTmyHtmlOutPut1
        END
        
        --Finally Output the HTML
        SET @mySQL = 'INSERT INTO dbo.[' + Replace(@myHtmlOutPut, '''', '''''') + '] ([mySort], [ShipAckContentHtml], [Error], [Warning]) Select [mySort], [ShipAckContentHtml], [Error], [Warning] From #myHtmlOutPut Order By [mySort] '
         IF @pDebugOn = 1
        BEGIN
            SELECT @mySQL As [@mySQL_INSERT INTO_GUID]
        END 
        BEGIN TRANSACTION INSERTmyHtmlOutPut2
            EXEC(@mySQL)
        COMMIT TRANSACTION INSERTmyHtmlOutPut2
        
        --Cleanup
        Drop Table #myHtmlOutPut
        
        --Exit SP
        RETURN 1
    END
    ----------------------------------------------------------------------------------------------------------------------------------------------------------------
	--PreviewONLY -- END                                                                                                                          PreviewONLY -- END
	----------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    ----------------------------------------------------------------------------------------------------------------------------------------------------------------
    --DoUpdateShipAck -- BEGIN                                                                                                              DoUpdateShipAck -- BEGIN
	----------------------------------------------------------------------------------------------------------------------------------------------------------------
    IF @pDoUpdateShipAck = 1 AND @pPreviewONLY = 0 AND @pInsertPreviewONLY = 0
    BEGIN
        
        --Need couple extra vars
        DECLARE @Customer As nvarchar(40)
        DECLARE @ParentOrderID As bigint
        DECLARE @Description As nvarchar(500)
        DECLARE @AdjUser As nvarchar(20)
        DECLARE @AdjDate As datetime
        DECLARE @Status As nvarchar(40)
        --------------------------------------------------------------------------------------------------------------------------
        --BEGIN CURSOR GET CurUpdSHIPAcks
        --------------------------------------------------------------------------------------------------------------------------
        Declare CurUpdSHIPAcks Cursor for
            SELECT [ID]
                , [ShipID]
                , [HalfBbl]
                , [TentHalfBbl]
                , [SixthBbl]
                , [TentSixthBbl]
                , [Pallets]
                , [TentPallets]
                , [BadPallets]
                , [ForeignKegs]
                , [DateReceived]
                , Cast(Left([Brewer], 40) as nvarchar(40)) As [Customer]
                , (SELECT Top 1 [ParentOrderID] FROM [dbo].[Shipments] WHERE [ID] = [dbo].[tmpMassUploadShipAck].[ShipID]) As [ParentOrderID]
                , 'Shipment Updated to Acknowledged by Brewer-MassUploadShipAck' As [Description]
                , @pAdjUser As [AdjUser]
                , GetDate() As [AdjDate]
                , 'Acknowledged' As [Status]
            FROM [dbo].[tmpMassUploadShipAck]
            WHERE [MassUploadShipAckKey] = @pMassUploadShipAckKey
            Order By [mySort]
        --------------------------------------------------------------------------------------------------------------------------
        --END CURSOR GET CurUpdSHIPAcks
        --------------------------------------------------------------------------------------------------------------------------
--                @ID bigint   
--                @ShipID bigint                 	--SHIPMENTS.[ID]    
--                @HalfBbl bigint                 --SHIPMENTS.[HalfBbl]
--                @TentHalfBbl bigint             --SHIPMENTS.[TentHalfBbl]
--                @SixthBbl bigint                --SHIPMENTS.[SixthBbl]
--                @TentSixthBbl bigint            --SHIPMENTS.[TentSixthBbl]
--                @Pallets bigint                 --SHIPMENTS.[Pallets]
--                @TentPallets bigint             --SHIPMENTS.[TentPallets]
--                @BadPallets bigint              --SHIPMENTS.[BadPallets]
--                @ForeignKegs bigint             --SHIPMENTS.[ForeignKegs]
--                @DateReceived datetime          --SHIPMENTS.[DateReceived]
--                @Customer
--                @ParentOrderID
--                @Description
--                @AdjUser
--                @AdjDate
--                @Status
        Open CurUpdSHIPAcks
            Fetch Next From CurUpdSHIPAcks into
                @ID 
                , @ShipID            	    
                , @HalfBbl
                , @TentHalfBbl
                , @SixthBbl
                , @TentSixthBbl
                , @Pallets
                , @TentPallets
                , @BadPallets
                , @ForeignKegs
                , @DateReceived
                , @Customer
                , @ParentOrderID
                , @Description
                , @AdjUser
                , @AdjDate
                , @Status
            While @@FEtch_Status = 0		
            BEGIN
                --BEGIN TRANSACTION TRANSHIPMENTS--------------------------------------------------------BEGIN TRANSACTION TRANSHIPMENTS
                PRINT 'BEFORE TRANSACTION TRANSHIPMENTS'
                BEGIN TRANSACTION 
                    BEGIN TRY
                        --BEGIN UPD SHIPMENTS --------------------------------------------BEGIN UPD SHIPMENTS
                        UPDATE dbo.[SHIPMENTS] SET [HalfBbl] = @HalfBbl
                            , [SixthBbl] = @SixthBbl
                            , [Pallets] = @Pallets
                            , [BadPallets] = @BadPallets
                            , [ForeignKegs] = @ForeignKegs
                            , [TentHalfBbl] = @TentHalfBbl
                            , [TentSixthBbl] = @TentSixthBbl
                            , [TentPallets] = @TentPallets
                            , [DateReceived] = @DateReceived
                            , [Status] = @Status
                        WHERE dbo.[SHIPMENTS].[ID] = @ShipID
                        --END   UPD SHIPMENTS --------------------------------------------END   UPD SHIPMENTS
                        --BEGIN INSERT SECLOG --------------------------------------------BEGIN INSERT SECLOG
                        INSERT INTO [dbo].[SECLOG]([Customer], [ShipID], [ParentOrderID], [Description], [AdjDate], [AdjUser], [Status], [Halfbbl], [Sixthbbl], [Pallets], [HalfbblDistributor], [SixthbblDistributor], [PalletsDistributor], [HalfBblM], [SixthBblM], [HalfBblL], [SixthBblL], [tmpName], [BadPallets]) 
                        SELECT @Customer            As [Customer]
                        , @ShipID                   As [ShipID]
                        , @ParentOrderID            As [ParentOrderID]
                        , @Description              As [Description]
                        , @AdjDate                  As [AdjDate]
                        , @AdjUser                  As [AdjUser]
                        , @Status                   As [Status]
                        , @HalfBbl                  As [Halfbbl]
                        , @SixthBbl                 As [Sixthbbl]
                        , @Pallets                  As [Pallets]
                        , Null                      As [HalfbblDistributor]
                        , Null                      As [SixthbblDistributor]
                        , Null                      As [PalletsDistributor]
                        , Null                      As [HalfBblM]
                        , Null                      As [SixthBblM]
                        , Null                      As [HalfBblL]
                        , Null                      As [SixthBblL]
                        , @pMassUploadShipAckKey    As [tmpName]
                        , @BadPallets               As [BadPallets]
                        --END   INSERT SECLOG --------------------------------------------END   INSERT SECLOG
                        --BEGIN UPD tmpMassUploadShipAck------------------------BEGIN UPD tmpMassUploadShipAck
                        UPDATE [dbo].[tmpMassUploadShipAck] SET [Processed] = 1 WHERE [ID] = @ID
                        --END   UPD tmpMassUploadShipAck------------------------END   UPD tmpMassUploadShipAck
                    END TRY
                    BEGIN CATCH
                        PRINT 'In CATCH TRANSHIPMENTS'
                        IF(@@TRANCOUNT > 0)
                        BEGIN
                            ROLLBACK TRANSACTION 
                            PRINT 'ROLLED BACK TRANSACTION TRANSHIPMENTS'
                        END
                    END CATCH
                IF(@@TRANCOUNT > 0)
                BEGIN
                    COMMIT TRANSACTION
                    PRINT 'AFTER TRANSACTION TRANSHIPMENTS'
                END
                --END   TRANSACTION TRANSHIPMENTS--------------------------------------------------------END   TRANSACTION TRANSHIPMENTS
                Fetch Next From CurUpdSHIPAcks into
                    @ID 
                    , @ShipID            	    
                    , @HalfBbl
                    , @TentHalfBbl
                    , @SixthBbl
                    , @TentSixthBbl
                    , @Pallets
                    , @TentPallets
                    , @BadPallets
                    , @ForeignKegs
                    , @DateReceived
                    , @Customer
                    , @ParentOrderID
                    , @Description
                    , @AdjUser
                    , @AdjDate
                    , @Status
            END
        Close CurUpdSHIPAcks
        Deallocate CurUpdSHIPAcks   
    END
    ----------------------------------------------------------------------------------------------------------------------------------------------------------------
    --DoUpdateShipAck -- END                                                                                                              DoUpdateShipAck -- END
	----------------------------------------------------------------------------------------------------------------------------------------------------------------
END
GO
