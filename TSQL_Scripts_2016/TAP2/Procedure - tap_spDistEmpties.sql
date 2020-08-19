
/**********************************************************************************************
*       Jira:           2018.09.07 TAP2-46 Optimization Dynamic SQL to SQL SP Iteration [2]
*                       https://microstartap3.atlassian.net/browse/TAP2-46
*    	Description:	Traffic Empties Report [dbo].[INVENTORYREPORT]
*		Returns:        
*       Note:           
*               Original Author:		Unknowm
*               Original Create date:   Unknown
*               2018.09.25 2018.09.18 TAP2-55 Widen the BillState and ShipState fields in the USER table
*        CREATE TABLE tmpDISTEMPTIES --> 2018.10.08 Converted to actual temp table #myTmpDISTEMPTIES
*            ShipState nvarchar(10), ------------>2018.09.25 NEEDS FIX ShipState nvarchar(20)
*        CREATE TABLE tmpDISTEMPTIES2 --> 2018.10.08 Converted to actual temp table #myTmpDISTEMPTIES2
*            ShipState nvarchar(10),------------>2018.09.25 NEEDS FIX ShipState nvarchar(20)
*        CREATE TABLE tmpDISTEMPTIES3 --> 2018.10.08 Converted to actual temp table #myTmpDISTEMPTIES3
            ShipState nvarchar(10),------------>2018.09.25 NEEDS FIX ShipState nvarchar(20)
*	Author: 	Brad Skidmore
*	Date: 		10/08/2018
**********************************************************************************************/
CREATE PROCEDURE [dbo].[tap_spDistEmpties]
@pTAPAppVersion varchar(50)='TAP2.5',--1. Versioning, just in case [different strokes for different folks]
@pRegion nvarchar(50),               --2. dbo.[Users].Region
@pOrderBy nvarchar(100),             --3. Fields to order by e.g. [ShipState], [Name]
@pCustSel nvarchar(50),              --4. dbo.[INVENTORYREPORT].[UserID]
@pCompany bigint,                    --5. Set @pCompany = -1 for ALL Companies in results.  Otherwise, ALL other companies besides @pCompany value will be removed from results 
@pDebugOn bit=0                      --6. Debugging?  SET @pDebugOn = 1 IF NOT SET @pDebugOn = 0  Will return a robust set of queries in order to interrogate/reconcile the sp results.
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
        
        IF NULLIF(@pRegion, '') Is Null 
        BEGIN
            SET @pRegion = ''
        END
        
        IF NULLIF(@pOrderBy, '') Is Null
        BEGIN
            SET @pOrderBy = ''
        END
        ELSE
        BEGIN
            SET @pOrderBy = ' ORDER BY ' + @pOrderBy 
        END
        
        IF NULLIF(@pCustSel, '') Is Null
        BEGIN
            SET @pCustSel = '%%%%'
        END
        
        IF NULLIF(@pCompany, 0) Is Null
        BEGIN
            SET @pCompany = -1
        END
        
        IF NULLIF(@pDebugOn, null) Is Null
        BEGIN
            SET @pDebugOn = 0
        END
    END
    --------------------------------------------------------------------------------------------------
    --END Validate Params:                                                         END Validate Params
    --------------------------------------------------------------------------------------------------  
    --Debug
    IF @pDebugOn = 1
    BEGIN
        --List Current Params
        SELECT 
            @pTAPAppVersion As [@pTAPAppVersion]
            , @pRegion As [@pRegion]
            , @pOrderBy As [@pOrderBy]
            , @pCustSel As [@pCustSel]
            , @pCompany As [@pCompany]
            , @pDebugOn As [@pDebugOn]
    END
    --Debug
    
    SET NOCOUNT ON
    /* - NOT USED as of 2018.10.08
    --------------------------------------------------------------------------------------------------
    --BEGIN #myTmpDISTIDS                                                          BEGIN #myTmpDISTIDS
    --------------------------------------------------------------------------------------------------
    IF OBJECT_ID('tempdb..[#myTmpDISTIDS]') Is Not Null
    BEGIN
        DROP TABLE #myTmpDISTIDS
    END
    CREATE TABLE #myTmpDISTIDS (
        [ID] nvarchar(20)
    )
    --------------------------------------------------------------------------------------------------
    --END #myTmpDISTIDS                                                              END #myTmpDISTIDS
    --------------------------------------------------------------------------------------------------
    */
    
    --------------------------------------------------------------------------------------------------
    --BEGIN #myTmpDISTEMPTIES                                                  BEGIN #myTmpDISTEMPTIES
    --------------------------------------------------------------------------------------------------
    IF OBJECT_ID('tempdb..[#myTmpDISTEMPTIES]') Is Not Null
    BEGIN
        DROP TABLE #myTmpDISTEMPTIES
    END
    CREATE TABLE #myTmpDISTEMPTIES	
    (
        [Company] bigint
        , [UserID] nvarchar(20)
        , [Name] nvarchar(200)
        , [ShipAddress1] nvarchar(200)
        , [ShipAddress2] nvarchar(200)
        , [ShipCity] nvarchar(60)
        , [ShipZip] nvarchar(20)
        , [ShipState] nvarchar(20)
        , [ShipContact] nvarchar(50)
        , [ShipContact2] nvarchar(50)
        , [ShipPhone] nvarchar(50)
        , [ShipPhone2] nvarchar(50)
        , [ShipFax] nvarchar(50)
        , [DateReported] datetime
        , [LastEmptyID] bigint
        , [HalfEmpty] bigint
        , [SixthEmpty] bigint
        , [AvgHalfEmpty] bigint
        , [AvgSixthEmpty] bigint
    )
    --------------------------------------------------------------------------------------------------
    --END #myTmpDISTEMPTIES                                                      END #myTmpDISTEMPTIES
    --------------------------------------------------------------------------------------------------
    
    --------------------------------------------------------------------------------------------------
    --BEGIN #myTmpDISTEMPTIES2                                                BEGIN #myTmpDISTEMPTIES2
    --------------------------------------------------------------------------------------------------
    IF OBJECT_ID('tempdb..[#myTmpDISTEMPTIES2]') Is Not Null
    BEGIN
        DROP TABLE #myTmpDISTEMPTIES2
    END
    CREATE TABLE #myTmpDISTEMPTIES2	
    (
        [Company] bigint
        , [UserID] nvarchar(20)
        , [Name] nvarchar(200)
        , [ShipAddress1] nvarchar(200)
        , [ShipAddress2] nvarchar(200)
        , [ShipCity] nvarchar(60)
        , [ShipZip] nvarchar(20)
        , [ShipState] nvarchar(20)
        , [ShipContact] nvarchar(50)
        , [ShipContact2] nvarchar(50)
        , [ShipPhone] nvarchar(50)
        , [ShipPhone2] nvarchar(50)
        , [ShipFax] nvarchar(50)
        , [DateReported] datetime
        , [LastEmptyID] bigint
        , [HalfEmpty] bigint
        , [SixthEmpty] bigint
        , [AvgHalfEmpty] bigint
        , [AvgSixthEmpty] bigint
        , [HalfDerived] bigint
        , [SixthDerived] bigint
        , [HalfSinceFull] bigint
        , [SixthSinceFull] bigint
    )
    --------------------------------------------------------------------------------------------------
    --END #myTmpDISTEMPTIES2                                                END #myTmpDISTEMPTIES2
    --------------------------------------------------------------------------------------------------     
    
    --------------------------------------------------------------------------------------------------
    --BEGIN #myTmpDISTEMPTIES3                                                BEGIN #myTmpDISTEMPTIES3
    --------------------------------------------------------------------------------------------------
    IF OBJECT_ID('tempdb..[#myTmpDISTEMPTIES3]') Is Not Null
    BEGIN
        DROP TABLE #myTmpDISTEMPTIES3
    END
    CREATE TABLE #myTmpDISTEMPTIES3	
    (
        [Company] bigint
        , [UserID] nvarchar(20)
        , [Name] nvarchar(200)
        , [ShipAddress1] nvarchar(200)
        , [ShipAddress2] nvarchar(200)
        , [ShipCity] nvarchar(60)
        , [ShipZip] nvarchar(20)
        , [ShipState] nvarchar(20)
        , [ShipContact] nvarchar(50)
        , [ShipContact2] nvarchar(50)
        , [ShipPhone] nvarchar(50)
        , [ShipPhone2] nvarchar(50)
        , [ShipFax] nvarchar(50)
        , [DateReported] datetime
        , [LastEmptyID] bigint
        , [HalfEmpty] bigint
        , [SixthEmpty] bigint
        , [AvgHalfEmpty] bigint
        , [AvgSixthEmpty] bigint
        , [HalfDerived] bigint
        , [SixthDerived] bigint
        , [HalfSinceFull] bigint
        , [SixthSinceFull] bigint
        , [DateReported2] datetime
        , [LastEmptyID2] bigint
        , [HalfEmpty2] bigint
        , [SixthEmpty2] bigint
        , [Comments] text
        , [TempComments] text
    )
    --------------------------------------------------------------------------------------------------
    --END #myTmpDISTEMPTIES3                                                    END #myTmpDISTEMPTIES3
    --------------------------------------------------------------------------------------------------

    /*
    --- NOT USED as of 2018.10.08
    --Find the inventory report numbers 
    INSERT INTO #myTmpDISTIDS
    select distinct userid from inventoryreport
    --- NOT USED as of 2018.10.08
    */
    
    
    --------------------------------------------------------------------------------------------------
    --BEGIN Insert into #myTmpDISTEMPTIES                          BEGIN Insert into #myTmpDISTEMPTIES
    --------------------------------------------------------------------------------------------------
    /* Grab the core info and shove it into a table and then we can re-work it */
    INSERT 
    INTO #myTmpDISTEMPTIES 
    (
        [Company]
        , [UserID]
        , [Name]
        , [ShipAddress1]
        , [ShipAddress2]
        , [ShipCity]
        , [ShipZip]
        , [ShipState]
        , [ShipContact]
        , [ShipContact2]
        , [ShipPhone]
        , [ShipPhone2]
        , [ShipFax]
        , [DateReported]
        , [LastEmptyID]
        , [HalfEmpty]
        , [SixthEmpty]
        , [AvgHalfEmpty]
        , [AvgSixthEmpty]
    )
    Select
        b.[Company]
        , Rtrim(a.[UserID]) As [UserID]
        , Replace(b.[Name], '''', '') As [Name]
        , b.[ShipAddress1] As [ShipAddress1]
        , b.[ShipAddress2] As [ShipAddress2]
        , b.[ShipCity] As [ShipCity]
        , b.[ShipZip] As [ShipZip]
        , b.[ShipState] As [ShipState]
        , b.[ShipContact] As [ShipContact]
        , b.[ShipContact2] As [ShipContact2]
        , b.[ShipPhone] As [ShipPhone]
        , b.[ShipPhone2] As [ShipPhone2]
        , b.[ShipFax] As [ShipFax]
        , (
            Select Top 1 c.[ReportPeriod] As [DateReported]
            From dbo.[INVENTORYREPORT] c 
            Where c.[UserID] = a.[UserID] 
            Order By c.[DateReported] Desc
        ) As [DateReported]
        , (
            Select Top 1 Coalesce(d.[ID], 0) As [LastEmptyID]
            From dbo.[INVENTORYREPORT] d
            Where d.[UserID] = a.[UserID] 
            Order By d.[DateReported] Desc , d.[ID] desc
        ) As [LastEmptyID]
        , (
            Select Top 1 Coalesce(e.[HalfEmpty], 0) As [HalfEmpty]
            From dbo.[INVENTORYREPORT] e
            Where e.[UserID] = a.[UserID] 
            Order By e.[DateReported] Desc , e.[ID] desc
        ) As [HalfEmpty]
        , (
            Select Top 1 Coalesce(f.[SixthEmpty], 0) As [SixthEmpty]
            From dbo.[INVENTORYREPORT] f
            Where f.[UserID] = a.[UserID] 
            Order By f.[DateReported] Desc , f.[ID] desc
        ) As [SixthEmpty]
        , (
            Select Coalesce(avg(g.[HalfEmpty]), 0) As [AvgHalfEmpty] 
            From dbo.[INVENTORYREPORT] g 
            Where g.[UserID] = a.[UserID] 
            AND g.[HalfEmpty] Is Not Null 
            AND Coalesce(g.[HalfEmpty], 0) > 0
        ) As [AvgHalfEmpty]
        , (
            Select Coalesce(avg(h.[SixthEmpty]),0)  As [AvgSixthEmpty] 
            From dbo.[INVENTORYREPORT] h 
            Where h.[UserID] = a.[UserID]  
            AND h.[SixthEmpty] Is Not Null 
            AND Coalesce(h.[SixthEmpty], 0) > 0
        ) As [AvgSixthEmpty] 
    From dbo.[INVENTORYREPORT] a
        Inner Join dbo.[Users] b 
            On (a.[UserID] = b.[ID])
    Where b.[CustType] != 'Brewer' 
    AND b.[PermDistEmpties] != 'yes' 
    AND Coalesce(a.[UserID], '') Like @pCustSel 
    AND Len(b.[ID]) > 0 
    AND b.[ID] Is Not Null
    AND a.[InvReportType] = 'Distributor Inventory' 
    AND Coalesce(b.[Region], '') = @pRegion 
    Group By
        b.[Company]
        , a.[UserID]
        , b.[Name]
        , b.[ShipPhone]
        , b.[ShipPhone2]
        , b.[ShipFax]
        , b.[ShipContact]
        , b.[ShipContact2]
        , b.[ShipState]
        , b.[ShipAddress1]
        , b.[ShipAddress2]
        , b.[ShipCity]
        , b.[ShipZip]
    Order By b.[Name]
    --------------------------------------------------------------------------------------------------
    --END Insert into #myTmpDISTEMPTIES                              END Insert into #myTmpDISTEMPTIES
    --------------------------------------------------------------------------------------------------
    --------------------------------------------------------------------------------------------------
    --BEGIN company filtering                                                  BEGIN company filtering
    --------------------------------------------------------------------------------------------------
    --Debug
    IF @pDebugOn = 1
    BEGIN
        --List Current Params
        SELECT 'BeforeRemoveCompany' As [#myTmpDISTEMPTIES], * From #myTmpDISTEMPTIES
    END
    --Debug
    
    /*** DBM 12/14/2015 - Implement company filtering ***/
    IF @pCompany <> -1
    BEGIN
        DELETE FROM #myTmpDISTEMPTIES WHERE [Company] <> @pCompany
    END
    
    --Debug
    IF @pDebugOn = 1
    BEGIN
        --List Current Params
        SELECT 'AfterRemoveCompany' As [#myTmpDISTEMPTIES], * From #myTmpDISTEMPTIES
    END
    --Debug
    --------------------------------------------------------------------------------------------------
    --END company filtering                                                      END company filtering
    --------------------------------------------------------------------------------------------------

    /*
    --- NOT USED as of 2018.10.08
    -- Throw in the guys who don't belong in the system
    INSERT INTO #myTmpDISTEMPTIES (UserID,Name,ShipAddress1,ShipAddress2,ShipCity,ShipZip,ShipState,ShipContact,ShipContact2,ShipPhone,ShipPhone2,ShipFax,DateReported,LastEmptyID,HalfEmpty,SixthEmpty,AvgHalfEmpty,AvgSixthEmpty)
    select
    Rtrim(ID),
    Replace(name, '''', ''),
    ShipAddress1,
    ShipAddress2,
    ShipCity,
    ShipZip,
    shipstate,
    shipcontact,
    shipcontact2,
    shipphone,
    shipphone2,
    shipfax,
    0,0,0,0,0,0
    from users a where Enabler = 'yes' AND custtype = 'Distributor' AND permdistempties != 'yes' AND name not like '%%Widmer%%' AND name not like '%%Redhook%%' AND ID >= '22000' AND ID not in (select ID from #myTmpDISTIDS) AND name not like '%%Focus%%' AND len(ID) >= 5 order by name
    --- NOT USED as of 2018.10.08
    */


    --------------------------------------------------------------------------------------------------
    --BEGIN Insert into #myTmpDISTEMPTIES2                        BEGIN Insert into #myTmpDISTEMPTIES2
    --------------------------------------------------------------------------------------------------
    /* Parse it properly */
    INSERT 
    INTO #myTmpDISTEMPTIES2 
        (
              [Company]
            , [UserID]
            , [Name]
            , [ShipAddress1]
            , [ShipAddress2]
            , [ShipCity]
            , [ShipZip]
            , [ShipState]
            , [ShipContact]
            , [ShipContact2]
            , [ShipPhone]
            , [ShipPhone2]
            , [ShipFax]
            , [DateReported]
            , [LastEmptyID]
            , [HalfEmpty]
            , [SixthEmpty]
            , [AvgHalfEmpty]
            , [AvgSixthEmpty]
            , [HalfDerived]
            , [SixthDerived]
            , [HalfSinceFull]
            , [SixthSinceFull]
        )
    SELECT 
        a.[Company]
        , a.[UserID]
        , a.[Name]
        , a.[ShipAddress1]
        , a.[ShipAddress2]
        , a.[ShipCity]
        , a.[ShipZip]
        , a.[ShipState]
        , a.[ShipContact]
        , a.[ShipContact2]
        , a.[ShipPhone]
        , a.[ShipPhone2]
        , a.[ShipFax]
        , a.[DateReported]
        , a.[LastEmptyID]
        , a.[HalfEmpty]
        , a.[SixthEmpty]
        , a.[AvgHalfEmpty]
        , a.[AvgSixthEmpty]
        , Coalesce([HalfEmpty], 0)-( 
            Select Coalesce(sum([halfbbl]), 0) 
                From dbo.[Shipments] 
                Where [distributor] = a.[UserID] 
                And [MoveType] like 'Distributor to%%' 
                And [MoveType] not like 'Distributor to%%Rev%%' 
                And [Status] != 'Canceled' 
                And [DateShipped] >= [DateReported]
          ) As [HalfDerived]
        , Coalesce([SixthEmpty], 0)-(
            Select Coalesce(sum([sixthbbl]), 0) 
            From dbo.[Shipments] 
            Where [distributor] = a.[UserID]  
            And [MoveType] like 'Distributor to%%' 
            And [MoveType] not like 'Distributor to%%Rev%%' 
            And [Status] != 'Canceled' 
            And [DateShipped] >= [DateReported]
          ) As [SixthDerived]
        , ( 
            Select Coalesce(sum([halfbbl]), 0) 
            From dbo.[Shipments] 
            Where [distributor] = a.[UserID]
            AND [MoveType] like '%%to Distributor' 
            AND [Status] != 'Canceled' 
            AND Month([DateShipped]) = Month(getdate())-1 
            AND Year([DateShipped]) = Year(getdate())
          ) As [HalfSinceFull]
        , ( 
            Select Coalesce(sum(sixthbbl), 0) 
            From dbo.[Shipments] 
            Where [distributor] = a.[UserID] 
            AND [MoveType] like '%%to Distributor' 
            AND [Status] != 'Canceled' 
            AND Month([DateShipped]) = Month(getdate())-1 
            AND Year([DateShipped]) = Year(getdate())
          ) As [SixthSinceFull] 
    FROM #myTmpDISTEMPTIES a


    UPDATE #myTmpDISTEMPTIES2 SET [HalfDerived] = 0 WHERE [HalfDerived] < 0
    UPDATE #myTmpDISTEMPTIES2 SET [SixthDerived] = 0 WHERE [SixthDerived] < 0
    --------------------------------------------------------------------------------------------------
    --END Insert into #myTmpDISTEMPTIES2                            END Insert into #myTmpDISTEMPTIES2
    --------------------------------------------------------------------------------------------------
    
    --Debug
    IF @pDebugOn = 1
    BEGIN
        --List Current Params
        SELECT '#myTmpDISTEMPTIES2' As [#myTmpDISTEMPTIES2], * From #myTmpDISTEMPTIES2
    END
    --Debug
    
    --------------------------------------------------------------------------------------------------
    --BEGIN Insert into #myTmpDISTEMPTIES3                        BEGIN Insert into #myTmpDISTEMPTIES3
    --------------------------------------------------------------------------------------------------
    /* Go find the time before last and shove that into the database */
    INSERT 
    INTO #myTmpDISTEMPTIES3 
    (
        [Company]
        , [UserID]
        , [Name]
        , [ShipAddress1]
        , [ShipAddress2]
        , [ShipCity]
        , [ShipZip]
        , [ShipState]
        , [ShipContact]
        , [ShipContact2]
        , [ShipPhone]
        , [ShipPhone2]
        , [ShipFax]
        , [DateReported]
        , [LastEmptyID]
        , [HalfEmpty]
        , [SixthEmpty]
        , [AvgHalfEmpty]
        , [AvgSixthEmpty]
        , [HalfDerived]
        , [SixthDerived]
        , [HalfSinceFull]
        , [SixthSinceFull]
        , [DateReported2]
        , [LastEmptyID2]
        , [HalfEmpty2]
        , [SixthEmpty2]
        , [Comments]
        , [TempComments]
    )
    SELECT 
        a.[Company]
        , a.[UserID]
        , a.[Name]
        , a.[ShipAddress1]
        , a.[ShipAddress2]
        , a.[ShipCity]
        , a.[ShipZip]
        , a.[ShipState]
        , a.[ShipContact]
        , a.[ShipContact2]
        , a.[ShipPhone]
        , a.[ShipPhone2]
        , a.[ShipFax]
        , a.[DateReported]
        , a.[LastEmptyID]
        , a.[HalfEmpty]
        , a.[SixthEmpty]
        , a.[AvgHalfEmpty]
        , a.[AvgSixthEmpty]
        , a.[HalfDerived]
        , a.[SixthDerived]
        , a.[HalfSinceFull]
        , a.[SixthSinceFull]
        , ( 
            Select Top 1 Coalesce([ReportPeriod], 0) 
            From dbo.[INVENTORYREPORT]  
            Where [UserID] = a.[UserID] AND [ID] < a.[LastEmptyID] 
            Order By [DateReported] Desc
            ) As [DateReported2]
        , ( 
            Select Top 1 Coalesce(ID, 0) 
            From dbo.[INVENTORYREPORT]  
            where [UserID] = a.[UserID] AND [DateReported] < a.[DateReported] 
            Order By [DateReported] Desc, [ID] Desc
            ) As [LastEmptyID2]
        , ( 
            Select Top 1 Coalesce(halfempty, 0) 
            From dbo.[INVENTORYREPORT]  
            where [UserID] = a.[UserID] AND [ID] < a.LastEmptyID 
            Order By [DateReported] Desc
                , [ID] Desc
            ) As [HalfEmpty2]
        , ( 
            Select Top 1 Coalesce(sixthempty, 0) 
            From dbo.[INVENTORYREPORT] 
            where [UserID] = a.[UserID] AND [ID] < a.LastEmptyID 
            Order By [DateReported] Desc
                , [ID] Desc
            ) As [SixthEmpty2]
        , ( 
            SELECT 
                LEFT(l.[list], LEN(l.[list])-1) 
                FROM (  
                        SELECT [subject] + '~~~' AS [text()] 
                        FROM dbo.[UsersEvents] 
                        WHERE ([Name] = 'Empty Keg Count Comment' OR [Name] = 'Old DAVE Comment') AND [UserID] = a.[UserID] AND len([Subject]) > 0 
                        ORDER BY [ID] Desc FOR XML PATH('')
                    ) l (list)
            ) As [Comments]
        , ( 
            Select Coalesce([TempComments], '') 
            From dbo.[Users] 
            Where [ID] = a.[UserID]
            ) As [TempComments]
    FROM #myTmpDISTEMPTIES2 a

    /*** NULLs Were Failing Out ***/
    UPDATE #myTmpDISTEMPTIES3 SET [halfempty2] = 0 WHERE [halfempty2] Is Null
    UPDATE #myTmpDISTEMPTIES3 SET [sixthempty2] = 0 WHERE [sixthempty2] Is Null
    --------------------------------------------------------------------------------------------------
    --END Insert into #myTmpDISTEMPTIES3                            END Insert into #myTmpDISTEMPTIES3
    --------------------------------------------------------------------------------------------------
    
    --Debug
    IF @pDebugOn = 1
    BEGIN
        --List Current Params
        SELECT '#myTmpDISTEMPTIES3' As [#myTmpDISTEMPTIES3], * From #myTmpDISTEMPTIES3
    END
    --Debug
    

    /* Go find the last time before last and output */
    DECLARE @query as varchar(3000)
    SET @query = ''
    SET @query = @query + 'SELECT '
    SET @query = @query + ' [Company] '
    SET @query = @query + ' , [UserID] '
    SET @query = @query + ' , [Name] '
    SET @query = @query + ' , [ShipAddress1] '
    SET @query = @query + ' , [ShipAddress2] '
    SET @query = @query + ' , [ShipCity] '
    SET @query = @query + ' , [ShipZip] '
    SET @query = @query + ' , [ShipState] '
    SET @query = @query + ' , [ShipContact] '
    SET @query = @query + ' , [ShipContact2] '
    SET @query = @query + ' , [ShipPhone] '
    SET @query = @query + ' , [ShipPhone2] '
    SET @query = @query + ' , [ShipFax] '
    SET @query = @query + ' , [DateReported] '
    SET @query = @query + ' , [LastEmptyID] '
    SET @query = @query + ' , [HalfEmpty] '
    SET @query = @query + ' , [SixthEmpty] '
    SET @query = @query + ' , [AvgHalfEmpty] '
    SET @query = @query + ' , [AvgSixthEmpty] '
    SET @query = @query + ' , [HalfDerived] '
    SET @query = @query + ' , [SixthDerived] '
    SET @query = @query + ' , [HalfSinceFull] '
    SET @query = @query + ' , [SixthSinceFull] '
    SET @query = @query + ' , [DateReported2] '
    SET @query = @query + ' , [LastEmptyID2] '
    SET @query = @query + ' , [HalfEmpty2] '
    SET @query = @query + ' , [SixthEmpty2] '
    SET @query = @query + ' , [Comments] ' 
    SET @query = @query + ' , [TempComments] '
    SET @query = @query + ', ( '
        SET @query = @query + 'Select Top 1 Coalesce([ReportPeriod], 0) '
        SET @query = @query + 'From dbo.[INVENTORYREPORT] '
        SET @query = @query + 'Where [UserID] = a.[UserID] AND [DateReported] < a.[DateReported2] '
        SET @query = @query + 'Order By [DateReported] Desc '
    SET @query = @query + ') As [DateReported3] '
    SET @query = @query + ', ( '
        SET @query = @query + 'Select Top 1 Coalesce([HalfEmpty], 0) '
        SET @query = @query + 'From dbo.[INVENTORYREPORT] '
        SET @query = @query + 'Where [UserID] = a.[UserID] AND [DateReported] < a.[DateReported2] ' 
        SET @query = @query + 'Order By [DateReported] Desc, [ID] Desc '
    SET @query = @query + ') As [HalfEmpty3] '
    SET @query = @query + ', ( '
        SET @query = @query + 'Select Top 1 Coalesce([SixthEmpty],0) '
        SET @query = @query + 'From dbo.[INVENTORYREPORT]  '
        SET @query = @query + 'Where [UserID] = a.[UserID] AND [DateReported] < a.[DateReported2] ' 
        SET @query = @query + 'Order By [DateReported] Desc, [ID] Desc '
    SET @query = @query + ') As [SixthEmpty3]  '
    SET @query = @query + 'FROM #myTmpDISTEMPTIES3 a '
    SET @query = @query + @pOrderBy 
    
    --Return Results from #myTmpDISTEMPTIES3 With Added: [DateReported3], [HalfEmpty3], amd [SixthEmpty3]
    Exec(@query)
    
    --Clean up
    Drop Table #myTmpDISTEMPTIES3
    Drop Table #myTmpDISTEMPTIES2
    Drop Table #myTmpDISTEMPTIES
    
END
GO
