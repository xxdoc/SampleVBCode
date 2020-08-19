
/**********************************************************************************************
*       Jira:           2018.09.07 TAP2-46 Optimization Dynamic SQL to SQL SP Iteration [2]
*                       https://microstartap3.atlassian.net/browse/TAP2-46
*    	Description:	Daily Empty Keg Report [dbo].[INVENTORYREPORT]
*		Returns:        @stopdate as EmptyDate
*                        , sum(halfempty) as HB
*                        , sum(sixthempty) as SB
*                        , sum(halfempty)+sum(sixthempty) as Total
*                        , sum(halfemptyold) as HBOld
*                        , sum(sixthemptyold) as SBOld
*                        , sum(halfemptyold)+sum(sixthemptyold) as TotalOld 
*       Note:           
*               Original Author:		Unknowm
*               Original Create date:   Unknown
*               2018.09.25 2018.09.18 TAP2-55 Widen the BillState and ShipState fields in the USER table
*               https://microstartap3.atlassian.net/browse/TAP2-55
*               CREATE TABLE tmpDISTEMPTIESNew (2018.10.03 Converted to actual temporary Table)
*               ShipState nvarchar(10), ------------>2018.09.25 NEEDS FIX ShipState nvarchar(20)
*	Author: 	Brad Skidmore
*	Date: 		10/03/2018
**********************************************************************************************/
CREATE PROCEDURE [dbo].[tap_spDistEmptiesDaily]
@pTAPAppVersion varchar(50)='TAP2.5',   --1. Versioning, just in case [different strokes for different folks]
@pStartdate datetime,                   --2. Top 1 [dbo].[INVENTORYREPORT] row where [dbo].[INVENTORYREPORT].[DateReported] >= @pStartdate
@pStopdate datetime,                    --3. Top 1 [dbo].[INVENTORYREPORT] row where [dbo].[INVENTORYREPORT].[DateReported] <= @pStopdate
@pHideDisabled bit,                     --4. Set @pHideDisabled = 1 in order to remove rows where [dbo].[USERS].[Enabler] <> 'yes'
@pCompany bigint,                       --5. Set @pCompany = -1 for ALL Companies in results.  Otherwise, ALL other companies besides @pCompany value will be removed from results 
@pDebugOn bit=0                         --6. Debugging?  SET @pDebugOn = 1 IF NOT SET @pDebugOn = 0  Will return a robust set of queries in order to interrogate/reconcile the sp results.
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
        
        IF @pStartdate Is Null 
        BEGIN
            SET @pStartdate = Cast('1753-01-01' as datetime)
        END
        
        IF @pStopdate Is Null 
        BEGIN
            SET @pStopdate = Cast('1753-01-01' as datetime)
        END
        
        IF NULLIF(@pHideDisabled, Null) Is Null
        BEGIN
            SET @pHideDisabled = 0
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
            , @pStartdate As [@pStartdate]
            , @pStopdate As [@pStopdate]
            , @pHideDisabled As [@pHideDisabled]
            , @pCompany As [@pCompany]
            , @pDebugOn As [@pDebugOn]
    END
    --Debug
    
    SET NOCOUNT ON
    --------------------------------------------------------------------------------------------------
    --BEGIN #myDISTEMPTIESNew                                                   EGIN #myDISTEMPTIESNew
    --------------------------------------------------------------------------------------------------
    IF OBJECT_ID('tempdb..[#myDISTEMPTIESNew]') Is Not Null
    BEGIN
        DROP TABLE #myDISTEMPTIESNew
    END
    CREATE TABLE #myDISTEMPTIESNew	
    (
        [Company] bigint
        , [Enabler] bit
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
        , [HalfEmptyOld] bigint
        , [SixthEmptyOld] bigint
        , [HalfEmpty] bigint
        , [SixthEmpty] bigint
        , [AvgHalfEmpty] bigint
        , [AvgSixthEmpty] bigint 
    )
    
    /* Grab the core info and shove it into a table and then we can re-work it */
    INSERT 
    INTO #myDISTEMPTIESNew 
    (
        [Company]
        , [Enabler]
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
        , [HalfEmptyOld]
        , [SixthEmptyOld]
        , [HalfEmpty]
        , [SixthEmpty]
        , [AvgHalfEmpty]
        , [AvgSixthEmpty]
    )
    Select b.[Company] As [Company]
        , (Case When Coalesce(b.[Enabler], '') = 'yes' Then 1 Else 0 End) As [Enabler]
        , Rtrim(a.[UserID]) As [UserID]
        , Replace(b.[Name], '''', '') As [Name]
        , b.[ShipAddress1] As [ShipAddress1]
        , b.[ShipAddress2] As [ShipAddress2]
        , b.[ShipCity] As [ShipCity]
        , b.[ShipZip] As [ShipZip]
        , b.[shipstate] As [ShipState]
        , b.[shipcontact] As [ShipContact]
        , b.[shipcontact2] As [ShipContact2]
        , b.[ShipPhone] As [ShipPhone]
        , b.[ShipPhone2] As [ShipPhone2]
        , b.[ShipFax] As [ShipFax]
        , (
            Select Top 1 c.[DateReported] 
            From dbo.[INVENTORYREPORT] c 
            Where c.[UserID] = a.[UserID] And c.[DateReported] <= @pStopdate 
            Order By 
                c.[DateReported] desc
                , c.[ID] desc
            ) As [DateReported]
        , (
            Select Top 1 Coalesce(d.[HalfEmpty], 0) 
            From dbo.[INVENTORYREPORT] d 
            Where d.[UserID] = a.[UserID] And d.[DateReported] <= @pStopdate 
            Order By d.[DateReported] desc
                , d.[ID] desc
            ) As [HalfEmptyOld]
        , (
            Select Top 1 Coalesce(e.[SixthEmpty], 0) 
            From dbo.[INVENTORYREPORT] e
            Where e.[UserID] = a.[UserID] And e.[DateReported] <= @pStopdate 
            Order By e.[DateReported] desc
                , e.[ID] desc
            ) As [SixthEmptyOld]
        , (
            Select Top 1 Coalesce(f.[HalfEmpty], 0) 
            From dbo.[INVENTORYREPORT] f 
            Where f.[UserID] = a.[UserID] And f.[DateReported] <= @pStopdate 
            Order By f.[DateReported] desc
                , f.[HalfEmpty] desc
                , f.[SixthEmpty] desc
            ) As [HalfEmpty]
        , (
            Select Top 1 Coalesce(g.[SixthEmpty], 0) 
            From dbo.[INVENTORYREPORT] g
            Where g.[UserID] = a.[UserID] And g.[DateReported] <= @pStopdate 
            Order By g.[DateReported] desc
                , g.[HalfEmpty] desc
                , g.[SixthEmpty] desc
            ) As [SixthEmpty]
        , (
            Select Coalesce(avg(h.[HalfEmpty]), 0) 
            From dbo.[INVENTORYREPORT] h 
            Where h.[UserID] = a.[UserID] And h.[HalfEmpty] Is Not Null And h.[DateReported] <= @pStopdate
            ) As [AvgHalfEmpty]
        , (
            Select Coalesce(avg(i.[SixthEmpty]), 0) 
            From dbo.[INVENTORYREPORT] i
            Where i.[UserID] = a.[UserID] And i.[SixthEmpty] Is Not Null And i.[DateReported] <= @pStopdate
            ) As [AvgSixthEmpty] 
    From dbo.[INVENTORYREPORT] a
        Inner Join dbo.[USERS] b
            On (a.[UserID] = b.[ID])
    Where Coalesce(b.[PermDistEmpties], '') <> 'yes' 
    And a.[DateReported] <= @pStopdate 
    And Len(Coalesce(b.[ID], '')) > 0 
    And a.[InvReportType] = 'Distributor Inventory' 
    Group By 
        b.[Company]
        , b.[Enabler]
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
    --END #myDISTEMPTIESNew                                                      END #myDISTEMPTIESNew
    --------------------------------------------------------------------------------------------------
    
    --------------------------------------------------------------------------------------------------
    --BEGIN company filtering                                                  BEGIN company filtering
    --------------------------------------------------------------------------------------------------
    /*** DBM 12/14/2015 - Implement company filtering ***/
    IF @pCompany <> -1
    BEGIN
        DELETE FROM #myDISTEMPTIESNew WHERE [Company] <> @pCompany
    END
    
    IF @pHideDisabled = 1
    BEGIN
        DELETE FROM #myDISTEMPTIESNew WHERE [Enabler] = 0
    END
    --------------------------------------------------------------------------------------------------
    --END company filtering                                                      END company filtering
    --------------------------------------------------------------------------------------------------
    --Debug
    IF @pDebugOn = 1
    BEGIN
        Select * From #myDISTEMPTIESNew
    END
    --Debug
    --------------------------------------------------------------------------------------------------
    --BEGIN Iterate through  Totals                                      BEGIN Iterate through  Totals
    --------------------------------------------------------------------------------------------------
    IF OBJECT_ID('tempdb..[#myDISTEMPTIESTotals]') Is Not Null
    BEGIN
        DROP TABLE #myDISTEMPTIESTotals
    END
    CREATE TABLE #myDISTEMPTIESTotals	
    (
        [IDSort] bigint IDENTITY(1,1) NOT NULL
        , [EmptyDate] datetime
        , [HB] bigint
        , [SB] bigint
        , [Total] bigint
        , [HBOld] bigint
        , [SBOld] bigint
        , [TotalOld] bigint
    )
    
    --Start with the @pStopdate
    DECLARE @myCurDate datetime
    SET @myCurDate = @pStopdate
    
    WHILE (@myCurDate >= @pStartdate)  
    BEGIN
        --Insert the Results for each Date starting with the Top @pStopdate
        --Then remove each day and rerun the totals inserting them into #myDISTEMPTIESTotals until we reach the @pStartdate 
        INSERT
        INTO #myDISTEMPTIESTotals 
        (
            [EmptyDate] 
            , [HB]
            , [SB]
            , [Total]
            , [HBOld]
            , [SBOld]
            , [TotalOld]
        )
        Select @myCurDate As [EmptyDate]
            , Cast(sum(halfempty) as bigint) As [HB]
            , Cast(sum(sixthempty) as bigint) As [SB]
            , Cast(( sum(halfempty) + sum(sixthempty) ) as bigint) As [Total]
            , Cast(sum(halfemptyold) as bigint) As [HBOld]
            , Cast(sum(sixthemptyold) as bigint) As [SBOld]
            , Cast(( sum(halfemptyold) + sum(sixthemptyold) ) as bigint) As [TotalOld]
        from #myDISTEMPTIESNew
        
        --Remove the Top Date
        DELETE FROM #myDISTEMPTIESNew WHERE [DateReported] >= @myCurDate
        
        --Increment Date back one day
        SET @myCurDate = DATEADD(day, -1, @myCurDate)
        --If the next date is before the Start date then bail
        IF (@myCurDate < @pStartdate) 
            BREAK  
        ELSE  
            CONTINUE  
    END  
    --------------------------------------------------------------------------------------------------
    --END Iterate through  Totals                                          END Iterate through  Totals 
    --------------------------------------------------------------------------------------------------
    
    --Return the Results
    Select [EmptyDate], [HB], [SB], [Total], [HBOld], [SBOld], [TotalOld] 
    From #myDISTEMPTIESTotals
    Order By [IDSort]
    
    --Clean Up
    DROP TABLE #myDISTEMPTIESNew
    DROP TABLE #myDISTEMPTIESTotals
END
GO
