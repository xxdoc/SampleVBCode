
/**********************************************************************************************
*       Jira:           2018.10.24 TAP-90 Duplicate Relationships need Scrubbed
*                       https://microstartap3.atlassian.net/browse/TAP2-90
*    	Description:	Duplicate Relationships need Scrubbed
*		Returns:        Before and after results of the update
*       Note:           the effective dates for all the disabled must be one day previous to the enabled 
*                       and each subsequent disabled duplicate will be one day previous to the other duplicate disabled relationship.
*                       The Script IS 'IDEMPOTENT' Meaning it can be ran again and without causing any harm.
*	Author: 	Brad Skidmore
*	Date: 		10/25/2018
*   Date:       10/26/2018  Add a View just incase there are any multiple duplicates that are disabled without a corresponding Enabled Version.
*                           These records shouldn’t pose a problem since there is no enabled version.  It's just that the fix needs at least one enabled
*                           Version of the duplicates in order to know where to start counting down the dates.  Hence the view to manually un-obfuscate the view.
**********************************************************************************************/
CREATE PROCEDURE [dbo].[tap_90_idempotentFix_RELATIONSHIPSeffDate]
AS
BEGIN
    --------------------------------------------------------------------------------------------------------------------------
    --BEGIN create temp table for updating RELATIONSHIPS Effective dates 
    --------------------------------------------------------------------------------------------------------------------------
    IF OBJECT_ID('tempdb..[#myUpdRELATIONSHIPSeffDate]') Is Not Null
    BEGIN
        DROP TABLE #myUpdRELATIONSHIPSeffDate
    END
    
    CREATE TABLE #myUpdRELATIONSHIPSeffDate	
    (
        [ID] bigint,
        [EffectiveDate] datetime
    )
    --------------------------------------------------------------------------------------------------------------------------
    --END create temp table for updating RELATIONSHIPS Effective dates 
    --------------------------------------------------------------------------------------------------------------------------  
    
    --------------------------------------------------------------------------------------------------------------------------
    --BEGIN create temp table for Query RELATIONSHIPS Grouping Then populate
    --------------------------------------------------------------------------------------------------------------------------
    IF OBJECT_ID('tempdb..[#myRELATIONSHIPSquery]') Is Not Null
    BEGIN
        DROP TABLE #myRELATIONSHIPSquery
    END
    
    CREATE TABLE #myRELATIONSHIPSquery	
    (
        [Brewer]        	nvarchar(50) NULL,
        [Distributor]   	nvarchar(50) NULL,
        [EffectiveDate] 	datetime NULL,      
        [FeeSchedule]   	bigint NULL,
        [FeeScheduleMap]	bigint NULL,
        [CountOfEffDate]    int
    )
    
    INSERT INTO #myRELATIONSHIPSquery
    SELECT [Brewer]
        , [Distributor]
        , [EffectiveDate]
        , [FeeSchedule]
        , [FeeScheduleMap]
        , Count([EffectiveDate]) As [CountOfEffDate] 
    FROM [dbo].[RELATIONSHIPS] 
    Where [Status] = 'Disabled' 
    Group by [Brewer]
        , [Distributor]
        , [EffectiveDate] 
        , [FeeSchedule]
        , [FeeScheduleMap] 
    Having Count([EffectiveDate]) > 1
    Order By [Brewer], [Distributor], [FeeSchedule], [FeeScheduleMap]
    
    --------------------------------------------------------------------------------------------------------------------------
    --END create temp table for Query RELATIONSHIPS Grouping Then populate
    --------------------------------------------------------------------------------------------------------------------------  
        
    --------------------------------------------------------------------------------------------------------------------------
    --BEGIN DECLARE Vars for RELATIONSHIPS
    --------------------------------------------------------------------------------------------------------------------------
    DECLARE @ID bigint
    DECLARE @Brewer nvarchar(50)
    DECLARE @Distributor nvarchar(50)
    DECLARE @FeeSchedule bigint
    DECLARE @FeeScheduleMap bigint
    DECLARE @Status nvarchar(50)
    DECLARE @Mileage bigint
    DECLARE @EffectiveDate datetime 
    --------------------------------------------------------------------------------------------------------------------------
    --END DECLARE Vars for RELATIONSHIPS
    --------------------------------------------------------------------------------------------------------------------------
    
    --------------------------------------------------------------------------------------------------------------------------
    --BEGIN DECLARE Vars for Previous Row RELATIONSHIPS
    --------------------------------------------------------------------------------------------------------------------------
    DECLARE @ID_Prev bigint
    DECLARE @Brewer_Prev nvarchar(50)
    DECLARE @Distributor_Prev nvarchar(50)
    DECLARE @FeeSchedule_Prev bigint
    DECLARE @FeeScheduleMap_Prev bigint
    DECLARE @Status_Prev nvarchar(50)
    DECLARE @Mileage_Prev bigint
    DECLARE @EffectiveDate_Prev datetime 
    
    
    DECLARE @ID_Prev_Enabled bigint
    DECLARE @Brewer_Prev_Enabled nvarchar(50)
    DECLARE @Distributor_Prev_Enabled nvarchar(50)
    DECLARE @FeeSchedule_Prev_Enabled bigint
    DECLARE @FeeScheduleMap_Prev_Enabled bigint
    DECLARE @Status_Prev_Enabled nvarchar(50)
    DECLARE @Mileage_Prev_Enabled bigint
    DECLARE @EffectiveDate_Prev_Enabled datetime 
    --This is the ONLY date value to be used to actually do the updates
    DECLARE @EffectiveDate_Update datetime
    DECLARE @EffectiveDate_Update_MINUS datetime
    DECLARE @DayOffset int
    --------------------------------------------------------------------------------------------------------------------------
    --END DECLARE Vars for Previous Row RELATIONSHIPS
    --------------------------------------------------------------------------------------------------------------------------
    
    --------------------------------------------------------------------------------------------------------------------------
    --BEGIN CURSOR GET CurGetRELATIONSHIPS
    --------------------------------------------------------------------------------------------------------------------------
    Declare CurGetRELATIONSHIPS Cursor for
        SELECT [ID]
            , [Brewer]
            , [Distributor]
            , [FeeSchedule]
            , [FeeScheduleMap]
            , [Status]
            , [Mileage]
            , [EffectiveDate] 
        FROM [dbo].[RELATIONSHIPS] 
        Where [Brewer] In (SELECT [Brewer] FROM #myRELATIONSHIPSquery)
        And [Distributor] In (SELECT [Distributor] FROM #myRELATIONSHIPSquery)
        Order By [Brewer], [Distributor], [FeeSchedule], [FeeScheduleMap], [Status] Desc, [EffectiveDate]
    Open CurGetRELATIONSHIPS
        Fetch Next From CurGetRELATIONSHIPS into
            @ID
            ,@Brewer
            ,@Distributor
            ,@FeeSchedule
            ,@FeeScheduleMap
            ,@Status
            ,@Mileage
            ,@EffectiveDate 
        While @@FEtch_Status = 0		
            BEGIN
                --Need to be sure to use only the Enabled Date to reduce the disabled dates
                --If there is no matching Enabled Brewer/Distributer then just use the same date to minus away from.
                IF Coalesce(@Brewer_Prev_Enabled, '') = Coalesce(@Brewer, '')
                    AND Coalesce(@Distributor_Prev_Enabled, '') = Coalesce(@Distributor, '') 
                    AND Coalesce(@FeeSchedule_Prev_Enabled, '') = Coalesce(@FeeSchedule , '')
                    AND Coalesce(@FeeScheduleMap_Prev_Enabled, '') = Coalesce(@FeeScheduleMap, '')
                    AND Coalesce(@Status_Prev_Enabled, '') = 'Enabled' 
                BEGIN
                    SET @EffectiveDate_Update = DATEADD(day, -1 * @DayOffset, @EffectiveDate_Prev_Enabled)
                    SET @DayOffset = @DayOffset + 1
                END
                ELSE
                BEGIN
                    SET @EffectiveDate_Update = @EffectiveDate
                END
            
                IF Coalesce(@Brewer_Prev, '') = Coalesce(@Brewer, '')
                    AND Coalesce(@Distributor_Prev, '') = Coalesce(@Distributor, '') 
                    AND Coalesce(@FeeSchedule_Prev, '') = Coalesce(@FeeSchedule , '')
                    AND Coalesce(@FeeScheduleMap_Prev, '') = Coalesce(@FeeScheduleMap, '') 
                    AND Coalesce(@Status, '') = 'Disabled' 
                    AND Cast(Coalesce(@EffectiveDate, '1753-01-01') As Datetime) >= Cast(Coalesce(@EffectiveDate_Update, '1753-01-01') As Datetime)
                BEGIN
                --------------------------------------------------------------------------------------------------------------------------
                ---BEGIN DO the preinsert into #myUpdRELATIONSHIPSeffDate only after looping through all the relationship records
                --------------------------------------------------------------------------------------------------------------------------
                    SET @EffectiveDate_Update_MINUS = DATEADD(day, -1, @EffectiveDate_Update)
                    INSERT 
                    INTO #myUpdRELATIONSHIPSeffDate 
                    (
                        [ID],
                        [EffectiveDate] 
                    )
                    SELECT @ID As [ID], @EffectiveDate_Update_MINUS As [EffectiveDate]
                --------------------------------------------------------------------------------------------------------------------------
                ---END DO the preinsert into #myUpdRELATIONSHIPSeffDate only after looping through all the relationship records
                --------------------------------------------------------------------------------------------------------------------------
                END
                
                --Record the current row to be compared to the next row
                SET @ID_Prev = @ID
                SET @Brewer_Prev = @Brewer
                SET @Distributor_Prev = @Distributor
                SET @FeeSchedule_Prev = @FeeSchedule
                SET @FeeScheduleMap_Prev  = @FeeScheduleMap
                SET @Status_Prev = @Status
                SET @Mileage_Prev = @Mileage
                SET @EffectiveDate_Prev = @EffectiveDate
                
                IF Coalesce(@Status, '') = 'Enabled'
                BEGIN
                    SET @DayOffset = 0
                    SET @ID_Prev_Enabled = @ID
                    SET @Brewer_Prev_Enabled = @Brewer
                    SET @Distributor_Prev_Enabled = @Distributor
                    SET @FeeSchedule_Prev_Enabled = @FeeSchedule
                    SET @FeeScheduleMap_Prev_Enabled  = @FeeScheduleMap
                    SET @Status_Prev_Enabled = @Status
                    SET @Mileage_Prev_Enabled = @Mileage
                    SET @EffectiveDate_Prev_Enabled = @EffectiveDate
                END
                
            Fetch Next From CurGetRELATIONSHIPS into
                @ID
                ,@Brewer
                ,@Distributor
                ,@FeeSchedule
                ,@FeeScheduleMap
                ,@Status
                ,@Mileage
                ,@EffectiveDate 
            END
    Close CurGetRELATIONSHIPS
    Deallocate CurGetRELATIONSHIPS   
    --------------------------------------------------------------------------------------------------------------------------
    --END CURSOR GET CurGetRELATIONSHIPS
    --------------------------------------------------------------------------------------------------------------------------
    
    --View the Updates before and after they are updated
    SELECT [ID] As [ID_VIEW_ALL_ASSOCIATED_RECORDS]
        , [Brewer]
        , [Distributor]
        , [FeeSchedule]
        , [FeeScheduleMap]
        , [Status]
        , [Mileage]
        , [EffectiveDate] 
    FROM [dbo].[RELATIONSHIPS] 
    Where [Brewer] In (SELECT [Brewer] FROM #myRELATIONSHIPSquery)
    And [Distributor] In (SELECT [Distributor] FROM #myRELATIONSHIPSquery)
    Order By [Brewer], [Distributor], [FeeSchedule], [FeeScheduleMap], [Status] Desc, [EffectiveDate]
    
    --View The Id to Update
    Select [ID] As [ID_TO_BE_UPDATED], [EffectiveDate] FROM #myUpdRELATIONSHIPSeffDate
    
    --View The relationships before they are updated
    Select [ID] As [ID_BEFORE_UPDATED], [Brewer], [Distributor], [FeeSchedule], [FeeScheduleMap], [Status], [Mileage], [EffectiveDate] 
    FROM [dbo].[RELATIONSHIPS]
    WHERE [ID] In (Select [ID] FROM #myUpdRELATIONSHIPSeffDate)
    Order By [Brewer], [Distributor], [FeeSchedule], [FeeScheduleMap], [Status] Desc, [EffectiveDate]
    
    --------------------------------------------------------------------------------------------------------------------------
    -----BEGIN TRANSACTION UpdRelationships------------------------------------------------BEGIN TRANSACTION UpdRelationships
    --------------------------------------------------------------------------------------------------------------------------
    
    --------------------------------------------------------------------------------------------------------------------------
    --BEGIN CURSOR UPDATE CurUpdRELATIONSHIPS
    --------------------------------------------------------------------------------------------------------------------------
    Declare CurUpdRELATIONSHIPS Cursor for
        SELECT [ID], [EffectiveDate] FROM #myUpdRELATIONSHIPSeffDate
    
    Open CurUpdRELATIONSHIPS
        Fetch Next From CurUpdRELATIONSHIPS into
            @ID
            ,@EffectiveDate 
        While @@FEtch_Status = 0		
            BEGIN
                --------------------------------------------------------------------------------------------------------------------------
                -----BEGIN TRANSACTION UpdRelationships------------------------------------------------BEGIN TRANSACTION UpdRelationships
                --------------------------------------------------------------------------------------------------------------------------
                PRINT 'BEFORE TRANSACTION UpdRelationships'
                BEGIN TRANSACTION 
                    BEGIN TRY
                        UPDATE [dbo].[RELATIONSHIPS] SET [EffectiveDate] = @EffectiveDate WHERE [ID] = @ID
                    END TRY
                    BEGIN CATCH
                        PRINT 'In CATCH UpdRelationships'
                        IF(@@TRANCOUNT > 0)
                        BEGIN
                            ROLLBACK TRANSACTION 
                            PRINT 'ROLLED BACK TRANSACTION UpdRelationships'
                        END
                    END CATCH
                IF(@@TRANCOUNT > 0)
                BEGIN
                    COMMIT TRANSACTION
                    PRINT 'AFTER TRANSACTION UpdRelationships'
                END
                --------------------------------------------------------------------------------------------------------------------------
                -----BEGIN TRANSACTION UpdRelationships------------------------------------------------BEGIN TRANSACTION UpdRelationships
                --------------------------------------------------------------------------------------------------------------------------
            Fetch Next From CurUpdRELATIONSHIPS into
                @ID
                ,@EffectiveDate 
            END
    Close CurUpdRELATIONSHIPS
    Deallocate CurUpdRELATIONSHIPS   
    --------------------------------------------------------------------------------------------------------------------------
    --END CURSOR UPDATE CurUpdRELATIONSHIPS
    --------------------------------------------------------------------------------------------------------------------------
    
    --View The relationships AFTER they are updated
    Select [ID] As [ID_AFTER_UPDATED], [Brewer], [Distributor], [FeeSchedule], [FeeScheduleMap], [Status], [Mileage], [EffectiveDate] 
    FROM [dbo].[RELATIONSHIPS]
    WHERE [ID] In (Select [ID] FROM #myUpdRELATIONSHIPSeffDate)
    Order By [Brewer], [Distributor], [FeeSchedule], [FeeScheduleMap], [Status] Desc, [EffectiveDate]
    
    --View the duplicates that needs dates changed in sequence
    SELECT [ID] As [ID_VIEW_ANYTHING_THAT_NEEDS_MORE_FIXING]
        , [Brewer]
        , [Distributor]
        , [FeeSchedule]
        , [FeeScheduleMap]
        , [Status]
        , [Mileage]
        , [EffectiveDate] 
    FROM [dbo].[RELATIONSHIPS] 
    Where [Brewer] In (
        SELECT [Brewer]
        FROM [dbo].[RELATIONSHIPS] 
        Where [Status] = 'Disabled' 
        Group by [Brewer]
            , [Distributor]
            , [EffectiveDate] 
            , [FeeSchedule]
            , [FeeScheduleMap] 
        Having Count([EffectiveDate]) > 1
    )
    And [Distributor] In (
        SELECT [Distributor]
        FROM [dbo].[RELATIONSHIPS] 
        Where [Status] = 'Disabled' 
        Group by [Brewer]
            , [Distributor]
            , [EffectiveDate] 
            , [FeeSchedule]
            , [FeeScheduleMap] 
        Having Count([EffectiveDate]) > 1
    )
    Order By [Brewer], [Distributor], [FeeSchedule], [FeeScheduleMap], [Status] Desc, [EffectiveDate]
    
    --Cleanup
    IF OBJECT_ID('tempdb..[#myUpdRELATIONSHIPSeffDate]') Is Not Null
    BEGIN
        DROP TABLE #myUpdRELATIONSHIPSeffDate
    END
    IF OBJECT_ID('tempdb..[#myRELATIONSHIPSquery]') Is Not Null
    BEGIN
        DROP TABLE #myRELATIONSHIPSquery
    END      
END
GO
