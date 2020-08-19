
/**********************************************************************************************
*       Jira:           2018.09.07 TAP2-45 Upload Shipment Ack [brewPortal.asp]-ImportFile
*                       https://microstartap3.atlassian.net/browse/TAP2-45
*    	Description:	TAP Insert Into tapCancelProcess
*		Returns:        1 After Inserting ProcessName and ProcessKeyGUID
*       Note:           The Process that is actually running must check dbo.tapCancelProcess with the @pProcessName and @pProcessKeyGUID                        
*	                    Pass in @pRequestCancelDateTime and leave @pActualCancelDateTime NULL to Make the Request for Cancel
*                       Pass in @pActualCancelDateTime and leave @pRequestCancelDateTime NULL to Record that the cancel Actually Occured from the Process that canceled itself.
*                       When Both @pRequestCancelDateTime And @pRequestCancelDateTime are passed in Null and @pProcessName and @pProcessKeyGUID are valid strings, 
*                           then if there is actually an entiry in [dbo].[tapCancelProcess] then a result bit [CancelProcess] = 1 otherwise  [CancelProcess] = 0
*	Author: 	Brad Skidmore
*	Date: 		9/26/2018
**********************************************************************************************/
CREATE PROCEDURE [dbo].[tap_spInsUpdViewTapCancelProcess]
@pTAPAppVersion varchar(50)='TAP2.5',   --1. Versioning, just in case [different strokes for different folks]
@pProcessName           nvarchar(50),   --2. Name of the process e.g. massUploads-ShipAck
@pProcessKeyGUID        nvarchar(100),  --3. Guid of the process to unique i.d. this process from all others.
@pRequestCancelDateTime datetime=null,  --4. Pass in @pRequestCancelDateTime and leave @pActualCancelDateTime null to record the cancel request
@pActualCancelDateTime  datetime=null,  --5. Pass in @pActualCancelDateTime and leave @pRequestCancelDateTime null to record when the cancel actually occurred from the process which canceled itself.
@pDebugOn bit=0                         --6. Debugging?  SET to 1 IF NOT SET to 0  Will return a robust set of queries in order to interrogate/reconcile the sp results.
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
        
        IF NULLIF(@pProcessName, '') Is Null
        BEGIN
            SET @pProcessName = ''
        END
        
        IF NULLIF(@pProcessKeyGUID, '') Is Null
        BEGIN
            SET @pProcessKeyGUID = ''
        END
        
        IF @pRequestCancelDateTime Is Not Null AND @pActualCancelDateTime Is Not Null
        BEGIN
            SET @pRequestCancelDateTime = Null
            SET @pActualCancelDateTime = Null
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
    --2018.09.21 TAP2-45 Table to store processes that should cancel - https://microstartap3.atlassian.net/browse/TAP2-45
    --Any process that have been canceld by the user should check this table and gracefully stop doing whatever it was they were doing 
    --    CREATE TABLE [dbo].[tapCancelProcess]  (
    --    [ID]                    bigint IDENTITY(1,1) NOT NULL,
    --	[ProcessName]         	nvarchar(50) NULL,
    --	[ProcessKeyGUID]     	nvarchar(100) NULL,
    --	[RequestCancelDateTime] datetime NULL,
    --	[ActualCancelDateTime]  datetime NULL,
    --	[Processed]             bit NULL DEFAULT ((0)),
    --	[EntryDate]             datetime NULL DEFAULT ((GetDate()))
    --	)
--
--    DECLARE @ProcessName nvarchar(50)
--    DECLARE @ProcessKeyGUID nvarchar(100)
--    DECLARE RequestCancelDateTime datetime
--    DECLARE @ActualCancelDateTime datetime
--    DECLARE @Processed bit
    
    
    --------------------------------------------------------------------------------------------------
    --END Required fields                                                          END Required fields
    -------------------------------------------------------------------------------------------------- 
    --Debug
    IF @pDebugOn = 1
    BEGIN
        --List Current Params
        SELECT 
            @pTAPAppVersion As [@pTAPAppVersion]
            , @pProcessName As [@pProcessName]
            , @pProcessKeyGUID As [@pProcessKeyGUID]
            , @pRequestCancelDateTime As [@pRequestCancelDateTime]
            , @pActualCancelDateTime As [@pActualCancelDateTime]
            , @pDebugOn As [@pDebugOn]
    END
    --Debug
    
    --------------------------------------------------------------------------------------------------
    --BEGIN RequestCancel                                                          BEGIN RequestCancel
    --------------------------------------------------------------------------------------------------
    IF @pRequestCancelDateTime Is Not Null AND @pActualCancelDateTime Is Null And @pProcessName <> '' And @pProcessKeyGUID <> ''
    BEGIN
        --Only INSERT once per GUID.  If multiple request come in for the same guid that's just awful.
        INSERT INTO [dbo].[tapCancelProcess]([ProcessName], [ProcessKeyGUID], [RequestCancelDateTime], [ActualCancelDateTime]) 
        SELECT 
            @pProcessName As [ProcessName]
            , @pProcessKeyGUID As [ProcessKeyGUID]
            , @pRequestCancelDateTime As [RequestCancelDateTime]
            , Null As [ActualCancelDateTime]
        WHERE (SELECT TOP 1 [ID] FROM [dbo].[tapCancelProcess] WHERE [ProcessKeyGUID] = @pProcessKeyGUID AND [ProcessName] = @pProcessName) Is Null
        
        --Debug
        IF @pDebugOn = 1
        BEGIN
            SELECT [ID], [ProcessName], [ProcessKeyGUID], [RequestCancelDateTime], [ActualCancelDateTime], [Processed], [EntryDate] 
            FROM [dbo].[tapCancelProcess]
            WHERE [ProcessKeyGUID] = @pProcessKeyGUID
        END
        --Debug
    
        IF (SELECT TOP 1 [ID] FROM [dbo].[tapCancelProcess] WHERE [ProcessKeyGUID] = @pProcessKeyGUID AND [ProcessName] = @pProcessName) Is Not Null
        BEGIN
            RETURN 1 
        END
        ELSE
        BEGIN
            RETURN 0
        END
    END
    --------------------------------------------------------------------------------------------------
    --END RequestCancel                                                              END RequestCancel 
    --------------------------------------------------------------------------------------------------
    
    --------------------------------------------------------------------------------------------------
    --BEGIN ActualCancel                                                            BEGIN ActualCancel
    --------------------------------------------------------------------------------------------------
    IF @pActualCancelDateTime Is Not Null AND @pRequestCancelDateTime Is Null And @pProcessName <> '' And @pProcessKeyGUID <> ''
    BEGIN
        --Only Update once per GUID.  If multiple Actual Cancels come in for the same guid that's just awful.
        UPDATE [dbo].[tapCancelProcess] SET 
            [ActualCancelDateTime] = @pActualCancelDateTime
            , [Processed] = 1
        WHERE [ProcessKeyGUID] = @pProcessKeyGUID
        AND [ProcessName] = @pProcessName
        AND [Processed] = 0
        
        --Debug
        IF @pDebugOn = 1
        BEGIN
            SELECT [ID], [ProcessName], [ProcessKeyGUID], [RequestCancelDateTime], [ActualCancelDateTime], [Processed], [EntryDate] 
            FROM [dbo].[tapCancelProcess]
            WHERE [ProcessKeyGUID] = @pProcessKeyGUID
        END
        --Debug
        
        IF (SELECT TOP 1 [ID] FROM [dbo].[tapCancelProcess] WHERE [ProcessKeyGUID] = @pProcessKeyGUID AND [ProcessName] = @pProcessName AND [Processed] = 1) Is Not Null
        BEGIN
            RETURN 1 
        END
        ELSE
        BEGIN
            RETURN 0
        END
    END
    --------------------------------------------------------------------------------------------------
    --END ActualCancel                                                                END ActualCancel
    --------------------------------------------------------------------------------------------------
    
    --------------------------------------------------------------------------------------------------
    --BEGIN View IsCancelProcess                                            BEGIN View IsCancelProcess 
    --------------------------------------------------------------------------------------------------
    IF (@pRequestCancelDateTime Is Null AND @pActualCancelDateTime Is Null) AND (@pProcessName <> '' And @pProcessKeyGUID <> '')
    BEGIN
        --A Process is checking to see if it needs to be canceled
        IF @pDebugOn = 1
        BEGIN
            SELECT '(ProcessName: ' + @pProcessName + ' pProcessKeyGUID: ' + @pProcessKeyGUID + 'is checking to see if it needs to be canceled)' As [DebugMessage]
        END
        
        DECLARE @CancelProcess Bit
        
        SELECT TOP 1 @CancelProcess = Count([ID]) FROM [dbo].[tapCancelProcess] WHERE [ProcessName] = @pProcessName AND [ProcessKeyGUID] = @pProcessKeyGUID
        SELECT @CancelProcess As [CancelProcess]
        RETURN @CancelProcess
    END
    --------------------------------------------------------------------------------------------------
    --END View IsCancelProcess                                                END View IsCancelProcess 
    --------------------------------------------------------------------------------------------------
    
    IF @pDebugOn = 1
    BEGIN
        SELECT 'SOMETHING IS VERY VERY VERY WRONG!'
    END
    RETURN 0
END
GO
