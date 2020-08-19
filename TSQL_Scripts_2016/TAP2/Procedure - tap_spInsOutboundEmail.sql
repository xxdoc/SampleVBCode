
/**********************************************************************************************
*       Jira:           2018.08.31 TAP2-37 Log Keg Collar Order Confirmation Emails
*                       https://microstartap3.atlassian.net/browse/TAP2-37
*    	Description:	TAP Insert Into OutboundEmail
*		Returns:        nothing
*                       
*       Note:                                     
*	
*	Author: 	Brad Skidmore
*	Date: 		9/6/2018
**********************************************************************************************/
CREATE PROCEDURE [dbo].[tap_spInsOutboundEmail] 
@pSender varchar(50)='',
@pSubject         	varchar(200)='',
@pRecipientList   	text='',--text
@pRecipientCCList 	text=null,--text
@pRecipientBCCList	text=null,--text
@pBody            	text='',--text
@pPriority        	int=0,
@pContentType     	varchar(50)=null,
@pAttachmentPath  	varchar(250)=null,
@pAttachmentBody  	text=null,--text
@pWhenQueued      	datetime=null,
--@pWhenSent          datetime=null,
@pProcessed       	bit=0,
@pSendingUserId   	nchar(10)=null,
@pEmailType       	varchar(50)=null,
--@pMessageId       	nchar(70)=null,
@pDebugOn           bit=0
AS
BEGIN
--[dbo].[OutboundEmail]    
--[id] int IDENTITY(1,1) NOT NULL,
--[Sender]          	varchar(50) NOT NULL,
--[Subject]         	varchar(200) NULL,
--[RecipientList]   	text NOT NULL,
--[RecipientCCList] 	text NULL,
--[RecipientBCCList]	text NULL,
--[Body]            	text NOT NULL,
--[Priority]        	int NULL CONSTRAINT [DF_OutboundEmail_Priority]  DEFAULT ((0)),
--[ContentType]     	varchar(50) NULL,
--[AttachmentPath]  	varchar(250) NULL,
--[AttachmentBody]  	text NULL,
--[WhenQueued]      	datetime NOT NULL CONSTRAINT [DF_OutboundEmail_WhenQueued]  DEFAULT (getdate()),
--[WhenSent]        	datetime NULL, --NOT PASSED IN 
--[Processed]       	bit NOT NULL CONSTRAINT [DF_OutboundEmail_Processed]  DEFAULT ((0)),
--[SendingUserId]   	nchar(10) NULL,
--[EmailType]       	varchar(50) NULL,
--[MessageId]       	nchar(70) NULL --NOT PASSED IN 
    
    IF @pDebugOn = 1
    BEGIN
        --List Current Params
        SELECT @pSender As [@pSender] 
        , @pSubject As [@pSubject] 
        , @pRecipientList As [@pRecipientList] 
        , @pRecipientCCList As [@pRecipientCCList] 
        , @pRecipientBCCList As [@pRecipientBCCList] 
        , @pBody As [@pBody] 
        , @pPriority As [@pPriority] 
        , @pContentType As [@pContentType] 
        , @pAttachmentPath As [@pAttachmentPath] 
        , @pAttachmentBody As [@pAttachmentBody] 
        , @pWhenQueued As [@pWhenQueued] 
        --, @pWhenSent As [@pWhenSent] 
        , @pProcessed As [@pProcessed] 
        , @pSendingUserId As [@pSendingUserId] 
        , @pEmailType As [@pEmailType] 
        --, @pMessageId As [@pMessageId] 
        , @pDebugOn As [@pDebugOn] 
        
        SELECT
         Cast((CASE WHEN NULLIF(@pSender, '') Is Null THEN '' ELSE @pSender END) As varchar(50)) As [Sender]
        , Cast((CASE WHEN NULLIF(@pSubject, '') Is Null THEN Null ELSE @pSubject END) As varchar(200)) As [Subject]
        , Cast((CASE WHEN NULLIF(Cast(@pRecipientList as varchar(8000)), '') Is Null THEN '' ELSE @pRecipientList END) As text) As [RecipientList]
        , Cast((CASE WHEN NULLIF(Cast(@pRecipientCCList as varchar(8000)), '') Is Null THEN Null ELSE @pRecipientCCList END) As text) As [RecipientCCList]
        , Cast((CASE WHEN NULLIF(Cast(@pRecipientBCCList as varchar(8000)), '') Is Null THEN Null ELSE @pRecipientBCCList END) As text) As [RecipientBCCList]
        , Cast((CASE WHEN NULLIF(Cast(@pBody as varchar(8000)), '') Is Null THEN '' ELSE @pBody END) As text) As [Body]
        , Cast((CASE WHEN NULLIF(@pPriority, '') Is Null THEN 0 ELSE @pPriority END) As int) As [Priority]
        , Cast((CASE WHEN NULLIF(@pContentType, '') Is Null THEN Null ELSE @pContentType END) As varchar(50)) As [ContentType]
        , Cast((CASE WHEN NULLIF(@pAttachmentPath, '') Is Null THEN Null ELSE @pAttachmentPath END) As varchar(250)) As [AttachmentPath]
        , Cast((CASE WHEN NULLIF(Cast(@pAttachmentBody as varchar(8000)), '') Is Null THEN Null ELSE @pAttachmentBody END) As text) As [AttachmentBody]
        , Cast((CASE WHEN NULLIF(@pWhenQueued, '') Is Null THEN getdate() ELSE @pWhenQueued END) As datetime) As [WhenQueued]
        --, Cast((CASE WHEN NULLIF(@pWhenSent, '') Is Null THEN Null ELSE @pWhenSent END) As datetime) As [WhenSent]
        , Cast((CASE WHEN NULLIF(@pProcessed, '') Is Null THEN 0 ELSE @pProcessed END) As bit) As [Processed]
        , Cast((CASE WHEN NULLIF(@pSendingUserId, '') Is Null THEN Null ELSE @pSendingUserId END) As nchar(10)) As [SendingUserId]
        , Cast((CASE WHEN NULLIF(@pEmailType, '') Is Null THEN Null ELSE @pEmailType END) As varchar(50)) As [EmailType]
        --, Cast((CASE WHEN NULLIF(@pMessageId, '') Is Null THEN Null ELSE @pMessageId END) As nchar(70)) As [MessageId] 
    END

    INSERT INTO	[dbo].[OutboundEmail](
         [Sender]
        , [Subject]
        , [RecipientList]
        , [RecipientCCList]
        , [RecipientBCCList]
        , [Body]
        , [Priority]
        , [ContentType]
        , [AttachmentPath]
        , [AttachmentBody]
        , [WhenQueued]
        --, [WhenSent]
        , [Processed]
        , [SendingUserId]
        , [EmailType])
        --, [MessageId]) 
    SELECT
         Cast((CASE WHEN NULLIF(@pSender, '') Is Null THEN '' ELSE @pSender END) As varchar(50)) As [Sender]
        , Cast((CASE WHEN NULLIF(@pSubject, '') Is Null THEN Null ELSE @pSubject END) As varchar(200)) As [Subject]
        , Cast((CASE WHEN NULLIF(Cast(@pRecipientList as varchar(8000)), '') Is Null THEN '' ELSE @pRecipientList END) As text) As [RecipientList]
        , Cast((CASE WHEN NULLIF(Cast(@pRecipientCCList as varchar(8000)), '') Is Null THEN Null ELSE @pRecipientCCList END) As text) As [RecipientCCList]
        , Cast((CASE WHEN NULLIF(Cast(@pRecipientBCCList as varchar(8000)), '') Is Null THEN Null ELSE @pRecipientBCCList END) As text) As [RecipientBCCList]
        , Cast((CASE WHEN NULLIF(Cast(@pBody as varchar(8000)), '') Is Null THEN '' ELSE @pBody END) As text) As [Body]
        , Cast((CASE WHEN NULLIF(@pPriority, '') Is Null THEN 0 ELSE @pPriority END) As int) As [Priority]
        , Cast((CASE WHEN NULLIF(@pContentType, '') Is Null THEN Null ELSE @pContentType END) As varchar(50)) As [ContentType]
        , Cast((CASE WHEN NULLIF(@pAttachmentPath, '') Is Null THEN Null ELSE @pAttachmentPath END) As varchar(250)) As [AttachmentPath]
        , Cast((CASE WHEN NULLIF(Cast(@pAttachmentBody as varchar(8000)), '') Is Null THEN Null ELSE @pAttachmentBody END) As text) As [AttachmentBody]
        , Cast((CASE WHEN NULLIF(@pWhenQueued, '') Is Null THEN getdate() ELSE @pWhenQueued END) As datetime) As [WhenQueued]
        --, Cast((CASE WHEN NULLIF(@pWhenSent, '') Is Null THEN Null ELSE @pWhenSent END) As datetime) As [WhenSent]
        , Cast((CASE WHEN NULLIF(@pProcessed, '') Is Null THEN 0 ELSE @pProcessed END) As bit) As [Processed]
        , Cast((CASE WHEN NULLIF(@pSendingUserId, '') Is Null THEN Null ELSE @pSendingUserId END) As nchar(10)) As [SendingUserId]
        , Cast((CASE WHEN NULLIF(@pEmailType, '') Is Null THEN Null ELSE @pEmailType END) As varchar(50)) As [EmailType]
        --, Cast((CASE WHEN NULLIF(@pMessageId, '') Is Null THEN Null ELSE @pMessageId END) As nchar(70)) As [MessageId] 
END
GO
