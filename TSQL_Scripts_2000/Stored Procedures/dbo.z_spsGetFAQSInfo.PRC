SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spsGetFAQSInfo]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spsGetFAQSInfo]
GO

/******************************************************************
*
*    	Description:	FAQS Info
*			
*	
*	Author: 	Brad Skidmore
*	Date: 		2/10/2004
*
******************************************************************/


CREATE PROCEDURE [dbo].[z_spsGetFAQSInfo]
	@bHideDeleted		bit =1, 
	@UID			int,
	@FAQSID			Int=null,
	@OrderBy		varchar(500)=null,
	@GroupBy		varchar(500)=null,
	@SearchBy		varchar(1000)=null,
--updating the TypeOfLoss Table
	@bDelete		bit=0,
	@bAdd			bit=0,
	@bEdit			bit=0,
--Update variables
	@Question		varchar(8000)=null,
	@Answer			varchar(8000)=null,
	@IsDeleted		bit=null,
	@DateLastUpdated 	datetime =null,
--ViewHistory
	@bViewHistory		bit=NULL,
	@FromDate		varchar(50)=Null,
	@ToDate			varchar(50)=Null
	
AS
DECLARE @sSQL Varchar(8000)
DECLARE @sSQL1 Varchar(8000)
DECLARE @sSQL2 Varchar(8000)
DECLARE @sSQL3 Varchar(8000)
DECLARE @sSQL4 Varchar(8000)
DECLARE @sSQL5 Varchar(8000)
DECLARE @sSQL6 Varchar(8000)
DECLARE @sSQL7 Varchar(8000)
DECLARE @sSQL8 Varchar(8000)
DECLARE @sSQL9 Varchar(8000)
DECLARE @sSQL10 Varchar(8000)
DECLARE @CurDate 		DateTime
SET @sSQL1 = ''
SET @sSQL2 = ''
SET @sSQL3 = ''
SET @sSQL4 = ''
SET @sSQL5 = ''
SET @sSQL6 = ''
SET @sSQL7 = ''
SET @sSQL8 = ''
SET @sSQL9 = ''
SET @sSQL10 = ''
SET @CurDate = GetDate()
--See if the DateLastupdated is null
IF IsDate(@DateLastUpdated) =0
BEGIN
	SET @DateLastUpdated = GetDate()
END
--Clean Dates
SET @FromDate =			dbo.CleanFromOrToDate(@CurDate, @FromDate, 1)
SET @ToDate = 			dbo.CleanFromOrToDate(@CurDate, @ToDate, 0)

IF @FAQSID Is Not Null And @bViewHistory is Null
BEGIN
	IF @bDelete=0 And @bAdd=0 And @bEdit=0
	BEGIN
		SELECT 
			[FAQSID],
			[Question],
			[Answer],
			[IsDeleted],
			[DateLastUpdated],
			[UpdateByUserID]
		FROM	FAQS
		WHERE 	[FAQSID] = @FAQSID
	END
	ELSE
	BEGIN
		IF @bDelete =1
		BEGIN
			UPDATE FAQS SET 	[IsDeleted]=	(CASE 	WHEN [IsDeleted]= 0
									THEN 1
									ELSE 0
									END
								),
						[DateLastUpdated]=@DateLastUpdated ,
						[UpdateByUserID]=@UID
			WHERE FAQSID = @FAQSID
		
		END
		ELSE 
		BEGIN
			IF @bAdd=1
			BEGIN
				INSERT INTO FAQS(
							[Question],
							[Answer],
							[IsDeleted],
							[DateLastUpdated],
							[UpdateByUserID]
						)
				SELECT 			@Question As [Question],
							@Answer As [Answer],
							@IsDeleted As [IsDeleted] ,
							@DateLastUpdated As [DateLastUpdated],
							@UID As [UpdateByUserID] 
					
			END
			ELSE 
			BEGIN
				IF @bEdit=1
				BEGIN
					Update FAQS Set 		[Question] = @Question ,
									[Answer] = @Answer ,
									[IsDeleted] = @IsDeleted   ,
									[DateLastUpdated] = @DateLastUpdated  ,
									[UpdateByUserID] = @UID   
					WHERE 	FAQSID = @FAQSID
				END
			END
		END
	END
END
ELSE
BEGIN
	IF @GroupBy Is Not Null
	BEGIN
		SET @sSQL = 'SELECT RetFAQS.' + @GroupBy + ' '
	END
	ELSE
	BEGIN
		SET @sSQL = 'SELECT RetFAQS.* '
	END	
	SET @sSQL = @sSQL + 'FROM( ' 
	SET @sSQL = @sSQL + 'SELECT '
	SET @sSQL = @sSQL + '[FAQSID], '
	SET @sSQL = @sSQL + 'Convert(VarChar(8000),[Question]) As [Question], '
	SET @sSQL = @sSQL + 'Convert(VarChar(8000),[Answer]) As [Answer], '
	SET @sSQL = @sSQL + '[IsDeleted] , '
	SET @sSQL = @sSQL + '[DateLastUpdated] , '
	SET @sSQL = @sSQL + '(SELECT	USERNAME '
	SET @sSQL = @sSQL + 'FROM 		USERS '
	SET @sSQL = @sSQL + 'WHERE 	USERSID = F.[UpdateByUserID]) As [UpdateByUserName],  '
	SET @sSQL = @sSQL + '[UpdateByUserID]  '

	IF @bViewHistory Is Null
	BEGIN
		SET @sSQL = @sSQL + 'FROM FAQS F '
		SET @sSQL = @sSQL + ') RetFAQS '
		IF @bHideDeleted = 1
		BEGIN
			SET @sSQL = @sSQL + 'WHERE IsDeleted = 0 '
		END
		ELSE
		BEGIN
			SET @sSQL = @sSQL + 'WHERE IsDeleted Is Not Null '
		END
		--Searching
		IF @SearchBy Is Not Null
		BEGIN
			SET @sSQL1 = @SearchBy + ' '
		END
	END
	ELSE
	BEGIN
		SET @sSQL2 = @sSQL2 + 'FROM FAQSHistory F '
		SET @sSQL2 = @sSQL2 + ') RetFAQS '
		SET @sSQL2 = @sSQL2 + 'WHERE [DateLastUpdated] >= Convert(datetime,''' + @FromDate + ''') '
		SET @sSQL2 = @sSQL2 + 'AND [DateLastUpdated] <= Convert(datetime,''' + @ToDate + ''') '
		IF @FAQSID <> -1
		BEGIN
			SET @sSQL2 = @sSQL2 + 'AND [FAQSID] = ' + Cast(@FAQSid as varchar(10)) + ' '
		END
		--Searching
		IF @SearchBy Is Not Null
		BEGIN
			SET @sSQL3 = @SearchBy + ' '
		END
	END
	IF @bViewHistory Is Null
	BEGIN
		--Grouping
		IF @GroupBy Is Not Null
		BEGIN
			SET @sSQL4 = 'GROUP BY ' + @GroupBy + ' '			
		END
		---Sorting 
		IF @GroupBy Is Null
		BEGIN
			IF @OrderBy Is Null OR @OrderBy = 'Null'
			BEGIN
				SET @sSQL5 = @sSQL5 + 'ORDER BY [IsDeleted] Asc, [Question] '
			END
			ELSE
			BEGIN
				SET @sSQL6 = 'ORDER BY  [IsDeleted] Asc, ' + @OrderBy + ' '
			END
		END
	END
	ELSE
	BEGIN
		--Grouping
		IF @GroupBy Is Not Null
		BEGIN
			SET @sSQL7 = 'GROUP BY ' + @GroupBy + ' '			
		END
		---Sorting
		IF @GroupBy Is Null
		BEGIN
			SET @sSQL8 =  'ORDER BY  ' + @OrderBy + ' '
		END
	END
	EXEC(@sSQL + @sSQL1 + @sSQL2 + @sSQL3 + @sSQL4 + @sSQL5 + @sSQL6 + @sSQL7 + @sSQL8 + @sSQL9 + @sSQL10)
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[z_spsGetFAQSInfo]  TO [CFUSER]
GO

