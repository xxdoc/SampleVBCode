SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spDelCompanyUsersInfo]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spDelCompanyUsersInfo]
GO

/******************************************************************
*
*    	Description:	Delete Company Users Info
*			
*	
*	Author: 	Brad Skidmore
*	Date: 		10/14/2003
*	Date		2/17/2004 Added Date Span For Delete From History Table
******************************************************************/


CREATE PROCEDURE [dbo].[z_spDelCompanyUsersInfo]
	@UsersID		Int,
	@selCompanyID		int=null,
	@selClientCompanyID	int=null,
	@bRemoveHistoryOnly	int=0,
	@dtBeginHistoryDate	datetime=null,
	@dtEndHistoryDate	datetime=null
AS
IF @bRemoveHistoryOnly = 1 
BEGIN
	--Format the Begin and End History Dates
	IF IsDate(@dtBeginHistoryDate) = 0 
	BEGIN
		SET @dtBeginHistoryDate = Convert(datetime,dbo.FormatDate(GetDate(),'mm/dd/yyy'))
	END
	ELSE
	BEGIN
		SET @dtBeginHistoryDate = Convert(datetime,dbo.FormatDate(@dtBeginHistoryDate,'mm/dd/yyy'))
	END
	
	IF IsDate(@dtEndHistoryDate) = 0 
	BEGIN
		SET @dtEndHistoryDate = GetDate()
	END
	ELSE
	BEGIN
		SET @dtEndHistoryDate = Convert(datetime,dbo.FormatDate(@dtEndHistoryDate,'mm/dd/yyy')) + ' 23:59:59.999'
	END

	IF @UsersID = -1
	BEGIN
		IF @selCompanyID Is Not Null And @selClientCompanyID Is Not Null
		BEGIN
			DELETE FROM CompanyUsersHistory  
			WHERE  CompanyID = @selClientCompanyID
			AND 	DateLastUpdated >= @dtBeginHistoryDate
			AND 	DateLastUpdated <= @dtEndHistoryDate	
		END
		IF @selCompanyID Is Not  Null And @selClientCompanyID Is Null
		BEGIN
			DELETE FROM CompanyUsersHistory 
			WHERE  CompanyID = @selCompanyID
			AND 	DateLastUpdated >= @dtBeginHistoryDate
			AND 	DateLastUpdated <= @dtEndHistoryDate	
		END
		IF @selCompanyID Is Not Null  And @selClientCompanyID Is Null
		BEGIN
			DELETE FROM AdjusterUsersSoftwareHistory
			WHERE 	DateLastUpdated >= @dtBeginHistoryDate
			AND 	DateLastUpdated <= @dtEndHistoryDate	
			DELETE FROM AdjusterUsersUpdatesHistory
			WHERE 	DateLastUpdated >= @dtBeginHistoryDate
			AND 	DateLastUpdated <= @dtEndHistoryDate	
			DELETE FROM UsersHistory 
			WHERE 	DateLastUpdated >= @dtBeginHistoryDate
			AND 	DateLastUpdated <= @dtEndHistoryDate	
		END
	END
	ELSE
	BEGIN
		IF @selCompanyID Is Not Null And @selClientCompanyID Is Not Null
		BEGIN
			DELETE FROM CompanyUsersHistory  
			WHERE  CompanyID = @selClientCompanyID AND UsersID = @UsersID
			AND 	DateLastUpdated >= @dtBeginHistoryDate
			AND 	DateLastUpdated <= @dtEndHistoryDate	
		END
		IF @selCompanyID Is Not  Null And @selClientCompanyID Is Null
		BEGIN
			DELETE FROM CompanyUsersHistory 
			WHERE  CompanyID = @selCompanyID  AND UsersID = @UsersID
			AND 	DateLastUpdated >= @dtBeginHistoryDate
			AND 	DateLastUpdated <= @dtEndHistoryDate	
		END
		IF @selCompanyID Is Not Null  And @selClientCompanyID Is Null
		BEGIN
			DELETE FROM AdjusterUsersSoftwareHistory    
			WHERE UsersID = @UsersID
			AND 	DateLastUpdated >= @dtBeginHistoryDate
			AND 	DateLastUpdated <= @dtEndHistoryDate	
			DELETE FROM AdjusterUsersUpdatesHistory    
			WHERE UsersID = @UsersID
			AND 	DateLastUpdated >= @dtBeginHistoryDate
			AND 	DateLastUpdated <= @dtEndHistoryDate	
			DELETE FROM UsersHistory  
			WHERE UsersID = @UsersID
			AND 	DateLastUpdated >= @dtBeginHistoryDate
			AND 	DateLastUpdated <= @dtEndHistoryDate	
		END
		
	END
END
ELSE
BEGIN
	IF @selCompanyID Is Not Null And @selClientCompanyID Is Not Null
	BEGIN
		IF @UsersID = -1
		BEGIN
			DELETE FROM CompanyUsers WHERE CompanyID = @selClientCompanyID
		END
		ELSE
		BEGIN
			DELETE FROM CompanyUsers WHERE UsersID = @UsersID AND CompanyID = @selClientCompanyID
		END
	END
	IF @selCompanyID Is Not  Null And @selClientCompanyID Is Null
	BEGIN
		IF @UsersID = -1
		BEGIN
			DELETE FROM CompanyUsers WHERE CompanyID = @selCompanyID
		END
		ELSE
		BEGIN
			DELETE FROM CompanyUsers WHERE UsersID = @UsersID AND CompanyID = @selCompanyID
		END
		
	END
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[z_spDelCompanyUsersInfo]  TO [CFUSER]
GO

