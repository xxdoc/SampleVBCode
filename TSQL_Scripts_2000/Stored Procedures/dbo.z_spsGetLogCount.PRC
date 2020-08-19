SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spsGetLogCount]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spsGetLogCount]
GO

/******************************************************************
*
*    	Description:	Get Log Count
*			Returns the Total Record Counts for
			HTTP (Archive and Current)
			FTP (Archive and Current)
			Total Records for All Log Tables
*	
*	Author: 	Brad Skidmore
*	Date: 		4/26/2004
*
******************************************************************/
CREATE PROCEDURE [dbo].[z_spsGetLogCount]
AS

Select 	LogCounts.CountOFHTTPLogArchive,
	LogCounts.CountOFHTTPLog,
	LogCounts.CountOFFTPLogArchive,
	LogCounts.CountOFFTPLog,
	(CountOFHTTPLogArchive + CountOFHTTPLog + CountOFFTPLogArchive + CountOFFTPLog) As TotalCount
FROM 
	(SELECT	(Select count(HTTPLogArchiveID) from HTTPLogArchive)As CountOFHTTPLogArchive,
		(Select count(HTTPLogID) from HTTPLog )As CountOFHTTPLog,
		(Select count(FTPLogArchiveID) from FTPLogArchive )As CountOFFTPLogArchive,
		(Select count(FTPLogID) from FTPLog )As CountOFFTPLog
	) As LogCounts

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[z_spsGetLogCount]  TO [CFUSER]
GO

