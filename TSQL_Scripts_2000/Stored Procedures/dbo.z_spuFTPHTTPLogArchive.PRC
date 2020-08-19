SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spuFTPHTTPLogArchive]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spuFTPHTTPLogArchive]
GO

CREATE PROCEDURE [dbo].[z_spuFTPHTTPLogArchive] AS

--First get the count of records from the tables that
--are about to be archived.
DECLARE @CountOfFTPArch	int
DECLARE @CountOFHTTPArch int

SELECT 	@CountOfFTPArch = 	(	
				SELECT 	COUNT([FTPLogID])
				FROM 	FTPlog
				WHERE 	[LogTime] < getdate()-90 --Select records older than 90 days)
				),
	@CountOFHTTPArch =	(	
				SELECT 	COUNT([HTTPLogID])
				FROM 	HTTPlog
				WHERE 	[LogTime] < getdate()-90 --Select records older than 90 days)
				)

     INSERT INTO FTPLogArchive(
				[FTPLogID] ,
				[ClientHost],
				[username],
				[LogTime] ,
				[service] ,
				[machine] ,
				[serverip] ,
				[processingtime],
				[bytesrecvd] ,
				[bytessent] ,
				[servicestatus] ,
				[win32status] ,
				[operation] ,
				[target] ,
				[parameters]
				)
         	SELECT 		
				[FTPLogID] ,
				[ClientHost],
				[username] ,
				[LogTime] ,
				[service] ,
				[machine] ,
				[serverip] ,
				[processingtime],
				[bytesrecvd] ,
				[bytessent] ,
				[servicestatus],
				[win32status] ,
				[operation] ,
				[target] ,
				[parameters]
         	FROM FTPlog
		WHERE LogTime < getdate()-90 --Select records older than 90 days
		--Then Delete from HTTP Log
		DELETE FROM FTPLog 
		WHERE LogTime < getdate()-90 --Select records older than 90 days


	--Do the http here too instead of in the http update
	INSERT INTO HTTPLogArchive(
				[HTTPLogID] ,
				[ClientHost],
				[username],
				[LogTime] ,
				[service] ,
				[machine] ,
				[serverip] ,
				[processingtime],
				[bytesrecvd] ,
				[bytessent] ,
				[servicestatus] ,
				[win32status] ,
				[operation] ,
				[target] ,
				[parameters]
				)
         	SELECT 		
				[HTTPLogID] ,
				[ClientHost],
				[username] ,
				[LogTime] ,
				[service] ,
				[machine] ,
				[serverip] ,
				[processingtime],
				[bytesrecvd] ,
				[bytessent] ,
				[servicestatus],
				[win32status] ,
				[operation] ,
				[target] ,
				[parameters]
         	FROM HTTPlog
		WHERE LogTime < getdate()-90 --Select records older than 90 days
		--Then Delete from HTTP Log
		DELETE FROM HTTPLog 
		WHERE LogTime < getdate()-90 --Select records older than 90 days
--Now return the same count that was just Archived
SELECT 	@CountOfFTPArch 	As [CountOfFTPArch],
	@CountOFHTTPArch	As [CountOFHTTPArch]

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[z_spuFTPHTTPLogArchive]  TO [CFUSER]
GO

