SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spsClaimScrub]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[spsClaimScrub]
GO

/******************************************************************
*
*    	Description:	Populates Multiple Lookup RS for Scrubbing
*			Batches Table
*	
*	Author: 	Brad Skidmore
*	Date: 		02/26/2003
*			09/16/2005 -- UPDATED TO CHECK FOR BLANK CITY NAMES... 
					
******************************************************************/

CREATE PROCEDURE [dbo].[spsClaimScrub] 
AS

--09/16/2005  Correct for blank city names
UPDATE Batches SET losscity = 	(
				SELECT [MACity] 
				FROM Assignments 
				WHERE AssignmentsID = B.AssignmentsID
				)
FROM 	Batches B
WHERE 	Len(B.[losscity]) < 2 
AND 	B.[lossstate] =	(
			SELECT [MAState] 
			FROM Assignments 
			WHERE AssignmentsID = B.AssignmentsID
			)

SELECT
	ECSADJ.*
FROM
	OpenQuery(ClaimProdV1,'SELECT 	FIRST_NAME, 
						LAST_NAME, 
						SITE, SITE_STATE, 
						SS_NUM 
				FROM 		ECSADJ') 
	as ECSADJ
	Inner Join
	Batches Bat
	On ECSADJ.SS_NUM = Bat.SSN
Where
	Bat.SSN > 0 
AND 
	Bat.COPIED Is Null 
AND 
	Bat.ECUPDATED Is Null 
Or 
	Bat.ECUPDATED = 0


SELECT
	BILLDATA.*
FROM
-- 	OpenQuery(ClaimProdV1,'SELECT 	IB
-- 				 FROM 		BILLDATA') 
	(
		SELECT 'TEMPFIX_11-11-2004' As IB
	)
	as BILLDATA
	Inner Join
	Batches Bat
	On BILLDATA.IB = Bat.IBNUMBER
Where
	Bat.SSN > 0 
AND 
	Bat.COPIED Is Null 
AND 
	Bat.ECUPDATED Is Null 
Or 
	Bat.ECUPDATED = 0


SELECT
	TXTAX.*
FROM
	OpenQuery(ClaimProdV1,'SELECT	CITY_NAME,
						STATE
				 FROM 		TXTAX
				 WHERE LEN(RTRIM(TXTAX.ALIAS)) > 0') 
	as TXTAX


SELECT
	GPTXTAX.*
FROM
	OpenQuery(ClaimProdV1,'SELECT	CITY_NAME,
						STATE
				FROM 		TXTAX
				WHERE LEN(RTRIM(TXTAX.ALIAS)) > 0
				GROUP BY	STATE') 
	as GPTXTAX

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[spsClaimScrub]  TO [CFUSER]
GO

