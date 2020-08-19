SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spuUpdateNewAssignments]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[spuUpdateNewAssignments]
GO

/******************************************************************
*
*    	Description:	Update Assignments that were just imported from VS1
*			VS1 to VS2 Conversion
*	
*	Author: 	Brad Skidmore
*	Date: 		07/17/2003
*	Updated		10/28/2003
******************************************************************/


CREATE PROCEDURE [dbo].[spuUpdateNewAssignments]
	
AS
Update assignments set 

	[ID] = (CASE 	WHEN [ID] is NULL 
			THEN assignmentsID
			ELSE [ID]
			END
		),
	[IBNUM] = (CASE 	WHEN LEFT(IBNUM,3) ='***'
				THEN 	(	SELECT distinct top 1 Co.CODE + cast(assignmentsID as varchar(20))
						FROM Company Co RIGHT JOIN ClientCompanyCatSpec CCS on CO.CompanyID = CCS.ClientCompanyID 
						WHERE CCS.Clientcompanycatspecid = A.Clientcompanycatspecid
					)
				ELSE IBNUM
				END
		),
	[SPVersion] = (CASE	WHEN [SPVersion] Is null
				THEN	(
						SELECT SP.SPVersion
						FROM ((Assignments ASS INNER JOIN ClientCompanyCatSpec CCS ON ASS.ClientCompanyCatSpecID = CCS.ClientCompanyCatSpecID)
							INNER JOIN ClientCompanyCat CC ON CCS.ClientCompanyID = CC.ClientCompanyID And CCS.CATID = CC.CATID)
							INNER JOIN SoftwarePackage SP  ON (SP.ClientCompanyID = CC.ClientCompanyID And SP.CATID = CC.CATID)
						WHERE Ass.AssignmentsID = A.AssignmentsID
					) 
				ELSE SPVersion
				END
		), 
	[AppDedClassTypeIDOrder] = (CASE 	WHEN AppDedClassTypeIDOrder is NULL
						THEN '1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26'
						ELSE AppDedClassTypeIDOrder
						END
				)
FROM Assignments A


--2. Update assignMents Address fields
exec spuCleanAssignmentsAddressFields
--Get rif of history after doing this VS1 updated
truncate table AssignmentsHistory

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

