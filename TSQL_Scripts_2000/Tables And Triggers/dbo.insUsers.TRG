SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[insUsers]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[insUsers]
GO

CREATE TRIGGER insUsers
ON dbo.Users
AFTER INSERT
AS
--1 After Adding a User Also Add the UserID to AdjusterUsersSoftware Table 
insert into AdjusterUsersSoftware (
	[UsersID],
	[IBPrefix],
	[DateLastUpdated],
	[UpdateByUserID]
	)
SELECT
	INS.[UsersID],
	Left(INS.[FirstName],1) + Left(INS.[LastName],1) As IBPrefix, --Trye to use the First I and Last I as default IB Prefix
	INS.[DateLastUpdated],
	INS.[UpdateByUserID]
FROM 	INSERTED INS

--2 After Adding a User Also Add the UserID to AdjusterUsersUpdates Table 
insert into AdjusterUsersUpdates (
	[UsersID],
	[DateLastUpdated],
	[UpdateByUserID]
	)
SELECT
	INS.[UsersID],
	INS.[DateLastUpdated],
	INS.[UpdateByUserID]
FROM 	INSERTED INS

--3 After Adding a User Also Add the UserID to ECSADJUSers Table 
insert into ECSADJUsers (
	[UsersID]
	)
SELECT
	INS.[UsersID]
FROM 	INSERTED INS

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

