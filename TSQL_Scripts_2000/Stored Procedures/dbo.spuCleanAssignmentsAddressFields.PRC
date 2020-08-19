SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spuCleanAssignmentsAddressFields]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[spuCleanAssignmentsAddressFields]
GO

/******************************************************************
*
*    	Description:	VS1 Conversion
*			Clean Assignments Address Fields
*	
*	Author: 	Brad Skidmore
*	Date: 		07/16/2003
*	Updated		4/7/2004 Use SP FillAddressFields and UpdateAddress
******************************************************************/


CREATE PROCEDURE [dbo].[spuCleanAssignmentsAddressFields] 
AS

DECLARE @AssignmentsID int
DECLARE @psAddress varchar(500)
DECLARE @psZip varchar(20)
DECLARE @psZip4 varchar(20)
DECLARE @psState varchar(50)
DECLARE @psCity varchar(50)
DECLARE @psStreet varchar(200)

--1. Update the Mailing Address fields 
Declare MyMailingAddress Cursor for
	SELECT 	AssignmentsID,
		MailingAddress
	from 	Assignments
	where 	MailingAddress is not null
Open MyMailingAddress

Fetch Next From MyMailingAddress into @AssignmentsID, @psAddress

WHile @@FEtch_Status = 0
	Begin
		SET @psZip = '0'
		SET @psZip4 = '0'
		SET @psState = ''
		SET @psCity = ''
		SET @psStreet = ''
		exec z_spuFillAddressFields
			@psAddress ,
			@psStreet output,
			@psCity output,
			@psState output,
			@psZip output
		exec z_spuUpdateAddress
			@psAddress output,
			@psStreet,
			@psCity ,
			@psState ,
			@psZip 

		--Clean the Zip Code
		SET @psZip = REPLACE(@psZip,'`','')
		SET @psZIP = REPLACE(@psZIP,'K','')
		--Check For ZIP4
		IF CHARINDEX('-',@psZip) > 0
		BEGIN	
			SET @psZip4 = @psZip
			SET @psZip = Left(@psZip4, CHARINDEX('-',@psZip4)-1)
			SET @psZip4 = Replace(@psZip4, @psZip + '-','') 
		END
		UPDATE Assignments SET  
					MailingAddress = UPPER(@psAddress),
					MAStreet = UPPER(@psStreet),
					MACity = UPPER(@psCity),
					MAState = UPPER(@psState),
					MAZIP = @psZip,
					MAZIP4 = @psZip4
		WHERE
			AssignmentsID = @AssignmentsID
						
		Fetch Next From MyMailingAddress into @AssignmentsID, @psAddress
	End

Close MyMailingAddress
Deallocate MyMailingAddress

--2. Update the Property Address fields 
Declare MyPropertyAddress Cursor for
	SELECT 	AssignmentsID,
		PropertyAddress
	from 	Assignments
	where 	MailingAddress is not null
Open MyPropertyAddress

Fetch Next From MyPropertyAddress into @AssignmentsID, @psAddress

WHile @@FEtch_Status = 0
	Begin
		SET @psZip = '0'
		SET @psZip4 = '0'
		SET @psState = ''
		SET @psCity = ''
		SET @psStreet = ''
		exec z_spuFillAddressFields
			@psAddress ,
			@psStreet output,
			@psCity output,
			@psState output,
			@psZip output
		exec z_spuUpdateAddress
			@psAddress output,
			@psStreet,
			@psCity ,
			@psState ,
			@psZip 

		--Clean the Zip Code
		SET @psZip = REPLACE(@psZip,'`','')
		SET @psZIP = REPLACE(@psZIP,'K','')
		--Check For ZIP4
		IF CHARINDEX('-',@psZip) > 0
		BEGIN	
			SET @psZip4 = @psZip
			SET @psZip = Left(@psZip4, CHARINDEX('-',@psZip4)-1)
			SET @psZip4 = Replace(@psZip4, @psZip + '-','') 
		END
		UPDATE Assignments SET  
					PropertyAddress = UPPER(@psAddress),
					PAStreet = UPPER(@psStreet),
					PACity = UPPER(@psCity),
					PAState = UPPER(@psState),
					PAZIP = @psZip,
					PAZIP4 = @psZip4
		WHERE
			AssignmentsID = @AssignmentsID
						
		Fetch Next From MyPropertyAddress into @AssignmentsID, @psAddress
	End

Close MyPropertyAddress
Deallocate MyPropertyAddress

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

