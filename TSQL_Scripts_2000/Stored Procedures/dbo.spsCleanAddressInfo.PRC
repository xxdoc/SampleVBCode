SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spsCleanAddressInfo]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[spsCleanAddressInfo]
GO

/******************************************************************
*
*    	Description:	Return Clean Address Info
*			
*	
*	Author: 	Brad Skidmore
*	Date: 		07/16/2003
*	Updated		4/7/2004 Use SP FillAddressFields and UpdateAddress
******************************************************************/


CREATE PROCEDURE [dbo].[spsCleanAddressInfo] 
	@sOLDAddress varchar(500)
AS

DECLARE @psAddress varchar(500)
DECLARE @psZip varchar(20)
DECLARE @psState varchar(50)
DECLARE @psCity varchar(50)
DECLARE @psStreet varchar(200)

SET @psAddress = @sOldAddress
--1. First Get the Address Info parts
exec z_spuFillAddressFields
	@psAddress ,
	@psStreet output,
	@psCity output,
	@psState output,
	@psZip output

--2. Next update the Address String
exec z_spuUpdateAddress
	@psAddress output,
	@psStreet,
	@psCity ,
	@psState ,
	@psZip 

--3. Return The Cleaned Address parts
SELECT Address = UPPER(@psAddress),
	Street = UPPER(@psStreet),
	City = UPPER(@psCity),
	State = UPPER(@psState), 
	ZIP = @psZip

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[spsCleanAddressInfo]  TO [CFUSER]
GO

