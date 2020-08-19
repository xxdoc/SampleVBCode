SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spuCleanAndPopulateAddressInfo]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spuCleanAndPopulateAddressInfo]
GO

/******************************************************************
*
*    	Description:	Clean Address from one Address line
*			Populate parts of Address if possible
*			
*	Author: 	Brad Skidmore
*	Date: 		1/14/2004
*	
******************************************************************/
CREATE PROCEDURE [dbo].[z_spuCleanAndPopulateAddressInfo]
	@Adress		varchar(500) = '' OUTPUT ,
	@Street		varchar(200) = '' OUTPUT ,
	@City		varchar(50) = '' OUTPUT ,
	@State		varchar(50) = '' OUTPUT ,
	@Zip		int = 0 OUTPUT ,
	@Zip4		int = 0 OUTPUT,
	@OtherPostCode	varchar(20) = '' OUTPUT 
AS
--Need to parse out any ZIP 4 items
DECLARE	@sTempZIP	varchar(20)
DECLARE @sTempZip4	varchar(20)

--Create A temp table to Hold Cleaned Address info
Create Table #TempAddress	(
				Address varchar(500),
				Street	varchar(200),
				City	varchar(50),
				State	varchar(50),
				Zip	varchar(20)
				)

INSERT #TempAddress EXEC spsCleanAddressInfo  @Adress

DECLARE RSCleanAddress Cursor for
SELECT * FROM #TempAddress

OPEN RSCleanAddress

Fetch Next From RSCleanAddress into @Adress, @Street, @City, @State, @sTempZIP

Close RSCleanAddress
Deallocate RSCleanAddress

Drop Table #TempAddress

--Clean the Zip Code
SET @sTempZip = REPLACE(@sTempZip,'`','')
SET @sTempZIP = REPLACE(@sTempZIP,'K','')
--Check For ZIP4
IF CHARINDEX('-',@sTempZip) > 0
BEGIN	
	SET @sTempZip4 = @sTempZip
	SET @sTempZip = Left(@sTempZip4, CHARINDEX('-',@sTempZip4)-1)
	SET @sTempZip4 = Replace(@sTempZip4, @sTempZip + '-','') 
END
IF IsNumeric(@sTempZip) > 0 
BEGIN
	SET @Zip = Convert(int,LTrim(RTrim(@sTempZip)))
END
IF IsNumeric(@sTempZip4) > 0
BEGIN
	SET @Zip4 = Convert(int,LTrim(RTrim(@sTempZip4)))
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[z_spuCleanAndPopulateAddressInfo]  TO [CFUSER]
GO

