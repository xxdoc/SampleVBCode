SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spuUpdateAddress]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spuUpdateAddress]
GO

/******************************************************************
*
*    	Description:	Update Address String
*			
*	Author: 	Brad Skidmore
*	Date: 		4/7/2004
*	
******************************************************************/
CREATE PROCEDURE [dbo].[z_spuUpdateAddress]
	@psAddress		varchar(500) = ''	OUTPUT	,
	@psStreet		varchar(200) = ''		,
	@psCity			varchar(50) = '' 		,
	@psState		varchar(50) = '' 		,
	@psZip			varchar(50) = '' 		 
	
AS
DECLARE @lPos 		int
DECLARE @sTemp 		varchar(1000)
DECLARE @sAddress 	varchar(1000)
DECLARE @VBCRLF 	varchar(2)
DECLARE @VBNullString 	varchar(1)
	
--Set Vars    
SET @VBCRLF = dbo.GetVBCRLF()
SET @VBNullString = ''
SET @sAddress = LTRIM(RTRIM(Replace(@psAddress, @VBCRLF, @VBNullString)))

--Update the Address string       
SET @sTemp = LTrim(RTrim(Replace(@psStreet, @vbCrLf, @vbNullString))) + REPLICATE(' ', 2) + @vbCrLf
SET @sAddress = @sTemp
SET @sTemp = LTrim(RTrim(Replace(@psCity, @vbCrLf, @vbNullString)))
SET @sTemp = Replace(@sTemp, ',', @vbNullString) + ', '
SET @sAddress = @sAddress + @sTemp
SET @sTemp = LTrim(RTrim(Replace(@psState, @vbCrLf, @vbNullString))) + ' '
SET @sAddress = @sAddress + @sTemp
SET @sTemp = LTrim(RTrim(Replace(@psZip, @vbCrLf, @vbNullString)))
SET @sAddress = @sAddress + @sTemp
    
     
IF Right(@sAddress, 5) = @vbCrLf + ', ' + ' ' 
BEGIN
	SET @lPos = dbo.InStrRev(@sAddress, @vbCrLf, default) - 1
	IF @lPos > 0
	BEGIN
		SET @sAddress = RTrim(left(@sAddress, @lPos))
	END
END
    
SET @psAddress = @sAddress

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[z_spuUpdateAddress]  TO [CFUSER]
GO

