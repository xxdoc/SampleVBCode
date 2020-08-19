SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spuFillAddressFields]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spuFillAddressFields]
GO

/******************************************************************
*
*    	Description:	Populate Address Field Outputs
*			
*	Author: 	Brad Skidmore
*	Date: 		4/7/2004
*	
******************************************************************/
CREATE PROCEDURE [dbo].[z_spuFillAddressFields]
	@psAddress		varchar(500) = ''		,
	@psStreet		varchar(200) = ''	OUTPUT ,
	@psCity			varchar(50) = '' 	OUTPUT ,
	@psState		varchar(50) = '' 	OUTPUT ,
	@psZip			varchar(50) = '' 	OUTPUT 
	
AS
DECLARE @bGoToAddress	bit
DECLARE @S_z		varchar(3)
DECLARE @sTemp 		varchar(1000)
DECLARE @sAddress 	varchar(1000)
DECLARE @sValTemp 	varchar(1000)
DECLARE @lPos 		int
DECLARE @VBCRLF 	varchar(2)
DECLARE @VBNullString 	varchar(1)
	
--Set Vars    
SET @S_z = '¶Ññ' --Begin SQL String Field
SET @bGoToAddress = 0
SET @VBCRLF = dbo.GetVBCRLF()
SET @VBNullString = ''
SET @sAddress = LTRIM(RTRIM(Replace(@psAddress, @VBCRLF, @VBNullString)))
    
--Zip code
IF CHARINDEX(' ', @sAddress, 1) > 0 And @bGoToAddress = 0
BEGIN	
	SET @lPos = dbo.InStrRev(@sAddress, ' ', default)
	SET @sTemp = LTrim(RTrim(SUBSTRING(@sAddress, @lPos,DATALENGTH(@sAddress))))
	SET @sValTemp = Replace(@sTemp, '-', @vbNullString)
        IF ISNUMERIC(@sValTemp) > 0 
	BEGIN
    		IF DATALENGTH(@sTemp) >= 5 
		BEGIN
	                SET @lPos = dbo.InStrRev(@sAddress, @sTemp, default) - 1
	                IF @lPos >= 0 
			BEGIN
	                    	SET @sAddress = LTrim(RTrim(Left(@sAddress, @lPos)))
	                END
	                SET @sTemp = Replace(@sTemp, ',', @vbNullString)
	                SET @psZip = @sTemp
		END
            	ELSE
		BEGIN
	                SET @psZip = @vbNullString
	                SET @psState = @vbNullString
	                SET @psCity = @vbNullString
	                SET @bGoToAddress = 1
            	END
	END
        ELSE
	BEGIN
        	SET @psZip = @vbNullString
            	SET @psState = @vbNullString
            	SET @psCity = @vbNullString
           	SET @bGoToAddress = 1
        END
END
ELSE
BEGIN
        SET @psZip = @vbNullString
        SET @psState = @vbNullString
        SET @psCity = @vbNullString
        SET @bGoToAddress = 1
END

    
--State
IF DATALENGTH(@sAddress) > 2 And @bGoToAddress = 0
BEGIN	
        SET @sTemp = Right(@sAddress, 2)
        IF ISNUMERIC(@sTemp) = 0 
	BEGIN
		SET @lPos = dbo.InStrRev(@sAddress, @sTemp, default) - 1
		IF @lPos >= 0 
		BEGIN
			SET @sAddress = LTrim(RTrim(Left(@sAddress, @lPos)))
		END
		SET @sTemp = Replace(@sTemp, ',', @vbNullString)
		SET @psState = @sTemp
	END
        ELSE
	BEGIN
            SET @psState = @vbNullString
            SET @psCity = @vbNullString
            SET @bGoToAddress = 1
       	END
END
ELSE
BEGIN	
	SET @psState = @vbNullString
	SET @psCity = @vbNullString
	SET @bGoToAddress = 1
END
    
--City
IF CHARINDEX(@S_z, @sAddress, 1) > 0 And @bGoToAddress = 0
BEGIN
	SET @lPos = dbo.InStrRev(@sAddress, @S_z, default)
        SET @sTemp = LTrim(RTrim(SUBSTRING(@sAddress, @lPos, DATALENGTH(@sAddress))))
        IF ISNUMERIC(@sTemp) = 0 
	BEGIN
           	SET @lPos = dbo.InStrRev(@sAddress, @sTemp, default) - 1
            	IF @lPos >= 0 
		BEGIN
                	SET @sAddress = Ltrim(RTrim(left(@sAddress, @lPos)))
            	END
            	SET @sTemp = Replace(@sTemp, ',', @vbNullString)
           	SET @sTemp = Replace(@sTemp, @S_z, @vbNullString)
            	SET @sTemp = Replace(@sTemp, Char(32), Char(160))
            	SET @psCity = @sTemp
	END
        ELSE
	BEGIN
            SET @psCity = @vbNullString
        END
END
ELSE IF CHARINDEX(REPLICATE(Char(32), 2), @sAddress, 1) > 0 And @bGoToAddress = 0
BEGIN
	SET @lPos = dbo.InStrRev(@sAddress, REPLICATE(Char(32), 2), default)
        SET @sTemp = LTrim(RTrim(SUBSTRING(@sAddress, @lPos, DATALENGTH(@sAddress))))
        IF ISNUMERIC(@sTemp) = 0
	BEGIN
            	SET @lPos = dbo.InStrRev(@sAddress, @sTemp, default) - 1
            	IF @lPos >= 0 
		BEGIN
                	SET @sAddress = LTrim(RTrim(left(@sAddress, @lPos)))
            	END
            	SET @sTemp = Replace(@sTemp, ',', @vbNullString)
            	SET @sTemp = Replace(@sTemp, @S_z, @vbNullString)
           	SET @sTemp = Replace(@sTemp, Char(32), Char(160))
            	SET @psCity = @sTemp
	END
        ELSE
	BEGIN
            	SET @psCity = @vbNullString
        END
END
ELSE IF CHARINDEX(REPLICATE (Char(32), 1), @sAddress, 1) > 0 And @bGoToAddress = 0
BEGIN
	SET @lPos = dbo.InStrRev(@sAddress, REPLICATE(Char(32), 1), default)
        SET @sTemp = Ltrim(RTrim(SUBSTRING(@sAddress, @lPos, DATALENGTH(@sAddress))))
        IF ISNUMERIC(@sTemp) = 0 
	BEGIN
           	SET @lPos = dbo.InStrRev(@sAddress, @sTemp, default) - 1
            	IF @lPos >= 0 
		BEGIN
                	SET @sAddress = LTrim(RTrim(left(@sAddress, @lPos)))
            	END
		SET @sTemp = Replace(@sTemp, ',', @vbNullString)
		SET @sTemp = Replace(@sTemp, @S_z, @vbNullString)
		SET @sTemp = Replace(@sTemp, Char(32), Char(160))
		SET @psCity = @sTemp
	END
        ELSE
	BEGIN
            SET @psCity = @vbNullString
	END
END
ELSE
BEGIN
	SET @psCity = @vbNullString
END

--Address
SET @sAddress = Replace(@sAddress, ',', @vbNullString)
SET @sAddress = Replace(@sAddress, @S_z, @vbNullString)
SET @psStreet = @sAddress

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[z_spuFillAddressFields]  TO [CFUSER]
GO

