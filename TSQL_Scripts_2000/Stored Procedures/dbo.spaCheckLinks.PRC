SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spaCheckLinks]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[spaCheckLinks]
GO

Create Procedure spaCheckLinks
	@Links	Varchar(500),
	@lCheckV2ECcarFarmersclsLossXML01Exports bit =0
As

Set NoCount On

Declare @CurrentLink	Varchar(500)
Declare @CurrentEnd	Integer
Declare @CurrentStart	Integer
Declare @TotalResult	Integer
Declare @CurrentDB	Varchar(100)
Declare @CurrentTable	Varchar(100)
Declare @CommaPos	Integer
Declare @Counter	Integer

Set @CurrentStart = 0
Set @CurrentEnd = 0
Set @TotalResult = 1

While @CurrentEnd < Len(@Links)
	Begin
		If charIndex('|',@Links,@CurrentStart) > 0
			Begin
				Set @CurrentEnd = charIndex('|',@Links,@CurrentStart) + 1
				Set @CurrentLink = Substring(@Links,@CurrentStart,@CurrentEnd - @CurrentStart - 1)
				Set @CurrentStart = @CurrentEnd
			End
		Else
			Begin
				Set @CurrentEnd = Len(@Links)
				Set @CurrentLink = Substring(@Links,@CurrentStart,@CurrentEnd - @CurrentStart + 1)
			End

		--See if there is a table with the link
		Set @CommaPos = CharIndex(',',@CurrentLink)
		If @CommaPos > 0
			Begin
				Set @CurrentDB = RTrim(LTrim(SubString(@CurrentLink,1,@CommaPos - 1)))
				Set @CurrentTable = RTrim(LTrim(SubString(@CurrentLink,@CommaPos + 1,Len(@CurrentLink) - @CommaPos + 1)))
			End
		Else
			Begin
				Set @CurrentDB = RTrim(LTrim(@CurrentLink))
				Set @CurrentTable = ''
			End
		
		--Make sure the linked server really is linked
		If Exists(Select * from master..sysservers where srvname = @CurrentDB)
			Begin
				If @CurrentTable = ''
					Begin
						Select @Counter = Count(*) From master..SYSREMOTE_Tables (@CurrentDB)
					End
				Else
					Begin
						Select @Counter = Count(*) From master..SYSREMOTE_Tables (@CurrentDB) Where TABLE_NAME = @CurrentTable
					End
	
				If @Counter = 0
					Begin
						Set @TotalResult = 0
						Set @CurrentEnd =  Len(@Links)
					End
			End
		Else
			Begin
				Set @TotalResult = 0
				Set @CurrentEnd =  Len(@Links)
			End
	End

Set NoCount Off

Select @TotalResult As Valid

--'5.18.2005 BGS Check for XML Exports only if flagged
--ONly return a recordset if flagged !!!
-- Farmers Class 5.18.2005
IF @lCheckV2ECcarFarmersclsLossXML01Exports = 1
BEGIN
	--Currently V2ECcarFarmers.clsLossXML01 needs to check the following tables
	--5.18.2005	FarmersXML01_LocationUpdateQue
	--5.18.2005	FarmersXML01_ActivityUpdateQue
	DECLARE @SendToCRN as bit 

	SELECT 	TOP 1 @SendToCRN = [SendToCRN]
	FROM	FarmersXML01_LocationUpdateQue
	WHERE	[SendToCRN] = 1
	
	-- If no Exports are found for FarmersXML01_LocationUpdateQue 
	-- Then check FarmersXML01_ActivityUpdateQue
	IF @SendToCRN Is Null
	BEGIN
		SELECT 	TOP 1 @SendToCRN = [SendToCRN]
		FROM	FarmersXML01_ActivityUpdateQue
		WHERE	[SendToCRN] = 1
	END
	
	IF @SendToCRN Is Null
	BEGIN
		SELECT 0 As [bV2ECcarFarmersclsLossXML01Exports]
	END
	ELSE
	BEGIN
		SELECT @SendToCRN As [bV2ECcarFarmersclsLossXML01Exports]
	END	
END
--10.7.2005 Scrub IBFee
Exec spuScrubIBFee

--10.7.2005 Scrub INsured name with "/" in it
UPDATE Assignments SET 
	[Insured] = Replace([Insured],'/',' '), 
	[DownloadMe] = 1
WHERE 	[Insured] Like '%/%'
AND 	[DateLastUpdated] >= DateAdd(dd,-1,GetDate())


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[spaCheckLinks]  TO [CFUSER]
GO

