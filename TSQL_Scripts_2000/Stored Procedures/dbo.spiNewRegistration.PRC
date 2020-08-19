SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spiNewRegistration]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[spiNewRegistration]
GO

CREATE PROCEDURE [dbo].[spiNewRegistration] 
@FirstName 	Varchar(50),
@Middle		Varchar(5),
@LastName  	Varchar(50),
@SSNum		Varchar(15),
@Address	Varchar(50),
@City		Varchar(50),
@State		Varchar(50),
@Phone		Varchar(50),
@Zip		Varchar(50),
@EntryDate	DateTime,
@Banquet	Varchar(50),
@Spouse		Varchar(50),
@Lunch		Varchar(50),
@Thurs		Varchar(50),
@FriAM		Varchar(50),
@FriPM		Varchar(50),
@SatAM		Varchar(50),
@SatPM		Varchar(50),
@SunAM		Varchar(50),
@Texas		Varchar(50),
@Oklahoma	Varchar(50),
@Mississippi	Varchar(50),
@NorthCarolina	Varchar(50),
@NewHampshire	Varchar(50),
@Georgia	Varchar(50),
@Wyoming	Varchar(50),
@Florida	Varchar(50),
@TotalDue	Varchar(50),
@hBillingName	Varchar(50),
@hBillingAddr	Varchar(50),
@hCity		Varchar(50),
@hState		Varchar(50),
@hZip		Varchar(50),
@hPhone		Varchar(50),
@hRoom1A	Varchar(50),
@hRoom1B	Varchar(50),
@hRoom1C	Varchar(50),
@hRoom1D	Varchar(50),
@hRoom2A	Varchar(50),
@hRoom2B	Varchar(50),
@hRoom2C	Varchar(50),
@hRoom2D	Varchar(50),
@hRoom3A	Varchar(50),
@hRoom3B	Varchar(50),
@hRoom3C	Varchar(50),
@hRoom3D	Varchar(50),
@hArrivalDate	DateTime,
@hDepartureDate DateTime,
@hNonSmoking	bit,
@hSmoking	bit,
@hKingBed	bit,
@h2QueenBeds	bit,
@hConnectingRms bit,
@hWheelchair	bit,
@hCCType	Varchar(50),
@hCCNumber	Varchar(50),
@hCCExpiration  Varchar(50),
@hNumberOfRms	numeric(9)

AS

INSERT INTO SEMI04
	(FirstName, Middle, LastName, SSNum, Address, City, State,
	 Phone, Zip, DAT__DATE9, Banquet, Spouse, Lunch, Thurs,
	 FriAM, FriPM, SatAM, SatPM, SunAM, Texas, Oklahoma,
	 Mississippi, NorthCarolina, NewHampshire, Georgia,
	 Wyoming, Florida, TotalDue, 
	 hBillingName, hBillingAddr, hCity, hState, hZip,
	 hPhone, hRoom1A, hRoom1B, hRoom1C, hRoom1D,
	 hRoom2A, hRoom2B, hRoom2C, hRoom2D, 
	 hRoom3A, hRoom3B, hRoom3C, hRoom3D,
	 hArrivalDate, hDepartureDate, hNonSmoking, hSmoking,
	 hKingBed, h2QueenBeds, hConnectingRms, hWheelchair,
	 hCCType, hCCNumber, hCCExpiration, hNumberOfRms)
VALUES
	(@FirstName, @Middle, @LastName, @SSNum, @Address, @City, @State,
 	 @Phone, @Zip, @EntryDate, @Banquet, @Spouse, @Lunch, @Thurs,
	 @FriAM, @FriPM, @SatAM, @SatPM, @SunAM, @Texas, @Oklahoma,
	 @Mississippi, @NorthCarolina, @NewHampshire, @Georgia,
	 @Wyoming, @Florida, @TotalDue, 
	 @hBillingName, @hBillingAddr, @hCity, @hState, @hZip,
	 @hPhone, @hRoom1A, @hRoom1B, @hRoom1C, @hRoom1D,
	 @hRoom2A, @hRoom2B, @hRoom2C, @hRoom2D,
	 @hRoom3A, @hRoom3B, @hRoom3C, @hRoom3D,
	 @hArrivalDate, @hDepartureDate, @hNonSmoking, @hSmoking,
	 @hKingBed, @h2QueenBeds, @hConnectingRms, @hWheelchair,
	 @hCCType, @hCCNumber, @hCCExpiration, @hNumberOfRms)

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

