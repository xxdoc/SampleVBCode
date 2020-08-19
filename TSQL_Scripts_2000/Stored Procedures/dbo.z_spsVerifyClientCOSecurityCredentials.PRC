SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spsVerifyClientCOSecurityCredentials]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spsVerifyClientCOSecurityCredentials]
GO

/******************************************************************
*
*    	Description:	Output True Or False If Security Credentials Pass
*			Uses Function VerifyClientCOSecurityCredentials
*			Out Puts @VerifiedCredentials 1 or 0
*	
*	Author: 	Brad Skidmore
*	Date: 		5/11/2005
*
******************************************************************/


CREATE PROCEDURE [dbo].[z_spsVerifyClientCOSecurityCredentials]
	@UsersID 		int,
	@CompanyCode 		varchar(3),
	@VerifiedCredentials 	Bit=0 OUTPUT
AS
--Clean Dates
SET @VerifiedCredentials = dbo.VerifyClientCOSecurityCredentials(@UsersID, @CompanyCode)

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[z_spsVerifyClientCOSecurityCredentials]  TO [CFUSER]
GO

