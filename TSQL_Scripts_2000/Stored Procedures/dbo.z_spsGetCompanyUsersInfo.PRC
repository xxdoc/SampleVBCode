SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spsGetCompanyUsersInfo]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spsGetCompanyUsersInfo]
GO

/******************************************************************
*
*    	Description:	Company Users Info
*			
*	
*	Author: 	Brad Skidmore
*	Date: 		09/26/2003
*	Date:		10/9/2003  BGS
			Added User Client Co, User profiles, 
			User rpt to Manager, user rpts to Coordinator
*	DATE:		1/13/2004 BGS
*			Added FillName for populating dropdowns With Adjuster Names
*	DATE:		1/30/2004 Fixed Bug when viewing User reports to MGR or Coordinator
			and Hide deleted is unchecked.  SP was processing as if the 
			Hidedeleted was never un checked. Fixed that.
*	Date		2/12/2004 
*			Added AdjMaxOpenAssign, This will be used by z_spuInitCCAssignAdjToZipCode
			To help figure out where to Assgn Claims
******************************************************************/

CREATE PROCEDURE [dbo].[z_spsGetCompanyUsersInfo]
	@bHideDeleted		bit =1, 
	@UID			int,
	@UsersID		int=null,
--Current Selected Company
	@selCompany		int=null,
--Current Selected Client Company
	@selClientCompany	int=null,
--Current Selected Profile
	@selUserProfile		int=null,
--Current Selected User Reports To Manager
	@selUserReportsToManager	int=null,
--Current Selected User Reports To Coordinator
	@selUserReportsToCoordinator	int=null,
--Current Selected User Assgn To CAT
	@selUserAssgnToCAT 		int=null,
	@OrderBy		varchar(500)=null,
	@GroupBy		varchar(500)=null,
	@SearchBy		varchar(1000)=null,
--updating the CompanyUSers Table
	@bDelete		bit=0,
	@bAdd			bit=0,
	@bEdit			bit=0,
--Update variables
	@UserName		varchar(20)=null,
	@PassWord		varchar(255)=null,
	@FirstName		varchar(20)=null,
	@LastName		varchar(20)=null,
	@SSN			int=null,
	@Email			varchar(30)=null,
	@ContactPhone		varchar(50)=null,
	@EmergencyPhone		varchar(50)=null,
	@Address		varchar(200)=null,
	@City			varchar(50)=null,
	@State			varchar(50)=null,
	@Zip			int=null,
	@ZIP4			int=null,
	@OtherPostCode		varchar(20)=null,
	@UsersActive		bit=null,
	@Active			bit=null,
	@UsersActiveDate	datetime=Null,
	@ActiveDate		datetime=Null,
	@UsersInactiveDate	datetime=Null,
	@InactiveDate		datetime=Null,
	@UsersSecurityLevel	int=null,
	@SecurityLevel		int=null,
	@UsersComments		varchar(1000)=null,
	@Comments		varchar(255)=null,
	@UsersDateLastUpdated	datetime=Null,
	@DateLastUpdated	datetime=Null,
--ViewHistory
	@bViewHistory		bit=NULL,
	@FromActiveDate		varchar(50)='*',
	@ToActiveDate		varchar(50)='*',
	@FromInactiveDate	varchar(50)='*',
	@ToInactiveDate		varchar(50)='*',
	@FromDate		varchar(50)=Null,
	@ToDate			varchar(50)=Null,
--Edit Adjuster User Software
	@AUSLicenseDaysLeft	smallint=null,
	@AUSIBPrefix		varchar(5)=null,
	@AUSSingleFileSendAuthority Bit=Null,
--Apply Adjuster User Updates
	@bAdjusterUsersUpdatesYesNo bit=Null,
--Assignment Type List
	@AssignmentTypeIDList	varchar(255)=null --Default to first Assignment Type ID in the list when adding 
AS
DECLARE @lFillNameCount int
SET @lFillNameCount = (SELECT MAX(LEN(LastName)) FROM USERS)
SET @lFillNameCount = @lFillNameCount + Len(', ')
SET @lFillNameCount = @lFillNameCount + (SELECT MAX(LEN(FirstName)) FROM USERS)
SET @lFillNameCount = @lFillNameCount + Len('...(')
SET @lFillNameCount = @lFillNameCount + (SELECT MAX(LEN(UserName)) FROM USERS)
SET @lFillNameCount = @lFillNameCount + Len(')')
DECLARE @sSQL1 varchar(8000)
DECLARE @sSQL2 varchar(8000)
DECLARE @sSQL3 varchar(8000)
DECLARE @sSQL4 varchar(8000)
DECLARE @sSQL5 varchar(8000)
DECLARE @sSQL6 varchar(8000)
DECLARE @sSQL7 varchar(8000)
DECLARE @sSQL8 varchar(8000)
DECLARE @sSQL9 varchar(8000)
DECLARE @sSQL10 varchar(8000)
DECLARE @selClientORCompany int
DECLARE @sSQLUP as varchar(8000)
DECLARE @sSQLTemp as varchar(8000)
DECLARE @sSQLTemp2 as varchar(8000)
DECLARE @iCount as int
DECLARE @sUserProfileTable As varchar(50)

--Check for client company or Company
IF @selCompany Is Not Null And @selClientCompany Is Null
BEGIN
	SET @selClientORCompany = @selCompany
END
ELSE IF @selCompany Is Null And @selClientCompany Is Not Null
BEGIN
	SET @selClientORCompany = @selClientCompany
END
ELSE IF @selCompany Is Null And @selClientCompany Is Null
BEGIN
	SET @selClientORCompany = Null
END
ELSE IF @selCompany Is Not Null And @selClientCompany Is Not Null
BEGIN
	SET @selClientORCompany = @selClientCompany
END

DECLARE @CurDate 		DateTime
SET @CurDate = GetDate()
--See if the DateLastupdated is null
IF IsDate(@DateLastUpdated) =0
BEGIN
	SET @DateLastUpdated = GetDate()
END
IF IsDate(@UsersDateLastUpdated) =0
BEGIN
	SET @UsersDateLastUpdated = GetDate()
END
--Clean Dates
SET @FromDate =			dbo.CleanFromOrToDate(@CurDate, @FromDate, 1)
SET @ToDate = 			dbo.CleanFromOrToDate(@CurDate, @ToDate, 0)
--Clean String Dates
SET @FromActiveDate = 		dbo.CleanFromOrToDateString(@CurDate, @FromActiveDate, 1)
SET @ToActiveDate = 		dbo.CleanFromOrToDateString(@CurDate, @ToActiveDate, 0)
SET @FromInactiveDate = 	dbo.CleanFromOrToDateString(@CurDate, @FromInactiveDate, 1)
SET @ToInactiveDate = 		dbo.CleanFromOrToDateString(@CurDate, @ToInactiveDate, 0)



IF @UsersID Is Not Null And @bViewHistory is Null Or (@UsersID Is null And @selCompany Is Null And @selClientCompany Is null)
BEGIN
	IF @UsersID Is Null And @selCompany Is Null And @selClientCompany Is Null
	BEGIN
		SET @sSQL1 =''
		SET @sSQL1 = @sSQL1 + 'SELECT '
		SET @sSQL1 = @sSQL1 + 'U.[UsersID], '
		SET @sSQL1 = @sSQL1 + 'U.[UserName], '
		SET @sSQL1 = @sSQL1 + 'U.[PassWord] , '
		SET @sSQL1 = @sSQL1 + 'U.[FirstName] , '
		SET @sSQL1 = @sSQL1 + 'U.[LastName] , '
		SET @sSQL1 = @sSQL1 + 'U.[LastName] + '', '' + U.[FirstName] + REPLICATE(''.'', ' + convert(varchar(10),@lFillNameCount) + '  - LEN(U.[LastName] + '', '' + U.[FirstName] + ''('' + U.[UserName] + '')'')) + ''('' + U.[UserName] + '')'' As FillName, '
		SET @sSQL1 = @sSQL1 + 'U.[SSN] , '
		SET @sSQL1 = @sSQL1 + 'U.[Email], '
		SET @sSQL1 = @sSQL1 + 'U.[ContactPhone], '
		SET @sSQL1 = @sSQL1 + 'U.[EmergencyPhone], '
		SET @sSQL1 = @sSQL1 + 'U.[Address] , '
		SET @sSQL1 = @sSQL1 + 'U.[City] , '
		SET @sSQL1 = @sSQL1 + 'U.[State] , '
		SET @sSQL1 = @sSQL1 + 'U.[Zip] , '
		SET @sSQL1 = @sSQL1 + 'U.[ZIP4] , '
		SET @sSQL1 = @sSQL1 + 'U.[OtherPostCode], '
		SET @sSQL1 = @sSQL1 + 'U.[Active] As UsersActive, '
		SET @sSQL1 = @sSQL1 + 'U.[ActiveDate] As UsersActiveDate, '
		SET @sSQL1 = @sSQL1 + 'U.[InactiveDate] As UsersInactiveDate, '
		SET @sSQL1 = @sSQL1 + 'U.[SecurityLevel] As UsersSecurityLevel , '
		SET @sSQL1 = @sSQL1 + 'U.[Comments] As UsersComments, '
		SET @sSQL1 = @sSQL1 + 'U.[DateLastUpdated] As UsersDateLastUpdated, '
		SET @sSQL1 = @sSQL1 + 'U.[UpdateByUserID] As UsersUpdateByUserID '
		SET @sSQL1 = @sSQL1 + 'FROM Users U '
		SET @sSQL1 = @sSQL1 + 'WHERE 	U.[UsersID] Is Not Null '
		IF @SearchBy Is Not Null
		BEGIN
			SET @sSQL1 = @sSQL1 + @SearchBy + ' '
		END
		IF @bHideDeleted = 1
		BEGIN
			SET @sSQL1 = @sSQL1 + 'AND U.Active = 1 '
		END
		IF @OrderBy Is Not Null
		BEGIN
			SET @sSQL1 = @sSQL1 + @OrderBy + ' '
		END 
		Print Len(@sSQL1)
		Print @sSQL1
		EXEC(@sSQL1)
	END
	ELSE
	BEGIN
		IF @bDelete=0 And @bAdd=0 And @bEdit=0
		BEGIN
			SELECT 
				U.[UsersID],
				U.[UserName],
				U.[PassWord] ,
				U.[FirstName] ,
				U.[LastName] ,
				U.[LastName] + ', ' + U.[FirstName] + REPLICATE('.', @lFillNameCount - LEN(U.[LastName] + ', ' + U.[FirstName] + '(' + U.[UserName] + ')')) + '(' + U.[UserName] + ')' As FillName, 
				U.[SSN] ,
				U.[Email],
				U.[ContactPhone],
				U.[EmergencyPhone],
				U.[Address] ,
				U.[City] ,
				U.[State] ,
				U.[Zip] ,
				U.[ZIP4] ,
				U.[OtherPostCode],
				U.[Active] As UsersActive,
				CU.[Active],
				U.[ActiveDate] As UsersActiveDate,
				CU.[ActiveDate],
				U.[InactiveDate] As UsersInactiveDate,
				CU.[InactiveDate],
				U.[SecurityLevel] As UsersSecurityLevel ,
				CU.[SecurityLevel] ,
				U.[Comments] As UsersComments,
				CU.[Comments],
				CU.[AssignmentTypeIDList],
				U.[DateLastUpdated] As UsersDateLastUpdated,
				CU.[DateLastUpdated],
				U.[UpdateByUserID] As UsersUpdateByUserID,
				CU.[UpdateByUserID] 
			FROM Users U 	INNER JOIN CompanyUsers CU On U.UsersID = CU.UsersID
					AND CU.CompanyID = @selClientORCompany 
			WHERE 	U.[UsersID] = @UsersID
		END
		ELSE
		BEGIN
	/****************************************
	*		***DELETE***
	****************************************/
			IF @bDelete =1
			BEGIN
				IF @selUserProfile is Null And @selUserReportsToManager is Null And @selUserReportsToCoordinator is Null And @selUserAssgnToCAT is Null 
				BEGIN---Only Update Users Table if not Removing From the Above Child Tables...
					--Only Update Company users Table with delete
					UPDATE CompanyUsers SET [InactiveDate]=(CASE 	WHEN CU.[InactiveDate]Is Null
										THEN Convert(datetime,dbo.FormatDate(GetDate(),'mm/dd/yyy'))
										ELSE Null
										END
									),
								[Active] = 	(CASE 	WHEN CU.[InactiveDate]Is Null
										THEN 0
										ELSE 1
										END
									),
							[DateLastUpdated]=@DateLastUpdated ,
							[UpdateByUserID]=@UID
					FROM Users U 	INNER JOIN CompanyUsers CU On U.UsersID = CU.UsersID
							AND CU.CompanyID = @selClientORCompany
					WHERE 	U.[UsersID] = @UsersID
				END
				ELSE
				BEGIN---See Which Child Table Removing From ...
					IF @selUserProfile is Not Null 
					BEGIN--User Profile 
						SET @sUserProfileTable = ''
						SET @sSQLUP =''
						SET @sUserProfileTable = (SELECT [TableName] FROM UserProfile WHERE UserProfileID = @selUserProfile)
						SET @sSQLUP = @sSQLUP + 'UPDATE ' + @sUserProfileTable + ' SET [Active] = 0, '
						SET @sSQLUP = @sSQLUP + '[DateLastUpdated] =''' + cast(@DateLastUpdated as varchar(50)) + ''' , '
						SET @sSQLUP = @sSQLUP + '[UpdateByUserID] = ' + cast(@UID as varchar(10)) + ' '
						SET @sSQLUP = @sSQLUP + 'FROM CompanyUsers CU INNER JOIN ' + @sUserProfileTable + ' UPT On CU.UsersID = UPT.UsersID '
						SET @sSQLUP = @sSQLUP + 'AND CU.CompanyID = UPT.CompanyID '
						SET @sSQLUP = @sSQLUP + 'WHERE 	CU.[UsersID] = ' + cast(@UsersID as varchar(10)) + ' '
						SET @sSQLUP = @sSQLUP + 'AND  CU.[CompanyID] = ' + Cast(@selClientORCompany as varchar(10)) + ' '
						SET @sSQLUP = @sSQLUP + 'AND UPT.[UsersID] = ' + cast(@UsersID as varchar(10)) + ' ' 
						SET @sSQLUP = @sSQLUP + 'AND  UPT.[CompanyID] = ' + Cast(@selClientORCompany as varchar(10)) + ' '
						Print Len(@sSQLUP) 
						Print @sSQLUP
						EXEC(@sSQLUP)
					END
	
					IF @selUserReportsToManager is Not Null 
					BEGIN--User Reports to Manager 
						SET @sSQLUP =''
						SET @sSQLUP = @sSQLUP + 'UPDATE UserReportsToManager SET [Active] = 0, '
						SET @sSQLUP = @sSQLUP + '[DateLastUpdated] =''' + cast(@DateLastUpdated as varchar(50)) + ''' , '
						SET @sSQLUP = @sSQLUP + '[UpdateByUserID] = ' + cast(@UID as varchar(10)) + ' '
						SET @sSQLUP = @sSQLUP + 'FROM CompanyUsers CU INNER JOIN UserReportsToManager URTM On CU.UsersID = URTM.UsersID '
						SET @sSQLUP = @sSQLUP + 'AND CU.CompanyID = URTM.CompanyID '
						SET @sSQLUP = @sSQLUP + 'WHERE 	URTM.ReportsToUsersID = ' + cast(@selUserReportsToManager as varchar(10)) + ' '
						SET @sSQLUP = @sSQLUP +	'AND CU.[UsersID] = ' + cast(@UsersID as varchar(10)) + ' '
						SET @sSQLUP = @sSQLUP + 'AND  CU.[CompanyID] = ' + Cast(@selClientORCompany as varchar(10)) + ' '
						SET @sSQLUP = @sSQLUP + 'AND URTM.[UsersID] = ' + cast(@UsersID as varchar(10)) + ' ' 
						SET @sSQLUP = @sSQLUP + 'AND  URTM.[CompanyID] = ' + Cast(@selClientORCompany as varchar(10)) + ' '
						Print Len(@sSQLUP) 
						Print @sSQLUP
						EXEC(@sSQLUP)
					END
					
					IF @selUserReportsToCoordinator is Not Null 
					BEGIN--User Reports to Coordinator
						SET @sSQLUP =''
						SET @sSQLUP = @sSQLUP + 'UPDATE UserReportsToCoordinator SET [Active] = 0, '
						SET @sSQLUP = @sSQLUP + '[DateLastUpdated] =''' + cast(@DateLastUpdated as varchar(50)) + ''' , '
						SET @sSQLUP = @sSQLUP + '[UpdateByUserID] = ' + cast(@UID as varchar(10)) + ' '
						SET @sSQLUP = @sSQLUP + 'FROM CompanyUsers CU INNER JOIN UserReportsToCoordinator URTC On CU.UsersID = URTC.UsersID '
						SET @sSQLUP = @sSQLUP + 'AND CU.CompanyID = URTC.CompanyID '
						SET @sSQLUP = @sSQLUP + 'WHERE 	URTC.ReportsToUsersID = ' + cast(@selUserReportsToCoordinator as varchar(10)) + ' '
						SET @sSQLUP = @sSQLUP +	'AND CU.[UsersID] = ' + cast(@UsersID as varchar(10)) + ' '
						SET @sSQLUP = @sSQLUP + 'AND  CU.[CompanyID] = ' + Cast(@selClientORCompany as varchar(10)) + ' '
						SET @sSQLUP = @sSQLUP + 'AND URTC.[UsersID] = ' + cast(@UsersID as varchar(10)) + ' ' 
						SET @sSQLUP = @sSQLUP + 'AND  URTC.[CompanyID] = ' + Cast(@selClientORCompany as varchar(10)) + ' '
						Print Len(@sSQLUP) 

						Print @sSQLUP
						EXEC(@sSQLUP)
					END
					
					IF @selUserAssgnToCAT is Not Null 
					BEGIN--User Assgn To CAT
						SET @sSQLUP =''
						SET @sSQLUP = @sSQLUP + 'UPDATE ClientCompanyUsersCat SET [Active] = 0, '
						SET @sSQLUP = @sSQLUP + '[DateLastUpdated] =''' + cast(@DateLastUpdated as varchar(50)) + ''' , '
						SET @sSQLUP = @sSQLUP + '[UpdateByUserID] = ' + cast(@UID as varchar(10)) + ' '
						SET @sSQLUP = @sSQLUP + 'FROM CompanyUsers CU INNER JOIN ClientCompanyUsersCat CCUC On CU.UsersID = CCUC.UsersID '
						SET @sSQLUP = @sSQLUP + 'AND CU.CompanyID = CCUC.ClientCompanyID '
						SET @sSQLUP = @sSQLUP + 'WHERE 	CCUC.CATID = ' + cast(@selUserAssgnToCAT as varchar(10)) + ' '
						SET @sSQLUP = @sSQLUP +	'AND CU.[UsersID] = ' + cast(@UsersID as varchar(10)) + ' '
						SET @sSQLUP = @sSQLUP + 'AND  CU.[CompanyID] = ' + Cast(@selClientORCompany as varchar(10)) + ' '
						SET @sSQLUP = @sSQLUP + 'AND CCUC.[UsersID] = ' + cast(@UsersID as varchar(10)) + ' ' 
						SET @sSQLUP = @sSQLUP + 'AND  CCUC.[ClientCompanyID] = ' + Cast(@selClientORCompany as varchar(10)) + ' '
						Print Len(@sSQLUP) 
						Print @sSQLUP
						EXEC(@sSQLUP)
					END 
				END 
			END
			ELSE 
			BEGIN
	/****************************************
	*		***ADD***
	****************************************/
				IF @bAdd=1
				BEGIN
					IF 	@selUserProfile is Null 
						And @selUserReportsToManager is Null 
						And @selUserReportsToCoordinator is Null 
						And @selUserAssgnToCAT is Null 
					BEGIN---Only Add to Users Table if not Adding to the above Child Tables...
						DECLARE @iAddThisUsersID int
						IF (SELECT Count([UserName]) FROM Users WHERE [UserName] = @UserName ) = 0 
						BEGIN
							INSERT INTO Users(
										[UserName],
										[PassWord] ,
										[FirstName] ,
										[LastName] ,
										[SSN] ,
										[Email],
										[ContactPhone],
										[EmergencyPhone],
										[Address] ,
										[City] ,
										[State],
										[Zip] ,
										[ZIP4] ,
										[OtherPostCode],
										[Active] ,
										[ActiveDate],
										[InactiveDate],
										[SecurityLevel] ,
										[Comments],
										[DateLastUpdated],
										[UpdateByUserID] 
									)
							SELECT 			@UserName As [UserName],
										@PassWord As [PassWord],
										@FirstName As [FirstName],
										@LastName As [LastName], 
										@SSN As [SSN],
										@Email As [Email],
										@ContactPhone As [ContactPhone],
										@EmergencyPhone As [EmergencyPhone],
										@Address As [Address],
										@City As [City],
										@State As [State],
										@Zip As [Zip],
										@ZIP4 As [ZIP4],
										@OtherPostCode As [OtherPostCode],
										@UsersActive As [Active],
										@UsersActiveDate As [ActiveDate],
										@UsersInactiveDate As [InactiveDate],
										@UsersSecurityLevel As [SecurityLevel],
										@UsersComments As [Comments],
										@UsersDateLastUpdated As [DateLastUpdated],
										@UID As [UpdateByUserID] 

							SET @iAddThisUsersID = IDENT_CURRENT('Users')
						END
						ELSE
						BEGIN
							SET @iAddThisUsersID = (SELECT [UsersID] FROM Users WHERE [UserName] = @UserName)
						END
						IF (SELECT Count([UsersID]) FROM CompanyUsers WHERE [UsersID] = @iAddThisUsersID AND CompanyID = @selClientORCompany ) = 0 
						BEGIN
							--Also insert into the Company Users Table
							INSERT INTO CompanyUsers (	[CompanyID],
											[UsersID] ,
											[Active] ,
											[ActiveDate],
											[InactiveDate],
											[SecurityLevel] ,
											[Comments] ,
											[AssignmentTypeIDList],
											[DateLastUpdated],
											[UpdateByUserID] 
										)
											--Only Insert COMPANY Or CLIENT COMPANY!
							SELECT 				@selClientORCompany As [CompanyID],
											@iAddThisUsersID As [UsersID] ,
											@Active As [Active] ,
											@ActiveDate As [ActiveDate],
											@InactiveDate As [InactiveDate],
											@SecurityLevel As [SecurityLevel] ,
											@Comments As [Comments] ,
											(CASE	WHEN  @AssignmentTypeIDList Is Null
													THEN '1'
													ELSE @AssignmentTypeIDList
													END
											) As [AssignmentTypeIDList], -- Default to first Assignment Type ID
											@DateLastUpdated As [DateLastUpdated],
											@UID As [UpdateByUserID] 
						END
						ELSE
						BEGIN
							--If Adding user to Companyuser table and they are already there...
							--Update the inactive date to null and active flag to true.
							UPDATE CompanyUsers SET [Active] = 1 ,
										[InactiveDate] = null,
										[SecurityLevel]= @SecurityLevel,
										[Comments] = @Comments,
										[AssignmentTypeIDList] = 	(CASE	WHEN  @AssignmentTypeIDList Is Null
															THEN '1'
															ELSE @AssignmentTypeIDList
															END
														) ,
										[DateLastUpdated] = @DateLastUpdated,
										[UpdateByUserID] = @UID
							FROM Users U 	INNER JOIN CompanyUsers CU On U.UsersID = CU.UsersID
									AND CU.CompanyID = @selClientORCompany
							WHERE 	U.[UsersID] = @UsersID
						END
						
						--IF Adding User to Client Company ...
						--Default to exising profiles Setup On the Main Company
	 					--**NOTE**
						--DON'T Default Profiles When Adding Users From One Company To The Next
						--i.e. Adding a user from Eberl's Company to Wardlaw... 
						--User Profiles need to be set up from scratch when making a Company To Company Move
						--*END NOTE**
						IF @selCompany Is Null and @selClientCompany Is Not Null
						BEGIN
							--GET the Company for this Client Company
							SET @selCompany = (SELECT [IsClientOf] FROM Company WHERE CompanyID = @selClientCompany)
							
							--Need to Loop through each profile and Add them to the User for this Client Company
							---USer Profile (Table Names)
							DECLARE @UPUserProfileID1 int
							DECLARE @UPTableName1 varchar(20)
							DECLARE @UPDescription1 varchar(255)
							DECLARE @UPActive1 bit
							DECLARE @UPSortOrder1 int
							Declare CurUserProfile1 Cursor for
							SELECT 
								[UserProfileID],
								[TableName], 
								[Description] , 
								[Active] , 
								[SortOrder]
							FROM UserProfile UP
							WHERE Active = 1
							ORDER BY SortOrder
							--End User Profile (Table Names)
							Open CurUserProfile1
							Fetch Next From CurUserProfile1 into 	@UPUserProfileID1,
												@UPTableName1, 
												@UPDescription1 , 
												@UPActive1 , 
												@UPSortOrder1
							While @@FEtch_Status = 0
							BEGIN 	
								SET @sSQLUP = ''
								SET @sSQLUP = @sSQLUP + 'DECLARE @iCount Int '
								SET @sSQLUP = @sSQLUP + 'SET @iCount = ('
								SET @sSQLUP = @sSQLUP + 'SELECT Count([UsersID]) '
								SET @sSQLUP = @sSQLUP + 'FROM ' + @UPTableName1 + ' UPT '
								SET @sSQLUP = @sSQLUP + 'WHERE UPT.[UsersID] = ' + cast(@UsersID as varchar(10)) + ' ' 
								SET @sSQLUP = @sSQLUP + 'AND  UPT.[CompanyID] = ' + Cast(@selCompany as varchar(10)) + ' '
								SET @sSQLUP = @sSQLUP + 'AND  UPT.[Active] = 1 '
								SET @sSQLUP = @sSQLUP + ')'
								SET @sSQLUP = @sSQLUP + 'IF @iCount = 1 '
								SET @sSQLUP = @sSQLUP + 'BEGIN '
								SET @sSQLUP = @sSQLUP + '	DECLARE @bHideDeleted int '
								SET @sSQLUP = @sSQLUP + '	DECLARE @UID int '
								SET @sSQLUP = @sSQLUP + '	DECLARE @UsersID int '
								SET @sSQLUP = @sSQLUP + '	DECLARE @selCompany int '
								SET @sSQLUP = @sSQLUP + '	DECLARE @selClientCompany int '
								SET @sSQLUP = @sSQLUP + '	DECLARE @selUserProfile int '
								SET @sSQLUP = @sSQLUP + '	DECLARE @selUserReportsToManager int '
								SET @sSQLUP = @sSQLUP + '	DECLARE @selUserReportsToCoordinator int '
								SET @sSQLUP = @sSQLUP + '	DECLARE @selUserAssgnToCAT int '
								SET @sSQLUP = @sSQLUP + '	DECLARE @OrderBy varchar(500) '
								SET @sSQLUP = @sSQLUP + '	DECLARE @GroupBy varchar(500) '
								SET @sSQLUP = @sSQLUP + '	DECLARE @SearchBy varchar(1000) '
								SET @sSQLUP = @sSQLUP + '	DECLARE @bDelete bit '
								SET @sSQLUP = @sSQLUP + '	DECLARE @bAdd bit '
								SET @sSQLUP = @sSQLUP + '	DECLARE @bEdit bit '
								SET @sSQLUP = @sSQLUP + '	SET @bHideDeleted = ' + cast(@bHideDeleted as varchar(10)) + ' ' 
								SET @sSQLUP = @sSQLUP + '	SET @UID = ' + cast(@UID as varchar(10)) + ' ' 

								SET @sSQLUP = @sSQLUP + '	SET @UsersID = ' + cast(@UsersID as varchar(10)) + ' ' 
								SET @sSQLUP = @sSQLUP + '	SET @selCompany = null '
								SET @sSQLUP = @sSQLUP + '	SET @selClientCompany = ' + cast(@selClientCompany as varchar(10)) + ' ' 
								SET @sSQLUP = @sSQLUP + '	SET @selUserProfile = ' + cast(@UPUserProfileID1 as varchar(10)) + ' ' 
								SET @sSQLUP = @sSQLUP + '	SET @selUserReportsToManager = null '
								SET @sSQLUP = @sSQLUP + '	SET @selUserReportsToCoordinator = null '
								SET @sSQLUP = @sSQLUP + '	SET @selUserAssgnToCAT = null '
								SET @sSQLUP = @sSQLUP + '	SET @OrderBy = null '
								SET @sSQLUP = @sSQLUP + '	SET @GroupBy = null '
								SET @sSQLUP = @sSQLUP + '	SET @SearchBy = null '
								SET @sSQLUP = @sSQLUP + '	SET @bDelete = 0 '
								SET @sSQLUP = @sSQLUP + '	SET @bAdd = 1 '
								SET @sSQLUP = @sSQLUP + '	SET @bEdit = 0 '
								SET @sSQLUP = @sSQLUP + '	EXEC z_spsGetCompanyUsersInfo ' 	
								SET @sSQLUP = @sSQLUP + '	@bHideDeleted, '
								SET @sSQLUP = @sSQLUP + '	@UID , ' 
								SET @sSQLUP = @sSQLUP + '	@UsersID, '
								SET @sSQLUP = @sSQLUP + '	@selCompany, '
								SET @sSQLUP = @sSQLUP + '	@selClientCompany, '
								SET @sSQLUP = @sSQLUP + '	@selUserProfile, '
								SET @sSQLUP = @sSQLUP + '	@selUserReportsToManager, '
								SET @sSQLUP = @sSQLUP + '	@selUserReportsToCoordinator, '
								SET @sSQLUP = @sSQLUP + '	@selUserAssgnToCAT, '
								SET @sSQLUP = @sSQLUP + '	@OrderBy, '
								SET @sSQLUP = @sSQLUP + '	@GroupBy, '
								SET @sSQLUP = @sSQLUP + '	@SearchBy, '
								SET @sSQLUP = @sSQLUP + '	@bDelete, '
								SET @sSQLUP = @sSQLUP + '	@bAdd, '
								SET @sSQLUP = @sSQLUP + '	@bEdit '
								SET @sSQLUP = @sSQLUP + 'END '
								
								--Execute User profile
								Print Len(@sSQLUP) 
								Print @sSQLUP
								EXEC(@sSQLUP)
								
								Fetch Next From CurUserProfile1 into 	@UPUserProfileID1,
													@UPTableName1, 
													@UPDescription1 , 
													@UPActive1 , 
													@UPSortOrder1
							END
							Close CurUserProfile1
							Deallocate CurUserProfile1
						END				
					END
					ELSE
					BEGIN---See Which Child Table Adding to ...
						IF @selUserProfile is Not Null 
						BEGIN--User Profile 
							SET @sUserProfileTable = ''
							SET @sSQLUP =''
							SET @sUserProfileTable = (SELECT [TableName] FROM UserProfile WHERE UserProfileID = @selUserProfile)
							--Build SELECT for Profile Table to see if the User is already there
							SET @sSQLUP = @sSQLUP + 'DECLARE @selClientORCompany Int '
							SET @sSQLUP = @sSQLUP + 'DECLARE @UsersID Int '
							SET @sSQLUP = @sSQLUP + 'DECLARE @DateLastUpdated DateTime '
							SET @sSQLUP = @sSQLUP + 'DECLARE @UID Int '
							SET @sSQLUP = @sSQLUP + 'DECLARE @iCount Int '
							SET @sSQLUP = @sSQLUP + 'SET @selClientORCompany = ' + cast(@selClientORCompany as varchar(10)) + ' '
							SET @sSQLUP = @sSQLUP + 'SET @UsersID = ' + cast(@UsersID as varchar(10)) + ' '
							SET @sSQLUP = @sSQLUP + 'SET @DateLastUpdated = ''' + cast(@DateLastUpdated as varchar(50)) + ''' '
							SET @sSQLUP = @sSQLUP + 'SET @UID = ' + cast(@UID as varchar(10)) + ' '
							SET @sSQLUP = @sSQLUP + 'SET @iCount = ('
							SET @sSQLUP = @sSQLUP + 'SELECT Count([UsersID]) '
							SET @sSQLUP = @sSQLUP + 'FROM ' + @sUserProfileTable + ' UPT '
							SET @sSQLUP = @sSQLUP + 'WHERE UPT.[UsersID] = ' + cast(@UsersID as varchar(10)) + ' ' 
							SET @sSQLUP = @sSQLUP + 'AND  UPT.[CompanyID] = ' + Cast(@selClientORCompany as varchar(10)) + ' '
							SET @sSQLUP = @sSQLUP + ')'
							SET @sSQLUP = @sSQLUP + 'IF @iCount = 0 '
							SET @sSQLUP = @sSQLUP + 'BEGIN /*Since its not in there yet Insert it */ '
							SET @sSQLUP = @sSQLUP + '	INSERT INTO ' + @sUserProfileTable + '( '
							SET @sSQLUP = @sSQLUP + '			[CompanyID], '
							SET @sSQLUP = @sSQLUP + '			[UsersID], '
							SET @sSQLUP = @sSQLUP + '			[Active], '
							SET @sSQLUP = @sSQLUP + '			[DateLastUpdated], '
							SET @sSQLUP = @sSQLUP + '			[UpdateByUserID] '
							SET @sSQLUP = @sSQLUP + '	) '
							SET @sSQLUP = @sSQLUP + '	SELECT 	@selClientORCompany As [CompanyID], '
							SET @sSQLUP = @sSQLUP + '			@UsersID As [UsersID], '
							SET @sSQLUP = @sSQLUP + '			1 As [Active], '
							SET @sSQLUP = @sSQLUP + '			@DateLastUpdated As [DateLastUpdated], '
							SET @sSQLUP = @sSQLUP + '			@UID As [UpdateByUserID] '
							SET @sSQLUP = @sSQLUP + 'END '
							SET @sSQLUP = @sSQLUP + 'ELSE '
							SET @sSQLUP = @sSQLUP + 'BEGIN /*Its already there make sure it is flagged as Active*/ '
							SET @sSQLUP = @sSQLUP + '	UPDATE ' + @sUserProfileTable + ' 	SET [Active] = 1, '
							SET @sSQLUP = @sSQLUP + '						[DateLastUpdated] =''' + cast(@DateLastUpdated as varchar(50)) + ''' , '
							SET @sSQLUP = @sSQLUP + '						[UpdateByUserID] = ' + cast(@UID as varchar(10)) + ' '
							SET @sSQLUP = @sSQLUP + '	FROM CompanyUsers CU INNER JOIN ' + @sUserProfileTable + ' UPT On CU.UsersID = UPT.UsersID '
							SET @sSQLUP = @sSQLUP + '		AND CU.CompanyID = UPT.CompanyID '
							SET @sSQLUP = @sSQLUP + '	WHERE 	CU.[UsersID] = ' + cast(@UsersID as varchar(10)) + ' '
							SET @sSQLUP = @sSQLUP + '	AND  CU.[CompanyID] = ' + Cast(@selClientORCompany as varchar(10)) + ' '
							SET @sSQLUP = @sSQLUP + '	AND UPT.[UsersID] = ' + cast(@UsersID as varchar(10)) + ' ' 
							SET @sSQLUP = @sSQLUP + '	AND  UPT.[CompanyID] = ' + Cast(@selClientORCompany as varchar(10)) + ' '
							SET @sSQLUP = @sSQLUP + 'END '
	
							--Execute Add to User Profile
							Print Len(@sSQLUP) 
							Print @sSQLUP
							EXEC(@sSQLUP)
						END
		
						IF @selUserReportsToManager is Not Null 
						BEGIN--User Reports to Manager 
							SET @iCount = (	SELECT Count([UsersID]) 
									FROM UserReportsToManager 
									WHERE [UsersID] = @UsersID
									AND  [CompanyID] = @selClientORCompany
									)
							IF @iCount = 0 
							BEGIN--Since its not in there yet Insert it
								INSERT INTO UserReportsToManager (
													[UsersID],
													[CompanyID],
													[ReportsToUsersID],
													[Active] ,
													[DateLastUpdated],
													[UpdateByUserID] 
												)
								SELECT					@UsersID As [UsersID], 
													@selClientORCompany As [CompanyID], 
													@selUserReportsToManager AS [ReportsToUsersID], 
													1 As [Active] ,
													@DateLastUpdated As [DateLastUpdated], 
													@UID As [UpdateByUserID]
							END
							ELSE	
							BEGIN--Its already there make sure it is flagged as Active
								SET @sSQLUP =''
								SET @sSQLUP = @sSQLUP + 'UPDATE UserReportsToManager SET [Active] = 1, '
								SET @sSQLUP = @sSQLUP + '[ReportsToUsersID] = ' + cast(@selUserReportsToManager as varchar(10)) + ', '
								SET @sSQLUP = @sSQLUP + '[DateLastUpdated] =''' + cast(@DateLastUpdated as varchar(50)) + ''' , '

								SET @sSQLUP = @sSQLUP + '[UpdateByUserID] = ' + cast(@UID as varchar(10)) + ' '
								SET @sSQLUP = @sSQLUP + 'FROM CompanyUsers CU INNER JOIN UserReportsToManager URTM On CU.UsersID = URTM.UsersID '
								SET @sSQLUP = @sSQLUP + 'AND CU.CompanyID = URTM.CompanyID '
								SET @sSQLUP = @sSQLUP +	'WHERE CU.[UsersID] = ' + cast(@UsersID as varchar(10)) + ' '
								SET @sSQLUP = @sSQLUP + 'AND  CU.[CompanyID] = ' + Cast(@selClientORCompany as varchar(10)) + ' '
								SET @sSQLUP = @sSQLUP + 'AND URTM.[UsersID] = ' + cast(@UsersID as varchar(10)) + ' ' 
								SET @sSQLUP = @sSQLUP + 'AND  URTM.[CompanyID] = ' + Cast(@selClientORCompany as varchar(10)) + ' '
								Print Len(@sSQLUP) 
								Print @sSQLUP
								EXEC(@sSQLUP)
							END
						END
						
						IF @selUserReportsToCoordinator is Not Null 
						BEGIN--User Reports to Coordinator
							SET @iCount = (	SELECT Count([UsersID]) 
									FROM UserReportsToCoordinator
									WHERE [UsersID] = @UsersID
									AND  [CompanyID] = @selClientORCompany
									)
							IF @iCount = 0 
							BEGIN--Since its not in there yet Insert it
								INSERT INTO UserReportsToCoordinator (
													[UsersID],
													[CompanyID],
													[ReportsToUsersID],
													[Active] ,
													[DateLastUpdated],
													[UpdateByUserID] 
												)
								SELECT					@UsersID As [UsersID], 
													@selClientORCompany As [CompanyID], 
													@selUserReportsToCoordinator AS [ReportsToUsersID], 
													1 As [Active] ,
													@DateLastUpdated As [DateLastUpdated], 
													@UID As [UpdateByUserID]
							END
							ELSE	
							BEGIN--Its already there make sure it is flagged as Active
								SET @sSQLUP =''
								SET @sSQLUP = @sSQLUP + 'UPDATE UserReportsToCoordinator SET [Active] = 1, '
								SET @sSQLUP = @sSQLUP + '[ReportsToUsersID] = ' + cast(@selUserReportsToCoordinator as varchar(10)) + ', '
								SET @sSQLUP = @sSQLUP + '[DateLastUpdated] =''' + cast(@DateLastUpdated as varchar(50)) + ''' , '
								SET @sSQLUP = @sSQLUP + '[UpdateByUserID] = ' + cast(@UID as varchar(10)) + ' '
								SET @sSQLUP = @sSQLUP + 'FROM CompanyUsers CU INNER JOIN UserReportsToCoordinator URTC On CU.UsersID = URTC.UsersID '
								SET @sSQLUP = @sSQLUP + 'AND CU.CompanyID = URTC.CompanyID '
								SET @sSQLUP = @sSQLUP +	'WHERE CU.[UsersID] = ' + cast(@UsersID as varchar(10)) + ' '
								SET @sSQLUP = @sSQLUP + 'AND  CU.[CompanyID] = ' + Cast(@selClientORCompany as varchar(10)) + ' '
								SET @sSQLUP = @sSQLUP + 'AND URTC.[UsersID] = ' + cast(@UsersID as varchar(10)) + ' ' 
								SET @sSQLUP = @sSQLUP + 'AND  URTC.[CompanyID] = ' + Cast(@selClientORCompany as varchar(10)) + ' '
								Print Len(@sSQLUP) 
								Print @sSQLUP
								EXEC(@sSQLUP)
							END
						END
						
						IF @selUserAssgnToCAT is Not Null 
						BEGIN--User Assgn To CAT
							SET @iCount = (	SELECT Count([UsersID]) 
									FROM ClientCompanyUsersCat
									WHERE [UsersID] = @UsersID
									AND  [ClientCompanyID] = @selClientORCompany
									AND [CATID] = @selUserAssgnToCAT
									)
							IF @iCount = 0 
							BEGIN--Since its not in there yet Insert it
								INSERT INTO ClientCompanyUsersCat (
													[ClientCompanyID] ,
													[CATID] ,
													[UsersID],
													[Active] ,
													[DateLastUpdated],
													[UpdateByUserID] 
												)
								SELECT					@selClientORCompany As [ClientCompanyID], 
													@selUserAssgnToCAT As [CATID] ,
													@UsersID As [UsersID], 
													1 As [Active] ,
													@DateLastUpdated As [DateLastUpdated], 
													@UID As [UpdateByUserID]
							END
							ELSE	
							BEGIN--Its already there make sure it is flagged as Active
								SET @sSQLUP =''
								SET @sSQLUP = @sSQLUP + 'UPDATE ClientCompanyUsersCat SET [Active] = 1, '
								SET @sSQLUP = @sSQLUP + '[DateLastUpdated] =''' + cast(@DateLastUpdated as varchar(50)) + ''' , '
								SET @sSQLUP = @sSQLUP + '[UpdateByUserID] = ' + cast(@UID as varchar(10)) + ' '
								SET @sSQLUP = @sSQLUP + 'FROM CompanyUsers CU INNER JOIN ClientCompanyUsersCat CCUC On CU.UsersID = CCUC.UsersID '
								SET @sSQLUP = @sSQLUP + 'AND CU.CompanyID = CCUC.ClientCompanyID '
								SET @sSQLUP = @sSQLUP + 'WHERE 	CCUC.CATID = ' + cast(@selUserAssgnToCAT as varchar(10)) + ' '
								SET @sSQLUP = @sSQLUP +	'AND CU.[UsersID] = ' + cast(@UsersID as varchar(10)) + ' '
								SET @sSQLUP = @sSQLUP + 'AND  CU.[CompanyID] = ' + Cast(@selClientORCompany as varchar(10)) + ' '
								SET @sSQLUP = @sSQLUP + 'AND CCUC.[UsersID] = ' + cast(@UsersID as varchar(10)) + ' ' 
								SET @sSQLUP = @sSQLUP + 'AND  CCUC.[ClientCompanyID] = ' + Cast(@selClientORCompany as varchar(10)) + ' '
								Print Len(@sSQLUP) 
								Print @sSQLUP
								EXEC(@sSQLUP)
							END
						END 
					END 
				END
				ELSE 
				BEGIN
	/****************************************
	*		***EDIT***
	****************************************/
					IF 	@bEdit=1 
						And @AUSLicenseDaysLeft Is Null 
						And @AUSIBPrefix Is Null 
						And @AUSSingleFileSendAuthority Is Null 
						And @bAdjusterUsersUpdatesYesNo Is Null
					BEGIN
						Update Users Set 	[UserName] = @UserName,
									[PassWord] = @PassWord ,
									[FirstName] = @FirstName ,
									[LastName] = @LastName ,
									[SSN] = @SSN ,
									[Email] = @Email,
									[ContactPhone] = @ContactPhone,
									[EmergencyPhone] = @EmergencyPhone,
									[Address] = @Address ,
									[City] = @City ,
									[State] = @State,
									[Zip] = @Zip ,
									[ZIP4] = @ZIP4 ,
									[OtherPostCode] = @OtherPostCode,
									[Active] = @UsersActive ,
									[ActiveDate] = @UsersActiveDate,
									[InactiveDate] = @UsersInactiveDate,
									[SecurityLevel] = @UsersSecurityLevel ,
									[Comments] = @UsersComments,
									[DateLastUpdated] = @UsersDateLastUpdated,
									[UpdateByUserID] = @UID
						WHERE 	[UsersID] = @UsersID
						--Also Edit the Company Users Table
						UPDATE CompanyUsers SET [Active] = @Active ,
									[ActiveDate] = @ActiveDate,
									[InactiveDate] = @InactiveDate,
									[SecurityLevel] = @SecurityLevel ,
									[Comments] = @Comments,
									--[Flag] Not updated Here!
									[AssignmentTypeIDList] =(CASE	WHEN  @AssignmentTypeIDList Is Null
													THEN [AssignmentTypeIDList]
													ELSE @AssignmentTypeIDList
													END
												) ,
									[DateLastUpdated] = @DateLastUpdated,
									[UpdateByUserID] = @UID
						FROM Users U 	INNER JOIN CompanyUsers CU On U.UsersID = CU.UsersID
								AND CU.CompanyID = @selClientORCompany
						WHERE 	U.[UsersID] = @UsersID
					END
					ELSE 
					BEGIN --Edit Adjuster User Software Or Apply Adjuster User Updates
						Update Users Set 	[Comments] = @UsersComments,
									[DateLastUpdated] = @UsersDateLastUpdated,
									[UpdateByUserID] = @UID
						WHERE 	[UsersID] = @UsersID
						--@AUSLicenseDaysLeft
						IF @AUSLicenseDaysLeft Is Not Null 
						BEGIN
							UPDATE AdjusterUsersSoftware SET 	[LicenseDaysLeft] = @AUSLicenseDaysLeft,
												[ResetLicense] = 1,
												[DateLastUpdated] = @UsersDateLastUpdated,
												[UpdateByUserID] = @UID
							FROM AdjusterUsersSoftware
							WHERE AdjusterUsersSoftware.UsersID = @UsersID
						END
						--@AUSIBPrefix
						IF @AUSIBPrefix Is Not Null
						BEGIN
							UPDATE AdjusterUsersSoftware SET 	[IBPrefix] = @AUSIBPrefix,
												[ResetIBPrefix] = 1,
												[DateLastUpdated] = @UsersDateLastUpdated,
												[UpdateByUserID] = @UID
							FROM AdjusterUsersSoftware
							WHERE AdjusterUsersSoftware.UsersID = @UsersID
						END
						--@AUSSingleFileSendAuthority
						IF @AUSSingleFileSendAuthority Is Not Null
						BEGIN
							UPDATE AdjusterUsersSoftware SET 	[SingleFileSendAuthority] = @AUSSingleFileSendAuthority,
												[DateLastUpdated] = @UsersDateLastUpdated,
												[UpdateByUserID] = @UID
							FROM AdjusterUsersSoftware
							WHERE AdjusterUsersSoftware.UsersID = @UsersID
						END
						--@bAdjusterUsersUpdatesYesNo
						IF  @bAdjusterUsersUpdatesYesNo = 1
						BEGIN
							--Do not Allow Adjuster Updates IF...
							-- SSN that would cause a Duplicate SSN
							--Integrity of other settings are not Correct.
							Update Users Set 	
										[FirstName] = (CASE	WHEN  RTRIM(AUU.FirstName) = '' Or AUU.FirstName Is Null
													THEN USERS.[FirstName]
													ELSE UPPER(AUU.FirstName)	
													END
												) ,
										[LastName] = (CASE	WHEN  RTRIM(AUU.LastName) = '' Or AUU.LastName Is Null
													THEN USERS.[LastName]
													ELSE UPPER(AUU.LastName)	
													END
												) ,
										[SSN] = (CASE 		WHEN (SELECT COUNT(SSN) FROM USERS WHERE SSN = AUU.SSN And SSN Is Not Null)>0 OR AUU.SSN Is Null
													THEN USERS.[SSN]
													ELSE AUU.SSN
													END
											),
										[Email] = (CASE	WHEN  RTRIM(AUU.Email) = '' Or AUU.Email Is Null
													THEN USERS.[Email]
													ELSE AUU.Email	
													END
											) ,
										[ContactPhone] = (CASE	WHEN  RTRIM(AUU.ContactPhone) = '' Or AUU.ContactPhone Is Null
													THEN USERS.[ContactPhone]
													ELSE AUU.ContactPhone	
													END
												) ,
										[EmergencyPhone] =(CASE	WHEN  RTRIM(AUU.EmergencyPhone) = '' Or AUU.EmergencyPhone Is Null
													THEN USERS.[EmergencyPhone]
													ELSE AUU.EmergencyPhone	
													END
												) ,
										[Address] = (CASE	WHEN  RTRIM(AUU.Address) = '' Or AUU.Address Is Null
													THEN USERS.[Address]
													ELSE UPPER(AUU.Address)	
													END
												)  ,
										[City] = (CASE	WHEN  RTRIM(AUU.City) = '' Or AUU.City Is Null
													THEN USERS.[City]
													ELSE UPPER(AUU.City)	
													END
												) ,
										[State] = (CASE	WHEN  RTRIM(AUU.State) = '' Or AUU.State Is Null
													THEN USERS.[State]
													ELSE UPPER(AUU.State)	
													END
												),
										[Zip] = (CASE	WHEN AUU.Zip Is Null
												THEN USERS.[Zip]
												ELSE AUU.Zip
												END
												) ,
										[ZIP4] = (CASE	WHEN AUU.Zip4 Is Null
												THEN USERS.[Zip4]
												ELSE AUU.Zip4
												END
												) ,
										[OtherPostCode] = (CASE	WHEN  RTRIM(AUU.OtherPostCode) = '' Or AUU.OtherPostCode Is Null
													THEN USERS.[OtherPostCode]
													ELSE AUU.OtherPostCode	
													END
												),
										[DateLastUpdated] = @UsersDateLastUpdated,
										[UpdateByUserID] = @UID
							FROM USERS INNER JOIN AdjusterUsersUpdates AUU ON USERS.UsersID = AUU.UsersID
							WHERE 	USERS.[UsersID] = @UsersID
						END
					END
				END
			END
		END
	END
END
ELSE
BEGIN
	IF @GroupBy Is Not Null
	BEGIN
		SET @sSQL1 = 'SELECT RetUsers.' + @GroupBy + ' '
	END
	ELSE
	BEGIN
		SET @sSQL1 = 'SELECT RetUsers.* '
	END	
	SET @sSQL1 = @sSQL1 + 'FROM( ' 
	SET @sSQL1 = @sSQL1 + 'SELECT '
	SET @sSQL1 = @sSQL1 + 'U.[UsersID], '
	SET @sSQL1 = @sSQL1 + 'U.[UserName], '
	SET @sSQL1 = @sSQL1 + 'U.[PassWord], '
	SET @sSQL1 = @sSQL1 + 'U.[FirstName], '
	SET @sSQL1 = @sSQL1 + 'U.[LastName], '
	SET @sSQL1 = @sSQL1 + 'U.[LastName] + '', '' + U.[FirstName] + REPLICATE(''.'', ' + convert(varchar(10),@lFillNameCount) + '  - LEN(U.[LastName] + '', '' + U.[FirstName] + ''('' + U.[UserName] + '')'')) + ''('' + U.[UserName] + '')'' As FillName, '
	SET @sSQL1 = @sSQL1 + 'U.[SSN], '
	SET @sSQL1 = @sSQL1 + 'U.[Email], '
	SET @sSQL1 = @sSQL1 + 'U.[ContactPhone], '
	SET @sSQL1 = @sSQL1 + 'U.[EmergencyPhone], '
	SET @sSQL1 = @sSQL1 + 'U.[Address], '
	SET @sSQL1 = @sSQL1 + 'U.[City], '
	SET @sSQL1 = @sSQL1 + 'U.[State], '
	SET @sSQL1 = @sSQL1 + 'U.[Zip], '
	SET @sSQL1 = @sSQL1 + 'U.[ZIP4], '
	SET @sSQL1 = @sSQL1 + 'U.[OtherPostCode], '
	SET @sSQL1 = @sSQL1 + 'U.[Active] As UsersActive, '
	SET @sSQL1 = @sSQL1 + 'CU.[Active], '
	SET @sSQL1 = @sSQL1 + 'U.[ActiveDate] As UsersActiveDate, '
	SET @sSQL1 = @sSQL1 + 'CU.[ActiveDate], '
	SET @sSQL1 = @sSQL1 + 'U.[InactiveDate] As UsersInActiveDate, '
	SET @sSQL1 = @sSQL1 + 'CU.[InactiveDate], '
	SET @sSQL1 = @sSQL1 + '(SELECT	Name '
	SET @sSQL1 = @sSQL1 + 'FROM 	SecurityLevel '
	SET @sSQL1 = @sSQL1 + 'WHERE 	SecurityLevel.[SecurityLevel] = U.[SecurityLevel]) As [UsersSecLevelName],  '
	SET @sSQL1 = @sSQL1 + '(SELECT	Description '
	SET @sSQL1 = @sSQL1 + 'FROM 	SecurityLevel '
	SET @sSQL1 = @sSQL1 + 'WHERE 	SecurityLevel.[SecurityLevel] = U.[SecurityLevel]) As [UsersSecLevelDescription],  '
	SET @sSQL1 = @sSQL1 + 'U.[SecurityLevel] As UsersSecurityLevel, '
	SET @sSQL1 = @sSQL1 + '(SELECT	Name '
	SET @sSQL1 = @sSQL1 + 'FROM 	SecurityLevel '
	SET @sSQL1 = @sSQL1 + 'WHERE 	SecurityLevel.[SecurityLevel] = CU.[SecurityLevel]) As [SecLevelName],  '
	SET @sSQL1 = @sSQL1 + '(SELECT	Description '
	SET @sSQL1 = @sSQL1 + 'FROM 	SecurityLevel '
	SET @sSQL1 = @sSQL1 + 'WHERE 	SecurityLevel.[SecurityLevel] = CU.[SecurityLevel]) As [SecLevelDescription],  '
	SET @sSQL1 = @sSQL1 + 'CU.[SecurityLevel], '
	SET @sSQL1 = @sSQL1 + 'U.[Comments] As UsersComments, '
	SET @sSQL1 = @sSQL1 + 'CU.[Comments], '
	SET @sSQL1 = @sSQL1 + 'CU.[AssignmentTypeIDList], '
	SET @sSQL1 = @sSQL1 + '(SELECT	MaxOpenAssgn '
	SET @sSQL1 = @sSQL1 + 'FROM 	Adjuster '
	SET @sSQL1 = @sSQL1 + 'WHERE 	Adjuster.[CompanyID] = CU.[CompanyID] And Adjuster.[UsersID] = CU.[USersID]) As [AdjMaxOpenAssgn],  '
	--Initi the Temp sql
	SET @sSQLTemp = ''
	SET @sSQLTemp = @sSQLTemp + '(SELECT	HomeBaseZip '
	SET @sSQLTemp = @sSQLTemp + 'FROM 	Adjuster '
	SET @sSQLTemp = @sSQLTemp + 'WHERE 	Adjuster.[CompanyID] = CU.[CompanyID] And Adjuster.[UsersID] = CU.[USersID]) '
	SET @sSQL1 = @sSQL1 + @sSQLTemp + 'As [AdjHomeBaseZip], '
	-- Init Temp2 sql
	SET @sSQLTemp2 = ''
	SET @sSQLTemp2 = @sSQLTemp2  + 'SELECT 	TOP 1 ClientCompanyCatSpecID '
	SET @sSQLTemp2 = @sSQLTemp2  + 'FROM 	ClientCoAdjusterSpec '
	SET @sSQLTemp2 = @sSQLTemp2  + 'WHERE 	UsersID = (SELECT UsersID FROM 	Users WHERE UserName = ''UUNASSIGNED'') '
	SET @sSQLTemp2 = @sSQLTemp2  + 'ORDER BY ClientCompanyCatSpecID Desc '
	SET @sSQL1 = @sSQL1 + '(SELECT	State '
	SET @sSQL1 = @sSQL1 + 'FROM 	ZIPS '
	SET @sSQL1 = @sSQL1 + 'WHERE 	Zip = Cast(' + @sSQLTemp + ' As Varchar(53)) '
	--SET @sSQL1 = @sSQL1 + 'AND A.ClientCompanyCatSpecID IN (' + @sSQLTemp2 + ') '
	SET @sSQL1 = @sSQL1 + ') As [AdjHomeBaseState],  '
	SET @sSQL1 = @sSQL1 + '(SELECT	City '
	SET @sSQL1 = @sSQL1 + 'FROM 	ZIPS '
	SET @sSQL1 = @sSQL1 + 'WHERE 	Zip = Cast(' + @sSQLTemp + ' As Varchar(53)) '
	--SET @sSQL1 = @sSQL1 + 'AND A.ClientCompanyCatSpecID IN (' + @sSQLTemp2 + ') '
	SET @sSQL1 = @sSQL1 + ') As [AdjHomeBaseCity],  '
	SET @sSQL1 = @sSQL1 + '(SELECT	MaxRangeFromHomeBaseZip '
	SET @sSQL1 = @sSQL1 + 'FROM 	Adjuster '
	SET @sSQL1 = @sSQL1 + 'WHERE 	Adjuster.[CompanyID] = CU.[CompanyID] And Adjuster.[UsersID] = CU.[USersID]) As [AdjMaxRangeFromHomeBaseZip],  '
	IF @selCompany Is Not Null And @bViewHistory Is Null
	BEGIN
	--Adjuster Users Software
		SET @sSQL1 = @sSQL1 + 'Cast(AUS.[VersionInfo] as Varchar(8000)) As [AUSVersionInfo], '
		SET @sSQL1 = @sSQL1 + 'AUS.[LicenseDaysLeft] As [AUSLicenseDaysLeft], '
		SET @sSQL1 = @sSQL1 + 'AUS.[ResetLicense] As [AUSResetLicense], '
		SET @sSQL1 = @sSQL1 + 'AUS.[IBPrefix] As [AUSIBPrefix], '
		SET @sSQL1 = @sSQL1 + 'AUS.[ResetIBPrefix] As [AUSResetIBPrefix], '
		SET @sSQL1 = @sSQL1 + 'AUS.[SingleFileSendAuthority] As [AUSSingleFileSendAuthority], '
		SET @sSQL1 = @sSQL1 + 'AUS.[DateLastUpdated] As [AUSDateLastUpdated], '
		SET @sSQL1 = @sSQL1 + '(SELECT	USERNAME '
		SET @sSQL1 = @sSQL1 + 'FROM 	USERS '
		SET @sSQL1 = @sSQL1 + 'WHERE 	USERSID = AUS.[UpdateByUserID]) As [AUSUpdateByUserName],  '
		SET @sSQL1 = @sSQL1 + 'AUS.[UpdateByUserID] As AUSUpdateByUserID, '
	--Adjuster Users Updates
		SET @sSQL1 = @sSQL1 + 'AUU.[FirstName] As [AUUFirstName], '
		SET @sSQL1 = @sSQL1 + 'AUU.[LastName] As [AUULastName], '
		SET @sSQL1 = @sSQL1 + 'AUU.[SSN] As [AUUSSN], '
		SET @sSQL1 = @sSQL1 + 'AUU.[Email] As [AUUEmail], '
		SET @sSQL1 = @sSQL1 + 'AUU.[ContactPhone] As [AUUContactPhone], '
		SET @sSQL1 = @sSQL1 + 'AUU.[EmergencyPhone] As [AUUEmergencyPhone], '
		SET @sSQL1 = @sSQL1 + 'AUU.[Address] As [AUUAddress], '
		SET @sSQL1 = @sSQL1 + 'AUU.[City] As [AUUCity], '
		SET @sSQL1 = @sSQL1 + 'AUU.[State] As [AUUState], '
		SET @sSQL1 = @sSQL1 + 'AUU.[Zip] As [AUUZip], '
		SET @sSQL1 = @sSQL1 + 'AUU.[ZIP4] As [AUUZIP4], '
		SET @sSQL1 = @sSQL1 + 'AUU.[OtherPostCode] As [AUUOtherPostCode], '
		SET @sSQL1 = @sSQL1 + 'AUU.[DateLastUpdated] As [AUUDateLastUpdated], '
		SET @sSQL1 = @sSQL1 + '(SELECT	USERNAME '
		SET @sSQL1 = @sSQL1 + 'FROM 	USERS '
		SET @sSQL1 = @sSQL1 + 'WHERE 	USERSID = AUU.[UpdateByUserID]) As [AUUUpdateByUserName],  '
		SET @sSQL1 = @sSQL1 + 'AUU.[UpdateByUserID] As AUUUpdateByUserID, '
	END
--*******************************************************************************
--*			Begin ClientCompany Association List			*
--*										*
--VVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV
	SET @sSQL2 = ''
	IF @selCompany Is Not Null And @bViewHistory Is Null
	BEGIN
		DECLARE @ClientCoCompanyID int
		DECLARE @ClientCoName Varchar(30)
		DECLARE @ClientCoCode Varchar(3)
		DECLARE @ClientCoComments Varchar(150)
		DECLARE @ClientCoIsClientOf int
		DECLARE @ClientCoActive int
		Declare CurClientCompanies Cursor for
		SELECT 
			[CompanyID],
			[Name], 
			[Code] , 
			[Comments] , 
			[IsClientOf],
			[Active]
		FROM Company C
		WHERE IsClientOF = @selCompany
		Open CurClientCompanies
		Fetch Next From CurClientCompanies into @ClientCoCompanyID,
							@ClientCoName, 
							@ClientCoCode , 
							@ClientCoComments , 
							@ClientCoIsClientOf,
							@ClientCoActive
		While @@FEtch_Status = 0
		BEGIN
			SET @sSQL2 = @sSQL2 + '(SELECT Count(CompanyUsers.UsersID) as CountOfUserID '	
			SET @sSQL2 = @sSQL2 + 'FROM CompanyUsers '
			SET @sSQL2 = @sSQL2 + 'WHERE CompanyUsers.CompanyID = ' + Cast(@ClientCoCompanyID as Varchar(10)) + ' '
			SET @sSQL2 = @sSQL2 + 'AND CompanyUsers.UsersID = U.[UsersID] ' 
			SET @sSQL2 = @sSQL2 + 'AND Active = 1 ' 
			SET @sSQL2 = @sSQL2 + 'GROUP BY CompanyID '
			SET @sSQL2 = @sSQL2 + ') As [ClientCompany' + Cast(@ClientCoCode As Varchar(30)) + '], '
			Fetch Next From CurClientCompanies into @ClientCoCompanyID,
								@ClientCoName, 
								@ClientCoCode , 
								@ClientCoComments , 
								@ClientCoIsClientOf,
								@ClientCoActive
		END
		Close CurClientCompanies
		Deallocate CurClientCompanies
		
	END
--^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
--*			End ClientCompany Association List			*
--*										*
--*******************************************************************************
--*******************************************************************************
--*			Begin Users Profile Association List			*
--*										*
--VVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV
	SET @sSQL3 = ''
	IF @selCompany Is Not Null And @bViewHistory Is Null
	BEGIN
		---USer Profile (Table Names)
		DECLARE @UPUserProfileID int
		DECLARE @UPTableName varchar(20)
		DECLARE @UPDescription varchar(255)
		DECLARE @UPActive bit
		DECLARE @UPSortOrder int
		
		IF @bHideDeleted =1
		BEGIN
			Declare CurUserProfile Cursor for
			SELECT 
				[UserProfileID],
				[TableName], 
				[Description] , 
				[Active] , 
				[SortOrder]
			FROM UserProfile UP
			WHERE Active = 1
			ORDER BY SortOrder
		END
		ELSE
		BEGIN
			Declare CurUserProfile Cursor for
			SELECT 
				[UserProfileID],
				[TableName], 
				[Description] , 
				[Active] , 
				[SortOrder]
			FROM UserProfile UP
			WHERE Active IN(1,0)
			ORDER BY SortOrder
		END
		
		--End User Profile (Table Names)
		Open CurUserProfile
		Fetch Next From CurUserProfile into 	@UPUserProfileID,
							@UPTableName, 
							@UPDescription , 
							@UPActive , 
							@UPSortOrder
		While @@FEtch_Status = 0
		BEGIN
			SET @sSQL3 = @sSQL3 + '(SELECT Count(' + Cast( @UPTableName as Varchar(20)) + '.UsersID) as CountOfUserID '	
			SET @sSQL3 = @sSQL3 + 'FROM ' + Cast( @UPTableName as Varchar(20)) + ' '
			SET @sSQL3 = @sSQL3 + 'WHERE ' + Cast( @UPTableName as Varchar(20)) + '.CompanyID = CU.CompanyID '
			SET @sSQL3 = @sSQL3 + 'AND ' + Cast( @UPTableName as Varchar(20)) + '.UsersID = CU.[UsersID] ' 
			SET @sSQL3 = @sSQL3 + 'AND ' + Cast( @UPTableName as Varchar(20)) + '.Active = 1 ' 
			SET @sSQL3 = @sSQL3 + 'GROUP BY ' + Cast( @UPTableName as Varchar(20)) + '.CompanyID '
			SET @sSQL3 = @sSQL3 + ') As [UserProfile' + Cast( @UPTableName as Varchar(20)) + '], '
			Fetch Next From CurUserProfile into 	@UPUserProfileID,
								@UPTableName, 
								@UPDescription , 
								@UPActive , 
								@UPSortOrder
		END
		Close CurUserProfile
		Deallocate CurUserProfile
		
	END
--^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
--*			End Users Profile Association List			*
--*										*
--*******************************************************************************

--*******************************************************************************
--*		Begin User Reports to Manager Association List			*
--*										*
--VVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV
	SET @sSQL4 = ''
	IF @selCompany Is Not Null And @selClientCompany Is Not Null And @selUserProfile Is not Null And @bViewHistory Is Null
	BEGIN
		DECLARE @URTMUsersID int
		DECLARE @URTMUserName varchar(20)
		DECLARE @URTMLastName varchar(20)
		DECLARE @URTMFirstName varchar(20)
		
		IF @bHideDeleted =1
		BEGIN
			Declare CurUserReportsToManager Cursor for
			SELECT 
				URTM.[UsersID],
				URTM.[UserName], 
				URTM.[LastName] , 
				URTM.[FirstName] 
			FROM Users URTM INNER JOIN CompanyUsers CURTM ON URTM.UsersID = CURTM.USersID And CURTM.CompanyID = @selClientCompany
					INNER JOIN Manager M ON CURTM.UsersID = M.UsersID And M.CompanyID = @selClientCompany
			WHERE 	URTM.Active = 1
			AND 	CURTM.Active =1
			AND 	M.Active =1
			ORDER BY URTM.[LastName]
		END
		ELSE
		BEGIN
			Declare CurUserReportsToManager Cursor for
			SELECT 
				URTM.[UsersID],
				URTM.[UserName], 
				URTM.[LastName] , 
				URTM.[FirstName] 
			FROM Users URTM INNER JOIN CompanyUsers CURTM ON URTM.UsersID = CURTM.USersID And CURTM.CompanyID = @selClientCompany
					INNER JOIN Manager M ON CURTM.UsersID = M.UsersID And M.CompanyID = @selClientCompany
			WHERE 	URTM.Active = 1
			AND 	CURTM.Active IN(1,0)
			AND 	M.Active =1
			ORDER BY URTM.[LastName]
		END
		
		Open CurUserReportsToManager
		Fetch Next From CurUserReportsToManager into 	@URTMUsersID,
								@URTMUserName, 
								@URTMLastName , 
								@URTMFirstName 
		While @@FEtch_Status = 0
		BEGIN
			SET @sSQL4 = @sSQL4 + '(SELECT Count(URTM.UsersID) as CountOfUserID '	
			SET @sSQL4 = @sSQL4 + 'FROM UserReportsToManager URTM '
			SET @sSQL4 = @sSQL4 + 'WHERE URTM.CompanyID = CU.[CompanyID] '
			SET @sSQL4 = @sSQL4 + 'AND URTM.UsersID = CU.[UsersID] '
			SET @sSQL4 = @sSQL4 + 'AND URTM.ReportsToUsersID = ' + Cast( @URTMUsersID as Varchar(10)) + ' '
			SET @sSQL4 = @sSQL4 + 'AND URTM.Active = 1 ' 
			SET @sSQL4 = @sSQL4 + 'GROUP BY URTM.CompanyID '
			SET @sSQL4 = @sSQL4 + ') As [UserReportsToManager' + Cast( @URTMUserName as Varchar(30)) + '], '
			Fetch Next From CurUserReportsToManager into 	@URTMUsersID,
									@URTMUserName, 
									@URTMLastName, 
									@URTMFirstName
		END
		Close CurUserReportsToManager
		Deallocate CurUserReportsToManager
		
	END
--^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
--*		End User Reports to Manager Association List			*
--*										*
--*******************************************************************************
--*******************************************************************************
--*		Begin User Reports to Coordinator Association List			*
--*										*
--VVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV
	SET @sSQL5 = ''
	IF @selCompany Is Not Null And @selClientCompany Is Not Null And @selUserProfile Is not Null And @bViewHistory Is Null
	BEGIN
		DECLARE @URTCUsersID int
		DECLARE @URTCUserName varchar(20)
		DECLARE @URTCLastName varchar(20)
		DECLARE @URTCFirstName varchar(20)
		
		IF @bHideDeleted =1
		BEGIN
			Declare CurUserReportsToCoordinator Cursor for
			SELECT 
				URTC.[UsersID],
				URTC.[UserName], 
				URTC.[LastName] , 
				URTC.[FirstName] 
			FROM Users URTC INNER JOIN CompanyUsers CURTC ON URTC.UsersID = CURTC.USersID And CURTC.CompanyID = @selClientCompany
					INNER JOIN Coordinator C ON CURTC.UsersID = C.UsersID And C.CompanyID = @selClientCompany
			WHERE 	URTC.Active = 1
			AND 	CURTC.Active = 1
			AND 	C.Active =1
			ORDER BY URTC.[LastName]
		END
		ELSE
		BEGIN
			Declare CurUserReportsToCoordinator Cursor for
			SELECT 
				URTC.[UsersID],
				URTC.[UserName], 
				URTC.[LastName] , 
				URTC.[FirstName] 
			FROM Users URTC INNER JOIN CompanyUsers CURTC ON URTC.UsersID = CURTC.USersID And CURTC.CompanyID = @selClientCompany
					INNER JOIN Coordinator C ON CURTC.UsersID = C.UsersID And C.CompanyID = @selClientCompany
			WHERE 	URTC.Active = 1
			AND 	CURTC.Active IN(1,0)
			AND 	C.Active =1
			ORDER BY URTC.[LastName]
		END
		
		Open CurUserReportsToCoordinator
		Fetch Next From CurUserReportsToCoordinator into 	@URTCUsersID,
									@URTCUserName, 
									@URTCLastName , 
									@URTCFirstName 
		While @@FEtch_Status = 0
		BEGIN
			SET @sSQL5 = @sSQL5 + '(SELECT Count(URTC.UsersID) as CountOfUserID '	
			SET @sSQL5 = @sSQL5 + 'FROM UserReportsToCoordinator URTC '
			SET @sSQL5 = @sSQL5 + 'WHERE URTC.CompanyID = CU.[CompanyID] '
			SET @sSQL5 = @sSQL5 + 'AND URTC.UsersID = CU.[UsersID] '
			SET @sSQL5 = @sSQL5 + 'AND URTC.ReportsToUsersID = ' + Cast( @URTCUsersID as Varchar(10)) + ' '
			SET @sSQL5 = @sSQL5 + 'AND URTC.Active = 1 ' 
			SET @sSQL5 = @sSQL5 + 'GROUP BY URTC.CompanyID '
			SET @sSQL5 = @sSQL5 + ') As [UserReportsToCoordinator' + Cast( @URTCUserName as Varchar(30)) + '], '
			Fetch Next From CurUserReportsToCoordinator into 	@URTCUsersID,
										@URTCUserName, 
										@URTCLastName, 
										@URTCFirstName
		END
		Close CurUserReportsToCoordinator
		Deallocate CurUserReportsToCoordinator
		
	END
--^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
--*		End User Reports to Coordinator Association List		*
--*										*
--*******************************************************************************

--*******************************************************************************
--*			Begin User Assgn to CAT Association List		*
--*										*
--VVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV
	SET @sSQL6 = ''
	IF @selCompany Is Not Null And @selClientCompany Is Not Null And @selUserProfile Is not Null And @bViewHistory Is Null
	BEGIN
		DECLARE @UATCClientCompanyID int
		DECLARE @UATCCatID int
		DECLARE @UATCName varchar(20)
		DECLARE @UATCDescription varchar(255)
		DECLARE @UATCSiteAddress varchar(200)
		DECLARE @UATCSACity varchar(50)
		DECLARE @UATCSAState varchar(50)
		DECLARE @UATCSAZip int
		DECLARE @UATCSAZip4 int
		IF @bHideDeleted =1 
		BEGIN
			Declare CurUserAssgnToCAT Cursor for
			SELECT 
				UATC.[ClientCompanyID],
				UATC.[CatID],
				CAT.[Name], 
				CAT.[Description] , 
				UATC.[SiteAddress], 
				UATC.[SACity],
				UATC.[SAState], 
				UATC.[SAZip] , 
				UATC.[SAZip4] 
			FROM ClientCompanyCat UATC INNER JOIN CAT ON UATC.CATID = CAT.CATID And UATC.ClientCompanyID = @selClientCompany
			WHERE 	UATC.InactiveDate Is Null
			ORDER BY CAT.[Name]
		END
		ELSE
		BEGIN
			Declare CurUserAssgnToCAT Cursor for
			SELECT 
				UATC.[ClientCompanyID],
				UATC.[CatID],
				CAT.[Name], 
				CAT.[Description] , 
				UATC.[SiteAddress], 
				UATC.[SACity],
				UATC.[SAState], 
				UATC.[SAZip] , 
				UATC.[SAZip4] 
			FROM ClientCompanyCat UATC INNER JOIN CAT ON UATC.CATID = CAT.CATID And UATC.ClientCompanyID = @selClientCompany
			ORDER BY CAT.[Name]
		END
		Open CurUserAssgnToCAT
		Fetch Next From CurUserAssgnToCAT into 			
									@UATCClientCompanyID,
									@UATCCatID,
									@UATCName, 
									@UATCDescription , 
									@UATCSiteAddress,
									@UATCSACity,
									@UATCSAState,
									@UATCSAZip,
									@UATCSAZip4
		While @@FEtch_Status = 0
		BEGIN
			SET @sSQL6 = @sSQL6 + '(SELECT Count(UATC.UsersID) as CountOfUserID '	
			SET @sSQL6 = @sSQL6 + 'FROM ClientCompanyUsersCat UATC '
			SET @sSQL6 = @sSQL6 + 'WHERE UATC.ClientCompanyID = CU.[CompanyID] '
			SET @sSQL6 = @sSQL6 + 'AND UATC.ClientCompanyID = ' + Cast( @UATCClientCompanyID as Varchar(10)) + ' '
			SET @sSQL6 = @sSQL6 + 'AND UATC.CATID = ' + Cast( @UATCCatID as Varchar(10)) + ' '
			SET @sSQL6 = @sSQL6 + 'AND UATC.UsersID = CU.[UsersID] '
			SET @sSQL6 = @sSQL6 + 'AND UATC.Active = 1 ' 
			SET @sSQL6 = @sSQL6 + 'GROUP BY UATC.ClientCompanyID, UATC.CATID '
			SET @sSQL6 = @sSQL6 + ') As [UserAssgnToCAT' + Cast( @UATCName  as Varchar(20)) + '_' + Cast( @UATCClientCompanyID as Varchar(10)) + '_' + Cast( @UATCCatID as Varchar(10)) + '], '
			Fetch Next From CurUserAssgnToCAT into 			
										@UATCClientCompanyID,
										@UATCCatID,
										@UATCName, 
										@UATCDescription , 
										@UATCSiteAddress,
										@UATCSACity,
										@UATCSAState,
										@UATCSAZip,
										@UATCSAZip4
		END
		Close CurUserAssgnToCAT
		Deallocate CurUserAssgnToCAT
		
	END
--^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
--*			End User Assgn to CAT Association List			*
--*										*
--*******************************************************************************
	SET @sSQL7 = ''
	SET @sSQL8 = ''
	SET @sSQL9 = ''
	SET @sSQL10 =''
	SET @sSQL10 = 'U.[DateLastUpdated] As UsersDateLastUpdated, '
	SET @sSQL10 = @sSQL10 + '(SELECT	USERNAME '
	SET @sSQL10 = @sSQL10 + 'FROM 	USERS '
	SET @sSQL10 = @sSQL10 + 'WHERE 	USERSID = U.[UpdateByUserID]) As [UsersUpdateByUserName],  '
	SET @sSQL10 = @sSQL10 + 'U.[UpdateByUserID] As UsersUpdateByUserID, '
	SET @sSQL10 = @sSQL10 + 'CU.[DateLastUpdated], '
	SET @sSQL10 = @sSQL10 + '(SELECT	USERNAME '
	SET @sSQL10 = @sSQL10 + 'FROM 	USERS '
	SET @sSQL10 = @sSQL10 + 'WHERE 	USERSID = CU.[UpdateByUserID]) As [UpdateByUserName],  '
	SET @sSQL10 = @sSQL10 + 'CU.[UpdateByUserID] As UpdateByUserID '

	IF @bViewHistory Is Null
	BEGIN
		SET @sSQL10 = @sSQL10 + 'FROM Users U '
		IF @bViewHistory Is Null
		BEGIN
		--Include AdjusterUsersSoftware
			SET @sSQL10 = @sSQL10 + 'INNER JOIN AdjusterUsersSoftware AUS ON U.UsersID = AUS.UsersID '
		--Include AdjusterUSersUpdates
			SET @sSQL10 = @sSQL10 + 'INNER JOIN AdjusterUsersUpdates AUU ON U.UsersID = AUU.UsersID '
		END
		--Check for Selected Company 
		SET @sSQL10 = @sSQL10 + 'INNER JOIN CompanyUsers CU ON U.UsersID = CU.UsersID '
		SET @sSQL10 = @sSQL10 + 'AND CU.CompanyID = ' + Cast(@selClientORCompany  as varchar(10)) + ' '
		IF @selUserProfile Is not Null
		Begin
			SET @sSQL10 = @sSQL10 + 'INNER JOIN ' + (SELECT TableName FROM USERPROFILE WHERE UserProfileID = @selUserProfile) + ' UP2 '
			SET @sSQL10 = @sSQL10 + 'ON CU.CompanyID = UP2.CompanyID AND CU.UsersID = UP2.UsersID AND UP2.Active = 1 '
		END
		IF @selUserProfile Is not Null and @selUserReportsToManager Is not Null
		Begin
			SET @sSQL10 = @sSQL10 + 'INNER JOIN UserReportsToManager URTM ON UP2.UsersID = URTM.UsersID AND URTM.ReportsToUsersID = ' + Cast(@selUserReportsToManager  as varchar(10)) + ' ' 
			SET @sSQL10 = @sSQL10 + 'AND URTM.CompanyID = ' + Cast(@selClientORCompany  as varchar(10)) + ' '
			SET @sSQL10 = @sSQL10 + 'AND URTM.Active = 1 '
		END
		IF @selUserProfile Is not Null and @selUserReportsToCoordinator Is not Null
		Begin
			SET @sSQL10 = @sSQL10 + 'INNER JOIN UserReportsToCoordinator URTC ON UP2.UsersID = URTC.UsersID AND URTC.ReportsToUsersID = ' + Cast(@selUserReportsToCoordinator  as varchar(10)) + ' ' 
			SET @sSQL10 = @sSQL10 + 'AND URTC.CompanyID = ' + Cast(@selClientORCompany  as varchar(10)) + ' '
			SET @sSQL10 = @sSQL10 + 'AND URTC.Active = 1 '
		END
		IF @selUserProfile Is not Null And @selUserAssgnToCAT Is Not Null
		Begin
			SET @sSQL10 = @sSQL10 + 'INNER JOIN ClientCompanyUsersCat UATC '
			SET @sSQL10 = @sSQL10 + 'ON CU.UsersID = UATC.UsersID AND CU.CompanyID = UATC.ClientCompanyID AND UATC.CATID = ' + Cast(@selUserAssgnToCAT  as varchar(10)) + ' '
			SET @sSQL10 = @sSQL10 + 'AND UATC.Active = 1 '
		END
		IF @bHideDeleted = 1
		BEGIN
			SET @sSQL10 = @sSQL10 + 'AND CU.Active = 1 '
			SET @sSQL10 = @sSQL10 + 'AND U.Active = 1 '
		END

		SET SET @sSQL10 = @sSQL10 + ') RetUsers '
		IF @bHideDeleted = 1
		BEGIN
			SET @sSQL10 = @sSQL10 + 'WHERE [InactiveDate] Is Null '
		END
		ELSE
		BEGIN
			SET @sSQL10 = @sSQL10 + 'WHERE [ActiveDate] Is Not Null '
		END
		--Searching
		IF @SearchBy Is Not Null
		BEGIN
			SET @sSQL10 = @sSQL10 + @SearchBy + ' '
			--Searching Dates
			IF @FromActiveDate <> '*'
			BEGIN
				SET @sSQL10 = @sSQL10 + 'AND [ActiveDate] >= Convert(datetime,''' + @FromActiveDate + ''') '
			END
			IF @ToActiveDate <> '*'
			BEGIN
				SET @sSQL10 = @sSQL10 + 'AND [ActiveDate] <= Convert(datetime,''' + @ToActiveDate + ''') '
			END
			IF @FromInactiveDate <> '*'
			BEGIN
				SET @sSQL10 = @sSQL10 + 'AND [InactiveDate] >= Convert(datetime,''' + @FromInactiveDate + ''') '
			END
			IF @ToInactiveDate <> '*'
			BEGIN
				SET @sSQL10 = @sSQL10 + 'AND [InactiveDate] <= Convert(datetime,''' + @ToInactiveDate + ''') '
			END
		END
	END
	ELSE
	BEGIN
		SET @sSQL10 = @sSQL10 + 'FROM CompanyUsers CUS '

		--Check for Selected Company 
		SET @sSQL10 = @sSQL10 + 'LEFT JOIN Users U ON U.UsersID = CUS.UsersID '
		SET @sSQL10 = @sSQL10 +'LEFT JOIN CompanyUsersHistory CU On CU.UsersID = CUS.UsersID AND CU.CompanyID = CUS.CompanyID '
		SET @sSQL10 = @sSQL10 + 'AND CU.CompanyID = ' + Cast(@selClientORCompany  as varchar(10)) + ' '
		IF @bHideDeleted = 1
		BEGIN
			SET @sSQL10 = @sSQL10 + 'AND CU.Active = 1 '
			SET @sSQL10 = @sSQL10 + 'AND U.Active = 1 '
		END

		SET @sSQL10 = @sSQL10 + ') RetUsers '
		SET @sSQL10 = @sSQL10 + 'WHERE [DateLastUpdated] >= Convert(datetime,''' + @FromDate + ''') '
		SET @sSQL10 = @sSQL10 + 'AND [DateLastUpdated] <= Convert(datetime,''' + @ToDate + ''') '
		IF @UsersID <> -1 
		BEGIN
			SET @sSQL10 = @sSQL10 + 'AND [UsersID] = ' + Cast(@UsersID as varchar(10)) + ' '
		END
		--Searching
		IF @SearchBy Is Not Null
		BEGIN
			SET @sSQL10 = @sSQL10 + @SearchBy + ' '
			--Searching Dates
			IF @FromActiveDate <> '*'
			BEGIN
				SET @sSQL10 = @sSQL10 + 'AND [ActiveDate] >= Convert(datetime,''' + @FromActiveDate + ''') '
			END



			IF @ToActiveDate <> '*'
			BEGIN
				SET @sSQL10 = @sSQL10 + 'AND [ActiveDate] <= Convert(datetime,''' + @ToActiveDate + ''') '
			END
			IF @FromInactiveDate <> '*'
			BEGIN
				SET @sSQL10 = @sSQL10 + 'AND [InactiveDate] >= Convert(datetime,''' + @FromInactiveDate + ''') '
			END
			IF @ToInactiveDate <> '*'
			BEGIN
				SET @sSQL10 = @sSQL10 + 'AND [InactiveDate] <= Convert(datetime,''' + @ToInactiveDate + ''') '
			END
		END
		
	END
	IF @bViewHistory Is Null
	BEGIN
		--Grouping
		IF @GroupBy Is Not Null
		BEGIN
			SET @sSQL10 = @sSQL10 + 'GROUP BY ' + @GroupBy + ' '			
		END
		---Sorting 
		IF @GroupBy Is Null
		BEGIN
			IF @OrderBy Is Null OR @OrderBy = 'Null'
			BEGIN
				SET @sSQL10 = @sSQL10 + 'ORDER BY [Active] Desc, [LastName] '
			END
			ELSE
			BEGIN
				SET @sSQL10 = @sSQL10 + 'ORDER BY  [Active] Desc, ' + @OrderBy + ' '
			END
		END
	END
	ELSE
	BEGIN
		--Grouping
		IF @GroupBy Is Not Null
		BEGIN
			SET @sSQL10 = @sSQL10 + 'GROUP BY ' + @GroupBy + ' '			
		END
		---Sorting
		IF @GroupBy Is Null
		BEGIN
			SET @sSQL10 = @sSQL10 + 'ORDER BY  ' + @OrderBy + ' '
		END
	END
	Print Len(@sSQL1) + Len(@sSQL2) + Len(@sSQL3) + Len(@sSQL4) + Len(@sSQL5) + Len(@sSQL6) + Len(@sSQL7) + Len(@sSQL8) + Len(@sSQL9) + Len(@sSQL10)
	Print @sSQL1
	Print @sSQL2
	Print @sSQL3
	Print @sSQL4
	Print @sSQL5
	Print @sSQL6
	Print @sSQL7
	Print @sSQL8
	Print @sSQL9
	Print @sSQL10
	EXEC(@sSQL1 + @sSQL2 + @sSQL3 + @sSQL4 + @sSQL5 + @sSQL6 + @sSQL7 + @sSQL8 + @sSQL9 + @sSQL10)
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[z_spsGetCompanyUsersInfo]  TO [CFUSER]
GO

