SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[z_spsGetProdTableCounts]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[z_spsGetProdTableCounts]
GO

/******************************************************************
*
*    	Description:	Get All the Productions Table Counts,
*			These inscludes the Assignments table 
*			and it's Dependant Tables.
*	
*	Author: 	Brad Skidmore
*	Date: 		5/12/2005
*
******************************************************************/
CREATE PROCEDURE [dbo].[z_spsGetProdTableCounts]
AS
SELECT 	
	(SELECT Count(AssignmentsID) FROM Assignments )  		As AssignmentsCount,
	(SELECT Count(AssignmentsID) FROM AssignmentsHistory )  	As AssignmentsHistoryCount,
	(SELECT Count(AssignmentsID) FROM BillingCount )  		As BillingCountCount,
	(SELECT Count(AssignmentsID) FROM PolicyLimits )  		As PolicyLimitsCount,
	(SELECT Count(AssignmentsID) FROM PolicyLimitsHistory )  	As PolicyLimitsHistoryCount,
	(SELECT Count(AssignmentsID) FROM RTIB )  			As RTIBCount,
	(SELECT Count(AssignmentsID) FROM RTIBHistory )  		As RTIBHistoryCount,
	(SELECT Count(AssignmentsID) FROM RTIBFee )  			As RTIBFeeCount,
	(SELECT Count(AssignmentsID) FROM RTIBFeeHistory )  		As RTIBFeeHistoryCount,
	(SELECT Count(AssignmentsID) FROM IB )  			As IBCount,
	(SELECT Count(AssignmentsID) FROM IBHistory )  			As IBHistoryCount,
	(SELECT Count(AssignmentsID) FROM IBFee )  			As IBFeeCount,
	(SELECT Count(AssignmentsID) FROM IBFeeHistory )  		As IBFeeHistoryCount,
	(SELECT Count(AssignmentsID) FROM RTChecks )  			As RTChecksCount,
	(SELECT Count(AssignmentsID) FROM RTChecksHistory )  		As RTChecksHistoryCount,
	(SELECT Count(AssignmentsID) FROM RTIndemnity )  		As RTIndemnityCount,
	(SELECT Count(AssignmentsID) FROM RTIndemnityHistory )  	As RTIndemnityHistoryCount,
	(SELECT Count(AssignmentsID) FROM RTActivityLog )  		As RTActivityLogCount,
	(SELECT Count(AssignmentsID) FROM RTActivityLogHistory )  	As RTActivityLogHistoryCount,
	(SELECT Count(AssignmentsID) FROM RTActivityLogInfo )  		As RTActivityLogInfoCount,
	(SELECT Count(AssignmentsID) FROM RTActivityLogInfoHistory )  	As RTActivityLogInfoHistoryCount,
	(SELECT Count(AssignmentsID) FROM RTPhotoReport )  		As RTPhotoReportCount,
	(SELECT Count(AssignmentsID) FROM RTPhotoReportHistory )  	As RTPhotoReportHistoryCount,
	(SELECT Count(AssignmentsID) FROM RTPhotoLog )  		As RTPhotoLogCount,
	(SELECT Count(AssignmentsID) FROM RTPhotoLogHistory )  		As RTPhotoLogHistoryCount,
	(SELECT Count(AssignmentsID) FROM RTWSDiagram )  		As RTWSDiagramCount,
	(SELECT Count(AssignmentsID) FROM RTWSDiagramHistory )  	As RTWSDiagramHistoryCount,
	(SELECT Count(AssignmentsID) FROM RTAttachments )  		As RTAttachmentsCount,
	(SELECT Count(AssignmentsID) FROM RTAttachmentsHistory )	As RTAttachmentsHistoryCount,
	(SELECT Count(AssignmentsID) FROM MiscReportParam )  		As MiscReportParamCount,
	(SELECT Count(AssignmentsID) FROM MiscReportParam01 )  		As MiscReportParam01Count,
	(SELECT Count(AssignmentsID) FROM MiscReportParam02 )  		As MiscReportParam02Count,
	(SELECT Count(AssignmentsID) FROM MiscReportParam03 )  		As MiscReportParam03Count,
	(SELECT Count(AssignmentsID) FROM MiscReportParam04 )  		As MiscReportParam04Count,
	(SELECT Count(AssignmentsID) FROM MiscReportParam05 )  		As MiscReportParam05Count,
	(SELECT Count(AssignmentsID) FROM MiscReportParam06 )  		As MiscReportParam06Count,
	(SELECT Count(AssignmentsID) FROM MiscReportParam07 )  		As MiscReportParam07Count,
	(SELECT Count(AssignmentsID) FROM MiscReportParam08 )  		As MiscReportParam08Count,
	(SELECT Count(AssignmentsID) FROM MiscReportParam09 )  		As MiscReportParam09Count,
	(SELECT Count(AssignmentsID) FROM MiscReportParam10 )  		As MiscReportParam10Count,
	(SELECT Count(AssignmentsID) FROM MiscReportParam11 )  		As MiscReportParam11Count,
	(SELECT Count(AssignmentsID) FROM MiscReportParam12 )  		As MiscReportParam12Count,
	(SELECT Count(AssignmentsID) FROM MiscReportParam13 )  		As MiscReportParam13Count,
	(SELECT Count(AssignmentsID) FROM MiscReportParam14 )  		As MiscReportParam14Count,
	(SELECT Count(AssignmentsID) FROM MiscReportParam15 )  		As MiscReportParam15Count,
	(SELECT Count(AssignmentsID) FROM MiscReportParam16 )  		As MiscReportParam16Count,
	(SELECT Count(AssignmentsID) FROM MiscReportParam17 )  		As MiscReportParam17Count,
	(SELECT Count(AssignmentsID) FROM MiscReportParam18 )  		As MiscReportParam18Count,
	(SELECT Count(AssignmentsID) FROM MiscReportParam19 )  		As MiscReportParam19Count,
	(SELECT Count(AssignmentsID) FROM MiscReportParam20 )  		As MiscReportParam20Count,
	(SELECT Count(AssignmentsID) FROM MiscReportParam21 )  		As MiscReportParam21Count,
	(SELECT Count(AssignmentsID) FROM MiscReportParam22 )  		As MiscReportParam22Count,
	(SELECT Count(AssignmentsID) FROM MiscReportParam23 )  		As MiscReportParam23Count,
	(SELECT Count(AssignmentsID) FROM MiscReportParam24 )  		As MiscReportParam24Count,
	(SELECT Count(AssignmentsID) FROM MiscReportParam25 )  		As MiscReportParam25Count,
	(SELECT Count(AssignmentsID) FROM MiscReportParam26 )  		As MiscReportParam26Count,
	(SELECT Count(AssignmentsID) FROM MiscReportParam27 )  		As MiscReportParam27Count,
	(SELECT Count(AssignmentsID) FROM MiscReportParam28 )  		As MiscReportParam28Count,
	(SELECT Count(AssignmentsID) FROM MiscReportParam29 )  		As MiscReportParam29Count,
	(SELECT Count(AssignmentsID) FROM MiscReportParam30 )  		As MiscReportParam30Count,
	(SELECT Count(AssignmentsID) FROM Package )  			As PackageCount,
	(SELECT Count(AssignmentsID) FROM PackageHistory )  		As PackageHistoryCount,
	(SELECT Count(AssignmentsID) FROM PackageItem )  		As PackageItemCount,
	(SELECT Count(AssignmentsID) FROM PackageItemHistory )  	As PackageItemHistoryCount,
	(SELECT Count(AssignmentsID) FROM Batches )  			As BatchesCount,
	(SELECT Count(AssignmentsID) FROM BatchesHistory )  		As BatchesHistoryCount,
	(SELECT Count(AssignmentsID) FROM FarmersXML01_LocationUpdateQue )  		As FarmersXML01_LocationUpdateQueCount,
	(SELECT Count(AssignmentsID) FROM FarmersXML01_ActivityUpdateQue )  		As FarmersXML01_ActivityUpdateQueCount,
--************************************BillAssignments*****************************************
	(SELECT Count(BillAssignmentID) FROM BillAssignment )  				As BillAssignmentCount,
	(SELECT Count(BillAssignmentID) FROM BillBillingCount )  			As BillBillingCountCount,
	(SELECT Count(BillAssignmentID) FROM IBStateFarm )  				As IBStateFarmCount,
--************************************LookUp Tables*******************************************
	(SELECT Count(Version) FROM  DB_VERSION )  					As DB_VERSIONCount,
	(SELECT Count(AssignmentTypeID) FROM AssignmentType )  				As AssignmentTypeCount,
	(SELECT Count(AssignmentTypeID) FROM AssignmentTypeHistory )  			As AssignmentTypeHistoryCount,
	(SELECT Count(StatusID) FROM Status )  						As StatusCount,
	(SELECT Count(StatusID) FROM StatusHistory )  					As StatusHistoryCount,
	(SELECT Count(Zip) FROM Zips )  						As ZipsCount,
	(SELECT Count([SecurityLevel]) FROM SecurityLevel )  				As SecurityLevelCount,
	(SELECT Count([SecurityLevel]) FROM SecurityLevelHistory )  			As SecurityLevelHistoryCount,
	(SELECT Count(FAQSID) FROM FAQS )  						As FAQSCount,
	(SELECT Count(FAQSID) FROM FAQSHistory )  					As FAQSHistoryCount,
	(SELECT Count(FeeScheduleID) FROM FeeSchedule )  				As FeeScheduleCount,
	(SELECT Count(FeeScheduleID) FROM FeeScheduleHistory )  			As FeeScheduleHistoryCount,
	(SELECT Count(FeeScheduleID) FROM FeeScheduleFeeTypes )  			As FeeScheduleFeeTypesCount,
	(SELECT Count(FeeScheduleID) FROM FeeScheduleFeeTypesHistory )  		As FeeScheduleFeeTypesHistoryCount,
	(SELECT Count(FeeScheduleID) FROM FeeScheduleLevels )  				As FeeScheduleLevelsCount,
	(SELECT Count(FeeScheduleID) FROM FeeScheduleLevelsHistory )  			As FeeScheduleLevelsHistoryCount,
	(SELECT Count(TypeOfLossID) FROM TypeOfLoss )  					As TypeOfLossCount,
	(SELECT Count(TypeOfLossID) FROM TypeOfLossHistory )  				As TypeOfLossHistoryCount,
	(SELECT Count(ClassTypeID) FROM ClassType )  					As ClassTypeCount,
	(SELECT Count(ClassTypeID) FROM ClassTypeHistory )  				As ClassTypeHistoryCount,
	(SELECT Count(ClassOfLossID) FROM ClassOfLoss )  				As ClassOfLossCount,
	(SELECT Count(ClassOfLossID) FROM ClassOfLossHistory )  			As ClassOfLossHistoryCount,
	(SELECT Count(StateID) FROM State )  						As StateCount,
	(SELECT Count(StateID) FROM StateHistory )  					As StateHistoryCount,
	(SELECT Count(CompanyID) FROM Company )  					As CompanyCount,
	(SELECT Count(CompanyID) FROM CompanyHistory )  				As CompanyHistoryCount,
	(SELECT Count(CATID) FROM CAT )  						As CATCount,
	(SELECT Count(CATID) FROM CATHistory )  					As CATHistoryCount,
	(SELECT Count(CATID) FROM ClientCompanyCatSpec )  				As ClientCompanyCatSpecCount,
	(SELECT Count(CATID) FROM ClientCompanyCatSpecHistory )  			As ClientCompanyCatSpecHistoryCount,
	(SELECT Count(CATID) FROM ClientCompanyCat )  					As ClientCompanyCatCount,
	(SELECT Count(CATID) FROM ClientCompanyCatHistory )  				As ClientCompanyCatHistoryCount,
--************************************Software************************************************
	(SELECT Count(SoftwarePackageID) FROM SoftwarePackage )  			As SoftwarePackageCount,
	(SELECT Count(SoftwarePackageID) FROM SoftwarePackageHistory )  		As SoftwarePackageHistoryCount,
	(SELECT Count(SoftwarePackageID) FROM SoftwarePackageApplication )  		As SoftwarePackageApplicationCount,
	(SELECT Count(SoftwarePackageID) FROM SoftwarePackageDocument )  		As SoftwarePackageDocumentCount,
	(SELECT Count(SoftwarePackageID) FROM SoftwarePackageRegSetting )  		As SoftwarePackageRegSettingCount,
	(SELECT Count(ApplicationID) FROM Application )  				As ApplicationCount,
	(SELECT Count(ApplicationID) FROM ApplicationHistory )  			As ApplicationHistoryCount,
	(SELECT Count(DocumentID) FROM Document )  					As DocumentCount,
	(SELECT Count(DocumentID) FROM DocumentHistory )  				As DocumentHistoryCount,
	(SELECT Count(RegSettingID) FROM RegSetting )  					As RegSettingCount,
	(SELECT Count(RegSettingID) FROM RegSettingHistory )  				As RegSettingHistoryCount,
--************************************ECS ADJ*************************************************
	(SELECT Count(UsersID) FROM ECSADJUsers )  					As ECSADJUsersCount,
	(SELECT Count(EvaluationID) FROM AdjusterEvaluations )  			As AdjusterEvaluationsCount,
	(SELECT Count(EvaluationID) FROM SF_AdjusterEvaluations )  			As SF_AdjusterEvaluationsCount,
--************************************User Tables*********************************************
	(SELECT Count(UsersID) FROM CompanyUsers )  					As CompanyUsersCount,
	(SELECT Count(UsersID) FROM ClientCompanyUsersCat )  				As ClientCompanyUsersCatCount, 
	(SELECT Count(UsersID) FROM Users )  						As UsersCount,
	(SELECT Count(UsersID) FROM UsersHistory )  					As UsersHistoryCount,
	(SELECT Count(UserProfileID) FROM UserProfile )  				As UserProfileCount,  
	(SELECT Count(UserProfileID) FROM UserProfileHistory  )  			As UserProfileHistoryCount, 
	(SELECT Count(UsersID) FROM AdjusterUsersSoftware )  				As AdjusterUsersSoftwareCount, 
	(SELECT Count(UsersID) FROM AdjusterUsersSoftwareHistory )  			As AdjusterUsersSoftwareHistoryCount,  
	(SELECT Count(UsersID) FROM AdjusterUsersUpdates )  				As AdjusterUsersUpdatesCount,  
	(SELECT Count(UsersID) FROM AdjusterUsersUpdatesHistory )  			As AdjusterUsersUpdatesHistoryCount,  
	(SELECT Count(UsersID) FROM Client )  						As ClientCount,  
	(SELECT Count(UsersID) FROM Admin )  						As AdminCount,  
	(SELECT Count(UsersID) FROM Coordinator )  					As CoordinatorCount,  
	(SELECT Count(UsersID) FROM UserReportsToCoordinator )  			As UserReportsToCoordinatorCount,  
	(SELECT Count(UsersID) FROM Manager )  						As ManagerCount,  
	(SELECT Count(UsersID) FROM UserReportsToManager )  				As UserReportsToManagerCount,  
	(SELECT Count(UsersID) FROM Adjuster )  					As AdjusterCount,  
	(SELECT Count(ClientCoAdjusterSpecID) FROM ClientCoAdjusterSpec )  		As ClientCoAdjusterSpecCount, 
	(SELECT Count(ClientCoAdjusterSpecID) FROM ClientCoAdjusterSpecHistory )  	As ClientCoAdjusterSpecHistoryCount,  
	(SELECT Count(UsersID) FROM Dispatcher )  					As DispatcherCount,  
	(SELECT Count(UsersID) FROM Employee )  					As EmployeeCount,  
	(SELECT Count(UsersID) FROM [Temporary] )  					As TemporaryCount,  
	(SELECT Count(UsersID) FROM Accounting  )  					As AccountingCount	
 
	

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[z_spsGetProdTableCounts]  TO [CFUSER]
GO

