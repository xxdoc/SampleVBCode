if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_UserCompanySpec_CompanyAdjusterUsers]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[ClientCoAdjusterSpec] DROP CONSTRAINT FK_UserCompanySpec_CompanyAdjusterUsers
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_UserReportsToManager_Adjuster]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[UserReportsToManager] DROP CONSTRAINT FK_UserReportsToManager_Adjuster
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Adjuster_AdjusterUsersSoftware]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[Adjuster] DROP CONSTRAINT FK_Adjuster_AdjusterUsersSoftware
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_UsersSoftwareHistory_UsersSoftware]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[AdjusterUsersSoftwareHistory] DROP CONSTRAINT FK_UsersSoftwareHistory_UsersSoftware
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Adjuster_AdjusterUsersUpdates]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[Adjuster] DROP CONSTRAINT FK_Adjuster_AdjusterUsersUpdates
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_UsersUpdatesHistory_UsersUpdates]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[AdjusterUsersUpdatesHistory] DROP CONSTRAINT FK_UsersUpdatesHistory_UsersUpdates
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_ApplicationHistory_Application]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[ApplicationHistory] DROP CONSTRAINT FK_ApplicationHistory_Application
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_SoftwarePackageApplication_Application]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[SoftwarePackageApplication] DROP CONSTRAINT FK_SoftwarePackageApplication_Application
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Assignments_AssignmentType]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[Assignments] DROP CONSTRAINT FK_Assignments_AssignmentType
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_AssignmentTypeHistory_AssignmentType]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[AssignmentTypeHistory] DROP CONSTRAINT FK_AssignmentTypeHistory_AssignmentType
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_BillAssignment_AssignmentType]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[BillAssignment] DROP CONSTRAINT FK_BillAssignment_AssignmentType
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_CAT_AssignmentType]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[CAT] DROP CONSTRAINT FK_CAT_AssignmentType
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_AssignmentsHistory_Assignments]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[AssignmentsHistory] DROP CONSTRAINT FK_AssignmentsHistory_Assignments
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Batches_Assignments]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[Batches] DROP CONSTRAINT FK_Batches_Assignments
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_BillingCount_Assignments]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[BillingCount] DROP CONSTRAINT FK_BillingCount_Assignments
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_IB_Assignments]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[IB] DROP CONSTRAINT FK_IB_Assignments
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_IBFee_Assignments]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[IBFee] DROP CONSTRAINT FK_IBFee_Assignments
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_MiscReportParam_Assignments]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[MiscReportParam] DROP CONSTRAINT FK_MiscReportParam_Assignments
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_MiscReportParam01_Assignments]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[MiscReportParam01] DROP CONSTRAINT FK_MiscReportParam01_Assignments
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_MiscReportParam02_Assignments]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[MiscReportParam02] DROP CONSTRAINT FK_MiscReportParam02_Assignments
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_MiscReportParam03_Assignments]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[MiscReportParam03] DROP CONSTRAINT FK_MiscReportParam03_Assignments
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_MiscReportParam04_Assignments]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[MiscReportParam04] DROP CONSTRAINT FK_MiscReportParam04_Assignments
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_MiscReportParam05_Assignments]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[MiscReportParam05] DROP CONSTRAINT FK_MiscReportParam05_Assignments
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_MiscReportParam06_Assignments]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[MiscReportParam06] DROP CONSTRAINT FK_MiscReportParam06_Assignments
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_MiscReportParam07_Assignments]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[MiscReportParam07] DROP CONSTRAINT FK_MiscReportParam07_Assignments
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_MiscReportParam08_Assignments]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[MiscReportParam08] DROP CONSTRAINT FK_MiscReportParam08_Assignments
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_MiscReportParam09_Assignments]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[MiscReportParam09] DROP CONSTRAINT FK_MiscReportParam09_Assignments
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_MiscReportParam10_Assignments]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[MiscReportParam10] DROP CONSTRAINT FK_MiscReportParam10_Assignments
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_MiscReportParam11_Assignments]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[MiscReportParam11] DROP CONSTRAINT FK_MiscReportParam11_Assignments
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_MiscReportParam12_Assignments]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[MiscReportParam12] DROP CONSTRAINT FK_MiscReportParam12_Assignments
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_MiscReportParam13_Assignments]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[MiscReportParam13] DROP CONSTRAINT FK_MiscReportParam13_Assignments
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_MiscReportParam14_Assignments]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[MiscReportParam14] DROP CONSTRAINT FK_MiscReportParam14_Assignments
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_MiscReportParam15_Assignments]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[MiscReportParam15] DROP CONSTRAINT FK_MiscReportParam15_Assignments
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_MiscReportParam16_Assignments]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[MiscReportParam16] DROP CONSTRAINT FK_MiscReportParam16_Assignments
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_MiscReportParam17_Assignments]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[MiscReportParam17] DROP CONSTRAINT FK_MiscReportParam17_Assignments
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_MiscReportParam18_Assignments]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[MiscReportParam18] DROP CONSTRAINT FK_MiscReportParam18_Assignments
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_MiscReportParam19_Assignments]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[MiscReportParam19] DROP CONSTRAINT FK_MiscReportParam19_Assignments
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_MiscReportParam20_Assignments]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[MiscReportParam20] DROP CONSTRAINT FK_MiscReportParam20_Assignments
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_MiscReportParam21_Assignments]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[MiscReportParam21] DROP CONSTRAINT FK_MiscReportParam21_Assignments
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_MiscReportParam22_Assignments]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[MiscReportParam22] DROP CONSTRAINT FK_MiscReportParam22_Assignments
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_MiscReportParam23_Assignments]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[MiscReportParam23] DROP CONSTRAINT FK_MiscReportParam23_Assignments
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_MiscReportParam24_Assignments]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[MiscReportParam24] DROP CONSTRAINT FK_MiscReportParam24_Assignments
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_MiscReportParam25_Assignments]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[MiscReportParam25] DROP CONSTRAINT FK_MiscReportParam25_Assignments
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_MiscReportParam26_Assignments]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[MiscReportParam26] DROP CONSTRAINT FK_MiscReportParam26_Assignments
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_MiscReportParam27_Assignments]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[MiscReportParam27] DROP CONSTRAINT FK_MiscReportParam27_Assignments
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_MiscReportParam28_Assignments]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[MiscReportParam28] DROP CONSTRAINT FK_MiscReportParam28_Assignments
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_MiscReportParam29_Assignments]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[MiscReportParam29] DROP CONSTRAINT FK_MiscReportParam29_Assignments
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_MiscReportParam30_Assignments]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[MiscReportParam30] DROP CONSTRAINT FK_MiscReportParam30_Assignments
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Package_Assignments]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[Package] DROP CONSTRAINT FK_Package_Assignments
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_PackageItem_Assignments]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[PackageItem] DROP CONSTRAINT FK_PackageItem_Assignments
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_PolicyLimits_Assignments]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[PolicyLimits] DROP CONSTRAINT FK_PolicyLimits_Assignments
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_RTActivityLog_Assignments]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[RTActivityLog] DROP CONSTRAINT FK_RTActivityLog_Assignments
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_RTActivityLogInfo_Assignments]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[RTActivityLogInfo] DROP CONSTRAINT FK_RTActivityLogInfo_Assignments
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_RTAttachments_Assignments]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[RTAttachments] DROP CONSTRAINT FK_RTAttachments_Assignments
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_RTChecks_Assignments]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[RTChecks] DROP CONSTRAINT FK_RTChecks_Assignments
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_RTIB_Assignments]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[RTIB] DROP CONSTRAINT FK_RTIB_Assignments
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_RTIBFee_Assignments]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[RTIBFee] DROP CONSTRAINT FK_RTIBFee_Assignments
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_RTIndemnity_Assignments]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[RTIndemnity] DROP CONSTRAINT FK_RTIndemnity_Assignments
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_RTPhotoLog_Assignments]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[RTPhotoLog] DROP CONSTRAINT FK_RTPhotoLog_Assignments
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_RTPhotoReport_Assignments]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[RTPhotoReport] DROP CONSTRAINT FK_RTPhotoReport_Assignments
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_RTWSDiagram_Assignments]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[RTWSDiagram] DROP CONSTRAINT FK_RTWSDiagram_Assignments
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_XML_Trans_Assignments]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[XML_Trans] DROP CONSTRAINT FK_XML_Trans_Assignments
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_BatchesHistory_Batches]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[BatchesHistory] DROP CONSTRAINT FK_BatchesHistory_Batches
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Batches_BillAssignment]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[Batches] DROP CONSTRAINT FK_Batches_BillAssignment
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_BillBillingCount_BillAssignment]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[BillBillingCount] DROP CONSTRAINT FK_BillBillingCount_BillAssignment
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_IBStateFarm_BillAssignment]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[IBStateFarm] DROP CONSTRAINT FK_IBStateFarm_BillAssignment
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_IBStateFarm_BillBillingCount]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[IBStateFarm] DROP CONSTRAINT FK_IBStateFarm_BillBillingCount
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_IB_BillingCount]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[IB] DROP CONSTRAINT FK_IB_BillingCount
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_RTActivityLog_BillingCount]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[RTActivityLog] DROP CONSTRAINT FK_RTActivityLog_BillingCount
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_RTChecks_BillingCount]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[RTChecks] DROP CONSTRAINT FK_RTChecks_BillingCount
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_RTIB_BillingCount]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[RTIB] DROP CONSTRAINT FK_RTIB_BillingCount
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_RTPhotoLog_BillingCount]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[RTPhotoLog] DROP CONSTRAINT FK_RTPhotoLog_BillingCount
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_CATHistory_CAT]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[CATHistory] DROP CONSTRAINT FK_CATHistory_CAT
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_CompanyCat_CAT]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[ClientCompanyCat] DROP CONSTRAINT FK_CompanyCat_CAT
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_CompanyCatSpec_CAT]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[ClientCompanyCatSpec] DROP CONSTRAINT FK_CompanyCatSpec_CAT
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_ClassOfLoss_ClassOfLoss]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[ClassOfLoss] DROP CONSTRAINT FK_ClassOfLoss_ClassOfLoss
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_ClassOfLossHistory_ClassOfLoss1]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[ClassOfLossHistory] DROP CONSTRAINT FK_ClassOfLossHistory_ClassOfLoss1
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_RTIndemnity_ClassOfLoss]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[RTIndemnity] DROP CONSTRAINT FK_RTIndemnity_ClassOfLoss
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_ClassOfLoss_ClassType]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[ClassOfLoss] DROP CONSTRAINT FK_ClassOfLoss_ClassType
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_ClassTypeHistory_ClassType]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[ClassTypeHistory] DROP CONSTRAINT FK_ClassTypeHistory_ClassType
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_PolicyLimits_ClassType]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[PolicyLimits] DROP CONSTRAINT FK_PolicyLimits_ClassType
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Assignments_ClientCoAdjusterSpec]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[Assignments] DROP CONSTRAINT FK_Assignments_ClientCoAdjusterSpec
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_BillAssignment_ClientCoAdjusterSpec]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[BillAssignment] DROP CONSTRAINT FK_BillAssignment_ClientCoAdjusterSpec
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_AdjusterSpecHistory_AdjusterSpec]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[ClientCoAdjusterSpecHistory] DROP CONSTRAINT FK_AdjusterSpecHistory_AdjusterSpec
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_CompanyCatHistory_CompanyCat]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[ClientCompanyCatHistory] DROP CONSTRAINT FK_CompanyCatHistory_CompanyCat
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_CompanyCatSpec_CompanyCat]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[ClientCompanyCatSpec] DROP CONSTRAINT FK_CompanyCatSpec_CompanyCat
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_UsrersCat_CompanyCat]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[ClientCompanyUsersCat] DROP CONSTRAINT FK_UsrersCat_CompanyCat
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_SoftwarePackage_ClientCompanyCat]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[SoftwarePackage] DROP CONSTRAINT FK_SoftwarePackage_ClientCompanyCat
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Assignments_ClientCompanyCatSpec]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[Assignments] DROP CONSTRAINT FK_Assignments_ClientCompanyCatSpec
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Batches_CompanyCatSpec]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[Batches] DROP CONSTRAINT FK_Batches_CompanyCatSpec
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_BillAssignment_ClientCompanyCatSpec]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[BillAssignment] DROP CONSTRAINT FK_BillAssignment_ClientCompanyCatSpec
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_ClientCoAdjusterSpec_ClientCompanyCatSpec]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[ClientCoAdjusterSpec] DROP CONSTRAINT FK_ClientCoAdjusterSpec_ClientCompanyCatSpec
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_CompanyCatSpecHistory_CompanyCatSpec]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[ClientCompanyCatSpecHistory] DROP CONSTRAINT FK_CompanyCatSpecHistory_CompanyCatSpec
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_CAT_Company]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[CAT] DROP CONSTRAINT FK_CAT_Company
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_ClassOfLoss_Company]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[ClassOfLoss] DROP CONSTRAINT FK_ClassOfLoss_Company
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_CompanyCat_Company1]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[ClientCompanyCat] DROP CONSTRAINT FK_CompanyCat_Company1
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_CompanyCatSpec_Company]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[ClientCompanyCatSpec] DROP CONSTRAINT FK_CompanyCatSpec_Company
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Company_Company]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[Company] DROP CONSTRAINT FK_Company_Company
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_CompanyHistory_Company]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[CompanyHistory] DROP CONSTRAINT FK_CompanyHistory_Company
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_CompanyUsers_Company]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[CompanyUsers] DROP CONSTRAINT FK_CompanyUsers_Company
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_FeeSchedule_Company]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[FeeSchedule] DROP CONSTRAINT FK_FeeSchedule_Company
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_TransType_Company]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[TransType] DROP CONSTRAINT FK_TransType_Company
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_TypeOfLoss_Company]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[TypeOfLoss] DROP CONSTRAINT FK_TypeOfLoss_Company
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_XML_Trans_Company]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[XML_Trans] DROP CONSTRAINT FK_XML_Trans_Company
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_XML_Trans_Company1]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[XML_Trans] DROP CONSTRAINT FK_XML_Trans_Company1
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_XML_Trans_Company2]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[XML_Trans] DROP CONSTRAINT FK_XML_Trans_Company2
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Accounting_CompanyUsers]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[Accounting] DROP CONSTRAINT FK_Accounting_CompanyUsers
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_CompanyAdjusterUsers_CompanyUsers]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[Adjuster] DROP CONSTRAINT FK_CompanyAdjusterUsers_CompanyUsers
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_CompanyAdminUsers_CompanyUsers]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[Admin] DROP CONSTRAINT FK_CompanyAdminUsers_CompanyUsers
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Client_CompanyUsers]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[Client] DROP CONSTRAINT FK_Client_CompanyUsers
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_CompanyUsersCat_CompanyUsers]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[ClientCompanyUsersCat] DROP CONSTRAINT FK_CompanyUsersCat_CompanyUsers
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_CompanyUsersHistory_CompanyUsers]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[CompanyUsersHistory] DROP CONSTRAINT FK_CompanyUsersHistory_CompanyUsers
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Coordinator_CompanyUsers]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[Coordinator] DROP CONSTRAINT FK_Coordinator_CompanyUsers
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Dispatcher_CompanyUsers]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[Dispatcher] DROP CONSTRAINT FK_Dispatcher_CompanyUsers
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Employee_CompanyUsers]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[Employee] DROP CONSTRAINT FK_Employee_CompanyUsers
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_CompanyManagerUsers_CompanyUsers]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[Manager] DROP CONSTRAINT FK_CompanyManagerUsers_CompanyUsers
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Temporary_CompanyUsers]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[Temporary] DROP CONSTRAINT FK_Temporary_CompanyUsers
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_UserReportsToCoordinator_Coordinator]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[UserReportsToCoordinator] DROP CONSTRAINT FK_UserReportsToCoordinator_Coordinator
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_DocumentHistory_Document]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[DocumentHistory] DROP CONSTRAINT FK_DocumentHistory_Document
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_SoftwarePackageDocument_Document]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[SoftwarePackageDocument] DROP CONSTRAINT FK_SoftwarePackageDocument_Document
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_FAQSHistory_FAQS]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[FAQSHistory] DROP CONSTRAINT FK_FAQSHistory_FAQS
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_CompanyCat_FeeSchedule]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[ClientCompanyCat] DROP CONSTRAINT FK_CompanyCat_FeeSchedule
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_FeeScheduleFeeTypes_FeeSchedule]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[FeeScheduleFeeTypes] DROP CONSTRAINT FK_FeeScheduleFeeTypes_FeeSchedule
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_FeeScheduleHistory_FeeSchedule]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[FeeScheduleHistory] DROP CONSTRAINT FK_FeeScheduleHistory_FeeSchedule
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_FeeScheduleLevels_FeeSchedule]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[FeeScheduleLevels] DROP CONSTRAINT FK_FeeScheduleLevels_FeeSchedule
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_FeeScheduleFeeTypesHistory_FeeScheduleFeeTypes]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[FeeScheduleFeeTypesHistory] DROP CONSTRAINT FK_FeeScheduleFeeTypesHistory_FeeScheduleFeeTypes
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_IBFee_FeeScheduleFeeTypes]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[IBFee] DROP CONSTRAINT FK_IBFee_FeeScheduleFeeTypes
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_RTIBFee_FeeScheduleFeeTypes]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[RTIBFee] DROP CONSTRAINT FK_RTIBFee_FeeScheduleFeeTypes
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_FeeScheduleLevelsHistory_FeeScheduleLevels]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[FeeScheduleLevelsHistory] DROP CONSTRAINT FK_FeeScheduleLevelsHistory_FeeScheduleLevels
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_SecurityGroup_Group]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[SecurityGroup] DROP CONSTRAINT FK_SecurityGroup_Group
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_UserGroup_Group]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[UsersGroup] DROP CONSTRAINT FK_UserGroup_Group
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_IBFee_IB]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[IBFee] DROP CONSTRAINT FK_IBFee_IB
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_IBHistory_IB]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[IBHistory] DROP CONSTRAINT FK_IBHistory_IB
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_IBFeeHistory_IBFee]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[IBFeeHistory] DROP CONSTRAINT FK_IBFeeHistory_IBFee
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_UserReportsToCoordinator_Manager]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[UserReportsToCoordinator] DROP CONSTRAINT FK_UserReportsToCoordinator_Manager
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_UserReportsToManager_Manager]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[UserReportsToManager] DROP CONSTRAINT FK_UserReportsToManager_Manager
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_PackageHistory_Package]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[PackageHistory] DROP CONSTRAINT FK_PackageHistory_Package
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_PackageItem_Package]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[PackageItem] DROP CONSTRAINT FK_PackageItem_Package
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_PackageItemHistory_PackageItem]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[PackageItemHistory] DROP CONSTRAINT FK_PackageItemHistory_PackageItem
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_PolicyLimitsHistory_PolicyLimits]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[PolicyLimitsHistory] DROP CONSTRAINT FK_PolicyLimitsHistory_PolicyLimits
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_RTActivityLogHistory_RTActivityLog]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[RTActivityLogHistory] DROP CONSTRAINT FK_RTActivityLogHistory_RTActivityLog
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_RTActivityLogInfoHistory_RTActivityLogInfo]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[RTActivityLogInfoHistory] DROP CONSTRAINT FK_RTActivityLogInfoHistory_RTActivityLogInfo
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_PackageItem_RTAttachments]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[PackageItem] DROP CONSTRAINT FK_PackageItem_RTAttachments
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_RTAttachmentsHistory_RTAttachments]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[RTAttachmentsHistory] DROP CONSTRAINT FK_RTAttachmentsHistory_RTAttachments
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_RTChecksHistory_RTChecks]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[RTChecksHistory] DROP CONSTRAINT FK_RTChecksHistory_RTChecks
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_RTIndemnity_RTChecks]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[RTIndemnity] DROP CONSTRAINT FK_RTIndemnity_RTChecks
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_RTIBFee_RTIB]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[RTIBFee] DROP CONSTRAINT FK_RTIBFee_RTIB
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_RTIBHistory_RTIB]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[RTIBHistory] DROP CONSTRAINT FK_RTIBHistory_RTIB
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_RTIBFeeHistory_RTIBFee]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[RTIBFeeHistory] DROP CONSTRAINT FK_RTIBFeeHistory_RTIBFee
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_RTIndemnityHistory_RTIndemnity]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[RTIndemnityHistory] DROP CONSTRAINT FK_RTIndemnityHistory_RTIndemnity
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_RTPhotoLogHistory_RTPhotoLog]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[RTPhotoLogHistory] DROP CONSTRAINT FK_RTPhotoLogHistory_RTPhotoLog
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_RTPhotoLog_RTPhotoReport]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[RTPhotoLog] DROP CONSTRAINT FK_RTPhotoLog_RTPhotoReport
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_RTPhotoReportHistory_RTPhotoReport]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[RTPhotoReportHistory] DROP CONSTRAINT FK_RTPhotoReportHistory_RTPhotoReport
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_RTWSDiagramHistory_RTWSDiagram]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[RTWSDiagramHistory] DROP CONSTRAINT FK_RTWSDiagramHistory_RTWSDiagram
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_RegSettingHistory_RegSetting]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[RegSettingHistory] DROP CONSTRAINT FK_RegSettingHistory_RegSetting
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_SoftwarePackageRegSetting_RegSetting]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[SoftwarePackageRegSetting] DROP CONSTRAINT FK_SoftwarePackageRegSetting_RegSetting
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_SecurityGroup_Security]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[SecurityGroup] DROP CONSTRAINT FK_SecurityGroup_Security
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_SPSecurity_Security]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[SPSecurity] DROP CONSTRAINT FK_SPSecurity_Security
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_SASecirityPackage_SecurityArea1]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[SASecurityPackage] DROP CONSTRAINT FK_SASecirityPackage_SecurityArea1
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_SISecurityArea_SecurityArea]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[SISecurityArea] DROP CONSTRAINT FK_SISecurityArea_SecurityArea
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_SecurityArea_SecurityAreaType]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[SecurityArea] DROP CONSTRAINT FK_SecurityArea_SecurityAreaType
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_SecurityItems_SecurityItemType]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[SecurityItems] DROP CONSTRAINT FK_SecurityItems_SecurityItemType
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_SISecurityArea_SecurityItems]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[SISecurityArea] DROP CONSTRAINT FK_SISecurityArea_SecurityItems
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_CompanyUsers_SecurityLevel]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[CompanyUsers] DROP CONSTRAINT FK_CompanyUsers_SecurityLevel
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_SecurityLevelHistory_SecurityLevel]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[SecurityLevelHistory] DROP CONSTRAINT FK_SecurityLevelHistory_SecurityLevel
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Users_SecurityLevel]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[Users] DROP CONSTRAINT FK_Users_SecurityLevel
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_SASecirityPackage_SecurityPackage]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[SASecurityPackage] DROP CONSTRAINT FK_SASecirityPackage_SecurityPackage
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_SPSecurity_SecurityPackage]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[SPSecurity] DROP CONSTRAINT FK_SPSecurity_SecurityPackage
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_SoftwarePackageApplication_SoftwarePackage]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[SoftwarePackageApplication] DROP CONSTRAINT FK_SoftwarePackageApplication_SoftwarePackage
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_SoftwarePackageDocument_SoftwarePackage]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[SoftwarePackageDocument] DROP CONSTRAINT FK_SoftwarePackageDocument_SoftwarePackage
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_SoftwarePackageHistory_SoftwarePackage]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[SoftwarePackageHistory] DROP CONSTRAINT FK_SoftwarePackageHistory_SoftwarePackage
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_SoftwarePackageRegSetting_SoftwarePackage]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[SoftwarePackageRegSetting] DROP CONSTRAINT FK_SoftwarePackageRegSetting_SoftwarePackage
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_StateHistory_State]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[StateHistory] DROP CONSTRAINT FK_StateHistory_State
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Assignments_Status]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[Assignments] DROP CONSTRAINT FK_Assignments_Status
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_StatusHistory_Status]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[StatusHistory] DROP CONSTRAINT FK_StatusHistory_Status
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_XML_Trans_TransType]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[XML_Trans] DROP CONSTRAINT FK_XML_Trans_TransType
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Assignments_TypeOfLoss]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[Assignments] DROP CONSTRAINT FK_Assignments_TypeOfLoss
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_CompanyCat_TypeOfLoss]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[ClientCompanyCat] DROP CONSTRAINT FK_CompanyCat_TypeOfLoss
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_RTChecks_TypeOfLoss]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[RTChecks] DROP CONSTRAINT FK_RTChecks_TypeOfLoss
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_RTIndemnity_TypeOfLoss]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[RTIndemnity] DROP CONSTRAINT FK_RTIndemnity_TypeOfLoss
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_TypeOfLossHistory_TypeOfLoss]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[TypeOfLossHistory] DROP CONSTRAINT FK_TypeOfLossHistory_TypeOfLoss
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_UserProfileHistory_UserProfile]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[UserProfileHistory] DROP CONSTRAINT FK_UserProfileHistory_UserProfile
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_UsersUpdates_Users]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[AdjusterUsersUpdates] DROP CONSTRAINT FK_UsersUpdates_Users
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_CompanyUsers_Users]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[CompanyUsers] DROP CONSTRAINT FK_CompanyUsers_Users
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_ECSADJUsers_Users]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[ECSADJUsers] DROP CONSTRAINT FK_ECSADJUsers_Users
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_UserGroup_Users]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[UsersGroup] DROP CONSTRAINT FK_UserGroup_Users
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_UsersHistory_Users]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[UsersHistory] DROP CONSTRAINT FK_UsersHistory_Users
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[updAdjusterUsersSoftwareHistory]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[updAdjusterUsersSoftwareHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[insAdjusterUsersSoftwareHistory]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[insAdjusterUsersSoftwareHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[updAdjusterUsersUpdatesHistory]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[updAdjusterUsersUpdatesHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[insApplication]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[insApplication]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[updApplicationHistory]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[updApplicationHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[updAssignmentTypeHistory]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[updAssignmentTypeHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[updAssignmentsHistory]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[updAssignmentsHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[updBatchesHistory]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[updBatchesHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[insBillAssignment]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[insBillAssignment]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[insBillBillingCount]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[insBillBillingCount]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[updCATHistory]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[updCATHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[updClassOfLossHistory]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[updClassOfLossHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[updClassTypeHistory]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[updClassTypeHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[updClientCoAdjusterSpecHistory]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[updClientCoAdjusterSpecHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[insClientCoAdjusterSpec]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[insClientCoAdjusterSpec]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[insClientCompanyCat]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[insClientCompanyCat]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[updClientCompanyCatHistory]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[updClientCompanyCatHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[insClientCompanyCatSpec]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[insClientCompanyCatSpec]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[updClientCompanyCatSpecHistory]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[updClientCompanyCatSpecHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[updCompanyHistory]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[updCompanyHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[updCompanyUsersHistory]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[updCompanyUsersHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[insDocument]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[insDocument]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[updDocumentHistory]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[updDocumentHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[updFAQSHistory]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[updFAQSHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[updFtpArchive]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[updFtpArchive]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[updFeeScheduleHistory]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[updFeeScheduleHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[updFeeScheduleFeeTypesHistory]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[updFeeScheduleFeeTypesHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[updFeeScheduleLevelsHistory]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[updFeeScheduleLevelsHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[updIBHistory]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[updIBHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[updIB]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[updIB]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[updIBFeeHistory]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[updIBFeeHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[updIBStateFarm]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[updIBStateFarm]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[updPackageHistory]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[updPackageHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[updPackageItemHistory]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[updPackageItemHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[updPolicyLimitsHistory]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[updPolicyLimitsHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[updRTActivityLogHistory]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[updRTActivityLogHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[updRTActivityLogInfoHistory]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[updRTActivityLogInfoHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[updRTAttachmentsHistory]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[updRTAttachmentsHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[updRTChecksHistory]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[updRTChecksHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[updRTIBHistory]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[updRTIBHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[updRTIBFeeHistory]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[updRTIBFeeHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[updRTIndemnityHistory]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[updRTIndemnityHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[updRTPhotoLogHistory]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[updRTPhotoLogHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[updRTPhotoReportHistory]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[updRTPhotoReportHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[updRTWSDiagramHistory]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[updRTWSDiagramHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[insRegSetting]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[insRegSetting]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[updRegSettingHistory]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[updRegSettingHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[updSecurityLevelHistory]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[updSecurityLevelHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[updSoftwarePackageHistory]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[updSoftwarePackageHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[insSoftwarePackageApplication]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[insSoftwarePackageApplication]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[insSoftwarePackageDocument]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[insSoftwarePackageDocument]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[insSoftwarePackageRegSetting]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[insSoftwarePackageRegSetting]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[updStateHistory]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[updStateHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[updStatusHistory]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[updStatusHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[updTypeOfLossHistory]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[updTypeOfLossHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[updUserProfileHistory]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[updUserProfileHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[updUsersHistory]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[updUsersHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[insUsers]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[insUsers]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[insXML_Trans]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[insXML_Trans]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Accounting]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Accounting]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Adjuster]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Adjuster]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[AdjusterEvaluations]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[AdjusterEvaluations]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[AdjusterUsersSoftware]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[AdjusterUsersSoftware]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[AdjusterUsersSoftwareHistory]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[AdjusterUsersSoftwareHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[AdjusterUsersUpdates]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[AdjusterUsersUpdates]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[AdjusterUsersUpdatesHistory]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[AdjusterUsersUpdatesHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Admin]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Admin]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Application]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Application]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ApplicationHistory]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ApplicationHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[AssignmentType]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[AssignmentType]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[AssignmentTypeHistory]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[AssignmentTypeHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Assignments]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Assignments]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[AssignmentsHistory]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[AssignmentsHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Batches]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Batches]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[BatchesHistory]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[BatchesHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[BillAssignment]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[BillAssignment]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[BillBillingCount]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[BillBillingCount]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[BillingCount]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[BillingCount]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[CAT]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[CAT]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[CATHistory]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[CATHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ClassOfLoss]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ClassOfLoss]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ClassOfLossHistory]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ClassOfLossHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ClassType]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ClassType]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ClassTypeHistory]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ClassTypeHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Client]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Client]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ClientCoAdjusterSpec]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ClientCoAdjusterSpec]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ClientCoAdjusterSpecHistory]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ClientCoAdjusterSpecHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ClientCompanyCat]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ClientCompanyCat]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ClientCompanyCatHistory]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ClientCompanyCatHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ClientCompanyCatSpec]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ClientCompanyCatSpec]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ClientCompanyCatSpecHistory]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ClientCompanyCatSpecHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ClientCompanyUsersCat]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ClientCompanyUsersCat]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Company]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Company]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[CompanyHistory]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[CompanyHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[CompanyUsers]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[CompanyUsers]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[CompanyUsersHistory]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[CompanyUsersHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Coordinator]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Coordinator]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[DB_VERSION]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[DB_VERSION]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Dispatcher]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Dispatcher]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Document]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Document]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[DocumentHistory]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[DocumentHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ECSADJUsers]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ECSADJUsers]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Employee]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Employee]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FAQS]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[FAQS]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FAQSHistory]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[FAQSHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FTPLog]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[FTPLog]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FTPLogArchive]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[FTPLogArchive]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FeeSchedule]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[FeeSchedule]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FeeScheduleFeeTypes]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[FeeScheduleFeeTypes]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FeeScheduleFeeTypesHistory]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[FeeScheduleFeeTypesHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FeeScheduleHistory]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[FeeScheduleHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FeeScheduleLevels]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[FeeScheduleLevels]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FeeScheduleLevelsHistory]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[FeeScheduleLevelsHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Group]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Group]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[HTTPLog]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[HTTPLog]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[HTTPLogArchive]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[HTTPLogArchive]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[IB]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[IB]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[IBFee]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[IBFee]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[IBFeeHistory]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[IBFeeHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[IBHistory]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[IBHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[IBStateFarm]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[IBStateFarm]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Manager]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Manager]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[MiscReportParam]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[MiscReportParam]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[MiscReportParam01]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[MiscReportParam01]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[MiscReportParam02]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[MiscReportParam02]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[MiscReportParam03]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[MiscReportParam03]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[MiscReportParam04]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[MiscReportParam04]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[MiscReportParam05]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[MiscReportParam05]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[MiscReportParam06]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[MiscReportParam06]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[MiscReportParam07]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[MiscReportParam07]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[MiscReportParam08]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[MiscReportParam08]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[MiscReportParam09]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[MiscReportParam09]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[MiscReportParam10]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[MiscReportParam10]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[MiscReportParam11]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[MiscReportParam11]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[MiscReportParam12]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[MiscReportParam12]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[MiscReportParam13]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[MiscReportParam13]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[MiscReportParam14]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[MiscReportParam14]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[MiscReportParam15]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[MiscReportParam15]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[MiscReportParam16]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[MiscReportParam16]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[MiscReportParam17]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[MiscReportParam17]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[MiscReportParam18]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[MiscReportParam18]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[MiscReportParam19]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[MiscReportParam19]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[MiscReportParam20]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[MiscReportParam20]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[MiscReportParam21]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[MiscReportParam21]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[MiscReportParam22]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[MiscReportParam22]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[MiscReportParam23]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[MiscReportParam23]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[MiscReportParam24]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[MiscReportParam24]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[MiscReportParam25]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[MiscReportParam25]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[MiscReportParam26]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[MiscReportParam26]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[MiscReportParam27]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[MiscReportParam27]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[MiscReportParam28]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[MiscReportParam28]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[MiscReportParam29]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[MiscReportParam29]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[MiscReportParam30]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[MiscReportParam30]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Package]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Package]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[PackageHistory]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[PackageHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[PackageItem]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[PackageItem]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[PackageItemHistory]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[PackageItemHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[PolicyLimits]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[PolicyLimits]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[PolicyLimitsHistory]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[PolicyLimitsHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[RTActivityLog]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[RTActivityLog]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[RTActivityLogHistory]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[RTActivityLogHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[RTActivityLogInfo]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[RTActivityLogInfo]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[RTActivityLogInfoHistory]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[RTActivityLogInfoHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[RTAttachments]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[RTAttachments]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[RTAttachmentsHistory]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[RTAttachmentsHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[RTChecks]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[RTChecks]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[RTChecksHistory]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[RTChecksHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[RTIB]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[RTIB]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[RTIBFee]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[RTIBFee]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[RTIBFeeHistory]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[RTIBFeeHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[RTIBHistory]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[RTIBHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[RTIndemnity]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[RTIndemnity]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[RTIndemnityHistory]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[RTIndemnityHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[RTPhotoLog]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[RTPhotoLog]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[RTPhotoLogHistory]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[RTPhotoLogHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[RTPhotoReport]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[RTPhotoReport]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[RTPhotoReportHistory]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[RTPhotoReportHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[RTWSDiagram]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[RTWSDiagram]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[RTWSDiagramHistory]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[RTWSDiagramHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[RegSetting]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[RegSetting]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[RegSettingHistory]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[RegSettingHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SASecurityPackage]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[SASecurityPackage]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SF_AdjusterEvaluations]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[SF_AdjusterEvaluations]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SISecurityArea]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[SISecurityArea]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SPSecurity]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[SPSecurity]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Security]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Security]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SecurityArea]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[SecurityArea]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SecurityAreaType]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[SecurityAreaType]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SecurityGroup]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[SecurityGroup]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SecurityItemType]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[SecurityItemType]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SecurityItems]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[SecurityItems]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SecurityLevel]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[SecurityLevel]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SecurityLevelHistory]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[SecurityLevelHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SecurityPackage]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[SecurityPackage]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SoftwarePackage]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[SoftwarePackage]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SoftwarePackageApplication]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[SoftwarePackageApplication]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SoftwarePackageDocument]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[SoftwarePackageDocument]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SoftwarePackageHistory]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[SoftwarePackageHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SoftwarePackageRegSetting]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[SoftwarePackageRegSetting]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[State]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[State]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[StateHistory]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[StateHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Status]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Status]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[StatusHistory]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[StatusHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Temporary]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Temporary]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[TransType]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[TransType]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[TypeOfLoss]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[TypeOfLoss]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[TypeOfLossHistory]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[TypeOfLossHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[UserProfile]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[UserProfile]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[UserProfileHistory]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[UserProfileHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[UserReportsToCoordinator]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[UserReportsToCoordinator]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[UserReportsToManager]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[UserReportsToManager]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Users]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Users]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[UsersGroup]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[UsersGroup]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[UsersHistory]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[UsersHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Work_ListText]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Work_ListText]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[XML_Trans]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[XML_Trans]
GO

CREATE TABLE [dbo].[Accounting] (
	[CompanyID] [int] NOT NULL ,
	[UsersID] [int] NOT NULL ,
	[Active] [int] NOT NULL ,
	[DateLastUpdated] [datetime] NULL ,
	[UpdateByUserID] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Adjuster] (
	[CompanyID] [int] NOT NULL ,
	[UsersID] [int] NOT NULL ,
	[MaxOpenAssgn] [int] NOT NULL ,
	[HomeBaseZip] [int] NULL ,
	[MaxRangeFromHomeBaseZip] [int] NULL ,
	[Active] [bit] NOT NULL ,
	[DateLastUpdated] [datetime] NULL ,
	[UpdateByUserID] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[AdjusterEvaluations] (
	[EvaluationID] [int] IDENTITY (1, 1) NOT NULL ,
	[EvaluationDate] [datetime] NULL ,
	[CatCode] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ArrivalDate] [datetime] NULL ,
	[DepartureDate] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[CatOfficeLocation] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[CompanyName] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[AdjUID] [int] NULL ,
	[AdjFirstName] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[AdjLastName] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[UserID] [int] NULL ,
	[NumFilesAssigned] [int] NULL ,
	[NumFilesInspected] [int] NULL ,
	[NumFilesClosed] [int] NULL ,
	[NumFilesAverageEstimate] [money] NULL ,
	[DaysOnTheStorm] [int] NULL ,
	[AverageClosingsPerDay] [float] NULL ,
	[EstimatingPlatform] [nvarchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[EstimatingSkills] [int] NULL ,
	[ComputerSkills] [int] NULL ,
	[AccuracyofScope] [int] NULL ,
	[PrioritizationofAssignments] [int] NULL ,
	[Productivity] [int] NULL ,
	[Professionalism] [int] NULL ,
	[TelephoneFollowup] [int] NULL ,
	[SubmitsAccurateBilling] [int] NULL ,
	[FollowSupervisoryDirection] [int] NULL ,
	[FileSevCompHeavy] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[FileSevCompModerate] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[FileSevCompLight] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[GeneralComments] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[CatRT] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[CatSupervisor] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[WindHail] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Hurricane] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Flood] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Other] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Earthquake] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[FreezeStorm] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Commercial] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[OtherLossesHandled] [nvarchar] (500) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[txtEstimatingSkills] [varchar] (500) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[txtComputerSkills] [varchar] (500) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[txtAccuracyofScope] [nvarchar] (500) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[txtPrioritizationofAssignments] [varchar] (500) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[txtProductivity] [varchar] (500) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[txtProfessionalism] [varchar] (500) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[txtTelephoneFollowup] [varchar] (500) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[txtSubmitsAccurateBilling] [varchar] (500) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[txtFollowSupervisoryDirection] [varchar] (500) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[EnteredBy] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DateEntered] [datetime] NULL ,
	[ModifiedBy] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ModifiedDate] [datetime] NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[AdjusterUsersSoftware] (
	[UsersID] [int] NOT NULL ,
	[VersionInfo] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[LicenseDaysLeft] [smallint] NULL ,
	[ResetLicense] [bit] NULL ,
	[IBPrefix] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ResetIBPrefix] [bit] NULL ,
	[SingleFileSendAuthority] [bit] NOT NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[AdjusterUsersSoftwareHistory] (
	[AdjusterUsersSoftwareHistoryID] [int] IDENTITY (1, 1) NOT NULL ,
	[UsersID] [int] NOT NULL ,
	[VersionInfo] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[LicenseDaysLeft] [smallint] NULL ,
	[ResetLicense] [bit] NULL ,
	[IBPrefix] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ResetIBPrefix] [bit] NULL ,
	[SingleFileSendAuthority] [bit] NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[AdjusterUsersUpdates] (
	[UsersID] [int] NOT NULL ,
	[FirstName] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[LastName] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[SSN] [int] NULL ,
	[Email] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ContactPhone] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[EmergencyPhone] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Address] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[City] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[State] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Zip] [int] NULL ,
	[ZIP4] [int] NULL ,
	[OtherPostCode] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[AdjusterUsersUpdatesHistory] (
	[AdjusterUsersUpdatesHistoryID] [int] IDENTITY (1, 1) NOT NULL ,
	[UsersID] [int] NULL ,
	[FirstName] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[LastName] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[SSN] [int] NULL ,
	[Email] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ContactPhone] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[EmergencyPhone] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Address] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[City] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[State] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Zip] [int] NULL ,
	[ZIP4] [int] NULL ,
	[OtherPostCode] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Admin] (
	[CompanyID] [int] NOT NULL ,
	[UsersID] [int] NOT NULL ,
	[Active] [bit] NOT NULL ,
	[DateLastUpdated] [datetime] NULL ,
	[UpdateByUserID] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Application] (
	[ApplicationID] [int] IDENTITY (1, 1) NOT NULL ,
	[AppNameBase] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[AppName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Description] [varchar] (500) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Version] [int] NOT NULL ,
	[MajorVS] [int] NOT NULL ,
	[MinorVS] [int] NOT NULL ,
	[RevisionVS] [int] NOT NULL ,
	[SPVersionBase] [int] NOT NULL ,
	[SPVersion] [int] NOT NULL ,
	[VersionDate] [datetime] NOT NULL ,
	[ProjectName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ClassName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[SectionLevel01] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[SectionLevel02] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[SectionLevel03] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[SectionLevel04] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[SectionLevel05] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[InstallFileLocation] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[SPName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[SELF_REG] [bit] NOT NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ApplicationHistory] (
	[ApplicationID] [int] NOT NULL ,
	[AppNameBase] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[AppName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Description] [varchar] (500) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Version] [int] NOT NULL ,
	[MajorVS] [int] NOT NULL ,
	[MinorVS] [int] NOT NULL ,
	[RevisionVS] [int] NOT NULL ,
	[SPVersionBase] [int] NULL ,
	[SPVersion] [int] NULL ,
	[VersionDate] [datetime] NOT NULL ,
	[ProjectName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ClassName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[SectionLevel01] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[SectionLevel02] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[SectionLevel03] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[SectionLevel04] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[SectionLevel05] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[InstallFileLocation] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[SPName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[SELF_REG] [bit] NOT NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[AssignmentType] (
	[AssignmentTypeID] [int] IDENTITY (1, 1) NOT NULL ,
	[Type] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Description] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[AdminComments] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[AssignmentTypeHistory] (
	[AssignmentTypeHistoryID] [int] IDENTITY (1, 1) NOT NULL ,
	[AssignmentTypeID] [int] NOT NULL ,
	[Type] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Description] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[AdminComments] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Assignments] (
	[AssignmentsID] [int] IDENTITY (20000, 1) NOT NULL ,
	[ID] [int] NULL ,
	[AssignmentTypeID] [int] NOT NULL ,
	[ClientCompanyCatSpecID] [int] NOT NULL ,
	[AdjusterSpecID] [int] NOT NULL ,
	[AdjusterSpecIDDisplay] [int] NULL ,
	[SPVersion] [int] NOT NULL ,
	[IBNUM] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[CLIENTNUM] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[PolicyNo] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[PolicyDescription] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Insured] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[MailingAddress] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[MAStreet] [varchar] (150) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[MACity] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[MAState] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[MAZIP] [int] NULL ,
	[MAZIP4] [int] NULL ,
	[MAOtherPostCode] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[HomePhone] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[BusinessPhone] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[PropertyAddress] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[PAStreet] [varchar] (150) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[PACity] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[PAState] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[PAZIP] [int] NULL ,
	[PAZIP4] [int] NULL ,
	[PAOtherPostCode] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[MortgageeName] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[AgentNo] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ReportedBy] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ReportedByPhone] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Deductible] [money] NULL ,
	[AppDedClassTypeIDOrder] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[LRFormat] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[LossReport] [ntext] COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[LRPrintedDate] [datetime] NULL ,
	[DownLoadLossReport] [bit] NOT NULL ,
	[UploadLossReport] [bit] NOT NULL ,
	[StatusID] [int] NOT NULL ,
	[TypeOfLossID] [int] NULL ,
	[XactTypeOfLoss] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[SentToXact] [bit] NOT NULL ,
	[LossDate] [datetime] NULL ,
	[AssignedDate] [datetime] NULL ,
	[ReceivedDate] [datetime] NULL ,
	[ContactDate] [datetime] NULL ,
	[InspectedDate] [datetime] NULL ,
	[CloseDate] [datetime] NULL ,
	[Reassigned] [bit] NULL ,
	[DateReassigned] [datetime] NULL ,
	[RAAdjusterSpecID] [int] NULL ,
	[IsLocked] [bit] NOT NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DownLoadMe] [bit] NOT NULL ,
	[UpLoadMe] [bit] NOT NULL ,
	[DownLoadAll] [bit] NOT NULL ,
	[UpLoadAll] [bit] NOT NULL ,
	[AdminComments] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[MiscDelimSettings] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[AssignmentsHistory] (
	[AssignmentsHistoryID] [int] IDENTITY (1, 1) NOT NULL ,
	[AssignmentsID] [int] NOT NULL ,
	[ID] [int] NULL ,
	[AssignmentTypeID] [int] NOT NULL ,
	[ClientCompanyCatSpecID] [int] NOT NULL ,
	[AdjusterSpecID] [int] NOT NULL ,
	[AdjusterSpecIDDisplay] [int] NULL ,
	[SPVersion] [int] NULL ,
	[IBNUM] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[CLIENTNUM] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[PolicyNo] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[PolicyDescription] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Insured] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[MailingAddress] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[MAStreet] [varchar] (150) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[MACity] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[MAState] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[MAZIP] [int] NULL ,
	[MAZIP4] [int] NULL ,
	[MAOtherPostCode] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[HomePhone] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[BusinessPhone] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[PropertyAddress] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[PAStreet] [varchar] (150) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[PACity] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[PAState] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[PAZIP] [int] NULL ,
	[PAZIP4] [int] NULL ,
	[PAOtherPostCode] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[MortgageeName] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[AgentNo] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ReportedBy] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ReportedByPhone] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Deductible] [money] NULL ,
	[AppDedClassTypeIDOrder] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[LRFormat] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[LossReport] [ntext] COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[LRPrintedDate] [datetime] NULL ,
	[DownLoadLossReport] [bit] NOT NULL ,
	[UploadLossReport] [bit] NOT NULL ,
	[StatusID] [int] NOT NULL ,
	[TypeOfLossID] [int] NULL ,
	[XactTypeOfLoss] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[SentToXact] [bit] NOT NULL ,
	[LossDate] [datetime] NULL ,
	[AssignedDate] [datetime] NULL ,
	[ReceivedDate] [datetime] NULL ,
	[ContactDate] [datetime] NULL ,
	[InspectedDate] [datetime] NULL ,
	[CloseDate] [datetime] NULL ,
	[Reassigned] [bit] NULL ,
	[DateReassigned] [datetime] NULL ,
	[RAAdjusterSpecID] [int] NULL ,
	[IsLocked] [bit] NOT NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DownLoadMe] [bit] NOT NULL ,
	[UpLoadMe] [bit] NOT NULL ,
	[DownLoadAll] [bit] NOT NULL ,
	[UpLoadAll] [bit] NOT NULL ,
	[AdminComments] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[MiscDelimSettings] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[Batches] (
	[BatchesID] [int] IDENTITY (1, 1) NOT NULL ,
	[AssignmentsID] [int] NULL ,
	[ClientCompanyCatSpecID] [int] NOT NULL ,
	[ssn] [numeric](9, 0) NULL ,
	[ibnumber] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[date] [datetime] NULL ,
	[EnteredDate] [datetime] NULL ,
	[adj_name] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[adjuster_n] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[claimnumber] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[insuredname] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[loss_loc] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[losscity] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[lossstate] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[dateofloss] [datetime] NULL ,
	[dateclosed] [datetime] NULL ,
	[grossloss] [decimal](20, 5) NULL ,
	[totalservice] [decimal](20, 5) NULL ,
	[administrative] [decimal](20, 5) NULL ,
	[misccharge] [decimal](20, 5) NULL ,
	[taxestotal] [decimal](20, 5) NULL ,
	[totalfee] [decimal](20, 5) NULL ,
	[catsite] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Void] [bit] NOT NULL ,
	[billingdup] [bit] NULL ,
	[ecupdated] [bit] NULL ,
	[copied] [int] NULL ,
	[duplicate] [bit] NULL ,
	[Comments] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Reassigned] [int] NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL ,
	[BillAssignmentID] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[BatchesHistory] (
	[BatchesHistoryID] [int] IDENTITY (1, 1) NOT NULL ,
	[BatchesID] [int] NOT NULL ,
	[AssignmentsID] [int] NULL ,
	[ClientCompanyCatSpecID] [int] NOT NULL ,
	[ssn] [numeric](9, 0) NULL ,
	[ibnumber] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[date] [datetime] NULL ,
	[EnteredDate] [datetime] NULL ,
	[adj_name] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[adjuster_n] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[claimnumber] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[insuredname] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[loss_loc] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[losscity] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[lossstate] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[dateofloss] [datetime] NULL ,
	[dateclosed] [datetime] NULL ,
	[grossloss] [decimal](20, 5) NULL ,
	[totalservice] [decimal](20, 5) NULL ,
	[administrative] [decimal](20, 5) NULL ,
	[misccharge] [decimal](20, 5) NULL ,
	[taxestotal] [decimal](20, 5) NULL ,
	[totalfee] [decimal](20, 5) NULL ,
	[catsite] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Void] [bit] NOT NULL ,
	[billingdup] [bit] NULL ,
	[ecupdated] [bit] NULL ,
	[copied] [int] NULL ,
	[duplicate] [bit] NULL ,
	[Comments] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Reassigned] [int] NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL ,
	[BillAssignmentID] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[BillAssignment] (
	[BillAssignmentID] [int] IDENTITY (20000, 1) NOT NULL ,
	[AssignmentTypeID] [int] NOT NULL ,
	[ClientCompanyCatSpecID] [int] NOT NULL ,
	[AdjusterSpecID] [int] NOT NULL ,
	[IBNUM] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[CLIENTNUM] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[PolicyNo] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Insured] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[LossLoc1] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[LossLoc2] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[LossLocCity] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[LossLocState] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[LossLocZipcode] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[LossDate] [datetime] NOT NULL ,
	[CloseDate] [datetime] NULL ,
	[IsLocked] [bit] NOT NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[AdminComments] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[MiscDelimSettings] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[BillBillingCount] (
	[BillBillingCountID] [int] IDENTITY (20000, 1) NOT NULL ,
	[BillAssignmentID] [int] NOT NULL ,
	[Rebill] [int] NOT NULL ,
	[Supplement] [int] NOT NULL ,
	[AdminComments] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[BillingCount] (
	[BillingCountID] [int] IDENTITY (20000, 1) NOT NULL ,
	[AssignmentsID] [int] NOT NULL ,
	[ID] [int] NULL ,
	[IDAssignments] [int] NULL ,
	[Rebill] [int] NOT NULL ,
	[Supplement] [int] NOT NULL ,
	[DownLoadMe] [bit] NOT NULL ,
	[UpLoadMe] [bit] NOT NULL ,
	[AdminComments] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[CAT] (
	[CATID] [int] IDENTITY (1, 1) NOT NULL ,
	[CompanyID] [int] NOT NULL ,
	[AssignmentTypeID] [int] NOT NULL ,
	[Name] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Description] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ActiveDate] [datetime] NULL ,
	[InactiveDate] [datetime] NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[CATHistory] (
	[CATHistoryID] [int] IDENTITY (1, 1) NOT NULL ,
	[CATID] [int] NOT NULL ,
	[CompanyID] [int] NULL ,
	[AssignmentTypeID] [int] NULL ,
	[Name] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Description] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ActiveDate] [datetime] NULL ,
	[InactiveDate] [datetime] NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ClassOfLoss] (
	[ClassOfLossID] [int] IDENTITY (1, 1) NOT NULL ,
	[ClientCompanyID] [int] NOT NULL ,
	[ClassTypeID] [int] NOT NULL ,
	[IsSubSetOFClassOfLossID] [int] NULL ,
	[Code] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Description] [varchar] (150) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ClassOfLossHistory] (
	[ClassOfLossHistoryID] [int] IDENTITY (1, 1) NOT NULL ,
	[ClassOfLossID] [int] NOT NULL ,
	[ClientCompanyID] [int] NOT NULL ,
	[ClassTypeID] [int] NOT NULL ,
	[IsSubSetOFClassOfLossID] [int] NULL ,
	[Code] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Description] [varchar] (150) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ClassType] (
	[ClassTypeID] [int] IDENTITY (1, 1) NOT NULL ,
	[Class] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Description] [varchar] (150) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ClassTypeHistory] (
	[ClassTypeHistoryID] [int] IDENTITY (1, 1) NOT NULL ,
	[ClassTypeID] [int] NOT NULL ,
	[Class] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Description] [varchar] (150) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Client] (
	[CompanyID] [int] NOT NULL ,
	[UsersID] [int] NOT NULL ,
	[Active] [bit] NOT NULL ,
	[DateLastUpdated] [datetime] NULL ,
	[UpdateByUserID] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ClientCoAdjusterSpec] (
	[ClientCoAdjusterSpecID] [int] IDENTITY (1, 1) NOT NULL ,
	[ClientCompanyID] [int] NOT NULL ,
	[UsersID] [int] NOT NULL ,
	[ACID] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ACIDDescription] [varchar] (150) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Comments] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ClientCompanyCatSpecID] [int] NULL ,
	[ZipCode] [int] NULL ,
	[ZipCodeCount] [int] NULL ,
	[ActiveDate] [datetime] NOT NULL ,
	[InactiveDate] [datetime] NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ClientCoAdjusterSpecHistory] (
	[AdjusterSpecHistoryID] [int] IDENTITY (1, 1) NOT NULL ,
	[ClientCoAdjusterSpecID] [int] NOT NULL ,
	[ClientCompanyID] [int] NOT NULL ,
	[UsersID] [int] NOT NULL ,
	[ACID] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ACIDDescription] [varchar] (150) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Comments] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ClientCompanyCatSpecID] [int] NULL ,
	[ZipCode] [int] NULL ,
	[ZipCodeCount] [int] NULL ,
	[ActiveDate] [datetime] NOT NULL ,
	[InactiveDate] [datetime] NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ClientCompanyCat] (
	[ClientCompanyID] [int] NOT NULL ,
	[CATID] [int] NOT NULL ,
	[BillingCode] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[TypeOfLossID] [int] NULL ,
	[FeeScheduleID] [int] NULL ,
	[SiteAddress] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[SACity] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[SAState] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[SAZip] [int] NULL ,
	[SAZip4] [int] NULL ,
	[SAOtherPostCode] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ActiveDate] [datetime] NULL ,
	[InactiveDate] [datetime] NULL ,
	[AssignByZipDefault] [bit] NOT NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ClientCompanyCatHistory] (
	[ClientCompanyCatHistoryID] [int] IDENTITY (1, 1) NOT NULL ,
	[ClientCompanyID] [int] NOT NULL ,
	[CATID] [int] NOT NULL ,
	[BillingCode] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[TypeOfLossID] [int] NULL ,
	[FeeScheduleID] [int] NULL ,
	[SiteAddress] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[SACity] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[SAState] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[SAZip] [int] NULL ,
	[SAZip4] [int] NULL ,
	[SAOtherPostCode] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ActiveDate] [datetime] NULL ,
	[InactiveDate] [datetime] NULL ,
	[AssignByZipDefault] [bit] NOT NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ClientCompanyCatSpec] (
	[ClientCompanyCatSpecID] [int] IDENTITY (1, 1) NOT NULL ,
	[ClientCompanyID] [int] NOT NULL ,
	[CATID] [int] NOT NULL ,
	[CatCode] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Comments] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ActiveDate] [datetime] NOT NULL ,
	[InactiveDate] [datetime] NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ClientCompanyCatSpecHistory] (
	[ClientCompanyCatSpecHistoryID] [int] IDENTITY (1, 1) NOT NULL ,
	[ClientCompanyCatSpecID] [int] NOT NULL ,
	[ClientCompanyID] [int] NOT NULL ,
	[CATID] [int] NOT NULL ,
	[CatCode] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Comments] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ActiveDate] [datetime] NOT NULL ,
	[InactiveDate] [datetime] NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ClientCompanyUsersCat] (
	[ClientCompanyID] [int] NOT NULL ,
	[CATID] [int] NOT NULL ,
	[UsersID] [int] NOT NULL ,
	[Active] [bit] NOT NULL ,
	[DateLastUpdated] [datetime] NULL ,
	[UpdateByUserID] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Company] (
	[CompanyID] [int] IDENTITY (1, 1) NOT NULL ,
	[Name] [varchar] (30) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[DBName] [varchar] (30) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Code] [varchar] (3) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[CarrierPrefix] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Comments] [varchar] (150) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[IsClientOf] [int] NULL ,
	[EnableSingleFile] [bit] NOT NULL ,
	[SingleFileEmail] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[PDFJpegQuality] [int] NOT NULL ,
	[Active] [bit] NOT NULL ,
	[LogoImageName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[CompanyHistory] (
	[CompanyHistoryID] [int] IDENTITY (1, 1) NOT NULL ,
	[CompanyID] [int] NOT NULL ,
	[Name] [varchar] (30) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[DBName] [varchar] (30) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Code] [varchar] (3) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[CarrierPrefix] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Comments] [varchar] (150) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[IsClientOf] [int] NULL ,
	[EnableSingleFile] [bit] NOT NULL ,
	[SingleFileEmail] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[PDFJpegQuality] [int] NOT NULL ,
	[Active] [bit] NOT NULL ,
	[LogoImageName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[CompanyUsers] (
	[CompanyID] [int] NOT NULL ,
	[UsersID] [int] NOT NULL ,
	[Active] [bit] NOT NULL ,
	[ActiveDate] [datetime] NOT NULL ,
	[InactiveDate] [datetime] NULL ,
	[SecurityLevel] [int] NOT NULL ,
	[Comments] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Flag] [bit] NOT NULL ,
	[AssignmentTypeIDList] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[CompanyUsersHistory] (
	[CompanyUsersHistoryID] [int] IDENTITY (1, 1) NOT NULL ,
	[CompanyID] [int] NOT NULL ,
	[UsersID] [int] NOT NULL ,
	[Active] [bit] NOT NULL ,
	[ActiveDate] [datetime] NOT NULL ,
	[InactiveDate] [datetime] NULL ,
	[SecurityLevel] [int] NOT NULL ,
	[Comments] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Flag] [bit] NOT NULL ,
	[AssignmentTypeIDList] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Coordinator] (
	[CompanyID] [int] NOT NULL ,
	[UsersID] [int] NOT NULL ,
	[Active] [bit] NOT NULL ,
	[DateLastUpdated] [datetime] NULL ,
	[UpdateByUserID] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[DB_VERSION] (
	[Version] [int] NOT NULL ,
	[Comments] [varchar] (500) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[InstallFileLocation] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[SPName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[MainUtilInstallFileLocation] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[MainUtilSPName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[MainARVInstallFileLocation] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[MainARVSPName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[MainEXEInstallFileLocation] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[MainEXESPName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[MainFTPEXEInstallFileLocation] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[MainFTPEXESPName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Dispatcher] (
	[CompanyID] [int] NOT NULL ,
	[UsersID] [int] NOT NULL ,
	[Active] [bit] NOT NULL ,
	[DateLastUpdated] [datetime] NULL ,
	[UpdateByUserID] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Document] (
	[DocumentID] [int] IDENTITY (1, 1) NOT NULL ,
	[DocNameBase] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[DocName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Description] [varchar] (500) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Version] [int] NOT NULL ,
	[SPVersionBase] [int] NOT NULL ,
	[SPVersion] [int] NOT NULL ,
	[VersionDate] [datetime] NOT NULL ,
	[SectionLevel01] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[SectionLevel02] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[SectionLevel03] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[SectionLevel04] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[SectionLevel05] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[InstallFileLocation] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[SPName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[DocumentHistory] (
	[DocumentID] [int] NOT NULL ,
	[DocNameBase] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DocName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Description] [varchar] (500) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Version] [int] NOT NULL ,
	[SPVersionBase] [int] NULL ,
	[SPVersion] [int] NULL ,
	[VersionDate] [datetime] NOT NULL ,
	[SectionLevel01] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[SectionLevel02] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[SectionLevel03] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[SectionLevel04] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[SectionLevel05] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[InstallFileLocation] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[SPName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ECSADJUsers] (
	[UsersID] [int] NOT NULL ,
	[AdjUID] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Employee] (
	[CompanyID] [int] NOT NULL ,
	[UsersID] [int] NOT NULL ,
	[Active] [bit] NOT NULL ,
	[DateLastUpdated] [datetime] NULL ,
	[UpdateByUserID] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[FAQS] (
	[FAQSID] [int] IDENTITY (1, 1) NOT NULL ,
	[Question] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Answer] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[FAQSHistory] (
	[FAQSHistoryID] [int] IDENTITY (1, 1) NOT NULL ,
	[FAQSID] [int] NOT NULL ,
	[Question] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Answer] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[FTPLog] (
	[FTPLogID] [int] IDENTITY (1, 1) NOT NULL ,
	[ClientHost] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[username] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[LogTime] [datetime] NOT NULL ,
	[service] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[machine] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[serverip] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[processingtime] [int] NOT NULL ,
	[bytesrecvd] [int] NOT NULL ,
	[bytessent] [int] NOT NULL ,
	[servicestatus] [int] NOT NULL ,
	[win32status] [int] NOT NULL ,
	[operation] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[target] [varchar] (500) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[parameters] [varchar] (500) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[FTPLogArchive] (
	[FTPLogArchiveID] [int] IDENTITY (1, 1) NOT NULL ,
	[FTPLogID] [int] NOT NULL ,
	[ClientHost] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[username] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[LogTime] [datetime] NOT NULL ,
	[service] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[machine] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[serverip] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[processingtime] [int] NOT NULL ,
	[bytesrecvd] [int] NOT NULL ,
	[bytessent] [int] NOT NULL ,
	[servicestatus] [int] NOT NULL ,
	[win32status] [int] NOT NULL ,
	[operation] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[target] [varchar] (500) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[parameters] [varchar] (500) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[FeeSchedule] (
	[FeeScheduleID] [int] IDENTITY (1, 1) NOT NULL ,
	[ClientCompanyID] [int] NOT NULL ,
	[ScheduleName] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Description] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[NumOfLevels] [int] NULL ,
	[NumOfFeeTypes] [int] NULL ,
	[FeeServiceHourlyRate] [money] NOT NULL ,
	[TaxPercent] [decimal](18, 4) NOT NULL ,
	[InitialOptions] [varchar] (3000) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Options] [varchar] (3000) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[DefaultAppDedClassTypeIDOrder] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[FeeScheduleFeeTypes] (
	[FeeScheduleFeeTypesID] [int] IDENTITY (1, 1) NOT NULL ,
	[FeeScheduleID] [int] NOT NULL ,
	[TypeNum] [int] NULL ,
	[Name] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Description] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[FeeAmount] [money] NOT NULL ,
	[IsExpense] [bit] NULL ,
	[MaxNumberOfItems] [int] NULL ,
	[MaxFeeAmount] [money] NULL ,
	[IsMiscAmount] [bit] NULL ,
	[UseFormula] [bit] NULL ,
	[VBFormula] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[FeeScheduleFeeTypesHistory] (
	[FeeScheduleFeeTypesHistoryID] [int] IDENTITY (1, 1) NOT NULL ,
	[FeeScheduleFeeTypesID] [int] NOT NULL ,
	[FeeScheduleID] [int] NOT NULL ,
	[TypeNum] [int] NULL ,
	[Name] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Description] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[FeeAmount] [money] NOT NULL ,
	[IsExpense] [bit] NULL ,
	[MaxNumberOfItems] [int] NULL ,
	[MaxFeeAmount] [money] NULL ,
	[IsMiscAmount] [bit] NULL ,
	[UseFormula] [bit] NULL ,
	[VBFormula] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[FeeScheduleHistory] (
	[FeeScheduleHistoryID] [int] IDENTITY (1, 1) NOT NULL ,
	[FeeScheduleID] [int] NOT NULL ,
	[ClientCompanyID] [int] NOT NULL ,
	[ScheduleName] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Description] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[NumOfLevels] [int] NULL ,
	[NumOfFeeTypes] [int] NULL ,
	[FeeServiceHourlyRate] [money] NOT NULL ,
	[TaxPercent] [decimal](18, 4) NOT NULL ,
	[InitialOptions] [varchar] (3000) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Options] [varchar] (3000) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[DefaultAppDedClassTypeIDOrder] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[FeeScheduleLevels] (
	[FeeScheduleLevelsID] [int] IDENTITY (1, 1) NOT NULL ,
	[FeeScheduleID] [int] NOT NULL ,
	[LevelNum] [int] NULL ,
	[LevelMax] [money] NOT NULL ,
	[LevelPctApp] [decimal](18, 4) NOT NULL ,
	[LevelMin] [money] NOT NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[FeeScheduleLevelsHistory] (
	[FeeScheduleLevelsHistoryID] [int] IDENTITY (1, 1) NOT NULL ,
	[FeeScheduleLevelsID] [int] NOT NULL ,
	[FeeScheduleID] [int] NOT NULL ,
	[LevelNum] [int] NULL ,
	[LevelMax] [money] NOT NULL ,
	[LevelPctApp] [decimal](18, 4) NOT NULL ,
	[LevelMin] [money] NOT NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Group] (
	[GroupID] [int] IDENTITY (1, 1) NOT NULL ,
	[Name] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Description] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[HTTPLog] (
	[HTTPLogID] [int] IDENTITY (1, 1) NOT NULL ,
	[ClientHost] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[username] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[LogTime] [datetime] NOT NULL ,
	[service] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[machine] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[serverip] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[processingtime] [int] NOT NULL ,
	[bytesrecvd] [int] NOT NULL ,
	[bytessent] [int] NOT NULL ,
	[servicestatus] [int] NOT NULL ,
	[win32status] [int] NOT NULL ,
	[operation] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[target] [varchar] (500) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[parameters] [varchar] (500) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[HTTPLogArchive] (
	[HTTPLogArchiveID] [int] IDENTITY (1, 1) NOT NULL ,
	[HTTPLogID] [int] NOT NULL ,
	[ClientHost] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[username] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[LogTime] [datetime] NOT NULL ,
	[service] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[machine] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[serverip] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[processingtime] [int] NOT NULL ,
	[bytesrecvd] [int] NOT NULL ,
	[bytessent] [int] NOT NULL ,
	[servicestatus] [int] NOT NULL ,
	[win32status] [int] NOT NULL ,
	[operation] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[target] [varchar] (500) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[parameters] [varchar] (500) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[IB] (
	[IBID] [int] IDENTITY (20000, 1) NOT NULL ,
	[AssignmentsID] [int] NOT NULL ,
	[BillingCountID] [int] NOT NULL ,
	[ID] [int] NULL ,
	[IDAssignments] [int] NULL ,
	[IDBillingCount] [int] NOT NULL ,
	[IB00_lssn] [numeric](9, 0) NULL ,
	[IB01_sSubToCarrier] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[IB02_sIBNumber] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[IB05_sLocation] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[IB05a_sState] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[IB06_dtDateClosed] [datetime] NULL ,
	[IB07_sAdjusterName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[IB09_sSALN] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[IB09a_sPolicyNo] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[IB10_sInsuredName] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[IB11_sLossLocation] [varchar] (254) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[IB12_dtDateOfLoss] [datetime] NULL ,
	[IB13_cGrossLoss] [money] NULL ,
	[IB14_cDepreciation] [money] NULL ,
	[IB14a_sSupplement] [int] NOT NULL ,
	[IB14b_sRebilled] [int] NOT NULL ,
	[IB15_cDeductible] [money] NULL ,
	[IB15a_cLessExcessLimits] [money] NULL ,
	[IB15b_sExcessLimDesc] [varchar] (254) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[IB15c_cLessMiscellaneous] [money] NULL ,
	[IB15d_cMiscellaneousDesc] [varchar] (254) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[IB16_cNetClaim] [money] NULL ,
	[IB17_cServiceFee] [money] NULL ,
	[IB17a_cMiscServiceFee] [money] NULL ,
	[IB18_sServiceFeeComment] [varchar] (254) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[IB18a_sMiscServiceFeeComment] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[IB25_cServiceFeeSubTotal] [money] NULL ,
	[IB29a_sMiscExpenseFeeComment] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[IB29b_cMiscExpenseFee] [money] NULL ,
	[IB30_cTotalExpenses] [money] NULL ,
	[IB31_dTaxPercent] [numeric](8, 3) NULL ,
	[IB32_cTaxAmount] [money] NULL ,
	[IB33_cTotalAdjustingFee] [money] NULL ,
	[IB33a_sAccountCode] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[FeeScheduleID] [int] NULL ,
	[Void] [bit] NOT NULL ,
	[FeeByTime] [bit] NULL ,
	[UseActivityTime] [bit] NULL ,
	[DownLoadMe] [bit] NOT NULL ,
	[UpLoadMe] [bit] NOT NULL ,
	[Comments] [varchar] (150) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[AdminComments] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[IBFee] (
	[IBFeeID] [int] IDENTITY (20000, 1) NOT NULL ,
	[AssignmentsID] [int] NOT NULL ,
	[IBID] [int] NOT NULL ,
	[ID] [int] NULL ,
	[IDAssignments] [int] NULL ,
	[IDIB] [int] NULL ,
	[FeeScheduleFeeTypesID] [int] NOT NULL ,
	[NumberOfItems] [int] NOT NULL ,
	[Amount] [money] NOT NULL ,
	[Comment] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[DownLoadMe] [bit] NOT NULL ,
	[UpLoadMe] [bit] NOT NULL ,
	[AdminComments] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[IBFeeHistory] (
	[IBFeeHistoryID] [int] IDENTITY (20000, 1) NOT NULL ,
	[IBFeeID] [int] NOT NULL ,
	[AssignmentsID] [int] NULL ,
	[IBID] [int] NULL ,
	[ID] [int] NULL ,
	[IDAssignments] [int] NULL ,
	[IDIB] [int] NULL ,
	[FeeScheduleFeeTypesID] [int] NULL ,
	[NumberOfItems] [int] NULL ,
	[Amount] [money] NULL ,
	[Comment] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DownLoadMe] [bit] NULL ,
	[UpLoadMe] [bit] NULL ,
	[AdminComments] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DateLastUpdated] [datetime] NULL ,
	[UpdateByUserID] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[IBHistory] (
	[IBHistoryID] [int] IDENTITY (1, 1) NOT NULL ,
	[IBID] [int] NOT NULL ,
	[AssignmentsID] [int] NOT NULL ,
	[BillingCountID] [int] NOT NULL ,
	[ID] [int] NULL ,
	[IDAssignments] [int] NULL ,
	[IDBillingCount] [int] NOT NULL ,
	[IB00_lssn] [numeric](9, 0) NULL ,
	[IB01_sSubToCarrier] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[IB02_sIBNumber] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[IB05_sLocation] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[IB05a_sState] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[IB06_dtDateClosed] [datetime] NULL ,
	[IB07_sAdjusterName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[IB09_sSALN] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[IB09a_sPolicyNo] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[IB10_sInsuredName] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[IB11_sLossLocation] [varchar] (254) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[IB12_dtDateOfLoss] [datetime] NULL ,
	[IB13_cGrossLoss] [money] NULL ,
	[IB14_cDepreciation] [money] NULL ,
	[IB14a_sSupplement] [int] NOT NULL ,
	[IB14b_sRebilled] [int] NOT NULL ,
	[IB15_cDeductible] [money] NULL ,
	[IB15a_cLessExcessLimits] [money] NULL ,
	[IB15b_sExcessLimDesc] [varchar] (254) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[IB15c_cLessMiscellaneous] [money] NULL ,
	[IB15d_cMiscellaneousDesc] [varchar] (254) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[IB16_cNetClaim] [money] NULL ,
	[IB17_cServiceFee] [money] NULL ,
	[IB17a_cMiscServiceFee] [money] NULL ,
	[IB18_sServiceFeeComment] [varchar] (254) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[IB18a_sMiscServiceFeeComment] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[IB25_cServiceFeeSubTotal] [money] NULL ,
	[IB29a_sMiscExpenseFeeComment] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[IB29b_cMiscExpenseFee] [money] NULL ,
	[IB30_cTotalExpenses] [money] NULL ,
	[IB31_dTaxPercent] [numeric](8, 3) NULL ,
	[IB32_cTaxAmount] [money] NULL ,
	[IB33_cTotalAdjustingFee] [money] NULL ,
	[IB33a_sAccountCode] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[FeeScheduleID] [int] NULL ,
	[Void] [bit] NOT NULL ,
	[FeeByTime] [bit] NULL ,
	[UseActivityTime] [bit] NULL ,
	[DownLoadMe] [bit] NOT NULL ,
	[UpLoadMe] [bit] NOT NULL ,
	[Comments] [varchar] (150) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[AdminComments] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[IBStateFarm] (
	[IBStateFarmID] [int] IDENTITY (20000, 1) NOT NULL ,
	[BillAssignmentID] [int] NOT NULL ,
	[BillBillingCountID] [int] NOT NULL ,
	[lssn] [numeric](9, 0) NOT NULL ,
	[IBNumber] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[PolicyNo] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Insured] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[LossLoc1] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[LossLoc2] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[LossLocCity] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[LossLocState] [varchar] (4) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[LossLocZipcode] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[LossDate] [datetime] NULL ,
	[CloseDate] [datetime] NULL ,
	[GrossLoss] [money] NOT NULL ,
	[Supplement] [int] NOT NULL ,
	[SupplementExplain] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[AdditionalLoss] [money] NOT NULL ,
	[Rebilled] [int] NOT NULL ,
	[OrigIBIBNumber] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[OrigIBTotalFee] [money] NOT NULL ,
	[RebillExplain] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[MultiClaimBldgUnitNum] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ClientCompanyCatSpecID] [int] NOT NULL ,
	[SeverityCode] [int] NOT NULL ,
	[ServiceFeeBase] [money] NOT NULL ,
	[ServiceFeeCovAExterior] [money] NOT NULL ,
	[ServiceFeeCovAFraming] [money] NOT NULL ,
	[ServiceFeeCovAInterior] [money] NOT NULL ,
	[ServiceFeeCovB] [money] NOT NULL ,
	[ServiceFeeALE] [money] NOT NULL ,
	[OutBuildCount] [int] NOT NULL ,
	[OutBuildPerItemCharge] [money] NOT NULL ,
	[ServiceFeeOutBuildings] [money] NOT NULL ,
	[ServiceFeeSteepCharge] [money] NOT NULL ,
	[ServiceFeeTwoStory] [money] NOT NULL ,
	[ServiceFeeMoreThan50Squares] [money] NOT NULL ,
	[ServiceFeeWoodSlateTileConRoof] [money] NOT NULL ,
	[ServiceFeeAdditionalDamage] [money] NOT NULL ,
	[ServiceFeeRopeAndHarness] [money] NOT NULL ,
	[ServiceFeeMisc] [money] NOT NULL ,
	[MiscFeesExplain] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ServiceFeeTotal] [money] NOT NULL ,
	[ExpensePagerPhoneExplain] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ExpensePagerPhone] [money] NOT NULL ,
	[ExpenseOtherExplain] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ExpenseOther] [money] NOT NULL ,
	[SumTotalServiceFeeAndExpense] [money] NOT NULL ,
	[TaxPercent] [numeric](8, 3) NOT NULL ,
	[TaxesTotal] [money] NOT NULL ,
	[TotalFee] [money] NOT NULL ,
	[Void] [bit] NOT NULL ,
	[Comments] [varchar] (150) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[AdminComments] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Manager] (
	[CompanyID] [int] NOT NULL ,
	[UsersID] [int] NOT NULL ,
	[Active] [bit] NOT NULL ,
	[DateLastUpdated] [datetime] NULL ,
	[UpdateByUserID] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[MiscReportParam] (
	[MiscReportParamID] [int] IDENTITY (1, 1) NOT NULL ,
	[AssignmentsID] [int] NOT NULL ,
	[ID] [int] NULL ,
	[IDAssignments] [int] NULL ,
	[Number] [int] NULL ,
	[ProjectName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ClassName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ParamName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ParamCaption] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ParamDataType] [int] NOT NULL ,
	[ParamValue] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[SortMe] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DownLoadMe] [bit] NOT NULL ,
	[UpLoadMe] [bit] NOT NULL ,
	[AdminComments] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[MiscReportParam01] (
	[MiscReportParamID] [int] IDENTITY (1, 1) NOT NULL ,
	[AssignmentsID] [int] NOT NULL ,
	[ID] [int] NULL ,
	[IDAssignments] [int] NULL ,
	[Number] [int] NULL ,
	[ProjectName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ClassName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ParamName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ParamCaption] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ParamDataType] [int] NOT NULL ,
	[ParamValue] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[SortMe] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DownLoadMe] [bit] NOT NULL ,
	[UpLoadMe] [bit] NOT NULL ,
	[AdminComments] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[MiscReportParam02] (
	[MiscReportParamID] [int] IDENTITY (1, 1) NOT NULL ,
	[AssignmentsID] [int] NOT NULL ,
	[ID] [int] NULL ,
	[IDAssignments] [int] NULL ,
	[Number] [int] NULL ,
	[ProjectName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ClassName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ParamName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ParamCaption] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ParamDataType] [int] NOT NULL ,
	[ParamValue] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[SortMe] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DownLoadMe] [bit] NOT NULL ,
	[UpLoadMe] [bit] NOT NULL ,
	[AdminComments] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[MiscReportParam03] (
	[MiscReportParamID] [int] IDENTITY (1, 1) NOT NULL ,
	[AssignmentsID] [int] NOT NULL ,
	[ID] [int] NULL ,
	[IDAssignments] [int] NULL ,
	[Number] [int] NULL ,
	[ProjectName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ClassName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ParamName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ParamCaption] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ParamDataType] [int] NOT NULL ,
	[ParamValue] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[SortMe] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DownLoadMe] [bit] NOT NULL ,
	[UpLoadMe] [bit] NOT NULL ,
	[AdminComments] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[MiscReportParam04] (
	[MiscReportParamID] [int] IDENTITY (1, 1) NOT NULL ,
	[AssignmentsID] [int] NOT NULL ,
	[ID] [int] NULL ,
	[IDAssignments] [int] NULL ,
	[Number] [int] NULL ,
	[ProjectName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ClassName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ParamName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ParamCaption] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ParamDataType] [int] NOT NULL ,
	[ParamValue] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[SortMe] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DownLoadMe] [bit] NOT NULL ,
	[UpLoadMe] [bit] NOT NULL ,
	[AdminComments] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[MiscReportParam05] (
	[MiscReportParamID] [int] IDENTITY (1, 1) NOT NULL ,
	[AssignmentsID] [int] NOT NULL ,
	[ID] [int] NULL ,
	[IDAssignments] [int] NULL ,
	[Number] [int] NULL ,
	[ProjectName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ClassName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ParamName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ParamCaption] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ParamDataType] [int] NOT NULL ,
	[ParamValue] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[SortMe] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DownLoadMe] [bit] NOT NULL ,
	[UpLoadMe] [bit] NOT NULL ,
	[AdminComments] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[MiscReportParam06] (
	[MiscReportParamID] [int] IDENTITY (1, 1) NOT NULL ,
	[AssignmentsID] [int] NOT NULL ,
	[ID] [int] NULL ,
	[IDAssignments] [int] NULL ,
	[Number] [int] NULL ,
	[ProjectName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ClassName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ParamName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ParamCaption] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ParamDataType] [int] NOT NULL ,
	[ParamValue] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[SortMe] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DownLoadMe] [bit] NOT NULL ,
	[UpLoadMe] [bit] NOT NULL ,
	[AdminComments] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[MiscReportParam07] (
	[MiscReportParamID] [int] IDENTITY (1, 1) NOT NULL ,
	[AssignmentsID] [int] NOT NULL ,
	[ID] [int] NULL ,
	[IDAssignments] [int] NULL ,
	[Number] [int] NULL ,
	[ProjectName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ClassName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ParamName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ParamCaption] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ParamDataType] [int] NOT NULL ,
	[ParamValue] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[SortMe] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DownLoadMe] [bit] NOT NULL ,
	[UpLoadMe] [bit] NOT NULL ,
	[AdminComments] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[MiscReportParam08] (
	[MiscReportParamID] [int] IDENTITY (1, 1) NOT NULL ,
	[AssignmentsID] [int] NOT NULL ,
	[ID] [int] NULL ,
	[IDAssignments] [int] NULL ,
	[Number] [int] NULL ,
	[ProjectName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ClassName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ParamName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ParamCaption] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ParamDataType] [int] NOT NULL ,
	[ParamValue] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[SortMe] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DownLoadMe] [bit] NOT NULL ,
	[UpLoadMe] [bit] NOT NULL ,
	[AdminComments] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[MiscReportParam09] (
	[MiscReportParamID] [int] IDENTITY (1, 1) NOT NULL ,
	[AssignmentsID] [int] NOT NULL ,
	[ID] [int] NULL ,
	[IDAssignments] [int] NULL ,
	[Number] [int] NULL ,
	[ProjectName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ClassName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ParamName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ParamCaption] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ParamDataType] [int] NOT NULL ,
	[ParamValue] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[SortMe] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DownLoadMe] [bit] NOT NULL ,
	[UpLoadMe] [bit] NOT NULL ,
	[AdminComments] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[MiscReportParam10] (
	[MiscReportParamID] [int] IDENTITY (1, 1) NOT NULL ,
	[AssignmentsID] [int] NOT NULL ,
	[ID] [int] NULL ,
	[IDAssignments] [int] NULL ,
	[Number] [int] NULL ,
	[ProjectName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ClassName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ParamName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ParamCaption] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ParamDataType] [int] NOT NULL ,
	[ParamValue] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[SortMe] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DownLoadMe] [bit] NOT NULL ,
	[UpLoadMe] [bit] NOT NULL ,
	[AdminComments] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[MiscReportParam11] (
	[MiscReportParamID] [int] IDENTITY (1, 1) NOT NULL ,
	[AssignmentsID] [int] NOT NULL ,
	[ID] [int] NULL ,
	[IDAssignments] [int] NULL ,
	[Number] [int] NULL ,
	[ProjectName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ClassName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ParamName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ParamCaption] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ParamDataType] [int] NOT NULL ,
	[ParamValue] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[SortMe] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DownLoadMe] [bit] NOT NULL ,
	[UpLoadMe] [bit] NOT NULL ,
	[AdminComments] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[MiscReportParam12] (
	[MiscReportParamID] [int] IDENTITY (1, 1) NOT NULL ,
	[AssignmentsID] [int] NOT NULL ,
	[ID] [int] NULL ,
	[IDAssignments] [int] NULL ,
	[Number] [int] NULL ,
	[ProjectName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ClassName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ParamName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ParamCaption] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ParamDataType] [int] NOT NULL ,
	[ParamValue] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[SortMe] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DownLoadMe] [bit] NOT NULL ,
	[UpLoadMe] [bit] NOT NULL ,
	[AdminComments] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[MiscReportParam13] (
	[MiscReportParamID] [int] IDENTITY (1, 1) NOT NULL ,
	[AssignmentsID] [int] NOT NULL ,
	[ID] [int] NULL ,
	[IDAssignments] [int] NULL ,
	[Number] [int] NULL ,
	[ProjectName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ClassName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ParamName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ParamCaption] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ParamDataType] [int] NOT NULL ,
	[ParamValue] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[SortMe] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DownLoadMe] [bit] NOT NULL ,
	[UpLoadMe] [bit] NOT NULL ,
	[AdminComments] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[MiscReportParam14] (
	[MiscReportParamID] [int] IDENTITY (1, 1) NOT NULL ,
	[AssignmentsID] [int] NOT NULL ,
	[ID] [int] NULL ,
	[IDAssignments] [int] NULL ,
	[Number] [int] NULL ,
	[ProjectName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ClassName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ParamName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ParamCaption] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ParamDataType] [int] NOT NULL ,
	[ParamValue] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[SortMe] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DownLoadMe] [bit] NOT NULL ,
	[UpLoadMe] [bit] NOT NULL ,
	[AdminComments] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[MiscReportParam15] (
	[MiscReportParamID] [int] IDENTITY (1, 1) NOT NULL ,
	[AssignmentsID] [int] NOT NULL ,
	[ID] [int] NULL ,
	[IDAssignments] [int] NULL ,
	[Number] [int] NULL ,
	[ProjectName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ClassName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ParamName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ParamCaption] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ParamDataType] [int] NOT NULL ,
	[ParamValue] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[SortMe] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DownLoadMe] [bit] NOT NULL ,
	[UpLoadMe] [bit] NOT NULL ,
	[AdminComments] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[MiscReportParam16] (
	[MiscReportParamID] [int] IDENTITY (1, 1) NOT NULL ,
	[AssignmentsID] [int] NOT NULL ,
	[ID] [int] NULL ,
	[IDAssignments] [int] NULL ,
	[Number] [int] NULL ,
	[ProjectName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ClassName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ParamName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ParamCaption] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ParamDataType] [int] NOT NULL ,
	[ParamValue] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[SortMe] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DownLoadMe] [bit] NOT NULL ,
	[UpLoadMe] [bit] NOT NULL ,
	[AdminComments] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[MiscReportParam17] (
	[MiscReportParamID] [int] IDENTITY (1, 1) NOT NULL ,
	[AssignmentsID] [int] NOT NULL ,
	[ID] [int] NULL ,
	[IDAssignments] [int] NULL ,
	[Number] [int] NULL ,
	[ProjectName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ClassName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ParamName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ParamCaption] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ParamDataType] [int] NOT NULL ,
	[ParamValue] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[SortMe] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DownLoadMe] [bit] NOT NULL ,
	[UpLoadMe] [bit] NOT NULL ,
	[AdminComments] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[MiscReportParam18] (
	[MiscReportParamID] [int] IDENTITY (1, 1) NOT NULL ,
	[AssignmentsID] [int] NOT NULL ,
	[ID] [int] NULL ,
	[IDAssignments] [int] NULL ,
	[Number] [int] NULL ,
	[ProjectName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ClassName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ParamName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ParamCaption] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ParamDataType] [int] NOT NULL ,
	[ParamValue] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[SortMe] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DownLoadMe] [bit] NOT NULL ,
	[UpLoadMe] [bit] NOT NULL ,
	[AdminComments] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[MiscReportParam19] (
	[MiscReportParamID] [int] IDENTITY (1, 1) NOT NULL ,
	[AssignmentsID] [int] NOT NULL ,
	[ID] [int] NULL ,
	[IDAssignments] [int] NULL ,
	[Number] [int] NULL ,
	[ProjectName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ClassName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ParamName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ParamCaption] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ParamDataType] [int] NOT NULL ,
	[ParamValue] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[SortMe] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DownLoadMe] [bit] NOT NULL ,
	[UpLoadMe] [bit] NOT NULL ,
	[AdminComments] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[MiscReportParam20] (
	[MiscReportParamID] [int] IDENTITY (1, 1) NOT NULL ,
	[AssignmentsID] [int] NOT NULL ,
	[ID] [int] NULL ,
	[IDAssignments] [int] NULL ,
	[Number] [int] NULL ,
	[ProjectName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ClassName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ParamName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ParamCaption] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ParamDataType] [int] NOT NULL ,
	[ParamValue] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[SortMe] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DownLoadMe] [bit] NOT NULL ,
	[UpLoadMe] [bit] NOT NULL ,
	[AdminComments] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[MiscReportParam21] (
	[MiscReportParamID] [int] IDENTITY (1, 1) NOT NULL ,
	[AssignmentsID] [int] NOT NULL ,
	[ID] [int] NULL ,
	[IDAssignments] [int] NULL ,
	[Number] [int] NULL ,
	[ProjectName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ClassName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ParamName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ParamCaption] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ParamDataType] [int] NOT NULL ,
	[ParamValue] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[SortMe] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DownLoadMe] [bit] NOT NULL ,
	[UpLoadMe] [bit] NOT NULL ,
	[AdminComments] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[MiscReportParam22] (
	[MiscReportParamID] [int] IDENTITY (1, 1) NOT NULL ,
	[AssignmentsID] [int] NOT NULL ,
	[ID] [int] NULL ,
	[IDAssignments] [int] NULL ,
	[Number] [int] NULL ,
	[ProjectName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ClassName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ParamName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ParamCaption] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ParamDataType] [int] NOT NULL ,
	[ParamValue] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[SortMe] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DownLoadMe] [bit] NOT NULL ,
	[UpLoadMe] [bit] NOT NULL ,
	[AdminComments] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[MiscReportParam23] (
	[MiscReportParamID] [int] IDENTITY (1, 1) NOT NULL ,
	[AssignmentsID] [int] NOT NULL ,
	[ID] [int] NULL ,
	[IDAssignments] [int] NULL ,
	[Number] [int] NULL ,
	[ProjectName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ClassName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ParamName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ParamCaption] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ParamDataType] [int] NOT NULL ,
	[ParamValue] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[SortMe] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DownLoadMe] [bit] NOT NULL ,
	[UpLoadMe] [bit] NOT NULL ,
	[AdminComments] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[MiscReportParam24] (
	[MiscReportParamID] [int] IDENTITY (1, 1) NOT NULL ,
	[AssignmentsID] [int] NOT NULL ,
	[ID] [int] NULL ,
	[IDAssignments] [int] NULL ,
	[Number] [int] NULL ,
	[ProjectName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ClassName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ParamName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ParamCaption] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ParamDataType] [int] NOT NULL ,
	[ParamValue] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[SortMe] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DownLoadMe] [bit] NOT NULL ,
	[UpLoadMe] [bit] NOT NULL ,
	[AdminComments] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[MiscReportParam25] (
	[MiscReportParamID] [int] IDENTITY (1, 1) NOT NULL ,
	[AssignmentsID] [int] NOT NULL ,
	[ID] [int] NULL ,
	[IDAssignments] [int] NULL ,
	[Number] [int] NULL ,
	[ProjectName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ClassName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ParamName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ParamCaption] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ParamDataType] [int] NOT NULL ,
	[ParamValue] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[SortMe] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DownLoadMe] [bit] NOT NULL ,
	[UpLoadMe] [bit] NOT NULL ,
	[AdminComments] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[MiscReportParam26] (
	[MiscReportParamID] [int] IDENTITY (1, 1) NOT NULL ,
	[AssignmentsID] [int] NOT NULL ,
	[ID] [int] NULL ,
	[IDAssignments] [int] NULL ,
	[Number] [int] NULL ,
	[ProjectName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ClassName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ParamName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ParamCaption] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ParamDataType] [int] NOT NULL ,
	[ParamValue] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[SortMe] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DownLoadMe] [bit] NOT NULL ,
	[UpLoadMe] [bit] NOT NULL ,
	[AdminComments] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[MiscReportParam27] (
	[MiscReportParamID] [int] IDENTITY (1, 1) NOT NULL ,
	[AssignmentsID] [int] NOT NULL ,
	[ID] [int] NULL ,
	[IDAssignments] [int] NULL ,
	[Number] [int] NULL ,
	[ProjectName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ClassName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ParamName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ParamCaption] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ParamDataType] [int] NOT NULL ,
	[ParamValue] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[SortMe] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DownLoadMe] [bit] NOT NULL ,
	[UpLoadMe] [bit] NOT NULL ,
	[AdminComments] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[MiscReportParam28] (
	[MiscReportParamID] [int] IDENTITY (1, 1) NOT NULL ,
	[AssignmentsID] [int] NOT NULL ,
	[ID] [int] NULL ,
	[IDAssignments] [int] NULL ,
	[Number] [int] NULL ,
	[ProjectName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ClassName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ParamName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ParamCaption] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ParamDataType] [int] NOT NULL ,
	[ParamValue] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[SortMe] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DownLoadMe] [bit] NOT NULL ,
	[UpLoadMe] [bit] NOT NULL ,
	[AdminComments] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[MiscReportParam29] (
	[MiscReportParamID] [int] IDENTITY (1, 1) NOT NULL ,
	[AssignmentsID] [int] NOT NULL ,
	[ID] [int] NULL ,
	[IDAssignments] [int] NULL ,
	[Number] [int] NULL ,
	[ProjectName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ClassName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ParamName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ParamCaption] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ParamDataType] [int] NOT NULL ,
	[ParamValue] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[SortMe] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DownLoadMe] [bit] NOT NULL ,
	[UpLoadMe] [bit] NOT NULL ,
	[AdminComments] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[MiscReportParam30] (
	[MiscReportParamID] [int] IDENTITY (1, 1) NOT NULL ,
	[AssignmentsID] [int] NOT NULL ,
	[ID] [int] NULL ,
	[IDAssignments] [int] NULL ,
	[Number] [int] NULL ,
	[ProjectName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ClassName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ParamName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ParamCaption] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ParamDataType] [int] NOT NULL ,
	[ParamValue] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[SortMe] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DownLoadMe] [bit] NOT NULL ,
	[UpLoadMe] [bit] NOT NULL ,
	[AdminComments] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Package] (
	[PackageID] [int] IDENTITY (20000, 1) NOT NULL ,
	[AssignmentsID] [int] NOT NULL ,
	[ID] [int] NULL ,
	[IDAssignments] [int] NULL ,
	[CreateDate] [datetime] NOT NULL ,
	[PackageStatus] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Name] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Description] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Number] [int] NOT NULL ,
	[SendMe] [bit] NOT NULL ,
	[SentDate] [datetime] NULL ,
	[SentToEmail] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DownLoadMe] [bit] NOT NULL ,
	[UpLoadMe] [bit] NOT NULL ,
	[AdminComments] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[PackageHistory] (
	[PackageHistoryID] [int] IDENTITY (1, 1) NOT NULL ,
	[PackageID] [int] NOT NULL ,
	[AssignmentsID] [int] NOT NULL ,
	[ID] [int] NULL ,
	[IDAssignments] [int] NULL ,
	[CreateDate] [datetime] NOT NULL ,
	[PackageStatus] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Name] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Description] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Number] [int] NULL ,
	[SendMe] [bit] NOT NULL ,
	[SentDate] [datetime] NULL ,
	[SentToEmail] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DownLoadMe] [bit] NOT NULL ,
	[UpLoadMe] [bit] NOT NULL ,
	[AdminComments] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[PackageItem] (
	[PackageItemID] [int] IDENTITY (20000, 1) NOT NULL ,
	[PackageID] [int] NOT NULL ,
	[AssignmentsID] [int] NOT NULL ,
	[ID] [int] NULL ,
	[IDPackage] [int] NULL ,
	[IDAssignments] [int] NULL ,
	[ReportFormat] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[RTAttachmentsID] [int] NULL ,
	[IDRTAttachments] [int] NULL ,
	[Number] [int] NULL ,
	[AttachmentName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[SortOrder] [int] NOT NULL ,
	[Name] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Description] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[IsCoApprove] [bit] NOT NULL ,
	[CoApproveDate] [datetime] NULL ,
	[CoApproveDesc] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[IsClientCoReject] [bit] NOT NULL ,
	[ClientCoRejectDate] [datetime] NULL ,
	[ClientCoRejectDesc] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[IsClientCoDelete] [bit] NOT NULL ,
	[ClientCoDeleteDate] [datetime] NULL ,
	[ClientCoDeleteDesc] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[IsClientCoApprove] [bit] NOT NULL ,
	[ClientCoApproveDate] [datetime] NULL ,
	[ClientCoApproveDesc] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[PackageItemGUID] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[SendMe] [bit] NOT NULL ,
	[SentDate] [datetime] NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DownLoadMe] [bit] NOT NULL ,
	[UpLoadMe] [bit] NOT NULL ,
	[AdminComments] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[PackageItemHistory] (
	[PackageItemHistoryID] [int] IDENTITY (1, 1) NOT NULL ,
	[PackageItemID] [int] NOT NULL ,
	[PackageID] [int] NOT NULL ,
	[AssignmentsID] [int] NOT NULL ,
	[ID] [int] NULL ,
	[IDPackage] [int] NULL ,
	[IDAssignments] [int] NULL ,
	[ReportFormat] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[RTAttachmentsID] [int] NULL ,
	[IDRTAttachments] [int] NULL ,
	[Number] [int] NULL ,
	[AttachmentName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[SortOrder] [int] NOT NULL ,
	[Name] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Description] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[IsCoApprove] [bit] NOT NULL ,
	[CoApproveDate] [datetime] NULL ,
	[CoApproveDesc] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[IsClientCoReject] [bit] NOT NULL ,
	[ClientCoRejectDate] [datetime] NULL ,
	[ClientCoRejectDesc] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[IsClientCoDelete] [bit] NOT NULL ,
	[ClientCoDeleteDate] [datetime] NULL ,
	[ClientCoDeleteDesc] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[IsClientCoApprove] [bit] NOT NULL ,
	[ClientCoApproveDate] [datetime] NULL ,
	[ClientCoApproveDesc] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[PackageItemGUID] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[SendMe] [bit] NOT NULL ,
	[SentDate] [datetime] NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DownLoadMe] [bit] NOT NULL ,
	[UpLoadMe] [bit] NOT NULL ,
	[AdminComments] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[PolicyLimits] (
	[PolicyLimitsID] [int] IDENTITY (20000, 1) NOT NULL ,
	[AssignmentsID] [int] NOT NULL ,
	[ID] [int] NULL ,
	[IDAssignments] [int] NULL ,
	[ClassTypeID] [int] NOT NULL ,
	[LimitAmount] [money] NOT NULL ,
	[RCSaidProp] [money] NOT NULL ,
	[Reserves] [money] NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DownLoadMe] [bit] NOT NULL ,
	[UpLoadMe] [bit] NOT NULL ,
	[AdminComments] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[PolicyLimitsHistory] (
	[PolicyLimitsHistoryID] [int] IDENTITY (1, 1) NOT NULL ,
	[PolicyLimitsID] [int] NOT NULL ,
	[AssignmentsID] [int] NOT NULL ,
	[ID] [int] NULL ,
	[IDAssignments] [int] NULL ,
	[ClassTypeID] [int] NOT NULL ,
	[LimitAmount] [money] NOT NULL ,
	[RCSaidProp] [money] NOT NULL ,
	[Reserves] [money] NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DownLoadMe] [bit] NOT NULL ,
	[UpLoadMe] [bit] NOT NULL ,
	[AdminComments] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[RTActivityLog] (
	[RTActivityLogID] [int] IDENTITY (20000, 1) NOT NULL ,
	[AssignmentsID] [int] NOT NULL ,
	[BillingCountID] [int] NULL ,
	[ID] [int] NULL ,
	[IDAssignments] [int] NULL ,
	[IDBillingCount] [int] NULL ,
	[ServiceTime] [numeric](10, 2) NULL ,
	[ActDate] [datetime] NULL ,
	[ActText] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ActTime] [datetime] NULL ,
	[PageBreakAfter] [bit] NOT NULL ,
	[BlankPageAfter] [bit] NOT NULL ,
	[BlankRowsAfter] [int] NOT NULL ,
	[IsMgrEntry] [bit] NOT NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DownLoadMe] [bit] NOT NULL ,
	[UpLoadMe] [bit] NOT NULL ,
	[AdminComments] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[RTActivityLogHistory] (
	[RTActivityLogHistoryID] [int] IDENTITY (1, 1) NOT NULL ,
	[RTActivityLogID] [int] NOT NULL ,
	[AssignmentsID] [int] NOT NULL ,
	[BillingCountID] [int] NULL ,
	[ID] [int] NULL ,
	[IDAssignments] [int] NULL ,
	[IDBillingCount] [int] NULL ,
	[ServiceTime] [numeric](10, 2) NULL ,
	[ActDate] [datetime] NULL ,
	[ActText] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ActTime] [datetime] NULL ,
	[PageBreakAfter] [bit] NOT NULL ,
	[BlankPageAfter] [bit] NOT NULL ,
	[BlankRowsAfter] [int] NOT NULL ,
	[IsMgrEntry] [bit] NOT NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DownLoadMe] [bit] NOT NULL ,
	[UpLoadMe] [bit] NOT NULL ,
	[AdminComments] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[RTActivityLogInfo] (
	[AssignmentsID] [int] NOT NULL ,
	[IDAssignments] [int] NULL ,
	[AL01_sPresentDurringInspection] [bit] NULL ,
	[AL02_sExplainedEstimate] [bit] NULL ,
	[AL03_sExplainedRCV] [bit] NULL ,
	[AL03_sExplainedRCVNA] [bit] NULL ,
	[AL04_sConfirmMortgageeIsCorrect] [bit] NULL ,
	[AL04_sConfirmMortgageeIsCorrectNA] [bit] NULL ,
	[AL05_sExplainedMortgageeChecks] [bit] NULL ,
	[AL05_sExplainedMortgageeChecksNA] [bit] NULL ,
	[AL06_sConfirmedCoverage] [bit] NULL ,
	[AL07_sPriorLoss] [bit] NULL ,
	[AL07_sPriorLossNA] [bit] NULL ,
	[AL08_sSalvage] [bit] NULL ,
	[AL09_sSubrogation] [bit] NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DownLoadMe] [bit] NOT NULL ,
	[UpLoadMe] [bit] NOT NULL ,
	[AdminComments] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[RTActivityLogInfoHistory] (
	[RTActivityLogInfoHistoryID] [int] IDENTITY (1, 1) NOT NULL ,
	[AssignmentsID] [int] NULL ,
	[IDAssignments] [int] NULL ,
	[AL01_sPresentDurringInspection] [bit] NULL ,
	[AL02_sExplainedEstimate] [bit] NULL ,
	[AL03_sExplainedRCV] [bit] NULL ,
	[AL03_sExplainedRCVNA] [bit] NULL ,
	[AL04_sConfirmMortgageeIsCorrect] [bit] NULL ,
	[AL04_sConfirmMortgageeIsCorrectNA] [bit] NULL ,
	[AL05_sExplainedMortgageeChecks] [bit] NULL ,
	[AL05_sExplainedMortgageeChecksNA] [bit] NULL ,
	[AL06_sConfirmedCoverage] [bit] NULL ,
	[AL07_sPriorLoss] [bit] NULL ,
	[AL07_sPriorLossNA] [bit] NULL ,
	[AL08_sSalvage] [bit] NULL ,
	[AL09_sSubrogation] [bit] NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DownLoadMe] [bit] NOT NULL ,
	[UpLoadMe] [bit] NOT NULL ,
	[AdminComments] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[RTAttachments] (
	[RTAttachmentsID] [int] IDENTITY (20000, 1) NOT NULL ,
	[AssignmentsID] [int] NOT NULL ,
	[ID] [int] NULL ,
	[IDAssignments] [int] NULL ,
	[AttachDate] [datetime] NOT NULL ,
	[SortOrder] [int] NOT NULL ,
	[Description] [varchar] (254) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[AttachName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Attachment] [varchar] (254) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[DownloadAttachment] [bit] NOT NULL ,
	[UpLoadAttachment] [bit] NOT NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DownLoadMe] [bit] NOT NULL ,
	[UpLoadMe] [bit] NOT NULL ,
	[AdminComments] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[RTAttachmentsHistory] (
	[RTAttachmentsHistoryID] [int] IDENTITY (1, 1) NOT NULL ,
	[RTAttachmentsID] [int] NOT NULL ,
	[AssignmentsID] [int] NOT NULL ,
	[ID] [int] NULL ,
	[IDAssignments] [int] NULL ,
	[AttachDate] [datetime] NULL ,
	[SortOrder] [int] NULL ,
	[Description] [varchar] (254) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[AttachName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Attachment] [varchar] (254) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DownloadAttachment] [bit] NOT NULL ,
	[UpLoadAttachment] [bit] NOT NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DownLoadMe] [bit] NOT NULL ,
	[UpLoadMe] [bit] NOT NULL ,
	[AdminComments] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[RTChecks] (
	[RTChecksID] [int] IDENTITY (20000, 1) NOT NULL ,
	[AssignmentsID] [int] NOT NULL ,
	[BillingCountID] [int] NULL ,
	[ID] [int] NULL ,
	[IDAssignments] [int] NOT NULL ,
	[IDBillingCount] [int] NULL ,
	[CheckNum] [int] NOT NULL ,
	[RT42_ClassOfLossID] [int] NULL ,
	[RT43_TypeOfLossID] [int] NULL ,
	[RT50_sInsuredPayeeName] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[RT51_sPayeeNames] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[RT52_sAddress] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[RT53_cAmountOfCheck] [money] NULL ,
	[AppliedDeductible] [money] NULL ,
	[RT54_CompanyCatSpecID] [int] NOT NULL ,
	[tempCHeckName] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[PrintOnIB] [bit] NOT NULL ,
	[PrintedDate] [datetime] NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DownLoadMe] [bit] NOT NULL ,
	[UpLoadMe] [bit] NOT NULL ,
	[AdminComments] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[RTChecksHistory] (
	[RTChecksHistoryID] [int] IDENTITY (1, 1) NOT NULL ,
	[RTChecksID] [int] NOT NULL ,
	[AssignmentsID] [int] NOT NULL ,
	[BillingCountID] [int] NULL ,
	[ID] [int] NULL ,
	[IDAssignments] [int] NOT NULL ,
	[IDBillingCount] [int] NULL ,
	[CheckNum] [int] NOT NULL ,
	[RT42_ClassOfLossID] [int] NULL ,
	[RT43_TypeOfLossID] [int] NULL ,
	[RT50_sInsuredPayeeName] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[RT51_sPayeeNames] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[RT52_sAddress] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[RT53_cAmountOfCheck] [money] NULL ,
	[AppliedDeductible] [money] NULL ,
	[RT54_CompanyCatSpecID] [int] NOT NULL ,
	[tempCHeckName] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[PrintOnIB] [bit] NOT NULL ,
	[PrintedDate] [datetime] NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DownLoadMe] [bit] NOT NULL ,
	[UpLoadMe] [bit] NOT NULL ,
	[AdminComments] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[RTIB] (
	[AssignmentsID] [int] NOT NULL ,
	[BillingCountID] [int] NOT NULL ,
	[IDAssignments] [int] NOT NULL ,
	[IDBillingCount] [int] NOT NULL ,
	[RT00_lSSN] [numeric](9, 0) NULL ,
	[RT01_sSubToCarrier] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[RT02_sIBNumber] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[RT05_sLocation] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[RT05a_sState] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[RT06_dtDateClosed] [datetime] NULL ,
	[RT07_sAdjusterName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[RT09_sSALN] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[RT09a_sPolicyNo] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[RT10_sInsuredName] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[RT11_sLossLocation] [varchar] (254) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[RT12_dtDateOfLoss] [datetime] NULL ,
	[RT13_cGrossLoss] [money] NULL ,
	[RT14_cDepreciation] [money] NULL ,
	[RT14a_sSupplement] [int] NOT NULL ,
	[RT14b_sRebilled] [int] NOT NULL ,
	[RT15_cDeductible] [money] NULL ,
	[RT15a_cLessExcessLimits] [money] NULL ,
	[RT15b_sExcessLimDesc] [varchar] (254) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[RT15c_cLessMiscellaneous] [money] NULL ,
	[RT15d_cMiscellaneousDesc] [varchar] (254) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[RT16_cNetClaim] [money] NULL ,
	[RT17_cServiceFee] [money] NULL ,
	[RT17a_cMiscServiceFee] [money] NULL ,
	[RT18_sServiceFeeComment] [varchar] (254) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[RT18a_sMiscServiceFeeComment] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[RT25_cServiceFeeSubTotal] [money] NULL ,
	[RT29a_sMiscExpenseFeeComment] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[RT29b_cMiscExpenseFee] [money] NULL ,
	[RT30_cTotalExpenses] [money] NULL ,
	[RT31_dTaxPercent] [numeric](8, 3) NULL ,
	[RT32_cTaxAmount] [money] NULL ,
	[RT33_cTotalAdjustingFee] [money] NULL ,
	[RT33a_sAccountCode] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[FeeScheduleID] [int] NULL ,
	[Void] [bit] NOT NULL ,
	[FeeByTime] [bit] NOT NULL ,
	[UseActivityTime] [bit] NOT NULL ,
	[DownLoadMe] [bit] NOT NULL ,
	[UpLoadMe] [bit] NOT NULL ,
	[Comments] [varchar] (150) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[AdminComments] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[RTIBFee] (
	[RTIBFeeID] [int] IDENTITY (20000, 1) NOT NULL ,
	[AssignmentsID] [int] NOT NULL ,
	[ID] [int] NULL ,
	[IDAssignments] [int] NULL ,
	[FeeScheduleFeeTypesID] [int] NOT NULL ,
	[NumberOfItems] [int] NOT NULL ,
	[Amount] [money] NOT NULL ,
	[Comment] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[DownLoadMe] [bit] NOT NULL ,
	[UpLoadMe] [bit] NOT NULL ,
	[AdminComments] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[RTIBFeeHistory] (
	[RTIBFeeHistoryID] [int] IDENTITY (20000, 1) NOT NULL ,
	[RTIBFeeID] [int] NOT NULL ,
	[AssignmentsID] [int] NULL ,
	[ID] [int] NULL ,
	[IDAssignments] [int] NULL ,
	[FeeScheduleFeeTypesID] [int] NULL ,
	[NumberOfItems] [int] NULL ,
	[Amount] [money] NULL ,
	[Comment] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DownLoadMe] [bit] NULL ,
	[UpLoadMe] [bit] NULL ,
	[AdminComments] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DateLastUpdated] [datetime] NULL ,
	[UpdateByUserID] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[RTIBHistory] (
	[RTIBHistoryID] [int] IDENTITY (1, 1) NOT NULL ,
	[AssignmentsID] [int] NOT NULL ,
	[BillingCountID] [int] NOT NULL ,
	[IDAssignments] [int] NOT NULL ,
	[IDBillingCount] [int] NOT NULL ,
	[RT00_lSSN] [numeric](9, 0) NULL ,
	[RT01_sSubToCarrier] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[RT02_sIBNumber] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[RT05_sLocation] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[RT05a_sState] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[RT06_dtDateClosed] [datetime] NULL ,
	[RT07_sAdjusterName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[RT09_sSALN] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[RT09a_sPolicyNo] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[RT10_sInsuredName] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[RT11_sLossLocation] [varchar] (254) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[RT12_dtDateOfLoss] [datetime] NULL ,
	[RT13_cGrossLoss] [money] NULL ,
	[RT14_cDepreciation] [money] NULL ,
	[RT14a_sSupplement] [int] NULL ,
	[RT14b_sRebilled] [int] NULL ,
	[RT15_cDeductible] [money] NULL ,
	[RT15a_cLessExcessLimits] [money] NULL ,
	[RT15b_sExcessLimDesc] [varchar] (254) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[RT15c_cLessMiscellaneous] [money] NULL ,
	[RT15d_cMiscellaneousDesc] [varchar] (254) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[RT16_cNetClaim] [money] NULL ,
	[RT17_cServiceFee] [money] NULL ,
	[RT17a_cMiscServiceFee] [money] NULL ,
	[RT18_sServiceFeeComment] [varchar] (254) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[RT18a_sMiscServiceFeeComment] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[RT25_cServiceFeeSubTotal] [money] NULL ,
	[RT29a_sMiscExpenseFeeComment] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[RT29b_cMiscExpenseFee] [money] NULL ,
	[RT30_cTotalExpenses] [money] NULL ,
	[RT31_dTaxPercent] [numeric](8, 3) NULL ,
	[RT32_cTaxAmount] [money] NULL ,
	[RT33_cTotalAdjustingFee] [money] NULL ,
	[RT33a_sAccountCode] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[FeeScheduleID] [int] NULL ,
	[Void] [bit] NOT NULL ,
	[FeeByTime] [bit] NULL ,
	[UseActivityTime] [bit] NULL ,
	[DownLoadMe] [bit] NULL ,
	[UpLoadMe] [bit] NULL ,
	[Comments] [varchar] (150) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[AdminComments] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[RTIndemnity] (
	[RTIndemnityID] [int] IDENTITY (20000, 1) NOT NULL ,
	[AssignmentsID] [int] NOT NULL ,
	[RTChecksID] [int] NULL ,
	[ID] [int] NULL ,
	[IDAssignments] [int] NULL ,
	[IDRTChecks] [int] NULL ,
	[ACVClaim] [money] NOT NULL ,
	[ACVLessExcessLimits] [money] NOT NULL ,
	[SpecialLimits] [money] NOT NULL ,
	[ExcessLimits] [money] NOT NULL ,
	[Miscellaneous] [money] NULL ,
	[MiscDescription] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[IsAddAmountOfInsurance] [bit] NOT NULL ,
	[ExcessAbsorbsDeductible] [bit] NULL ,
	[AppliedDeductible] [money] NULL ,
	[NonRecoverableDep] [money] NOT NULL ,
	[RecoverableDep] [money] NOT NULL ,
	[ReplacementCost] [money] NOT NULL ,
	[TypeOfLossID] [int] NOT NULL ,
	[ClassOfLossID] [int] NOT NULL ,
	[Description] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[IsPreviousPayment] [bit] NOT NULL ,
	[PPayDatePaid] [datetime] NULL ,
	[PPayAmountPaid] [money] NULL ,
	[PPayCheckNumber] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DownLoadMe] [bit] NOT NULL ,
	[UpLoadMe] [bit] NOT NULL ,
	[AdminComments] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[RTIndemnityHistory] (
	[RTIndemnityHistoryID] [int] IDENTITY (1, 1) NOT NULL ,
	[RTIndemnityID] [int] NOT NULL ,
	[AssignmentsID] [int] NOT NULL ,
	[RTChecksID] [int] NULL ,
	[ID] [int] NULL ,
	[IDAssignments] [int] NOT NULL ,
	[IDRTChecks] [int] NULL ,
	[ACVClaim] [money] NOT NULL ,
	[ACVLessExcessLimits] [money] NOT NULL ,
	[SpecialLimits] [money] NOT NULL ,
	[ExcessLimits] [money] NOT NULL ,
	[Miscellaneous] [money] NULL ,
	[MiscDescription] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[IsAddAmountOfInsurance] [bit] NOT NULL ,
	[ExcessAbsorbsDeductible] [bit] NULL ,
	[AppliedDeductible] [money] NULL ,
	[NonRecoverableDep] [money] NOT NULL ,
	[RecoverableDep] [money] NOT NULL ,
	[ReplacementCost] [money] NOT NULL ,
	[TypeOfLossID] [int] NOT NULL ,
	[ClassOfLossID] [int] NOT NULL ,
	[Description] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[IsPreviousPayment] [bit] NOT NULL ,
	[PPayDatePaid] [datetime] NULL ,
	[PPayAmountPaid] [money] NULL ,
	[PPayCheckNumber] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DownLoadMe] [bit] NOT NULL ,
	[UpLoadMe] [bit] NOT NULL ,
	[AdminComments] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[RTPhotoLog] (
	[RTPhotoLogID] [int] IDENTITY (20000, 1) NOT NULL ,
	[RTPhotoReportID] [int] NOT NULL ,
	[AssignmentsID] [int] NOT NULL ,
	[BillingCountID] [int] NULL ,
	[ID] [int] NULL ,
	[IDRTPhotoReport] [int] NULL ,
	[IDAssignments] [int] NULL ,
	[IDBillingCount] [int] NULL ,
	[PhotoDate] [datetime] NULL ,
	[SortOrder] [int] NULL ,
	[Description] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[PhotoName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Photo] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DownloadPhoto] [bit] NOT NULL ,
	[UpLoadPhoto] [bit] NOT NULL ,
	[PhotoThumb] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DownloadPhotoThumb] [bit] NOT NULL ,
	[UpLoadPhotoThumb] [bit] NOT NULL ,
	[PhotoHighRes] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DownloadPhotoHighRes] [bit] NOT NULL ,
	[UploadPhotoHighRes] [bit] NOT NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DownLoadMe] [bit] NOT NULL ,
	[UpLoadMe] [bit] NOT NULL ,
	[AdminComments] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[RTPhotoLogHistory] (
	[RTPhotoLogHistoryID] [int] IDENTITY (1, 1) NOT NULL ,
	[RTPhotoLogID] [int] NOT NULL ,
	[RTPhotoReportID] [int] NULL ,
	[AssignmentsID] [int] NOT NULL ,
	[BillingCountID] [int] NULL ,
	[ID] [int] NULL ,
	[IDRTPhotoReport] [int] NULL ,
	[IDAssignments] [int] NULL ,
	[IDBillingCount] [int] NULL ,
	[PhotoDate] [datetime] NULL ,
	[SortOrder] [int] NULL ,
	[Description] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[PhotoName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Photo] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DownloadPhoto] [bit] NOT NULL ,
	[UpLoadPhoto] [bit] NOT NULL ,
	[PhotoThumb] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DownloadPhotoThumb] [bit] NOT NULL ,
	[UpLoadPhotoThumb] [bit] NOT NULL ,
	[PhotoHighRes] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DownloadPhotoHighRes] [bit] NOT NULL ,
	[UploadPhotoHighRes] [bit] NOT NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DownLoadMe] [bit] NOT NULL ,
	[UpLoadMe] [bit] NOT NULL ,
	[AdminComments] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[RTPhotoReport] (
	[RTPhotoReportID] [int] IDENTITY (1, 1) NOT NULL ,
	[AssignmentsID] [int] NOT NULL ,
	[ID] [int] NULL ,
	[IDAssignments] [int] NULL ,
	[Name] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Description] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Number] [int] NOT NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DownLoadMe] [bit] NOT NULL ,
	[UpLoadMe] [bit] NOT NULL ,
	[AdminComments] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[RTPhotoReportHistory] (
	[RTPhotoReportHistoryID] [int] IDENTITY (1, 1) NOT NULL ,
	[RTPhotoReportID] [int] NOT NULL ,
	[AssignmentsID] [int] NOT NULL ,
	[ID] [int] NULL ,
	[IDAssignments] [int] NULL ,
	[Name] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Description] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Number] [int] NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DownLoadMe] [bit] NOT NULL ,
	[UpLoadMe] [bit] NOT NULL ,
	[AdminComments] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[RTWSDiagram] (
	[RTWSDiagramID] [int] IDENTITY (1, 1) NOT NULL ,
	[AssignmentsID] [int] NOT NULL ,
	[ID] [int] NULL ,
	[IDAssignments] [int] NULL ,
	[Name] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Description] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Number] [int] NOT NULL ,
	[DiagramPhotoName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[DownloadDiagramPhoto] [bit] NOT NULL ,
	[UploadDiagramPhoto] [bit] NOT NULL ,
	[DiagramXML] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DownLoadMe] [bit] NOT NULL ,
	[UpLoadMe] [bit] NOT NULL ,
	[AdminComments] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[RTWSDiagramHistory] (
	[RTWSDiagramHistoryID] [int] IDENTITY (1, 1) NOT NULL ,
	[RTWSDiagramID] [int] NOT NULL ,
	[AssignmentsID] [int] NOT NULL ,
	[ID] [int] NULL ,
	[IDAssignments] [int] NULL ,
	[Name] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Description] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Number] [int] NOT NULL ,
	[DiagramPhotoName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[DownloadDiagramPhoto] [bit] NOT NULL ,
	[UploadDiagramPhoto] [bit] NOT NULL ,
	[DiagramXML] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DownLoadMe] [bit] NOT NULL ,
	[UpLoadMe] [bit] NOT NULL ,
	[AdminComments] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[RegSetting] (
	[RegSettingID] [int] IDENTITY (1, 1) NOT NULL ,
	[RegNameBase] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[RegName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Description] [varchar] (500) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Version] [int] NOT NULL ,
	[SPVersionBase] [int] NOT NULL ,
	[SPVersion] [int] NOT NULL ,
	[VersionDate] [datetime] NOT NULL ,
	[SectionLevel01] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[SectionLevel02] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[SectionLevel03] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[SectionLevel04] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[SectionLevel05] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[SPName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[RegSettingHistory] (
	[RegSettingID] [int] NOT NULL ,
	[RegNameBase] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[RegName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Description] [varchar] (500) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Version] [int] NOT NULL ,
	[SPVersionBase] [int] NULL ,
	[SPVersion] [int] NULL ,
	[VersionDate] [datetime] NOT NULL ,
	[SectionLevel01] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[SectionLevel02] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[SectionLevel03] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[SectionLevel04] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[SectionLevel05] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[SPName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[SASecurityPackage] (
	[SecurityAreaID] [int] NOT NULL ,
	[SecurityPackageID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[SF_AdjusterEvaluations] (
	[EvaluationID] [int] IDENTITY (1, 1) NOT NULL ,
	[EvaluationDate] [datetime] NULL ,
	[CatCode] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ArrivalDate] [datetime] NULL ,
	[DepartureDate] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[CatOfficeLocation] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[SectionMgr] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DivisionMgr] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[SF_AdjusterID] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[AdjUID] [int] NULL ,
	[UserID] [int] NULL ,
	[AdjFirstName] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[AdjLastName] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[NumFilesAssigned] [int] NULL ,
	[NumFilesInspected] [int] NULL ,
	[NumFilesClosed] [int] NULL ,
	[NumFilesAverageEstimate] [money] NULL ,
	[EstimatingPlatform] [nvarchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[EstimatingSkills] [int] NULL ,
	[ComputerSkills] [int] NULL ,
	[AccuracyofScope] [int] NULL ,
	[PrioritizationofAssignments] [int] NULL ,
	[Productivity] [int] NULL ,
	[Professionalism] [int] NULL ,
	[TelephoneFollowup] [int] NULL ,
	[SubmitsAccurateBilling] [int] NULL ,
	[FutureAssignments] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[CatTeamMgr] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[WindHail] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Hurricane] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Flood] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Other] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Earthquake] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[FreezeStorm] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Commercial] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[OtherLossesHandled] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[txtEvalNotes] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[EnteredBy] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DateEntered] [datetime] NULL ,
	[ModifiedBy] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ModifiedDate] [datetime] NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[SISecurityArea] (
	[SecurityItemsID] [int] NOT NULL ,
	[SecurityAreaID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[SPSecurity] (
	[SecurityPackageID] [int] NOT NULL ,
	[SecurityID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Security] (
	[SecurityID] [int] IDENTITY (1, 1) NOT NULL ,
	[Name] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Description] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DateLastUpdated] [datetime] NULL ,
	[UpdateByUserID] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[SecurityArea] (
	[SecurityAreaID] [int] IDENTITY (1, 1) NOT NULL ,
	[SecurityAreaTypeID] [int] NULL ,
	[AreaName] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[AreaDescription] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DateLastUpdated] [datetime] NULL ,
	[UpdateByUserID] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[SecurityAreaType] (
	[SecurityAreaTypeID] [int] IDENTITY (1, 1) NOT NULL ,
	[TypeName] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[TypeDescription] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DateLastUpdated] [datetime] NULL ,
	[UpdateByUserID] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[SecurityGroup] (
	[SecurityID] [int] NOT NULL ,
	[GroupID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[SecurityItemType] (
	[SecurityItemTypeID] [int] IDENTITY (1, 1) NOT NULL ,
	[TypeName] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[TypeDescription] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DateLastUpdated] [datetime] NULL ,
	[UpdateByUserID] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[SecurityItems] (
	[SecurityItemsID] [int] IDENTITY (1, 1) NOT NULL ,
	[SecurityItemTypeID] [int] NULL ,
	[ItemName] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ItemDescription] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DateLastUpdated] [datetime] NULL ,
	[UpdateByUserID] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[SecurityLevel] (
	[SecurityLevel] [int] IDENTITY (1, 1) NOT NULL ,
	[Name] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Description] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[SecurityLevelHistory] (
	[SecurityLevelHistoryID] [int] IDENTITY (1, 1) NOT NULL ,
	[SecurityLevel] [int] NOT NULL ,
	[Name] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Description] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[SecurityPackage] (
	[SecurityPackageID] [int] IDENTITY (1, 1) NOT NULL ,
	[PackageName] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Description] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DateLastUpdated] [datetime] NULL ,
	[UpdateByUserID] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[SoftwarePackage] (
	[SoftWarePackageID] [int] IDENTITY (1, 1) NOT NULL ,
	[ClientCompanyID] [int] NOT NULL ,
	[CATID] [int] NOT NULL ,
	[PackageName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Description] [varchar] (500) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[SPVersion] [int] NOT NULL ,
	[VersionDate] [datetime] NOT NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[SoftwarePackageApplication] (
	[ApplicationID] [int] NOT NULL ,
	[SoftWarePackageID] [int] NOT NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[SoftwarePackageDocument] (
	[DocumentID] [int] NOT NULL ,
	[SoftWarePackageID] [int] NOT NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[SoftwarePackageHistory] (
	[SoftWarePackageHistoryID] [int] IDENTITY (1, 1) NOT NULL ,
	[SoftWarePackageID] [int] NOT NULL ,
	[ClientCompanyID] [int] NOT NULL ,
	[CATID] [int] NOT NULL ,
	[PackageName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Description] [varchar] (500) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[SPVersion] [int] NOT NULL ,
	[VersionDate] [datetime] NOT NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[SoftwarePackageRegSetting] (
	[RegSettingID] [int] NOT NULL ,
	[SoftWarePackageID] [int] NOT NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[State] (
	[StateID] [int] IDENTITY (1, 1) NOT NULL ,
	[Name] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Code] [varchar] (3) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Comments] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[StateHistory] (
	[StateHistoryID] [int] IDENTITY (1, 1) NOT NULL ,
	[StateID] [int] NOT NULL ,
	[Name] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Code] [varchar] (3) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Comments] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Status] (
	[StatusID] [int] IDENTITY (1, 1) NOT NULL ,
	[StatusAlias] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Status] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Description] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[AdminComments] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[StatusHistory] (
	[StatusHistoryID] [int] IDENTITY (1, 1) NOT NULL ,
	[StatusID] [int] NOT NULL ,
	[StatusAlias] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Status] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Description] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[AdminComments] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Temporary] (
	[CompanyID] [int] NOT NULL ,
	[UsersID] [int] NOT NULL ,
	[Active] [bit] NOT NULL ,
	[DateLastUpdated] [datetime] NULL ,
	[UpdateByUserID] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[TransType] (
	[TransTypeID] [int] IDENTITY (1, 1) NOT NULL ,
	[ClientCompanyID] [int] NOT NULL ,
	[Inbound] [bit] NOT NULL ,
	[AllowProcessDefault] [bit] NOT NULL ,
	[TransType] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Definition] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[TypeOfLoss] (
	[TypeOfLossID] [int] IDENTITY (1, 1) NOT NULL ,
	[ClientCompanyID] [int] NOT NULL ,
	[TypeOfLoss] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Description] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Code] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[TypeOfLossHistory] (
	[TypeOfLossHistoryID] [int] IDENTITY (1, 1) NOT NULL ,
	[TypeOfLossID] [int] NOT NULL ,
	[ClientCompanyID] [int] NOT NULL ,
	[TypeOfLoss] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Description] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Code] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[IsDeleted] [bit] NOT NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[UserProfile] (
	[UserProfileID] [int] IDENTITY (1, 1) NOT NULL ,
	[TableName] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Description] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Active] [bit] NOT NULL ,
	[SortOrder] [int] NOT NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[UserProfileHistory] (
	[UserProfileHistoryID] [int] IDENTITY (1, 1) NOT NULL ,
	[UserProfileID] [int] NOT NULL ,
	[TableName] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Description] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Active] [bit] NOT NULL ,
	[SortOrder] [int] NOT NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[UserReportsToCoordinator] (
	[UsersID] [int] NOT NULL ,
	[CompanyID] [int] NOT NULL ,
	[ReportsToUsersID] [int] NOT NULL ,
	[Active] [bit] NOT NULL ,
	[DateLastUpdated] [datetime] NULL ,
	[UpdateByUserID] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[UserReportsToManager] (
	[UsersID] [int] NOT NULL ,
	[CompanyID] [int] NOT NULL ,
	[ReportsToUsersID] [int] NOT NULL ,
	[Active] [bit] NOT NULL ,
	[DateLastUpdated] [datetime] NULL ,
	[UpdateByUserID] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Users] (
	[UsersID] [int] IDENTITY (1, 1) NOT NULL ,
	[UserName] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[PassWord] [varchar] (300) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[FirstName] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[LastName] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[SSN] [int] NULL ,
	[Email] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ContactPhone] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[EmergencyPhone] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Address] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[City] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[State] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Zip] [int] NULL ,
	[ZIP4] [int] NULL ,
	[OtherPostCode] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Active] [bit] NOT NULL ,
	[ActiveDate] [datetime] NOT NULL ,
	[InactiveDate] [datetime] NULL ,
	[SecurityLevel] [int] NOT NULL ,
	[Comments] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[UsersGroup] (
	[UsersID] [int] NOT NULL ,
	[GroupID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[UsersHistory] (
	[UsersHistoryID] [int] IDENTITY (1, 1) NOT NULL ,
	[UsersID] [int] NOT NULL ,
	[UserName] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[PassWord] [varchar] (300) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[FirstName] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[LastName] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[SSN] [int] NULL ,
	[Email] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ContactPhone] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[EmergencyPhone] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Address] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[City] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[State] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Zip] [int] NULL ,
	[ZIP4] [int] NULL ,
	[OtherPostCode] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Active] [bit] NULL ,
	[ActiveDate] [datetime] NULL ,
	[InactiveDate] [datetime] NULL ,
	[SecurityLevel] [int] NOT NULL ,
	[Comments] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Work_ListText] (
	[Work_ListTextID] [uniqueidentifier] NOT NULL ,
	[ListText] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[XML_Trans] (
	[XML_TransID] [int] IDENTITY (1, 1) NOT NULL ,
	[Mod20] [int] NOT NULL ,
	[ClientCompanyID] [int] NOT NULL ,
	[TransFromCompanyID] [int] NOT NULL ,
	[TransToCompanyID] [int] NOT NULL ,
	[AssignmentsID] [int] NULL ,
	[InBound] [bit] NOT NULL ,
	[AllowProcess] [bit] NOT NULL ,
	[XMLDoc] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[TransTypeID] [int] NULL ,
	[TransType] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[TransDescription] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Comments] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[IsErr] [bit] NOT NULL ,
	[ErrMess] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[DatePosted] [datetime] NOT NULL ,
	[DateProcessed] [datetime] NULL ,
	[DateLastUpdated] [datetime] NOT NULL ,
	[remote_addr] [varchar] (500) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[UpdateByUserID] [int] NOT NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

ALTER TABLE [dbo].[Accounting] WITH NOCHECK ADD 
	CONSTRAINT [PK_Accounting] PRIMARY KEY  CLUSTERED 
	(
		[CompanyID],
		[UsersID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Adjuster] WITH NOCHECK ADD 
	CONSTRAINT [PK_CompanyAdjusterUsers] PRIMARY KEY  CLUSTERED 
	(
		[CompanyID],
		[UsersID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[AdjusterEvaluations] WITH NOCHECK ADD 
	CONSTRAINT [PK_AdjusterEvaluations] PRIMARY KEY  CLUSTERED 
	(
		[EvaluationID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[AdjusterUsersSoftware] WITH NOCHECK ADD 
	CONSTRAINT [PK_UsersSoftware] PRIMARY KEY  CLUSTERED 
	(
		[UsersID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[AdjusterUsersSoftwareHistory] WITH NOCHECK ADD 
	CONSTRAINT [PK_UsersSoftwareHistory] PRIMARY KEY  CLUSTERED 
	(
		[AdjusterUsersSoftwareHistoryID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[AdjusterUsersUpdates] WITH NOCHECK ADD 
	CONSTRAINT [PK_UsersUpdates] PRIMARY KEY  CLUSTERED 
	(
		[UsersID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[AdjusterUsersUpdatesHistory] WITH NOCHECK ADD 
	CONSTRAINT [PK_UsersUpdatesHistory] PRIMARY KEY  CLUSTERED 
	(
		[AdjusterUsersUpdatesHistoryID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Admin] WITH NOCHECK ADD 
	CONSTRAINT [PK_CompanyAdminUsers] PRIMARY KEY  CLUSTERED 
	(
		[CompanyID],
		[UsersID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Application] WITH NOCHECK ADD 
	CONSTRAINT [PK_Application] PRIMARY KEY  CLUSTERED 
	(
		[ApplicationID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[ApplicationHistory] WITH NOCHECK ADD 
	CONSTRAINT [PK_ApplicationHistory] PRIMARY KEY  CLUSTERED 
	(
		[ApplicationID],
		[Version]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[AssignmentType] WITH NOCHECK ADD 
	CONSTRAINT [PK_AssignmentType] PRIMARY KEY  CLUSTERED 
	(
		[AssignmentTypeID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[AssignmentTypeHistory] WITH NOCHECK ADD 
	CONSTRAINT [PK_AssignmentTypeHistory] PRIMARY KEY  CLUSTERED 
	(
		[AssignmentTypeHistoryID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Assignments] WITH NOCHECK ADD 
	CONSTRAINT [PK_Assignments] PRIMARY KEY  CLUSTERED 
	(
		[AssignmentsID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[AssignmentsHistory] WITH NOCHECK ADD 
	CONSTRAINT [PK_AssignmentsHistory] PRIMARY KEY  CLUSTERED 
	(
		[AssignmentsHistoryID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Batches] WITH NOCHECK ADD 
	CONSTRAINT [PK_Batches] PRIMARY KEY  CLUSTERED 
	(
		[BatchesID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[BatchesHistory] WITH NOCHECK ADD 
	CONSTRAINT [PK_BatchesHistory] PRIMARY KEY  CLUSTERED 
	(
		[BatchesHistoryID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[BillAssignment] WITH NOCHECK ADD 
	CONSTRAINT [PK_BillAssignment] PRIMARY KEY  CLUSTERED 
	(
		[BillAssignmentID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[BillBillingCount] WITH NOCHECK ADD 
	CONSTRAINT [PK_BillBillingCount] PRIMARY KEY  CLUSTERED 
	(
		[BillBillingCountID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[BillingCount] WITH NOCHECK ADD 
	CONSTRAINT [PK_BillingCount] PRIMARY KEY  CLUSTERED 
	(
		[BillingCountID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[CAT] WITH NOCHECK ADD 
	CONSTRAINT [PK_CAT] PRIMARY KEY  CLUSTERED 
	(
		[CATID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[CATHistory] WITH NOCHECK ADD 
	CONSTRAINT [PK_CATHistory] PRIMARY KEY  CLUSTERED 
	(
		[CATHistoryID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[ClassOfLoss] WITH NOCHECK ADD 
	CONSTRAINT [PK_ClassOfLoss] PRIMARY KEY  CLUSTERED 
	(
		[ClassOfLossID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[ClassOfLossHistory] WITH NOCHECK ADD 
	CONSTRAINT [PK_ClassOfLossHistory] PRIMARY KEY  CLUSTERED 
	(
		[ClassOfLossHistoryID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[ClassType] WITH NOCHECK ADD 
	CONSTRAINT [PK_ClassType] PRIMARY KEY  CLUSTERED 
	(
		[ClassTypeID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[ClassTypeHistory] WITH NOCHECK ADD 
	CONSTRAINT [PK_ClassTypeHistory] PRIMARY KEY  CLUSTERED 
	(
		[ClassTypeHistoryID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Client] WITH NOCHECK ADD 
	CONSTRAINT [PK_Client] PRIMARY KEY  CLUSTERED 
	(
		[CompanyID],
		[UsersID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[ClientCoAdjusterSpec] WITH NOCHECK ADD 
	CONSTRAINT [PK_UserCompanySpec] PRIMARY KEY  CLUSTERED 
	(
		[ClientCoAdjusterSpecID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[ClientCoAdjusterSpecHistory] WITH NOCHECK ADD 
	CONSTRAINT [PK_AdjusterSpecHistory] PRIMARY KEY  CLUSTERED 
	(
		[AdjusterSpecHistoryID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[ClientCompanyCat] WITH NOCHECK ADD 
	CONSTRAINT [PK_CompanyCat] PRIMARY KEY  CLUSTERED 
	(
		[ClientCompanyID],
		[CATID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[ClientCompanyCatHistory] WITH NOCHECK ADD 
	CONSTRAINT [PK_CompanyCatHistory] PRIMARY KEY  CLUSTERED 
	(
		[ClientCompanyCatHistoryID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[ClientCompanyCatSpec] WITH NOCHECK ADD 
	CONSTRAINT [PK_CompanyCatSpec] PRIMARY KEY  CLUSTERED 
	(
		[ClientCompanyCatSpecID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[ClientCompanyCatSpecHistory] WITH NOCHECK ADD 
	CONSTRAINT [PK_CompanyCatSpecHistory] PRIMARY KEY  CLUSTERED 
	(
		[ClientCompanyCatSpecHistoryID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[ClientCompanyUsersCat] WITH NOCHECK ADD 
	CONSTRAINT [PK_UsrersCat] PRIMARY KEY  CLUSTERED 
	(
		[ClientCompanyID],
		[CATID],
		[UsersID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Company] WITH NOCHECK ADD 
	CONSTRAINT [PK_Company] PRIMARY KEY  CLUSTERED 
	(
		[CompanyID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[CompanyHistory] WITH NOCHECK ADD 
	CONSTRAINT [PK_CompanyHistory] PRIMARY KEY  CLUSTERED 
	(
		[CompanyHistoryID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[CompanyUsers] WITH NOCHECK ADD 
	CONSTRAINT [PK_CompanyUsers] PRIMARY KEY  CLUSTERED 
	(
		[CompanyID],
		[UsersID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[CompanyUsersHistory] WITH NOCHECK ADD 
	CONSTRAINT [PK_CompanyUsersHistory] PRIMARY KEY  CLUSTERED 
	(
		[CompanyUsersHistoryID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Coordinator] WITH NOCHECK ADD 
	CONSTRAINT [PK_Coordinator] PRIMARY KEY  CLUSTERED 
	(
		[CompanyID],
		[UsersID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Dispatcher] WITH NOCHECK ADD 
	CONSTRAINT [PK_Dispatcher] PRIMARY KEY  CLUSTERED 
	(
		[CompanyID],
		[UsersID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Document] WITH NOCHECK ADD 
	CONSTRAINT [PK_Document] PRIMARY KEY  CLUSTERED 
	(
		[DocumentID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[DocumentHistory] WITH NOCHECK ADD 
	CONSTRAINT [PK_DocumentHistory] PRIMARY KEY  CLUSTERED 
	(
		[DocumentID],
		[Version]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[ECSADJUsers] WITH NOCHECK ADD 
	CONSTRAINT [PK_ESCADJUsers] PRIMARY KEY  CLUSTERED 
	(
		[UsersID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Employee] WITH NOCHECK ADD 
	CONSTRAINT [PK_Employee] PRIMARY KEY  CLUSTERED 
	(
		[CompanyID],
		[UsersID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[FAQS] WITH NOCHECK ADD 
	CONSTRAINT [PK_FAQS] PRIMARY KEY  CLUSTERED 
	(
		[FAQSID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[FAQSHistory] WITH NOCHECK ADD 
	CONSTRAINT [PK_FAQSHistory] PRIMARY KEY  CLUSTERED 
	(
		[FAQSHistoryID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[FTPLog] WITH NOCHECK ADD 
	CONSTRAINT [PK_FTPLog] PRIMARY KEY  CLUSTERED 
	(
		[FTPLogID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[FTPLogArchive] WITH NOCHECK ADD 
	CONSTRAINT [PK_FTPLogArchive] PRIMARY KEY  CLUSTERED 
	(
		[FTPLogArchiveID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[FeeSchedule] WITH NOCHECK ADD 
	CONSTRAINT [PK_FeeSchedule] PRIMARY KEY  CLUSTERED 
	(
		[FeeScheduleID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[FeeScheduleFeeTypes] WITH NOCHECK ADD 
	CONSTRAINT [PK_FeeScheduleFeeTypes] PRIMARY KEY  CLUSTERED 
	(
		[FeeScheduleFeeTypesID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[FeeScheduleFeeTypesHistory] WITH NOCHECK ADD 
	CONSTRAINT [PK_FeeScheduleFeeTypesHistory] PRIMARY KEY  CLUSTERED 
	(
		[FeeScheduleFeeTypesHistoryID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[FeeScheduleHistory] WITH NOCHECK ADD 
	CONSTRAINT [PK_FeeScheduleHistory] PRIMARY KEY  CLUSTERED 
	(
		[FeeScheduleHistoryID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[FeeScheduleLevels] WITH NOCHECK ADD 
	CONSTRAINT [PK_FeeScheduleLevels] PRIMARY KEY  CLUSTERED 
	(
		[FeeScheduleLevelsID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[FeeScheduleLevelsHistory] WITH NOCHECK ADD 
	CONSTRAINT [PK_FeeScheduleLevelsHistory] PRIMARY KEY  CLUSTERED 
	(
		[FeeScheduleLevelsHistoryID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Group] WITH NOCHECK ADD 
	CONSTRAINT [PK_Group] PRIMARY KEY  CLUSTERED 
	(
		[GroupID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[HTTPLog] WITH NOCHECK ADD 
	CONSTRAINT [PK_HTTPLog] PRIMARY KEY  CLUSTERED 
	(
		[HTTPLogID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[HTTPLogArchive] WITH NOCHECK ADD 
	CONSTRAINT [PK_HTTPLogArchive] PRIMARY KEY  CLUSTERED 
	(
		[HTTPLogArchiveID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[IB] WITH NOCHECK ADD 
	CONSTRAINT [PK_IB] PRIMARY KEY  CLUSTERED 
	(
		[IBID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[IBFee] WITH NOCHECK ADD 
	CONSTRAINT [PK_IBFee] PRIMARY KEY  CLUSTERED 
	(
		[IBFeeID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[IBFeeHistory] WITH NOCHECK ADD 
	CONSTRAINT [PK_IBFeeHistory] PRIMARY KEY  CLUSTERED 
	(
		[IBFeeHistoryID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[IBHistory] WITH NOCHECK ADD 
	CONSTRAINT [PK_IBHistory] PRIMARY KEY  CLUSTERED 
	(
		[IBHistoryID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[IBStateFarm] WITH NOCHECK ADD 
	CONSTRAINT [PK_IBStateFarm] PRIMARY KEY  CLUSTERED 
	(
		[IBStateFarmID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Manager] WITH NOCHECK ADD 
	CONSTRAINT [PK_CompanyManagerUsers] PRIMARY KEY  CLUSTERED 
	(
		[CompanyID],
		[UsersID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[MiscReportParam] WITH NOCHECK ADD 
	CONSTRAINT [PK_MiscReportParam] PRIMARY KEY  CLUSTERED 
	(
		[MiscReportParamID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[MiscReportParam01] WITH NOCHECK ADD 
	CONSTRAINT [PK_MiscReportParam01] PRIMARY KEY  CLUSTERED 
	(
		[MiscReportParamID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[MiscReportParam02] WITH NOCHECK ADD 
	CONSTRAINT [PK_MiscReportParam02] PRIMARY KEY  CLUSTERED 
	(
		[MiscReportParamID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[MiscReportParam03] WITH NOCHECK ADD 
	CONSTRAINT [PK_MiscReportParam03] PRIMARY KEY  CLUSTERED 
	(
		[MiscReportParamID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[MiscReportParam04] WITH NOCHECK ADD 
	CONSTRAINT [PK_MiscReportParam04] PRIMARY KEY  CLUSTERED 
	(
		[MiscReportParamID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[MiscReportParam05] WITH NOCHECK ADD 
	CONSTRAINT [PK_MiscReportParam05] PRIMARY KEY  CLUSTERED 
	(
		[MiscReportParamID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[MiscReportParam06] WITH NOCHECK ADD 
	CONSTRAINT [PK_MiscReportParam06] PRIMARY KEY  CLUSTERED 
	(
		[MiscReportParamID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[MiscReportParam07] WITH NOCHECK ADD 
	CONSTRAINT [PK_MiscReportParam07] PRIMARY KEY  CLUSTERED 
	(
		[MiscReportParamID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[MiscReportParam08] WITH NOCHECK ADD 
	CONSTRAINT [PK_MiscReportParam08] PRIMARY KEY  CLUSTERED 
	(
		[MiscReportParamID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[MiscReportParam09] WITH NOCHECK ADD 
	CONSTRAINT [PK_MiscReportParam09] PRIMARY KEY  CLUSTERED 
	(
		[MiscReportParamID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[MiscReportParam10] WITH NOCHECK ADD 
	CONSTRAINT [PK_MiscReportParam10] PRIMARY KEY  CLUSTERED 
	(
		[MiscReportParamID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[MiscReportParam11] WITH NOCHECK ADD 
	CONSTRAINT [PK_MiscReportParam11] PRIMARY KEY  CLUSTERED 
	(
		[MiscReportParamID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[MiscReportParam12] WITH NOCHECK ADD 
	CONSTRAINT [PK_MiscReportParam12] PRIMARY KEY  CLUSTERED 
	(
		[MiscReportParamID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[MiscReportParam13] WITH NOCHECK ADD 
	CONSTRAINT [PK_MiscReportParam13] PRIMARY KEY  CLUSTERED 
	(
		[MiscReportParamID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[MiscReportParam14] WITH NOCHECK ADD 
	CONSTRAINT [PK_MiscReportParam14] PRIMARY KEY  CLUSTERED 
	(
		[MiscReportParamID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[MiscReportParam15] WITH NOCHECK ADD 
	CONSTRAINT [PK_MiscReportParam15] PRIMARY KEY  CLUSTERED 
	(
		[MiscReportParamID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[MiscReportParam16] WITH NOCHECK ADD 
	CONSTRAINT [PK_MiscReportParam16] PRIMARY KEY  CLUSTERED 
	(
		[MiscReportParamID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[MiscReportParam17] WITH NOCHECK ADD 
	CONSTRAINT [PK_MiscReportParam17] PRIMARY KEY  CLUSTERED 
	(
		[MiscReportParamID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[MiscReportParam18] WITH NOCHECK ADD 
	CONSTRAINT [PK_MiscReportParam18] PRIMARY KEY  CLUSTERED 
	(
		[MiscReportParamID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[MiscReportParam19] WITH NOCHECK ADD 
	CONSTRAINT [PK_MiscReportParam19] PRIMARY KEY  CLUSTERED 
	(
		[MiscReportParamID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[MiscReportParam20] WITH NOCHECK ADD 
	CONSTRAINT [PK_MiscReportParam20] PRIMARY KEY  CLUSTERED 
	(
		[MiscReportParamID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[MiscReportParam21] WITH NOCHECK ADD 
	CONSTRAINT [PK_MiscReportParam21] PRIMARY KEY  CLUSTERED 
	(
		[MiscReportParamID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[MiscReportParam22] WITH NOCHECK ADD 
	CONSTRAINT [PK_MiscReportParam22] PRIMARY KEY  CLUSTERED 
	(
		[MiscReportParamID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[MiscReportParam23] WITH NOCHECK ADD 
	CONSTRAINT [PK_MiscReportParam23] PRIMARY KEY  CLUSTERED 
	(
		[MiscReportParamID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[MiscReportParam24] WITH NOCHECK ADD 
	CONSTRAINT [PK_MiscReportParam24] PRIMARY KEY  CLUSTERED 
	(
		[MiscReportParamID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[MiscReportParam25] WITH NOCHECK ADD 
	CONSTRAINT [PK_MiscReportParam25] PRIMARY KEY  CLUSTERED 
	(
		[MiscReportParamID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[MiscReportParam26] WITH NOCHECK ADD 
	CONSTRAINT [PK_MiscReportParam26] PRIMARY KEY  CLUSTERED 
	(
		[MiscReportParamID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[MiscReportParam27] WITH NOCHECK ADD 
	CONSTRAINT [PK_MiscReportParam27] PRIMARY KEY  CLUSTERED 
	(
		[MiscReportParamID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[MiscReportParam28] WITH NOCHECK ADD 
	CONSTRAINT [PK_MiscReportParam28] PRIMARY KEY  CLUSTERED 
	(
		[MiscReportParamID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[MiscReportParam29] WITH NOCHECK ADD 
	CONSTRAINT [PK_MiscReportParam29] PRIMARY KEY  CLUSTERED 
	(
		[MiscReportParamID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[MiscReportParam30] WITH NOCHECK ADD 
	CONSTRAINT [PK_MiscReportParam30] PRIMARY KEY  CLUSTERED 
	(
		[MiscReportParamID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Package] WITH NOCHECK ADD 
	CONSTRAINT [PK_Package] PRIMARY KEY  CLUSTERED 
	(
		[PackageID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[PackageHistory] WITH NOCHECK ADD 
	CONSTRAINT [PK_PackageHistory] PRIMARY KEY  CLUSTERED 
	(
		[PackageHistoryID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[PackageItem] WITH NOCHECK ADD 
	CONSTRAINT [PK_PackageItem] PRIMARY KEY  CLUSTERED 
	(
		[PackageItemID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[PackageItemHistory] WITH NOCHECK ADD 
	CONSTRAINT [PK_PackageItemHistory] PRIMARY KEY  CLUSTERED 
	(
		[PackageItemHistoryID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[PolicyLimits] WITH NOCHECK ADD 
	CONSTRAINT [PK_PolicyLimits] PRIMARY KEY  CLUSTERED 
	(
		[PolicyLimitsID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[PolicyLimitsHistory] WITH NOCHECK ADD 
	CONSTRAINT [PK_PolicyLimitsHistory] PRIMARY KEY  CLUSTERED 
	(
		[PolicyLimitsHistoryID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[RTActivityLog] WITH NOCHECK ADD 
	CONSTRAINT [PK_RTActivityLog] PRIMARY KEY  CLUSTERED 
	(
		[RTActivityLogID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[RTActivityLogHistory] WITH NOCHECK ADD 
	CONSTRAINT [PK_RTActivityLogHistory] PRIMARY KEY  CLUSTERED 
	(
		[RTActivityLogHistoryID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[RTActivityLogInfo] WITH NOCHECK ADD 
	CONSTRAINT [PK_RTActivityLogInfo] PRIMARY KEY  CLUSTERED 
	(
		[AssignmentsID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[RTActivityLogInfoHistory] WITH NOCHECK ADD 
	CONSTRAINT [PK_RTActivityLogInfoHistory] PRIMARY KEY  CLUSTERED 
	(
		[RTActivityLogInfoHistoryID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[RTAttachments] WITH NOCHECK ADD 
	CONSTRAINT [PK_RTAttachments] PRIMARY KEY  CLUSTERED 
	(
		[RTAttachmentsID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[RTAttachmentsHistory] WITH NOCHECK ADD 
	CONSTRAINT [PK_RTAttachmentsHistory] PRIMARY KEY  CLUSTERED 
	(
		[RTAttachmentsHistoryID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[RTChecks] WITH NOCHECK ADD 
	CONSTRAINT [PK_RTChecks] PRIMARY KEY  CLUSTERED 
	(
		[RTChecksID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[RTChecksHistory] WITH NOCHECK ADD 
	CONSTRAINT [PK_RTChecksHistory] PRIMARY KEY  CLUSTERED 
	(
		[RTChecksHistoryID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[RTIB] WITH NOCHECK ADD 
	CONSTRAINT [PK_RTIB] PRIMARY KEY  CLUSTERED 
	(
		[AssignmentsID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[RTIBFee] WITH NOCHECK ADD 
	CONSTRAINT [PK_RTIBFee] PRIMARY KEY  CLUSTERED 
	(
		[RTIBFeeID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[RTIBFeeHistory] WITH NOCHECK ADD 
	CONSTRAINT [PK_RTIBFeeHistory] PRIMARY KEY  CLUSTERED 
	(
		[RTIBFeeHistoryID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[RTIBHistory] WITH NOCHECK ADD 
	CONSTRAINT [PK_RTIBHistory] PRIMARY KEY  CLUSTERED 
	(
		[RTIBHistoryID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[RTIndemnity] WITH NOCHECK ADD 
	CONSTRAINT [PK_RTIndemnity] PRIMARY KEY  CLUSTERED 
	(
		[RTIndemnityID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[RTIndemnityHistory] WITH NOCHECK ADD 
	CONSTRAINT [PK_RTIndemnityHistory] PRIMARY KEY  CLUSTERED 
	(
		[RTIndemnityHistoryID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[RTPhotoLog] WITH NOCHECK ADD 
	CONSTRAINT [PK_RTPhotoLog] PRIMARY KEY  CLUSTERED 
	(
		[RTPhotoLogID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[RTPhotoLogHistory] WITH NOCHECK ADD 
	CONSTRAINT [PK_RTPhotoLogHistory] PRIMARY KEY  CLUSTERED 
	(
		[RTPhotoLogHistoryID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[RTPhotoReport] WITH NOCHECK ADD 
	CONSTRAINT [PK_RTPhotoReport] PRIMARY KEY  CLUSTERED 
	(
		[RTPhotoReportID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[RTPhotoReportHistory] WITH NOCHECK ADD 
	CONSTRAINT [PK_RTPhotoReportHistory] PRIMARY KEY  CLUSTERED 
	(
		[RTPhotoReportHistoryID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[RTWSDiagram] WITH NOCHECK ADD 
	CONSTRAINT [PK_RTWSDiagram] PRIMARY KEY  CLUSTERED 
	(
		[RTWSDiagramID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[RTWSDiagramHistory] WITH NOCHECK ADD 
	CONSTRAINT [PK_RTWSDiagramHistory] PRIMARY KEY  CLUSTERED 
	(
		[RTWSDiagramHistoryID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[RegSetting] WITH NOCHECK ADD 
	CONSTRAINT [PK_RegSetting] PRIMARY KEY  CLUSTERED 
	(
		[RegSettingID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[RegSettingHistory] WITH NOCHECK ADD 
	CONSTRAINT [PK_RegSettingHistory] PRIMARY KEY  CLUSTERED 
	(
		[RegSettingID],
		[Version]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[SASecurityPackage] WITH NOCHECK ADD 
	CONSTRAINT [PK_SASecirityPackage] PRIMARY KEY  CLUSTERED 
	(
		[SecurityAreaID],
		[SecurityPackageID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[SF_AdjusterEvaluations] WITH NOCHECK ADD 
	CONSTRAINT [PK_SF_AdjusterEvaluations] PRIMARY KEY  CLUSTERED 
	(
		[EvaluationID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[SISecurityArea] WITH NOCHECK ADD 
	CONSTRAINT [PK_SISecurityArea] PRIMARY KEY  CLUSTERED 
	(
		[SecurityItemsID],
		[SecurityAreaID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[SPSecurity] WITH NOCHECK ADD 
	CONSTRAINT [PK_SPSecurity] PRIMARY KEY  CLUSTERED 
	(
		[SecurityPackageID],
		[SecurityID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Security] WITH NOCHECK ADD 
	CONSTRAINT [PK_Security] PRIMARY KEY  CLUSTERED 
	(
		[SecurityID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[SecurityArea] WITH NOCHECK ADD 
	CONSTRAINT [PK_SecurityArea] PRIMARY KEY  CLUSTERED 
	(
		[SecurityAreaID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[SecurityAreaType] WITH NOCHECK ADD 
	CONSTRAINT [PK_SecurityAreaType] PRIMARY KEY  CLUSTERED 
	(
		[SecurityAreaTypeID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[SecurityGroup] WITH NOCHECK ADD 
	CONSTRAINT [PK_SecurityGroup] PRIMARY KEY  CLUSTERED 
	(
		[SecurityID],
		[GroupID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[SecurityItemType] WITH NOCHECK ADD 
	CONSTRAINT [PK_SecurityItemType] PRIMARY KEY  CLUSTERED 
	(
		[SecurityItemTypeID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[SecurityItems] WITH NOCHECK ADD 
	CONSTRAINT [PK_SecurityItems] PRIMARY KEY  CLUSTERED 
	(
		[SecurityItemsID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[SecurityLevel] WITH NOCHECK ADD 
	CONSTRAINT [PK_SecurityLevel] PRIMARY KEY  CLUSTERED 
	(
		[SecurityLevel]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[SecurityLevelHistory] WITH NOCHECK ADD 
	CONSTRAINT [PK_SecurityLevelHistory] PRIMARY KEY  CLUSTERED 
	(
		[SecurityLevelHistoryID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[SecurityPackage] WITH NOCHECK ADD 
	CONSTRAINT [PK_SecurityPackage] PRIMARY KEY  CLUSTERED 
	(
		[SecurityPackageID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[SoftwarePackage] WITH NOCHECK ADD 
	CONSTRAINT [PK_SoftwarePackage] PRIMARY KEY  CLUSTERED 
	(
		[SoftWarePackageID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[SoftwarePackageApplication] WITH NOCHECK ADD 
	CONSTRAINT [PK_SoftwarePackageApplication] PRIMARY KEY  CLUSTERED 
	(
		[ApplicationID],
		[SoftWarePackageID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[SoftwarePackageDocument] WITH NOCHECK ADD 
	CONSTRAINT [PK_SoftwarePackageDocument] PRIMARY KEY  CLUSTERED 
	(
		[DocumentID],
		[SoftWarePackageID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[SoftwarePackageHistory] WITH NOCHECK ADD 
	CONSTRAINT [PK_SoftwarePackageHistory] PRIMARY KEY  CLUSTERED 
	(
		[SoftWarePackageHistoryID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[SoftwarePackageRegSetting] WITH NOCHECK ADD 
	CONSTRAINT [PK_SoftwarePackageRegSetting] PRIMARY KEY  CLUSTERED 
	(
		[RegSettingID],
		[SoftWarePackageID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[State] WITH NOCHECK ADD 
	CONSTRAINT [PK_State] PRIMARY KEY  CLUSTERED 
	(
		[StateID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[StateHistory] WITH NOCHECK ADD 
	CONSTRAINT [PK_StateHistory] PRIMARY KEY  CLUSTERED 
	(
		[StateHistoryID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Status] WITH NOCHECK ADD 
	CONSTRAINT [PK_AssStatus] PRIMARY KEY  CLUSTERED 
	(
		[StatusID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[StatusHistory] WITH NOCHECK ADD 
	CONSTRAINT [PK_StatusHistory] PRIMARY KEY  CLUSTERED 
	(
		[StatusHistoryID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Temporary] WITH NOCHECK ADD 
	CONSTRAINT [PK_Temporary] PRIMARY KEY  CLUSTERED 
	(
		[CompanyID],
		[UsersID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[TransType] WITH NOCHECK ADD 
	CONSTRAINT [PK_TransType] PRIMARY KEY  CLUSTERED 
	(
		[TransTypeID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[TypeOfLoss] WITH NOCHECK ADD 
	CONSTRAINT [PK_AssTypeOfLoss] PRIMARY KEY  CLUSTERED 
	(
		[TypeOfLossID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[TypeOfLossHistory] WITH NOCHECK ADD 
	CONSTRAINT [PK_TypeOfLossHistory] PRIMARY KEY  CLUSTERED 
	(
		[TypeOfLossHistoryID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[UserProfile] WITH NOCHECK ADD 
	CONSTRAINT [PK_UserProfile] PRIMARY KEY  CLUSTERED 
	(
		[UserProfileID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[UserProfileHistory] WITH NOCHECK ADD 
	CONSTRAINT [PK_UserProfileHistory] PRIMARY KEY  CLUSTERED 
	(
		[UserProfileHistoryID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[UserReportsToCoordinator] WITH NOCHECK ADD 
	CONSTRAINT [PK_UserReportsToCoordinator] PRIMARY KEY  CLUSTERED 
	(
		[UsersID],
		[CompanyID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[UserReportsToManager] WITH NOCHECK ADD 
	CONSTRAINT [PK_UserReportsToManager] PRIMARY KEY  CLUSTERED 
	(
		[UsersID],
		[CompanyID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Users] WITH NOCHECK ADD 
	CONSTRAINT [PK_Users] PRIMARY KEY  CLUSTERED 
	(
		[UsersID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[UsersGroup] WITH NOCHECK ADD 
	CONSTRAINT [PK_UserGroup] PRIMARY KEY  CLUSTERED 
	(
		[UsersID],
		[GroupID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[UsersHistory] WITH NOCHECK ADD 
	CONSTRAINT [PK_UsersHistory] PRIMARY KEY  CLUSTERED 
	(
		[UsersHistoryID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Work_ListText] WITH NOCHECK ADD 
	CONSTRAINT [PK_Work_ListText] PRIMARY KEY  CLUSTERED 
	(
		[Work_ListTextID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[XML_Trans] WITH NOCHECK ADD 
	CONSTRAINT [PK_XML_Trans] PRIMARY KEY  CLUSTERED 
	(
		[XML_TransID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Adjuster] ADD 
	CONSTRAINT [DF_Adjuster_MaxOpenAssgn] DEFAULT (0) FOR [MaxOpenAssgn],
	CONSTRAINT [DF_Adjuster_HomeBaseZip] DEFAULT (0) FOR [HomeBaseZip],
	CONSTRAINT [DF_Adjuster_MaxRangeFromHomeBaseZip] DEFAULT (0) FOR [MaxRangeFromHomeBaseZip],
	CONSTRAINT [DF_Adjuster_Active] DEFAULT (1) FOR [Active]
GO

ALTER TABLE [dbo].[AdjusterUsersSoftware] ADD 
	CONSTRAINT [DF_UsersSoftware_LicenseDaysLeft] DEFAULT (0) FOR [LicenseDaysLeft],
	CONSTRAINT [DF_UsersSoftware_ResetLicense] DEFAULT (0) FOR [ResetLicense],
	CONSTRAINT [DF_AdjusterUsersSoftware_IBPrefix] DEFAULT ('AA') FOR [IBPrefix],
	CONSTRAINT [DF_UsersSoftware_ResetIBPrefix] DEFAULT (0) FOR [ResetIBPrefix],
	CONSTRAINT [DF_AdjusterUsersSoftware_SingleFileSendAuthority] DEFAULT (0) FOR [SingleFileSendAuthority],
	CONSTRAINT [IX_AdjusterUsersSoftware] UNIQUE  NONCLUSTERED 
	(
		[IBPrefix]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[AdjusterUsersSoftwareHistory] ADD 
	CONSTRAINT [DF_UsersSoftwareHistory_LicenseDaysLeft] DEFAULT (0) FOR [LicenseDaysLeft],
	CONSTRAINT [DF_UsersSoftwareHistory_ResetLicense] DEFAULT (0) FOR [ResetLicense],
	CONSTRAINT [DF_UsersSoftwareHistory_ResetIBPrefix] DEFAULT (0) FOR [ResetIBPrefix],
	CONSTRAINT [DF_AdjusterUsersSoftwareHistory_SingleFileSendAuthority] DEFAULT (0) FOR [SingleFileSendAuthority]
GO

ALTER TABLE [dbo].[Admin] ADD 
	CONSTRAINT [DF_Admin_Active] DEFAULT (1) FOR [Active]
GO

ALTER TABLE [dbo].[Application] ADD 
	CONSTRAINT [DF_Application_Version] DEFAULT (1) FOR [Version],
	CONSTRAINT [DF_Application_SectionLevel01] DEFAULT ('') FOR [SectionLevel01],
	CONSTRAINT [DF_Application_SectionLevel02] DEFAULT ('') FOR [SectionLevel02],
	CONSTRAINT [DF_Application_SectionLevel03] DEFAULT ('') FOR [SectionLevel03],
	CONSTRAINT [DF_Application_SectionLevel04] DEFAULT ('') FOR [SectionLevel04],
	CONSTRAINT [DF_Application_SectionLevel05] DEFAULT ('') FOR [SectionLevel05],
	CONSTRAINT [DF_Application_SELF_REG] DEFAULT (1) FOR [SELF_REG],
	CONSTRAINT [DF_Application_IsDeleted] DEFAULT (0) FOR [IsDeleted],
	CONSTRAINT [IX_Application] UNIQUE  NONCLUSTERED 
	(
		[AppName]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[ApplicationHistory] ADD 
	CONSTRAINT [DF_ApplicationHistory_Version] DEFAULT (1) FOR [Version],
	CONSTRAINT [DF_ApplicationHistory_SectionLevel01] DEFAULT ('') FOR [SectionLevel01],
	CONSTRAINT [DF_ApplicationHistory_SectionLevel02] DEFAULT ('') FOR [SectionLevel02],
	CONSTRAINT [DF_ApplicationHistory_SectionLevel03] DEFAULT ('') FOR [SectionLevel03],
	CONSTRAINT [DF_ApplicationHistory_SectionLevel04] DEFAULT ('') FOR [SectionLevel04],
	CONSTRAINT [DF_ApplicationHistory_SectionLevel05] DEFAULT ('') FOR [SectionLevel05],
	CONSTRAINT [DF_ApplicationHistory_SELF_REG] DEFAULT (1) FOR [SELF_REG],
	CONSTRAINT [DF_ApplicationHistory_IsDeleted] DEFAULT (0) FOR [IsDeleted],
	CONSTRAINT [IX_ApplicationHistory] UNIQUE  NONCLUSTERED 
	(
		[ApplicationID],
		[Version]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[AssignmentType] ADD 
	CONSTRAINT [DF_AssignmentType_Description] DEFAULT ('') FOR [Description],
	CONSTRAINT [DF_AssignmentType_AdminComments] DEFAULT ('') FOR [AdminComments],
	CONSTRAINT [DF_AssignmentType_IsDeleted] DEFAULT (0) FOR [IsDeleted],
	CONSTRAINT [IX_AssignmentType] UNIQUE  NONCLUSTERED 
	(
		[Type]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[AssignmentTypeHistory] ADD 
	CONSTRAINT [DF_AssignmentTypeHistory_Description] DEFAULT ('') FOR [Description],
	CONSTRAINT [DF_AssignmentTypeHistory_AdminComments] DEFAULT ('') FOR [AdminComments],
	CONSTRAINT [DF_AssignmentTypeHistory_IsDeleted] DEFAULT (0) FOR [IsDeleted]
GO

ALTER TABLE [dbo].[Assignments] ADD 
	CONSTRAINT [DF_Assignments_AssignmentTypeID] DEFAULT (1) FOR [AssignmentTypeID],
	CONSTRAINT [DF_Assignments_DownLoadLossReport] DEFAULT (0) FOR [DownLoadLossReport],
	CONSTRAINT [DF_Assignments_UploadLossReport] DEFAULT (0) FOR [UploadLossReport],
	CONSTRAINT [DF_Assignments_SentToXact] DEFAULT (0) FOR [SentToXact],
	CONSTRAINT [DF_Assignments_IsLocked] DEFAULT (0) FOR [IsLocked],
	CONSTRAINT [DF_Assignments_IsDeleted] DEFAULT (0) FOR [IsDeleted],
	CONSTRAINT [DF_Assignments_DownLoadMe] DEFAULT (0) FOR [DownLoadMe],
	CONSTRAINT [DF_Assignments_UploadMe] DEFAULT (0) FOR [UpLoadMe],
	CONSTRAINT [DF_Assignments_DownLoadAll] DEFAULT (0) FOR [DownLoadAll],
	CONSTRAINT [DF_Assignments_UpLoadAll] DEFAULT (0) FOR [UpLoadAll],
	CONSTRAINT [DF_Assignments_MiscDelimSettings] DEFAULT ('') FOR [MiscDelimSettings],
	CONSTRAINT [IX_Assignments] UNIQUE  NONCLUSTERED 
	(
		[ClientCompanyCatSpecID],
		[AdjusterSpecID],
		[CLIENTNUM]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] ,
	CONSTRAINT [IX_Assignments_IBNUM] UNIQUE  NONCLUSTERED 
	(
		[IBNUM]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

 CREATE  INDEX [IX_Assignments_DateLastUpdated] ON [dbo].[Assignments]([DateLastUpdated], [CLIENTNUM], [IBNUM]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IX_Assignments_1] ON [dbo].[Assignments]([AssignmentsID]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

ALTER TABLE [dbo].[AssignmentsHistory] ADD 
	CONSTRAINT [DF_AssignmentsHistory_AssignmentTypeID] DEFAULT (1) FOR [AssignmentTypeID],
	CONSTRAINT [DF_AssignmentsHistory_DownLoadLossReport] DEFAULT (0) FOR [DownLoadLossReport],
	CONSTRAINT [DF_AssignmentsHistory_UploadLossReport] DEFAULT (0) FOR [UploadLossReport],
	CONSTRAINT [DF_AssignmentsHistory_SentToXact] DEFAULT (0) FOR [SentToXact],
	CONSTRAINT [DF_AssignmentsHistory_IsLocked] DEFAULT (0) FOR [IsLocked],
	CONSTRAINT [DF_AssignmentsHistory_IsDeleted] DEFAULT (0) FOR [IsDeleted],
	CONSTRAINT [DF_AssignmentsHistory_DownLoadMe] DEFAULT (0) FOR [DownLoadMe],
	CONSTRAINT [DF_AssignmentsHistory_UploadMe] DEFAULT (0) FOR [UpLoadMe],
	CONSTRAINT [DF_AssignmentsHistory_DownLoadAll] DEFAULT (0) FOR [DownLoadAll],
	CONSTRAINT [DF_AssignmentsHistory_UpLoadAll] DEFAULT (0) FOR [UpLoadAll],
	CONSTRAINT [DF_AssignmentsHistory_MiscDelimSettings] DEFAULT ('') FOR [MiscDelimSettings]
GO

ALTER TABLE [dbo].[Batches] ADD 
	CONSTRAINT [DF_Batches_Void] DEFAULT (0) FOR [Void],
	CONSTRAINT [DF_Batches_Reassigned_1] DEFAULT (0) FOR [Reassigned],
	CONSTRAINT [IX_Batches_ibnumber] UNIQUE  NONCLUSTERED 
	(
		[ibnumber]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[BatchesHistory] ADD 
	CONSTRAINT [DF_BatchesHistory_Void] DEFAULT (0) FOR [Void],
	CONSTRAINT [DF_BatchesHistory_Reassigned] DEFAULT (0) FOR [Reassigned]
GO

ALTER TABLE [dbo].[BillAssignment] ADD 
	CONSTRAINT [DF_BillAssignment_AssignmentTypeID] DEFAULT (1) FOR [AssignmentTypeID],
	CONSTRAINT [DF_BillAssignment_IsLocked] DEFAULT (0) FOR [IsLocked],
	CONSTRAINT [DF_BillAssignment_IsDeleted] DEFAULT (0) FOR [IsDeleted],
	CONSTRAINT [DF_BillAssignment_MiscDelimSettings] DEFAULT ('') FOR [MiscDelimSettings],
	CONSTRAINT [IX_BillAssignment] UNIQUE  NONCLUSTERED 
	(
		[ClientCompanyCatSpecID],
		[AdjusterSpecID],
		[CLIENTNUM]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] ,
	CONSTRAINT [IX_BillAssignment_IBNUM] UNIQUE  NONCLUSTERED 
	(
		[IBNUM]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

 CREATE  INDEX [IX_BillAssignment_DateLastUpdated] ON [dbo].[BillAssignment]([DateLastUpdated], [CLIENTNUM], [IBNUM]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IX_BillAssignment_1] ON [dbo].[BillAssignment]([BillAssignmentID]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

ALTER TABLE [dbo].[BillBillingCount] ADD 
	CONSTRAINT [DF_BillBillingCount_Rebill] DEFAULT (0) FOR [Rebill],
	CONSTRAINT [DF_BillBillingCount_Supplement] DEFAULT (0) FOR [Supplement]
GO

ALTER TABLE [dbo].[BillingCount] ADD 
	CONSTRAINT [DF_BillingCount_Rebill] DEFAULT (0) FOR [Rebill],
	CONSTRAINT [DF_BillingCount_Supplement] DEFAULT (0) FOR [Supplement],
	CONSTRAINT [DF_BillingCount_DownLoadMe] DEFAULT (0) FOR [DownLoadMe],
	CONSTRAINT [DF_BillingCount_UpLoadMe] DEFAULT (0) FOR [UpLoadMe],
	CONSTRAINT [IX_AssignmentID_Supplement] UNIQUE  NONCLUSTERED 
	(
		[AssignmentsID],
		[Supplement]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[CAT] ADD 
	CONSTRAINT [DF_CAT_Description] DEFAULT ('') FOR [Description]
GO

 CREATE  INDEX [IX_CAT] ON [dbo].[CAT]([Name]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

ALTER TABLE [dbo].[ClassOfLoss] ADD 
	CONSTRAINT [DF_ClassOfLoss_IsDeleted] DEFAULT (0) FOR [IsDeleted],
	CONSTRAINT [IX_ClassOfLoss_CompanyID_ClassTypeID] UNIQUE  NONCLUSTERED 
	(
		[ClientCompanyID],
		[ClassTypeID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[ClassOfLossHistory] ADD 
	CONSTRAINT [DF_ClassOfLossHistory_IsDeleted] DEFAULT (0) FOR [IsDeleted]
GO

ALTER TABLE [dbo].[ClassType] ADD 
	CONSTRAINT [DF_ClassType_IsDeleted] DEFAULT (0) FOR [IsDeleted],
	CONSTRAINT [IX_ClassType] UNIQUE  NONCLUSTERED 
	(
		[Class]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[ClassTypeHistory] ADD 
	CONSTRAINT [DF_ClassTypeHistory_IsDeleted] DEFAULT (0) FOR [IsDeleted]
GO

ALTER TABLE [dbo].[Client] ADD 
	CONSTRAINT [DF_Client_Active] DEFAULT (1) FOR [Active]
GO

ALTER TABLE [dbo].[ClientCoAdjusterSpec] ADD 
	CONSTRAINT [DF_AdjusterSpec_ActiveDate] DEFAULT (getdate()) FOR [ActiveDate],
	CONSTRAINT [IX_AdjusterSpec] UNIQUE  NONCLUSTERED 
	(
		[ClientCompanyID],
		[ACID],
		[ActiveDate]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[ClientCoAdjusterSpecHistory] ADD 
	CONSTRAINT [DF_AdjusterSpecHistory_ActiveDate] DEFAULT (getdate()) FOR [ActiveDate]
GO

ALTER TABLE [dbo].[ClientCompanyCat] ADD 
	CONSTRAINT [DF_ClientCompanyCat_AssignByZipDefault] DEFAULT (0) FOR [AssignByZipDefault]
GO

ALTER TABLE [dbo].[ClientCompanyCatHistory] ADD 
	CONSTRAINT [DF_ClientCompanyCatHistory_AssignByZipDefault] DEFAULT (0) FOR [AssignByZipDefault]
GO

ALTER TABLE [dbo].[ClientCompanyCatSpec] ADD 
	CONSTRAINT [DF_CompanyCatSpec_ActiveDate] DEFAULT (getdate()) FOR [ActiveDate],
	CONSTRAINT [IX_CompanyCatSpec] UNIQUE  NONCLUSTERED 
	(
		[ClientCompanyID],
		[CatCode],
		[ActiveDate]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[ClientCompanyCatSpecHistory] ADD 
	CONSTRAINT [DF_CompanyCatSpecHistory_ActiveDate] DEFAULT (getdate()) FOR [ActiveDate]
GO

ALTER TABLE [dbo].[ClientCompanyUsersCat] ADD 
	CONSTRAINT [DF_ClientCompanyUsersCat_Active] DEFAULT (1) FOR [Active]
GO

ALTER TABLE [dbo].[Company] ADD 
	CONSTRAINT [DF_Company_DBName] DEFAULT ('') FOR [DBName],
	CONSTRAINT [DF_Company_CarrierPrefix] DEFAULT ('') FOR [CarrierPrefix],
	CONSTRAINT [DF_Company_EnableSingleFile] DEFAULT (0) FOR [EnableSingleFile],
	CONSTRAINT [DF_Company_PDFJpegQuality] DEFAULT (100) FOR [PDFJpegQuality],
	CONSTRAINT [DF_Company_Active] DEFAULT (1) FOR [Active],
	CONSTRAINT [DF_Company_LogoImageName] DEFAULT ('') FOR [LogoImageName],
	CONSTRAINT [IX_Company] UNIQUE  NONCLUSTERED 
	(
		[Code]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[CompanyHistory] ADD 
	CONSTRAINT [DF_CompanyHistory_DBName] DEFAULT ('') FOR [DBName],
	CONSTRAINT [DF_CompanyHistory_CarrierPrefix] DEFAULT ('') FOR [CarrierPrefix],
	CONSTRAINT [DF_CompanyHistory_EnableSingleFile] DEFAULT (0) FOR [EnableSingleFile],
	CONSTRAINT [DF_CompanyHistory_PDFJpegQuality] DEFAULT (100) FOR [PDFJpegQuality],
	CONSTRAINT [DF_CompanyHistory_Active] DEFAULT (1) FOR [Active],
	CONSTRAINT [DF_CompanyHistory_LogoImageName] DEFAULT ('') FOR [LogoImageName]
GO

ALTER TABLE [dbo].[CompanyUsers] ADD 
	CONSTRAINT [DF_CompanyUsers_Active] DEFAULT (1) FOR [Active],
	CONSTRAINT [DF_CompanyUsers_ActiveDate] DEFAULT (getdate()) FOR [ActiveDate],
	CONSTRAINT [DF_CompanyUsers_SecurityLevel] DEFAULT (1) FOR [SecurityLevel],
	CONSTRAINT [DF_CompanyUsers_Flag] DEFAULT (0) FOR [Flag]
GO

ALTER TABLE [dbo].[CompanyUsersHistory] ADD 
	CONSTRAINT [DF_CompanyUsersHistory_Active] DEFAULT (1) FOR [Active],
	CONSTRAINT [DF_CompanyUsersHistory_SecurityLevel] DEFAULT (1) FOR [SecurityLevel],
	CONSTRAINT [DF_CompanyUsersHistory_Flag] DEFAULT (0) FOR [Flag]
GO

ALTER TABLE [dbo].[Coordinator] ADD 
	CONSTRAINT [DF_Coordinator_Active] DEFAULT (1) FOR [Active]
GO

ALTER TABLE [dbo].[DB_VERSION] ADD 
	CONSTRAINT [DF_DB_VERSION_Version] DEFAULT (1) FOR [Version],
	CONSTRAINT [DF_DB_VERSION_InstallFileLocation] DEFAULT ('') FOR [InstallFileLocation],
	CONSTRAINT [DF_DB_VERSION_SPName] DEFAULT ('') FOR [SPName],
	CONSTRAINT [DF_DB_VERSION_MainUtilInstallFileLocation] DEFAULT ('') FOR [MainUtilInstallFileLocation],
	CONSTRAINT [DF_DB_VERSION_MainUtilSPName] DEFAULT ('') FOR [MainUtilSPName],
	CONSTRAINT [DF_DB_VERSION_MainARVInstallFileLocation] DEFAULT ('') FOR [MainARVInstallFileLocation],
	CONSTRAINT [DF_DB_VERSION_MainARVSPName] DEFAULT ('') FOR [MainARVSPName],
	CONSTRAINT [DF_DB_VERSION_MainEXEInstallFileLocation] DEFAULT ('') FOR [MainEXEInstallFileLocation],
	CONSTRAINT [DF_DB_VERSION_MainEXESPName] DEFAULT ('') FOR [MainEXESPName],
	CONSTRAINT [DF_DB_VERSION_MainFTPEXEInstallFileLocation] DEFAULT ('') FOR [MainFTPEXEInstallFileLocation],
	CONSTRAINT [DF_DB_VERSION_MainFTPEXESPName] DEFAULT ('') FOR [MainFTPEXESPName]
GO

ALTER TABLE [dbo].[Dispatcher] ADD 
	CONSTRAINT [DF_Dispatcher_Active] DEFAULT (1) FOR [Active]
GO

ALTER TABLE [dbo].[Document] ADD 
	CONSTRAINT [DF_Document_Version] DEFAULT (1) FOR [Version],
	CONSTRAINT [DF_Document_SectionLevel01] DEFAULT ('') FOR [SectionLevel01],
	CONSTRAINT [DF_Document_SectionLevel02] DEFAULT ('') FOR [SectionLevel02],
	CONSTRAINT [DF_Document_SectionLevel03] DEFAULT ('') FOR [SectionLevel03],
	CONSTRAINT [DF_Document_SectionLevel04] DEFAULT ('') FOR [SectionLevel04],
	CONSTRAINT [DF_Document_SectionLevel05] DEFAULT ('') FOR [SectionLevel05],
	CONSTRAINT [DF_Document_IsDeleted] DEFAULT (0) FOR [IsDeleted],
	CONSTRAINT [IX_Document_1] UNIQUE  NONCLUSTERED 
	(
		[DocName]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[DocumentHistory] ADD 
	CONSTRAINT [DF_DocumentHistory_SectionLevel01] DEFAULT ('') FOR [SectionLevel01],
	CONSTRAINT [DF_DocumentHistory_SectionLevel02] DEFAULT ('') FOR [SectionLevel02],
	CONSTRAINT [DF_DocumentHistory_SectionLevel03] DEFAULT ('') FOR [SectionLevel03],
	CONSTRAINT [DF_DocumentHistory_SectionLevel04] DEFAULT ('') FOR [SectionLevel04],
	CONSTRAINT [DF_DocumentHistory_SectionLevel05] DEFAULT ('') FOR [SectionLevel05],
	CONSTRAINT [DF_DocumentHistory_IsDeleted] DEFAULT (0) FOR [IsDeleted],
	CONSTRAINT [IX_DocumentHistory] UNIQUE  NONCLUSTERED 
	(
		[DocumentID],
		[Version]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Employee] ADD 
	CONSTRAINT [DF_Employee_Active] DEFAULT (1) FOR [Active]
GO

ALTER TABLE [dbo].[FAQS] ADD 
	CONSTRAINT [DF_FAQS_IsDeleted] DEFAULT (0) FOR [IsDeleted]
GO

ALTER TABLE [dbo].[FAQSHistory] ADD 
	CONSTRAINT [DF_FAQSHistory_IsDeleted] DEFAULT (0) FOR [IsDeleted]
GO

 CREATE  INDEX [IX_FTPLog_LogTime] ON [dbo].[FTPLog]([LogTime]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IX_FTPLogArchive] ON [dbo].[FTPLogArchive]([LogTime]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

ALTER TABLE [dbo].[FeeSchedule] ADD 
	CONSTRAINT [DF_FeeSchedule_NumOfLevels] DEFAULT (0) FOR [NumOfLevels],
	CONSTRAINT [DF_FeeSchedule_NumOfFeeTypes] DEFAULT (0) FOR [NumOfFeeTypes],
	CONSTRAINT [DF_FeeSchedule_Fee10] DEFAULT (0) FOR [FeeServiceHourlyRate],
	CONSTRAINT [DF_FeeSchedule_Fee11] DEFAULT (0) FOR [TaxPercent],
	CONSTRAINT [DF_FeeSchedule_InitialOptions] DEFAULT ('') FOR [InitialOptions],
	CONSTRAINT [DF_FeeSchedule_Options] DEFAULT ('') FOR [Options],
	CONSTRAINT [DF_FeeSchedule_DefaultAppDedClassTypeIDOrder] DEFAULT ('1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26') FOR [DefaultAppDedClassTypeIDOrder],
	CONSTRAINT [DF_FeeSchedule_IsDeleted] DEFAULT (0) FOR [IsDeleted],
	CONSTRAINT [IX_FeeSchedule] UNIQUE  NONCLUSTERED 
	(
		[ClientCompanyID],
		[ScheduleName]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[FeeScheduleFeeTypes] ADD 
	CONSTRAINT [DF_FeeScheduleFeeTypes_TypeNum] DEFAULT (0) FOR [TypeNum],
	CONSTRAINT [DF_FeeScheduleFeeTypes_Description] DEFAULT ('') FOR [Description],
	CONSTRAINT [DF_FeeScheduleFeeTypes_FeeAount] DEFAULT (0.00) FOR [FeeAmount],
	CONSTRAINT [DF_FeeScheduleFeeTypes_IsExpense] DEFAULT (0) FOR [IsExpense],
	CONSTRAINT [DF_FeeScheduleFeeTypes_MaxNumberOfItems] DEFAULT (1) FOR [MaxNumberOfItems],
	CONSTRAINT [DF_FeeScheduleFeeTypes_MaxFeeAmount] DEFAULT (0) FOR [MaxFeeAmount],
	CONSTRAINT [DF_FeeScheduleFeeTypes_IsMiscAmount] DEFAULT (0) FOR [IsMiscAmount],
	CONSTRAINT [DF_FeeScheduleFeeTypes_UseFormula] DEFAULT (0) FOR [UseFormula],
	CONSTRAINT [DF_FeeScheduleFeeTypes_VBFormula] DEFAULT ('') FOR [VBFormula],
	CONSTRAINT [DF_FeeScheduleFeeTypes_IsDeleted] DEFAULT (0) FOR [IsDeleted]
GO

ALTER TABLE [dbo].[FeeScheduleFeeTypesHistory] ADD 
	CONSTRAINT [DF_FeeScheduleFeeTypesHistory_FeeAmount] DEFAULT (0.00) FOR [FeeAmount],
	CONSTRAINT [DF_FeeScheduleFeeTypesHistory_IsExpense] DEFAULT (0) FOR [IsExpense],
	CONSTRAINT [DF_FeeScheduleFeeTypesHistory_MaxFeeAmount] DEFAULT (0) FOR [MaxFeeAmount],
	CONSTRAINT [DF_FeeScheduleFeeTypesHistory_IsMiscAmount] DEFAULT (0) FOR [IsMiscAmount],
	CONSTRAINT [DF_FeeScheduleFeeTypesHistory_UseFormula] DEFAULT (0) FOR [UseFormula],
	CONSTRAINT [DF_FeeScheduleFeeTypesHistory_VBFormula] DEFAULT ('') FOR [VBFormula],
	CONSTRAINT [DF_FeeScheduleFeeTypesHistory_IsDeleted] DEFAULT (0) FOR [IsDeleted]
GO

ALTER TABLE [dbo].[FeeScheduleHistory] ADD 
	CONSTRAINT [DF_FeeScheduleHistory_NumOfLevels] DEFAULT (0) FOR [NumOfLevels],
	CONSTRAINT [DF_FeeScheduleHistory_FeeServiceHourlyRate] DEFAULT (0) FOR [FeeServiceHourlyRate],
	CONSTRAINT [DF_FeeScheduleHistory_TaxPercent] DEFAULT (0) FOR [TaxPercent],
	CONSTRAINT [DF_FeeScheduleHistory_InitialOptions] DEFAULT ('') FOR [InitialOptions],
	CONSTRAINT [DF_FeeScheduleHistory_Options] DEFAULT ('') FOR [Options],
	CONSTRAINT [DF_FeeScheduleHistory_DefaultAppDedClassTypeIDOrder] DEFAULT ('1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26') FOR [DefaultAppDedClassTypeIDOrder],
	CONSTRAINT [DF_FeeScheduleHistory_IsDeleted] DEFAULT (0) FOR [IsDeleted]
GO

ALTER TABLE [dbo].[FeeScheduleLevels] ADD 
	CONSTRAINT [DF_FeeScheduleLevels_LevelNum] DEFAULT (0) FOR [LevelNum],
	CONSTRAINT [DF_FeeScheduleLevels_IsDeleted] DEFAULT (0) FOR [IsDeleted]
GO

ALTER TABLE [dbo].[FeeScheduleLevelsHistory] ADD 
	CONSTRAINT [DF_FeeScheduleLevelsHistory_IsDeleted] DEFAULT (0) FOR [IsDeleted]
GO

 CREATE  INDEX [IX_HTTPLog_LogTime] ON [dbo].[HTTPLog]([LogTime]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IX_HTTPLogArchive] ON [dbo].[HTTPLogArchive]([LogTime]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

ALTER TABLE [dbo].[IB] ADD 
	CONSTRAINT [DF_IB_IB09a_sPolicyNo] DEFAULT ('') FOR [IB09a_sPolicyNo],
	CONSTRAINT [DF_IB_IB14a_sSupplement] DEFAULT (0) FOR [IB14a_sSupplement],
	CONSTRAINT [DF_IB_IB14b_sRebilled] DEFAULT (0) FOR [IB14b_sRebilled],
	CONSTRAINT [DF_IB_IB15c_cLessMiscellaneous] DEFAULT (0) FOR [IB15c_cLessMiscellaneous],
	CONSTRAINT [DF_IB_IB15d_cMiscellaneousDesc] DEFAULT ('') FOR [IB15d_cMiscellaneousDesc],
	CONSTRAINT [DF_IB_Void] DEFAULT (0) FOR [Void],
	CONSTRAINT [DF_IB_FeeByTime] DEFAULT (0) FOR [FeeByTime],
	CONSTRAINT [DF_IB_UseActivityTime] DEFAULT (0) FOR [UseActivityTime],
	CONSTRAINT [DF_IB_DownLoadMe] DEFAULT (0) FOR [DownLoadMe],
	CONSTRAINT [DF_IB_UploadMe] DEFAULT (0) FOR [UpLoadMe],
	CONSTRAINT [DF_IB_Comments] DEFAULT ('') FOR [Comments],
	CONSTRAINT [IX_IB_ibnumber] UNIQUE  NONCLUSTERED 
	(
		[BillingCountID],
		[IB02_sIBNumber]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[IBFee] ADD 
	CONSTRAINT [DF_IBFee_NumberOfItems] DEFAULT (1) FOR [NumberOfItems],
	CONSTRAINT [DF_IBFee_Amount] DEFAULT (0) FOR [Amount],
	CONSTRAINT [DF_IBFee_Comment] DEFAULT ('') FOR [Comment],
	CONSTRAINT [DF_IBFee_DownLoadMe] DEFAULT (0) FOR [DownLoadMe],
	CONSTRAINT [DF_IBFee_UpLoadMe] DEFAULT (0) FOR [UpLoadMe]
GO

ALTER TABLE [dbo].[IBFeeHistory] ADD 
	CONSTRAINT [DF_IBFeeHistory_DownLoadMe] DEFAULT (0) FOR [DownLoadMe],
	CONSTRAINT [DF_IBFeeHistory_UpLoadMe] DEFAULT (0) FOR [UpLoadMe]
GO

ALTER TABLE [dbo].[IBHistory] ADD 
	CONSTRAINT [DF_IBHistory_IB14a_sSupplement] DEFAULT (0) FOR [IB14a_sSupplement],
	CONSTRAINT [DF_IBHistory_IB14b_sRebilled] DEFAULT (0) FOR [IB14b_sRebilled],
	CONSTRAINT [DF_IBHistory_Void] DEFAULT (0) FOR [Void],
	CONSTRAINT [DF_IBHistory_FeeByTime] DEFAULT (0) FOR [FeeByTime],
	CONSTRAINT [DF_IBHistory_UseActivityTime] DEFAULT (0) FOR [UseActivityTime],
	CONSTRAINT [DF_IBHistory_DownLoadMe] DEFAULT (0) FOR [DownLoadMe],
	CONSTRAINT [DF_IBHistory_UploadMe] DEFAULT (0) FOR [UpLoadMe]
GO

ALTER TABLE [dbo].[IBStateFarm] ADD 
	CONSTRAINT [DF_IBStateFarm_Supplement] DEFAULT (0) FOR [Supplement],
	CONSTRAINT [DF_IBStateFarm_Rebilled] DEFAULT (0) FOR [Rebilled],
	CONSTRAINT [DF__StateFarm__Sever__2DDFDDD3] DEFAULT (1) FOR [SeverityCode],
	CONSTRAINT [DF__StateFarm__OutBu__2ED4020C] DEFAULT (0) FOR [OutBuildCount],
	CONSTRAINT [DF_IBStateFarm_Void] DEFAULT (0) FOR [Void],
	CONSTRAINT [DF_IBStateFarm_Comments] DEFAULT ('') FOR [Comments],
	CONSTRAINT [IX_IBStateFarm_IBNumber] UNIQUE  NONCLUSTERED 
	(
		[BillBillingCountID],
		[IBNumber]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Manager] ADD 
	CONSTRAINT [DF_Manager_Active] DEFAULT (1) FOR [Active]
GO

ALTER TABLE [dbo].[MiscReportParam] ADD 
	CONSTRAINT [DF_MiscReportParam_SortMe] DEFAULT ('') FOR [SortMe],
	CONSTRAINT [DF_MiscReportParam_IsDeleted] DEFAULT (0) FOR [IsDeleted],
	CONSTRAINT [DF_MiscReportParam_DownLoadMe] DEFAULT (0) FOR [DownLoadMe],
	CONSTRAINT [DF_MiscReportParam_UpLoadMe] DEFAULT (0) FOR [UpLoadMe],
	CONSTRAINT [IX_MiscReportParam_UniuqueParam] UNIQUE  NONCLUSTERED 
	(
		[AssignmentsID],
		[Number],
		[ProjectName],
		[ClassName],
		[ParamName]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[MiscReportParam01] ADD 
	CONSTRAINT [DF_MiscReportParam01_SortMe] DEFAULT ('') FOR [SortMe],
	CONSTRAINT [DF_MiscReportParam01_IsDeleted] DEFAULT (0) FOR [IsDeleted],
	CONSTRAINT [DF_MiscReportParam01_DownLoadMe] DEFAULT (0) FOR [DownLoadMe],
	CONSTRAINT [DF_MiscReportParam01_UpLoadMe] DEFAULT (0) FOR [UpLoadMe],
	CONSTRAINT [IX_MiscReportParam01_UniuqueParam] UNIQUE  NONCLUSTERED 
	(
		[AssignmentsID],
		[Number],
		[ProjectName],
		[ClassName],
		[ParamName]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[MiscReportParam02] ADD 
	CONSTRAINT [DF_MiscReportParam02_SortMe] DEFAULT ('') FOR [SortMe],
	CONSTRAINT [DF_MiscReportParam02_IsDeleted] DEFAULT (0) FOR [IsDeleted],
	CONSTRAINT [DF_MiscReportParam02_DownLoadMe] DEFAULT (0) FOR [DownLoadMe],
	CONSTRAINT [DF_MiscReportParam02_UpLoadMe] DEFAULT (0) FOR [UpLoadMe],
	CONSTRAINT [IX_MiscReportParam02_UniuqueParam] UNIQUE  NONCLUSTERED 
	(
		[AssignmentsID],
		[Number],
		[ProjectName],
		[ClassName],
		[ParamName]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[MiscReportParam03] ADD 
	CONSTRAINT [DF_MiscReportParam03_SortMe] DEFAULT ('') FOR [SortMe],
	CONSTRAINT [DF_MiscReportParam03_IsDeleted] DEFAULT (0) FOR [IsDeleted],
	CONSTRAINT [DF_MiscReportParam03_DownLoadMe] DEFAULT (0) FOR [DownLoadMe],
	CONSTRAINT [DF_MiscReportParam03_UpLoadMe] DEFAULT (0) FOR [UpLoadMe],
	CONSTRAINT [IX_MiscReportParam03_UniuqueParam] UNIQUE  NONCLUSTERED 
	(
		[AssignmentsID],
		[Number],
		[ProjectName],
		[ClassName],
		[ParamName]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[MiscReportParam04] ADD 
	CONSTRAINT [DF_MiscReportParam04_SortMe] DEFAULT ('') FOR [SortMe],
	CONSTRAINT [DF_MiscReportParam04_IsDeleted] DEFAULT (0) FOR [IsDeleted],
	CONSTRAINT [DF_MiscReportParam04_DownLoadMe] DEFAULT (0) FOR [DownLoadMe],
	CONSTRAINT [DF_MiscReportParam04_UpLoadMe] DEFAULT (0) FOR [UpLoadMe],
	CONSTRAINT [IX_MiscReportParam04_UniuqueParam] UNIQUE  NONCLUSTERED 
	(
		[AssignmentsID],
		[Number],
		[ProjectName],
		[ClassName],
		[ParamName]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[MiscReportParam05] ADD 
	CONSTRAINT [DF_MiscReportParam05_SortMe] DEFAULT ('') FOR [SortMe],
	CONSTRAINT [DF_MiscReportParam05_IsDeleted] DEFAULT (0) FOR [IsDeleted],
	CONSTRAINT [DF_MiscReportParam05_DownLoadMe] DEFAULT (0) FOR [DownLoadMe],
	CONSTRAINT [DF_MiscReportParam05_UpLoadMe] DEFAULT (0) FOR [UpLoadMe],
	CONSTRAINT [IX_MiscReportParam05_UniuqueParam] UNIQUE  NONCLUSTERED 
	(
		[AssignmentsID],
		[Number],
		[ProjectName],
		[ClassName],
		[ParamName]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[MiscReportParam06] ADD 
	CONSTRAINT [DF_MiscReportParam06_SortMe] DEFAULT ('') FOR [SortMe],
	CONSTRAINT [DF_MiscReportParam06_IsDeleted] DEFAULT (0) FOR [IsDeleted],
	CONSTRAINT [DF_MiscReportParam06_DownLoadMe] DEFAULT (0) FOR [DownLoadMe],
	CONSTRAINT [DF_MiscReportParam06_UpLoadMe] DEFAULT (0) FOR [UpLoadMe],
	CONSTRAINT [IX_MiscReportParam06_UniuqueParam] UNIQUE  NONCLUSTERED 
	(
		[AssignmentsID],
		[Number],
		[ProjectName],
		[ClassName],
		[ParamName]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[MiscReportParam07] ADD 
	CONSTRAINT [DF_MiscReportParam07_SortMe] DEFAULT ('') FOR [SortMe],
	CONSTRAINT [DF_MiscReportParam07_IsDeleted] DEFAULT (0) FOR [IsDeleted],
	CONSTRAINT [DF_MiscReportParam07_DownLoadMe] DEFAULT (0) FOR [DownLoadMe],
	CONSTRAINT [DF_MiscReportParam07_UpLoadMe] DEFAULT (0) FOR [UpLoadMe],
	CONSTRAINT [IX_MiscReportParam07_UniuqueParam] UNIQUE  NONCLUSTERED 
	(
		[AssignmentsID],
		[Number],
		[ProjectName],
		[ClassName],
		[ParamName]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[MiscReportParam08] ADD 
	CONSTRAINT [DF_MiscReportParam08_SortMe] DEFAULT ('') FOR [SortMe],
	CONSTRAINT [DF_MiscReportParam08_IsDeleted] DEFAULT (0) FOR [IsDeleted],
	CONSTRAINT [DF_MiscReportParam08_DownLoadMe] DEFAULT (0) FOR [DownLoadMe],
	CONSTRAINT [DF_MiscReportParam08_UpLoadMe] DEFAULT (0) FOR [UpLoadMe],
	CONSTRAINT [IX_MiscReportParam08_UniuqueParam] UNIQUE  NONCLUSTERED 
	(
		[AssignmentsID],
		[Number],
		[ProjectName],
		[ClassName],
		[ParamName]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[MiscReportParam09] ADD 
	CONSTRAINT [DF_MiscReportParam09_SortMe] DEFAULT ('') FOR [SortMe],
	CONSTRAINT [DF_MiscReportParam09_IsDeleted] DEFAULT (0) FOR [IsDeleted],
	CONSTRAINT [DF_MiscReportParam09_DownLoadMe] DEFAULT (0) FOR [DownLoadMe],
	CONSTRAINT [DF_MiscReportParam09_UpLoadMe] DEFAULT (0) FOR [UpLoadMe],
	CONSTRAINT [IX_MiscReportParam09_UniuqueParam] UNIQUE  NONCLUSTERED 
	(
		[AssignmentsID],
		[Number],
		[ProjectName],
		[ClassName],
		[ParamName]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[MiscReportParam10] ADD 
	CONSTRAINT [DF_MiscReportParam10_SortMe] DEFAULT ('') FOR [SortMe],
	CONSTRAINT [DF_MiscReportParam10_IsDeleted] DEFAULT (0) FOR [IsDeleted],
	CONSTRAINT [DF_MiscReportParam10_DownLoadMe] DEFAULT (0) FOR [DownLoadMe],
	CONSTRAINT [DF_MiscReportParam10_UpLoadMe] DEFAULT (0) FOR [UpLoadMe],
	CONSTRAINT [IX_MiscReportParam10_UniuqueParam] UNIQUE  NONCLUSTERED 
	(
		[AssignmentsID],
		[Number],
		[ProjectName],
		[ClassName],
		[ParamName]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[MiscReportParam11] ADD 
	CONSTRAINT [DF_MiscReportParam11_SortMe] DEFAULT ('') FOR [SortMe],
	CONSTRAINT [DF_MiscReportParam11_IsDeleted] DEFAULT (0) FOR [IsDeleted],
	CONSTRAINT [DF_MiscReportParam11_DownLoadMe] DEFAULT (0) FOR [DownLoadMe],
	CONSTRAINT [DF_MiscReportParam11_UpLoadMe] DEFAULT (0) FOR [UpLoadMe],
	CONSTRAINT [IX_MiscReportParam11_UniuqueParam] UNIQUE  NONCLUSTERED 
	(
		[AssignmentsID],
		[Number],
		[ProjectName],
		[ClassName],
		[ParamName]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[MiscReportParam12] ADD 
	CONSTRAINT [DF_MiscReportParam12_SortMe] DEFAULT ('') FOR [SortMe],
	CONSTRAINT [DF_MiscReportParam12_IsDeleted] DEFAULT (0) FOR [IsDeleted],
	CONSTRAINT [DF_MiscReportParam12_DownLoadMe] DEFAULT (0) FOR [DownLoadMe],
	CONSTRAINT [DF_MiscReportParam12_UpLoadMe] DEFAULT (0) FOR [UpLoadMe],
	CONSTRAINT [IX_MiscReportParam12_UniuqueParam] UNIQUE  NONCLUSTERED 
	(
		[AssignmentsID],
		[Number],
		[ProjectName],
		[ClassName],
		[ParamName]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[MiscReportParam13] ADD 
	CONSTRAINT [DF_MiscReportParam13_SortMe] DEFAULT ('') FOR [SortMe],
	CONSTRAINT [DF_MiscReportParam13_IsDeleted] DEFAULT (0) FOR [IsDeleted],
	CONSTRAINT [DF_MiscReportParam13_DownLoadMe] DEFAULT (0) FOR [DownLoadMe],
	CONSTRAINT [DF_MiscReportParam13_UpLoadMe] DEFAULT (0) FOR [UpLoadMe],
	CONSTRAINT [IX_MiscReportParam13_UniuqueParam] UNIQUE  NONCLUSTERED 
	(
		[AssignmentsID],
		[Number],
		[ProjectName],
		[ClassName],
		[ParamName]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[MiscReportParam14] ADD 
	CONSTRAINT [DF_MiscReportParam14_SortMe] DEFAULT ('') FOR [SortMe],
	CONSTRAINT [DF_MiscReportParam14_IsDeleted] DEFAULT (0) FOR [IsDeleted],
	CONSTRAINT [DF_MiscReportParam14_DownLoadMe] DEFAULT (0) FOR [DownLoadMe],
	CONSTRAINT [DF_MiscReportParam14_UpLoadMe] DEFAULT (0) FOR [UpLoadMe],
	CONSTRAINT [IX_MiscReportParam14_UniuqueParam] UNIQUE  NONCLUSTERED 
	(
		[AssignmentsID],
		[Number],
		[ProjectName],
		[ClassName],
		[ParamName]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[MiscReportParam15] ADD 
	CONSTRAINT [DF_MiscReportParam15_SortMe] DEFAULT ('') FOR [SortMe],
	CONSTRAINT [DF_MiscReportParam15_IsDeleted] DEFAULT (0) FOR [IsDeleted],
	CONSTRAINT [DF_MiscReportParam15_DownLoadMe] DEFAULT (0) FOR [DownLoadMe],
	CONSTRAINT [DF_MiscReportParam15_UpLoadMe] DEFAULT (0) FOR [UpLoadMe],
	CONSTRAINT [IX_MiscReportParam15_UniuqueParam] UNIQUE  NONCLUSTERED 
	(
		[AssignmentsID],
		[Number],
		[ProjectName],
		[ClassName],
		[ParamName]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[MiscReportParam16] ADD 
	CONSTRAINT [DF_MiscReportParam16_SortMe] DEFAULT ('') FOR [SortMe],
	CONSTRAINT [DF_MiscReportParam16_IsDeleted] DEFAULT (0) FOR [IsDeleted],
	CONSTRAINT [DF_MiscReportParam16_DownLoadMe] DEFAULT (0) FOR [DownLoadMe],
	CONSTRAINT [DF_MiscReportParam16_UpLoadMe] DEFAULT (0) FOR [UpLoadMe],
	CONSTRAINT [IX_MiscReportParam16_UniuqueParam] UNIQUE  NONCLUSTERED 
	(
		[AssignmentsID],
		[Number],
		[ProjectName],
		[ClassName],
		[ParamName]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[MiscReportParam17] ADD 
	CONSTRAINT [DF_MiscReportParam17_SortMe] DEFAULT ('') FOR [SortMe],
	CONSTRAINT [DF_MiscReportParam17_IsDeleted] DEFAULT (0) FOR [IsDeleted],
	CONSTRAINT [DF_MiscReportParam17_DownLoadMe] DEFAULT (0) FOR [DownLoadMe],
	CONSTRAINT [DF_MiscReportParam17_UpLoadMe] DEFAULT (0) FOR [UpLoadMe],
	CONSTRAINT [IX_MiscReportParam17_UniuqueParam] UNIQUE  NONCLUSTERED 
	(
		[AssignmentsID],
		[Number],
		[ProjectName],
		[ClassName],
		[ParamName]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[MiscReportParam18] ADD 
	CONSTRAINT [DF_MiscReportParam18_SortMe] DEFAULT ('') FOR [SortMe],
	CONSTRAINT [DF_MiscReportParam18_IsDeleted] DEFAULT (0) FOR [IsDeleted],
	CONSTRAINT [DF_MiscReportParam18_DownLoadMe] DEFAULT (0) FOR [DownLoadMe],
	CONSTRAINT [DF_MiscReportParam18_UpLoadMe] DEFAULT (0) FOR [UpLoadMe],
	CONSTRAINT [IX_MiscReportParam18_UniuqueParam] UNIQUE  NONCLUSTERED 
	(
		[AssignmentsID],
		[Number],
		[ProjectName],
		[ClassName],
		[ParamName]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[MiscReportParam19] ADD 
	CONSTRAINT [DF_MiscReportParam19_SortMe] DEFAULT ('') FOR [SortMe],
	CONSTRAINT [DF_MiscReportParam19_IsDeleted] DEFAULT (0) FOR [IsDeleted],
	CONSTRAINT [DF_MiscReportParam19_DownLoadMe] DEFAULT (0) FOR [DownLoadMe],
	CONSTRAINT [DF_MiscReportParam19_UpLoadMe] DEFAULT (0) FOR [UpLoadMe],
	CONSTRAINT [IX_MiscReportParam19_UniuqueParam] UNIQUE  NONCLUSTERED 
	(
		[AssignmentsID],
		[Number],
		[ProjectName],
		[ClassName],
		[ParamName]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[MiscReportParam20] ADD 
	CONSTRAINT [DF_MiscReportParam20_SortMe] DEFAULT ('') FOR [SortMe],
	CONSTRAINT [DF_MiscReportParam20_IsDeleted] DEFAULT (0) FOR [IsDeleted],
	CONSTRAINT [DF_MiscReportParam20_DownLoadMe] DEFAULT (0) FOR [DownLoadMe],
	CONSTRAINT [DF_MiscReportParam20_UpLoadMe] DEFAULT (0) FOR [UpLoadMe],
	CONSTRAINT [IX_MiscReportParam20_UniuqueParam] UNIQUE  NONCLUSTERED 
	(
		[AssignmentsID],
		[Number],
		[ProjectName],
		[ClassName],
		[ParamName]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[MiscReportParam21] ADD 
	CONSTRAINT [DF_MiscReportParam21_SortMe] DEFAULT ('') FOR [SortMe],
	CONSTRAINT [DF_MiscReportParam21_IsDeleted] DEFAULT (0) FOR [IsDeleted],
	CONSTRAINT [DF_MiscReportParam21_DownLoadMe] DEFAULT (0) FOR [DownLoadMe],
	CONSTRAINT [DF_MiscReportParam21_UpLoadMe] DEFAULT (0) FOR [UpLoadMe],
	CONSTRAINT [IX_MiscReportParam21_UniuqueParam] UNIQUE  NONCLUSTERED 
	(
		[AssignmentsID],
		[Number],
		[ProjectName],
		[ClassName],
		[ParamName]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[MiscReportParam22] ADD 
	CONSTRAINT [DF_MiscReportParam22_SortMe] DEFAULT ('') FOR [SortMe],
	CONSTRAINT [DF_MiscReportParam22_IsDeleted] DEFAULT (0) FOR [IsDeleted],
	CONSTRAINT [DF_MiscReportParam22_DownLoadMe] DEFAULT (0) FOR [DownLoadMe],
	CONSTRAINT [DF_MiscReportParam22_UpLoadMe] DEFAULT (0) FOR [UpLoadMe],
	CONSTRAINT [IX_MiscReportParam22_UniuqueParam] UNIQUE  NONCLUSTERED 
	(
		[AssignmentsID],
		[Number],
		[ProjectName],
		[ClassName],
		[ParamName]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[MiscReportParam23] ADD 
	CONSTRAINT [DF_MiscReportParam23_SortMe] DEFAULT ('') FOR [SortMe],
	CONSTRAINT [DF_MiscReportParam23_IsDeleted] DEFAULT (0) FOR [IsDeleted],
	CONSTRAINT [DF_MiscReportParam23_DownLoadMe] DEFAULT (0) FOR [DownLoadMe],
	CONSTRAINT [DF_MiscReportParam23_UpLoadMe] DEFAULT (0) FOR [UpLoadMe],
	CONSTRAINT [IX_MiscReportParam23_UniuqueParam] UNIQUE  NONCLUSTERED 
	(
		[AssignmentsID],
		[Number],
		[ProjectName],
		[ClassName],
		[ParamName]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[MiscReportParam24] ADD 
	CONSTRAINT [DF_MiscReportParam24_SortMe] DEFAULT ('') FOR [SortMe],
	CONSTRAINT [DF_MiscReportParam24_IsDeleted] DEFAULT (0) FOR [IsDeleted],
	CONSTRAINT [DF_MiscReportParam24_DownLoadMe] DEFAULT (0) FOR [DownLoadMe],
	CONSTRAINT [DF_MiscReportParam24_UpLoadMe] DEFAULT (0) FOR [UpLoadMe],
	CONSTRAINT [IX_MiscReportParam24_UniuqueParam] UNIQUE  NONCLUSTERED 
	(
		[AssignmentsID],
		[Number],
		[ProjectName],
		[ClassName],
		[ParamName]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[MiscReportParam25] ADD 
	CONSTRAINT [DF_MiscReportParam25_SortMe] DEFAULT ('') FOR [SortMe],
	CONSTRAINT [DF_MiscReportParam25_IsDeleted] DEFAULT (0) FOR [IsDeleted],
	CONSTRAINT [DF_MiscReportParam25_DownLoadMe] DEFAULT (0) FOR [DownLoadMe],
	CONSTRAINT [DF_MiscReportParam25_UpLoadMe] DEFAULT (0) FOR [UpLoadMe],
	CONSTRAINT [IX_MiscReportParam25_UniuqueParam] UNIQUE  NONCLUSTERED 
	(
		[AssignmentsID],
		[Number],
		[ProjectName],
		[ClassName],
		[ParamName]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[MiscReportParam26] ADD 
	CONSTRAINT [DF_MiscReportParam26_SortMe] DEFAULT ('') FOR [SortMe],
	CONSTRAINT [DF_MiscReportParam26_IsDeleted] DEFAULT (0) FOR [IsDeleted],
	CONSTRAINT [DF_MiscReportParam26_DownLoadMe] DEFAULT (0) FOR [DownLoadMe],
	CONSTRAINT [DF_MiscReportParam26_UpLoadMe] DEFAULT (0) FOR [UpLoadMe],
	CONSTRAINT [IX_MiscReportParam26_UniuqueParam] UNIQUE  NONCLUSTERED 
	(
		[AssignmentsID],
		[Number],
		[ProjectName],
		[ClassName],
		[ParamName]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[MiscReportParam27] ADD 
	CONSTRAINT [DF_MiscReportParam27_SortMe] DEFAULT ('') FOR [SortMe],
	CONSTRAINT [DF_MiscReportParam27_IsDeleted] DEFAULT (0) FOR [IsDeleted],
	CONSTRAINT [DF_MiscReportParam27_DownLoadMe] DEFAULT (0) FOR [DownLoadMe],
	CONSTRAINT [DF_MiscReportParam27_UpLoadMe] DEFAULT (0) FOR [UpLoadMe],
	CONSTRAINT [IX_MiscReportParam27_UniuqueParam] UNIQUE  NONCLUSTERED 
	(
		[AssignmentsID],
		[Number],
		[ProjectName],
		[ClassName],
		[ParamName]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[MiscReportParam28] ADD 
	CONSTRAINT [DF_MiscReportParam28_SortMe] DEFAULT ('') FOR [SortMe],
	CONSTRAINT [DF_MiscReportParam28_IsDeleted] DEFAULT (0) FOR [IsDeleted],
	CONSTRAINT [DF_MiscReportParam28_DownLoadMe] DEFAULT (0) FOR [DownLoadMe],
	CONSTRAINT [DF_MiscReportParam28_UpLoadMe] DEFAULT (0) FOR [UpLoadMe],
	CONSTRAINT [IX_MiscReportParam28_UniuqueParam] UNIQUE  NONCLUSTERED 
	(
		[AssignmentsID],
		[Number],
		[ProjectName],
		[ClassName],
		[ParamName]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[MiscReportParam29] ADD 
	CONSTRAINT [DF_MiscReportParam29_SortMe] DEFAULT ('') FOR [SortMe],
	CONSTRAINT [DF_MiscReportParam29_IsDeleted] DEFAULT (0) FOR [IsDeleted],
	CONSTRAINT [DF_MiscReportParam29_DownLoadMe] DEFAULT (0) FOR [DownLoadMe],
	CONSTRAINT [DF_MiscReportParam29_UpLoadMe] DEFAULT (0) FOR [UpLoadMe],
	CONSTRAINT [IX_MiscReportParam29_UniuqueParam] UNIQUE  NONCLUSTERED 
	(
		[AssignmentsID],
		[Number],
		[ProjectName],
		[ClassName],
		[ParamName]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[MiscReportParam30] ADD 
	CONSTRAINT [DF_MiscReportParam30_SortMe] DEFAULT ('') FOR [SortMe],
	CONSTRAINT [DF_MiscReportParam30_IsDeleted] DEFAULT (0) FOR [IsDeleted],
	CONSTRAINT [DF_MiscReportParam30_DownLoadMe] DEFAULT (0) FOR [DownLoadMe],
	CONSTRAINT [DF_MiscReportParam30_UpLoadMe] DEFAULT (0) FOR [UpLoadMe],
	CONSTRAINT [IX_MiscReportParam30_UniuqueParam] UNIQUE  NONCLUSTERED 
	(
		[AssignmentsID],
		[Number],
		[ProjectName],
		[ClassName],
		[ParamName]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Package] ADD 
	CONSTRAINT [DF_Package_PackageStatus] DEFAULT ('') FOR [PackageStatus],
	CONSTRAINT [DF_Package_Number] DEFAULT (1) FOR [Number],
	CONSTRAINT [DF_Package_SendMe] DEFAULT (0) FOR [SendMe],
	CONSTRAINT [DF_Package_IsDeleted] DEFAULT (0) FOR [IsDeleted],
	CONSTRAINT [DF_Package_DownLoadMe] DEFAULT (0) FOR [DownLoadMe],
	CONSTRAINT [DF_Package_UpLoadMe] DEFAULT (0) FOR [UpLoadMe]
GO

ALTER TABLE [dbo].[PackageHistory] ADD 
	CONSTRAINT [DF_PackageHistory_PackageStatus] DEFAULT ('') FOR [PackageStatus],
	CONSTRAINT [DF_PackageHistory_SendMe] DEFAULT (0) FOR [SendMe],
	CONSTRAINT [DF_PackageHistory_IsDeleted] DEFAULT (0) FOR [IsDeleted],
	CONSTRAINT [DF_PackageHistory_DownLoadMe] DEFAULT (0) FOR [DownLoadMe],
	CONSTRAINT [DF_PackageHistory_UpLoadMe] DEFAULT (0) FOR [UpLoadMe]
GO

ALTER TABLE [dbo].[PackageItem] ADD 
	CONSTRAINT [DF_PackageItem_ReportFormat] DEFAULT ('') FOR [ReportFormat],
	CONSTRAINT [DF_PackageItem_AttachmentName] DEFAULT ('') FOR [AttachmentName],
	CONSTRAINT [DF_PackageItem_IsCoApprove] DEFAULT (0) FOR [IsCoApprove],
	CONSTRAINT [DF_PackageItem_CoApproveDesc] DEFAULT ('') FOR [CoApproveDesc],
	CONSTRAINT [DF_PackageItem_IsClientCoReject] DEFAULT (0) FOR [IsClientCoReject],
	CONSTRAINT [DF_PackageItem_ClientCoRejectDesc] DEFAULT ('') FOR [ClientCoRejectDesc],
	CONSTRAINT [DF_PackageItem_IsClientCoDelete] DEFAULT (0) FOR [IsClientCoDelete],
	CONSTRAINT [DF_PackageItem_ClientCoDeleteDesc] DEFAULT ('') FOR [ClientCoDeleteDesc],
	CONSTRAINT [DF_PackageItem_IsClientCoApprove] DEFAULT (0) FOR [IsClientCoApprove],
	CONSTRAINT [DF_PackageItem_ClientCoApproveDesc] DEFAULT ('') FOR [ClientCoApproveDesc],
	CONSTRAINT [DF_PackageItem_PackageItemGUID] DEFAULT (newid()) FOR [PackageItemGUID],
	CONSTRAINT [DF_PackageItem_SendMe] DEFAULT (0) FOR [SendMe],
	CONSTRAINT [DF_PackageItem_IsDeleted] DEFAULT (0) FOR [IsDeleted],
	CONSTRAINT [DF_PackageItem_DownLoadMe] DEFAULT (0) FOR [DownLoadMe],
	CONSTRAINT [DF_PackageItem_UpLoadMe] DEFAULT (0) FOR [UpLoadMe]
GO

ALTER TABLE [dbo].[PackageItemHistory] ADD 
	CONSTRAINT [DF_PackageItemHistory_IsCoApprove] DEFAULT (0) FOR [IsCoApprove],
	CONSTRAINT [DF_PackageItemHistory_CoApproveDesc] DEFAULT ('') FOR [CoApproveDesc],
	CONSTRAINT [DF_PackageItemHistory_IsClientCoReject] DEFAULT (0) FOR [IsClientCoReject],
	CONSTRAINT [DF_PackageItemHistory_ClientCoRejectDesc] DEFAULT ('') FOR [ClientCoRejectDesc],
	CONSTRAINT [DF_PackageItemHistory_IsClientCoDelete] DEFAULT (0) FOR [IsClientCoDelete],
	CONSTRAINT [DF_PackageItemHistory_ClientCoDeleteDesc] DEFAULT ('') FOR [ClientCoDeleteDesc],
	CONSTRAINT [DF_PackageItemHistory_IsClientCoApprove] DEFAULT (0) FOR [IsClientCoApprove],
	CONSTRAINT [DF_PackageItemHistory_ClientCoApproveDesc] DEFAULT ('') FOR [ClientCoApproveDesc],
	CONSTRAINT [DF_PackageItemHistory_SendMe] DEFAULT (0) FOR [SendMe],
	CONSTRAINT [DF_PackageItemHistory_IsDeleted] DEFAULT (0) FOR [IsDeleted],
	CONSTRAINT [DF_PackageItemHistory_DownLoadMe] DEFAULT (0) FOR [DownLoadMe],
	CONSTRAINT [DF_PackageItemHistory_UpLoadMe] DEFAULT (0) FOR [UpLoadMe]
GO

ALTER TABLE [dbo].[PolicyLimits] ADD 
	CONSTRAINT [DF_PolicyLimits_LimitAmount] DEFAULT (0) FOR [LimitAmount],
	CONSTRAINT [DF_PolicyLimits_RCSaidProp] DEFAULT (0) FOR [RCSaidProp],
	CONSTRAINT [DF_PolicyLimits_Reserves] DEFAULT (0) FOR [Reserves],
	CONSTRAINT [DF_PolicyLimits_IsDeleted] DEFAULT (0) FOR [IsDeleted],
	CONSTRAINT [DF_PolicyLimits_DownLoadMe] DEFAULT (0) FOR [DownLoadMe],
	CONSTRAINT [DF_PolicyLimits_UpLoadMe] DEFAULT (0) FOR [UpLoadMe],
	CONSTRAINT [DF_PolicyLimits_AdminComments] DEFAULT ('') FOR [AdminComments],
	CONSTRAINT [IX_PolicyLimits] UNIQUE  NONCLUSTERED 
	(
		[AssignmentsID],
		[ID],
		[IDAssignments],
		[ClassTypeID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[PolicyLimitsHistory] ADD 
	CONSTRAINT [DF_PolicyLimitsHistory_LimitAmount] DEFAULT (0) FOR [LimitAmount],
	CONSTRAINT [DF_PolicyLimitsHistory_RCSaidProp] DEFAULT (0) FOR [RCSaidProp],
	CONSTRAINT [DF_PolicyLimitsHistory_IsDeleted] DEFAULT (0) FOR [IsDeleted],
	CONSTRAINT [DF_PolicyLimitsHistory_DownLoadMe] DEFAULT (0) FOR [DownLoadMe],
	CONSTRAINT [DF_PolicyLimitsHistory_UpLoadMe] DEFAULT (0) FOR [UpLoadMe],
	CONSTRAINT [DF_PolicyLimitsHistory_AdminComments] DEFAULT ('') FOR [AdminComments]
GO

ALTER TABLE [dbo].[RTActivityLog] ADD 
	CONSTRAINT [DF_RTActivityLog_PageBreakAfter] DEFAULT (0) FOR [PageBreakAfter],
	CONSTRAINT [DF_RTActivityLog_BlankPageAfter] DEFAULT (0) FOR [BlankPageAfter],
	CONSTRAINT [DF_RTActivityLog_BlankRowsAfter] DEFAULT (0) FOR [BlankRowsAfter],
	CONSTRAINT [DF_RTActivityLog_IsMgrEntry] DEFAULT (0) FOR [IsMgrEntry],
	CONSTRAINT [DF_RTActivityLog_IsDeleted] DEFAULT (0) FOR [IsDeleted],
	CONSTRAINT [DF_RTActivityLog_DownLoadMe] DEFAULT (0) FOR [DownLoadMe],
	CONSTRAINT [DF_RTActivityLog_UpLoadMe] DEFAULT (0) FOR [UpLoadMe]
GO

ALTER TABLE [dbo].[RTActivityLogHistory] ADD 
	CONSTRAINT [DF_RTActivityLogHistory_PageBreak] DEFAULT (0) FOR [PageBreakAfter],
	CONSTRAINT [DF_RTActivityLogHistory_BlankPageAfter] DEFAULT (0) FOR [BlankPageAfter],
	CONSTRAINT [DF_RTActivityLogHistory_BlankRowsAfter] DEFAULT (0) FOR [BlankRowsAfter],
	CONSTRAINT [DF_RTActivityLogHistory_IsMgrEntry] DEFAULT (0) FOR [IsMgrEntry],
	CONSTRAINT [DF_RTActivityLogHistory_IsDeleted] DEFAULT (0) FOR [IsDeleted],
	CONSTRAINT [DF_RTActivityLogHistory_DownLoadMe] DEFAULT (0) FOR [DownLoadMe],
	CONSTRAINT [DF_RTActivityLogHistory_UpLoadMe] DEFAULT (0) FOR [UpLoadMe]
GO

ALTER TABLE [dbo].[RTActivityLogInfo] ADD 
	CONSTRAINT [DF_RTActivityLogInfo_IsDeleted] DEFAULT (0) FOR [IsDeleted],
	CONSTRAINT [DF_RTActivityLogInfo_DownLoadMe] DEFAULT (0) FOR [DownLoadMe],
	CONSTRAINT [DF_RTActivityLogInfo_UpLoadMe] DEFAULT (0) FOR [UpLoadMe],
	CONSTRAINT [IX_RTActivityLogInfo] UNIQUE  NONCLUSTERED 
	(
		[AssignmentsID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[RTActivityLogInfoHistory] ADD 
	CONSTRAINT [DF_RTActivityLogInfoHistory_IsDeleted] DEFAULT (0) FOR [IsDeleted],
	CONSTRAINT [DF_RTActivityLogInfoHistory_DownLoadMe] DEFAULT (0) FOR [DownLoadMe],
	CONSTRAINT [DF_RTActivityLogInfoHistory_UpLoadMe] DEFAULT (0) FOR [UpLoadMe]
GO

ALTER TABLE [dbo].[RTAttachments] ADD 
	CONSTRAINT [DF_RTAttachments_Description] DEFAULT ('') FOR [Description],
	CONSTRAINT [DF_RTAttachments_DownloadAttachment] DEFAULT (0) FOR [DownloadAttachment],
	CONSTRAINT [DF_RTAttachments_UpLoadAttachment] DEFAULT (0) FOR [UpLoadAttachment],
	CONSTRAINT [DF_RTAttachments_IsDeleted] DEFAULT (0) FOR [IsDeleted],
	CONSTRAINT [DF_RTAttachments_DownLoadMe] DEFAULT (0) FOR [DownLoadMe],
	CONSTRAINT [DF_RTAttachments_UpLoadMe] DEFAULT (0) FOR [UpLoadMe],
	CONSTRAINT [DF_RTAttachments_AdminComments] DEFAULT ('') FOR [AdminComments],
	CONSTRAINT [IX_RTAttachments] UNIQUE  NONCLUSTERED 
	(
		[Attachment]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[RTAttachmentsHistory] ADD 
	CONSTRAINT [DF_RTAttachmentsHistory_DownloadAttachment] DEFAULT (0) FOR [DownloadAttachment],
	CONSTRAINT [DF_RTAttachmentsHistory_UpLoadAttachment] DEFAULT (0) FOR [UpLoadAttachment],
	CONSTRAINT [DF_RTAttachmentsHistory_IsDeleted] DEFAULT (0) FOR [IsDeleted],
	CONSTRAINT [DF_RTAttachmentsHistory_DownLoadMe] DEFAULT (0) FOR [DownLoadMe],
	CONSTRAINT [DF_RTAttachmentsHistory_UpLoadMe] DEFAULT (0) FOR [UpLoadMe]
GO

ALTER TABLE [dbo].[RTChecks] ADD 
	CONSTRAINT [DF_RTChecks_AppliedDeductible] DEFAULT (0) FOR [AppliedDeductible],
	CONSTRAINT [DF_RTChecks_PrintOnIB] DEFAULT (0) FOR [PrintOnIB],
	CONSTRAINT [DF_RTChecks_IsDeleted] DEFAULT (0) FOR [IsDeleted],
	CONSTRAINT [DF_RTChecks_DownLoadMe] DEFAULT (0) FOR [DownLoadMe],
	CONSTRAINT [DF_RTChecks_UpLoadMe] DEFAULT (0) FOR [UpLoadMe],
	CONSTRAINT [IX_RTChecks] UNIQUE  NONCLUSTERED 
	(
		[AssignmentsID],
		[BillingCountID],
		[CheckNum]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[RTChecksHistory] ADD 
	CONSTRAINT [DF_RTChecksHistory_AppliedDeductible] DEFAULT (0) FOR [AppliedDeductible],
	CONSTRAINT [DF_RTChecksHistory_PrintOnIB] DEFAULT (0) FOR [PrintOnIB],
	CONSTRAINT [DF_RTChecksHistory_IsDeleted] DEFAULT (0) FOR [IsDeleted],
	CONSTRAINT [DF_RTChecksHistory_DownLoadMe] DEFAULT (0) FOR [DownLoadMe],
	CONSTRAINT [DF_RTChecksHistory_UpLoadMe] DEFAULT (0) FOR [UpLoadMe]
GO

ALTER TABLE [dbo].[RTIB] ADD 
	CONSTRAINT [DF_RTIB_RT09a_sPolicyNo] DEFAULT ('') FOR [RT09a_sPolicyNo],
	CONSTRAINT [DF_RTIB_RT14a_sSupplement] DEFAULT (0) FOR [RT14a_sSupplement],
	CONSTRAINT [DF_RTIB_RT14b_sRebilled] DEFAULT (0) FOR [RT14b_sRebilled],
	CONSTRAINT [DF_RTIB_RT15c_cLessMiscellaneous] DEFAULT (0) FOR [RT15c_cLessMiscellaneous],
	CONSTRAINT [DF_RTIB_RT15d_cMiscellaneousDesc] DEFAULT ('') FOR [RT15d_cMiscellaneousDesc],
	CONSTRAINT [DF_RTIB_Void] DEFAULT (0) FOR [Void],
	CONSTRAINT [DF_RTIB_FeeByTime] DEFAULT (0) FOR [FeeByTime],
	CONSTRAINT [DF_RTIB_UseActivityTime] DEFAULT (0) FOR [UseActivityTime],
	CONSTRAINT [DF_RTIB_DownLoadMe] DEFAULT (0) FOR [DownLoadMe],
	CONSTRAINT [DF_RTIB_UploadMe] DEFAULT (0) FOR [UpLoadMe],
	CONSTRAINT [DF_RTIB_Comments] DEFAULT ('') FOR [Comments],
	CONSTRAINT [IX_RTIB_AssignmentsID] UNIQUE  NONCLUSTERED 
	(
		[AssignmentsID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] ,
	CONSTRAINT [IX_RTIB_BillingCountID] UNIQUE  NONCLUSTERED 
	(
		[BillingCountID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] ,
	CONSTRAINT [IX_RTIB_ibnumber] UNIQUE  NONCLUSTERED 
	(
		[RT02_sIBNumber],
		[AssignmentsID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[RTIBFee] ADD 
	CONSTRAINT [DF_RTIBFee_Amount] DEFAULT (0) FOR [Amount],
	CONSTRAINT [DF_RTIBFee_Comment] DEFAULT ('') FOR [Comment],
	CONSTRAINT [DF_RTIBFee_DownLoadMe] DEFAULT (0) FOR [DownLoadMe],
	CONSTRAINT [DF_RTIBFee_UpLoadMe] DEFAULT (0) FOR [UpLoadMe]
GO

ALTER TABLE [dbo].[RTIBFeeHistory] ADD 
	CONSTRAINT [DF_RTIBFeeHistory_DownLoadMe] DEFAULT (0) FOR [DownLoadMe],
	CONSTRAINT [DF_RTIBFeeHistory_UpLoadMe] DEFAULT (0) FOR [UpLoadMe]
GO

ALTER TABLE [dbo].[RTIBHistory] ADD 
	CONSTRAINT [DF_RTIBHistory_RT14a_sSupplement] DEFAULT (0) FOR [RT14a_sSupplement],
	CONSTRAINT [DF_RTIBHistory_RT14b_sRebilled] DEFAULT (0) FOR [RT14b_sRebilled],
	CONSTRAINT [DF_RTIBHistory_Void] DEFAULT (0) FOR [Void],
	CONSTRAINT [DF_RTIBHistory_FeeByTime] DEFAULT (0) FOR [FeeByTime],
	CONSTRAINT [DF_RTIBHistory_UseActivityTime] DEFAULT (0) FOR [UseActivityTime],
	CONSTRAINT [DF_RTIBHistory_DownLoadMe] DEFAULT (0) FOR [DownLoadMe],
	CONSTRAINT [DF_RTIBHistory_UpLoadMe] DEFAULT (0) FOR [UpLoadMe]
GO

 CREATE  INDEX [IX_RTIBHistory] ON [dbo].[RTIBHistory]([AssignmentsID], [BillingCountID]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

ALTER TABLE [dbo].[RTIndemnity] ADD 
	CONSTRAINT [DF_RTIndemnity_ACVClaim] DEFAULT (0) FOR [ACVClaim],
	CONSTRAINT [DF_RTIndemnity_ACVLessExcessLimits] DEFAULT (0) FOR [ACVLessExcessLimits],
	CONSTRAINT [DF_RTIndemnity_SpecialLimits] DEFAULT (0) FOR [SpecialLimits],
	CONSTRAINT [DF_RTIndemnity_ExcessLimits] DEFAULT (0) FOR [ExcessLimits],
	CONSTRAINT [DF_RTIndemnity_Miscellaneous] DEFAULT (0) FOR [Miscellaneous],
	CONSTRAINT [DF_RTIndemnity_MiscDescription] DEFAULT ('') FOR [MiscDescription],
	CONSTRAINT [DF_RTIndemnity_IsAddAmountOfInsurance] DEFAULT (0) FOR [IsAddAmountOfInsurance],
	CONSTRAINT [DF_RTIndemnity_ExcessAbsorbsDeductible] DEFAULT (1) FOR [ExcessAbsorbsDeductible],
	CONSTRAINT [DF_RTIndemnity_AppliedDeductible] DEFAULT (0) FOR [AppliedDeductible],
	CONSTRAINT [DF_RTIndemnity_NonRecoverableDep] DEFAULT (0) FOR [NonRecoverableDep],
	CONSTRAINT [DF_RTIndemnity_RecoverableDep] DEFAULT (0) FOR [RecoverableDep],
	CONSTRAINT [DF_RTIndemnity_ReplacementCost] DEFAULT (0) FOR [ReplacementCost],
	CONSTRAINT [DF_RTIndemnity_IsPreviousPayment] DEFAULT (0) FOR [IsPreviousPayment],
	CONSTRAINT [DF_RTIndemnity_PPayAmountPaid] DEFAULT (0) FOR [PPayAmountPaid],
	CONSTRAINT [DF_RTIndemnity_PPayCheckNumber] DEFAULT ('') FOR [PPayCheckNumber],
	CONSTRAINT [DF_RTIndemnity_IsDeleted] DEFAULT (0) FOR [IsDeleted],
	CONSTRAINT [DF_RTIndemnity_DownLoadMe] DEFAULT (0) FOR [DownLoadMe],
	CONSTRAINT [DF_RTIndemnity_UpLoadMe] DEFAULT (0) FOR [UpLoadMe]
GO

ALTER TABLE [dbo].[RTIndemnityHistory] ADD 
	CONSTRAINT [DF_RTIndemnityHistory_ACVClaim] DEFAULT (0) FOR [ACVClaim],
	CONSTRAINT [DF_RTIndemnityHistory_ACVLessExcessLimits] DEFAULT (0) FOR [ACVLessExcessLimits],
	CONSTRAINT [DF_RTIndemnityHistory_SpecialLimits] DEFAULT (0) FOR [SpecialLimits],
	CONSTRAINT [DF_RTIndemnityHistory_ExcessLimits] DEFAULT (0) FOR [ExcessLimits],
	CONSTRAINT [DF_RTIndemnityHistory_Micellaneous] DEFAULT (0) FOR [Miscellaneous],
	CONSTRAINT [DF_RTIndemnityHistory_MiscDescription] DEFAULT ('') FOR [MiscDescription],
	CONSTRAINT [DF_RTIndemnityHistory_IsAddAmountOfInsurance] DEFAULT (0) FOR [IsAddAmountOfInsurance],
	CONSTRAINT [DF_RTIndemnityHistory_ExcessAbsorbsDeductible] DEFAULT (1) FOR [ExcessAbsorbsDeductible],
	CONSTRAINT [DF_RTIndemnityHistory_AppliedDeductible] DEFAULT (0) FOR [AppliedDeductible],
	CONSTRAINT [DF_RTIndemnityHistory_NonRecoverableDep] DEFAULT (0) FOR [NonRecoverableDep],
	CONSTRAINT [DF_RTIndemnityHistory_RecoverableDep] DEFAULT (0) FOR [RecoverableDep],
	CONSTRAINT [DF_RTIndemnityHistory_ReplacementCost] DEFAULT (0) FOR [ReplacementCost],
	CONSTRAINT [DF_RTIndemnityHistory_IsPreviousPayment] DEFAULT (0) FOR [IsPreviousPayment],
	CONSTRAINT [DF_RTIndemnityHistory_PPayAmountPaid] DEFAULT (0) FOR [PPayAmountPaid],
	CONSTRAINT [DF_RTIndemnityHistory_PPayCheckNumber] DEFAULT ('') FOR [PPayCheckNumber],
	CONSTRAINT [DF_RTIndemnityHistory_IsDeleted] DEFAULT (0) FOR [IsDeleted],
	CONSTRAINT [DF_RTIndemnityHistory_DownLoadMe] DEFAULT (0) FOR [DownLoadMe],
	CONSTRAINT [DF_RTIndemnityHistory_UpLoadMe] DEFAULT (0) FOR [UpLoadMe]
GO

ALTER TABLE [dbo].[RTPhotoLog] ADD 
	CONSTRAINT [DF_RTPhotoLog_Photo] DEFAULT ('') FOR [Photo],
	CONSTRAINT [DF_RTPhotoLog_DownloadPhoto] DEFAULT (0) FOR [DownloadPhoto],
	CONSTRAINT [DF_RTPhotoLog_UpLoadPhoto] DEFAULT (0) FOR [UpLoadPhoto],
	CONSTRAINT [DF_RTPhotoLog_PhotoThumb] DEFAULT ('') FOR [PhotoThumb],
	CONSTRAINT [DF_RTPhotoLog_DownloadPhotoThumb] DEFAULT (0) FOR [DownloadPhotoThumb],
	CONSTRAINT [DF_RTPhotoLog_UpLoadPhotoThumb] DEFAULT (0) FOR [UpLoadPhotoThumb],
	CONSTRAINT [DF_RTPhotoLog_PhotoHighRes] DEFAULT ('') FOR [PhotoHighRes],
	CONSTRAINT [DF_RTPhotoLog_DownloadPhotoHighRes] DEFAULT (0) FOR [DownloadPhotoHighRes],
	CONSTRAINT [DF_RTPhotoLog_UploadPhotoHishRes] DEFAULT (0) FOR [UploadPhotoHighRes],
	CONSTRAINT [DF_RTPhotoLog_IsDeleted] DEFAULT (0) FOR [IsDeleted],
	CONSTRAINT [DF_RTPhotoLog_DownLoadMe] DEFAULT (0) FOR [DownLoadMe],
	CONSTRAINT [DF_RTPhotoLog_UpLoadMe] DEFAULT (0) FOR [UpLoadMe],
	CONSTRAINT [IX_RTPhotoLog] UNIQUE  NONCLUSTERED 
	(
		[RTPhotoLogID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[RTPhotoLogHistory] ADD 
	CONSTRAINT [DF_RTPhotoLogHistory_DownloadPhoto] DEFAULT (0) FOR [DownloadPhoto],
	CONSTRAINT [DF_RTPhotoLogHistory_UpLoadPhoto] DEFAULT (0) FOR [UpLoadPhoto],
	CONSTRAINT [DF_RTPhotoLogHistory_DownloadPhotoThumb] DEFAULT (0) FOR [DownloadPhotoThumb],
	CONSTRAINT [DF_RTPhotoLogHistory_UpLoadPhotoThumb] DEFAULT (0) FOR [UpLoadPhotoThumb],
	CONSTRAINT [DF_RTPhotoLogHistory_DownloadPhotoHighRes] DEFAULT (0) FOR [DownloadPhotoHighRes],
	CONSTRAINT [DF_RTPhotoLogHistory_UploadPhotoHishRes] DEFAULT (0) FOR [UploadPhotoHighRes],
	CONSTRAINT [DF_RTPhotoLogHistory_IsDeleted] DEFAULT (0) FOR [IsDeleted],
	CONSTRAINT [DF_RTPhotoLogHistory_DownLoadMe] DEFAULT (0) FOR [DownLoadMe],
	CONSTRAINT [DF_RTPhotoLogHistory_UpLoadMe] DEFAULT (0) FOR [UpLoadMe]
GO

ALTER TABLE [dbo].[RTPhotoReport] ADD 
	CONSTRAINT [DF_RTPhotoReport_Name] DEFAULT ('') FOR [Name],
	CONSTRAINT [DF_RTPhotoReport_Description] DEFAULT ('') FOR [Description],
	CONSTRAINT [DF_RTPhotoReport_Number] DEFAULT (1) FOR [Number],
	CONSTRAINT [DF_RTPhotoReport_IsDeleted] DEFAULT (0) FOR [IsDeleted],
	CONSTRAINT [DF_RTPhotoReport_DownLoadMe] DEFAULT (0) FOR [DownLoadMe],
	CONSTRAINT [DF_RTPhotoReport_UpLoadMe] DEFAULT (0) FOR [UpLoadMe],
	CONSTRAINT [DF_RTPhotoReport_AdminComments] DEFAULT ('') FOR [AdminComments]
GO

ALTER TABLE [dbo].[RTPhotoReportHistory] ADD 
	CONSTRAINT [DF_RTPhotoReportHistory_Name] DEFAULT ('') FOR [Name],
	CONSTRAINT [DF_RTPhotoReportHistory_Description] DEFAULT ('') FOR [Description],
	CONSTRAINT [DF_RTPhotoReportHistory_Number] DEFAULT (1) FOR [Number],
	CONSTRAINT [DF_RTPhotoReportHistory_IsDeleted] DEFAULT (0) FOR [IsDeleted],
	CONSTRAINT [DF_RTPhotoReportHistory_DownLoadMe] DEFAULT (0) FOR [DownLoadMe],
	CONSTRAINT [DF_RTPhotoReportHistory_UpLoadMe] DEFAULT (0) FOR [UpLoadMe],
	CONSTRAINT [DF_RTPhotoReportHistory_AdminComments] DEFAULT ('') FOR [AdminComments]
GO

ALTER TABLE [dbo].[RTWSDiagram] ADD 
	CONSTRAINT [DF_RTWSDiagram_Name] DEFAULT ('') FOR [Name],
	CONSTRAINT [DF_RTWSDiagram_Description] DEFAULT ('') FOR [Description],
	CONSTRAINT [DF_RTWSDiagram_Number] DEFAULT (1) FOR [Number],
	CONSTRAINT [DF_RTWSDiagram_DiagramPhotoName] DEFAULT ('') FOR [DiagramPhotoName],
	CONSTRAINT [DF_RTWSDiagram_DownloadDiagramPhoto] DEFAULT (0) FOR [DownloadDiagramPhoto],
	CONSTRAINT [DF_RTWSDiagram_UploadDiagramPhoto] DEFAULT (0) FOR [UploadDiagramPhoto],
	CONSTRAINT [DF_RTWSDiagram_DiagramXML] DEFAULT ('') FOR [DiagramXML],
	CONSTRAINT [DF_RTWSDiagram_IsDeleted] DEFAULT (0) FOR [IsDeleted],
	CONSTRAINT [DF_RTWSDiagram_DownLoadMe] DEFAULT (0) FOR [DownLoadMe],
	CONSTRAINT [DF_RTWSDiagram_UpLoadMe] DEFAULT (0) FOR [UpLoadMe],
	CONSTRAINT [DF_RTWSDiagram_AdminComments] DEFAULT ('') FOR [AdminComments]
GO

ALTER TABLE [dbo].[RTWSDiagramHistory] ADD 
	CONSTRAINT [DF_RTWSDiagramHistory_Name] DEFAULT ('') FOR [Name],
	CONSTRAINT [DF_RTWSDiagramHistory_Description] DEFAULT ('') FOR [Description],
	CONSTRAINT [DF_RTWSDiagramHistory_Number] DEFAULT (1) FOR [Number],
	CONSTRAINT [DF_RTWSDiagramHistory_DiagramPhotoName] DEFAULT ('') FOR [DiagramPhotoName],
	CONSTRAINT [DF_RTWSDiagramHistory_DownloadDiagramPhoto] DEFAULT (0) FOR [DownloadDiagramPhoto],
	CONSTRAINT [DF_RTWSDiagramHistory_UploadDiagramPhoto] DEFAULT (0) FOR [UploadDiagramPhoto],
	CONSTRAINT [DF_RTWSDiagramHistory_DiagramXML] DEFAULT ('') FOR [DiagramXML],
	CONSTRAINT [DF_RTWSDiagramHistory_IsDeleted] DEFAULT (0) FOR [IsDeleted],
	CONSTRAINT [DF_RTWSDiagramHistory_DownLoadMe] DEFAULT (0) FOR [DownLoadMe],
	CONSTRAINT [DF_RTWSDiagramHistory_UpLoadMe] DEFAULT (0) FOR [UpLoadMe],
	CONSTRAINT [DF_RTWSDiagramHistory_AdminComments] DEFAULT ('') FOR [AdminComments]
GO

ALTER TABLE [dbo].[RegSetting] ADD 
	CONSTRAINT [DF_RegSetting_Version] DEFAULT (1) FOR [Version],
	CONSTRAINT [DF_RegSetting_SectionLevel01] DEFAULT ('') FOR [SectionLevel01],
	CONSTRAINT [DF_RegSetting_SectionLevel02] DEFAULT ('') FOR [SectionLevel02],
	CONSTRAINT [DF_RegSetting_SectionLevel03] DEFAULT ('') FOR [SectionLevel03],
	CONSTRAINT [DF_RegSetting_SectionLevel04] DEFAULT ('') FOR [SectionLevel04],
	CONSTRAINT [DF_RegSetting_SectionLevel05] DEFAULT ('') FOR [SectionLevel05],
	CONSTRAINT [DF_RegSetting_IsDeleted] DEFAULT (0) FOR [IsDeleted],
	CONSTRAINT [IX_RegSetting] UNIQUE  NONCLUSTERED 
	(
		[RegName]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[RegSettingHistory] ADD 
	CONSTRAINT [DF_RegSettingHistory_Version] DEFAULT (1) FOR [Version],
	CONSTRAINT [DF_RegSettingHistory_SectionLevel01] DEFAULT ('') FOR [SectionLevel01],
	CONSTRAINT [DF_RegSettingHistory_SectionLevel02] DEFAULT ('') FOR [SectionLevel02],
	CONSTRAINT [DF_RegSettingHistory_SectionLevel03] DEFAULT ('') FOR [SectionLevel03],
	CONSTRAINT [DF_RegSettingHistory_SectionLevel04] DEFAULT ('') FOR [SectionLevel04],
	CONSTRAINT [DF_RegSettingHistory_SectionLevel05] DEFAULT ('') FOR [SectionLevel05],
	CONSTRAINT [DF_RegSettingHistory_IsDeleted] DEFAULT (0) FOR [IsDeleted]
GO

ALTER TABLE [dbo].[SecurityLevel] ADD 
	CONSTRAINT [DF_SecurityLevel_IsDeleted] DEFAULT (0) FOR [IsDeleted],
	CONSTRAINT [IX_SecurityLevel] UNIQUE  NONCLUSTERED 
	(
		[Name]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[SoftwarePackage] ADD 
	CONSTRAINT [DF_SoftwarePackage_Version] DEFAULT (1) FOR [SPVersion],
	CONSTRAINT [DF_SoftwarePackage_IsDeleted] DEFAULT (0) FOR [IsDeleted],
	CONSTRAINT [IX_SoftwarePackage_1] UNIQUE  NONCLUSTERED 
	(
		[ClientCompanyID],
		[CATID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[SoftwarePackageApplication] ADD 
	CONSTRAINT [DF_SoftwarePackageApplication_Active] DEFAULT (0) FOR [IsDeleted]
GO

ALTER TABLE [dbo].[SoftwarePackageDocument] ADD 
	CONSTRAINT [DF_SoftwarePackageDocument_Active] DEFAULT (0) FOR [IsDeleted]
GO

ALTER TABLE [dbo].[SoftwarePackageHistory] ADD 
	CONSTRAINT [DF_SoftwarePackageHistory_SPVersion] DEFAULT (1) FOR [SPVersion],
	CONSTRAINT [DF_SoftwarePackageHistory_IsDeleted] DEFAULT (0) FOR [IsDeleted]
GO

ALTER TABLE [dbo].[SoftwarePackageRegSetting] ADD 
	CONSTRAINT [DF_SoftwarePackageRegSetting_Active] DEFAULT (0) FOR [IsDeleted]
GO

ALTER TABLE [dbo].[State] ADD 
	CONSTRAINT [DF_State_IsDeleted] DEFAULT (0) FOR [IsDeleted],
	CONSTRAINT [IX_State] UNIQUE  NONCLUSTERED 
	(
		[Code]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[StateHistory] ADD 
	CONSTRAINT [DF_StateHistory_IsDeleted] DEFAULT (0) FOR [IsDeleted]
GO

ALTER TABLE [dbo].[Status] ADD 
	CONSTRAINT [DF_Status_StatusAlias] DEFAULT ('') FOR [StatusAlias],
	CONSTRAINT [DF_Status_IsDeleted] DEFAULT (0) FOR [IsDeleted]
GO

ALTER TABLE [dbo].[StatusHistory] ADD 
	CONSTRAINT [DF_StatusHistory_StatusAlias] DEFAULT ('') FOR [StatusAlias],
	CONSTRAINT [DF_StatusHistory_IsDeleted] DEFAULT (0) FOR [IsDeleted]
GO

ALTER TABLE [dbo].[Temporary] ADD 
	CONSTRAINT [DF_Temporary_Active] DEFAULT (1) FOR [Active]
GO

ALTER TABLE [dbo].[TransType] ADD 
	CONSTRAINT [DF_TransType_Inbound] DEFAULT (1) FOR [Inbound],
	CONSTRAINT [DF_TransType_AllowProcessDefault] DEFAULT (1) FOR [AllowProcessDefault],
	CONSTRAINT [DF_TransType_Definition] DEFAULT ('') FOR [Definition],
	CONSTRAINT [DF_TransType_DateLastUpdated] DEFAULT (getdate()) FOR [DateLastUpdated]
GO

ALTER TABLE [dbo].[TypeOfLoss] ADD 
	CONSTRAINT [DF_TypeOfLoss_IsDeleted] DEFAULT (0) FOR [IsDeleted],
	CONSTRAINT [IX_TypeOfLoss] UNIQUE  NONCLUSTERED 
	(
		[ClientCompanyID],
		[Code]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[TypeOfLossHistory] ADD 
	CONSTRAINT [DF_TypeOfLossHistory_IsDeleted] DEFAULT (0) FOR [IsDeleted]
GO

ALTER TABLE [dbo].[UserProfile] ADD 
	CONSTRAINT [DF_UserProfile_Active] DEFAULT (1) FOR [Active],
	CONSTRAINT [DF_UserProfile_SortOrder] DEFAULT (1) FOR [SortOrder],
	CONSTRAINT [IX_UserProfile] UNIQUE  NONCLUSTERED 
	(
		[TableName]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[UserProfileHistory] ADD 
	CONSTRAINT [DF_UserProfileHistory_Active] DEFAULT (1) FOR [Active]
GO

ALTER TABLE [dbo].[UserReportsToCoordinator] ADD 
	CONSTRAINT [DF_UserReportsToCoordinator_Active] DEFAULT (1) FOR [Active]
GO

ALTER TABLE [dbo].[UserReportsToManager] ADD 
	CONSTRAINT [DF_UserReportsToManager_Active] DEFAULT (1) FOR [Active]
GO

ALTER TABLE [dbo].[Users] ADD 
	CONSTRAINT [DF_Users_Active] DEFAULT (1) FOR [Active],
	CONSTRAINT [DF_Users_ActiveDate] DEFAULT (getdate()) FOR [ActiveDate],
	CONSTRAINT [DF_Users_SecurityLevel] DEFAULT (1) FOR [SecurityLevel],
	CONSTRAINT [IX_Users] UNIQUE  NONCLUSTERED 
	(
		[UserName]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Work_ListText] ADD 
	CONSTRAINT [DF_Work_ListText_Work_ListTextID] DEFAULT (newid()) FOR [Work_ListTextID]
GO

ALTER TABLE [dbo].[XML_Trans] ADD 
	CONSTRAINT [DF_XML_Trans_Mod20] DEFAULT (0) FOR [Mod20],
	CONSTRAINT [DF_XML_Trans_InBound] DEFAULT (1) FOR [InBound],
	CONSTRAINT [DF_XML_Trans_AllowProcess] DEFAULT (0) FOR [AllowProcess],
	CONSTRAINT [DF_XML_Trans_TransType] DEFAULT ('') FOR [TransType],
	CONSTRAINT [DF_XML_Trans_TransDescription] DEFAULT ('') FOR [TransDescription],
	CONSTRAINT [DF_XML_Trans_Comments] DEFAULT ('') FOR [Comments],
	CONSTRAINT [DF_XML_Trans_IsErr] DEFAULT (0) FOR [IsErr],
	CONSTRAINT [DF_XML_Trans_ErrMess] DEFAULT ('') FOR [ErrMess],
	CONSTRAINT [DF_XML_Trans_DatePosted] DEFAULT (getdate()) FOR [DatePosted],
	CONSTRAINT [DF_XML_Trans_DateLastUpdated] DEFAULT (getdate()) FOR [DateLastUpdated],
	CONSTRAINT [DF_XML_Trans_remote_addr] DEFAULT ('') FOR [remote_addr]
GO

ALTER TABLE [dbo].[Accounting] ADD 
	CONSTRAINT [FK_Accounting_CompanyUsers] FOREIGN KEY 
	(
		[CompanyID],
		[UsersID]
	) REFERENCES [dbo].[CompanyUsers] (
		[CompanyID],
		[UsersID]
	)
GO

ALTER TABLE [dbo].[Adjuster] ADD 
	CONSTRAINT [FK_Adjuster_AdjusterUsersSoftware] FOREIGN KEY 
	(
		[UsersID]
	) REFERENCES [dbo].[AdjusterUsersSoftware] (
		[UsersID]
	),
	CONSTRAINT [FK_Adjuster_AdjusterUsersUpdates] FOREIGN KEY 
	(
		[UsersID]
	) REFERENCES [dbo].[AdjusterUsersUpdates] (
		[UsersID]
	),
	CONSTRAINT [FK_CompanyAdjusterUsers_CompanyUsers] FOREIGN KEY 
	(
		[CompanyID],
		[UsersID]
	) REFERENCES [dbo].[CompanyUsers] (
		[CompanyID],
		[UsersID]
	) ON DELETE CASCADE 
GO

ALTER TABLE [dbo].[AdjusterUsersSoftwareHistory] ADD 
	CONSTRAINT [FK_UsersSoftwareHistory_UsersSoftware] FOREIGN KEY 
	(
		[UsersID]
	) REFERENCES [dbo].[AdjusterUsersSoftware] (
		[UsersID]
	) ON DELETE CASCADE 
GO

ALTER TABLE [dbo].[AdjusterUsersUpdates] ADD 
	CONSTRAINT [FK_UsersUpdates_Users] FOREIGN KEY 
	(
		[UsersID]
	) REFERENCES [dbo].[Users] (
		[UsersID]
	) ON DELETE CASCADE 
GO

ALTER TABLE [dbo].[AdjusterUsersUpdatesHistory] ADD 
	CONSTRAINT [FK_UsersUpdatesHistory_UsersUpdates] FOREIGN KEY 
	(
		[UsersID]
	) REFERENCES [dbo].[AdjusterUsersUpdates] (
		[UsersID]
	) ON DELETE CASCADE 
GO

ALTER TABLE [dbo].[Admin] ADD 
	CONSTRAINT [FK_CompanyAdminUsers_CompanyUsers] FOREIGN KEY 
	(
		[CompanyID],
		[UsersID]
	) REFERENCES [dbo].[CompanyUsers] (
		[CompanyID],
		[UsersID]
	) ON DELETE CASCADE 
GO

ALTER TABLE [dbo].[ApplicationHistory] ADD 
	CONSTRAINT [FK_ApplicationHistory_Application] FOREIGN KEY 
	(
		[ApplicationID]
	) REFERENCES [dbo].[Application] (
		[ApplicationID]
	)
GO

ALTER TABLE [dbo].[AssignmentTypeHistory] ADD 
	CONSTRAINT [FK_AssignmentTypeHistory_AssignmentType] FOREIGN KEY 
	(
		[AssignmentTypeID]
	) REFERENCES [dbo].[AssignmentType] (
		[AssignmentTypeID]
	)
GO

ALTER TABLE [dbo].[Assignments] ADD 
	CONSTRAINT [FK_Assignments_AssignmentType] FOREIGN KEY 
	(
		[AssignmentTypeID]
	) REFERENCES [dbo].[AssignmentType] (
		[AssignmentTypeID]
	),
	CONSTRAINT [FK_Assignments_ClientCoAdjusterSpec] FOREIGN KEY 
	(
		[AdjusterSpecID]
	) REFERENCES [dbo].[ClientCoAdjusterSpec] (
		[ClientCoAdjusterSpecID]
	),
	CONSTRAINT [FK_Assignments_ClientCompanyCatSpec] FOREIGN KEY 
	(
		[ClientCompanyCatSpecID]
	) REFERENCES [dbo].[ClientCompanyCatSpec] (
		[ClientCompanyCatSpecID]
	),
	CONSTRAINT [FK_Assignments_Status] FOREIGN KEY 
	(
		[StatusID]
	) REFERENCES [dbo].[Status] (
		[StatusID]
	),
	CONSTRAINT [FK_Assignments_TypeOfLoss] FOREIGN KEY 
	(
		[TypeOfLossID]
	) REFERENCES [dbo].[TypeOfLoss] (
		[TypeOfLossID]
	)
GO

ALTER TABLE [dbo].[AssignmentsHistory] ADD 
	CONSTRAINT [FK_AssignmentsHistory_Assignments] FOREIGN KEY 
	(
		[AssignmentsID]
	) REFERENCES [dbo].[Assignments] (
		[AssignmentsID]
	)
GO

ALTER TABLE [dbo].[Batches] ADD 
	CONSTRAINT [FK_Batches_Assignments] FOREIGN KEY 
	(
		[AssignmentsID]
	) REFERENCES [dbo].[Assignments] (
		[AssignmentsID]
	),
	CONSTRAINT [FK_Batches_BillAssignment] FOREIGN KEY 
	(
		[BillAssignmentID]
	) REFERENCES [dbo].[BillAssignment] (
		[BillAssignmentID]
	),
	CONSTRAINT [FK_Batches_CompanyCatSpec] FOREIGN KEY 
	(
		[ClientCompanyCatSpecID]
	) REFERENCES [dbo].[ClientCompanyCatSpec] (
		[ClientCompanyCatSpecID]
	)
GO

ALTER TABLE [dbo].[BatchesHistory] ADD 
	CONSTRAINT [FK_BatchesHistory_Batches] FOREIGN KEY 
	(
		[BatchesID]
	) REFERENCES [dbo].[Batches] (
		[BatchesID]
	)
GO

ALTER TABLE [dbo].[BillAssignment] ADD 
	CONSTRAINT [FK_BillAssignment_AssignmentType] FOREIGN KEY 
	(
		[AssignmentTypeID]
	) REFERENCES [dbo].[AssignmentType] (
		[AssignmentTypeID]
	),
	CONSTRAINT [FK_BillAssignment_ClientCoAdjusterSpec] FOREIGN KEY 
	(
		[AdjusterSpecID]
	) REFERENCES [dbo].[ClientCoAdjusterSpec] (
		[ClientCoAdjusterSpecID]
	),
	CONSTRAINT [FK_BillAssignment_ClientCompanyCatSpec] FOREIGN KEY 
	(
		[ClientCompanyCatSpecID]
	) REFERENCES [dbo].[ClientCompanyCatSpec] (
		[ClientCompanyCatSpecID]
	)
GO

ALTER TABLE [dbo].[BillBillingCount] ADD 
	CONSTRAINT [FK_BillBillingCount_BillAssignment] FOREIGN KEY 
	(
		[BillAssignmentID]
	) REFERENCES [dbo].[BillAssignment] (
		[BillAssignmentID]
	)
GO

ALTER TABLE [dbo].[BillingCount] ADD 
	CONSTRAINT [FK_BillingCount_Assignments] FOREIGN KEY 
	(
		[AssignmentsID]
	) REFERENCES [dbo].[Assignments] (
		[AssignmentsID]
	)
GO

ALTER TABLE [dbo].[CAT] ADD 
	CONSTRAINT [FK_CAT_AssignmentType] FOREIGN KEY 
	(
		[AssignmentTypeID]
	) REFERENCES [dbo].[AssignmentType] (
		[AssignmentTypeID]
	),
	CONSTRAINT [FK_CAT_Company] FOREIGN KEY 
	(
		[CompanyID]
	) REFERENCES [dbo].[Company] (
		[CompanyID]
	) ON DELETE CASCADE 
GO

ALTER TABLE [dbo].[CATHistory] ADD 
	CONSTRAINT [FK_CATHistory_CAT] FOREIGN KEY 
	(
		[CATID]
	) REFERENCES [dbo].[CAT] (
		[CATID]
	)
GO

ALTER TABLE [dbo].[ClassOfLoss] ADD 
	CONSTRAINT [FK_ClassOfLoss_ClassOfLoss] FOREIGN KEY 
	(
		[IsSubSetOFClassOfLossID]
	) REFERENCES [dbo].[ClassOfLoss] (
		[ClassOfLossID]
	),
	CONSTRAINT [FK_ClassOfLoss_ClassType] FOREIGN KEY 
	(
		[ClassTypeID]
	) REFERENCES [dbo].[ClassType] (
		[ClassTypeID]
	),
	CONSTRAINT [FK_ClassOfLoss_Company] FOREIGN KEY 
	(
		[ClientCompanyID]
	) REFERENCES [dbo].[Company] (
		[CompanyID]
	)
GO

ALTER TABLE [dbo].[ClassOfLossHistory] ADD 
	CONSTRAINT [FK_ClassOfLossHistory_ClassOfLoss1] FOREIGN KEY 
	(
		[ClassOfLossID]
	) REFERENCES [dbo].[ClassOfLoss] (
		[ClassOfLossID]
	)
GO

ALTER TABLE [dbo].[ClassTypeHistory] ADD 
	CONSTRAINT [FK_ClassTypeHistory_ClassType] FOREIGN KEY 
	(
		[ClassTypeID]
	) REFERENCES [dbo].[ClassType] (
		[ClassTypeID]
	)
GO

ALTER TABLE [dbo].[Client] ADD 
	CONSTRAINT [FK_Client_CompanyUsers] FOREIGN KEY 
	(
		[CompanyID],
		[UsersID]
	) REFERENCES [dbo].[CompanyUsers] (
		[CompanyID],
		[UsersID]
	) ON DELETE CASCADE 
GO

ALTER TABLE [dbo].[ClientCoAdjusterSpec] ADD 
	CONSTRAINT [FK_ClientCoAdjusterSpec_ClientCompanyCatSpec] FOREIGN KEY 
	(
		[ClientCompanyCatSpecID]
	) REFERENCES [dbo].[ClientCompanyCatSpec] (
		[ClientCompanyCatSpecID]
	),
	CONSTRAINT [FK_UserCompanySpec_CompanyAdjusterUsers] FOREIGN KEY 
	(
		[ClientCompanyID],
		[UsersID]
	) REFERENCES [dbo].[Adjuster] (
		[CompanyID],
		[UsersID]
	)
GO

ALTER TABLE [dbo].[ClientCoAdjusterSpecHistory] ADD 
	CONSTRAINT [FK_AdjusterSpecHistory_AdjusterSpec] FOREIGN KEY 
	(
		[ClientCoAdjusterSpecID]
	) REFERENCES [dbo].[ClientCoAdjusterSpec] (
		[ClientCoAdjusterSpecID]
	)
GO

ALTER TABLE [dbo].[ClientCompanyCat] ADD 
	CONSTRAINT [FK_CompanyCat_CAT] FOREIGN KEY 
	(
		[CATID]
	) REFERENCES [dbo].[CAT] (
		[CATID]
	),
	CONSTRAINT [FK_CompanyCat_Company1] FOREIGN KEY 
	(
		[ClientCompanyID]
	) REFERENCES [dbo].[Company] (
		[CompanyID]
	) ON DELETE CASCADE ,
	CONSTRAINT [FK_CompanyCat_FeeSchedule] FOREIGN KEY 
	(
		[FeeScheduleID]
	) REFERENCES [dbo].[FeeSchedule] (
		[FeeScheduleID]
	) ON DELETE CASCADE ,
	CONSTRAINT [FK_CompanyCat_TypeOfLoss] FOREIGN KEY 
	(
		[TypeOfLossID]
	) REFERENCES [dbo].[TypeOfLoss] (
		[TypeOfLossID]
	) ON DELETE CASCADE 
GO

ALTER TABLE [dbo].[ClientCompanyCatHistory] ADD 
	CONSTRAINT [FK_CompanyCatHistory_CompanyCat] FOREIGN KEY 
	(
		[ClientCompanyID],
		[CATID]
	) REFERENCES [dbo].[ClientCompanyCat] (
		[ClientCompanyID],
		[CATID]
	)
GO

ALTER TABLE [dbo].[ClientCompanyCatSpec] ADD 
	CONSTRAINT [FK_CompanyCatSpec_CAT] FOREIGN KEY 
	(
		[CATID]
	) REFERENCES [dbo].[CAT] (
		[CATID]
	),
	CONSTRAINT [FK_CompanyCatSpec_Company] FOREIGN KEY 
	(
		[ClientCompanyID]
	) REFERENCES [dbo].[Company] (
		[CompanyID]
	),
	CONSTRAINT [FK_CompanyCatSpec_CompanyCat] FOREIGN KEY 
	(
		[ClientCompanyID],
		[CATID]
	) REFERENCES [dbo].[ClientCompanyCat] (
		[ClientCompanyID],
		[CATID]
	)
GO

ALTER TABLE [dbo].[ClientCompanyCatSpecHistory] ADD 
	CONSTRAINT [FK_CompanyCatSpecHistory_CompanyCatSpec] FOREIGN KEY 
	(
		[ClientCompanyCatSpecID]
	) REFERENCES [dbo].[ClientCompanyCatSpec] (
		[ClientCompanyCatSpecID]
	)
GO

ALTER TABLE [dbo].[ClientCompanyUsersCat] ADD 
	CONSTRAINT [FK_CompanyUsersCat_CompanyUsers] FOREIGN KEY 
	(
		[ClientCompanyID],
		[UsersID]
	) REFERENCES [dbo].[CompanyUsers] (
		[CompanyID],
		[UsersID]
	),
	CONSTRAINT [FK_UsrersCat_CompanyCat] FOREIGN KEY 
	(
		[ClientCompanyID],
		[CATID]
	) REFERENCES [dbo].[ClientCompanyCat] (
		[ClientCompanyID],
		[CATID]
	) ON DELETE CASCADE 
GO

ALTER TABLE [dbo].[Company] ADD 
	CONSTRAINT [FK_Company_Company] FOREIGN KEY 
	(
		[IsClientOf]
	) REFERENCES [dbo].[Company] (
		[CompanyID]
	)
GO

ALTER TABLE [dbo].[CompanyHistory] ADD 
	CONSTRAINT [FK_CompanyHistory_Company] FOREIGN KEY 
	(
		[CompanyID]
	) REFERENCES [dbo].[Company] (
		[CompanyID]
	)
GO

ALTER TABLE [dbo].[CompanyUsers] ADD 
	CONSTRAINT [FK_CompanyUsers_Company] FOREIGN KEY 
	(
		[CompanyID]
	) REFERENCES [dbo].[Company] (
		[CompanyID]
	) ON DELETE CASCADE ,
	CONSTRAINT [FK_CompanyUsers_SecurityLevel] FOREIGN KEY 
	(
		[SecurityLevel]
	) REFERENCES [dbo].[SecurityLevel] (
		[SecurityLevel]
	),
	CONSTRAINT [FK_CompanyUsers_Users] FOREIGN KEY 
	(
		[UsersID]
	) REFERENCES [dbo].[Users] (
		[UsersID]
	)
GO

ALTER TABLE [dbo].[CompanyUsersHistory] ADD 
	CONSTRAINT [FK_CompanyUsersHistory_CompanyUsers] FOREIGN KEY 
	(
		[CompanyID],
		[UsersID]
	) REFERENCES [dbo].[CompanyUsers] (
		[CompanyID],
		[UsersID]
	) ON DELETE CASCADE 
GO

ALTER TABLE [dbo].[Coordinator] ADD 
	CONSTRAINT [FK_Coordinator_CompanyUsers] FOREIGN KEY 
	(
		[CompanyID],
		[UsersID]
	) REFERENCES [dbo].[CompanyUsers] (
		[CompanyID],
		[UsersID]
	) ON DELETE CASCADE 
GO

ALTER TABLE [dbo].[Dispatcher] ADD 
	CONSTRAINT [FK_Dispatcher_CompanyUsers] FOREIGN KEY 
	(
		[CompanyID],
		[UsersID]
	) REFERENCES [dbo].[CompanyUsers] (
		[CompanyID],
		[UsersID]
	) ON DELETE CASCADE 
GO

ALTER TABLE [dbo].[DocumentHistory] ADD 
	CONSTRAINT [FK_DocumentHistory_Document] FOREIGN KEY 
	(
		[DocumentID]
	) REFERENCES [dbo].[Document] (
		[DocumentID]
	)
GO

ALTER TABLE [dbo].[ECSADJUsers] ADD 
	CONSTRAINT [FK_ECSADJUsers_Users] FOREIGN KEY 
	(
		[UsersID]
	) REFERENCES [dbo].[Users] (
		[UsersID]
	) ON DELETE CASCADE 
GO

ALTER TABLE [dbo].[Employee] ADD 
	CONSTRAINT [FK_Employee_CompanyUsers] FOREIGN KEY 
	(
		[CompanyID],
		[UsersID]
	) REFERENCES [dbo].[CompanyUsers] (
		[CompanyID],
		[UsersID]
	) ON DELETE CASCADE 
GO

ALTER TABLE [dbo].[FAQSHistory] ADD 
	CONSTRAINT [FK_FAQSHistory_FAQS] FOREIGN KEY 
	(
		[FAQSID]
	) REFERENCES [dbo].[FAQS] (
		[FAQSID]
	)
GO

ALTER TABLE [dbo].[FeeSchedule] ADD 
	CONSTRAINT [FK_FeeSchedule_Company] FOREIGN KEY 
	(
		[ClientCompanyID]
	) REFERENCES [dbo].[Company] (
		[CompanyID]
	)
GO

ALTER TABLE [dbo].[FeeScheduleFeeTypes] ADD 
	CONSTRAINT [FK_FeeScheduleFeeTypes_FeeSchedule] FOREIGN KEY 
	(
		[FeeScheduleID]
	) REFERENCES [dbo].[FeeSchedule] (
		[FeeScheduleID]
	) ON DELETE CASCADE 
GO

ALTER TABLE [dbo].[FeeScheduleFeeTypesHistory] ADD 
	CONSTRAINT [FK_FeeScheduleFeeTypesHistory_FeeScheduleFeeTypes] FOREIGN KEY 
	(
		[FeeScheduleFeeTypesID]
	) REFERENCES [dbo].[FeeScheduleFeeTypes] (
		[FeeScheduleFeeTypesID]
	) ON DELETE CASCADE 
GO

ALTER TABLE [dbo].[FeeScheduleHistory] ADD 
	CONSTRAINT [FK_FeeScheduleHistory_FeeSchedule] FOREIGN KEY 
	(
		[FeeScheduleID]
	) REFERENCES [dbo].[FeeSchedule] (
		[FeeScheduleID]
	)
GO

ALTER TABLE [dbo].[FeeScheduleLevels] ADD 
	CONSTRAINT [FK_FeeScheduleLevels_FeeSchedule] FOREIGN KEY 
	(
		[FeeScheduleID]
	) REFERENCES [dbo].[FeeSchedule] (
		[FeeScheduleID]
	) ON DELETE CASCADE 
GO

ALTER TABLE [dbo].[FeeScheduleLevelsHistory] ADD 
	CONSTRAINT [FK_FeeScheduleLevelsHistory_FeeScheduleLevels] FOREIGN KEY 
	(
		[FeeScheduleLevelsID]
	) REFERENCES [dbo].[FeeScheduleLevels] (
		[FeeScheduleLevelsID]
	) ON DELETE CASCADE 
GO

ALTER TABLE [dbo].[IB] ADD 
	CONSTRAINT [FK_IB_Assignments] FOREIGN KEY 
	(
		[AssignmentsID]
	) REFERENCES [dbo].[Assignments] (
		[AssignmentsID]
	),
	CONSTRAINT [FK_IB_BillingCount] FOREIGN KEY 
	(
		[BillingCountID]
	) REFERENCES [dbo].[BillingCount] (
		[BillingCountID]
	)
GO

ALTER TABLE [dbo].[IBFee] ADD 
	CONSTRAINT [FK_IBFee_Assignments] FOREIGN KEY 
	(
		[AssignmentsID]
	) REFERENCES [dbo].[Assignments] (
		[AssignmentsID]
	),
	CONSTRAINT [FK_IBFee_FeeScheduleFeeTypes] FOREIGN KEY 
	(
		[FeeScheduleFeeTypesID]
	) REFERENCES [dbo].[FeeScheduleFeeTypes] (
		[FeeScheduleFeeTypesID]
	),
	CONSTRAINT [FK_IBFee_IB] FOREIGN KEY 
	(
		[IBID]
	) REFERENCES [dbo].[IB] (
		[IBID]
	) ON DELETE CASCADE 
GO

ALTER TABLE [dbo].[IBFeeHistory] ADD 
	CONSTRAINT [FK_IBFeeHistory_IBFee] FOREIGN KEY 
	(
		[IBFeeID]
	) REFERENCES [dbo].[IBFee] (
		[IBFeeID]
	) ON DELETE CASCADE 
GO

ALTER TABLE [dbo].[IBHistory] ADD 
	CONSTRAINT [FK_IBHistory_IB] FOREIGN KEY 
	(
		[IBID]
	) REFERENCES [dbo].[IB] (
		[IBID]
	)
GO

ALTER TABLE [dbo].[IBStateFarm] ADD 
	CONSTRAINT [FK_IBStateFarm_BillAssignment] FOREIGN KEY 
	(
		[BillAssignmentID]
	) REFERENCES [dbo].[BillAssignment] (
		[BillAssignmentID]
	),
	CONSTRAINT [FK_IBStateFarm_BillBillingCount] FOREIGN KEY 
	(
		[BillBillingCountID]
	) REFERENCES [dbo].[BillBillingCount] (
		[BillBillingCountID]
	)
GO

ALTER TABLE [dbo].[Manager] ADD 
	CONSTRAINT [FK_CompanyManagerUsers_CompanyUsers] FOREIGN KEY 
	(
		[CompanyID],
		[UsersID]
	) REFERENCES [dbo].[CompanyUsers] (
		[CompanyID],
		[UsersID]
	) ON DELETE CASCADE 
GO

ALTER TABLE [dbo].[MiscReportParam] ADD 
	CONSTRAINT [FK_MiscReportParam_Assignments] FOREIGN KEY 
	(
		[AssignmentsID]
	) REFERENCES [dbo].[Assignments] (
		[AssignmentsID]
	)
GO

ALTER TABLE [dbo].[MiscReportParam01] ADD 
	CONSTRAINT [FK_MiscReportParam01_Assignments] FOREIGN KEY 
	(
		[AssignmentsID]
	) REFERENCES [dbo].[Assignments] (
		[AssignmentsID]
	)
GO

ALTER TABLE [dbo].[MiscReportParam02] ADD 
	CONSTRAINT [FK_MiscReportParam02_Assignments] FOREIGN KEY 
	(
		[AssignmentsID]
	) REFERENCES [dbo].[Assignments] (
		[AssignmentsID]
	)
GO

ALTER TABLE [dbo].[MiscReportParam03] ADD 
	CONSTRAINT [FK_MiscReportParam03_Assignments] FOREIGN KEY 
	(
		[AssignmentsID]
	) REFERENCES [dbo].[Assignments] (
		[AssignmentsID]
	)
GO

ALTER TABLE [dbo].[MiscReportParam04] ADD 
	CONSTRAINT [FK_MiscReportParam04_Assignments] FOREIGN KEY 
	(
		[AssignmentsID]
	) REFERENCES [dbo].[Assignments] (
		[AssignmentsID]
	)
GO

ALTER TABLE [dbo].[MiscReportParam05] ADD 
	CONSTRAINT [FK_MiscReportParam05_Assignments] FOREIGN KEY 
	(
		[AssignmentsID]
	) REFERENCES [dbo].[Assignments] (
		[AssignmentsID]
	)
GO

ALTER TABLE [dbo].[MiscReportParam06] ADD 
	CONSTRAINT [FK_MiscReportParam06_Assignments] FOREIGN KEY 
	(
		[AssignmentsID]
	) REFERENCES [dbo].[Assignments] (
		[AssignmentsID]
	)
GO

ALTER TABLE [dbo].[MiscReportParam07] ADD 
	CONSTRAINT [FK_MiscReportParam07_Assignments] FOREIGN KEY 
	(
		[AssignmentsID]
	) REFERENCES [dbo].[Assignments] (
		[AssignmentsID]
	)
GO

ALTER TABLE [dbo].[MiscReportParam08] ADD 
	CONSTRAINT [FK_MiscReportParam08_Assignments] FOREIGN KEY 
	(
		[AssignmentsID]
	) REFERENCES [dbo].[Assignments] (
		[AssignmentsID]
	)
GO

ALTER TABLE [dbo].[MiscReportParam09] ADD 
	CONSTRAINT [FK_MiscReportParam09_Assignments] FOREIGN KEY 
	(
		[AssignmentsID]
	) REFERENCES [dbo].[Assignments] (
		[AssignmentsID]
	)
GO

ALTER TABLE [dbo].[MiscReportParam10] ADD 
	CONSTRAINT [FK_MiscReportParam10_Assignments] FOREIGN KEY 
	(
		[AssignmentsID]
	) REFERENCES [dbo].[Assignments] (
		[AssignmentsID]
	)
GO

ALTER TABLE [dbo].[MiscReportParam11] ADD 
	CONSTRAINT [FK_MiscReportParam11_Assignments] FOREIGN KEY 
	(
		[AssignmentsID]
	) REFERENCES [dbo].[Assignments] (
		[AssignmentsID]
	)
GO

ALTER TABLE [dbo].[MiscReportParam12] ADD 
	CONSTRAINT [FK_MiscReportParam12_Assignments] FOREIGN KEY 
	(
		[AssignmentsID]
	) REFERENCES [dbo].[Assignments] (
		[AssignmentsID]
	)
GO

ALTER TABLE [dbo].[MiscReportParam13] ADD 
	CONSTRAINT [FK_MiscReportParam13_Assignments] FOREIGN KEY 
	(
		[AssignmentsID]
	) REFERENCES [dbo].[Assignments] (
		[AssignmentsID]
	)
GO

ALTER TABLE [dbo].[MiscReportParam14] ADD 
	CONSTRAINT [FK_MiscReportParam14_Assignments] FOREIGN KEY 
	(
		[AssignmentsID]
	) REFERENCES [dbo].[Assignments] (
		[AssignmentsID]
	)
GO

ALTER TABLE [dbo].[MiscReportParam15] ADD 
	CONSTRAINT [FK_MiscReportParam15_Assignments] FOREIGN KEY 
	(
		[AssignmentsID]
	) REFERENCES [dbo].[Assignments] (
		[AssignmentsID]
	)
GO

ALTER TABLE [dbo].[MiscReportParam16] ADD 
	CONSTRAINT [FK_MiscReportParam16_Assignments] FOREIGN KEY 
	(
		[AssignmentsID]
	) REFERENCES [dbo].[Assignments] (
		[AssignmentsID]
	)
GO

ALTER TABLE [dbo].[MiscReportParam17] ADD 
	CONSTRAINT [FK_MiscReportParam17_Assignments] FOREIGN KEY 
	(
		[AssignmentsID]
	) REFERENCES [dbo].[Assignments] (
		[AssignmentsID]
	)
GO

ALTER TABLE [dbo].[MiscReportParam18] ADD 
	CONSTRAINT [FK_MiscReportParam18_Assignments] FOREIGN KEY 
	(
		[AssignmentsID]
	) REFERENCES [dbo].[Assignments] (
		[AssignmentsID]
	)
GO

ALTER TABLE [dbo].[MiscReportParam19] ADD 
	CONSTRAINT [FK_MiscReportParam19_Assignments] FOREIGN KEY 
	(
		[AssignmentsID]
	) REFERENCES [dbo].[Assignments] (
		[AssignmentsID]
	)
GO

ALTER TABLE [dbo].[MiscReportParam20] ADD 
	CONSTRAINT [FK_MiscReportParam20_Assignments] FOREIGN KEY 
	(
		[AssignmentsID]
	) REFERENCES [dbo].[Assignments] (
		[AssignmentsID]
	)
GO

ALTER TABLE [dbo].[MiscReportParam21] ADD 
	CONSTRAINT [FK_MiscReportParam21_Assignments] FOREIGN KEY 
	(
		[AssignmentsID]
	) REFERENCES [dbo].[Assignments] (
		[AssignmentsID]
	)
GO

ALTER TABLE [dbo].[MiscReportParam22] ADD 
	CONSTRAINT [FK_MiscReportParam22_Assignments] FOREIGN KEY 
	(
		[AssignmentsID]
	) REFERENCES [dbo].[Assignments] (
		[AssignmentsID]
	)
GO

ALTER TABLE [dbo].[MiscReportParam23] ADD 
	CONSTRAINT [FK_MiscReportParam23_Assignments] FOREIGN KEY 
	(
		[AssignmentsID]
	) REFERENCES [dbo].[Assignments] (
		[AssignmentsID]
	)
GO

ALTER TABLE [dbo].[MiscReportParam24] ADD 
	CONSTRAINT [FK_MiscReportParam24_Assignments] FOREIGN KEY 
	(
		[AssignmentsID]
	) REFERENCES [dbo].[Assignments] (
		[AssignmentsID]
	)
GO

ALTER TABLE [dbo].[MiscReportParam25] ADD 
	CONSTRAINT [FK_MiscReportParam25_Assignments] FOREIGN KEY 
	(
		[AssignmentsID]
	) REFERENCES [dbo].[Assignments] (
		[AssignmentsID]
	)
GO

ALTER TABLE [dbo].[MiscReportParam26] ADD 
	CONSTRAINT [FK_MiscReportParam26_Assignments] FOREIGN KEY 
	(
		[AssignmentsID]
	) REFERENCES [dbo].[Assignments] (
		[AssignmentsID]
	)
GO

ALTER TABLE [dbo].[MiscReportParam27] ADD 
	CONSTRAINT [FK_MiscReportParam27_Assignments] FOREIGN KEY 
	(
		[AssignmentsID]
	) REFERENCES [dbo].[Assignments] (
		[AssignmentsID]
	)
GO

ALTER TABLE [dbo].[MiscReportParam28] ADD 
	CONSTRAINT [FK_MiscReportParam28_Assignments] FOREIGN KEY 
	(
		[AssignmentsID]
	) REFERENCES [dbo].[Assignments] (
		[AssignmentsID]
	)
GO

ALTER TABLE [dbo].[MiscReportParam29] ADD 
	CONSTRAINT [FK_MiscReportParam29_Assignments] FOREIGN KEY 
	(
		[AssignmentsID]
	) REFERENCES [dbo].[Assignments] (
		[AssignmentsID]
	)
GO

ALTER TABLE [dbo].[MiscReportParam30] ADD 
	CONSTRAINT [FK_MiscReportParam30_Assignments] FOREIGN KEY 
	(
		[AssignmentsID]
	) REFERENCES [dbo].[Assignments] (
		[AssignmentsID]
	)
GO

ALTER TABLE [dbo].[Package] ADD 
	CONSTRAINT [FK_Package_Assignments] FOREIGN KEY 
	(
		[AssignmentsID]
	) REFERENCES [dbo].[Assignments] (
		[AssignmentsID]
	)
GO

ALTER TABLE [dbo].[PackageHistory] ADD 
	CONSTRAINT [FK_PackageHistory_Package] FOREIGN KEY 
	(
		[PackageID]
	) REFERENCES [dbo].[Package] (
		[PackageID]
	)
GO

ALTER TABLE [dbo].[PackageItem] ADD 
	CONSTRAINT [FK_PackageItem_Assignments] FOREIGN KEY 
	(
		[AssignmentsID]
	) REFERENCES [dbo].[Assignments] (
		[AssignmentsID]
	),
	CONSTRAINT [FK_PackageItem_Package] FOREIGN KEY 
	(
		[PackageID]
	) REFERENCES [dbo].[Package] (
		[PackageID]
	),
	CONSTRAINT [FK_PackageItem_RTAttachments] FOREIGN KEY 
	(
		[RTAttachmentsID]
	) REFERENCES [dbo].[RTAttachments] (
		[RTAttachmentsID]
	)
GO

ALTER TABLE [dbo].[PackageItemHistory] ADD 
	CONSTRAINT [FK_PackageItemHistory_PackageItem] FOREIGN KEY 
	(
		[PackageItemID]
	) REFERENCES [dbo].[PackageItem] (
		[PackageItemID]
	)
GO

ALTER TABLE [dbo].[PolicyLimits] ADD 
	CONSTRAINT [FK_PolicyLimits_Assignments] FOREIGN KEY 
	(
		[AssignmentsID]
	) REFERENCES [dbo].[Assignments] (
		[AssignmentsID]
	) ON DELETE CASCADE ,
	CONSTRAINT [FK_PolicyLimits_ClassType] FOREIGN KEY 
	(
		[ClassTypeID]
	) REFERENCES [dbo].[ClassType] (
		[ClassTypeID]
	)
GO

ALTER TABLE [dbo].[PolicyLimitsHistory] ADD 
	CONSTRAINT [FK_PolicyLimitsHistory_PolicyLimits] FOREIGN KEY 
	(
		[PolicyLimitsID]
	) REFERENCES [dbo].[PolicyLimits] (
		[PolicyLimitsID]
	) ON DELETE CASCADE 
GO

ALTER TABLE [dbo].[RTActivityLog] ADD 
	CONSTRAINT [FK_RTActivityLog_Assignments] FOREIGN KEY 
	(
		[AssignmentsID]
	) REFERENCES [dbo].[Assignments] (
		[AssignmentsID]
	),
	CONSTRAINT [FK_RTActivityLog_BillingCount] FOREIGN KEY 
	(
		[BillingCountID]
	) REFERENCES [dbo].[BillingCount] (
		[BillingCountID]
	)
GO

ALTER TABLE [dbo].[RTActivityLogHistory] ADD 
	CONSTRAINT [FK_RTActivityLogHistory_RTActivityLog] FOREIGN KEY 
	(
		[RTActivityLogID]
	) REFERENCES [dbo].[RTActivityLog] (
		[RTActivityLogID]
	)
GO

ALTER TABLE [dbo].[RTActivityLogInfo] ADD 
	CONSTRAINT [FK_RTActivityLogInfo_Assignments] FOREIGN KEY 
	(
		[AssignmentsID]
	) REFERENCES [dbo].[Assignments] (
		[AssignmentsID]
	)
GO

ALTER TABLE [dbo].[RTActivityLogInfoHistory] ADD 
	CONSTRAINT [FK_RTActivityLogInfoHistory_RTActivityLogInfo] FOREIGN KEY 
	(
		[AssignmentsID]
	) REFERENCES [dbo].[RTActivityLogInfo] (
		[AssignmentsID]
	)
GO

ALTER TABLE [dbo].[RTAttachments] ADD 
	CONSTRAINT [FK_RTAttachments_Assignments] FOREIGN KEY 
	(
		[AssignmentsID]
	) REFERENCES [dbo].[Assignments] (
		[AssignmentsID]
	)
GO

ALTER TABLE [dbo].[RTAttachmentsHistory] ADD 
	CONSTRAINT [FK_RTAttachmentsHistory_RTAttachments] FOREIGN KEY 
	(
		[RTAttachmentsID]
	) REFERENCES [dbo].[RTAttachments] (
		[RTAttachmentsID]
	)
GO

ALTER TABLE [dbo].[RTChecks] ADD 
	CONSTRAINT [FK_RTChecks_Assignments] FOREIGN KEY 
	(
		[AssignmentsID]
	) REFERENCES [dbo].[Assignments] (
		[AssignmentsID]
	),
	CONSTRAINT [FK_RTChecks_BillingCount] FOREIGN KEY 
	(
		[BillingCountID]
	) REFERENCES [dbo].[BillingCount] (
		[BillingCountID]
	),
	CONSTRAINT [FK_RTChecks_TypeOfLoss] FOREIGN KEY 
	(
		[RT43_TypeOfLossID]
	) REFERENCES [dbo].[TypeOfLoss] (
		[TypeOfLossID]
	)
GO

ALTER TABLE [dbo].[RTChecksHistory] ADD 
	CONSTRAINT [FK_RTChecksHistory_RTChecks] FOREIGN KEY 
	(
		[RTChecksID]
	) REFERENCES [dbo].[RTChecks] (
		[RTChecksID]
	)
GO

ALTER TABLE [dbo].[RTIB] ADD 
	CONSTRAINT [FK_RTIB_Assignments] FOREIGN KEY 
	(
		[AssignmentsID]
	) REFERENCES [dbo].[Assignments] (
		[AssignmentsID]
	),
	CONSTRAINT [FK_RTIB_BillingCount] FOREIGN KEY 
	(
		[BillingCountID]
	) REFERENCES [dbo].[BillingCount] (
		[BillingCountID]
	)
GO

ALTER TABLE [dbo].[RTIBFee] ADD 
	CONSTRAINT [FK_RTIBFee_Assignments] FOREIGN KEY 
	(
		[AssignmentsID]
	) REFERENCES [dbo].[Assignments] (
		[AssignmentsID]
	),
	CONSTRAINT [FK_RTIBFee_FeeScheduleFeeTypes] FOREIGN KEY 
	(
		[FeeScheduleFeeTypesID]
	) REFERENCES [dbo].[FeeScheduleFeeTypes] (
		[FeeScheduleFeeTypesID]
	),
	CONSTRAINT [FK_RTIBFee_RTIB] FOREIGN KEY 
	(
		[AssignmentsID]
	) REFERENCES [dbo].[RTIB] (
		[AssignmentsID]
	) ON DELETE CASCADE 
GO

ALTER TABLE [dbo].[RTIBFeeHistory] ADD 
	CONSTRAINT [FK_RTIBFeeHistory_RTIBFee] FOREIGN KEY 
	(
		[RTIBFeeID]
	) REFERENCES [dbo].[RTIBFee] (
		[RTIBFeeID]
	) ON DELETE CASCADE 
GO

ALTER TABLE [dbo].[RTIBHistory] ADD 
	CONSTRAINT [FK_RTIBHistory_RTIB] FOREIGN KEY 
	(
		[AssignmentsID]
	) REFERENCES [dbo].[RTIB] (
		[AssignmentsID]
	)
GO

ALTER TABLE [dbo].[RTIndemnity] ADD 
	CONSTRAINT [FK_RTIndemnity_Assignments] FOREIGN KEY 
	(
		[AssignmentsID]
	) REFERENCES [dbo].[Assignments] (
		[AssignmentsID]
	),
	CONSTRAINT [FK_RTIndemnity_ClassOfLoss] FOREIGN KEY 
	(
		[ClassOfLossID]
	) REFERENCES [dbo].[ClassOfLoss] (
		[ClassOfLossID]
	),
	CONSTRAINT [FK_RTIndemnity_RTChecks] FOREIGN KEY 
	(
		[RTChecksID]
	) REFERENCES [dbo].[RTChecks] (
		[RTChecksID]
	),
	CONSTRAINT [FK_RTIndemnity_TypeOfLoss] FOREIGN KEY 
	(
		[TypeOfLossID]
	) REFERENCES [dbo].[TypeOfLoss] (
		[TypeOfLossID]
	)
GO

ALTER TABLE [dbo].[RTIndemnityHistory] ADD 
	CONSTRAINT [FK_RTIndemnityHistory_RTIndemnity] FOREIGN KEY 
	(
		[RTIndemnityID]
	) REFERENCES [dbo].[RTIndemnity] (
		[RTIndemnityID]
	)
GO

ALTER TABLE [dbo].[RTPhotoLog] ADD 
	CONSTRAINT [FK_RTPhotoLog_Assignments] FOREIGN KEY 
	(
		[AssignmentsID]
	) REFERENCES [dbo].[Assignments] (
		[AssignmentsID]
	),
	CONSTRAINT [FK_RTPhotoLog_BillingCount] FOREIGN KEY 
	(
		[BillingCountID]
	) REFERENCES [dbo].[BillingCount] (
		[BillingCountID]
	),
	CONSTRAINT [FK_RTPhotoLog_RTPhotoReport] FOREIGN KEY 
	(
		[RTPhotoReportID]
	) REFERENCES [dbo].[RTPhotoReport] (
		[RTPhotoReportID]
	)
GO

ALTER TABLE [dbo].[RTPhotoLogHistory] ADD 
	CONSTRAINT [FK_RTPhotoLogHistory_RTPhotoLog] FOREIGN KEY 
	(
		[RTPhotoLogID]
	) REFERENCES [dbo].[RTPhotoLog] (
		[RTPhotoLogID]
	)
GO

ALTER TABLE [dbo].[RTPhotoReport] ADD 
	CONSTRAINT [FK_RTPhotoReport_Assignments] FOREIGN KEY 
	(
		[AssignmentsID]
	) REFERENCES [dbo].[Assignments] (
		[AssignmentsID]
	)
GO

ALTER TABLE [dbo].[RTPhotoReportHistory] ADD 
	CONSTRAINT [FK_RTPhotoReportHistory_RTPhotoReport] FOREIGN KEY 
	(
		[RTPhotoReportID]
	) REFERENCES [dbo].[RTPhotoReport] (
		[RTPhotoReportID]
	)
GO

ALTER TABLE [dbo].[RTWSDiagram] ADD 
	CONSTRAINT [FK_RTWSDiagram_Assignments] FOREIGN KEY 
	(
		[AssignmentsID]
	) REFERENCES [dbo].[Assignments] (
		[AssignmentsID]
	)
GO

ALTER TABLE [dbo].[RTWSDiagramHistory] ADD 
	CONSTRAINT [FK_RTWSDiagramHistory_RTWSDiagram] FOREIGN KEY 
	(
		[RTWSDiagramID]
	) REFERENCES [dbo].[RTWSDiagram] (
		[RTWSDiagramID]
	)
GO

ALTER TABLE [dbo].[RegSettingHistory] ADD 
	CONSTRAINT [FK_RegSettingHistory_RegSetting] FOREIGN KEY 
	(
		[RegSettingID]
	) REFERENCES [dbo].[RegSetting] (
		[RegSettingID]
	)
GO

ALTER TABLE [dbo].[SASecurityPackage] ADD 
	CONSTRAINT [FK_SASecirityPackage_SecurityArea1] FOREIGN KEY 
	(
		[SecurityAreaID]
	) REFERENCES [dbo].[SecurityArea] (
		[SecurityAreaID]
	),
	CONSTRAINT [FK_SASecirityPackage_SecurityPackage] FOREIGN KEY 
	(
		[SecurityPackageID]
	) REFERENCES [dbo].[SecurityPackage] (
		[SecurityPackageID]
	)
GO

ALTER TABLE [dbo].[SISecurityArea] ADD 
	CONSTRAINT [FK_SISecurityArea_SecurityArea] FOREIGN KEY 
	(
		[SecurityAreaID]
	) REFERENCES [dbo].[SecurityArea] (
		[SecurityAreaID]
	),
	CONSTRAINT [FK_SISecurityArea_SecurityItems] FOREIGN KEY 
	(
		[SecurityItemsID]
	) REFERENCES [dbo].[SecurityItems] (
		[SecurityItemsID]
	)
GO

ALTER TABLE [dbo].[SPSecurity] ADD 
	CONSTRAINT [FK_SPSecurity_Security] FOREIGN KEY 
	(
		[SecurityID]
	) REFERENCES [dbo].[Security] (
		[SecurityID]
	),
	CONSTRAINT [FK_SPSecurity_SecurityPackage] FOREIGN KEY 
	(
		[SecurityPackageID]
	) REFERENCES [dbo].[SecurityPackage] (
		[SecurityPackageID]
	)
GO

ALTER TABLE [dbo].[SecurityArea] ADD 
	CONSTRAINT [FK_SecurityArea_SecurityAreaType] FOREIGN KEY 
	(
		[SecurityAreaTypeID]
	) REFERENCES [dbo].[SecurityAreaType] (
		[SecurityAreaTypeID]
	)
GO

ALTER TABLE [dbo].[SecurityGroup] ADD 
	CONSTRAINT [FK_SecurityGroup_Group] FOREIGN KEY 
	(
		[GroupID]
	) REFERENCES [dbo].[Group] (
		[GroupID]
	),
	CONSTRAINT [FK_SecurityGroup_Security] FOREIGN KEY 
	(
		[SecurityID]
	) REFERENCES [dbo].[Security] (
		[SecurityID]
	)
GO

ALTER TABLE [dbo].[SecurityItems] ADD 
	CONSTRAINT [FK_SecurityItems_SecurityItemType] FOREIGN KEY 
	(
		[SecurityItemTypeID]
	) REFERENCES [dbo].[SecurityItemType] (
		[SecurityItemTypeID]
	)
GO

ALTER TABLE [dbo].[SecurityLevelHistory] ADD 
	CONSTRAINT [FK_SecurityLevelHistory_SecurityLevel] FOREIGN KEY 
	(
		[SecurityLevel]
	) REFERENCES [dbo].[SecurityLevel] (
		[SecurityLevel]
	)
GO

ALTER TABLE [dbo].[SoftwarePackage] ADD 
	CONSTRAINT [FK_SoftwarePackage_ClientCompanyCat] FOREIGN KEY 
	(
		[ClientCompanyID],
		[CATID]
	) REFERENCES [dbo].[ClientCompanyCat] (
		[ClientCompanyID],
		[CATID]
	) ON DELETE CASCADE 
GO

ALTER TABLE [dbo].[SoftwarePackageApplication] ADD 
	CONSTRAINT [FK_SoftwarePackageApplication_Application] FOREIGN KEY 
	(
		[ApplicationID]
	) REFERENCES [dbo].[Application] (
		[ApplicationID]
	),
	CONSTRAINT [FK_SoftwarePackageApplication_SoftwarePackage] FOREIGN KEY 
	(
		[SoftWarePackageID]
	) REFERENCES [dbo].[SoftwarePackage] (
		[SoftWarePackageID]
	) ON DELETE CASCADE 
GO

ALTER TABLE [dbo].[SoftwarePackageDocument] ADD 
	CONSTRAINT [FK_SoftwarePackageDocument_Document] FOREIGN KEY 
	(
		[DocumentID]
	) REFERENCES [dbo].[Document] (
		[DocumentID]
	),
	CONSTRAINT [FK_SoftwarePackageDocument_SoftwarePackage] FOREIGN KEY 
	(
		[SoftWarePackageID]
	) REFERENCES [dbo].[SoftwarePackage] (
		[SoftWarePackageID]
	) ON DELETE CASCADE 
GO

ALTER TABLE [dbo].[SoftwarePackageHistory] ADD 
	CONSTRAINT [FK_SoftwarePackageHistory_SoftwarePackage] FOREIGN KEY 
	(
		[SoftWarePackageID]
	) REFERENCES [dbo].[SoftwarePackage] (
		[SoftWarePackageID]
	)
GO

ALTER TABLE [dbo].[SoftwarePackageRegSetting] ADD 
	CONSTRAINT [FK_SoftwarePackageRegSetting_RegSetting] FOREIGN KEY 
	(
		[RegSettingID]
	) REFERENCES [dbo].[RegSetting] (
		[RegSettingID]
	),
	CONSTRAINT [FK_SoftwarePackageRegSetting_SoftwarePackage] FOREIGN KEY 
	(
		[SoftWarePackageID]
	) REFERENCES [dbo].[SoftwarePackage] (
		[SoftWarePackageID]
	) ON DELETE CASCADE 
GO

ALTER TABLE [dbo].[StateHistory] ADD 
	CONSTRAINT [FK_StateHistory_State] FOREIGN KEY 
	(
		[StateID]
	) REFERENCES [dbo].[State] (
		[StateID]
	)
GO

ALTER TABLE [dbo].[StatusHistory] ADD 
	CONSTRAINT [FK_StatusHistory_Status] FOREIGN KEY 
	(
		[StatusID]
	) REFERENCES [dbo].[Status] (
		[StatusID]
	)
GO

ALTER TABLE [dbo].[Temporary] ADD 
	CONSTRAINT [FK_Temporary_CompanyUsers] FOREIGN KEY 
	(
		[CompanyID],
		[UsersID]
	) REFERENCES [dbo].[CompanyUsers] (
		[CompanyID],
		[UsersID]
	) ON DELETE CASCADE 
GO

ALTER TABLE [dbo].[TransType] ADD 
	CONSTRAINT [FK_TransType_Company] FOREIGN KEY 
	(
		[ClientCompanyID]
	) REFERENCES [dbo].[Company] (
		[CompanyID]
	)
GO

ALTER TABLE [dbo].[TypeOfLoss] ADD 
	CONSTRAINT [FK_TypeOfLoss_Company] FOREIGN KEY 
	(
		[ClientCompanyID]
	) REFERENCES [dbo].[Company] (
		[CompanyID]
	)
GO

ALTER TABLE [dbo].[TypeOfLossHistory] ADD 
	CONSTRAINT [FK_TypeOfLossHistory_TypeOfLoss] FOREIGN KEY 
	(
		[TypeOfLossID]
	) REFERENCES [dbo].[TypeOfLoss] (
		[TypeOfLossID]
	)
GO

ALTER TABLE [dbo].[UserProfileHistory] ADD 
	CONSTRAINT [FK_UserProfileHistory_UserProfile] FOREIGN KEY 
	(
		[UserProfileID]
	) REFERENCES [dbo].[UserProfile] (
		[UserProfileID]
	)
GO

ALTER TABLE [dbo].[UserReportsToCoordinator] ADD 
	CONSTRAINT [FK_UserReportsToCoordinator_Coordinator] FOREIGN KEY 
	(
		[CompanyID],
		[ReportsToUsersID]
	) REFERENCES [dbo].[Coordinator] (
		[CompanyID],
		[UsersID]
	) ON DELETE CASCADE ,
	CONSTRAINT [FK_UserReportsToCoordinator_Manager] FOREIGN KEY 
	(
		[CompanyID],
		[UsersID]
	) REFERENCES [dbo].[Manager] (
		[CompanyID],
		[UsersID]
	)
GO

ALTER TABLE [dbo].[UserReportsToManager] ADD 
	CONSTRAINT [FK_UserReportsToManager_Adjuster] FOREIGN KEY 
	(
		[CompanyID],
		[UsersID]
	) REFERENCES [dbo].[Adjuster] (
		[CompanyID],
		[UsersID]
	) ON DELETE CASCADE ,
	CONSTRAINT [FK_UserReportsToManager_Manager] FOREIGN KEY 
	(
		[CompanyID],
		[ReportsToUsersID]
	) REFERENCES [dbo].[Manager] (
		[CompanyID],
		[UsersID]
	)
GO

ALTER TABLE [dbo].[Users] ADD 
	CONSTRAINT [FK_Users_SecurityLevel] FOREIGN KEY 
	(
		[SecurityLevel]
	) REFERENCES [dbo].[SecurityLevel] (
		[SecurityLevel]
	)
GO

ALTER TABLE [dbo].[UsersGroup] ADD 
	CONSTRAINT [FK_UserGroup_Group] FOREIGN KEY 
	(
		[GroupID]
	) REFERENCES [dbo].[Group] (
		[GroupID]
	),
	CONSTRAINT [FK_UserGroup_Users] FOREIGN KEY 
	(
		[UsersID]
	) REFERENCES [dbo].[Users] (
		[UsersID]
	)
GO

ALTER TABLE [dbo].[UsersHistory] ADD 
	CONSTRAINT [FK_UsersHistory_Users] FOREIGN KEY 
	(
		[UsersID]
	) REFERENCES [dbo].[Users] (
		[UsersID]
	)
GO

ALTER TABLE [dbo].[XML_Trans] ADD 
	CONSTRAINT [FK_XML_Trans_Assignments] FOREIGN KEY 
	(
		[AssignmentsID]
	) REFERENCES [dbo].[Assignments] (
		[AssignmentsID]
	),
	CONSTRAINT [FK_XML_Trans_Company] FOREIGN KEY 
	(
		[ClientCompanyID]
	) REFERENCES [dbo].[Company] (
		[CompanyID]
	),
	CONSTRAINT [FK_XML_Trans_Company1] FOREIGN KEY 
	(
		[TransFromCompanyID]
	) REFERENCES [dbo].[Company] (
		[CompanyID]
	),
	CONSTRAINT [FK_XML_Trans_Company2] FOREIGN KEY 
	(
		[TransToCompanyID]
	) REFERENCES [dbo].[Company] (
		[CompanyID]
	),
	CONSTRAINT [FK_XML_Trans_TransType] FOREIGN KEY 
	(
		[TransTypeID]
	) REFERENCES [dbo].[TransType] (
		[TransTypeID]
	)
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

--Because of the Restriction on Triggers when Ntext fields ...
/*
In a DELETE, INSERT, or UPDATE trigger, SQL Server does not allow text, ntext, or image column references 
in the inserted and deleted tables if the compatibility level is equal to 70. The text, ntext, and image values 
in the inserted and deleted tables cannot be accessed. To retrieve the new value in either an INSERT or UPDATE 
trigger, join the inserted table with the original update table. When the compatibility level is 65 or lower, 
null values are returned for inserted or deleted text, ntext, or image columns that allow null values; zero-length 
strings are returned if the columns are not nullable. 
If the compatibility level is 80 or higher, SQL Server allows the update of text, ntext, or image columns through 
the INSTEAD OF trigger on tables or views.
*/
--Use Instead OF to get around the above restriction
CREATE TRIGGER updAdjusterUsersSoftwareHistory
ON dbo.AdjusterUsersSoftware
INSTEAD OF UPDATE
AS
INSERT INTO AdjusterUsersSoftwareHistory	
	SELECT del.*
	FROM DELETED del

-- Now that the History table was updated first...
--Allow the original update to process...

Update AdjusterUsersSoftware SET
	[VersionInfo] = INS.VersionInfo,
	[LicenseDaysLeft] = INS.LicenseDaysLeft,
	[ResetLicense] = INS.ResetLicense,
	[IBPrefix] = dbo.VerifyNotDupIBPrefix(INS.IBPrefix, INS.UsersID),
	[ResetIBPrefix] = INS.ResetIBPrefix,
	[SingleFileSendAuthority] = INS.SingleFileSendAuthority,
	[DateLastUpdated] = INS.DateLastUpdated,
	[UpdateByUserID] = INS.UpdateByUserID
FROM AdjusterUsersSoftware U INNER JOIN INSERTED INS ON U.UsersID = INS.UsersID

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

--Because of the Restriction on Triggers when Ntext fields ...
/*
In a DELETE, INSERT, or UPDATE trigger, SQL Server does not allow text, ntext, or image column references 
in the inserted and deleted tables if the compatibility level is equal to 70. The text, ntext, and image values 
in the inserted and deleted tables cannot be accessed. To retrieve the new value in either an INSERT or UPDATE 
trigger, join the inserted table with the original update table. When the compatibility level is 65 or lower, 
null values are returned for inserted or deleted text, ntext, or image columns that allow null values; zero-length 
strings are returned if the columns are not nullable. 
If the compatibility level is 80 or higher, SQL Server allows the update of text, ntext, or image columns through 
the INSTEAD OF trigger on tables or views.
*/
--Use Instead OF to get around the above restriction
CREATE TRIGGER insAdjusterUsersSoftwareHistory
ON dbo.AdjusterUsersSoftware
INSTEAD OF INSERT
AS

INSERT INTO AdjusterUsersSoftware (
					[UsersID] ,
					[VersionInfo],
					[LicenseDaysLeft] ,
					[ResetLicense],
					[IBPrefix] ,
					[ResetIBPrefix] ,
					[SingleFileSendAuthority] ,
					[DateLastUpdated] ,
					[UpdateByUserID] 
				)
	SELECT
		INS.[UsersID] ,
		INS.[VersionInfo],
		INS.[LicenseDaysLeft],
		INS.[ResetLicense],
		dbo.VerifyNotDupIBPrefix(INS.[IBPrefix], INS.[UsersID]) AS [IBPrefix],
		INS.[ResetIBPrefix],
		INS.[SingleFileSendAuthority],	
		INS.[DateLastUpdated],
		INS.[UpdateByUserID]
	FROM  INSERTED INS

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE TRIGGER updAdjusterUsersUpdatesHistory
ON dbo.AdjusterUsersUpdates
FOR UPDATE
AS
INSERT INTO AdjusterUsersUpdatesHistory
	SELECT del.* 
	FROM DELETED del

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE TRIGGER insApplication
ON dbo.Application 
INSTEAD OF INSERT
AS

--As Well Update this records SPName and VersionDate And SPVersion and SPVersionBase
	INSERT INTO  Application (
					[AppNameBase],
					[AppName],
					[Description],
					[Version] ,
					[MajorVS],
					[MinorVS],
					[RevisionVS],
					[SPVersionBase],
					[SPVersion] ,
					[VersionDate],
					[ProjectName] ,
					[ClassName] ,
					[SectionLevel01],
					[SectionLevel02],
					[SectionLevel03],
					[SectionLevel04],
					[SectionLevel05],
					[InstallFileLocation],
					[SPName] ,
					[IsDeleted],
					[DateLastUpdated],
					[UpdateByUserID] 
				)
	SELECT 		
					INS.[AppNameBase],
					INS.[AppName],
					INS.[Description],
					INS.[Version],
					INS.[MajorVS],
					INS.[MinorVS],
					INS.[RevisionVS],
					(SELECT Max(SPVersion) FROM SoftwarePackage) As [SPVersionBase] ,
					(SELECT Max(SPVersion) FROM SoftwarePackage) As [SPVersion] ,
					GetDate() as [VersionDate],
					INS.[ProjectName] ,
					INS.[ClassName] ,
					INS.[SectionLevel01],
					INS.[SectionLevel02],
					INS.[SectionLevel03],
					INS.[SectionLevel04],
					INS.[SectionLevel05],
					INS.[InstallFileLocation],
					INS.[AppName] + '_V' + cast(INS.[Version] As VarChar(10)) + '.exe' As [SPName] ,
					INS.[IsDeleted],
					Getdate() as [DateLastUpdated] ,
					INS.[UpdateByUserID] 	
	FROM INSERTED INS

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE TRIGGER updApplicationHistory
ON dbo.Application
AFTER UPDATE
AS
--See if there are any Version Changes
DECLARE @CountVersion int
DECLARE @CountVersionDate int

SET @CountVersion = 	(
				SELECT Count(DEL.Version) As CountOFVersion
				FROM DELETED DEL INNER JOIN INSERTED INS
				ON DEL.ApplicationID = INS.ApplicationID
				WHERE (DEL.Version <> INS.Version)
				AND DEL.SPVersion Is Not Null	
			)

SET @CountVersionDate = 	(
				SELECT Count(DEL.VersionDate) As CountOFVersionDate
				FROM DELETED DEL INNER JOIN INSERTED INS
				ON DEL.ApplicationID = INS.ApplicationID
				WHERE (DEL.VersionDate <> INS.VersionDate)
				AND INS.VersionDate Is Not Null
				AND DEL.SPVersion Is Not Null	
				)

IF @CountVersion > 0
BEGIN
	INSERT INTO ApplicationHistory
		SELECT DEL.* 
		FROM DELETED DEL INNER JOIN INSERTED INS
		ON DEL.ApplicationID = INS.ApplicationID
		WHERE (DEL.Version <> INS.Version)
		AND DEL.SPVersion Is Not Null	

	--Then update the Softwarepackage Table which will in turn Update All dependant Tables
	--with new SPVersion.
	UPDATE SoftwarePackage SET SoftwarePackage.SPVersion =  (SELECT MAX(SoftwarePackage.SPVersion) +1  FROM SoftwarePackage),
					SoftwarePackage.VersionDate = getdate()
	FROM SoftwarePackage 
	WHERE SoftwarePackage.SoftwarePackageID IN 	(
									SELECT SPA.SoftwarePackageID
									FROM SoftwarePackageApplication SPA INNER JOIN (INSERTED INS INNER JOIN DELETED DEL ON INS.ApplicationID = DEL.ApplicationID) ON SPA.ApplicationID = INS.ApplicationID								
									WHERE (DEL.Version <> INS.Version)
									AND SPA.IsDeleted =0
								)
	--As Well Update this records SPName, VersionDate and SPVersionBase
	UPDATE Application SET 	Application.SPName = Application.AppName + '_V' + cast(Application.Version As VarChar(10)) + '.exe',
					Application.VersionDate = GetDate(),
					Application.SPVersionBase = Application.SPVersion
	FROM Application  INNER JOIN INSERTED INS On Application.ApplicationID = INS.ApplicationID INNER JOIN DELETED DEL On INS.ApplicationID = DEL.ApplicationID
END

ELSE
	IF @CountVersionDate > 0
	BEGIN
		--Update the software package VersionDate
		UPDATE SoftwarePackage SET 	SoftwarePackage.VersionDate = getdate()
		FROM SoftwarePackage 
		WHERE SoftwarePackage.SoftwarePackageID IN 	(
										SELECT SPA.SoftwarePackageID
										FROM SoftwarePackageApplication SPA INNER JOIN (INSERTED INS INNER JOIN DELETED DEL ON INS.ApplicationID = DEL.ApplicationID) ON SPA.ApplicationID = INS.ApplicationID								
										WHERE (DEL.VersionDate <> INS.VersionDate)
										AND INS.VersionDate Is Not Null
										AND DEL.SPVersion Is Not Null	
										AND SPA.IsDeleted =0
									)
		--Use Server DATE Time for Version Date on Updates
		UPDATE Application SET 	Application.VersionDate = GetDate()
		FROM Application  INNER JOIN INSERTED INS On Application.ApplicationID = INS.ApplicationID INNER JOIN DELETED DEL On INS.ApplicationID = DEL.ApplicationID

		--Besure that the Version Date on Development matches the Version Date On Production
		UPDATE DevWebV2..Application SET  
					DevWebV2..Application.[VersionDate] = ProdWebV2..Application.[VersionDate],
					DevWebV2..Application.[DatelastUpdated] = ProdWebV2..Application.[DatelastUpdated] 
		FROM DevWebV2..Application 
			INNER JOIN INSERTED INS On DevWebV2..Application.[AppName] = INS.[AppName] 
			INNER JOIN DELETED DEL On INS.[ApplicationID] = DEL.[ApplicationID] 
			INNER JOIN ProdWebV2..Application On ProdWebV2..Application.[AppName] = DEL.[AppName]
	END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE TRIGGER updAssignmentTypeHistory
ON dbo.AssignmentType
FOR UPDATE
AS
INSERT INTO AssignmentTypeHistory
	SELECT del.* 
	FROM DELETED del

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

--Because of the Restriction on Triggers when Ntext fields ...
/*
In a DELETE, INSERT, or UPDATE trigger, SQL Server does not allow text, ntext, or image column references 
in the inserted and deleted tables if the compatibility level is equal to 70. The text, ntext, and image values 
in the inserted and deleted tables cannot be accessed. To retrieve the new value in either an INSERT or UPDATE 
trigger, join the inserted table with the original update table. When the compatibility level is 65 or lower, 
null values are returned for inserted or deleted text, ntext, or image columns that allow null values; zero-length 
strings are returned if the columns are not nullable. 
If the compatibility level is 80 or higher, SQL Server allows the update of text, ntext, or image columns through 
the INSTEAD OF trigger on tables or views.
*/
--Use Instead OF to get around the above restriction
CREATE TRIGGER updAssignmentsHistory
ON dbo.Assignments
INSTEAD OF UPDATE
AS
INSERT INTO AssignmentsHistory	
	SELECT del.*
	FROM DELETED del
	

-- Now that the History table was updated first...
--Allow the original update to process...

Update Assignments SET
	[ID]= 			INS.ID,
	[AssignmentTypeID]=	INS.AssignmentTypeID,
	[ClientCompanyCatSpecID]=INS.ClientCompanyCatSpecID,
	[AdjusterSpecID]= 	INS.AdjusterSpecID,
	[AdjusterSpecIDDisplay]= INS.AdjusterSpecIDDisplay,
	[SPVersion]=		INS.SPVersion,
	[IBNUM]= 		INS.IBNUM,
	[CLIENTNUM]= 		INS.CLIENTNUM,
	[PolicyNo]= 		INS.PolicyNo,
	[PolicyDescription]= 	INS.PolicyDescription,
	[Insured]= 		INS.Insured,
	[MailingAddress]= 	INS.MailingAddress,
	[MAStreet]= 		INS.MAStreet,
	[MACity]=		INS.MACity,
	[MAState]= 		INS.MAState,
	[MAZIP]= 		INS.MAZIP,
	[MAZIP4]= 		INS.MAZIP4,
	[MAOtherPostCode]=	INS.[MAOtherPostCode],
	[HomePhone]= 		INS.HomePhone,
	[BusinessPhone]= 	INS.BusinessPhone,
	[PropertyAddress]= 	INS.PropertyAddress,
	[PAStreet]= 		INS.PAStreet,
	[PACity]= 		INS.PACity,
	[PAState]= 		INS.PAState,
	[PAZIP]= 		INS.PAZIP,
	[PAZIP4]= 		INS.PAZIP4,
	[PAOtherPostCode]=	INS.[PAOtherPostCode],	
	[MortgageeName]= 	INS.MortgageeName,
	[AgentNo]= 		INS.AgentNo,
	[ReportedBy]= 		INS.ReportedBy,
	[ReportedByPhone]= 	INS.ReportedByPhone,
	[Deductible]= 		INS.Deductible,
	[AppDedClassTypeIDOrder]= INS.AppDedClassTypeIDOrder,
	[LRFormat]= 		INS.LRFormat,
	[LossReport]= 		INS.LossReport,
	[LRPrintedDate]=	INS.LRPrintedDate,
	[DownLoadLossReport]=INS.DownLoadLossReport,
	[UpLoadLossReport]=	INS.UpLoadLossReport,
	[StatusID]= 		INS.StatusID,
	[TypeOfLossID]=		INS.[TypeOfLossID],
	[XactTypeOfLoss]=	INS.[XactTypeOfLoss],
	[SentToXact]=		INS.[SentToXact],
	[LossDate]= 		INS.LossDate,
	[AssignedDate]= 	INS.AssignedDate,
	[ReceivedDate]=		INS.[ReceivedDate],
	[ContactDate]= 		INS.ContactDate,
	[InspectedDate]=	INS.[InspectedDate],
	[CloseDate]= 		INS.CloseDate,
	[Reassigned]= 		INS.Reassigned,
	[DateReassigned]= 	INS.DateReassigned,
	[RAAdjusterSpecID]= 	INS.RAAdjusterSpecID,
	[IsLocked]=		INS.IsLocked,
	[IsDeleted]=		INS.IsDeleted,
	[DownloadMe]=		INS.DownloadME,
	[UpLoadMe]=		INS.UpLoadME,
	[DownloadAll]=		INS.DownloadAll,
	[UpLoadAll]=		INS.UpLoadAll,
	[AdminComments]=	INS.AdminComments,
	[MiscDelimSettings]=	INS.MiscDelimSettings,
	[DateLastUpdated]= 	INS.DateLastUpdated,
	[UpdateByUserID]= 	INS.UpdateByUserID 
FROM Assignments A INNER JOIN INSERTED INS ON A.AssignmentsID = INS.AssignmentsID

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE TRIGGER updBatchesHistory
ON dbo.Batches
FOR UPDATE
AS
INSERT INTO BatchesHistory
	SELECT del.* 
	FROM DELETED del

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE TRIGGER insBillAssignment
ON dbo.BillAssignment
AFTER INSERT
AS
BEGIN
	--Make the IBNUM
	UPDATE BillAssignment
	SET BillAssignment.[IBNUM] = 'SF' + Cast(BillAssignment.[BillAssignmentID] As VarChar(20)) 
	FROM INSERTED INS Inner Join BillAssignment On BillAssignment.[BillAssignmentID] = INS.[BillAssignmentID]

	--Insert the First Billing in BillBillingCount
	INSERT INTO BillBillingCount 
	(
		[BillAssignmentID],
		[Rebill],
		[Supplement],
		[AdminComments],
		[DateLastUpdated],
		[UpdateByUserID]
	)
	SELECT 
		INS.[BillAssignmentID] As [BillAssignmentID],
		0 As [Rebill],
		0 As [Supplement],
		'' As [AdminComments],
		INS.[DateLastUpdated] As [DateLastUpdated],
		INS.[UpdateByUserID] As [UpdateByUserID]
	FROM INSERTED INS Inner Join BillAssignment On BillAssignment.[BillAssignmentID] = INS.[BillAssignmentID]
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE TRIGGER insBillBillingCount
ON dbo.BillBillingCount
AFTER INSERT
AS
IF (SELECT COUNT(INS.[BillBillingCountID]) FROM INSERTED INS) = 1
BEGIN

	--Insert the First IB into IBStateFarm
	INSERT INTO IBStateFarm 
	(
		[BillAssignmentID],
		[BillBillingCountID],
		[lssn],
		[IBNumber],
		[PolicyNo],
		[Insured],
		[LossLoc1],
		[LossLoc2],
		[LossLocCity],
		[LossLocState],
		[LossLocZipcode],
		[LossDate],
		[CloseDate],
		[GrossLoss],
		[Supplement],
		[SupplementExplain],
		[AdditionalLoss],
		[Rebilled],
		[OrigIBIBNumber],
		[OrigIBTotalFee],
		[RebillExplain],
		[MultiClaimBldgUnitNum],
		[ClientCompanyCatSpecID],
		[SeverityCode],
		[ServiceFeeBase],
		[ServiceFeeCovAExterior],
		[ServiceFeeCovAFraming],
		[ServiceFeeCovAInterior],
		[ServiceFeeCovB],
		[ServiceFeeALE],
		[OutBuildCount],
		[OutBuildPerItemCharge],
		[ServiceFeeOutBuildings],
		[ServiceFeeSteepCharge],
		[ServiceFeeTwoStory],
		[ServiceFeeMoreThan50Squares],
		[ServiceFeeWoodSlateTileConRoof],
		[ServiceFeeAdditionalDamage],
		[ServiceFeeRopeAndHarness],
		[ServiceFeeMisc],
		[MiscFeesExplain],
		[ServiceFeeTotal],
		[ExpensePagerPhoneExplain],
		[ExpensePagerPhone],
		[ExpenseOtherExplain],
		[ExpenseOther],
		[SumTotalServiceFeeAndExpense],
		[TaxPercent],
		[TaxesTotal],
		[TotalFee],
		[Void],
		[Comments],
		[AdminComments],
		[DateLastUpdated],
		[UpdateByUserID]
	)
	SELECT 
		INS.[BillAssignmentID] As [BillAssignmentID],		--  [int] NOT NULL ,
		INS.[BillBillingCountID] As [BillBillingCountID],	--  [int] NOT NULL ,
		Cast(USERS.[SSN] As Numeric(9,0)) As [lssn],	--  [numeric](9, 0) NULL ,
		BASS.[IBNUM] As [IBNumber],			--  [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL  ,
		BASS.[PolicyNo] As [PolicyNo],			--  [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL  ,
		BASS.[Insured] As [Insured],			--  [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL  ,
		BASS.[LossLoc1] As [LossLoc1],			--  [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL  ,
		BASS.[LossLoc2] As [LossLoc2],			--  [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL  ,
		BASS.[LossLocCity] As [LossLocCity],		--  [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL  ,
		BASS.[LossLocState] As [LossLocState],		--  [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL  ,
		BASS.[LossLocZipcode] As [LossLocZipcode],	--  [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL  ,
		BASS.[LossDate] As [LossDate],			--  [datetime] NULL ,
		--Leave the CloseDate Null on the IB since this is 
		--what will trigger the insert into Batches for Ebill processing
		Null As [CloseDate],				--  [datetime] NULL ,
		0 As [GrossLoss],				--  [money] NOT NULL  ,
		INS.[Supplement] As [Supplement],		--  [int] NOT NULL ,
		'' As [SupplementExplain],			--  [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL  ,
		0 As [AdditionalLoss],				--  [money] NOT NULL  ,
		INS.[Supplement] As [Rebilled],				--  [int] NOT NULL ,
		'' As [OrigIBIBNumber],				--  [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
		0 As [OrigIBTotalFee],				--  [money] NULL ,
		'' As [RebillExplain],				--  [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
		'' As [MultiClaimBldgUnitNum],			--  [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
		BASS.[ClientCompanyCatSpecID] As [ClientCompanyCatSpecID],		--  [int] NOT NULL ,
		1 As [SeverityCode],				--  [int] NOT NULL ,
		0 As [ServiceFeeBase],				--  [money] NOT NULL ,
		0 As [ServiceFeeCovAExterior] ,			--  [money] NOT NULL ,
		0 As [ServiceFeeCovAFraming],			--  [money] NOT NULL ,
		0 As [ServiceFeeCovAInterior],			--  [money] NOT NULL ,
		0 As [ServiceFeeCovB],				--  [money] NOT NULL ,
		0 As [ServiceFeeALE],				--  [money] NOT NULL ,
		0 As [OutBuildCount],				--  [int] NOT NULL ,
		0 As [OutBuildPerItemCharge] ,			--  [money] NOT NULL ,
		0 As [ServiceFeeOutBuildings],			--  [money] NOT NULL ,
		0 As [ServiceFeeSteepCharge],			--  [money] NOT NULL ,
		0 As [ServiceFeeTwoStory] ,			--  [money] NOT NULL ,
		0 As [ServiceFeeMoreThan50Squares],		--  [money] NOT NULL ,
		0 As [ServiceFeeWoodSlateTileConRoof],		--  [money] NOT NULL ,
		0 As [ServiceFeeAdditionalDamage],		--  [money] NOT NULL ,	
		0 As [ServiceFeeRopeAndHarness],		--  [money] NOT NULL ,
		0 As [ServiceFeeMisc],				--  [money] NOT NULL ,
		'' As [MiscFeesExplain] ,			--  [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
		0 As [ServiceFeeTotal] ,			--  [money] NOT NULL ,
		'' As [ExpensePagerPhoneExplain],		--  [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
		0 As [ExpensePagerPhone] ,			--  [money] NOT NULL ,
		'' As [ExpenseOtherExplain],			--  [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
		0 As [ExpenseOther] ,				--  [money] NOT NULL ,
		8.250 As [SumTotalServiceFeeAndExpense] ,	--  [money] NOT NULL ,
		0 As [TaxPercent],				--  [numeric](8, 3) NOT NULL ,
		0 As [TaxesTotal],				--  [money] NOT NULL ,
		0 As [TotalFee],				--  [money] NOT NULL ,
		0 As [Void],					--  [bit] NOT NULL ,
		'' As [Comments],				--  [varchar] (150) COLLATE SQL_Latin1_General_CP1_CI_AS  NULL ,
		'' As [AdminComments],				--  [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
		INS.[DateLastUpdated] As [DateLastUpdated],	--  [datetime] NOT NULL ,
		INS.[UpdateByUserID] As [UpdateByUserID]	--  [int] NOT NULL 
	FROM INSERTED INS 
		Inner Join BillAssignment BASS On BASS.[BillAssignmentID] = INS.[BillAssignmentID]
		Inner Join AssignmentType AssType On AssType.[AssignmentTypeID] = BASS.[AssignmentTypeID]
		Inner Join ClientCompanyCatSpec CCCS On CCCS.[ClientCompanyCatSpecID] = BASS.[ClientCompanyCatSpecID]
		Inner Join ClientCoAdjusterSpec CCAS On CCAS.[ClientCoAdjusterSpecID] = BASS.[AdjusterSpecID]
		Inner Join Users ON Users.[UsersID] = CCAS.[UsersID]
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE TRIGGER updCATHistory
ON dbo.CAT
FOR UPDATE
AS
INSERT INTO CATHistory
	SELECT del.* 
	FROM DELETED del


---When the Active and INactive dates Change need to apply the changes 
--as appropriate to the ClientCompanyCats associated with the CAT table

--InactiveDate
UPDATE ClientCompanyCat set ClientCompanyCat.inactivedate = (	CASE 	WHEN cc.inactivedate is null 
								THEN INS.inactivedate 
								ELSE (	CASE 	WHEN cc.inactivedate > INS.inactivedate
										THEN INS.inactivedate
										ELSE cc.inactivedate
										END
									)
								END
							),
			ClientCompanyCat.DateLastUpdated = INS.DateLastUpdated,
			ClientCompanyCat.UpdateByUserID = INS.UpdateByUserID
FROM  INSERTED INS 
	INNER JOIN DELETED DEL ON INS.CatID = DEL.CatID AND (INS.inactivedate <> DEL.inactivedate  Or (DEL.inactivedate Is null And INS.inactivedate Is Not Null) Or (DEL.inactivedate Is Not null And INS.inactivedate Is Null)) 
	INNER JOIN ClientCompanyCat cc ON cc.catid = INS.catid

--ActiveDate
UPDATE ClientCompanyCat set ClientCompanyCat.activedate = (	CASE 	WHEN cc.activedate is null 
								THEN INS.activedate 
								ELSE(	CASE	WHEN  cc.activedate < INS.activedate
										THEN INS.activedate
										ELSE cc.activedate
										END
									)
								END
							),
			ClientCompanyCat.DateLastUpdated = INS.DateLastUpdated,
			ClientCompanyCat.UpdateByUserID = INS.UpdateByUserID
FROM  INSERTED INS 
	INNER JOIN DELETED DEL ON INS.CatID = DEL.CatID AND (INS.activedate <> DEL.activedate  Or (DEL.activedate Is null And INS.activedate Is Not Null) Or (DEL.activedate Is Not null And INS.activedate Is Null)) 
	INNER JOIN ClientCompanyCat cc ON cc.catid = INS.catid

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE TRIGGER updClassOfLossHistory
ON dbo.ClassOfLoss
FOR UPDATE
AS
INSERT INTO ClassOfLossHistory
	SELECT del.* 
	FROM DELETED del

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE TRIGGER updClassTypeHistory
ON dbo.ClassType
FOR UPDATE
AS
INSERT INTO ClassTypeHistory
	SELECT del.* 
	FROM DELETED del

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE TRIGGER updClientCoAdjusterSpecHistory
ON dbo.ClientCoAdjusterSpec
FOR UPDATE
AS
INSERT INTO ClientCoAdjusterSpecHistory
	SELECT del.* 
	FROM DELETED del

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE TRIGGER insClientCoAdjusterSpec
ON dbo.ClientCoAdjusterSpec 
AFTER INSERT
AS
--After Adding ACID (Adjuster Client Company identification) Update the Inactive date of Duplicate ACID when applicable
UPDATE ClientCoAdjusterSpec SET
			[InactiveDate] = (CASE 	WHEN CCS.[ClientCoAdjusterSpecID] <> INS.[ClientCoAdjusterSpecID] 
					  	THEN INS.[ActiveDate]
						ELSE CCS.[InactiveDate]
						END
					),
			[DateLastUpdated] = INS.[DateLastUpdated],
			[UpdateByUserID] = INS.[UpdateByUserID]
FROM 	ClientCoAdjusterSpec CCS INNER JOIN INSERTED INS ON CCS.ClientCompanyID = INS.ClientCompanyID 
			AND CCS.ACID = INS.ACID
WHERE CCS.InactiveDate Is Null

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE TRIGGER insClientCompanyCat
ON dbo.ClientCompanyCat 
AFTER INSERT
AS
--After Adding Company Cats Update the Software Package Table 
insert into SoftwarePackage (
	[ClientCompanyID],
	[CatID],
	[PackageName] ,
	[Description] ,
	[SPVersion] ,
	[VersionDate],
	[IsDeleted],
	[DateLastUpdated] ,
	[UpdateByUserID] 
	)
SELECT
	INS.[ClientCompanyID],
	INS.[CatID],
	'CAT_SP_' + Cat.Name As [PackageName],
	'Software Package for ' + C.Name + ' For CAT ' + CAT.Name As [Description],
	(CASE	WHEN (SELECT MAX(SPVersion) FROM SoftwarePackage) IS Null
		THEN 1
		ELSE (SELECT MAX(SPVersion) FROM SoftwarePackage)
		END
	) as [SPVersion] ,
	GetDate() as [VersionDate],
	0 As [IsDeleted],
	GetDate() As [DateLastUpdated] ,
	INS.[UpdateByUserID]
FROM 	INSERTED INS 	INNER JOIN CAT ON INS.CATID = CAT.CATID 
			INNER JOIN Company C ON INS.ClientCompanyID = C.CompanyID
WHERE C.IsClientOF Is Not Null

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE TRIGGER updClientCompanyCatHistory
ON dbo.ClientCompanyCat
FOR UPDATE
AS
INSERT INTO ClientCompanyCatHistory
	SELECT del.* 
	FROM DELETED del

--When the INactive dates ONLY !!! Change for the Client Company Cat....
--need to apply the changes as appropriate to the ClientCompanyCatSPecs associated with the ClientCompanyCat table
--InactiveDate
UPDATE ClientCompanyCatSpec set 
	ClientCompanyCatSpec.inactivedate = 	(	CASE 	WHEN cccs.inactivedate is null 
							THEN INS.inactivedate 
							ELSE (	CASE 	WHEN cccs.inactivedate > INS.inactivedate
									THEN INS.inactivedate
									ELSE cccs.inactivedate
									END
								)
							END
						),
	ClientCompanyCatSpec.DateLastUpdated = INS.DateLastUpdated,
	ClientCompanyCatSpec.UpdateByUserID = INS.UpdateByUserID
FROM  INSERTED INS 
	INNER JOIN DELETED DEL ON 
			INS.CatID = DEL.CatID 
			AND INS.ClientCompanyID = DEL.ClientCompanyID 
			AND (INS.inactivedate <> DEL.inactivedate  Or (DEL.inactivedate Is null And INS.inactivedate Is Not Null) Or (DEL.inactivedate Is Not null And INS.inactivedate Is Null)) 
	INNER JOIN ClientCompanyCatSpec cccs ON 
			cccs.catid = INS.catid
			AND cccs.ClientCompanyID = INS.ClientCompanyID

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE TRIGGER insClientCompanyCatSpec
ON dbo.ClientCompanyCatSpec 
AFTER INSERT
AS
--After Adding Cat code Update the Inactive date of Duplicate Cat code when applicable
UPDATE ClientCompanyCatSpec SET
			[InactiveDate] = (CASE 	WHEN CCS.[ClientCompanyCatSpecID] <> INS.[ClientCompanyCatSpecID] 
					  	THEN INS.[ActiveDate]
						ELSE CCS.[InactiveDate]
						END
					),
			[DateLastUpdated] = INS.[DateLastUpdated],
			[UpdateByUserID] = INS.[UpdateByUserID]
FROM 	ClientCompanyCatSpec CCS INNER JOIN INSERTED INS ON CCS.ClientCompanyID = INS.ClientCompanyID 
			AND CCS.CatCode = INS.CatCode
WHERE CCS.InactiveDate Is Null

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE TRIGGER updClientCompanyCatSpecHistory
ON dbo.ClientCompanyCatSpec
FOR UPDATE
AS
INSERT INTO ClientCompanyCatSpecHistory
	SELECT del.* 
	FROM DELETED del

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE TRIGGER updCompanyHistory
ON dbo.Company
FOR UPDATE
AS
INSERT INTO CompanyHistory
	SELECT del.* 
	FROM DELETED del

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE TRIGGER updCompanyUsersHistory
ON dbo.CompanyUsers
FOR UPDATE
AS
INSERT INTO CompanyUsersHistory
	SELECT del.* 
	FROM DELETED del

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE TRIGGER insDocument
 ON dbo.Document 
INSTEAD OF INSERT
AS

--As Well Update this records SPName and VersionDate And SPVersion and SPVersionBase
	INSERT INTO  Document (
					[DocNameBase],
					[DocName],
					[Description],
					[Version] ,
					[SPVersionBase] ,
					[SPVersion] ,
					[VersionDate],
					[SectionLevel01],
					[SectionLevel02],
					[SectionLevel03],
					[SectionLevel04],
					[SectionLevel05],
					[InstallFileLocation],
					[SPName] ,
					[IsDeleted],
					[DateLastUpdated] ,
					[UpdateByUserID] 
				)
	SELECT 		
					INS.[DocNameBase],
					INS.[DocName],
					INS.[Description],
					INS.[Version],
					(SELECT Max(SPVersion) FROM SoftwarePackage) As [SPVersionBase] ,
					(SELECT Max(SPVersion) FROM SoftwarePackage) As [SPVersion] ,
					GetDate() as [VersionDate],
					INS.[SectionLevel01],
					INS.[SectionLevel02],
					INS.[SectionLevel03],
					INS.[SectionLevel04],
					INS.[SectionLevel05],
					INS.[InstallFileLocation],
					INS.[DocName] + '_V' + cast(INS.[Version] As VarChar(10)) + '.exe' As [SPName] ,
					INS.[IsDeleted],
					Getdate() as [DateLastUpdated] ,
					INS.[UpdateByUserID] 	
	FROM INSERTED INS

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE TRIGGER updDocumentHistory
ON dbo.Document
AFTER UPDATE
AS
--See if there are any Version Changes
DECLARE @CountVersion int
DECLARE @CountVersionDate int

SET @CountVersion = 	(
				SELECT Count(DEL.Version) As CountOFVersion
				FROM DELETED DEL INNER JOIN INSERTED INS
				ON DEL.DocumentID = INS.DocumentID
				WHERE (DEL.Version <> INS.Version)
				AND DEL.SPVersion Is Not Null	
			)
SET @CountVersionDate = 	(
				SELECT Count(DEL.VersionDate) As CountOFVersionDate
				FROM DELETED DEL INNER JOIN INSERTED INS
				ON DEL.DocumentID = INS.DocumentID
				WHERE (DEL.VersionDate <> INS.VersionDate)
				AND INS.VersionDate Is Not Null 
				AND DEL.SPVersion Is Not Null	
				)

IF @CountVersion > 0
BEGIN
	INSERT INTO DocumentHistory
		SELECT DEL.* 
		FROM DELETED DEL INNER JOIN INSERTED INS
		ON DEL.DocumentID = INS.DocumentID
		WHERE (DEL.Version <> INS.Version )
		AND DEL.SPVersion Is Not Null	

	--Then update the Softwarepackage Table which will in turn Update All dependant Tables
	--with new SPVersion.
	UPDATE SoftwarePackage SET SoftwarePackage.SPVersion = (SELECT MAX(SoftwarePackage.SPVersion) +1  FROM SoftwarePackage),
					SoftwarePackage.VersionDate = GetDate()
	FROM SoftwarePackage 
	WHERE SoftwarePackage.SoftwarePackageID IN 	(
									SELECT SPD.SoftwarePackageID
									FROM SoftwarePackageDocument SPD INNER JOIN (INSERTED INS INNER JOIN DELETED DEL ON INS.DocumentID = DEL.DocumentID) ON SPD.DocumentID = INS.DocumentID								
									WHERE DEL.Version <> INS.Version
									AND SPD.IsDeleted =0
								)
	--As Well Update this records SPName, VersionDate and SPVersionBase
	UPDATE Document SET 	Document.SPName = Document.DocName + '_V' + cast(Document.Version As VarChar(10)) + '.exe',
					Document.VersionDate = GetDate(),
					Document.SPVersionBase = Document.SPVersion
	FROM Document  INNER JOIN INSERTED INS On Document.DocumentID = INS.DocumentID INNER JOIN DELETED DEL On INS.DocumentID = DEL.DocumentID
END

ELSE
	IF @CountVersionDate > 0
	BEGIN
		--Update the software package VersionDate
		UPDATE SoftwarePackage SET 	SoftwarePackage.VersionDate = getdate()
		FROM SoftwarePackage 
		WHERE SoftwarePackage.SoftwarePackageID IN 	(
										SELECT SPD.SoftwarePackageID
										FROM SoftwarePackageDocument SPD INNER JOIN (INSERTED INS INNER JOIN DELETED DEL ON INS.DocumentID = DEL.DocumentID) ON SPD.DocumentID = INS.DocumentID								
										WHERE (DEL.VersionDate <> INS.VersionDate)
										AND INS.VersionDate Is Not Null
										AND DEL.SPVersion Is Not Null	
										AND SPD.IsDeleted =0
									)
		--Use Server DATE Time for Version Date on Updates
		UPDATE Document SET 	Document.VersionDate = GetDate()
		FROM Document  INNER JOIN INSERTED INS On Document.DocumentID = INS.DocumentID INNER JOIN DELETED DEL On INS.DocumentID = DEL.DocumentID
		
		--Besure that the Version Date on Development matches the Version Date On Production
		UPDATE DevWebV2..Document SET  
					DevWebV2..Document.[VersionDate] = ProdWebV2..Document.[VersionDate],
					DevWebV2..Document.[DatelastUpdated] = ProdWebV2..Document.[DatelastUpdated] 
		FROM DevWebV2..Document 
			INNER JOIN INSERTED INS On DevWebV2..Document.[DocName] = INS.[DocName] 
			INNER JOIN DELETED DEL On INS.[DocumentID] = DEL.[DocumentID] 
			INNER JOIN ProdWebV2..Document On ProdWebV2..Document.[DocName] = DEL.[DocName]
	END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE TRIGGER updFAQSHistory
ON dbo.FAQS
--Use "Instead OF UPDATE" to get around TEXT MEMO BLOB Restrictions
INSTEAD OF UPDATE
AS
INSERT INTO FAQSHistory
	SELECT del.* 
	FROM DELETED del
-- Now that the History table was updated first...
--Allow the original update to process...

Update FAQS SET
	[Question] = INS.Question,
	[Answer] = INS.Answer,
	[IsDeleted] = INS.IsDeleted,
	[DateLastUpdated] = INS.DateLastUpdated,
	[UpdateByUserID] =INS.UpdateByUserID
FROM FAQS INNER JOIN INSERTED INS ON FAQS.FAQSID = INS.FAQSID

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE TRIGGER updFtpArchive
ON FTPlog
FOR INSERT AS

      INSERT INTO FTPLogArchive(
				[FTPLogID] ,
				[ClientHost],
				[username],
				[LogTime] ,
				[service] ,
				[machine] ,
				[serverip] ,
				[processingtime],
				[bytesrecvd] ,
				[bytessent] ,
				[servicestatus] ,
				[win32status] ,
				[operation] ,
				[target] ,
				[parameters]
				)
         	SELECT 		
				[FTPLogID] ,
				[ClientHost],
				[username] ,
				[LogTime] ,
				[service] ,
				[machine] ,
				[serverip] ,
				[processingtime],
				[bytesrecvd] ,
				[bytessent] ,
				[servicestatus],
				[win32status] ,
				[operation] ,
				[target] ,
				[parameters]
         	FROM FTPlog
		WHERE LogTime < getdate()-90 --Select records older than 90 days
		--Then Delete from HTTP Log
		DELETE FROM FTPLog 
		WHERE LogTime < getdate()-90 --Select records older than 90 days


	--Do the http here too instead of in the http update
	INSERT INTO HTTPLogArchive(
				[HTTPLogID] ,
				[ClientHost],
				[username],
				[LogTime] ,
				[service] ,
				[machine] ,
				[serverip] ,
				[processingtime],
				[bytesrecvd] ,
				[bytessent] ,
				[servicestatus] ,
				[win32status] ,
				[operation] ,
				[target] ,
				[parameters]
				)
         	SELECT 		
				[HTTPLogID] ,
				[ClientHost],
				[username] ,
				[LogTime] ,
				[service] ,
				[machine] ,
				[serverip] ,
				[processingtime],
				[bytesrecvd] ,
				[bytessent] ,
				[servicestatus],
				[win32status] ,
				[operation] ,
				[target] ,
				[parameters]
         	FROM HTTPlog
		WHERE LogTime < getdate()-90 --Select records older than 90 days
		--Then Delete from HTTP Log
		DELETE FROM HTTPLog 
		WHERE LogTime < getdate()-90 --Select records older than 90 days






GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE TRIGGER updFeeScheduleHistory
ON dbo.FeeSchedule
FOR UPDATE
AS
INSERT INTO FeeScheduleHistory
	SELECT del.* 
	FROM DELETED del

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE TRIGGER updFeeScheduleFeeTypesHistory
ON dbo.FeeScheduleFeeTypes
FOR UPDATE
AS
INSERT INTO FeeScheduleFeeTypesHistory
	SELECT del.* 
	FROM DELETED del

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE TRIGGER updFeeScheduleLevelsHistory
ON dbo.FeeScheduleLevels
FOR UPDATE
AS
INSERT INTO FeeScheduleLevelsHistory
	SELECT del.* 
	FROM DELETED del

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE TRIGGER updIBHistory
ON dbo.IB
FOR UPDATE
AS
INSERT INTO IBHistory
	SELECT del.* 
	FROM DELETED del

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE TRIGGER updIB
ON dbo.IB
AFTER UPDATE
AS
--1 After Updating an IB Insert Applicable record into Batches
--This is the Fee Bill that will be sent to Billings
--Only insert into batches if the ID of IB is > 0
--And the Count of inserted is 1.  This Trigger only handles single record Updates on the IB table Not Multi Updates
--This means the record has successfully been Synched with Client
IF (SELECT COUNT(INS.[ID]) FROM INSERTED INS) = 1
BEGIN
	IF (SELECT INS.[ID] FROM INSERTED INS) > 0
	BEGIN
		DECLARE @CurDate 	DateTime
		DECLARE @MYDATE		DateTime
		DECLARE @MYIBNUM 	VarChar(50)
		DECLARE @VBCRLF		VarChar(2)
		SET @CurDate = GetDate()
		SET @VBCRLF = dbo.GetVBCRLF()
		SET @MYDATE = @CurDate
		SET @MYDATE = dbo.CleanFromOrToDateString(@CurDate, @MYDATE, 1)
		SET @MYIBNUM = ( 
				SELECT
					(CASE	WHEN 	INS.[IB14a_sSupplement] > 0 And INS.[IB14b_sRebilled] > 0
						THEN	INS.[IB02_sIBNumber] + 'S' + cast(INS.[IB14a_sSupplement] As varchar(4)) + 'R' + cast(INS.[IB14b_sRebilled] As varchar(4))
						ELSE	(CASE	WHEN 	INS.[IB14a_sSupplement] > 0
								THEN	INS.[IB02_sIBNumber] + 'S' + cast(INS.[IB14a_sSupplement] As varchar(4))
								ELSE	(CASE	WHEN	INS.[IB14b_sRebilled] > 0
										THEN	INS.[IB02_sIBNumber] + 'R' + cast(INS.[IB14b_sRebilled] As varchar(4))
										ELSE	INS.[IB02_sIBNumber]
										END
									)
								END
							)
						END
					) AS [MYIBNUM]
				FROM INSERTED INS
				)		
		INSERT INTO Batches (
			
			[AssignmentsID],		-- [int] NOT NULL ,
			[ClientCompanyCatSpecID],	-- [int] NOT NULL ,
			[ssn],				-- [numeric](9, 0) NULL ,
			[ibnumber],			-- [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
			[date],				-- [datetime] NULL ,
			[EnteredDate],			-- [datetime] NULL ,
			[adj_name],			-- [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
			[adjuster_n],			-- [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
			[claimnumber],			-- [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
			[insuredname],			-- [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
			[loss_loc],			-- [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
			[losscity],			-- [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
			[lossstate],			-- [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
			[dateofloss],			-- [datetime] NULL ,
			[dateclosed],			-- [datetime] NULL ,
			[grossloss],			-- [decimal](20, 5) NULL ,
			[totalservice],			-- [decimal](20, 5) NULL ,
			[administrative],			-- [decimal](20, 5) NULL ,
			[misccharge],			-- [decimal](20, 5) NULL ,
			[taxestotal],			-- [decimal](20, 5) NULL ,
			[totalfee],			-- [decimal](20, 5) NULL ,
			[catsite],			-- [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
			[Void],				-- [bit] NOT NULL ,
			[billingdup],			-- [bit] NULL ,
			[ecupdated],			-- [bit] NULL ,
			[copied],			-- [int] NULL ,
			[duplicate],			-- [bit] NULL ,
			[Comments],			-- [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
			[Reassigned],			-- [int] NULL ,
			[DateLastUpdated],		-- [datetime] NOT NULL ,
			[UpdateByUserID]		-- [int] NOT NULL 
			)
		SELECT
			INS.[AssignmentsID] 				As [AssignmentsID],		--  [int] NOT NULL ,
			(
			SELECT 	[ClientCompanyCatSpecID]
			FROM	Assignments
			WHERE	[AssignmentsID] = INS.[AssignmentsID]
			) 						As [ClientCompanyCatSpecID],	--  [int] NOT NULL ,
			INS.[IB00_lssn]					As [ssn],			-- [numeric](9, 0) NULL ,
			@MYIBNUM					As [ibnumber],			-- [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
			@MYDATE						As [date],			-- [datetime] NULL ,
			Null						As [EnteredDate],		-- [datetime] NULL ,
			''						As [adj_name],			-- [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
			INS.[IB07_sAdjusterName]			As [adjuster_n],		-- [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
			INS.[IB09_sSALN]				As [claimnumber],		-- [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
			INS.[IB10_sInsuredName]				As [insuredname],		-- [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
			REPLACE(INS.[IB11_sLossLocation],@VBCRLF,'    ')	As [loss_loc],			-- [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
			(
			SELECT 	[PACITY]
			FROM 	Assignments
			WHERE	[AssignmentsID] = INS.[AssignmentsID]
			)
									As [losscity],			-- [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
			(
			SELECT 	[PASTATE]
			FROM 	Assignments
			WHERE	[AssignmentsID] = INS.[AssignmentsID]
			)
									As [lossstate],			-- [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
			INS.[IB12_dtDateOfLoss]				As [dateofloss],		-- [datetime] NULL ,
			INS.[IB06_dtDateClosed]				As [dateclosed],		-- [datetime] NULL ,
			INS.[IB13_cGrossLoss]				As [grossloss],			-- [decimal](20, 5) NULL ,
			INS.[IB25_cServiceFeeSubTotal]			As [totalservice],		-- [decimal](20, 5) NULL ,
			0						As [administrative],		-- [decimal](20, 5) NULL ,
			INS.[IB30_cTotalExpenses]			As [misccharge],		-- [decimal](20, 5) NULL ,
			INS.[IB32_cTaxAmount]				As [taxestotal],		-- [decimal](20, 5) NULL ,
			INS.[IB33_cTotalAdjustingFee]			As [totalfee],			-- [decimal](20, 5) NULL ,
			(
		    	SELECT [BillingCode] 
		    	FROM ClientCompanyCat 
		    	WHERE ClientCompanyID = 
		                ( 
		                SELECT   [ClientCompanyID] 
		                FROM     ClientCompanyCatSpec 
		                WHERE    [ClientCompanyCatSpecID] = 	(
									SELECT 	[ClientCompanyCatSpecID]
									FROM	Assignments
									WHERE	[AssignmentsID] = INS.[AssignmentsID]
									) 		
		                )
		    	AND [CATID] = 	(
		                	SELECT   [CatID] 
		                	FROM     ClientCompanyCatSpec 
		                	WHERE    [ClientCompanyCatSpecID] = 	(
										SELECT 	[ClientCompanyCatSpecID]
										FROM	Assignments
										WHERE	[AssignmentsID] = INS.[AssignmentsID]
										) 	
		                	)
		    	) 						As [catsite],			-- [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
			INS.[Void]					As [Void],			-- [bit] NOT NULL ,
			0						As [billingdup],		-- [bit] NULL ,
			0						As [ecupdated],			-- [bit] NULL ,
			0						As [copied],			-- [int] NULL ,
			0						As [duplicate],			-- [bit] NULL ,
			Left(INS.[Comments],100)			As [Comments],			-- [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
			0						As [Reassigned],		-- [int] NULL ,
			GetDate()					As [DateLastUpdated],		-- [datetime] NOT NULL ,
			INS.[UpdateByUserID]				As [UpdateByUserID]		-- [int] NOT NULL 
		FROM 	INSERTED INS
		WHERE 	@MYIBNUM NOT IN	(
						SELECT 	[ibnumber]
						FROM	BATCHES
					)
	END
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE TRIGGER updIBFeeHistory
ON dbo.IBFee
FOR UPDATE
AS
INSERT INTO IBFeeHistory
	SELECT del.* 
	FROM DELETED del

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE TRIGGER updIBStateFarm
ON dbo.IBStateFarm
AFTER UPDATE
AS
--1 After Updating an IB Insert Applicable record into Batches
--This is the Fee Bill that will be sent to Billings
--Only insert into batches if the CloseDate is not null
--And the Count of inserted is 1.  This Trigger only handles single record Updates on the IB table Not Multi Updates
IF (SELECT COUNT(INS.[IBStateFarmID]) FROM INSERTED INS) = 1
BEGIN
	IF (SELECT INS.[CloseDate] FROM INSERTED INS) Is Not Null
	BEGIN
		DECLARE @CurDate 	DateTime
		DECLARE @MYDATE		DateTime
		DECLARE @MYIBNUM 	VarChar(50)
		SET @CurDate = GetDate()
		SET @MYDATE = @CurDate
		SET @MYDATE = dbo.CleanFromOrToDateString(@CurDate, @MYDATE, 1)
		SET @MYIBNUM = ( 
				SELECT
					(CASE	WHEN 	INS.[Supplement] > 0 And INS.[Rebilled] > 0
						THEN	INS.[IBNumber] + 'S' + cast(INS.[Supplement] As varchar(4)) + 'R' + cast(INS.[Rebilled] As varchar(4))
						ELSE	(CASE	WHEN 	INS.[Supplement] > 0
								THEN	INS.[IBNumber] + 'S' + cast(INS.[Supplement] As varchar(4))
								ELSE	(CASE	WHEN	INS.[Rebilled] > 0
										THEN	INS.[IBNumber] + 'R' + cast(INS.[Rebilled] As varchar(4))
										ELSE	INS.[IBNumber]
										END
									)
								END
							)
						END
					) AS [MYIBNUM]
				FROM INSERTED INS
				)
		
		INSERT INTO Batches (
			[AssignmentsID],		-- [int] NULL ,
			[ClientCompanyCatSpecID],	-- [int] NOT NULL ,
			[ssn],				-- [numeric](9, 0) NULL ,
			[ibnumber],			-- [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
			[date],				-- [datetime] NULL ,
			[EnteredDate],			-- [datetime] NULL ,
			[adj_name],			-- [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
			[adjuster_n],			-- [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
			[claimnumber],			-- [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
			[insuredname],			-- [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
			[loss_loc],			-- [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
			[losscity],			-- [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
			[lossstate],			-- [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
			[dateofloss],			-- [datetime] NULL ,
			[dateclosed],			-- [datetime] NULL ,
			[grossloss],			-- [decimal](20, 5) NULL ,
			[totalservice],			-- [decimal](20, 5) NULL ,
			[administrative],		-- [decimal](20, 5) NULL ,
			[misccharge],			-- [decimal](20, 5) NULL ,
			[taxestotal],			-- [decimal](20, 5) NULL ,
			[totalfee],			-- [decimal](20, 5) NULL ,
			[catsite],			-- [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
			[Void],				-- [bit] NOT NULL ,
			[billingdup],			-- [bit] NULL ,
			[ecupdated],			-- [bit] NULL ,
			[copied],			-- [int] NULL ,
			[duplicate],			-- [bit] NULL ,
			[Comments],			-- [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
			[Reassigned],			-- [int] NULL ,
			[DateLastUpdated],		-- [datetime] NOT NULL ,
			[UpdateByUserID],		-- [int] NOT NULL ,
			[BillAssignmentID]		-- [int] NULL 
			)
		SELECT
			-- Since this is a State Farm Ebill Make sure the AssignmentsID is NULL !!!!
			Null 						As [AssignmentsID],		--  [int] NOT NULL ,
			CCCS.[ClientCompanyCatSpecID]			As [ClientCompanyCatSpecID],	--  [int] NOT NULL ,
			INS.[lssn]					As [ssn],			-- [numeric](9, 0) NULL ,
			@MYIBNUM					As [ibnumber],			-- [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
			@MYDATE						As [date],			-- [datetime] NULL ,
			Null						As [EnteredDate],		-- [datetime] NULL ,
			''						As [adj_name],			-- [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
			(	
				Users.[LastName] + ', ' + Users.[FirstName]
			)						As [adjuster_n],		-- [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
			BASS.[CLIENTNUM]				As [claimnumber],		-- [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
			INS.[Insured] 					As [insuredname],		-- [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,

			(
				INS.[LossLoc1] + ' ' + INS.[LossLoc2] + '    ' + INS.[LossLocCity] + ', ' + INS.[LossLocState] + ' ' + INS.[LossLocZipcode] 
			)						As [loss_loc],			-- [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
			INS.[LossLocCity]				As [losscity],			-- [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
			INS.[LossLocState]				As [lossstate],			-- [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
			INS.[LossDate]					As [dateofloss],		-- [datetime] NULL ,
			INS.[CloseDate]					As [dateclosed],		-- [datetime] NULL ,
			INS.[GrossLoss]					As [grossloss],			-- [decimal](20, 5) NULL ,
			INS.[ServiceFeeTotal]				As [totalservice],		-- [decimal](20, 5) NULL ,
			0						As [administrative],		-- [decimal](20, 5) NULL ,
			(
				INS.[ExpensePagerPhone] + INS.[ExpenseOther]
			)						As [misccharge],		-- [decimal](20, 5) NULL ,
			INS.[TaxesTotal]				As [taxestotal],		-- [decimal](20, 5) NULL ,
			INS.[TotalFee]					As [totalfee],			-- [decimal](20, 5) NULL ,
			CCC.[BillingCode]				As [catsite],			-- [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
			INS.[Void]					As [Void],			-- [bit] NOT NULL ,
			0						As [billingdup],		-- [bit] NULL ,
			0						As [ecupdated],			-- [bit] NULL ,
			0						As [copied],			-- [int] NULL ,
			0						As [duplicate],			-- [bit] NULL ,
			Left(INS.[Comments],100)			As [Comments],			-- [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
			0						As [Reassigned],		-- [int] NULL ,
			GetDate()					As [DateLastUpdated],		-- [datetime] NOT NULL ,
			INS.[UpdateByUserID]				As [UpdateByUserID],		-- [int] NOT NULL 
			INS.[BillAssignmentID]			As [BillAssignmentID]	--[int]  NULL
		FROM 	INSERTED INS
			Inner Join BillAssignment BASS On BASS.[BillAssignmentID] = INS.[BillAssignmentID]
			Inner Join AssignmentType AssType On AssType.[AssignmentTypeID] = BASS.[AssignmentTypeID]
			Inner Join ClientCompanyCatSpec CCCS On CCCS.[ClientCompanyCatSpecID] = BASS.[ClientCompanyCatSpecID]
			Inner Join ClientCompanyCat CCC On CCC.[ClientCompanyID] = CCCS.[ClientCompanyID] And CCC.[CATID] = CCCS.[CATID]
			Inner Join ClientCoAdjusterSpec CCAS On CCAS.[ClientCoAdjusterSpecID] = BASS.[AdjusterSpecID]
			Inner Join Users ON Users.[UsersID] = CCAS.[UsersID]
		WHERE 	@MYIBNUM NOT IN	(
						SELECT 	[ibnumber]
						FROM	BATCHES
					)
	END
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE TRIGGER updPackageHistory
ON dbo.Package
FOR UPDATE
AS
INSERT INTO PackageHistory
	SELECT del.* 
	FROM DELETED del

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE TRIGGER updPackageItemHistory
ON dbo.PackageItem
FOR UPDATE
AS
INSERT INTO PackageItemHistory
	SELECT del.* 
	FROM DELETED del

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE TRIGGER updPolicyLimitsHistory
ON dbo.PolicyLimits
FOR UPDATE
AS
INSERT INTO PolicyLimitsHistory
	SELECT del.* 
	FROM DELETED del

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

--Because of the Restriction on Triggers when Ntext fields ...
/*
In a DELETE, INSERT, or UPDATE trigger, SQL Server does not allow text, ntext, or image column references 
in the inserted and deleted tables if the compatibility level is equal to 70. The text, ntext, and image values 
in the inserted and deleted tables cannot be accessed. To retrieve the new value in either an INSERT or UPDATE 
trigger, join the inserted table with the original update table. When the compatibility level is 65 or lower, 
null values are returned for inserted or deleted text, ntext, or image columns that allow null values; zero-length 
strings are returned if the columns are not nullable. 
If the compatibility level is 80 or higher, SQL Server allows the update of text, ntext, or image columns through 
the INSTEAD OF trigger on tables or views.
*/
--Use Instead OF to get around the above restriction
CREATE TRIGGER updRTActivityLogHistory
ON dbo.RTActivityLog
INSTEAD OF UPDATE
AS
INSERT INTO RTActivityLogHistory	
	SELECT del.*
	FROM DELETED del
	

-- Now that the History table was updated first...
--Allow the original update to process...

Update RTActivityLog SET
	[AssignmentsID] = INS.AssignmentsID,
	[BillingCountID] = INS.BillingCountID,
	[ID] = INS.ID,
	[IDAssignments] = INS.IDAssignments,
	[IDBillingCount] = INS.IDBillingCount,
	[ServiceTime] = INS.ServiceTime,
	[ActDate] = INS.ActDate,
	[ActText] = INS.ActText,
	[ActTime] = INS.ActTime,
	[PageBreakAfter] = INS.PageBreakAfter,
	[BlankPageAfter] = INS.BlankPageAfter,
	[BlankRowsAfter] = INS.BlankRowsAfter,
	[IsMgrEntry]= INS.IsMgrEntry,
	[IsDeleted]= INS.IsDeleted,
	[DownloadME]= INS.DownloadMe,
	[UpLoadMe]= INS.UpLoadMe,
	[DateLastUpdated] = INS.DateLastUpdated,
	[UpdateByUserID] =INS.UpdateByUserID
FROM RTActivityLog RTAL INNER JOIN INSERTED INS ON RTAL.RTActivityLogID = INS.RTActivityLogID

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE TRIGGER updRTActivityLogInfoHistory
ON dbo.RTActivityLogInfo
FOR UPDATE
AS
INSERT INTO RTActivityLogInfoHistory
	SELECT del.* 
	FROM DELETED del

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE TRIGGER updRTAttachmentsHistory
ON dbo.RTAttachments
FOR UPDATE
AS
INSERT INTO RTAttachmentsHistory
	SELECT del.* 
	FROM DELETED del

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE TRIGGER updRTChecksHistory
ON dbo.RTChecks
FOR UPDATE
AS
INSERT INTO RTChecksHistory
	SELECT del.* 
	FROM DELETED del

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE TRIGGER updRTIBHistory
ON dbo.RTIB
FOR UPDATE
AS
INSERT INTO RTIBHistory
	SELECT del.* 
	FROM DELETED del

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE TRIGGER updRTIBFeeHistory
ON dbo.RTIBFee
FOR UPDATE
AS
INSERT INTO RTIBFeeHistory
	SELECT del.* 
	FROM DELETED del

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE TRIGGER updRTIndemnityHistory
ON dbo.RTIndemnity
FOR UPDATE
AS
INSERT INTO RTIndemnityHistory
	SELECT del.* 
	FROM DELETED del

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

--Because of the Restriction on Triggers when Ntext fields ...
/*
In a DELETE, INSERT, or UPDATE trigger, SQL Server does not allow text, ntext, or image column references 
in the inserted and deleted tables if the compatibility level is equal to 70. The text, ntext, and image values 
in the inserted and deleted tables cannot be accessed. To retrieve the new value in either an INSERT or UPDATE 
trigger, join the inserted table with the original update table. When the compatibility level is 65 or lower, 
null values are returned for inserted or deleted text, ntext, or image columns that allow null values; zero-length 
strings are returned if the columns are not nullable. 
If the compatibility level is 80 or higher, SQL Server allows the update of text, ntext, or image columns through 
the INSTEAD OF trigger on tables or views.
*/
--Use Instead OF to get around the above restriction
CREATE TRIGGER updRTPhotoLogHistory
ON dbo.RTPhotoLog
INSTEAD OF UPDATE
AS
INSERT INTO RTPhotoLogHistory
	SELECT del.* 
	FROM DELETED del

-- Now that the History table was updated first...
--Allow the original update to process...
Update RTPhotoLog SET
[RTPhotoReportID]=		INS.RTPhotoReportID,
[AssignmentsID]=		INS.AssignmentsID,
[BillingCountID]=		INS.BillingCountID,
[ID]=				INS.ID,
[IDRTPhotoReport]=		INS.IDRTPhotoReport,
[IDAssignments]=		INS.IDAssignments,
[IDBillingCount]=		INS.IDBillingCount,
[PhotoDate]=			INS.PhotoDate,
[SortOrder]=			INS.SortOrder,
[Description]=			INS.Description,
[PhotoName]=			INS.PhotoName,
[Photo]=			INS.Photo,
[DownloadPhoto]=		INS.DownloadPhoto,
[UpLoadPhoto]=			INS.UpLoadPhoto,
[PhotoThumb]=			INS.PhotoThumb,
[DownloadPhotoThumb]=		INS.DownloadPhotoThumb,
[UpLoadPhotoThumb]=		INS.UpLoadPhotoThumb,
[PhotoHighRes]=		INS.PhotoHighRes,
[DownloadPhotoHighRes]=		INS.DownloadPhotoHighRes,
[UploadPhotoHighRes]=		INS.UploadPhotoHighRes,
[IsDeleted]=			INS.IsDeleted,
[DownLoadMe]=			INS.DownLoadMe,
[UpLoadMe]=			INS.UpLoadMe,
[AdminComments]=		INS.AdminComments,
[DateLastUpdated]=		INS.DateLastUpdated,
[UpdateByUserID]=		INS.UpdateByUserID
FROM RTPhotoLog RTPL INNER JOIN INSERTED INS ON RTPL.RTPhotoLogID = INS.RTPhotoLogID

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE TRIGGER updRTPhotoReportHistory
ON dbo.RTPhotoReport
FOR UPDATE
AS
INSERT INTO RTPhotoReportHistory
	SELECT del.* 
	FROM DELETED del

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

--Because of the Restriction on Triggers when Ntext fields ...
/*
In a DELETE, INSERT, or UPDATE trigger, SQL Server does not allow text, ntext, or image column references 
in the inserted and deleted tables if the compatibility level is equal to 70. The text, ntext, and image values 
in the inserted and deleted tables cannot be accessed. To retrieve the new value in either an INSERT or UPDATE 
trigger, join the inserted table with the original update table. When the compatibility level is 65 or lower, 
null values are returned for inserted or deleted text, ntext, or image columns that allow null values; zero-length 
strings are returned if the columns are not nullable. 
If the compatibility level is 80 or higher, SQL Server allows the update of text, ntext, or image columns through 
the INSTEAD OF trigger on tables or views.
*/
--Use Instead OF to get around the above restriction
CREATE TRIGGER updRTWSDiagramHistory
ON dbo.RTWSDiagram
INSTEAD OF UPDATE
AS
INSERT INTO RTWSDiagramHistory
	SELECT del.* 
	FROM DELETED del

-- Now that the History table was updated first...
--Allow the original update to process...
Update RTWSDiagram SET
[AssignmentsID]= 		INS.[AssignmentsID],
[ID]= 				INS.[ID],
[IDAssignments]= 		INS.IDAssignments,
[Name]= 			INS.[Name],
[Description]= 			INS.[Description],
[Number]= 			INS.[Number],
[DiagramPhotoName]= 		INS.[DiagramPhotoName],
[DownloadDiagramPhoto]=		INS.DownloadDiagramPhoto,
[UploadDiagramPhoto]=		INS.UploadDiagramPhoto,
[DiagramXML]= 			INS.[DiagramXML],
[IsDeleted]= 			INS.[IsDeleted],
[DownLoadMe]= 			INS.[DownLoadMe],
[UpLoadMe]= 			INS.[UpLoadMe],
[AdminComments]= 		INS.[AdminComments],
[DateLastUpdated]= 		INS.[DateLastUpdated],
[UpdateByUserID]= 		INS.[UpdateByUserID] 
FROM RTWSDiagram RTWDS INNER JOIN INSERTED INS ON RTWDS.RTWSDiagramID = INS.RTWSDiagramID

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE TRIGGER insRegSetting
 ON dbo.RegSetting 
INSTEAD OF INSERT
AS

--As Well Update this records SPName and VersionDate And SPVersion and SPVersionBase
	INSERT INTO  RegSetting (
					[RegNameBase],
					[RegName],
					[Description],
					[Version] ,
					[SPVersionBase] ,
					[SPVersion] ,
					[VersionDate],
					[SectionLevel01],
					[SectionLevel02],
					[SectionLevel03],
					[SectionLevel04],
					[SectionLevel05],
					[SPName] ,
					[IsDeleted],
					[DateLastUpdated] ,
					[UpdateByUserID] 
				)
	SELECT 		
					INS.[RegNameBase],
					INS.[RegName],
					INS.[Description],
					INS.[Version],
					(SELECT Max(SPVersion) FROM SoftwarePackage) As [SPVersionBase] ,
					(SELECT Max(SPVersion) FROM SoftwarePackage) As [SPVersion] ,
					GetDate() as [VersionDate],
					INS.[SectionLevel01],
					INS.[SectionLevel02],
					INS.[SectionLevel03],
					INS.[SectionLevel04],
					INS.[SectionLevel05],
					INS.RegName + '_V' + cast(INS.Version As VarChar(10)) + '.exe' As [SPName] ,
					INS.[IsDeleted],
					Getdate() as [DateLastUpdated] ,
					INS.[UpdateByUserID] 	
	FROM INSERTED INS

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE TRIGGER updRegSettingHistory
ON dbo.RegSetting
AFTER UPDATE
AS
--See if there are any Version Changes
DECLARE @CountVersion int
DECLARE @CountVersionDate int

SET @CountVersion = 	(
				SELECT Count(DEL.Version) As CountOFVersion
				FROM DELETED DEL INNER JOIN INSERTED INS
				ON DEL.RegSettingID = INS.RegSettingID
				WHERE (DEL.Version <> INS.Version)
				AND DEL.SPVersion Is Not Null	
			)
SET @CountVersionDate = 	(
				SELECT Count(DEL.VersionDate) As CountOFVersionDate
				FROM DELETED DEL INNER JOIN INSERTED INS
				ON DEL.RegSettingID = INS.RegSettingID
				WHERE (DEL.VersionDate <> INS.VersionDate)
				AND INS.VersionDate Is Not Null
				AND DEL.SPVersion Is Not Null	
				)

IF @CountVersion > 0
BEGIN
	INSERT INTO RegSettingHistory
		SELECT DEL.* 
		FROM DELETED DEL INNER JOIN INSERTED INS
		ON DEL.RegSettingID = INS.RegSettingID
		WHERE (DEL.Version <> INS.Version )
		AND DEL.SPVersion Is Not Null	

	--Then update the Softwarepackage Table which will in turn Update All dependant Tables
	--with new SPVersion.
	UPDATE SoftwarePackage SET SoftwarePackage.SPVersion =  (SELECT MAX(SoftwarePackage.SPVersion) +1  FROM SoftwarePackage),
					SoftwarePackage.VersionDate = getdate()
	FROM SoftwarePackage 
	WHERE SoftwarePackage.SoftwarePackageID IN 	(
									SELECT SPR.SoftwarePackageID
									FROM SoftwarePackageRegSetting SPR INNER JOIN (INSERTED INS INNER JOIN DELETED DEL ON INS.RegSettingID = DEL.RegSettingID) ON SPR.RegSettingID = INS.RegSettingID								
									WHERE DEL.Version <> INS.Version
									AND SPR.IsDeleted =0
								)
	--As Well Update this records SPName, VersionDate and SPVersionBase
	UPDATE RegSetting SET 	RegSetting.SPName = RegSetting.RegName + '_V' + cast(RegSetting.Version As VarChar(10)) + '.exe',
					RegSetting.VersionDate = GetDate(),
					RegSetting.SPVersionBase = RegSetting.SPVersion
	FROM RegSetting  INNER JOIN INSERTED INS On RegSetting.RegSettingID = INS.RegSettingID INNER JOIN DELETED DEL On INS.RegSettingID = DEL.RegSettingID
	
END

ELSE
	IF @CountVersionDate > 0
	BEGIN
		--Update the software package VersionDate
		UPDATE SoftwarePackage SET 	SoftwarePackage.VersionDate = getdate()
		FROM SoftwarePackage 
		WHERE SoftwarePackage.SoftwarePackageID IN 	(
										SELECT SPR.SoftwarePackageID
										FROM SoftwarePackageRegSetting SPR INNER JOIN (INSERTED INS INNER JOIN DELETED DEL ON INS.RegSettingID = DEL.RegSettingID) ON SPR.RegSettingID = INS.RegSettingID								
										WHERE (DEL.VersionDate <> INS.VersionDate)
										AND INS.VersionDate Is Not Null
										AND DEL.SPVersion Is Not Null	
										AND SPR.IsDeleted =0
									)
		--Use Server DATE Time for Version Date on Updates
		UPDATE RegSetting SET RegSetting.VersionDate = GetDate()
		FROM RegSetting  INNER JOIN INSERTED INS On RegSetting.RegSettingID = INS.RegSettingID INNER JOIN DELETED DEL On INS.RegSettingID = DEL.RegSettingID

		--Besure that the Version Date on Development matches the Version Date On Production
		UPDATE DevWebV2..RegSetting SET  
					DevWebV2..RegSetting.[VersionDate] = ProdWebV2..RegSetting.[VersionDate],
					DevWebV2..RegSetting.[DatelastUpdated] = ProdWebV2..RegSetting.[DatelastUpdated] 
		FROM DevWebV2..RegSetting 
			INNER JOIN INSERTED INS On DevWebV2..RegSetting.[RegName] = INS.[RegName] 
			INNER JOIN DELETED DEL On INS.[RegSettingID] = DEL.[RegSettingID] 
			INNER JOIN ProdWebV2..RegSetting On ProdWebV2..RegSetting.[RegName] = DEL.[RegName]
	END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE TRIGGER updSecurityLevelHistory
ON dbo.SecurityLevel
FOR UPDATE
AS
INSERT INTO SecurityLevelHistory
	SELECT del.* 
	FROM DELETED del
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO



CREATE TRIGGER updSoftwarePackageHistory
ON dbo.SoftwarePackage
AFTER UPDATE
AS
--Update History table
INSERT INTO  SoftwarePackageHistory
	SELECT del.* 
	FROM DELETED del

--See if there are any Version Changes
DECLARE @CountSPVersion int

SET @CountSPVersion = (
				SELECT Count(DEL.SPVersion) As CountOFSPVersion
				FROM DELETED DEL INNER JOIN INSERTED INS
				ON DEL.SoftwarePackageID = INS.SoftwarePackageID
				WHERE DEL.SPVersion <> INS.SPVersion
			)
IF @CountSPVersion > 0
BEGIN

	--RegSetting...
	UPDATE RegSetting SET 	RegSetting.SPVersion = INS.SPVersion,
					RegSetting.SPVersionBase =	(CASE 	WHEN RegSetting.SPVersionBase IS Null Or RegSetting.SPVersionBase > INS.SPVersion
								  	THEN INS.SPVersion
									ELSE RegSetting.SPVersionBase
									END
								)
	FROM INSERTED INS INNER JOIN  SoftwarePackageRegSetting SPR ON INS.SoftwarePackageID = SPR.SoftwarePackageID
	
	--Also update all dependant tables with new SPVersion
	--Documents...
	UPDATE Document SET 	Document.SPVersion = INS.SPVersion,
					Document.SPVersionBase =	(CASE 	WHEN Document.SPVersionBase IS Null Or Document.SPVersionBase >  INS.SPVersion
								  	THEN INS.SPVersion
									ELSE Document.SPVersionBase
									END
								)
	FROM INSERTED INS INNER JOIN  SoftwarePackageDocument SPD ON INS.SoftwarePackageID = SPD.SoftwarePackageID
	
	
	--Application...
	UPDATE Application SET 	Application.SPVersion = INS.SPVersion,
					Application.SPVersionBase =	(CASE 	WHEN Application.SPVersionBase IS Null OR Application.SPVersionBase > INS.SPVersion
								  	THEN INS.SPVersion
									ELSE Application.SPVersionBase
									END
								)
	FROM INSERTED INS INNER JOIN  SoftwarePackageApplication SPA ON INS.SoftwarePackageID = SPA.SoftwarePackageID
	
	
	--Update SPVersion in Software package
	UPDATE SoftwarePackage Set SoftwarePackage.SPVersion = INS.SPVersion
	FROM INSERTED INS
	
	
END




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE TRIGGER insSoftwarePackageApplication
ON dbo.SoftwarePackageApplication 
AFTER INSERT
AS

	UPDATE Application set 	Application.SPVersion = SP.SPVersion,
			    	Application.SPVersionBase =	(CASE 	WHEN Application.SPVersionBase IS Null
								  	THEN SP.SPVersion
									ELSE Application.SPVersionBase
									END
								)
						
	FROM  SoftwarePackage SP  	INNER JOIN INSERTED INS On SP.SoftwarePackageID = INS.SoftwarePackageID 
					INNER JOIN Application  ON INS.ApplicationID = Application.ApplicationID
	WHERE INS.ApplicationID = Application.ApplicationID

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE TRIGGER insSoftwarePackageDocument
ON dbo.SoftwarePackageDocument 
AFTER INSERT
AS

	UPDATE Document set 	Document.SPVersion = SP.SPVersion,
			    	Document.SPVersionBase =	(CASE 	WHEN Document.SPVersionBase IS Null
								  	THEN SP.SPVersion
									ELSE Document.SPVersionBase
									END
								)
						
	FROM  SoftwarePackage SP  	INNER JOIN INSERTED INS On SP.SoftwarePackageID = INS.SoftwarePackageID 
					INNER JOIN Document  ON INS.DocumentID = Document.DocumentID
	WHERE INS.DocumentID = Document.DocumentID

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE TRIGGER insSoftwarePackageRegSetting
ON dbo.SoftwarePackageRegSetting 
AFTER INSERT
AS

	UPDATE RegSetting set RegSetting.SPVersion = SP.SPVersion,
				RegSetting.SPVersionBase =	(CASE 	WHEN RegSetting.SPVersionBase IS Null
								  	THEN SP.SPVersion
									ELSE RegSetting.SPVersionBase
									END
								)		
	FROM  SoftwarePackage SP  	INNER JOIN INSERTED INS On SP.SoftwarePackageID = INS.SoftwarePackageID 
					INNER JOIN RegSetting  ON INS.RegSettingID = RegSetting.RegSettingID
	WHERE INS.RegsettingID = RegSetting.RegSettingID

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE TRIGGER updStateHistory
ON dbo.State
FOR UPDATE
AS
INSERT INTO StateHistory
	SELECT del.* 
	FROM DELETED del

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE TRIGGER updStatusHistory
ON dbo.Status
FOR UPDATE
AS
INSERT INTO StatusHistory
	SELECT del.* 
	FROM DELETED del

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE TRIGGER updTypeOfLossHistory
ON dbo.TypeOfLoss
FOR UPDATE
AS
INSERT INTO TypeOfLossHistory
	SELECT del.* 
	FROM DELETED del

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE TRIGGER updUserProfileHistory
ON dbo.UserProfile
FOR UPDATE
AS
INSERT INTO UserProfileHistory
	SELECT del.* 
	FROM DELETED del

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE TRIGGER updUsersHistory
ON dbo.Users
FOR UPDATE
AS
INSERT INTO UsersHistory
	SELECT del.* 
	FROM DELETED del

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE TRIGGER insUsers
ON dbo.Users
AFTER INSERT
AS
--1 After Adding a User Also Add the UserID to AdjusterUsersSoftware Table 
insert into AdjusterUsersSoftware (
	[UsersID],
	[IBPrefix],
	[DateLastUpdated],
	[UpdateByUserID]
	)
SELECT
	INS.[UsersID],
	Left(INS.[FirstName],1) + Left(INS.[LastName],1) As IBPrefix, --Trye to use the First I and Last I as default IB Prefix
	INS.[DateLastUpdated],
	INS.[UpdateByUserID]
FROM 	INSERTED INS

--2 After Adding a User Also Add the UserID to AdjusterUsersUpdates Table 
insert into AdjusterUsersUpdates (
	[UsersID],
	[DateLastUpdated],
	[UpdateByUserID]
	)
SELECT
	INS.[UsersID],
	INS.[DateLastUpdated],
	INS.[UpdateByUserID]
FROM 	INSERTED INS

--3 After Adding a User Also Add the UserID to ECSADJUSers Table 
insert into ECSADJUsers (
	[UsersID]
	)
SELECT
	INS.[UsersID]
FROM 	INSERTED INS

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE TRIGGER insXML_Trans
ON dbo.XML_Trans 
AFTER INSERT
AS
	UPDATE XML_TRANS 
	SET XML_TRANS.MOD20 = INS.XML_TransID % 20
	FROM INSERTED INS Inner Join XML_Trans On XML_Trans.XML_TransID = INS.XML_TransID

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

