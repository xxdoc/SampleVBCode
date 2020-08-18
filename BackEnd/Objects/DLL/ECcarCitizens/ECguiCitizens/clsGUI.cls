VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Carrier GUI Polymorphism
'Carrier GUI InterFace Definiton
'Use Implements V2ECKeyBoard.clsCarGUI in ALL clsGUI For All Carrier GUI Objects

Implements V2ECKeyBoard.clsCarGUI

'Option Index Values for Yes / No / NA
Public Enum GuiOptValues
    YES = 0
    No
    NA
End Enum

'Used in Assignments List View
Public Enum GuiAssignments
    CLIENTNUM = 1
    Insured
    LossDate
    LossDateSort ' Hidden
    AssignedDate
    AssignedDateSort ' Hidden
    ReceivedDate
    ReceivedDateSort ' Hidden
    ContactDate
    ContactDateSort ' Hidden
    InspectedDate
    InspectedDateSort ' Hidden
    Status
    CloseDate
    CloseDateSort
    AssignmentType
    Islocked
    IsDeleted
    UpLoadMe
    SP
    CatName
    CatCode
    IBNUM
    MAStreet
    MAStreetSort ' Hidden
    MACity
    MAState
    MAZIP
    MAZIP4
    PAStreet
    PAStreetSort ' Hidden
    PACity
    PAState
    PAZIP
    PAZIP4
    Adjuster
    ACID
    DateReassigned
    DateReassignedSort ' Hidden
    DateLastUpdated
    DateLastUpdatedSort ' Hidden
    AdminComments
    RKey 'Hidden
End Enum

'Used in Assignments List
Public Enum GuiAssignmentsPic
    IsDeleted = 1
    Islocked
    UploadLossReport
    UpLoadMe
    AssignmentType_Property
    AssignmentType_Auto
    EmptyPic
    CalandarPic
End Enum

'-----------------------------------Used in frmClaim-----------------------
Public Enum GuiClaimOptions
    opt01_ClaimInfo = 0
    opt02_ActivityLog
    opt03_Indemnity
    opt04_BillingInformation
    opt05_Photos
    opt06_Reports
    opt07_Attachments
    opt08_Miscellaneous
    opt09_Print
End Enum

'----------------------------------Used in frmClaimInfo-------------------------
Public Enum GuiPolicyLimits
    ClassTypeClass = 1
    ClassTypeID 'Hidden
    ClassTypeDescription
    LimitAmount
    LimitAmountSort 'Hidden
    RCSaidProp
    RCSaidPropSort 'Hidden
    Reserves
    ReservesSort 'Hidden
    IsDeleted
    UpLoadMe
    DateLastUpdated
    DateLastUpdatedSort
    AdminComments
    ID 'Hidden
    IDAssignments 'Hidden
    SortOrder 'Hidden Only used in frmIndemnity when figuring out Deductible Apply order
End Enum

'Used in frmClaimInfo
Public Enum GuiPolicyLimitsPic
    IsDeleted = 1
    UpLoadMe
    EmptyPic
End Enum

'------------------------------------------Used in frmPhotos-----------------------
Public Enum GuiPhotoListView
    Thumb = 1
    SortOrder
    SortOrderSort ' Hidden
    Status
    IB
    PhotoDate
    PhotoDateSort ' Hidden
    Description
    PhotoName
    UpLoadPhoto
    UpLoadPhotoThumb
    UploadPhotoHighRes
    IsDeleted
    UpLoadMe
    DateLastUpdated
    DateLastUpdatedSort ' Hidden
    AdminComments
    ID ' Hidden
    IDRTPhotoReport 'Hidden
    IDAssignments ' Hidden
    RTPhotoLogID ' Hidden
    RTPhotoReportID ' hidden
    AssignmentsID ' Hidden
    BillingCountID ' Hidden
    IDBillingCount 'Hidden
    Photo 'Hidden
    DownloadPhoto 'hidden
    PhotoThumb 'hidden
    DownloadPhotoThumb 'hidden
    PhotoHighRes 'hidden
    DownloadPhotoHighRes 'hidden
    DownLoadMe ' hidden
    UpdateByUserID ' Hidden
End Enum
Public Enum GuiPhotoStatusList
    IsDeleted = 1
    UpLoadMe
    IsDeletedColHeader
    UpLoadMeColHeader
    EmptyPic
    PhotoUploaded
End Enum

Public Type GuiPhotoItem
    RTPhotoLogID As String
    RTPhotoReportID As String
    AssignmentsID As String
    BillingCountID As String
    ID As String
    IDRTPhotoReport As String
    IDAssignments As String
    IDBillingCount As String
    PhotoDate As String
    SortOrder As String
    Description As String
    PhotoName As String
    Photo As String
    DownloadPhoto As String
    UpLoadPhoto As String
    PhotoThumb As String
    DownloadPhotoThumb As String
    UpLoadPhotoThumb As String
    PhotoHighRes As String
    DownloadPhotoHighRes As String
    UploadPhotoHighRes As String
    IsDeleted As String
    DownLoadMe As String
    UpLoadMe As String
    AdminComments As String
    DateLastUpdated As String
    UpdateByUserID As String
End Type

'------------------------------------------Used in frmActivityLog-----------------------
Public Enum GuiActLogListView
    ActDate = 1
    ActDateSort ' Hidden
    ActTime
    ActTimeSort ' Hidden
    ActText
    ServiceTime
    ServiceTimeSort ' Hidden
    IB
    PageBreakAfter 'hidden until future Use
    BlankPageAfter 'hidden until future Use
    BlankRowsAfter 'hidden until future Use
    BlankRowsAfterSort ' Hidden
    IsMgrEntry 'hidden until future Use
    IsDeleted
    UpLoadMe
    DateLastUpdated
    DateLastUpdatedSort ' Hidden
    AdminComments
    ID ' Hidden
    IDAssignments ' Hidden
    RTActivityLogID ' Hidden
    AssignmentsID ' Hidden
    BillingCountID ' Hidden
    IDBillingCount 'Hidden
    DownLoadMe ' hidden
    UpdateByUserID ' Hidden
End Enum
Public Enum GuiActLogStatusList
    IsDeleted = 1
    UpLoadMe
End Enum

Public Type GuiActLogItem
    RTActivityLogID As String
    AssignmentsID As String
    BillingCountID As String
    ID As String
    IDAssignments As String
    IDBillingCount As String
    ServiceTime As String
    ActDate As String
    ActText As String
    ActTime As String
    PageBreakAfter As String
    BlankPageAfter As String
    BlankRowsAfter As String
    IsMgrEntry As String
    IsDeleted As String
    DownLoadMe As String
    UpLoadMe As String
    AdminComments As String
    DateLastUpdated As String
    UpdateByUserID As String
End Type

Public Type GuiActLogInfoItem
    AssignmentsID As String
    IDAssignments As String
    AL01_sPresentDurringInspection As String
    AL02_sExplainedEstimate As String
    AL03_sExplainedRCV As String
    AL03_sExplainedRCVNA As String
    AL04_sConfirmMortgageeIsCorrect As String
    AL04_sConfirmMortgageeIsCorrectNA As String
    AL05_sExplainedMortgageeChecks As String
    AL05_sExplainedMortgageeChecksNA As String
    AL06_sConfirmedCoverage As String
    AL07_sPriorLoss As String
    AL07_sPriorLossNA As String
    AL08_sSalvage As String
    AL09_sSubrogation As String
    IsDeleted As String
    DownLoadMe As String
    UpLoadMe As String
    AdminComments As String
    DateLastUpdated As String
    UpdateByUserID As String
End Type

'------------------------------------------Used in frmAttachments-----------------------
Public Enum GuiAttachListView
    AttachDate = 1
    AttachDateSort ' Hidden
    SortOrder
    SortOrderSort ' Hidden
    AttachName
    Description
    Attachment
    UploadAttachment
    IsDeleted
    UpLoadMe
    DateLastUpdated
    DateLastUpdatedSort ' Hidden
    AdminComments
    ID ' Hidden
    IDAssignments ' Hidden
    RTAttachmentsID ' Hidden
    AssignmentsID ' Hidden
    DownloadAttachment ' Hidden
    DownLoadMe ' hidden
    UpdateByUserID ' Hidden
End Enum

Public Enum GuiAttachStatusList
    IsDeleted = 1
    UpLoadMe
End Enum

Public Type GuiAttachItem
    RTAttachmentsID As String
    AssignmentsID As String
    ID As String
    IDAssignments As String
    AttachDate As String
    SortOrder As String
    Description As String
    AttachName As String
    Attachment As String
    DownloadAttachment As String
    UploadAttachment As String
    IsDeleted As String
    DownLoadMe As String
    UpLoadMe As String
    AdminComments As String
    DateLastUpdated As String
    UpdateByUserID As String
End Type

'------------------------------------------Used in frmReports-----------------------
Public Enum GuiRptParamsListView
    ParamCaption = 1
    ParamValue
    IsDeleted
    UpLoadMe
    DateLastUpdated
    DateLastUpdatedSort ' Hidden
    AdminComments
    Number ' hidden
    ProjectName ' hidden
    ClassName ' hidden
    ParamName ' hidden
    ParamDataType ' hidden
    SortMe ' hidden
    SortMeSort ' hidden
    ID ' Hidden
    IDAssignments ' Hidden
    MiscReportParamID ' Hidden
    AssignmentsID ' Hidden
    DownLoadMe ' hidden
    UpdateByUserID ' Hidden
End Enum

Public Enum GuiRptParamsStatusList
    IsDeleted = 1
    UpLoadMe
    ValueIsChecked
    ValueIsUnchecked
    ShowCalendarBox
    ShowUserDefinedToolBox
End Enum
'Use v2EcKeyBoard MiscReportParam UDT for GuiItem

'------------------------------------------Used in frmIndemnity-----------------------
'----------------------------Indemnity Frame
Public Enum GuiIndemListView
    PaymentRequest = 1
    ClassOfLoss
    ClassOfLossCode
    TypeOfLoss
    TypeOfLossCode
    Description
    ReplacementCost
    ReplacementCostSort 'Hidden
    RecoverableDep
    RecoverableDepSort 'Hidden
    NonRecoverableDep
    NonRecoverableDepSort 'Hidden
    ACVClaim
    ACVClaimSort 'Hidden
    SpecialLimits
    SpecialLimitsSort 'Hidden
    IsAddAmountOfInsurance
    ExcessAbsorbsDeductible
    AppliedDeductible
    AppliedDeductibleSort 'Hidden
    ExcessLimits
    ExcessLimitsSort 'hidden
    Miscellaneous
    MiscellaneousSort 'Hidden
    MiscellaneousDesc
    ACVLessExcessLimits
    ACVLessExcessLimitsSort 'Hidden
    IsPreviousPayment
    PPayDatePaid
    PPayDatePaidSort 'Hidden
    PPayAmountPaid
    PPayAmountPaidSort 'hidden
    PPayCheckNumber
    PPayCheckNumberSort 'Hidenn
    IsDeleted
    UpLoadMe
    DateLastUpdated
    DateLastUpdatedSort ' Hidden
    AdminComments
    ClassOfLossID ' Hidden
    TypeOfLossID 'Hidden
    ID ' Hidden
    IDAssignments ' Hidden
    RTIndemnityID ' Hidden
    AssignmentsID ' Hidden
    RTChecksID ' Hidden
    IDRTChecks 'Hidden
    DownLoadMe ' hidden
    UpdateByUserID ' Hidden
End Enum

Public Enum GuiIndemStatusList
    IsDeleted = 1
    UpLoadMe
    IsChecked
End Enum

Public Type GuiIndemItem
    RTIndemnityID As String
    AssignmentsID As String
    RTChecksID As String
    ID As String
    IDAssignments As String
    IDRTChecks As String
    ACVClaim As String
    ACVLessExcessLimits As String
    SpecialLimits As String
    ExcessLimits As String
    Miscellaneous As String
    MiscDescription As String
    IsAddAmountOfInsurance As String
    ExcessAbsorbsDeductible As String
    AppliedDeductible As String
    NonRecoverableDep As String
    RecoverableDep As String
    ReplacementCost As String
    TypeOfLossID As String
    ClassOfLossID As String
    Description As String
    IsPreviousPayment As String
    PPayDatePaid As String
    PPayAmountPaid As String
    PPayCheckNumber As String
    IsDeleted As String
    DownLoadMe As String
    UpLoadMe As String
    AdminComments As String
    DateLastUpdated As String
    UpdateByUserID As String
End Type

'---------------------Payment Request Frame
Public Enum GuiPayReqsListView
    CheckNum = 1
    CheckNumSort 'hidden
    PrintedDate
    PrintedDateSort 'Hiden
    NoOfRequests
    NoOfRequestsSort 'Hidden
    IB
    PrintOnIB
    ClassOfLoss
    ClassOfLossCode
    TypeOfLoss
    TypeOfLossCode
    RT50_sInsuredPayeeName
    RT51_sPayeeNames
    RT52_sAddress
    RT52_sAddressSort 'Hidden
    RT53_cAmountOfCheck
    RT53_cAmountOfCheckSort 'Hidden
    AppliedDeductible
    AppliedDeductibleSort 'Hidden
    CatCode
    IsDeleted
    UpLoadMe
    DateLastUpdated
    DateLastUpdatedSort ' Hidden
    AdminComments
    tempCHeckName 'Hidden
    RT42_ClassOfLossID 'Hidden
    RT43_TypeOfLossID 'Hidden
    RT54_CompanyCatSpecID 'Hidden
    ID ' Hidden
    IDAssignments ' Hidden
    RTChecksID ' Hidden
    AssignmentsID ' Hidden
    BillingCountID ' Hidden
    IDBillingCount 'Hidden
    DownLoadMe ' hidden
    UpdateByUserID ' Hidden
End Enum

Public Enum GuiPayReqsStatusList
    IsDeleted = 1
    UpLoadMe
    IsChecked
End Enum

Public Type GuiPayReqsItem
    RTChecksID As String
    AssignmentsID As String
    BillingCountID As String
    ID As String
    IDAssignments As String
    IDBillingCount As String
    CheckNum As String
    RT42_ClassOfLossID As String
    RT43_TypeOfLossID As String
    RT50_sInsuredPayeeName As String
    RT51_sPayeeNames As String
    RT52_sAddress As String
    RT53_cAmountOfCheck As String
    AppliedDeductible As String
    RT54_CompanyCatSpecID As String
    tempCHeckName As String
    PrintOnIB As String
    PrintedDate As String
    IsDeleted As String
    DownLoadMe As String
    UpLoadMe As String
    AdminComments As String
    DateLastUpdated As String
    UpdateByUserID As String
End Type

'----------Applied Deductible
Public Type AppliedDeductible
    ClassTypeID As String
    AppliedDeductible As Currency
End Type

'------------------------------------------Used in frmFeeSchedule-----------------------

'-----------------------FeeSchedLevels Frame
Public Enum GuiFeeSchedLevelsListView
    LevelNum = 1
    LevelNumSort
    LevelMax
    LevelMaxSort
    LevelPctApp
    LevelPctAppSort
    LevelMin
    LevelMinSort
    DateLastUpdated
    DateLastUpdatedSort ' Hidden
    UpdateByUserID ' Hidden
    FeeScheduleLevelsID ' Hidden
    FeeScheduleID ' Hidden
End Enum

Public Enum GuiFeeTypesStatusList
    IsDeleted = 1
    UpLoadMe
    IsChecked
End Enum

'-------------FeeTypes
Public Enum GuiFeeTypesListView
    TypeNum = 1
    TypeNumSort
    ftName
    Description
    FeeAmount
    FeeAmountSort
    IsExpense
    MaxNumberOfItems
    MaxNumberOfItemsSort
    MaxFeeAmount
    MaxFeeAmountSort
    UseFormula
    VBFormula
    DateLastUpdated
    DateLastUpdatedSort ' Hidden
    UpdateByUserID ' Hidden
    FeeScheduleFeeTypesID ' Hidden
    FeeScheduleID ' Hidden
End Enum

Public Type FSOptions
    optCAPTION As String
    optTYPE As String
    optDATATYPE As String
    optVBREGSTR As String
    optDEFAULTVALUE As String
    ctlINDEX As Long
End Type

'------------------------------------------Used in frmBillingInfo-----------------------

'Billing Count (Keeps track of Billing ie supplements = number of bills Rebills = times each bill revised

Public Enum GuiIBStatusList
    IsDeleted = 1
    UpLoadMe
    IsChecked
    IsUnchecked
    showPDF
    showDropDown
End Enum

Public Enum GuiOverrideFeesListView
    fsftDescription = 1
    Comment
    NumberOfItems
    Amount
    UpLoadMe
    DateLastUpdated
    AdminComments
    DownLoadMe 'hidden
    RTIBFeeID 'hidden
    AssignmentsID 'hidden
    ID 'hidden
    IDAssignments 'hidden
    UpdateByUserID 'hidden
    FeeScheduleFeeTypesID 'hidden
    fsftFeeScheduleID 'hidden
    fsftTypeNum 'hidden
    fsftName 'hidden
    fsftFeeAmount 'hidden
    fsftIsExpense 'hidden
    fsftMaxNumberOfItems 'hidden
    fsftMaxFeeAmount 'hidden
    fsftIsMiscAmount 'hidden
    fsftUseFormula 'hidden
    fsftVBFormula 'hidden
    fsftIsDeleted 'hidden
    fsftDateLastUpdated 'hidden
    fsftUpdateByUserID 'hidden
End Enum

Public Enum GuiServiceFeesListView
    fsftDescription = 1
    Comment
    NumberOfItems
    Amount
    UpLoadMe
    DateLastUpdated
    AdminComments
    DownLoadMe 'hidden
    RTIBFeeID 'hidden
    AssignmentsID 'hidden
    ID 'hidden
    IDAssignments 'hidden
    UpdateByUserID 'hidden
    FeeScheduleFeeTypesID 'hidden
    fsftFeeScheduleID 'hidden
    fsftTypeNum 'hidden
    fsftName 'hidden
    fsftFeeAmount 'hidden
    fsftIsExpense 'hidden
    fsftMaxNumberOfItems 'hidden
    fsftMaxFeeAmount 'hidden
    fsftIsMiscAmount 'hidden
    fsftUseFormula 'hidden
    fsftVBFormula 'hidden
    fsftIsDeleted 'hidden
    fsftDateLastUpdated 'hidden
    fsftUpdateByUserID 'hidden
End Enum

Public Enum GuiExpenseFeesListView
    fsftDescription = 1
    Comment
    NumberOfItems
    Amount
    UpLoadMe
    DateLastUpdated
    AdminComments
    DownLoadMe 'hidden
    RTIBFeeID 'hidden
    AssignmentsID 'hidden
    ID 'hidden
    IDAssignments 'hidden
    UpdateByUserID 'hidden
    FeeScheduleFeeTypesID 'hidden
    fsftFeeScheduleID 'hidden
    fsftTypeNum 'hidden
    fsftName 'hidden
    fsftFeeAmount 'hidden
    fsftIsExpense 'hidden
    fsftMaxNumberOfItems 'hidden
    fsftMaxFeeAmount 'hidden
    fsftIsMiscAmount 'hidden
    fsftUseFormula 'hidden
    fsftVBFormula 'hidden
    fsftIsDeleted 'hidden
    fsftDateLastUpdated 'hidden
    fsftUpdateByUserID 'hidden
End Enum


'------- Table UDT for Billing info
Public Type GuiBillingCountItem
    BillingCountID As String
    AssignmentsID As String
    ID As String
    IDAssignments As String
    Rebill As String
    Supplement As String
    DownLoadMe As String
    UpLoadMe As String
    AdminComments As String
    DateLastUpdated As String
    UpdateByUserID As String
End Type

'------Real Time IB  (Currently working on IB)
Public Type GuiRTIBItem
    AssignmentsID As String
    BillingCountID As String
    IDAssignments As String
    IDBillingCount As String
    RT00_lSSN As String
    RT01_sSubToCarrier As String
    RT02_sIBNumber As String
    RT05_sLocation As String
    RT05a_sState As String
    RT06_dtDateClosed As String
    RT07_sAdjusterName As String
    RT09_sSALN As String
    RT09a_sPolicyNo As String
    RT10_sInsuredName As String
    RT11_sLossLocation As String
    RT12_dtDateOfLoss As String
    RT13_cGrossLoss As String
    RT14_cDepreciation As String
    RT14a_sSupplement As String
    RT14b_sRebilled As String
    RT15_cDeductible As String
    RT15a_cLessExcessLimits As String
    RT15b_sExcessLimDesc As String
    RT15c_cLessMiscellaneous As String
    RT15d_cMiscellaneousDesc As String
    RT16_cNetClaim As String
    RT17_cServiceFee As String
    RT17a_cMiscServiceFee As String
    RT18_sServiceFeeComment As String
    RT18a_sMiscServiceFeeComment As String
    RT25_cServiceFeeSubTotal As String
    RT29a_sMiscExpenseFeeComment As String
    RT29b_cMiscExpenseFee As String
    RT30_cTotalExpenses As String
    RT31_dTaxPercent As String
    RT32_cTaxAmount As String
    RT33_cTotalAdjustingFee As String
    RT33a_sAccountCode As String
    FeeScheduleID As String
    Void As String
    FeeByTime As String
    UseActivityTime As String
    DownLoadMe As String
    UpLoadMe As String
    Comments As String
    AdminComments As String
    DateLastUpdated As String
    UpdateByUserID As String
End Type

Public Type GuiRTIBFeeItem
    RTIBFeeID As String
    AssignmentsID As String
    ID As String
    IDAssignments As String
    FeeScheduleFeeTypesID As String
    NumberOfItems As String
    Amount As String
    Comment As String
    DownLoadMe As String
    UpLoadMe As String
    AdminComments As String
    DateLastUpdated As String
    UpdateByUserID As String
End Type

'-------IB (Closed IB --- can be Reopened
Public Type GuiIBItem
    IBID As String
    AssignmentsID As String
    BillingCountID As String
    ID As String
    IDAssignments As String
    IDBillingCount As String
    IB00_lssn As String
    IB01_sSubToCarrier As String
    IB02_sIBNumber As String
    IB05_sLocation As String
    IB05a_sState As String
    IB06_dtDateClosed As String
    IB07_sAdjusterName As String
    IB09_sSALN As String
    IB09a_sPolicyNo As String
    IB10_sInsuredName As String
    IB11_sLossLocation As String
    IB12_dtDateOfLoss As String
    IB13_cGrossLoss As String
    IB14_cDepreciation As String
    IB14a_sSupplement As String
    IB14b_sRebilled As String
    IB15_cDeductible As String
    IB15a_cLessExcessLimits As String
    IB15b_sExcessLimDesc As String
    IB15c_cLessMiscellaneous As String
    IB15d_cMiscellaneousDesc As String
    IB16_cNetClaim As String
    IB17_cServiceFee As String
    IB17a_cMiscServiceFee As String
    IB18_sServiceFeeComment As String
    IB18a_sMiscServiceFeeComment As String
    IB25_cServiceFeeSubTotal As String
    IB29a_sMiscExpenseFeeComment As String
    IB29b_cMiscExpenseFee As String
    IB30_cTotalExpenses As String
    IB31_dTaxPercent As String
    IB32_cTaxAmount As String
    IB33_cTotalAdjustingFee As String
    IB33a_sAccountCode As String
    FeeScheduleID As String
    Void As String
    FeeByTime As String
    UseActivityTime As String
    DownLoadMe As String
    UpLoadMe As String
    Comments As String
    AdminComments As String
    DateLastUpdated As String
    UpdateByUserID As String
End Type

Public Type GuiIBFeeItem
    IBFeeID As String
    AssignmentsID As String
    IBID As String
    ID As String
    IDAssignments As String
    IDIB As String
    FeeScheduleFeeTypesID As String
    NumberOfItems As String
    Amount As String
    Comment As String
    DownLoadMe As String
    UpLoadMe As String
    AdminComments As String
    DateLastUpdated As String
    UpdateByUserID As String
End Type

'------------------------------------------Used in frmPrint-----------------------

Public Enum GuiPackageItemStatusList
    IsChecked = 1
    ItemHasBeenSent
    UpLoadMe
    IsDeleted
    IsApproved
    IsRejected
End Enum

Public Enum SavedPackagesListView
    Name = 1
    DateLastUpdated
    DateLastUpdatedSort
    DateCreated
    DateCreatedSort
End Enum

Public Enum GuiPackageItemListView
    SortOrder = 1
    SortOrderSort
    piName
    Description
    AttachmentName
    IsCoApprove
    CoApproveDate
    CoApproveDesc
    IsClientCoReject
    ClientCoRejectDate
    ClientCoRejectDesc
    IsClientCoApprove
    ClientCoApproveDate
    ClientCoApproveDesc
    IsClientCoDelete
    ClientCoDeleteDate
    ClientCoDeleteDesc
    SendMe
    SentDate
    IsDeleted
    UpLoadMe
    DateLastUpdated
    AdminComments
    'Hidden
    PackageItemID
    PackageID
    AssignmentsID
    ID
    IDPackage
    IDAssignments
    ReportFormat
    RTAttachmentsID
    IDRTAttachments
    Number
    PackageItemGUID
    DownLoadMe
    UpdateByUserID
End Enum

Public Type GuiPackageItem
    PackageItemID As String
    PackageID As String
    AssignmentsID As String
    ID As String
    IDPackage As String
    IDAssignments As String
    ReportFormat As String
    RTAttachmentsID As String
    IDRTAttachments As String
    Number As String
    AttachmentName As String
    SortOrder As String
    Name As String
    Description As String
    IsCoApprove As String
    CoApproveDate As String
    CoApproveDesc As String
    IsClientCoReject As String
    ClientCoRejectDate As String
    ClientCoRejectDesc As String
    IsClientCoDelete As String
    ClientCoDeleteDate As String
    ClientCoDeleteDesc As String
    IsClientCoApprove As String
    ClientCoApproveDate As String
    ClientCoApproveDesc As String
    PackageItemGUID As String
    SendMe As String
    SentDate As String
    IsDeleted As String
    DownLoadMe As String
    UpLoadMe As String
    AdminComments As String
    DateLastUpdated As String
    UpdateByUserID As String
End Type

Public Type GuiPackage
    PackageID As String
    AssignmentsID As String
    ID As String
    IDAssignments As String
    CreateDate As String
    PackageStatus As String
    Name As String
    Description As String
    Number As String
    SendMe As String
    SentDate As String
    SentToEmail As String
    IsDeleted As String
    DownLoadMe As String
    UpLoadMe As String
    AdminComments As String
    DateLastUpdated As String
    UpdateByUserID As String
End Type


'------------------------------------

Private mfrmClaimsList As frmClaimsList
Private mfrmAdjusterReports As frmAdjusterReports
Private mfrmFeeSchedule As frmFeeSchedule
Private madoRSTypeOfLoss As ADODB.Recordset
Private madoClassOfLoss As ADODB.Recordset
Private madoRSAssignmentType As ADODB.Recordset
Private madoRSACID As ADODB.Recordset
Private madoRSCatCode As ADODB.Recordset
Private madoRSState As ADODB.Recordset
Private madoRSClassType As ADODB.Recordset
Private madoRSFeeSchedule As ADODB.Recordset
Private madoRSFeeScheduleLevels As ADODB.Recordset
Private madoRSFeeScheduleFeeTypes As ADODB.Recordset

Private Property Set clsCarGUI_adoFeeScheduleFeeTypes(padoRS As ADODB.Recordset)
    Set madoRSFeeScheduleFeeTypes = padoRS
End Property
Private Property Let clsCarGUI_adoFeeScheduleFeeTypes(padoRS As ADODB.Recordset)
    Set madoRSFeeScheduleFeeTypes = padoRS
End Property
Private Property Get clsCarGUI_adoFeeScheduleFeeTypes() As ADODB.Recordset
    Set clsCarGUI_adoFeeScheduleFeeTypes = madoRSFeeScheduleFeeTypes
End Property

Private Property Set clsCarGUI_adoFeeScheduleLevels(padoRS As ADODB.Recordset)
    Set madoRSFeeScheduleLevels = padoRS
End Property
Private Property Let clsCarGUI_adoFeeScheduleLevels(padoRS As ADODB.Recordset)
    Set madoRSFeeScheduleLevels = padoRS
End Property
Private Property Get clsCarGUI_adoFeeScheduleLevels() As ADODB.Recordset
    Set clsCarGUI_adoFeeScheduleLevels = madoRSFeeScheduleLevels
End Property

Private Property Set clsCarGUI_adoFeeSchedule(padoRS As ADODB.Recordset)
    Set madoRSFeeSchedule = padoRS
End Property
Private Property Let clsCarGUI_adoFeeSchedule(padoRS As ADODB.Recordset)
    Set madoRSFeeSchedule = padoRS
End Property
Private Property Get clsCarGUI_adoFeeSchedule() As ADODB.Recordset
    Set clsCarGUI_adoFeeSchedule = madoRSFeeSchedule
End Property

Private Property Let clsCarGUI_adoClassType(padoRS As ADODB.Recordset)
    Set madoRSClassType = padoRS
End Property
Private Property Set clsCarGUI_adoClassType(padoRS As ADODB.Recordset)
    Set madoRSClassType = padoRS
End Property
Private Property Get clsCarGUI_adoClassType() As ADODB.Recordset
    Set clsCarGUI_adoClassType = madoRSClassType
End Property

Private Property Let clsCarGUI_adoRSState(padoRS As ADODB.Recordset)
    Set madoRSState = padoRS
End Property
Private Property Set clsCarGUI_adoRSState(padoRS As ADODB.Recordset)
    Set madoRSState = padoRS
End Property
Private Property Get clsCarGUI_adoRSState() As ADODB.Recordset
    Set clsCarGUI_adoRSState = madoRSState
End Property

Private Property Let clsCarGUI_adoRSCatCode(padoRS As ADODB.Recordset)
    Set madoRSCatCode = padoRS
End Property
Private Property Set clsCarGUI_adoRSCatCode(padoRS As ADODB.Recordset)
    Set madoRSCatCode = padoRS
End Property
Private Property Get clsCarGUI_adoRSCatCode() As ADODB.Recordset
    Set clsCarGUI_adoRSCatCode = madoRSCatCode
End Property

Private Property Let clsCarGUI_adoRSACID(padoRS As ADODB.Recordset)
    Set madoRSACID = padoRS
End Property
Private Property Set clsCarGUI_adoRSACID(padoRS As ADODB.Recordset)
    Set madoRSACID = padoRS
End Property
Private Property Get clsCarGUI_adoRSACID() As ADODB.Recordset
    Set clsCarGUI_adoRSACID = madoRSACID
End Property

Private Property Let clsCarGUI_adoRSAssignmentType(padoRS As ADODB.Recordset)
    Set madoRSAssignmentType = padoRS
End Property
Private Property Set clsCarGUI_adoRSAssignmentType(padoRS As ADODB.Recordset)
    Set madoRSAssignmentType = padoRS
End Property
Private Property Get clsCarGUI_adoRSAssignmentType() As ADODB.Recordset
    Set clsCarGUI_adoRSAssignmentType = madoRSAssignmentType
End Property

Private Property Let clsCarGUI_adoRSTypeOfLoss(padoRS As ADODB.Recordset)
    Set madoRSTypeOfLoss = padoRS
End Property
Private Property Set clsCarGUI_adoRSTypeOfLoss(padoRS As ADODB.Recordset)
    Set madoRSTypeOfLoss = padoRS
End Property
Private Property Get clsCarGUI_adoRSTypeOfLoss() As ADODB.Recordset
    Set clsCarGUI_adoRSTypeOfLoss = madoRSTypeOfLoss
End Property

Private Property Get clsCarGUI_MyClaimsList() As Object
    Set clsCarGUI_MyClaimsList = mfrmClaimsList
End Property

Private Property Get clsCarGUI_ClassName() As String
    clsCarGUI_ClassName = App.EXEName & ".clsGUI"
End Property

Private Property Get clsCarGUI_goForms() As Object
    Set clsCarGUI_goForms = Forms
End Property

Private Function clsCarGUI_SetUtilObject(oUtilObject As Object) As Boolean

    On Error Resume Next
    Set goUtil = oUtilObject
    If Err.Number = 0 Then
        clsCarGUI_SetUtilObject = True
    Else
        clsCarGUI_SetUtilObject = False
    End If

End Function

Private Function clsCarGUI_ShowCalendar(poTextBox As Object) As Boolean
    On Error GoTo EH
    Dim oCalendar As V2ECKeyBoard.clsCalendar
    Dim oTextBox As TextBox
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    Set oCalendar = New V2ECKeyBoard.clsCalendar
       
    If TypeOf poTextBox Is TextBox Then
        Set oTextBox = poTextBox
        If Not IsDate(oTextBox.Text) Then
            oTextBox.Text = vbNullString
            oCalendar.CurrentDate = Now()
        Else
            oCalendar.CurrentDate = oTextBox.Text
        End If
        oCalendar.ShowCalendar
        If oCalendar.CurrentDate = NULL_DATE Then
            If IsDate(oTextBox.Text) Then
                If CDate(oTextBox.Text) = NULL_DATE Then
                    oTextBox.Text = vbNullString
                End If
            Else
                oTextBox.Text = vbNullString
            End If
        Else
            oTextBox.Text = Format(oCalendar.CurrentDate, "MM/DD/YYYY")
        End If
    End If
    
    clsCarGUI_ShowCalendar = True
    
    'cleanup
    Set oCalendar = Nothing

    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    clsCarGUI_ShowCalendar = False
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarGUI_ClassName & vbCrLf & "Private Function clsCarGUI_ShowCalendar"
End Function

Private Sub clsCarGUI_ShowForm(psFormName As String, Optional pMode As VBRUN.FormShowConstants = vbModal, Optional oMnuItem As Object)
    On Error GoTo EH
    Dim oForm As Form
    Dim oClaimForm As Form
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    If goUtil.utFindSetForm(Forms, psFormName, oForm) Then
        Select Case UCase(psFormName)
            Case UCase("frmClaimsList")
                'if selecting the Claims List...
                'First check to see if there is an open claim
                If goUtil.utFindSetForm(Forms, "frmClaim", oClaimForm) Then
                    oClaimForm.Show vbModeless
                    If oClaimForm.WindowState <> vbMaximized Then
                       oClaimForm.WindowState = vbNormal
                    End If
                    'be sure the Claims Form is not visible
                    oForm.Visible = False
                Else
                    oForm.Show vbModeless
                    If oForm.WindowState <> vbMaximized Then
                       oForm.WindowState = vbNormal
                    End If
                End If
            Case Else
                oForm.Show vbModeless
                If oForm.WindowState <> vbMaximized Then
                   oForm.WindowState = vbNormal
                End If
        End Select
    Else
        Select Case UCase(psFormName)
            Case UCase("frmClaimsList")
                Set mfrmClaimsList = New frmClaimsList
                Set mfrmClaimsList.MyGUI = Me
                Load mfrmClaimsList
                mfrmClaimsList.Show pMode
                goUtil.utFormWinRegPos goUtil.gsMainAppEXEName, mfrmClaimsList, , , , , True
            Case UCase("frmAdjusterReports")
                Set mfrmAdjusterReports = New frmAdjusterReports
                Set mfrmAdjusterReports.MyGUI = Me
                Load mfrmAdjusterReports
                mfrmAdjusterReports.Show pMode
                goUtil.utFormWinRegPos goUtil.gsMainAppEXEName, mfrmAdjusterReports, , , , , True
            Case UCase("frmFeeSchedule")
                Set mfrmFeeSchedule = New frmFeeSchedule
                Set mfrmFeeSchedule.MyGUI = Me
                Load mfrmFeeSchedule
                mfrmFeeSchedule.Show pMode
                goUtil.utFormWinRegPos goUtil.gsMainAppEXEName, mfrmFeeSchedule, , , , , True
        End Select
    End If
    
    Set oForm = Nothing
    Set oClaimForm = Nothing
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & clsCarGUI_ClassName & vbCrLf & "Private Sub clsCarGUI_ShowForm" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Sub

Private Sub clsCarGUI_ShowAllForms(oECTray As Object, psNavPos As String, Optional psSkipForms As String)
    goUtil.utShowAllForms Forms, oECTray, psNavPos, psSkipForms
End Sub

Private Sub clsCarGUI_HideAllForms(Optional psSkipForms As String)
    goUtil.utHideAllForms Forms, psSkipForms
End Sub

Private Function clsCarGUI_SetGlobalObjects(pcolObjects As Collection) As Boolean
    On Error GoTo EH
    'Init options and Optins Vars
    Dim MyfrmFeeSchedule As frmFeeSchedule
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    Set goUtil = pcolObjects("goUtil")
    'Set Lookup Recordset Values (Ado Disconnected RS in clsGUI)

    clsCarGUI_SetadoRSTypeOfLoss
    clsCarGUI_SetadoRSAssignmentType
    clsCarGUI_SetadoRSACID
    clsCarGUI_SetadoRSCatCode
    clsCarGUI_SetadoRSState
    clsCarGUI_SetadoClassType
    clsCarGUI_SetadoRSFeeSchedule
    clsCarGUI_SetadoRSFeeScheduleLevels
    clsCarGUI_SetadoRSFeeScheduleFeeTypes
    
    'Fee Schedule
    'need to save Default Values for InitOptions and Otions to Registry
    'Loading and unloading the  FeeSchedule will do this.
    Sleep 100
    Set MyfrmFeeSchedule = New frmFeeSchedule
    Set MyfrmFeeSchedule.MyGUI = Me
    Load MyfrmFeeSchedule
    Unload MyfrmFeeSchedule
    Set MyfrmFeeSchedule = Nothing
    
    clsCarGUI_SetGlobalObjects = True
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    clsCarGUI_SetGlobalObjects = False
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarGUI_ClassName & vbCrLf & "Private Function clsCarGUI_SetGlobalObjects"
End Function


Private Function clsCarGUI_SetadoRSTypeOfLoss() As Boolean
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    'reset rs
    If Not madoRSTypeOfLoss Is Nothing Then
        Set madoRSTypeOfLoss = Nothing
    End If
    
    Set oConn = New ADODB.Connection
    Set madoRSTypeOfLoss = New ADODB.Recordset
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    sSQL = "SELECT RetTypeOfLoss.* "
    sSQL = sSQL & "FROM( "
    sSQL = sSQL & "SELECT "
    sSQL = sSQL & "[TypeOfLossID], "
    sSQL = sSQL & "(SELECT    NAME "
    sSQL = sSQL & "FROM       Company "
    sSQL = sSQL & "WHERE CompanyID = T.[ClientCompanyID]) As ClientCompanyName, "
    sSQL = sSQL & "[ClientCompanyID], "
    sSQL = sSQL & "[TypeOfLoss], "
    sSQL = sSQL & "[Description], "
    sSQL = sSQL & "[Code], "
    sSQL = sSQL & "[IsDeleted] , "
    sSQL = sSQL & "[DateLastUpdated] , "
    sSQL = sSQL & "(SELECT    USERNAME "
    sSQL = sSQL & "FROM       USERS "
    sSQL = sSQL & "WHERE  USERSID = T.[UpdateByUserID]) As [UpdateByUserName],  "
    sSQL = sSQL & "[UpdateByUserID]  "
    sSQL = sSQL & "FROM TypeOfLoss T "
    sSQL = sSQL & ") RetTypeOfLoss "
    sSQL = sSQL & "WHERE ClientCompanyID = " & goUtil.gsCurCar & " "
    sSQL = sSQL & "AND IsDeleted = False "
    sSQL = sSQL & "ORDER BY TypeOfLoss "
    
    
    madoRSTypeOfLoss.CursorLocation = adUseClient
    madoRSTypeOfLoss.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set madoRSTypeOfLoss.ActiveConnection = Nothing
    
    clsCarGUI_SetadoRSTypeOfLoss = True
    Set oConn = Nothing
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Set oConn = Nothing
    clsCarGUI_SetadoRSTypeOfLoss = False
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarGUI_ClassName & vbCrLf & "Public Function clsCarGUI_SetadoRSTypeOfLoss"
End Function

Private Function clsCarGUI_SetadoRSAssignmentType() As Boolean
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    'reset rs
    If Not madoRSAssignmentType Is Nothing Then
        Set madoRSAssignmentType = Nothing
    End If
    
    Set oConn = New ADODB.Connection
    Set madoRSAssignmentType = New ADODB.Recordset
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    sSQL = "SELECT RetAssignmentType.* "
    sSQL = sSQL & "FROM( "
    sSQL = sSQL & "SELECT "
    sSQL = sSQL & "[AssignmentTypeID], "
    sSQL = sSQL & "[Type], "
    sSQL = sSQL & "[Description], "
    sSQL = sSQL & "[AdminComments], "
    sSQL = sSQL & "[IsDeleted] , "
    sSQL = sSQL & "[DateLastUpdated] , "
    sSQL = sSQL & "(SELECT    USERNAME "
    sSQL = sSQL & "FROM       USERS "
    sSQL = sSQL & "WHERE  USERSID = S.[UpdateByUserID]) As [UpdateByUserName],  "
    sSQL = sSQL & "[UpdateByUserID]  "
    sSQL = sSQL & "FROM AssignmentType S "
    sSQL = sSQL & ") RetAssignmentType "
    sSQL = sSQL & "WHERE IsDeleted = False "
    sSQL = sSQL & "ORDER BY Type "
    
    
    madoRSAssignmentType.CursorLocation = adUseClient
    madoRSAssignmentType.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set madoRSAssignmentType.ActiveConnection = Nothing
    
    clsCarGUI_SetadoRSAssignmentType = True
    Set oConn = Nothing
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Set oConn = Nothing
    clsCarGUI_SetadoRSAssignmentType = False
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarGUI_ClassName & vbCrLf & "Private Function clsCarGUI_SetadoRSAssignmentType"
End Function

Private Function clsCarGUI_SetadoRSACID() As Boolean
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    'reset rs
    If Not madoRSACID Is Nothing Then
        Set madoRSACID = Nothing
    End If
    
    Set oConn = New ADODB.Connection
    Set madoRSACID = New ADODB.Recordset
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    sSQL = "SELECT RetClientCoAdjusterSpec.* "
    sSQL = sSQL & "FROM( "
    sSQL = sSQL & "SELECT "
    sSQL = sSQL & "ClientCoAdjusterSpecID, "
    sSQL = sSQL & "(SELECT    NAME "
    sSQL = sSQL & "FROM   Company "
    sSQL = sSQL & "WHERE CompanyID = C.[ClientCompanyID]) As ClientCompanyName, "
    sSQL = sSQL & "(SELECT    Comments "
    sSQL = sSQL & "FROM   Company "
    sSQL = sSQL & "WHERE CompanyID = C.[ClientCompanyID]) As ClientCompanyDesc, "
    sSQL = sSQL & "(SELECT    IsClientOF "
    sSQL = sSQL & "FROM Company "
    sSQL = sSQL & "WHERE CompanyID = C.[ClientCompanyID]) As IsClientOF, "
    sSQL = sSQL & "[ClientCompanyID], "
    sSQL = sSQL & "(SELECT    USERNAME "
    sSQL = sSQL & "FROM   USERS "
    sSQL = sSQL & "WHERE  USERSID = C.[UsersID]) As [UserName],  "
    sSQL = sSQL & "(SELECT    LastName + ', ' + FirstName "
    sSQL = sSQL & "FROM   USERS "
    sSQL = sSQL & "WHERE  USERSID = C.[UsersID]) As [LFName],  "
    sSQL = sSQL & "[UsersID], "
    sSQL = sSQL & "[ACID], "
    sSQL = sSQL & "[ACIDDescription], "
    sSQL = sSQL & "[Comments], "
    sSQL = sSQL & "[ClientCompanyCatSpecID], "
    sSQL = sSQL & "[ZipCode], "
    sSQL = sSQL & "[ZipCodeCount], "
    sSQL = sSQL & "[ActiveDate] , "
    sSQL = sSQL & "[InactiveDate] , "
    sSQL = sSQL & "[DateLastUpdated] , "
    sSQL = sSQL & "(SELECT    USERNAME "
    sSQL = sSQL & "FROM   USERS "
    sSQL = sSQL & "WHERE  USERSID = C.[UpdateByUserID]) As [UpdateByUserName],  "
    sSQL = sSQL & "[UpdateByUserID] "
    sSQL = sSQL & "FROM ClientCoAdjusterSpec C WHERE [ZipCode] Is Null "
    sSQL = sSQL & ") RetClientCoAdjusterSpec "
    sSQL = sSQL & "WHERE ClientCompanyID = " & goUtil.gsCurCar & " "
    sSQL = sSQL & "AND UsersID = " & goUtil.gsCurUsersID & " "
    sSQL = sSQL & "ORDER BY ACID "
    
    madoRSACID.CursorLocation = adUseClient
    madoRSACID.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set madoRSACID.ActiveConnection = Nothing
    
    clsCarGUI_SetadoRSACID = True
    
    Set oConn = Nothing
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Set oConn = Nothing
    clsCarGUI_SetadoRSACID = False
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarGUI_ClassName & vbCrLf & "Private Function clsCarGUI_SetadoRSACID"
End Function

Private Function clsCarGUI_SetadoRSCatCode() As Boolean
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    'reset rs
    If Not madoRSCatCode Is Nothing Then
        Set madoRSCatCode = Nothing
    End If
    
    Set oConn = New ADODB.Connection
    Set madoRSCatCode = New ADODB.Recordset
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    sSQL = "SELECT RetClientCompanyCatSpec.* "
    sSQL = sSQL & "FROM( "
    sSQL = sSQL & "SELECT "
    sSQL = sSQL & "ClientCompanyCatSpecID, "
    sSQL = sSQL & "(SELECT    NAME "
    sSQL = sSQL & "FROM   Company "
    sSQL = sSQL & "WHERE CompanyID = C.[ClientCompanyID]) As ClientCompanyName, "
    sSQL = sSQL & "(SELECT    Comments "
    sSQL = sSQL & "FROM   Company "
    sSQL = sSQL & "WHERE CompanyID = C.[ClientCompanyID]) As ClientCompanyDesc, "
    sSQL = sSQL & "(SELECT    IsClientOF "
    sSQL = sSQL & "FROM Company "
    sSQL = sSQL & "WHERE CompanyID = C.[ClientCompanyID]) As IsClientOF, "
    sSQL = sSQL & "[ClientCompanyID], "
    sSQL = sSQL & "(SELECT    NAME "
    sSQL = sSQL & "FROM   Cat "
    sSQL = sSQL & "WHERE CatID = C.[CATID]) As CatName, "
    sSQL = sSQL & "[CATID], "
    sSQL = sSQL & "[CatCode], "
    sSQL = sSQL & "[Comments], "
    sSQL = sSQL & "[ActiveDate] , "
    sSQL = sSQL & "[InactiveDate] , "
    sSQL = sSQL & "[DateLastUpdated] , "
    sSQL = sSQL & "(SELECT    USERNAME "
    sSQL = sSQL & "FROM   USERS "
    sSQL = sSQL & "WHERE  USERSID = C.[UpdateByUserID]) As [UpdateByUserName],  "
    sSQL = sSQL & "[UpdateByUserID] "
    sSQL = sSQL & "FROM ClientCompanyCatSpec C "
    sSQL = sSQL & ") RetClientCompanyCatSpec "
    sSQL = sSQL & "WHERE [InactiveDate] Is Null "
    sSQL = sSQL & "AND ClientCompanyID = " & goUtil.gsCurCar & " "
    sSQL = sSQL & "AND CATID = " & goUtil.gsCurCat & " "
    sSQL = sSQL & "ORDER BY CatCode "
    
    madoRSCatCode.CursorLocation = adUseClient
    madoRSCatCode.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set madoRSCatCode.ActiveConnection = Nothing
    
    clsCarGUI_SetadoRSCatCode = True
    
    Set oConn = Nothing
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Set oConn = Nothing
    clsCarGUI_SetadoRSCatCode = False
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarGUI_ClassName & vbCrLf & "Private Function clsCarGUI_SetadoRSCatCode"
End Function

Private Function clsCarGUI_SetadoRSState() As Boolean
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    'reset rs
    If Not madoRSState Is Nothing Then
        Set madoRSState = Nothing
    End If
    
    Set oConn = New ADODB.Connection
    Set madoRSState = New ADODB.Recordset
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    sSQL = "SELECT RetState.* "
    sSQL = sSQL & "FROM( "
    sSQL = sSQL & "SELECT "
    sSQL = sSQL & "[StateID], "
    sSQL = sSQL & "[Name], "
    sSQL = sSQL & "[Code], "
    sSQL = sSQL & "[Comments], "
    sSQL = sSQL & "[IsDeleted] , "
    sSQL = sSQL & "[DateLastUpdated] , "
    sSQL = sSQL & "(SELECT    USERNAME "
    sSQL = sSQL & "FROM       USERS "
    sSQL = sSQL & "WHERE  USERSID = S.[UpdateByUserID]) As [UpdateByUserName],  "
    sSQL = sSQL & "[UpdateByUserID]  "
    sSQL = sSQL & "FROM State S "
    sSQL = sSQL & ") RetState "
    sSQL = sSQL & "WHERE IsDeleted = 0 "
    sSQL = sSQL & "ORDER BY Name "
    
    madoRSState.CursorLocation = adUseClient
    madoRSState.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set madoRSState.ActiveConnection = Nothing
    
    clsCarGUI_SetadoRSState = True
    
    Set oConn = Nothing
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Set oConn = Nothing
    clsCarGUI_SetadoRSState = False
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarGUI_ClassName & vbCrLf & "Private Function clsCarGUI_SetadoRSState"
End Function

Private Function clsCarGUI_SetadoClassType() As Boolean
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    'reset rs
    If Not madoRSClassType Is Nothing Then
        Set madoRSClassType = Nothing
    End If
    
    Set oConn = New ADODB.Connection
    Set madoRSClassType = New ADODB.Recordset
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    sSQL = "SELECT RetClassType.* "
    sSQL = sSQL & "FROM( "
    sSQL = sSQL & "SELECT "
    sSQL = sSQL & "[ClassTypeID], "
    sSQL = sSQL & "[Class], "
    sSQL = sSQL & "[Description], "
    sSQL = sSQL & "[IsDeleted] , "
    sSQL = sSQL & "[DateLastUpdated] , "
    sSQL = sSQL & "(SELECT    USERNAME "
    sSQL = sSQL & "FROM       USERS "
    sSQL = sSQL & "WHERE  USERSID = CT.[UpdateByUserID]) As [UpdateByUserName],  "
    sSQL = sSQL & "[UpdateByUserID]  "
    sSQL = sSQL & "FROM ClassType CT "
    sSQL = sSQL & "WHERE ClassTypeID IN( "
                        sSQL = sSQL & "SELECT   ClassTypeID "
                        sSQL = sSQL & "FROM     ClassOfLoss "
                        sSQL = sSQL & "WHERE    ClientCompanyID = " & goUtil.gsCurCar & " "
                        sSQL = sSQL & "AND      IsDeleted = 0 "
                        sSQL = sSQL & ")"
    sSQL = sSQL & ") RetClassType "
    sSQL = sSQL & "WHERE IsDeleted = 0 "
    sSQL = sSQL & "ORDER BY Class "
    
    madoRSClassType.CursorLocation = adUseClient
    madoRSClassType.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set madoRSClassType.ActiveConnection = Nothing
    
    clsCarGUI_SetadoClassType = True
    
    Set oConn = Nothing
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Set oConn = Nothing
    clsCarGUI_SetadoClassType = False
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarGUI_ClassName & vbCrLf & "Private Function clsCarGUI_SetadoClassType"
End Function

Private Function clsCarGUI_SetadoRSFeeScheduleFeeTypes(Optional psFeeScheduleID As String) As Boolean
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    'reset rs
    If Not madoRSFeeScheduleFeeTypes Is Nothing Then
        Set madoRSFeeScheduleFeeTypes = Nothing
    End If
    
    Set oConn = New ADODB.Connection
    Set madoRSFeeScheduleFeeTypes = New ADODB.Recordset
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    sSQL = "SELECT RetFeeScheduleFeeTypes.* "
    sSQL = sSQL & "FROM( "
    sSQL = sSQL & "SELECT "
    sSQL = sSQL & "[FeeScheduleFeeTypesID], "
    sSQL = sSQL & "[FeeScheduleID], "
    sSQL = sSQL & "[TypeNum], "
    sSQL = sSQL & "[Name], "
    sSQL = sSQL & "[Description], "
    sSQL = sSQL & "[FeeAmount], "
    sSQL = sSQL & "[IsExpense], "
    sSQL = sSQL & "[MaxNumberOfItems], "
    sSQL = sSQL & "[MaxFeeAmount], "
    sSQL = sSQL & "[IsMiscAmount], "
    sSQL = sSQL & "[UseFormula], "
    sSQL = sSQL & "[VBFormula], "
    sSQL = sSQL & "[IsDeleted], "
    sSQL = sSQL & "[DateLastUpdated], "
    sSQL = sSQL & "(SELECT    USERNAME "
    sSQL = sSQL & "FROM   USERS "
    sSQL = sSQL & "WHERE  USERSID = F.[UpdateByUserID]) As [UpdateByUserName],  "
    sSQL = sSQL & "[UpdateByUserID] "
    sSQL = sSQL & "FROM FeeScheduleFeeTypes F "
    sSQL = sSQL & ") RetFeeScheduleFeeTypes "
    If psFeeScheduleID = vbNullString Then
        sSQL = sSQL & "WHERE [FeeScheduleID] = ( "
                                sSQL = sSQL & "SELECT   [FeeScheduleID] "
                                sSQL = sSQL & "FROM     ClientCompanyCat "
                                sSQL = sSQL & "WHERE    [ClientCompanyID] = " & goUtil.gsCurCar & " "
                                sSQL = sSQL & "AND      [CATID] = " & goUtil.gsCurCat & " "
                                sSQL = sSQL & ") "
    Else
        sSQL = sSQL & "WHERE [FeeScheduleID] = ( "
                                sSQL = sSQL & "SELECT   " & psFeeScheduleID & " As [FeeScheduleID] "
                                sSQL = sSQL & "FROM     ClientCompanyCat "
                                sSQL = sSQL & "WHERE    [ClientCompanyID] = " & goUtil.gsCurCar & " "
                                sSQL = sSQL & "AND      [CATID] = " & goUtil.gsCurCat & " "
                                sSQL = sSQL & ") "
    End If
    sSQL = sSQL & "ORDER BY [TypeNum] "
    
    madoRSFeeScheduleFeeTypes.CursorLocation = adUseClient
    madoRSFeeScheduleFeeTypes.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set madoRSFeeScheduleFeeTypes.ActiveConnection = Nothing
    
    clsCarGUI_SetadoRSFeeScheduleFeeTypes = True
    
    Set oConn = Nothing
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Set oConn = Nothing
    clsCarGUI_SetadoRSFeeScheduleFeeTypes = False
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarGUI_ClassName & vbCrLf & "Private Function clsCarGUI_SetadoRSFeeScheduleFeeTypes"
End Function

Private Function clsCarGUI_SetadoRSFeeScheduleLevels(Optional psFeeScheduleID As String) As Boolean
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    'reset rs
    If Not madoRSFeeScheduleLevels Is Nothing Then
        Set madoRSFeeScheduleLevels = Nothing
    End If

    Set oConn = New ADODB.Connection
    Set madoRSFeeScheduleLevels = New ADODB.Recordset
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    sSQL = "SELECT RetFeeScheduleLevels.* "
    sSQL = sSQL & "FROM( "
    sSQL = sSQL & "SELECT "
    sSQL = sSQL & "[FeeScheduleLevelsID], "
    sSQL = sSQL & "[FeeScheduleID], "
    sSQL = sSQL & "[LevelNum], "
    sSQL = sSQL & "[LevelMax], "
    sSQL = sSQL & "[LevelPctApp], "
    sSQL = sSQL & "[LevelMin], "
    sSQL = sSQL & "[IsDeleted], "
    sSQL = sSQL & "[DateLastUpdated], "
    sSQL = sSQL & "(SELECT    USERNAME "
    sSQL = sSQL & "FROM   USERS "
    sSQL = sSQL & "WHERE  USERSID = F.[UpdateByUserID]) As [UpdateByUserName],  "
    sSQL = sSQL & "[UpdateByUserID] "
    sSQL = sSQL & "FROM FeeScheduleLevels F "
    sSQL = sSQL & ") RetFeeScheduleLevels "
    If psFeeScheduleID = vbNullString Then
        sSQL = sSQL & "WHERE [FeeScheduleID] = ( "
                                sSQL = sSQL & "SELECT   [FeeScheduleID] "
                                sSQL = sSQL & "FROM     ClientCompanyCat "
                                sSQL = sSQL & "WHERE    [ClientCompanyID] = " & goUtil.gsCurCar & " "
                                sSQL = sSQL & "AND      [CATID] = " & goUtil.gsCurCat & " "
                                sSQL = sSQL & ") "
    Else
        sSQL = sSQL & "WHERE [FeeScheduleID] = ( "
                                sSQL = sSQL & "SELECT   " & psFeeScheduleID & " As [FeeScheduleID] "
                                sSQL = sSQL & "FROM     ClientCompanyCat "
                                sSQL = sSQL & "WHERE    [ClientCompanyID] = " & goUtil.gsCurCar & " "
                                sSQL = sSQL & "AND      [CATID] = " & goUtil.gsCurCat & " "
                                sSQL = sSQL & ") "
    End If
    sSQL = sSQL & "ORDER BY [LevelNum] "


    madoRSFeeScheduleLevels.CursorLocation = adUseClient
    madoRSFeeScheduleLevels.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set madoRSFeeScheduleLevels.ActiveConnection = Nothing

    clsCarGUI_SetadoRSFeeScheduleLevels = True
    
    Set oConn = Nothing
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Set oConn = Nothing
    clsCarGUI_SetadoRSFeeScheduleLevels = False
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarGUI_ClassName & vbCrLf & "Private Function clsCarGUI_SetadoRSFeeScheduleLevels"
End Function

Private Function clsCarGUI_SetadoRSFeeSchedule(Optional psFeeScheduleID As String) As Boolean
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    'reset rs
    If Not madoRSFeeSchedule Is Nothing Then
        Set madoRSFeeSchedule = Nothing
    End If
    
    Set oConn = New ADODB.Connection
    Set madoRSFeeSchedule = New ADODB.Recordset
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    sSQL = "SELECT RetFeeSchedule.* "
    sSQL = sSQL & "FROM( "
    sSQL = sSQL & "SELECT "
    sSQL = sSQL & "[FeeScheduleID], "
    sSQL = sSQL & "[ClientCompanyID], "
    sSQL = sSQL & "[ScheduleName], "
    sSQL = sSQL & "[Description], "
    sSQL = sSQL & "[NumOfLevels], "
    sSQL = sSQL & "[NumOfFeeTypes], "
    sSQL = sSQL & "[FeeServiceHourlyRate], "
    sSQL = sSQL & "[TaxPercent], "
    sSQL = sSQL & "[InitialOptions], "
    sSQL = sSQL & "[Options], "
    sSQL = sSQL & "[DefaultAppDedClassTypeIDOrder], "
    sSQL = sSQL & "[IsDeleted], "
    sSQL = sSQL & "[DateLastUpdated], "
    sSQL = sSQL & "(SELECT    USERNAME "
    sSQL = sSQL & "FROM   USERS "
    sSQL = sSQL & "WHERE  USERSID = F.[UpdateByUserID]) As [UpdateByUserName],  "
    sSQL = sSQL & "[UpdateByUserID] "
    sSQL = sSQL & "FROM FeeSchedule F "
    sSQL = sSQL & ") RetFeeSchedule "
    If psFeeScheduleID = vbNullString Then
        sSQL = sSQL & "WHERE [FeeScheduleID] = ( "
                                sSQL = sSQL & "SELECT   [FeeScheduleID] "
                                sSQL = sSQL & "FROM     ClientCompanyCat "
                                sSQL = sSQL & "WHERE    [ClientCompanyID] = " & goUtil.gsCurCar & " "
                                sSQL = sSQL & "AND      [CATID] = " & goUtil.gsCurCat & " "
                                sSQL = sSQL & ") "
    Else
        sSQL = sSQL & "WHERE [FeeScheduleID] = ( "
                                sSQL = sSQL & "SELECT   " & psFeeScheduleID & " As [FeeScheduleID] "
                                sSQL = sSQL & "FROM     ClientCompanyCat "
                                sSQL = sSQL & "WHERE    [ClientCompanyID] = " & goUtil.gsCurCar & " "
                                sSQL = sSQL & "AND      [CATID] = " & goUtil.gsCurCat & " "
                                sSQL = sSQL & ") "
    End If
    sSQL = sSQL & "AND [ClientCompanyID] = " & goUtil.gsCurCar & " "
    
    madoRSFeeSchedule.CursorLocation = adUseClient
    madoRSFeeSchedule.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set madoRSFeeSchedule.ActiveConnection = Nothing
    
    clsCarGUI_SetadoRSFeeSchedule = True
    
    Set oConn = Nothing
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Set oConn = Nothing
    clsCarGUI_SetadoRSFeeSchedule = False
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarGUI_ClassName & vbCrLf & "Private Function clsCarGUI_SetadoRSFeeSchedule"
End Function

Private Function clsCarGUI_CLEANUP() As Boolean
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    If Not mfrmClaimsList Is Nothing Then
        mfrmClaimsList.CLEANUP
        Unload mfrmClaimsList
        Set mfrmClaimsList = Nothing
    End If
    If Not mfrmAdjusterReports Is Nothing Then
        mfrmAdjusterReports.CLEANUP
        Unload mfrmAdjusterReports
        Set mfrmAdjusterReports = Nothing
    End If
    If Not mfrmFeeSchedule Is Nothing Then
        mfrmFeeSchedule.CLEANUP
        Unload mfrmFeeSchedule
        Set mfrmFeeSchedule = Nothing
    End If
    
    Set madoRSTypeOfLoss = Nothing
    Set madoRSAssignmentType = Nothing
    Set madoRSACID = Nothing
    Set madoRSCatCode = Nothing
    Set madoRSState = Nothing
    Set madoRSClassType = Nothing
    Set madoRSFeeSchedule = Nothing
    Set madoRSFeeScheduleLevels = Nothing
    Set madoRSFeeScheduleFeeTypes = Nothing
    
    clsCarGUI_CLEANUP = True
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    clsCarGUI_CLEANUP = False
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarGUI_ClassName & vbCrLf & "Private Function CLEANUP"
End Function

