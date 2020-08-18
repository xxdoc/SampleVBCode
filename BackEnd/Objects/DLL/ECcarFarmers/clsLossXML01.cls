VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLossXML01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Loss Report Polymorphism
'Loss Report InterFace Definiton
'Use Implements V2ECKeyBoard.clsCarLR in ALL Loss Report Classes For All Carrier Objects

Implements V2ECKeyBoard.clsCarLR

'BGS 3.13.2002 Need a UDT for Admin Loss Info

Public Type udtXML01AssignmentDetail
    UniqueID As String
    UnitNumber As String
    ClaimNumber As String
    DateAssigned As String
    CatastropheCode As String
    LossDate As String
    Type As String
    CauseOfLoss As String
    FirstName As String
    LastName As String
    AssignedTo As String
    AssignedToFirstName As String
    AssignedToLastName As String
End Type

Public Type udtXML01LossDetail
    UniqueID As String
    LossLocationAddress As String
    LossLocationAddress2 As String
    LossLocationCity As String
    LossLocationState As String
    LossLocationZip As String
    PropertyAddress As String
    PropertyCity As String
    PropertyState As String
    PropertyZip As String
    AffectedAreas As String
    LossDescription As String
End Type

Public Type udtXML01VehicleDetail
    UniqueID As String
    VehicleMake As String
    VehicleModel As String
    VehicleYear As String
    VIN As String
    PropertyItemName As String
    DamageDescription As String
    LossDescription As String
    LocationType As String
    LocationName As String
    LocationPhoneNumber As String
    LocationAddress As String
    LocationCity As String
    LocationState As String
    LocationPostalCode As String
End Type

Public Type udtXML01ContactDetail
    UniqueID As String
    AgentFirstName As String
    AgentLastName As String
    AgentPrimaryPhone As String
    ContactRowID As String
End Type

Public Type udtXML01Contacts
    ContactRowID As String
    MM As String
    FirstName As String
    LastName As String
    ContactRole As String
    BestPhoneNumber As String
    PrimaryPhoneNumber As String
    HomePhoneNumber As String
    CellularPhoneNumber As String
    FaxPhoneNumber As String
    EmailAddress As String
    PagerPhoneNumber As String
    PagerPIN As String
    WorkPhoneNumber As String
    RelationshipToInsured As String
    PrimaryAddressID As String
    CreatedUpdated As Date
End Type

Public Type udtXML01Address
    ContactRowID As String
    AddressID As String
    Type As String
    StreetAddress As String
    StreetAddress2 As String
    City As String
    State As String
    PostalCode As String
End Type

'Policy Detail
Public Type udtXML01PolicyDetail
    PolicyNumber As String
    Status As String
    CoverageStatus As String
    BalanceDue As String
    CompanyCode As String
    CompanyName As String
    RenewalDate As String
    CancellationDate As String
    NewBusinessDate As String
    PolicyDescription As String
    MortgageeName As String
    MortgageeAddress As String
    PolicyEdition As String
    LienHolderName As String
    LienHolderAddress As String
End Type

'Coverage
Public Type udtXML01Coverages
    CoverageRowID As String
    Coverage As String
    Limits As String
    Deductible1 As String
    Deductible2 As String
    Deductible3 As String
    Deductible4 As String
End Type

'Endorsements UDTs will be Stored in varyEndorsements
'since the max is 10 but there could be anywhere from
'4 to 10 EDs. which the XML01 Loss report does not print
'null strings for the remainaing 10 items.
Public Type udtXML01Endorsements
    EndorsementsRowID As String
    EndorsementNumber As String
    EndorsementDescription As String
End Type

Public Type udtXML01PaymentDetail '---------->PLH
    PaymentDetailRowID As String
    DateIssued As String
    PayeeLineOne As String
    PayeeLineTwo As String
    PayeeLineThree As String
    PayeeLineFour As String
    AccountType As String
    PaymentClass As String
    PaymentAmount As String
End Type

'PriorLossHist UDTs will be stored in colPLH
'Since there may be multiple PriorLoss entries.
'As well there can be continuation pages that have
'Prior Loss entries continued.
Public Type udtXML01PriorLossDetail '---------->PLH
    PriorLossRowID As String
    SALN As String
    ClaimSegmentNumber As String
    PolicyNumber As String
    LossCause As String
    ClaimClass As String
    LossDate As String
    SummaryAmount As String
End Type

'CommentsActLog UDTs will be stored in colCAL
'Since there may be multiple Comments Act Log entries.
'As well, there can be Continuation pages that have
'Comments Act log entries continued.
Public Type udtXML01Activities '---------->CAL
    ActivityRowID As String
    GMTCreated As String
    CreatedBy As String
    Type As String
    Description As String
    Comment As String
End Type

'This is the entire Loss Report (Main page)
Public Type udtXML01Loss
    AssignmentDetailRS As WDDXRecordset
    LossDetailRS As WDDXRecordset
    VehicleDetailRS As WDDXRecordset
    ContactDetailRS As WDDXRecordset
    ContactsRS As WDDXRecordset
    AddressRS As WDDXRecordset
    PolicyDetailRS As WDDXRecordset
    CoverageRS As WDDXRecordset
    EndorsementRS As WDDXRecordset
    PaymentDetailRS As WDDXRecordset
    PriorLossDetailRS As WDDXRecordset
    ActivitiesRS As WDDXRecordset
End Type

Public Type XML01LossReport
    XML01Loss As udtXML01Loss
    PrnData As String 'Contains Page1 and all continued page data
    PrnKey As String 'File path for this Loss Report
End Type

Public Enum TypeXML01
    XML01Apd = 0
    XML01Pro
End Enum

Private mudtXML01LossReport As XML01LossReport
Private moLRs As V2ECKeyBoard.clsLossReports
Private msAssignedToFirstName As String
Private msAssignedToLastName As String
Private msInsuredName As String
Private msWorkPhone As String
Private msHomePhone As String
Private msDateAssign As String
Private msAssignmentType As String
Private msStatus As String
Private msCatName As String
Private msCatCode As String
Private msAdjuster As String
Private msACID As String
Private msCLIENTNUM As String
Private msIBNUM As String
Private msTypeOfACID As String
Private mLossType As TypeXML01
Private msOleType As String
Private mbAbortProcessRawData As Boolean
Private msUniqueID As String
Private msClaimId As String
'*****************UPDATE UNIT FLAGS************************************
'*
'*
'**********************************************************************
'Flags to indicate address updates required
Private mbPropertyAddressUpdateRequired As Boolean
Private mbMaillingAddressUpdateRequired As Boolean
'Flags for other Assignments Fields
Private mbPolicyNoUpdateRequired As Boolean
Private mbPolicyDescriptionUpdateRequired As Boolean
Private mbInsuredUpdateRequired As Boolean
Private mbHomePhoneUpdateRequired As Boolean
Private mbBusinessPhoneUpdateRequired As Boolean
Private mbMortgageeNameUpdateRequired As Boolean
Private mbAgentNoUpdateRequired As Boolean
Private mbDeductibleUpdateRequired As Boolean
'Flags For Vendor update Synchro
Private mbFarmersXML01_LocationUpdateQueUpdateRequired As Boolean
'*****************UPDATE UNIT FLAGS************************************
'*
'*
'**********************************************************************

Public Property Get ClaimId() As String
    ClaimId = msClaimId
End Property
Public Property Get UniqueID() As String
    UniqueID = msUniqueID
End Property

Private Sub Class_Initialize()
    Set goUtil = New V2ECKeyBoard.clsUtil
End Sub

Private Sub Class_Terminate()
    
    If Not mudtXML01LossReport.XML01Loss.AssignmentDetailRS Is Nothing Then
        Set mudtXML01LossReport.XML01Loss.AssignmentDetailRS = Nothing
    End If
    If Not mudtXML01LossReport.XML01Loss.ContactDetailRS Is Nothing Then
        Set mudtXML01LossReport.XML01Loss.ContactDetailRS = Nothing
    End If
    If Not mudtXML01LossReport.XML01Loss.ContactDetailRS Is Nothing Then
        Set mudtXML01LossReport.XML01Loss.ContactDetailRS = Nothing
    End If
    If Not mudtXML01LossReport.XML01Loss.ContactsRS Is Nothing Then
        Set mudtXML01LossReport.XML01Loss.ContactsRS = Nothing
    End If
    If Not mudtXML01LossReport.XML01Loss.CoverageRS Is Nothing Then
        Set mudtXML01LossReport.XML01Loss.CoverageRS = Nothing
    End If
    If Not mudtXML01LossReport.XML01Loss.EndorsementRS Is Nothing Then
        Set mudtXML01LossReport.XML01Loss.EndorsementRS = Nothing
    End If
    If Not mudtXML01LossReport.XML01Loss.LossDetailRS Is Nothing Then
        Set mudtXML01LossReport.XML01Loss.LossDetailRS = Nothing
    End If
    If Not mudtXML01LossReport.XML01Loss.PaymentDetailRS Is Nothing Then
        Set mudtXML01LossReport.XML01Loss.PaymentDetailRS = Nothing
    End If
    If Not mudtXML01LossReport.XML01Loss.PriorLossDetailRS Is Nothing Then
        Set mudtXML01LossReport.XML01Loss.PriorLossDetailRS = Nothing
    End If
    If Not mudtXML01LossReport.XML01Loss.PolicyDetailRS Is Nothing Then
        Set mudtXML01LossReport.XML01Loss.PolicyDetailRS = Nothing
    End If
    If Not mudtXML01LossReport.XML01Loss.AddressRS Is Nothing Then
        Set mudtXML01LossReport.XML01Loss.AddressRS = Nothing
    End If
    If Not mudtXML01LossReport.XML01Loss.VehicleDetailRS Is Nothing Then
        Set mudtXML01LossReport.XML01Loss.VehicleDetailRS = Nothing
    End If
End Sub

Private Property Let clsCarLR_AbortProcessRawData(pbFlag As Boolean)
    mbAbortProcessRawData = pbFlag
End Property

Private Property Get clsCarLR_AbortProcessRawData() As Boolean
    clsCarLR_AbortProcessRawData = mbAbortProcessRawData
End Property

Private Property Let clsCarLR_Adjuster(psAdjuster As String)
    msAdjuster = psAdjuster
End Property

Private Property Get clsCarLR_Adjuster() As String
    clsCarLR_Adjuster = msAdjuster
End Property

Private Property Let clsCarLR_AssignmentType(psAssignmentType As String)
    msAssignmentType = psAssignmentType
End Property
Private Property Get clsCarLR_AssignmentType() As String
    clsCarLR_AssignmentType = msAssignmentType
End Property

Public Property Get clsCarLR_CatName() As String
    clsCarLR_CatName = msCatName
End Property
Public Property Let clsCarLR_CatName(psCatName As String)
    msCatName = psCatName
End Property

Public Property Get clsCarLR_CatCode() As String
    clsCarLR_CatCode = msCatCode
End Property
Public Property Let clsCarLR_CatCode(psCatCode As String)
    msCatCode = psCatCode
End Property

Private Property Let clsCarLR_CLIENTNUM(psCLIENTNUM As String)
    msCLIENTNUM = psCLIENTNUM
End Property

Private Property Get clsCarLR_CLIENTNUM() As String
    clsCarLR_CLIENTNUM = msCLIENTNUM
End Property

Private Property Let clsCarLR_DateAssign(psDateAssign As String)
    msDateAssign = psDateAssign
End Property

Private Property Get clsCarLR_DateAssign() As String
    clsCarLR_DateAssign = msDateAssign
End Property

Private Function clsCarLR_GetLRSQL(psAppName As String, plVersion As Long, Optional psSQLError As String) As String
    On Error GoTo EH
    Dim sSQL As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    'Need to ucase the appname because ME will and who knows
    'maybe other windows OS will UCase the app exe name why I have not a friggin clue
    Select Case UCase(psAppName)
'<--------------------------------------WEBCONTROL-------------------------------------->
        Case "WEBCONTROL", "V2WEBCONTROL"
            Select Case plVersion
                Case Is >= 200
                    sSQL = SQL_WEBCONTROL_20000(psSQLError)
                'iF Newer versions Require Changes to SQL ...
                'They will Have  go on top in Select Case
                    
            End Select

    End Select
    
    If sSQL = vbNullString Then
        psSQLError = psAppName & " VS " & plVersion & " Not supported." & vbCrLf
    Else
        clsCarLR_GetLRSQL = sSQL
    End If
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLR_ClassName & vbCrLf & "Public Function GetLRSQL"
End Function

Private Property Let clsCarLR_HomePhone(psHomePhone As String)
    msHomePhone = psHomePhone
End Property

Private Property Get clsCarLR_HomePhone() As String
    clsCarLR_HomePhone = msHomePhone
End Property

Private Property Let clsCarLR_IBNUM(psIBNUM As String)
    msIBNUM = psIBNUM
End Property

Private Property Get clsCarLR_IBNUM() As String
    clsCarLR_IBNUM = msIBNUM
End Property

Private Property Let clsCarLR_InsuredName(psInsuredName As String)
    msInsuredName = psInsuredName
End Property

Private Property Get clsCarLR_InsuredName() As String
    clsCarLR_InsuredName = msInsuredName
End Property

Public Property Let clsCarLR_LossType(pType As Variant)
    If CStr(pType) = "Property" Then
        mLossType = XML01Pro
    ElseIf CStr(pType) = "Auto" Then
        mLossType = XML01Apd
    Else
        mLossType = pType
    End If
End Property
Public Property Get clsCarLR_LossType() As Variant
    clsCarLR_LossType = mLossType
End Property

Public Property Get clsCarLR_LossReport() As Variant
    clsCarLR_LossReport = mudtXML01LossReport
End Property

Public Property Let clsCarLR_LRs(poLRs As V2ECKeyBoard.clsLossReports)
    Set moLRs = poLRs
End Property
Public Property Set clsCarLR_LRs(poLRs As V2ECKeyBoard.clsLossReports)
    Set moLRs = poLRs
End Property
Public Property Get clsCarLR_LRs() As V2ECKeyBoard.clsLossReports
    Set clsCarLR_LRs = moLRs
End Property

Private Property Get clsCarLR_MaxAllowedInChain() As Long
    clsCarLR_MaxAllowedInChain = 100 'XML01 has a lot of overhead when chaining, so only 100
End Property

Private Property Let clsCarLR_OleType(psOleType As String)
    msOleType = psOleType
End Property
Private Property Get clsCarLR_OleType() As String
    clsCarLR_OleType = msOleType
End Property

Public Property Let clsCarLR_PrnData(psData As String)
    mudtXML01LossReport.PrnData = psData
End Property
Public Property Get clsCarLR_PrnData() As String
    clsCarLR_PrnData = mudtXML01LossReport.PrnData
End Property

Public Property Let clsCarLR_PrnKey(psKEY As String)
    mudtXML01LossReport.PrnKey = psKEY
End Property
Public Property Get clsCarLR_PrnKey() As String
    clsCarLR_PrnKey = mudtXML01LossReport.PrnKey
End Property

Public Property Get clsCarLR_ClassName() As String
    clsCarLR_ClassName = App.EXEName & ".clsLossXML01"
End Property

Public Property Get clsCarLR_GenTextHeight() As Long
    clsCarLR_GenTextHeight = 165
End Property
Public Property Get clsCarLR_GenTextFontName() As String
    clsCarLR_GenTextFontName = "Courier New"
End Property
Public Property Get clsCarLR_GenTextFontSize() As Long
    clsCarLR_GenTextFontSize = 11
End Property

Public Property Let clsCarLR_ACID(psACID As String)
    msACID = psACID
End Property
Public Property Get clsCarLR_ACID() As String
    clsCarLR_ACID = msACID
End Property

Private Property Let clsCarLR_Status(psStatus As String)
    msStatus = psStatus
End Property

Private Property Get clsCarLR_Status() As String
    clsCarLR_Status = msStatus
End Property

Private Property Let clsCarLR_TypeOfACID(psTypeOfACID As String)
    msTypeOfACID = psTypeOfACID
End Property

Private Property Get clsCarLR_TypeOfACID() As String
    clsCarLR_TypeOfACID = msTypeOfACID
End Property

Public Function GetXMLResultDocResponse(psData As String) As String
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim sResult As String
    Dim sAccepted As String
    Dim sReasonNotAccepted As String
    Dim sData As String
    'Farmers Doc Resposnse XML Data Vars
    Dim sAdminComments As String
    Dim sUnitNumber As String
    Dim sUnitStatus As String
    Dim sGUID As String
    Dim sStatus As String
    Dim sComments As String
    'SQL Server params for Doc Approval/ Reject / Delete
    Dim sIsClientCoReject As String    'Bit = Null, ""
    Dim sClientCoRejectDate As String   'DateTime = Null, ""
    Dim sClientCoRejectDesc As String   'VarChar(100) = Null, ""
    Dim sIsClientCoDelete As String    'Bit = Null, ""
    Dim sClientCoDeleteDate As String   'DateTime = Null, ""
    Dim sClientCoDeleteDesc As String   'VarChar(100) = Null, ""
    Dim sIsClientCoApprove As String   'Bit = Null, ""
    Dim sClientCoApproveDate As String  'DateTime = Null, ""
    Dim sClientCoApproveDesc As String  'varchar(100)=null"
    'Data base Objects
    Dim oConn As New ADODB.Connection
    Dim sSQL As String
    Dim lRecordsAffected As Long
    
    
    sData = psData
    
    'need to parse out stuff from the Farmers Doc Response XML Data
    'Format given to us from Micahel Modica 6.14.2005 11:49 AM Moutain time
'    <Result UnitNumber="1234" UnitStatus="PendingApproval">
'     <ResultOutboundDocument GUID="xyz">
'      <Status>Accepted</Status>
'      <Comments>If Rejected or Deleted, this will be populated..optional</Comments>
'     </ResultOutboundDocument>
'    </Result>
    
    sUnitNumber = ParseNodeOrAttribute(sData, "UnitNumber=")
    sUnitStatus = ParseNodeOrAttribute(sData, "UnitStatus=")
    sGUID = ParseNodeOrAttribute(sData, "GUID=")
    sStatus = ParseNodeOrAttribute(sData, "Status")
    sComments = ParseNodeOrAttribute(sData, "Comments")
    
    'Set the Database Objects
    Set oConn = New ADODB.Connection
    OpenConnection oConn
    
    
    'Build the SP SQL
    'These ITEMS MUST BE NULL !!!!
    'These items affect the Company (Eberls) level
    'Updates to the status of a package
'    [PackageStatus]     = (CASE WHEN @PackageStatus Is Null     THEN [PackageStatus]    ELSE @PackageStatus     END),
'    [SendMe]        = (CASE WHEN @SendMe Is Null        THEN [SendMe]       ELSE @SendMe        END),
'    [SentDate]      = (CASE WHEN @SentDate Is Null      THEN [SentDate]     ELSE @SentDate      END),
'    [SentToEmail]       = (CASE WHEN @SentToEmail Is Null   THEN [SentToEma
    
    
    'END THESE ITEMS MUST BE NULL !!!!
    
    'Use the Unit Status to update the package Level Admin Comments.
    'For Farmers it will be a combination of Unit Number and the Status value
    sAdminComments = "Farmers Unit Status Update:" & vbCrLf
    sAdminComments = sAdminComments & "Unit Number: " & sUnitNumber & vbCrLf
    sAdminComments = sAdminComments & "Status: " & sUnitStatus
    
    'populate the SQL Server params for Doc Approval/ Reject / Delete
    'According to Email from Farmers from jeffrey.johnson@farmersinsurance.com
    'Dated Sent: Tuesday, May 24, 2005 3:06 PM
    'These are the only Status that will appear in the Doc Response XML under the
    '<Status></Status> Node.
    'Deleted , Rejected Or Approved
    'Use Reject instead of Rejected
    If InStr(1, sStatus, "Reject", vbTextCompare) > 0 Then
        sIsClientCoReject = "1, "
        sClientCoRejectDate = "'" & Format(Now(), "MM/DD/YYYY HH:mm:ss") & "', "
        sClientCoRejectDesc = "'" & goUtil.utCleanSQLString(sComments) & "', "
    Else
        sIsClientCoReject = "0, "
        sClientCoRejectDate = "Null, "
        sClientCoRejectDesc = "Null, "
    End If
    'Use Delete instead of Deleted
    If InStr(1, sStatus, "Delete", vbTextCompare) > 0 Then
        sIsClientCoDelete = "1, "
        sClientCoDeleteDate = "'" & Format(Now(), "MM/DD/YYYY HH:mm:ss") & "', "
        sClientCoDeleteDesc = "'" & goUtil.utCleanSQLString(sComments) & "', "
    Else
        sIsClientCoDelete = "0, "
        sClientCoDeleteDate = "Null, "
        sClientCoDeleteDesc = "Null, "
    End If
    'Use Approve instead of Approved
    If InStr(1, sStatus, "Approve", vbTextCompare) > 0 Then
        sIsClientCoApprove = "1, "
        sClientCoApproveDate = "'" & Format(Now(), "MM/DD/YYYY HH:mm:ss") & "', "
        'Since this is the last param don't include the comma
        sClientCoApproveDesc = "'" & goUtil.utCleanSQLString(sComments) & "' "
    Else
        sIsClientCoApprove = "0, "
        sClientCoApproveDate = "Null, "
        'Since this is the last param don't include the comma
        sClientCoApproveDesc = "Null "
    End If
    
    sSQL = "Exec z_spuUpdatePackageStatus "
    sSQL = sSQL & "Null, "                                                  '@UID               int,        The UID used for this update is CFUSER Figured out in the SP
    sSQL = sSQL & "Null, "                                                  '@AssignmentsID     int,        NA for Client Updates since it depends on the @PackageItemGUID
    sSQL = sSQL & "Null, "                                                  '@PackageNumber     int,        NA ditto @PackageItemGUID
    sSQL = sSQL & "Null, "                                                  '@SendMe            bit=null,   NA ditto @PackageItemGUID
    sSQL = sSQL & "Null, "                                                  '@StatusID          int=null,   NA ditto @PackageItemGUID
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(sAdminComments) & "', "     '@AdminComments     varchar(1000)=null,"
    sSQL = sSQL & "Null, "                                                  '@SentToEmail        varchar(50)=null,"
    sSQL = sSQL & "Null, "                                                  '@SentDate       DateTime=null,"
    sSQL = sSQL & "Null, "                                                  '@ApprovePackageItems    bit=null,"
    'Updating Status via at the Document Level
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(sGUID) & "', "              '@PackageItemGUID    varchar(50)=null,"
    sSQL = sSQL & "1, "                                                     '@bClientUpdate      bit=0,"
    'The Quotes will be figured out in the var for Reject / Delete / Approve
    sSQL = sSQL & sIsClientCoReject                                         '@IsClientCoReject   bit=null,"
    sSQL = sSQL & sClientCoRejectDate                                       '@ClientCoRejectDate datetime=null,"
    sSQL = sSQL & sClientCoRejectDesc                                       '@ClientCoRejectDesc varchar(100)=null,"
    sSQL = sSQL & sIsClientCoDelete                                         '@IsClientCoDelete   bit=null,"
    sSQL = sSQL & sClientCoDeleteDate                                       '@ClientCoDeleteDate datetime=null,"
    sSQL = sSQL & sClientCoDeleteDesc                                       '@ClientCoDeleteDesc varchar(100)=null,"
    sSQL = sSQL & sIsClientCoApprove                                        '@IsClientCoApprove  bit=null,"
    sSQL = sSQL & sClientCoApproveDate                                      '@ClientCoApproveDate    datetime=null,"
    sSQL = sSQL & sClientCoApproveDesc                                      '@ClientCoApproveDesc    varchar(100)=null"
    
    oConn.Execute sSQL, lRecordsAffected
    
    
    sAccepted = CStr(CBool(lRecordsAffected))
    
    If lRecordsAffected = 0 Then
        sReasonNotAccepted = "No Records Affected!"
    End If
    
    'cleanup
    Set oConn = Nothing
    
    GoTo GIVE_RESULT
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    sAccepted = "False"
    sReasonNotAccepted = "Error Number: " & lErrNum & vbCrLf
    sReasonNotAccepted = sReasonNotAccepted & "Error Description: " & sErrDesc
GIVE_RESULT:
    sResult = vbNullString
    sResult = sResult & "<?xml version=""1.0"" encoding=""utf-8"" standalone=""yes""?>"
    sResult = sResult & "<Result GUID=""" & CleanXML(sGUID) & """>"
        sResult = sResult & "<ResultDocResponse Accepted=""" & CleanXML(sAccepted) & """ Reason=""" & CleanXML(sReasonNotAccepted) & """ />"
    sResult = sResult & "</Result>"
    
    GetXMLResultDocResponse = sResult
    
End Function

Public Function ParseNodeOrAttribute(psData As String, psNodeName As String) As String
    On Error GoTo EH
    Dim sData As String
    Dim sNodeName As String
    
    sData = psData
    sNodeName = psNodeName
    
    'Check for Attribute
    If InStr(1, sNodeName, "=", vbBinaryCompare) > 0 Then
        'parsing an attribute
        sData = Mid(sData, InStr(1, sData, sNodeName, vbBinaryCompare) + Len(sNodeName) + 1)
        sData = left(sData, InStr(1, sData, Chr(34), vbBinaryCompare) - 1)
    Else
        'parsing a Node
        sData = Mid(sData, InStr(1, sData, "<" & sNodeName & ">", vbBinaryCompare) + Len("<" & sNodeName & ">"))
        sData = left(sData, InStr(1, sData, "</" & sNodeName & ">", vbBinaryCompare) - 1)
    End If
    
    
    ParseNodeOrAttribute = sData
    
    Exit Function
EH:
    ParseNodeOrAttribute = vbNullString
End Function

Public Function GetXMLResultClose(psEntityID As String, pbAccepted As Boolean, Optional psReasonNotAccepted As String = vbNullString) As String
    Dim sResult As String
    Dim sAccepted As String
    Dim sReasonNotAccepted As String
    
    sAccepted = CStr(pbAccepted)
    
    sReasonNotAccepted = psReasonNotAccepted
    
    sResult = ""
    sResult = sResult & "<?xml version=""1.0"" encoding=""utf-8"" standalone=""yes""?>"
    sResult = sResult & "<Result EntityId=""" & CleanXML(psEntityID) & """>"
        sResult = sResult & "<ResultClose Accepted=""" & CleanXML(sAccepted) & """ Reason=""" & sReasonNotAccepted & """ />"
    sResult = sResult & "</Result>"
    
    GetXMLResultClose = sResult
    
End Function


Public Function GetXMLResultUpdate(psEntityID As String, pbAccepted As Boolean, Optional psReasonNotAccepted As String = vbNullString) As String
    Dim sResult As String
    Dim sAccepted As String
    
    sAccepted = CStr(pbAccepted)
    
    sResult = ""
    sResult = sResult & "<?xml version=""1.0"" encoding=""utf-8"" standalone=""yes""?>"
    sResult = sResult & "<Result EntityId=""" & CleanXML(psEntityID) & """>"
        sResult = sResult & "<ResultUnitUpdate/>"
    sResult = sResult & "</Result>"
    
    GetXMLResultUpdate = sResult

End Function

Public Function GetXMLResultAssignment(psEntityID As String, pbAccepted As Boolean, Optional psReasonNotAccepted As String = vbNullString) As String
    Dim sResult As String
    Dim sAccepted As String
    Dim sReasonNotAccepted As String
    
    sAccepted = CStr(pbAccepted)
    
    sReasonNotAccepted = psReasonNotAccepted
    sReasonNotAccepted = Replace(sReasonNotAccepted, "[Microsoft][ODBC SQL Server Driver][SQL Server]", vbNullString, , , vbTextCompare)
    
    sResult = ""
    sResult = sResult & "<?xml version=""1.0"" encoding=""utf-8"" standalone=""yes""?>"
    sResult = sResult & "<Result EntityId=""" & CleanXML(psEntityID) & """>"
        sResult = sResult & "<ResultAssignment Accepted=""" & CleanXML(sAccepted) & """ Reason=""" & CleanXML(sReasonNotAccepted) & """ />"
    sResult = sResult & "</Result>"
    
    GetXMLResultAssignment = sResult
    
End Function

Public Function GetXMLResultError(psEntityID As String, psErrName As String, psErrDesc As String) As String
    Dim sResult As String
    Dim sErrorName As String
    Dim sErrorDesc As String
    
    sErrorName = psErrName
    sErrorDesc = psErrDesc
    
    sResult = ""
    sResult = sResult & "<?xml version=""1.0"" encoding=""utf-8"" standalone=""yes""?>"
    sResult = sResult & "<Result EntityId=""" & CleanXML(psEntityID) & """>"
        sResult = sResult & "<ResultError>"
            sResult = sResult & "<Name>" & CleanXML(sErrorName) & "</Name>"
            sResult = sResult & "<Description>" & CleanXML(sErrorDesc) & "</Description>"
        sResult = sResult & "</ResultError>"
    sResult = sResult & "</Result>"
    
    GetXMLResultError = sResult
    
End Function

Public Function UpdateUnit(psEditLRData As String, _
                            Optional pbMergeUnit As Boolean) As Boolean
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
   
    Dim oConn As New ADODB.Connection
    Dim adoRS As ADODB.Recordset
    Dim sSQL As String

    Dim oXML01 As clsLossXML01
    Dim LossReport  As XML01LossReport
    Dim LossReportEdit As XML01LossReport
    Dim sLRData As String
    Dim oXML01Edit As clsLossXML01
    Dim sEditLRData As String
    Dim sAssignmentsID As String
    Dim sUniqueID As String
    Dim sOldUniqueID As String
    Dim bRSIsNothing As Boolean
    Dim saryNewCols() As String
    Dim lNewColsPos As Long
    Dim lAddNewRowPos As Long
    Dim lEditRowPos As Long
    Dim oWddxStruc As WDDXStruct
    Dim oWddxSer As WDDXSerializer
    Dim sOldHeader As String
    Dim ContactDetail As udtXML01ContactDetail
    Dim Contact As udtXML01Contacts
    Dim sTemp As String
    Dim sTemp2 As String
    Dim sCovType As String
    Dim sUpdUnitId As String
    Dim sUpdClaimId As String
    Dim sUpdateUnitReport As String
    '************BEGIN Merging Unit to Existing Assignment**********
    'Existing Assignment Looup Vars
    Dim sCLIENTNUM As String
    Dim sACIDAssignedTo As String
    Dim sLRFormat As String
    'Insert Policy limits Vars
    
    '************BEGIN Merging Unit to Existing Assignment**********
    
'*****************UPDATE UNIT FLAGS************************************
'*
'*                      RESET
'*
'**********************************************************************
    'Flags to indicate address updates required
    mbPropertyAddressUpdateRequired = False
    mbMaillingAddressUpdateRequired = False
    'Flags for other Assignments Fields
    mbPolicyNoUpdateRequired = False
    mbPolicyDescriptionUpdateRequired = False
    mbInsuredUpdateRequired = False
    mbHomePhoneUpdateRequired = False
    mbBusinessPhoneUpdateRequired = False
    mbMortgageeNameUpdateRequired = False
    mbAgentNoUpdateRequired = False
    mbDeductibleUpdateRequired = False
    'Flags For Vendor update Synchro
    mbFarmersXML01_LocationUpdateQueUpdateRequired = False
'*****************UPDATE UNIT FLAGS************************************
'*
'*                      RESET
'*
'**********************************************************************
    
    sEditLRData = psEditLRData
    Set oXML01Edit = New clsLossXML01
    
    oXML01Edit.clsCarLR_PrnData = sEditLRData
    oXML01Edit.clsCarLR_PopulateMe
    
    LossReportEdit = oXML01Edit.clsCarLR_LossReport
    
    'Need to Get the Loss Report Data for the same Assignment/ Unit that
    'Is being edited, so well, It can be edited.
    Set oConn = New ADODB.Connection
    OpenConnection oConn
    
    sUpdUnitId = oXML01Edit.UniqueID
    sUpdClaimId = oXML01Edit.ClaimId
    
    If Not pbMergeUnit Then
        sSQL = "Exec z_spuEditECSFarmersXML01LossReport "
        sSQL = sSQL & "0, "     '@bMergeUnit         bit=0,
        sSQL = sSQL & "null, "  '@CLIENTNUM          varchar(20)=null,
        sSQL = sSQL & "null, "  '@ACIDAssignedTo         varchar(20)=null,
        sSQL = sSQL & "null, "  '@LRFormat           varchar(50)=null,
        sSQL = sSQL & "null, "  '@UnitLimit          varchar(20)='0',
        sSQL = sSQL & "null, "  '@UnitLimitOther         varchar(100)=' ',
        sSQL = sSQL & "null, "  '@UnitClass          varchar(1)='OTHER',
        sSQL = sSQL & "null, "  '@UnitCauseOfLoss        varchar(50)='UNKNOWN',
        sSQL = sSQL & "'" & sUpdUnitId & "' " '@UnitID             varchar(20)=null,
    Else
        'Need to Pass In PolicyLimits Stuff When Merging a Unit to
        'An existing Assignment Record...
        sSQL = "Exec z_spuEditECSFarmersXML01LossReport "
        sSQL = sSQL & "1, "     '@bMergeUnit         bit=0,
        sTemp = LossReportEdit.XML01Loss.AssignmentDetailRS.getField(1, "ClaimNumber")
        sSQL = sSQL & "'" & Replace(sTemp, "'", "''") & "', "  '@CLIENTNUM          varchar(20)=null,
        sTemp = LossReportEdit.XML01Loss.AssignmentDetailRS.getField(1, "AssignedTo")
        sSQL = sSQL & "'" & Replace(sTemp, "'", "''") & "', "  '@ACIDAssignedTo         varchar(20)=null,
        sTemp = oXML01Edit.clsCarLR_ClassName
        sSQL = sSQL & "'" & Replace(sTemp, "'", "''") & "', "  '@LRFormat           varchar(50)=null,
        sCovType = LossReportEdit.XML01Loss.AssignmentDetailRS.getField(1, "Type")
        sTemp = oXML01Edit.GetCoverageItemByCoverage(sCovType, "Limits")
        sTemp = Val(goUtil.utCleanValString(sTemp))
        sSQL = sSQL & "'" & Replace(sTemp, "'", "''") & "', " '@UnitLimit
        If sTemp = "0" Then
            sTemp = oXML01Edit.GetCoverageItemByCoverage(sCovType, "Limits")
        Else
            sTemp = vbNullString
        End If
        If sTemp = vbNullString Then
            sTemp = " "
        End If
        sSQL = sSQL & "'" & Replace(sTemp, "'", "''") & "', " '@UnitLimitOther
        If StrComp(sCovType, "Building", vbTextCompare) = 0 Then
            sTemp = "A"
        ElseIf StrComp(sCovType, "Contents", vbTextCompare) = 0 Then
            sTemp = "C"
        ElseIf StrComp(sCovType, "ALE", vbTextCompare) = 0 Then
            sTemp = "D"
        Else
            sTemp = "OTHER"
        End If
        sSQL = sSQL & "'" & Replace(sTemp, "'", "''") & "', " '@UnitClass
        sTemp = LossReportEdit.XML01Loss.AssignmentDetailRS.getField(1, "CauseOfLoss")
        sSQL = sSQL & "'" & Replace(sTemp, "'", "''") & "', " '@UnitCauseOfLoss
        sSQL = sSQL & "'" & sUpdUnitId & "' " '@UnitID             varchar(20)=null,
    End If
    
    'Use Disconnected Recordset
    Set adoRS = New ADODB.Recordset
    adoRS.CursorLocation = adUseClient
    adoRS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set adoRS.ActiveConnection = Nothing
    
    sLRData = adoRS.Fields("LRData")
    sAssignmentsID = adoRS.Fields("AssignmentsID")
    
    Set oXML01 = New clsLossXML01
    
    oXML01.clsCarLR_PrnData = sLRData
    oXML01.clsCarLR_PopulateMe
   
    'Populate the old Header
    sOldHeader = left(sLRData, InStr(1, sLRData, "<header", vbTextCompare) - 1)
    
    'Now that both the Edit LossReport and Existing Loss Report are populated...
    'Need to loop through the Edit Object and Update the Exisiting.
    
    If oXML01Edit.clsCarLR_LossType <> oXML01.clsCarLR_LossType Then
        Err.Raise -999, , "LossType does not match! "
    End If
    
    'Process the Update, Get the Merged XML from Auto or Property
    If oXML01Edit.clsCarLR_LossType = TypeXML01.XML01Apd Then
        sLRData = UpdateUnitAutomobileXML(sLRData, oXML01, oXML01Edit, sUpdateUnitReport)
    ElseIf oXML01Edit.clsCarLR_LossType = TypeXML01.XML01Pro Then
        sLRData = UpdateUnitPropertyXML(sLRData, oXML01, oXML01Edit, sUpdateUnitReport)
    End If
    
    LossReport = oXML01.clsCarLR_LossReport
    
    sSQL = "Exec z_spuEditECSFarmersXML01LossReport "
    If pbMergeUnit Then
        sSQL = sSQL & "1, "     '@bMergeUnit         bit=0,
        sSQL = sSQL & "null, "  '@CLIENTNUM          varchar(20)=null,
        sSQL = sSQL & "null, "  '@ACIDAssignedTo         varchar(20)=null,
        sSQL = sSQL & "null, "  '@LRFormat           varchar(50)=null,
        sSQL = sSQL & "null, "  '@UnitLimit          varchar(20)='0',
        sSQL = sSQL & "null, "  '@UnitLimitOther         varchar(100)=' ',
        If StrComp(sCovType, "Building", vbTextCompare) = 0 Then
            sTemp = "A"
        ElseIf StrComp(sCovType, "Contents", vbTextCompare) = 0 Then
            sTemp = "C"
        ElseIf StrComp(sCovType, "ALE", vbTextCompare) = 0 Then
            sTemp = "D"
        Else
            sTemp = "OTHER"
        End If
        sSQL = sSQL & "'" & Replace(sTemp, "'", "''") & "', " '@UnitClass
        sSQL = sSQL & "null, "  '@UnitCauseOfLoss        varchar(50)='UNKNOWN',
        sSQL = sSQL & "null, "  '@UnitID             varchar(20)=null,
    Else
        sSQL = sSQL & "0, "     '@bMergeUnit         bit=0,
        sSQL = sSQL & "null, "  '@CLIENTNUM          varchar(20)=null,
        sSQL = sSQL & "null, "  '@ACIDAssignedTo         varchar(20)=null,
        sSQL = sSQL & "null, "  '@LRFormat           varchar(50)=null,
        sSQL = sSQL & "null, "  '@UnitLimit          varchar(20)='0',
        sSQL = sSQL & "null, "  '@UnitLimitOther         varchar(100)=' ',
        sSQL = sSQL & "null, "  '@UnitClass          varchar(1)='OTHER',
        sSQL = sSQL & "null, "  '@UnitCauseOfLoss        varchar(50)='UNKNOWN',
        sSQL = sSQL & "null, "  '@UnitID             varchar(20)=null,
    End If
    sSQL = sSQL & sAssignmentsID & ", " '@AssignmentsID          int=null,
    sSQL = sSQL & "'" & Replace(sLRData, "'", "''") & "', "  '@LossReportPrnData      ntext=null
    'Check for Required Updates
    '--  For Mailing Address Update Synchro
    If mbMaillingAddressUpdateRequired Then
        ContactDetail.ContactRowID = LossReport.XML01Loss.ContactDetailRS.getField(1, "ContactRowID")
        Contact.PrimaryAddressID = oXML01.GetContactItem(ContactDetail.ContactRowID, "PrimaryAddressID")
        sTemp = oXML01.GetAddressItem(Contact.PrimaryAddressID, "StreetAddress")
        sTemp = sTemp & " " & oXML01.GetAddressItem(Contact.PrimaryAddressID, "StreetAddress2")
        sSQL = sSQL & "'" & Replace(sTemp, "'", "''") & "', "  '@MailAddress1           varchar(250)=null,
        
        sTemp = oXML01.GetAddressItem(Contact.PrimaryAddressID, "City") & ", "
        sTemp = sTemp & oXML01.GetAddressItem(Contact.PrimaryAddressID, "State") & " "
        sTemp = sTemp & oXML01.GetAddressItem(Contact.PrimaryAddressID, "PostalCode")
        sSQL = sSQL & "'" & Replace(sTemp, "'", "''") & "', "  '@MailAddress2           varchar(250)=null,
    Else
        sSQL = sSQL & "null, "
        sSQL = sSQL & "null, "
    End If
    '--  For Property Address Update Synchro
    If mbPropertyAddressUpdateRequired Then
        If oXML01.clsCarLR_LossType = TypeXML01.XML01Apd Then
            sTemp = LossReport.XML01Loss.VehicleDetailRS.getField(1, "LocationAddress")
            sSQL = sSQL & "'" & Replace(sTemp, "'", "''") & "', "  '@PropAddress1           varchar(250)=null,
            
            sTemp = LossReport.XML01Loss.VehicleDetailRS.getField(1, "LocationCity") & ", "
            sTemp = sTemp & LossReport.XML01Loss.VehicleDetailRS.getField(1, "LocationState") & " "
            sTemp = sTemp & LossReport.XML01Loss.VehicleDetailRS.getField(1, "LocationPostalCode")
            
            sSQL = sSQL & "'" & Replace(sTemp, "'", "''") & "', "  '@PropAddress2           varchar(250)=null,
        ElseIf oXML01.clsCarLR_LossType = TypeXML01.XML01Pro Then
            sTemp = LossReport.XML01Loss.LossDetailRS.getField(1, "LossLocationAddress")
            sTemp2 = LossReport.XML01Loss.LossDetailRS.getField(1, "LossLocationAddress2")
            If Trim(sTemp2) <> vbNullString Then
                sTemp = sTemp & " " & sTemp2
            End If
            sSQL = sSQL & "'" & Replace(sTemp, "'", "''") & "', "  '@PropAddress1           varchar(250)=null,
            
            sTemp = LossReport.XML01Loss.LossDetailRS.getField(1, "LossLocationCity") & ", "
            sTemp = sTemp & LossReport.XML01Loss.LossDetailRS.getField(1, "LossLocationState") & " "
            sTemp = sTemp & LossReport.XML01Loss.LossDetailRS.getField(1, "LossLocationZip")
            sSQL = sSQL & "'" & Replace(sTemp, "'", "''") & "', "  '@PropAddress2           varchar(250)=null,
        End If
    Else
        sSQL = sSQL & "null, "
        sSQL = sSQL & "null, "
    End If
    '--  Other Assignments fields Update synchro
    If mbPolicyNoUpdateRequired Then
        sTemp = LossReport.XML01Loss.PolicyDetailRS.getField(1, "PolicyNumber")
        sSQL = sSQL & "'" & Replace(sTemp, "'", "''") & "', "  '@PolicyNo           varchar(50)=null,
    Else
        sSQL = sSQL & "null, "
    End If
    
    If mbPolicyDescriptionUpdateRequired Then
        sTemp = LossReport.XML01Loss.PolicyDetailRS.getField(1, "PolicyDescription")
        sTemp = sTemp & " " & LossReport.XML01Loss.PolicyDetailRS.getField(1, "PolicyEdition")
        sSQL = sSQL & "'" & Replace(sTemp, "'", "''") & "', "  '@PolicyDescription      varchar(100)=null,
    Else
        sSQL = sSQL & "null, "
    End If
    
    If mbInsuredUpdateRequired Then
        sTemp = LossReport.XML01Loss.AssignmentDetailRS.getField(1, "FirstName")
        sTemp = sTemp & " " & LossReport.XML01Loss.AssignmentDetailRS.getField(1, "LastName")
        sSQL = sSQL & "'" & Replace(sTemp, "'", "''") & "', "  '@Insured            varchar(100)=null,
    Else
        sSQL = sSQL & "null, "
    End If
    ContactDetail.ContactRowID = LossReport.XML01Loss.ContactDetailRS.getField(1, "ContactRowID")
    ContactDetail.UniqueID = LossReport.XML01Loss.ContactDetailRS.getField(1, "ContactRowID")
    If mbHomePhoneUpdateRequired Then
        sTemp = oXML01.GetContactItem(ContactDetail.ContactRowID, "HomePhoneNumber")
        sSQL = sSQL & "'" & Replace(sTemp, "'", "''") & "', "  '@HomePhone          varchar(50)=null,
    Else
        sSQL = sSQL & "null, "
    End If
    If mbBusinessPhoneUpdateRequired Then
        sTemp = oXML01.GetContactItem(ContactDetail.ContactRowID, "WorkPhoneNumber")
        sSQL = sSQL & "'" & Replace(sTemp, "'", "''") & "', "  '@BusinessPhone          varchar(50)=null,
    Else
        sSQL = sSQL & "null, "
    End If
    If mbMortgageeNameUpdateRequired Then
        sTemp = LossReport.XML01Loss.PolicyDetailRS.getField(1, "MortgageeName")
        sSQL = sSQL & "'" & Replace(sTemp, "'", "''") & "', "  '@MortgageeName          varchar(100)=null,
    Else
        sSQL = sSQL & "null, "
    End If
    If mbAgentNoUpdateRequired Then
        sTemp = LossReport.XML01Loss.ContactDetailRS.getField(1, "AgentPrimaryPhone")
        sSQL = sSQL & "'" & Replace(sTemp, "'", "''") & "', "  '@AgentNo            varchar(50)=null,
    Else
        sSQL = sSQL & "null, "
    End If
    If mbDeductibleUpdateRequired Then
        If oXML01.clsCarLR_LossType = TypeXML01.XML01Apd Then
            sCovType = LossReport.XML01Loss.AssignmentDetailRS.getField(1, "Type")
            If StrComp(sCovType, "Comp", vbTextCompare) = 0 Then
                sCovType = "Comprehensive"
            End If
            sTemp = LTrim(Right(oXML01.GetCoverageItemByCoverage(sCovType, "Deductible1"), 10))
            sTemp = Val(goUtil.utCleanValString(sTemp))
            sSQL = sSQL & sTemp & ", "  '@Deductible             Money=null,
        ElseIf oXML01.clsCarLR_LossType = TypeXML01.XML01Pro Then
            sCovType = "Building"
            sTemp = LTrim(Right(oXML01.GetCoverageItemByCoverage(sCovType, "Deductible1"), 10))
            sTemp = Val(goUtil.utCleanValString(sTemp))
            sSQL = sSQL & sTemp & ", "  '@Deductible             Money=null,
        End If
    Else
        sSQL = sSQL & "null, "
    End If
    '--Flags For Vendor update Synchro
    If mbFarmersXML01_LocationUpdateQueUpdateRequired And oXML01.clsCarLR_LossType = TypeXML01.XML01Apd Then
        sSQL = sSQL & "'" & Replace(sUpdUnitId, "'", "''") & "', "  '@UpdUnitId          varchar(20)=null,
        sSQL = sSQL & "'" & Replace(sUpdClaimId, "'", "''") & "', "  '@UpdClaimId         varchar(20)=null,
        sTemp = LossReport.XML01Loss.VehicleDetailRS.getField(1, "ContactRowId")
        sSQL = sSQL & "'" & Replace(sTemp, "'", "''") & "', "  '@ContactRowId           varchar(20)=null,
        sTemp = LossReport.XML01Loss.VehicleDetailRS.getField(1, "UniqueID")
        sSQL = sSQL & "'" & Replace(sTemp, "'", "''") & "', "  '@PropertyRowId          varchar(20)=null,
        sTemp = LossReport.XML01Loss.VehicleDetailRS.getField(1, "PropertyDescription")
        sSQL = sSQL & "'" & Replace(sTemp, "'", "''") & "', "  '@PropertyDescription        varchar(50)='',
        sTemp = LossReport.XML01Loss.VehicleDetailRS.getField(1, "VehicleMake")
        sSQL = sSQL & "'" & Replace(sTemp, "'", "''") & "', "  '@VehicleMake            varchar(30)='',
        sTemp = LossReport.XML01Loss.VehicleDetailRS.getField(1, "VehicleModel")
        sSQL = sSQL & "'" & Replace(sTemp, "'", "''") & "', "  '@VehicleModel           varchar(30)='',
        sTemp = LossReport.XML01Loss.VehicleDetailRS.getField(1, "VehicleYear")
        sSQL = sSQL & "'" & Replace(sTemp, "'", "''") & "', "  '@VehicleYear            varchar(4)='',
        sTemp = LossReport.XML01Loss.VehicleDetailRS.getField(1, "InspectedVIN")
        sSQL = sSQL & "'" & Replace(sTemp, "'", "''") & "', "  '@InspectedVIN           varchar(30)='',
        sTemp = LossReport.XML01Loss.VehicleDetailRS.getField(1, "LocationType")
        sSQL = sSQL & "'" & Replace(sTemp, "'", "''") & "', "  '@LocationType           varchar(50)='',
        sTemp = LossReport.XML01Loss.VehicleDetailRS.getField(1, "LocationName")
        sSQL = sSQL & "'" & Replace(sTemp, "'", "''") & "', "  '@LocationName           varchar(50)='',
        sTemp = LossReport.XML01Loss.VehicleDetailRS.getField(1, "LocationAddress")
        sSQL = sSQL & "'" & Replace(sTemp, "'", "''") & "', "  '@LocationAddress        varchar(200)='',
        sTemp = LossReport.XML01Loss.VehicleDetailRS.getField(1, "LocationCity")
        sSQL = sSQL & "'" & Replace(sTemp, "'", "''") & "', "  '@LocationCity           varchar(50)='',
        sTemp = LossReport.XML01Loss.VehicleDetailRS.getField(1, "LocationCountry")
        sSQL = sSQL & "'" & Replace(sTemp, "'", "''") & "', "  '@LocationCountry        varchar(50)='',
        sTemp = LossReport.XML01Loss.VehicleDetailRS.getField(1, "LocationState")
        sSQL = sSQL & "'" & Replace(sTemp, "'", "''") & "', "  '@LocationState          varchar(50)='',
        sTemp = LossReport.XML01Loss.VehicleDetailRS.getField(1, "LocationPostalCode")
        sSQL = sSQL & "'" & Replace(sTemp, "'", "''") & "', "  '@LocationPostalCode     varchar(30)='',
        sTemp = LossReport.XML01Loss.VehicleDetailRS.getField(1, "LocationPhoneNumber")
        sSQL = sSQL & "'" & Replace(sTemp, "'", "''") & "', "  '@LocationPhoneNumber        varchar(50)='',
        sTemp = LossReport.XML01Loss.VehicleDetailRS.getField(1, "GMTUpdated")
        sSQL = sSQL & "'" & sTemp & "', "  '@GMTUpdated         datetime=NULL
    Else
        sSQL = sSQL & "null, "  '@UpdUnitId               varchar(20)=null,
        sSQL = sSQL & "null, "   '@UpdClaimId             varchar(20)=null,
        sSQL = sSQL & "null, "   '@ContactRowId           varchar(20)=null,
        sSQL = sSQL & "null, "   '@PropertyRowId          varchar(20)=null,
        sSQL = sSQL & "null, "   '@PropertyDescription    varchar(50)='',
        sSQL = sSQL & "null, "   '@VehicleMake            varchar(30)='',
        sSQL = sSQL & "null, "   '@VehicleModel           varchar(30)='',
        sSQL = sSQL & "null, "   '@VehicleYear            varchar(4)='',
        sSQL = sSQL & "null, "   '@InspectedVIN           varchar(30)='',
        sSQL = sSQL & "null, "   '@LocationType           varchar(50)='',
        sSQL = sSQL & "null, "   '@LocationName           varchar(50)='',
        sSQL = sSQL & "null, "   '@LocationAddress        varchar(200)='',
        sSQL = sSQL & "null, "   '@LocationCity           varchar(50)='',
        sSQL = sSQL & "null, "   '@LocationCountry        varchar(50)='',
        sSQL = sSQL & "null, "   '@LocationState          varchar(50)='',
        sSQL = sSQL & "null, "   '@LocationPostalCode     varchar(30)='',
        sSQL = sSQL & "null, "   '@LocationPhoneNumber    varchar(50)='',
        sSQL = sSQL & "null, "   '@GMTUpdated             datetime=NULL
    End If
    'LAST ITEM NO COMMA!!!
    sTemp = Trim(sUpdateUnitReport)
    If sTemp = vbNullString Then
        sSQL = sSQL & "null "
    Else
        sSQL = sSQL & "'" & Replace(sTemp, "'", "''") & "' " '@UpdateUnitReport
    End If
    
  
    oConn.Execute sSQL
    
    UpdateUnit = True
    
    'Cleanup
    Set adoRS = Nothing
    Set oConn = Nothing
    Set oXML01 = Nothing
    Set oXML01Edit = Nothing
    Set oWddxSer = Nothing
    Set oWddxStruc = Nothing
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Set adoRS = Nothing
    Set oConn = Nothing
    Set oXML01 = Nothing
    Set oXML01Edit = Nothing
    Set oWddxSer = Nothing
    Set oWddxStruc = Nothing
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLR_ClassName & vbCrLf & "Public Function UpdateUnit"
End Function

Public Function UpdateUnitAutomobileXML(psLRData As String, _
                                        poXML01 As clsLossXML01, _
                                        poXML01Edit As clsLossXML01, _
                                        psUpdateUnitReport As String) As String
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
   
    Dim oXML01 As clsLossXML01
    Dim LossReport  As XML01LossReport
    Dim sLRData As String
    Dim oXML01Edit As clsLossXML01
    Dim LossReportEdit As XML01LossReport
    Dim lRSPos As Long
    Dim saryColNames() As Variant
    Dim lColNamesPos As Long
    Dim sOldValue As String
    Dim sNewValue As String
    Dim sUniqueID As String
    Dim sOldUniqueID As String
    Dim bRSIsNothing As Boolean
    Dim saryNewCols() As String
    Dim lNewColsPos As Long
    Dim lAddNewRowPos As Long
    Dim lEditRowPos As Long
    Dim oWddxStruc As WDDXStruct
    Dim oWddxSer As WDDXSerializer
    Dim sOldHeader As String
    Dim ContactDetail As udtXML01ContactDetail
    Dim Contact As udtXML01Contacts
    Dim sTemp As String
    Dim sTemp2 As String
    Dim sUpdateMess As String
    Dim sTickCount As String
   
    sTickCount = goUtil.utGetTickCount
   
    Set oXML01 = poXML01
    Set oXML01Edit = poXML01Edit

    
    'DEBUG *************************************************************************
    
    goUtil.utSaveFileData "c:\temp\" & sTickCount & "_AutoMobile_01BEFOREUPDATE_.xml", psLRData
    'DEBUG***************************************************************************
    
    'Populate the old Header
    sOldHeader = left(psLRData, InStr(1, psLRData, "<header", vbTextCompare) - 1)
    
    'Now that both the Edit LossReport and Existing Loss Report are populated...
    'Need to loop through the Edit Object and Update the Exisiting.
    
    If oXML01Edit.clsCarLR_LossType <> oXML01.clsCarLR_LossType Then
        Err.Raise -999, , "LossType does not match! "
    End If
'******************************************************************************************************
'*
'*                                            AUTOMOBILE
'*
'*
'******************************************************************************************************
    'Now Update the DB with the New Loss Report Data for Automobile
    'Rebuild the Wddx Packet !
    Set oWddxSer = New WDDXSerializer
    Set oWddxStruc = New WDDXStruct
    oWddxStruc.setProp "ClassName", "V2ECcarFarmers.clsLossXML01"
    oWddxStruc.setProp "TransType", "IA_CRN_ASSIGN"
    oWddxStruc.setProp "LossType", "Auto"
    
    LossReport = oXML01.clsCarLR_LossReport
    LossReportEdit = oXML01Edit.clsCarLR_LossReport
    
    
    'BEGIN THE UPDATE REPORT HEADER
    
    sUpdateMess = String(50, "*") & vbCrLf
    sUpdateMess = sUpdateMess & "UNIT UPDATE REPORT" & String(4, vbTab) & Now() & vbCrLf
    sUpdateMess = sUpdateMess & "Initial Unit Number: " & String(25 - Len("Initial Unit Numer: "), Chr(32)) & String(1, vbTab) & LossReport.XML01Loss.AssignmentDetailRS.getField(1, "UnitNumber") & vbCrLf
    sUpdateMess = sUpdateMess & "Claim Number: " & String(25 - Len("Claim Number: "), Chr(32)) & String(1, vbTab) & LossReport.XML01Loss.AssignmentDetailRS.getField(1, "ClaimNumber") & vbCrLf
    sUpdateMess = sUpdateMess & "ClaimId: " & String(25 - Len("ClaimId: "), Chr(32)) & String(1, vbTab) & oXML01Edit.ClaimId & vbCrLf
    sUpdateMess = sUpdateMess & "UnitId: " & String(25 - Len("UnitId: "), Chr(32)) & String(1, vbTab) & oXML01Edit.UniqueID & vbCrLf
    sUpdateMess = sUpdateMess & String(50, "*") & vbCrLf
    
    '********************************AssignmentDetailRS**********************************************
    If Not LossReportEdit.XML01Loss.AssignmentDetailRS Is Nothing Then
        sUpdateMess = BuildUpdateMessBeginRS(sUpdateMess, "Assignment Detail")
        bRSIsNothing = False
        lAddNewRowPos = 0
        For lRSPos = 1 To LossReportEdit.XML01Loss.AssignmentDetailRS.getRowCount
            sUniqueID = LossReportEdit.XML01Loss.AssignmentDetailRS.getField(lRSPos, "UniqueID")
            saryColNames() = LossReportEdit.XML01Loss.AssignmentDetailRS.getColumnNames
            For lColNamesPos = LBound(saryColNames, 1) To UBound(saryColNames, 1)
                sOldUniqueID = oXML01.GetAssignmentDetailItem(sUniqueID, "UniqueID", lEditRowPos)
                If sOldUniqueID <> vbNullString Then
                    sOldValue = oXML01.GetAssignmentDetailItem(sUniqueID, CStr(saryColNames(lColNamesPos)))
                Else
                    sOldValue = vbNullString
                    lAddNewRowPos = 0
                End If
                sNewValue = oXML01Edit.GetAssignmentDetailItem(sUniqueID, CStr(saryColNames(lColNamesPos)))
                If LossReport.XML01Loss.AssignmentDetailRS Is Nothing Then
                    bRSIsNothing = True
                    Set LossReport.XML01Loss.AssignmentDetailRS = New WDDXRecordset
                    saryNewCols() = Split("UniqueID,UnitNumber,ClaimNumber,DateAssigned,CatastropheCode,LossDate,Type,CauseOfLoss,FirstName,LastName,AssignedTo,AssignedToFirstName,AssignedToLastName", ",")
                    For lNewColsPos = LBound(saryNewCols, 1) To UBound(saryNewCols, 1)
                        LossReport.XML01Loss.AssignmentDetailRS.addColumn saryNewCols(lNewColsPos)
                    Next
                    lAddNewRowPos = LossReportEdit.XML01Loss.AssignmentDetailRS.getRowCount
                    LossReport.XML01Loss.AssignmentDetailRS.addRows lAddNewRowPos
                    sUpdateMess = sUpdateMess & "Added Record" & vbCrLf
                End If
                If bRSIsNothing Then
                    LossReport.XML01Loss.AssignmentDetailRS.setField lRSPos, CStr(saryColNames(lColNamesPos)), sNewValue
                Else
                    If sOldUniqueID = vbNullString And lAddNewRowPos = 0 Then
                        LossReport.XML01Loss.AssignmentDetailRS.addRows 1
                        lAddNewRowPos = LossReport.XML01Loss.AssignmentDetailRS.getRowCount
                        LossReport.XML01Loss.AssignmentDetailRS.setField lAddNewRowPos, CStr(saryColNames(lColNamesPos)), sNewValue
                        sUpdateMess = sUpdateMess & "Added Record" & vbCrLf
                    ElseIf lAddNewRowPos > 0 Then
                        LossReport.XML01Loss.AssignmentDetailRS.setField lAddNewRowPos, CStr(saryColNames(lColNamesPos)), sNewValue
                    Else
                        LossReport.XML01Loss.AssignmentDetailRS.setField lEditRowPos, CStr(saryColNames(lColNamesPos)), sNewValue
                    End If
                End If
                'See if an Assignment Field Update For Automobile is Required
                If StrComp(sOldValue, sNewValue, vbBinaryCompare) <> 0 Then
                    sTemp = CStr(saryColNames(lColNamesPos))
                    sUpdateMess = sUpdateMess & BuildUpdateMess(sTemp, sOldValue, sNewValue, CBool(lAddNewRowPos))
                    If Not mbInsuredUpdateRequired Then
                        If StrComp(sTemp, "LastName", vbTextCompare) = 0 Then
                            mbInsuredUpdateRequired = True
                        ElseIf StrComp(sTemp, "FirstName", vbTextCompare) = 0 Then
                            mbInsuredUpdateRequired = True
                        End If
                    End If
                End If
            Next
        Next
        sUpdateMess = BuildUpdateMessEndRS(sUpdateMess)
    End If
    If Not LossReport.XML01Loss.AssignmentDetailRS Is Nothing Then
        oWddxStruc.setProp "AssignmentDetailRS", LossReport.XML01Loss.AssignmentDetailRS
    End If
    '********************************END AssignmentDetailRS**********************************************
    
    '******************************VehicleDetailRS***********************
    If Not LossReportEdit.XML01Loss.VehicleDetailRS Is Nothing Then
        sUpdateMess = BuildUpdateMessBeginRS(sUpdateMess, "Vehicle Detail")
        bRSIsNothing = False
        lAddNewRowPos = 0
        For lRSPos = 1 To LossReportEdit.XML01Loss.VehicleDetailRS.getRowCount
            sUniqueID = LossReportEdit.XML01Loss.VehicleDetailRS.getField(lRSPos, "UniqueID")
            saryColNames() = LossReportEdit.XML01Loss.VehicleDetailRS.getColumnNames
            For lColNamesPos = LBound(saryColNames, 1) To UBound(saryColNames, 1)
                sOldUniqueID = oXML01.GetVehicleDetailItem(sUniqueID, "UniqueID", lEditRowPos)
                If sOldUniqueID <> vbNullString Then
                    sOldValue = oXML01.GetVehicleDetailItem(sUniqueID, CStr(saryColNames(lColNamesPos)))
                Else
                    sOldValue = vbNullString
                    lAddNewRowPos = 0
                End If
                sNewValue = oXML01Edit.GetVehicleDetailItem(sUniqueID, CStr(saryColNames(lColNamesPos)))
                If LossReport.XML01Loss.VehicleDetailRS Is Nothing Then
                    bRSIsNothing = True
                    Set LossReport.XML01Loss.VehicleDetailRS = New WDDXRecordset
                    saryNewCols() = Split("UniqueID,ContactRowId,PropertyDescription,VehicleMake,VehicleModel,VehicleYear,InspectedVIN,VIN,PropertyItemName,DamageDescription,LossDescription,LocationType,LocationName,LocationPhoneNumber,LocationAddress,LocationCity,LocationCountry,LocationState,LocationPostalCode,GMTUpdated", ",")
                    For lNewColsPos = LBound(saryNewCols, 1) To UBound(saryNewCols, 1)
                        LossReport.XML01Loss.VehicleDetailRS.addColumn saryNewCols(lNewColsPos)
                    Next
                    lAddNewRowPos = LossReportEdit.XML01Loss.VehicleDetailRS.getRowCount
                    LossReport.XML01Loss.VehicleDetailRS.addRows lAddNewRowPos
                    sUpdateMess = sUpdateMess & "Added Record" & vbCrLf
                End If
                If bRSIsNothing Then
                    LossReport.XML01Loss.VehicleDetailRS.setField lRSPos, CStr(saryColNames(lColNamesPos)), sNewValue
                Else
                    If sOldUniqueID = vbNullString And lAddNewRowPos = 0 Then
                        LossReport.XML01Loss.VehicleDetailRS.addRows 1
                        lAddNewRowPos = LossReport.XML01Loss.VehicleDetailRS.getRowCount
                        LossReport.XML01Loss.VehicleDetailRS.setField lAddNewRowPos, CStr(saryColNames(lColNamesPos)), sNewValue
                        sUpdateMess = sUpdateMess & "Added Record" & vbCrLf
                    ElseIf lAddNewRowPos > 0 Then
                        LossReport.XML01Loss.VehicleDetailRS.setField lAddNewRowPos, CStr(saryColNames(lColNamesPos)), sNewValue
                    Else
                        LossReport.XML01Loss.VehicleDetailRS.setField lEditRowPos, CStr(saryColNames(lColNamesPos)), sNewValue
                    End If
                End If
                If StrComp(sOldValue, sNewValue, vbBinaryCompare) <> 0 Then
                    sTemp = CStr(saryColNames(lColNamesPos))
                    sUpdateMess = sUpdateMess & BuildUpdateMess(sTemp, sOldValue, sNewValue, CBool(lAddNewRowPos))
                    'See if a Property Address Update For Automobile is Required
                    If Not mbPropertyAddressUpdateRequired Then
                        If StrComp(sTemp, "LocationAddress", vbTextCompare) = 0 Then
                            mbPropertyAddressUpdateRequired = True
                            'this also means Vendor update synchro is Required
                            mbFarmersXML01_LocationUpdateQueUpdateRequired = True
                        ElseIf StrComp(sTemp, "LocationCity", vbTextCompare) = 0 Then
                            mbPropertyAddressUpdateRequired = True
                            'this also means Vendor update synchro is Required
                            mbFarmersXML01_LocationUpdateQueUpdateRequired = True
                        ElseIf StrComp(sTemp, "LocationState", vbTextCompare) = 0 Then
                            mbPropertyAddressUpdateRequired = True
                            'this also means Vendor update synchro is Required
                            mbFarmersXML01_LocationUpdateQueUpdateRequired = True
                        ElseIf StrComp(sTemp, "LocationPostalCode", vbTextCompare) = 0 Then
                            mbPropertyAddressUpdateRequired = True
                            'this also means Vendor update synchro is Required
                            mbFarmersXML01_LocationUpdateQueUpdateRequired = True
                        End If
                    End If
                    If Not mbFarmersXML01_LocationUpdateQueUpdateRequired Then
                        If StrComp(sTemp, "VehicleMake", vbTextCompare) = 0 Then
                            mbFarmersXML01_LocationUpdateQueUpdateRequired = True
                        ElseIf StrComp(sTemp, "UniqueID", vbTextCompare) = 0 Then
                            mbFarmersXML01_LocationUpdateQueUpdateRequired = True
                        ElseIf StrComp(sTemp, "ContactRowID", vbTextCompare) = 0 Then
                            mbFarmersXML01_LocationUpdateQueUpdateRequired = True
                        ElseIf StrComp(sTemp, "VehicleModel", vbTextCompare) = 0 Then
                            mbFarmersXML01_LocationUpdateQueUpdateRequired = True
                        ElseIf StrComp(sTemp, "VehicleYear", vbTextCompare) = 0 Then
                            mbFarmersXML01_LocationUpdateQueUpdateRequired = True
                        ElseIf StrComp(sTemp, "InspectedVIN", vbTextCompare) = 0 Then
                            mbFarmersXML01_LocationUpdateQueUpdateRequired = True
                        ElseIf StrComp(sTemp, "LocationType", vbTextCompare) = 0 Then
                            mbFarmersXML01_LocationUpdateQueUpdateRequired = True
                        ElseIf StrComp(sTemp, "LocationName", vbTextCompare) = 0 Then
                            mbFarmersXML01_LocationUpdateQueUpdateRequired = True
                        ElseIf StrComp(sTemp, "LocationCountry", vbTextCompare) = 0 Then
                            mbFarmersXML01_LocationUpdateQueUpdateRequired = True
                        ElseIf StrComp(sTemp, "LocationPhoneNumber", vbTextCompare) = 0 Then
                            mbFarmersXML01_LocationUpdateQueUpdateRequired = True
                        ElseIf StrComp(sTemp, "GMTUpdated", vbTextCompare) = 0 Then
                            mbFarmersXML01_LocationUpdateQueUpdateRequired = True
                        End If
                    End If
                End If
            Next
        Next
        sUpdateMess = BuildUpdateMessEndRS(sUpdateMess)
    End If
    If Not LossReport.XML01Loss.VehicleDetailRS Is Nothing Then
        oWddxStruc.setProp "VehicleDetailRS", LossReport.XML01Loss.VehicleDetailRS
    End If
    '******************************END VehicleDetailRS***********************
    
    '********************************ContactDetailRS**********************************************
    If Not LossReportEdit.XML01Loss.ContactDetailRS Is Nothing Then
        sUpdateMess = BuildUpdateMessBeginRS(sUpdateMess, "Contact Detail")
        bRSIsNothing = False
        lAddNewRowPos = 0
        For lRSPos = 1 To LossReportEdit.XML01Loss.ContactDetailRS.getRowCount
            sUniqueID = LossReportEdit.XML01Loss.ContactDetailRS.getField(lRSPos, "UniqueID")
            saryColNames() = LossReportEdit.XML01Loss.ContactDetailRS.getColumnNames
            For lColNamesPos = LBound(saryColNames, 1) To UBound(saryColNames, 1)
                sOldUniqueID = oXML01.GetContactDetailsItem(sUniqueID, "UniqueID", lEditRowPos)
                If sOldUniqueID <> vbNullString Then
                    sOldValue = oXML01.GetContactDetailsItem(sUniqueID, CStr(saryColNames(lColNamesPos)))
                Else
                    sOldValue = vbNullString
                    lAddNewRowPos = 0
                End If
                sNewValue = oXML01Edit.GetContactDetailsItem(sUniqueID, CStr(saryColNames(lColNamesPos)))
                If LossReport.XML01Loss.ContactDetailRS Is Nothing Then
                    bRSIsNothing = True
                    Set LossReport.XML01Loss.ContactDetailRS = New WDDXRecordset
                    saryNewCols() = Split("UniqueID,AgentFirstName,AgentLastName,AgentPrimaryPhone,ContactRowID", ",")
                    For lNewColsPos = LBound(saryNewCols, 1) To UBound(saryNewCols, 1)
                        LossReport.XML01Loss.ContactDetailRS.addColumn saryNewCols(lNewColsPos)
                    Next
                    lAddNewRowPos = LossReportEdit.XML01Loss.ContactDetailRS.getRowCount
                    LossReport.XML01Loss.ContactDetailRS.addRows lAddNewRowPos
                    sUpdateMess = sUpdateMess & "Added Record" & vbCrLf
                End If
                If bRSIsNothing Then
                    LossReport.XML01Loss.ContactDetailRS.setField lRSPos, CStr(saryColNames(lColNamesPos)), sNewValue
                Else
                    If sOldUniqueID = vbNullString And lAddNewRowPos = 0 Then
                        LossReport.XML01Loss.ContactDetailRS.addRows 1
                        lAddNewRowPos = LossReport.XML01Loss.ContactDetailRS.getRowCount
                        LossReport.XML01Loss.ContactDetailRS.setField lAddNewRowPos, CStr(saryColNames(lColNamesPos)), sNewValue
                        sUpdateMess = sUpdateMess & "Added Record" & vbCrLf
                    ElseIf lAddNewRowPos > 0 Then
                        LossReport.XML01Loss.ContactDetailRS.setField lAddNewRowPos, CStr(saryColNames(lColNamesPos)), sNewValue
                    Else
                        LossReport.XML01Loss.ContactDetailRS.setField lEditRowPos, CStr(saryColNames(lColNamesPos)), sNewValue
                    End If
                End If
                'See if an Assignment Field Update For Automobile is Required
                If StrComp(sOldValue, sNewValue, vbBinaryCompare) <> 0 Then
                    sTemp = CStr(saryColNames(lColNamesPos))
                    sUpdateMess = sUpdateMess & BuildUpdateMess(sTemp, sOldValue, sNewValue, CBool(lAddNewRowPos))
                    If Not mbAgentNoUpdateRequired Then
                        If StrComp(sTemp, "AgentPrimaryPhone", vbTextCompare) = 0 Then
                            mbAgentNoUpdateRequired = True
                        End If
                    End If
                End If
            Next
        Next
        sUpdateMess = BuildUpdateMessEndRS(sUpdateMess)
    End If
    If Not LossReport.XML01Loss.ContactDetailRS Is Nothing Then
        oWddxStruc.setProp "ContactDetailsRS", LossReport.XML01Loss.ContactDetailRS
        'Set the current contactID
        ContactDetail.ContactRowID = LossReport.XML01Loss.ContactDetailRS.getField(1, "ContactRowID")
    End If
    
    '********************************END ContactDetailRS**********************************************
    
    '********************************ContactsRS**********************************************
    If Not LossReportEdit.XML01Loss.ContactsRS Is Nothing Then
        sUpdateMess = BuildUpdateMessBeginRS(sUpdateMess, "Contacts")
        bRSIsNothing = False
        lAddNewRowPos = 0
        For lRSPos = 1 To LossReportEdit.XML01Loss.ContactsRS.getRowCount
            sUniqueID = LossReportEdit.XML01Loss.ContactsRS.getField(lRSPos, "ContactRowID")
            saryColNames() = LossReportEdit.XML01Loss.ContactsRS.getColumnNames
            For lColNamesPos = LBound(saryColNames, 1) To UBound(saryColNames, 1)
                sOldUniqueID = oXML01.GetContactItem(sUniqueID, "ContactRowID", lEditRowPos)
                If sOldUniqueID <> vbNullString Then
                    sOldValue = oXML01.GetContactItem(sUniqueID, CStr(saryColNames(lColNamesPos)))
                Else
                    sOldValue = vbNullString
                    lAddNewRowPos = 0
                End If
                sNewValue = oXML01Edit.GetContactItem(sUniqueID, CStr(saryColNames(lColNamesPos)))
                If LossReport.XML01Loss.ContactsRS Is Nothing Then
                    bRSIsNothing = True
                    Set LossReport.XML01Loss.ContactsRS = New WDDXRecordset
                    saryNewCols() = Split("ContactRowID,MM,FirstName,LastName,ContactRole,BestPhoneNumber,PrimaryPhoneNumber,HomePhoneNumber,CellularPhoneNumber,FaxPhoneNumber,EmailAddress,PagerPhoneNumber,PagerPIN,WorkPhoneNumber,RelationshipToInsured,PrimaryAddressID,CreatedUpdated", ",")
                    For lNewColsPos = LBound(saryNewCols, 1) To UBound(saryNewCols, 1)
                        LossReport.XML01Loss.ContactsRS.addColumn saryNewCols(lNewColsPos)
                    Next
                    lAddNewRowPos = LossReportEdit.XML01Loss.ContactsRS.getRowCount
                    LossReport.XML01Loss.ContactsRS.addRows lAddNewRowPos
                    sUpdateMess = sUpdateMess & "Added Record" & vbCrLf
                End If
                If bRSIsNothing Then
                    LossReport.XML01Loss.ContactsRS.setField lRSPos, CStr(saryColNames(lColNamesPos)), sNewValue
                Else
                    If sOldUniqueID = vbNullString And lAddNewRowPos = 0 Then
                        LossReport.XML01Loss.ContactsRS.addRows 1
                        lAddNewRowPos = LossReport.XML01Loss.ContactsRS.getRowCount
                        LossReport.XML01Loss.ContactsRS.setField lAddNewRowPos, CStr(saryColNames(lColNamesPos)), sNewValue
                        sUpdateMess = sUpdateMess & "Added Record" & vbCrLf
                    ElseIf lAddNewRowPos > 0 Then
                        LossReport.XML01Loss.ContactsRS.setField lAddNewRowPos, CStr(saryColNames(lColNamesPos)), sNewValue
                    Else
                        LossReport.XML01Loss.ContactsRS.setField lEditRowPos, CStr(saryColNames(lColNamesPos)), sNewValue
                    End If
                End If
                'See if an Assignment Field Update For Automobile is Required
                If StrComp(sOldValue, sNewValue, vbBinaryCompare) <> 0 And StrComp(ContactDetail.ContactRowID, sUniqueID, vbTextCompare) = 0 Then
                    sTemp = CStr(saryColNames(lColNamesPos))
                    sUpdateMess = sUpdateMess & BuildUpdateMess(sTemp, sOldValue, sNewValue, CBool(lAddNewRowPos))
                    If Not mbHomePhoneUpdateRequired Then
                        If StrComp(sTemp, "HomePhoneNumber", vbTextCompare) = 0 Then
                            mbHomePhoneUpdateRequired = True
                        End If
                    End If
                    If Not mbBusinessPhoneUpdateRequired Then
                        If StrComp(sTemp, "WorkPhoneNumber", vbTextCompare) = 0 Then
                            mbBusinessPhoneUpdateRequired = True
                        End If
                    End If
                ElseIf StrComp(sOldValue, sNewValue, vbBinaryCompare) <> 0 Then
                    sTemp = CStr(saryColNames(lColNamesPos))
                    sUpdateMess = sUpdateMess & BuildUpdateMess(sTemp, sOldValue, sNewValue, CBool(lAddNewRowPos))
                End If
            Next
        Next
        sUpdateMess = BuildUpdateMessEndRS(sUpdateMess)
    End If
    If Not LossReport.XML01Loss.ContactsRS Is Nothing Then
        oWddxStruc.setProp "ContactRS", LossReport.XML01Loss.ContactsRS
    End If
    '********************************END ContactsRS**********************************************
    
    '****************************AddressRS*****************************
    If Not LossReportEdit.XML01Loss.AddressRS Is Nothing Then
        sUpdateMess = BuildUpdateMessBeginRS(sUpdateMess, "Address")
        bRSIsNothing = False
        lAddNewRowPos = 0
        For lRSPos = 1 To LossReportEdit.XML01Loss.AddressRS.getRowCount
            sUniqueID = LossReportEdit.XML01Loss.AddressRS.getField(lRSPos, "AddressID")
            saryColNames() = LossReportEdit.XML01Loss.AddressRS.getColumnNames
            For lColNamesPos = LBound(saryColNames, 1) To UBound(saryColNames, 1)
                sOldUniqueID = oXML01.GetAddressItem(sUniqueID, "AddressID", lEditRowPos)
                If sOldUniqueID <> vbNullString Then
                    sOldValue = oXML01.GetAddressItem(sUniqueID, CStr(saryColNames(lColNamesPos)))
                Else
                    sOldValue = vbNullString
                    lAddNewRowPos = 0
                End If
                sNewValue = oXML01Edit.GetAddressItem(sUniqueID, CStr(saryColNames(lColNamesPos)))
                If LossReport.XML01Loss.AddressRS Is Nothing Then
                    bRSIsNothing = True
                    Set LossReport.XML01Loss.AddressRS = New WDDXRecordset
                    saryNewCols() = Split("ContactRowID,AddressID,Type,StreetAddress,StreetAddress2,City,State,PostalCode", ",")
                    For lNewColsPos = LBound(saryNewCols, 1) To UBound(saryNewCols, 1)
                        LossReport.XML01Loss.AddressRS.addColumn saryNewCols(lNewColsPos)
                    Next
                    lAddNewRowPos = LossReportEdit.XML01Loss.AddressRS.getRowCount
                    LossReport.XML01Loss.AddressRS.addRows lAddNewRowPos
                    sUpdateMess = sUpdateMess & "Added Record" & vbCrLf
                End If
                If bRSIsNothing Then
                    LossReport.XML01Loss.AddressRS.setField lRSPos, CStr(saryColNames(lColNamesPos)), sNewValue
                Else
                    If sOldUniqueID = vbNullString And lAddNewRowPos = 0 Then
                        LossReport.XML01Loss.AddressRS.addRows 1
                        lAddNewRowPos = LossReport.XML01Loss.AddressRS.getRowCount
                        LossReport.XML01Loss.AddressRS.setField lAddNewRowPos, CStr(saryColNames(lColNamesPos)), sNewValue
                        sUpdateMess = sUpdateMess & "Added Record" & vbCrLf
                    ElseIf lAddNewRowPos > 0 Then
                        LossReport.XML01Loss.AddressRS.setField lAddNewRowPos, CStr(saryColNames(lColNamesPos)), sNewValue
                    Else
                        LossReport.XML01Loss.AddressRS.setField lEditRowPos, CStr(saryColNames(lColNamesPos)), sNewValue
                    End If
                End If
                'See if a Mailing Address Update For Automobile is Required
                If StrComp(sOldValue, sNewValue, vbTextCompare) <> 0 Then
                    sTemp = CStr(saryColNames(lColNamesPos))
                    sUpdateMess = sUpdateMess & BuildUpdateMess(sTemp, sOldValue, sNewValue, CBool(lAddNewRowPos))
                    If Not mbMaillingAddressUpdateRequired Then
                        sTemp2 = oXML01Edit.GetAddressItem(sUniqueID, "Type")
                        If StrComp(sTemp2, "Mailing", vbTextCompare) = 0 Then
                            mbMaillingAddressUpdateRequired = True
                        End If
                    End If
                End If
            Next
        Next
        sUpdateMess = BuildUpdateMessEndRS(sUpdateMess)
    End If
    If Not LossReport.XML01Loss.AddressRS Is Nothing Then
        oWddxStruc.setProp "AddressRS", LossReport.XML01Loss.AddressRS
    End If
    '******************************END AdressRS***********************
    
    
    '********************************PolicyDetailRS**********************************************
    If Not LossReportEdit.XML01Loss.PolicyDetailRS Is Nothing Then
        sUpdateMess = BuildUpdateMessBeginRS(sUpdateMess, "Policy Detail")
        bRSIsNothing = False
        'Policy Details should only ever have 1 Row !!!
        lRSPos = 1
        'No Unique ID
        saryColNames() = LossReportEdit.XML01Loss.PolicyDetailRS.getColumnNames
        For lColNamesPos = LBound(saryColNames, 1) To UBound(saryColNames, 1)
            If Not LossReport.XML01Loss.PolicyDetailRS Is Nothing Then
                sOldValue = oXML01.GetPolicyDetailItem(CStr(saryColNames(lColNamesPos)))
            Else
                sOldValue = vbNullString
            End If
            sNewValue = oXML01Edit.GetPolicyDetailItem(CStr(saryColNames(lColNamesPos)))
            If LossReport.XML01Loss.PolicyDetailRS Is Nothing Then
                bRSIsNothing = True
                Set LossReport.XML01Loss.PolicyDetailRS = New WDDXRecordset
                saryNewCols() = Split("PolicyNumber,Status,CoverageStatus,BalanceDue,CompanyCode,CompanyName,RenewalDate,CancellationDate,NewBusinessDate,PolicyDescription,PolicyEdition,MortgageeName,MortgageeAddress,LienHolderName,LienHolderAddress", ",")
                For lNewColsPos = LBound(saryNewCols, 1) To UBound(saryNewCols, 1)
                    LossReport.XML01Loss.PolicyDetailRS.addColumn saryNewCols(lNewColsPos)
                Next
                lAddNewRowPos = lRSPos 'LossReportEdit.XML01Loss.PolicyDetailRS.getRowCount
                LossReport.XML01Loss.PolicyDetailRS.addRows lAddNewRowPos
                sUpdateMess = sUpdateMess & "Added Record" & vbCrLf
            End If
            LossReport.XML01Loss.PolicyDetailRS.setField lRSPos, CStr(saryColNames(lColNamesPos)), sNewValue
            'See if an Assignment Field Update For Automobile is Required
            If StrComp(sOldValue, sNewValue, vbBinaryCompare) <> 0 Then
                sTemp = CStr(saryColNames(lColNamesPos))
                sUpdateMess = sUpdateMess & BuildUpdateMess(sTemp, sOldValue, sNewValue, CBool(lAddNewRowPos))
                If Not mbPolicyNoUpdateRequired Then
                    If StrComp(sTemp, "PolicyNumber", vbTextCompare) = 0 Then
                        mbPolicyNoUpdateRequired = True
                    End If
                End If
                If Not mbPolicyDescriptionUpdateRequired Then
                    If StrComp(sTemp, "PolicyDescription", vbTextCompare) = 0 Then
                        mbPolicyDescriptionUpdateRequired = True
                    ElseIf StrComp(sTemp, "PolicyEdition", vbTextCompare) = 0 Then
                        mbPolicyDescriptionUpdateRequired = True
                    End If
                End If
                If Not mbMortgageeNameUpdateRequired Then
                    If StrComp(sTemp, "MortgageeName", vbTextCompare) = 0 Then
                        mbMortgageeNameUpdateRequired = True
                    End If
                End If
            End If
        Next
        sUpdateMess = BuildUpdateMessEndRS(sUpdateMess)
    End If
    If Not LossReport.XML01Loss.PolicyDetailRS Is Nothing Then
        oWddxStruc.setProp "PolicyDetailRS", LossReport.XML01Loss.PolicyDetailRS
    End If
    '********************************END PolicyDetailRS**********************************************
    
    
    '********************************CoverageRS**********************************************
    If Not LossReportEdit.XML01Loss.CoverageRS Is Nothing Then
        sUpdateMess = BuildUpdateMessBeginRS(sUpdateMess, "Coverage")
        bRSIsNothing = False
        lAddNewRowPos = 0
        For lRSPos = 1 To LossReportEdit.XML01Loss.CoverageRS.getRowCount
            sUniqueID = LossReportEdit.XML01Loss.CoverageRS.getField(lRSPos, "CoverageRowID")
            saryColNames() = LossReportEdit.XML01Loss.CoverageRS.getColumnNames
            For lColNamesPos = LBound(saryColNames, 1) To UBound(saryColNames, 1)
                sOldUniqueID = oXML01.GetCoverageItem(sUniqueID, "CoverageRowID", lEditRowPos)
                If sOldUniqueID <> vbNullString Then
                    sOldValue = oXML01.GetCoverageItem(sUniqueID, CStr(saryColNames(lColNamesPos)))
                Else
                    sOldValue = vbNullString
                    lAddNewRowPos = 0
                End If
                sNewValue = oXML01Edit.GetCoverageItem(sUniqueID, CStr(saryColNames(lColNamesPos)))
                If LossReport.XML01Loss.CoverageRS Is Nothing Then
                    bRSIsNothing = True
                    Set LossReport.XML01Loss.CoverageRS = New WDDXRecordset
                    saryNewCols() = Split("CoverageRowID,Coverage,Limits,Deductible1", ",")
                    For lNewColsPos = LBound(saryNewCols, 1) To UBound(saryNewCols, 1)
                        LossReport.XML01Loss.CoverageRS.addColumn saryNewCols(lNewColsPos)
                    Next
                    lAddNewRowPos = LossReportEdit.XML01Loss.CoverageRS.getRowCount
                    LossReport.XML01Loss.CoverageRS.addRows lAddNewRowPos
                    sUpdateMess = sUpdateMess & "Added Record" & vbCrLf
                End If
                If bRSIsNothing Then
                    LossReport.XML01Loss.CoverageRS.setField lRSPos, CStr(saryColNames(lColNamesPos)), sNewValue
                Else
                    If sOldUniqueID = vbNullString And lAddNewRowPos = 0 Then
                        LossReport.XML01Loss.CoverageRS.addRows 1
                        lAddNewRowPos = LossReport.XML01Loss.CoverageRS.getRowCount
                        LossReport.XML01Loss.CoverageRS.setField lAddNewRowPos, CStr(saryColNames(lColNamesPos)), sNewValue
                        sUpdateMess = sUpdateMess & "Added Record" & vbCrLf
                    ElseIf lAddNewRowPos > 0 Then
                        LossReport.XML01Loss.CoverageRS.setField lAddNewRowPos, CStr(saryColNames(lColNamesPos)), sNewValue
                    Else
                        LossReport.XML01Loss.CoverageRS.setField lEditRowPos, CStr(saryColNames(lColNamesPos)), sNewValue
                    End If
                End If
                'See if an Assignment Field Update For Property is Required
                If StrComp(sOldValue, sNewValue, vbBinaryCompare) <> 0 Then
                    sTemp = CStr(saryColNames(lColNamesPos))
                    sUpdateMess = sUpdateMess & BuildUpdateMess(sTemp, sOldValue, sNewValue, CBool(lAddNewRowPos))
                    If Not mbDeductibleUpdateRequired Then
                        If StrComp(sTemp, "Deductible1", vbTextCompare) = 0 Then
                            mbDeductibleUpdateRequired = True
                        End If
                    End If
                End If
            Next
        Next
        sUpdateMess = BuildUpdateMessEndRS(sUpdateMess)
    End If
    If Not LossReport.XML01Loss.CoverageRS Is Nothing Then
        oWddxStruc.setProp "CoverageRS", LossReport.XML01Loss.CoverageRS
    End If
    '********************************END CoverageRS**********************************************
    
    
    '********************************EndorsementRS**********************************************
    If Not LossReportEdit.XML01Loss.EndorsementRS Is Nothing Then
        sUpdateMess = BuildUpdateMessBeginRS(sUpdateMess, "Endorsement")
        bRSIsNothing = False
        lAddNewRowPos = 0
        For lRSPos = 1 To LossReportEdit.XML01Loss.EndorsementRS.getRowCount
            sUniqueID = LossReportEdit.XML01Loss.EndorsementRS.getField(lRSPos, "EndorsementsRowID")
            saryColNames() = LossReportEdit.XML01Loss.EndorsementRS.getColumnNames
            For lColNamesPos = LBound(saryColNames, 1) To UBound(saryColNames, 1)
                sOldUniqueID = oXML01.GetEndorsementItem(sUniqueID, "EndorsementsRowID", lEditRowPos)
                If sOldUniqueID <> vbNullString Then
                    sOldValue = oXML01.GetEndorsementItem(sUniqueID, CStr(saryColNames(lColNamesPos)))
                Else
                    sOldValue = vbNullString
                    lAddNewRowPos = 0
                End If
                sNewValue = oXML01Edit.GetEndorsementItem(sUniqueID, CStr(saryColNames(lColNamesPos)))
                If LossReport.XML01Loss.EndorsementRS Is Nothing Then
                    bRSIsNothing = True
                    Set LossReport.XML01Loss.EndorsementRS = New WDDXRecordset
                    saryNewCols() = Split("EndorsementsRowID,EndorsementNumber,EndorsementDescription", ",")
                    For lNewColsPos = LBound(saryNewCols, 1) To UBound(saryNewCols, 1)
                        LossReport.XML01Loss.EndorsementRS.addColumn saryNewCols(lNewColsPos)
                    Next
                    lAddNewRowPos = LossReportEdit.XML01Loss.EndorsementRS.getRowCount
                    LossReport.XML01Loss.EndorsementRS.addRows lAddNewRowPos
                    sUpdateMess = sUpdateMess & "Added Record" & vbCrLf
                End If
                If bRSIsNothing Then
                    LossReport.XML01Loss.EndorsementRS.setField lRSPos, CStr(saryColNames(lColNamesPos)), sNewValue
                Else
                    If sOldUniqueID = vbNullString And lAddNewRowPos = 0 Then
                        LossReport.XML01Loss.EndorsementRS.addRows 1
                        lAddNewRowPos = LossReport.XML01Loss.EndorsementRS.getRowCount
                        LossReport.XML01Loss.EndorsementRS.setField lAddNewRowPos, CStr(saryColNames(lColNamesPos)), sNewValue
                        sUpdateMess = sUpdateMess & "Added Record" & vbCrLf
                    ElseIf lAddNewRowPos > 0 Then
                        LossReport.XML01Loss.EndorsementRS.setField lAddNewRowPos, CStr(saryColNames(lColNamesPos)), sNewValue
                    Else
                        LossReport.XML01Loss.EndorsementRS.setField lEditRowPos, CStr(saryColNames(lColNamesPos)), sNewValue
                    End If
                End If
                If StrComp(sOldValue, sNewValue, vbBinaryCompare) <> 0 Then
                    sTemp = CStr(saryColNames(lColNamesPos))
                    sUpdateMess = sUpdateMess & BuildUpdateMess(sTemp, sOldValue, sNewValue, CBool(lAddNewRowPos))
                End If
            Next
        Next
        sUpdateMess = BuildUpdateMessEndRS(sUpdateMess)
    End If
    If Not LossReport.XML01Loss.EndorsementRS Is Nothing Then
        oWddxStruc.setProp "EndorsementRS", LossReport.XML01Loss.EndorsementRS
    End If
    '********************************END EndorsementRS**********************************************
    
    '********************************PaymentDetailRS**********************************************
    If Not LossReportEdit.XML01Loss.PaymentDetailRS Is Nothing Then
        sUpdateMess = BuildUpdateMessBeginRS(sUpdateMess, "Payment Detail")
        bRSIsNothing = False
        lAddNewRowPos = 0
        For lRSPos = 1 To LossReportEdit.XML01Loss.PaymentDetailRS.getRowCount
            sUniqueID = LossReportEdit.XML01Loss.PaymentDetailRS.getField(lRSPos, "PaymentDetailRowID")
            saryColNames() = LossReportEdit.XML01Loss.PaymentDetailRS.getColumnNames
            For lColNamesPos = LBound(saryColNames, 1) To UBound(saryColNames, 1)
                sOldUniqueID = oXML01.GetPaymentDetailItem(sUniqueID, "PaymentDetailRowID", lEditRowPos)
                If sOldUniqueID <> vbNullString Then
                    sOldValue = oXML01.GetPaymentDetailItem(sUniqueID, CStr(saryColNames(lColNamesPos)))
                Else
                    sOldValue = vbNullString
                    lAddNewRowPos = 0
                End If
                sNewValue = oXML01Edit.GetPaymentDetailItem(sUniqueID, CStr(saryColNames(lColNamesPos)))
                If LossReport.XML01Loss.PaymentDetailRS Is Nothing Then
                    bRSIsNothing = True
                    Set LossReport.XML01Loss.PaymentDetailRS = New WDDXRecordset
                    saryNewCols() = Split("PaymentDetailRowID,DateIssued,PayeeLineOne,PayeeLineTwo,PayeeLineThree,PayeeLineFour,AccountType,PaymentClass,PaymentAmount", ",")
                    For lNewColsPos = LBound(saryNewCols, 1) To UBound(saryNewCols, 1)
                        LossReport.XML01Loss.PaymentDetailRS.addColumn saryNewCols(lNewColsPos)
                    Next
                    lAddNewRowPos = LossReportEdit.XML01Loss.PaymentDetailRS.getRowCount
                    LossReport.XML01Loss.PaymentDetailRS.addRows lAddNewRowPos
                    sUpdateMess = sUpdateMess & "Added Record" & vbCrLf
                End If
                If bRSIsNothing Then
                    LossReport.XML01Loss.PaymentDetailRS.setField lRSPos, CStr(saryColNames(lColNamesPos)), sNewValue
                Else
                    If sOldUniqueID = vbNullString And lAddNewRowPos = 0 Then
                        LossReport.XML01Loss.PaymentDetailRS.addRows 1
                        lAddNewRowPos = LossReport.XML01Loss.PaymentDetailRS.getRowCount
                        LossReport.XML01Loss.PaymentDetailRS.setField lAddNewRowPos, CStr(saryColNames(lColNamesPos)), sNewValue
                        sUpdateMess = sUpdateMess & "Added Record" & vbCrLf
                    ElseIf lAddNewRowPos > 0 Then
                        LossReport.XML01Loss.PaymentDetailRS.setField lAddNewRowPos, CStr(saryColNames(lColNamesPos)), sNewValue
                    Else
                        LossReport.XML01Loss.PaymentDetailRS.setField lEditRowPos, CStr(saryColNames(lColNamesPos)), sNewValue
                    End If
                End If
                If StrComp(sOldValue, sNewValue, vbBinaryCompare) <> 0 Then
                    sTemp = CStr(saryColNames(lColNamesPos))
                    sUpdateMess = sUpdateMess & BuildUpdateMess(sTemp, sOldValue, sNewValue, CBool(lAddNewRowPos))
                End If
            Next
        Next
        sUpdateMess = BuildUpdateMessEndRS(sUpdateMess)
    End If
    If Not LossReport.XML01Loss.PaymentDetailRS Is Nothing Then
        oWddxStruc.setProp "PaymentDetailRS", LossReport.XML01Loss.PaymentDetailRS
    End If
    '********************************END PaymentDetailRS**********************************************
    
    
    '********************************PriorLossDetailRS**********************************************
    If Not LossReportEdit.XML01Loss.PriorLossDetailRS Is Nothing Then
        sUpdateMess = BuildUpdateMessBeginRS(sUpdateMess, "Prior Loss Detail")
        bRSIsNothing = False
        lAddNewRowPos = 0
        For lRSPos = 1 To LossReportEdit.XML01Loss.PriorLossDetailRS.getRowCount
            sUniqueID = LossReportEdit.XML01Loss.PriorLossDetailRS.getField(lRSPos, "PriorLossRowID")
            saryColNames() = LossReportEdit.XML01Loss.PriorLossDetailRS.getColumnNames
            For lColNamesPos = LBound(saryColNames, 1) To UBound(saryColNames, 1)
                sOldUniqueID = oXML01.GetPriorLossDetailItem(sUniqueID, "PriorLossRowID", lEditRowPos)
                If sOldUniqueID <> vbNullString Then
                    sOldValue = oXML01.GetPriorLossDetailItem(sUniqueID, CStr(saryColNames(lColNamesPos)))
                Else
                    sOldValue = vbNullString
                    lAddNewRowPos = 0
                End If
                sNewValue = oXML01Edit.GetPriorLossDetailItem(sUniqueID, CStr(saryColNames(lColNamesPos)))
                If LossReport.XML01Loss.PriorLossDetailRS Is Nothing Then
                    bRSIsNothing = True
                    Set LossReport.XML01Loss.PriorLossDetailRS = New WDDXRecordset
                    saryNewCols() = Split("PriorLossRowID,SALN,ClaimSegmentNumber,PolicyNumber,LossCause,ClaimClass,LossDate,SummaryAmount", ",")
                    For lNewColsPos = LBound(saryNewCols, 1) To UBound(saryNewCols, 1)
                        LossReport.XML01Loss.PriorLossDetailRS.addColumn saryNewCols(lNewColsPos)
                    Next
                    lAddNewRowPos = LossReportEdit.XML01Loss.PriorLossDetailRS.getRowCount
                    LossReport.XML01Loss.PriorLossDetailRS.addRows lAddNewRowPos
                    sUpdateMess = sUpdateMess & "Added Record" & vbCrLf
                End If
                If bRSIsNothing Then
                    LossReport.XML01Loss.PriorLossDetailRS.setField lRSPos, CStr(saryColNames(lColNamesPos)), sNewValue
                Else
                    If sOldUniqueID = vbNullString And lAddNewRowPos = 0 Then
                        LossReport.XML01Loss.PriorLossDetailRS.addRows 1
                        lAddNewRowPos = LossReport.XML01Loss.PriorLossDetailRS.getRowCount
                        LossReport.XML01Loss.PriorLossDetailRS.setField lAddNewRowPos, CStr(saryColNames(lColNamesPos)), sNewValue
                        sUpdateMess = sUpdateMess & "Added Record" & vbCrLf
                    ElseIf lAddNewRowPos > 0 Then
                        LossReport.XML01Loss.PriorLossDetailRS.setField lAddNewRowPos, CStr(saryColNames(lColNamesPos)), sNewValue
                    Else
                        LossReport.XML01Loss.PriorLossDetailRS.setField lEditRowPos, CStr(saryColNames(lColNamesPos)), sNewValue
                    End If
                End If
                If StrComp(sOldValue, sNewValue, vbBinaryCompare) <> 0 Then
                    sTemp = CStr(saryColNames(lColNamesPos))
                    sUpdateMess = sUpdateMess & BuildUpdateMess(sTemp, sOldValue, sNewValue, CBool(lAddNewRowPos))
                End If
            Next
        Next
        sUpdateMess = BuildUpdateMessEndRS(sUpdateMess)
    End If
    If Not LossReport.XML01Loss.PriorLossDetailRS Is Nothing Then
        oWddxStruc.setProp "PriorLossDetailRS", LossReport.XML01Loss.PriorLossDetailRS
    End If
    '********************************END PriorLossDetailRS**********************************************
    
    
    '********************************ActivitiesRS**********************************************
    If Not LossReportEdit.XML01Loss.ActivitiesRS Is Nothing Then
        sUpdateMess = BuildUpdateMessBeginRS(sUpdateMess, "Activities")
        bRSIsNothing = False
        lAddNewRowPos = 0
        For lRSPos = 1 To LossReportEdit.XML01Loss.ActivitiesRS.getRowCount
            sUniqueID = LossReportEdit.XML01Loss.ActivitiesRS.getField(lRSPos, "ActivityRowID")
            saryColNames() = LossReportEdit.XML01Loss.ActivitiesRS.getColumnNames
            For lColNamesPos = LBound(saryColNames, 1) To UBound(saryColNames, 1)
                sOldUniqueID = oXML01.GetActivitiesItem(sUniqueID, "ActivityRowID", lEditRowPos)
                If sOldUniqueID <> vbNullString Then
                    sOldValue = oXML01.GetActivitiesItem(sUniqueID, CStr(saryColNames(lColNamesPos)))
                Else
                    sOldValue = vbNullString
                    lAddNewRowPos = 0
                End If
                sNewValue = oXML01Edit.GetActivitiesItem(sUniqueID, CStr(saryColNames(lColNamesPos)))
                If LossReport.XML01Loss.ActivitiesRS Is Nothing Then
                    bRSIsNothing = True
                    Set LossReport.XML01Loss.ActivitiesRS = New WDDXRecordset
                    saryNewCols() = Split("ActivityRowID,GMTCreated,CreatedBy,Type,Description,Comment", ",")
                    For lNewColsPos = LBound(saryNewCols, 1) To UBound(saryNewCols, 1)
                        LossReport.XML01Loss.ActivitiesRS.addColumn saryNewCols(lNewColsPos)
                    Next
                    lAddNewRowPos = LossReportEdit.XML01Loss.ActivitiesRS.getRowCount
                    LossReport.XML01Loss.ActivitiesRS.addRows lAddNewRowPos
                    sUpdateMess = sUpdateMess & "Added Record" & vbCrLf
                End If
                If bRSIsNothing Then
                    LossReport.XML01Loss.ActivitiesRS.setField lRSPos, CStr(saryColNames(lColNamesPos)), sNewValue
                Else
                    If sOldUniqueID = vbNullString And lAddNewRowPos = 0 Then
                        LossReport.XML01Loss.ActivitiesRS.addRows 1
                        lAddNewRowPos = LossReport.XML01Loss.ActivitiesRS.getRowCount
                        LossReport.XML01Loss.ActivitiesRS.setField lAddNewRowPos, CStr(saryColNames(lColNamesPos)), sNewValue
                        sUpdateMess = sUpdateMess & "Added Record" & vbCrLf
                    ElseIf lAddNewRowPos > 0 Then
                        LossReport.XML01Loss.ActivitiesRS.setField lAddNewRowPos, CStr(saryColNames(lColNamesPos)), sNewValue
                    Else
                        LossReport.XML01Loss.ActivitiesRS.setField lEditRowPos, CStr(saryColNames(lColNamesPos)), sNewValue
                    End If
                End If
                If StrComp(sOldValue, sNewValue, vbBinaryCompare) <> 0 Then
                    sTemp = CStr(saryColNames(lColNamesPos))
                    sUpdateMess = sUpdateMess & BuildUpdateMess(sTemp, sOldValue, sNewValue, CBool(lAddNewRowPos))
                End If
            Next
        Next
        sUpdateMess = BuildUpdateMessEndRS(sUpdateMess)
    End If
    If Not LossReport.XML01Loss.ActivitiesRS Is Nothing Then
        oWddxStruc.setProp "ActivitiesRS", LossReport.XML01Loss.ActivitiesRS
    End If
    '********************************END ActivitiesRS**********************************************
    

'******************************************************************************************************
'*
'*                                            AUTOMOBILE
'*
'*
'******************************************************************************************************
    
    sLRData = oWddxSer.serialize(oWddxStruc)
    'Take out the new header
    sLRData = Mid(sLRData, InStr(1, sLRData, "<header>", vbTextCompare))
    sLRData = sOldHeader & sLRData
    
    'DEBUG *************************************************************************
    
    goUtil.utSaveFileData "c:\temp\" & sTickCount & "_AutoMobile_02AFTERUPDATE.xml", sLRData
    goUtil.utSaveFileData "c:\temp\" & sTickCount & "_AutoMobile_03UpdateReport.txt", sUpdateMess
    'DEBUG***************************************************************************
    psUpdateUnitReport = sUpdateMess
    UpdateUnitAutomobileXML = sLRData
    
    'Cleanup
    Set oWddxSer = Nothing
    Set oWddxStruc = Nothing
    Set oXML01 = Nothing
    Set oXML01Edit = Nothing

    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
  
    Set oXML01 = Nothing
    Set oXML01Edit = Nothing
    Set oWddxSer = Nothing
    Set oWddxStruc = Nothing
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLR_ClassName & vbCrLf & "Public Function UpdateUnitAutomobileXML"
End Function

Public Function UpdateUnitPropertyXML(psLRData As String, _
                                        poXML01 As clsLossXML01, _
                                        poXML01Edit As clsLossXML01, _
                                        psUpdateUnitReport As String) As String
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
   
    Dim oXML01 As clsLossXML01
    Dim LossReport  As XML01LossReport
    Dim sLRData As String
    Dim oXML01Edit As clsLossXML01
    Dim LossReportEdit As XML01LossReport
    Dim lRSPos As Long
    Dim saryColNames() As Variant
    Dim lColNamesPos As Long
    Dim sOldValue As String
    Dim sNewValue As String
    Dim sUniqueID As String
    Dim sOldUniqueID As String
    Dim bRSIsNothing As Boolean
    Dim saryNewCols() As String
    Dim lNewColsPos As Long
    Dim lAddNewRowPos As Long
    Dim lEditRowPos As Long
    Dim oWddxStruc As WDDXStruct
    Dim oWddxSer As WDDXSerializer
    Dim sOldHeader As String
    Dim ContactDetail As udtXML01ContactDetail
    Dim Contact As udtXML01Contacts
    Dim sTemp As String
    Dim sTemp2 As String
    Dim sRSName As String
    Dim sUpdateMess As String
    Dim sTickCount As String
   
    sTickCount = goUtil.utGetTickCount
   
    Set oXML01 = poXML01
    Set oXML01Edit = poXML01Edit

    
    'DEBUG *************************************************************************
    
    goUtil.utSaveFileData "c:\temp\" & sTickCount & "_Property_01BEFOREUPDATE_.xml", psLRData
    'DEBUG***************************************************************************
    
    'Populate the old Header
    sOldHeader = left(psLRData, InStr(1, psLRData, "<header", vbTextCompare) - 1)
    
    'Now that both the Edit LossReport and Existing Loss Report are populated...
    'Need to loop through the Edit Object and Update the Exisiting.
    
    If oXML01Edit.clsCarLR_LossType <> oXML01.clsCarLR_LossType Then
        Err.Raise -999, , "LossType does not match! "
    End If
'******************************************************************************************************
'*
'*                                            PROPERTY
'*
'*
'******************************************************************************************************
        
    'Now Update the DB with the New Loss Report Data for Automobile
    'Rebuild the Wddx Packet !
    Set oWddxSer = New WDDXSerializer
    Set oWddxStruc = New WDDXStruct
    oWddxStruc.setProp "ClassName", "V2ECcarFarmers.clsLossXML01"
    oWddxStruc.setProp "TransType", "IA_CRN_ASSIGN"
    oWddxStruc.setProp "LossType", "Homeowners"
    
    LossReport = oXML01.clsCarLR_LossReport
    LossReportEdit = oXML01Edit.clsCarLR_LossReport
    
    
    'BEGIN THE UPDATE REPORT HEADER
    
    sUpdateMess = String(50, "*") & vbCrLf
    sUpdateMess = sUpdateMess & "UNIT UPDATE REPORT" & String(4, vbTab) & Now() & vbCrLf
    sUpdateMess = sUpdateMess & "Initial Unit Number: " & String(25 - Len("Initial Unit Numer: "), Chr(32)) & String(1, vbTab) & LossReport.XML01Loss.AssignmentDetailRS.getField(1, "UnitNumber") & vbCrLf
    sUpdateMess = sUpdateMess & "Claim Number: " & String(25 - Len("Claim Number: "), Chr(32)) & String(1, vbTab) & LossReport.XML01Loss.AssignmentDetailRS.getField(1, "ClaimNumber") & vbCrLf
    sUpdateMess = sUpdateMess & "ClaimId: " & String(25 - Len("ClaimId: "), Chr(32)) & String(1, vbTab) & oXML01Edit.ClaimId & vbCrLf
    sUpdateMess = sUpdateMess & "UnitId: " & String(25 - Len("UnitId: "), Chr(32)) & String(1, vbTab) & oXML01Edit.UniqueID & vbCrLf
    sUpdateMess = sUpdateMess & String(50, "*") & vbCrLf
    
    
    '********************************AssignmentDetailRS**********************************************
    If Not LossReportEdit.XML01Loss.AssignmentDetailRS Is Nothing Then
        sUpdateMess = BuildUpdateMessBeginRS(sUpdateMess, "Assignment Detail")
        bRSIsNothing = False
        lAddNewRowPos = 0
        For lRSPos = 1 To LossReportEdit.XML01Loss.AssignmentDetailRS.getRowCount
            sUniqueID = LossReportEdit.XML01Loss.AssignmentDetailRS.getField(lRSPos, "UniqueID")
            saryColNames() = LossReportEdit.XML01Loss.AssignmentDetailRS.getColumnNames
            For lColNamesPos = LBound(saryColNames, 1) To UBound(saryColNames, 1)
                sOldUniqueID = oXML01.GetAssignmentDetailItem(sUniqueID, "UniqueID", lEditRowPos)
                If sOldUniqueID <> vbNullString Then
                    sOldValue = oXML01.GetAssignmentDetailItem(sUniqueID, CStr(saryColNames(lColNamesPos)))
                Else
                    sOldValue = vbNullString
                    lAddNewRowPos = 0
                End If
                sNewValue = oXML01Edit.GetAssignmentDetailItem(sUniqueID, CStr(saryColNames(lColNamesPos)))
                If LossReport.XML01Loss.AssignmentDetailRS Is Nothing Then
                    bRSIsNothing = True
                    Set LossReport.XML01Loss.AssignmentDetailRS = New WDDXRecordset
                    saryNewCols() = Split("UniqueID,UnitNumber,ClaimNumber,DateAssigned,CatastropheCode,LossDate,Type,CauseOfLoss,FirstName,LastName,AssignedTo,AssignedToFirstName,AssignedToLastName", ",")
                    For lNewColsPos = LBound(saryNewCols, 1) To UBound(saryNewCols, 1)
                        LossReport.XML01Loss.AssignmentDetailRS.addColumn saryNewCols(lNewColsPos)
                    Next
                    lAddNewRowPos = LossReportEdit.XML01Loss.AssignmentDetailRS.getRowCount
                    LossReport.XML01Loss.AssignmentDetailRS.addRows lAddNewRowPos
                    sUpdateMess = sUpdateMess & "Added Record" & vbCrLf
                End If
                If bRSIsNothing Then
                    LossReport.XML01Loss.AssignmentDetailRS.setField lRSPos, CStr(saryColNames(lColNamesPos)), sNewValue
                Else
                    If sOldUniqueID = vbNullString And lAddNewRowPos = 0 Then
                        LossReport.XML01Loss.AssignmentDetailRS.addRows 1
                        lAddNewRowPos = LossReport.XML01Loss.AssignmentDetailRS.getRowCount
                        LossReport.XML01Loss.AssignmentDetailRS.setField lAddNewRowPos, CStr(saryColNames(lColNamesPos)), sNewValue
                        sUpdateMess = sUpdateMess & "Added Record" & vbCrLf
                    ElseIf lAddNewRowPos > 0 Then
                        LossReport.XML01Loss.AssignmentDetailRS.setField lAddNewRowPos, CStr(saryColNames(lColNamesPos)), sNewValue
                    Else
                        LossReport.XML01Loss.AssignmentDetailRS.setField lEditRowPos, CStr(saryColNames(lColNamesPos)), sNewValue
                    End If
                End If
                'See if an Assignment Field Update For Property is Required
                If StrComp(sOldValue, sNewValue, vbBinaryCompare) <> 0 Then
                    sTemp = CStr(saryColNames(lColNamesPos))
                    sUpdateMess = sUpdateMess & BuildUpdateMess(sTemp, sOldValue, sNewValue, CBool(lAddNewRowPos))
                    If Not mbInsuredUpdateRequired Then
                        If StrComp(sTemp, "LastName", vbTextCompare) = 0 Then
                            mbInsuredUpdateRequired = True
                        ElseIf StrComp(sTemp, "FirstName", vbTextCompare) = 0 Then
                            mbInsuredUpdateRequired = True
                        End If
                    End If
                End If
            Next
        Next
        sUpdateMess = BuildUpdateMessEndRS(sUpdateMess)
    End If
    If Not LossReport.XML01Loss.AssignmentDetailRS Is Nothing Then
        oWddxStruc.setProp "AssignmentDetailRS", LossReport.XML01Loss.AssignmentDetailRS
    End If
    '********************************END AssignmentDetailRS**********************************************
    
    '******************************LossDetailRS***********************
    If Not LossReportEdit.XML01Loss.LossDetailRS Is Nothing Then
        sUpdateMess = BuildUpdateMessBeginRS(sUpdateMess, "Loss Detail")
        bRSIsNothing = False
        lAddNewRowPos = 0
        For lRSPos = 1 To LossReportEdit.XML01Loss.LossDetailRS.getRowCount
            sUniqueID = LossReportEdit.XML01Loss.LossDetailRS.getField(lRSPos, "UniqueID")
            saryColNames() = LossReportEdit.XML01Loss.LossDetailRS.getColumnNames
            For lColNamesPos = LBound(saryColNames, 1) To UBound(saryColNames, 1)
                sOldUniqueID = oXML01.GetLossDetailItem(sUniqueID, "UniqueID", lEditRowPos)
                If sOldUniqueID <> vbNullString Then
                    sOldValue = oXML01.GetLossDetailItem(sUniqueID, CStr(saryColNames(lColNamesPos)))
                Else
                    sOldValue = vbNullString
                    lAddNewRowPos = 0
                End If
                sNewValue = oXML01Edit.GetLossDetailItem(sUniqueID, CStr(saryColNames(lColNamesPos)))
                If LossReport.XML01Loss.LossDetailRS Is Nothing Then
                    bRSIsNothing = True
                    Set LossReport.XML01Loss.LossDetailRS = New WDDXRecordset
                    saryNewCols() = Split("UniqueID,LossLocationAddress,LossLocationAddress2,LossLocationCity,LossLocationState,LossLocationZip,PropertyAddress,PropertyCity,PropertyState,PropertyZip,AffectedAreas,LossDescription", ",")
                    For lNewColsPos = LBound(saryNewCols, 1) To UBound(saryNewCols, 1)
                        LossReport.XML01Loss.LossDetailRS.addColumn saryNewCols(lNewColsPos)
                    Next
                    lAddNewRowPos = LossReportEdit.XML01Loss.LossDetailRS.getRowCount
                    LossReport.XML01Loss.LossDetailRS.addRows lAddNewRowPos
                    sUpdateMess = sUpdateMess & "Added Record" & vbCrLf
                End If
                If bRSIsNothing Then
                    LossReport.XML01Loss.LossDetailRS.setField lRSPos, CStr(saryColNames(lColNamesPos)), sNewValue
                Else
                    If sOldUniqueID = vbNullString And lAddNewRowPos = 0 Then
                        LossReport.XML01Loss.LossDetailRS.addRows 1
                        lAddNewRowPos = LossReport.XML01Loss.LossDetailRS.getRowCount
                        LossReport.XML01Loss.LossDetailRS.setField lAddNewRowPos, CStr(saryColNames(lColNamesPos)), sNewValue
                        sUpdateMess = sUpdateMess & "Added Record" & vbCrLf
                    ElseIf lAddNewRowPos > 0 Then
                        LossReport.XML01Loss.LossDetailRS.setField lAddNewRowPos, CStr(saryColNames(lColNamesPos)), sNewValue
                    Else
                        LossReport.XML01Loss.LossDetailRS.setField lEditRowPos, CStr(saryColNames(lColNamesPos)), sNewValue
                    End If
                End If
                'See if a Property Address Update For Homeowners is Required
                If StrComp(sOldValue, sNewValue, vbBinaryCompare) <> 0 Then
                    sTemp = CStr(saryColNames(lColNamesPos))
                    sUpdateMess = sUpdateMess & BuildUpdateMess(sTemp, sOldValue, sNewValue, CBool(lAddNewRowPos))
                    If Not mbPropertyAddressUpdateRequired Then
                        If StrComp(sTemp, "LossLocationAddress", vbTextCompare) = 0 Then
                            mbPropertyAddressUpdateRequired = True
                        ElseIf StrComp(sTemp, "LossLocationAddress2", vbTextCompare) = 0 Then
                            mbPropertyAddressUpdateRequired = True
                        ElseIf StrComp(sTemp, "LossLocationCity", vbTextCompare) = 0 Then
                            mbPropertyAddressUpdateRequired = True
                        ElseIf StrComp(sTemp, "LossLocationState", vbTextCompare) = 0 Then
                            mbPropertyAddressUpdateRequired = True
                        ElseIf StrComp(sTemp, "LossLocationZip", vbTextCompare) = 0 Then
                            mbPropertyAddressUpdateRequired = True
                        End If
                    End If
                End If
            Next
        Next
        sUpdateMess = BuildUpdateMessEndRS(sUpdateMess)
    End If
    If Not LossReport.XML01Loss.LossDetailRS Is Nothing Then
        oWddxStruc.setProp "LossDetailRS", LossReport.XML01Loss.LossDetailRS
    End If
    '******************************END LossDetailRS***********************
    
    '********************************ContactDetailRS**********************************************
    If Not LossReportEdit.XML01Loss.ContactDetailRS Is Nothing Then
        sUpdateMess = BuildUpdateMessBeginRS(sUpdateMess, "Contact Details")
        bRSIsNothing = False
        lAddNewRowPos = 0
        For lRSPos = 1 To LossReportEdit.XML01Loss.ContactDetailRS.getRowCount
            sUniqueID = LossReportEdit.XML01Loss.ContactDetailRS.getField(lRSPos, "UniqueID")
            saryColNames() = LossReportEdit.XML01Loss.ContactDetailRS.getColumnNames
            For lColNamesPos = LBound(saryColNames, 1) To UBound(saryColNames, 1)
                sOldUniqueID = oXML01.GetContactDetailsItem(sUniqueID, "UniqueID", lEditRowPos)
                If sOldUniqueID <> vbNullString Then
                    sOldValue = oXML01.GetContactDetailsItem(sUniqueID, CStr(saryColNames(lColNamesPos)))
                Else
                    sOldValue = vbNullString
                    lAddNewRowPos = 0
                End If
                sNewValue = oXML01Edit.GetContactDetailsItem(sUniqueID, CStr(saryColNames(lColNamesPos)))
                If LossReport.XML01Loss.ContactDetailRS Is Nothing Then
                    bRSIsNothing = True
                    Set LossReport.XML01Loss.ContactDetailRS = New WDDXRecordset
                    saryNewCols() = Split("UniqueID,AgentFirstName,AgentLastName,AgentPrimaryPhone,ContactRowID", ",")
                    For lNewColsPos = LBound(saryNewCols, 1) To UBound(saryNewCols, 1)
                        LossReport.XML01Loss.ContactDetailRS.addColumn saryNewCols(lNewColsPos)
                    Next
                    lAddNewRowPos = LossReportEdit.XML01Loss.ContactDetailRS.getRowCount
                    LossReport.XML01Loss.ContactDetailRS.addRows lAddNewRowPos
                    sUpdateMess = sUpdateMess & "Added Record" & vbCrLf
                End If
                If bRSIsNothing Then
                    LossReport.XML01Loss.ContactDetailRS.setField lRSPos, CStr(saryColNames(lColNamesPos)), sNewValue
                Else
                    If sOldUniqueID = vbNullString And lAddNewRowPos = 0 Then
                        LossReport.XML01Loss.ContactDetailRS.addRows 1
                        lAddNewRowPos = LossReport.XML01Loss.ContactDetailRS.getRowCount
                        LossReport.XML01Loss.ContactDetailRS.setField lAddNewRowPos, CStr(saryColNames(lColNamesPos)), sNewValue
                        sUpdateMess = sUpdateMess & "Added Record" & vbCrLf
                    ElseIf lAddNewRowPos > 0 Then
                        LossReport.XML01Loss.ContactDetailRS.setField lAddNewRowPos, CStr(saryColNames(lColNamesPos)), sNewValue
                    Else
                        LossReport.XML01Loss.ContactDetailRS.setField lEditRowPos, CStr(saryColNames(lColNamesPos)), sNewValue
                    End If
                End If
                'See if an Assignment Field Update For Property is Required
                If StrComp(sOldValue, sNewValue, vbBinaryCompare) <> 0 Then
                    sTemp = CStr(saryColNames(lColNamesPos))
                    sUpdateMess = sUpdateMess & BuildUpdateMess(sTemp, sOldValue, sNewValue, CBool(lAddNewRowPos))
                    If Not mbAgentNoUpdateRequired Then
                        If StrComp(sTemp, "AgentPrimaryPhone", vbTextCompare) = 0 Then
                            mbAgentNoUpdateRequired = True
                        End If
                    End If
                End If
            Next
        Next
        sUpdateMess = BuildUpdateMessEndRS(sUpdateMess)
    End If
    If Not LossReport.XML01Loss.ContactDetailRS Is Nothing Then
        oWddxStruc.setProp "ContactDetailsRS", LossReport.XML01Loss.ContactDetailRS
        'Set the current contactID
        ContactDetail.ContactRowID = LossReport.XML01Loss.ContactDetailRS.getField(1, "ContactRowID")
    End If
    '********************************END ContactDetailRS**********************************************
    
    
    '********************************ContactsRS**********************************************
    If Not LossReportEdit.XML01Loss.ContactsRS Is Nothing Then
        sUpdateMess = BuildUpdateMessBeginRS(sUpdateMess, "Contacts")
        bRSIsNothing = False
        lAddNewRowPos = 0
        For lRSPos = 1 To LossReportEdit.XML01Loss.ContactsRS.getRowCount
            sUniqueID = LossReportEdit.XML01Loss.ContactsRS.getField(lRSPos, "ContactRowID")
            saryColNames() = LossReportEdit.XML01Loss.ContactsRS.getColumnNames
            For lColNamesPos = LBound(saryColNames, 1) To UBound(saryColNames, 1)
                sOldUniqueID = oXML01.GetContactItem(sUniqueID, "ContactRowID", lEditRowPos)
                If sOldUniqueID <> vbNullString Then
                    sOldValue = oXML01.GetContactItem(sUniqueID, CStr(saryColNames(lColNamesPos)))
                Else
                    sOldValue = vbNullString
                    lAddNewRowPos = 0
                End If
                sNewValue = oXML01Edit.GetContactItem(sUniqueID, CStr(saryColNames(lColNamesPos)))
                If LossReport.XML01Loss.ContactsRS Is Nothing Then
                    bRSIsNothing = True
                    Set LossReport.XML01Loss.ContactsRS = New WDDXRecordset
                    saryNewCols() = Split("ContactRowID,MM,FirstName,LastName,ContactRole,BestPhoneNumber,PrimaryPhoneNumber,HomePhoneNumber,CellularPhoneNumber,FaxPhoneNumber,EmailAddress,PagerPhoneNumber,PagerPIN,WorkPhoneNumber,RelationshipToInsured,PrimaryAddressID,CreatedUpdated", ",")
                    For lNewColsPos = LBound(saryNewCols, 1) To UBound(saryNewCols, 1)
                        LossReport.XML01Loss.ContactsRS.addColumn saryNewCols(lNewColsPos)
                    Next
                    lAddNewRowPos = LossReportEdit.XML01Loss.ContactsRS.getRowCount
                    LossReport.XML01Loss.ContactsRS.addRows lAddNewRowPos
                    sUpdateMess = sUpdateMess & "Added Record" & vbCrLf
                End If
                If bRSIsNothing Then
                    LossReport.XML01Loss.ContactsRS.setField lRSPos, CStr(saryColNames(lColNamesPos)), sNewValue
                Else
                    If sOldUniqueID = vbNullString And lAddNewRowPos = 0 Then
                        LossReport.XML01Loss.ContactsRS.addRows 1
                        lAddNewRowPos = LossReport.XML01Loss.ContactsRS.getRowCount
                        LossReport.XML01Loss.ContactsRS.setField lAddNewRowPos, CStr(saryColNames(lColNamesPos)), sNewValue
                        sUpdateMess = sUpdateMess & "Added Record" & vbCrLf
                    ElseIf lAddNewRowPos > 0 Then
                        LossReport.XML01Loss.ContactsRS.setField lAddNewRowPos, CStr(saryColNames(lColNamesPos)), sNewValue
                    Else
                        LossReport.XML01Loss.ContactsRS.setField lEditRowPos, CStr(saryColNames(lColNamesPos)), sNewValue
                    End If
                End If
                'See if an Assignment Field Update For Automobile is Required
                If StrComp(sOldValue, sNewValue, vbBinaryCompare) <> 0 And StrComp(ContactDetail.ContactRowID, sUniqueID, vbTextCompare) = 0 Then
                    sTemp = CStr(saryColNames(lColNamesPos))
                    sUpdateMess = sUpdateMess & BuildUpdateMess(sTemp, sOldValue, sNewValue, CBool(lAddNewRowPos))
                    If Not mbHomePhoneUpdateRequired Then
                        If StrComp(sTemp, "HomePhoneNumber", vbTextCompare) = 0 Then
                            mbHomePhoneUpdateRequired = True
                        End If
                    End If
                    If Not mbBusinessPhoneUpdateRequired Then
                        If StrComp(sTemp, "WorkPhoneNumber", vbTextCompare) = 0 Then
                            mbBusinessPhoneUpdateRequired = True
                        End If
                    End If
                ElseIf StrComp(sOldValue, sNewValue, vbBinaryCompare) <> 0 Then
                    sTemp = CStr(saryColNames(lColNamesPos))
                    sUpdateMess = sUpdateMess & BuildUpdateMess(sTemp, sOldValue, sNewValue, CBool(lAddNewRowPos))
                End If
            Next
        Next
        sUpdateMess = BuildUpdateMessEndRS(sUpdateMess)
    End If
    If Not LossReport.XML01Loss.ContactsRS Is Nothing Then
        oWddxStruc.setProp "ContactRS", LossReport.XML01Loss.ContactsRS
    End If
    '********************************END ContactsRS**********************************************
    
     '****************************AddressRS*****************************
    If Not LossReportEdit.XML01Loss.AddressRS Is Nothing Then
        sUpdateMess = BuildUpdateMessBeginRS(sUpdateMess, "Address")
        bRSIsNothing = False
        lAddNewRowPos = 0
        For lRSPos = 1 To LossReportEdit.XML01Loss.AddressRS.getRowCount
            sUniqueID = LossReportEdit.XML01Loss.AddressRS.getField(lRSPos, "AddressID")
            saryColNames() = LossReportEdit.XML01Loss.AddressRS.getColumnNames
            For lColNamesPos = LBound(saryColNames, 1) To UBound(saryColNames, 1)
                sOldUniqueID = oXML01.GetAddressItem(sUniqueID, "AddressID", lEditRowPos)
                If sOldUniqueID <> vbNullString Then
                    sOldValue = oXML01.GetAddressItem(sUniqueID, CStr(saryColNames(lColNamesPos)))
                Else
                    sOldValue = vbNullString
                    lAddNewRowPos = 0
                End If
                sNewValue = oXML01Edit.GetAddressItem(sUniqueID, CStr(saryColNames(lColNamesPos)))
                If LossReport.XML01Loss.AddressRS Is Nothing Then
                    bRSIsNothing = True
                    Set LossReport.XML01Loss.AddressRS = New WDDXRecordset
                    saryNewCols() = Split("ContactRowID,AddressID,Type,StreetAddress,StreetAddress2,City,State,PostalCode", ",")
                    For lNewColsPos = LBound(saryNewCols, 1) To UBound(saryNewCols, 1)
                        LossReport.XML01Loss.AddressRS.addColumn saryNewCols(lNewColsPos)
                    Next
                    lAddNewRowPos = LossReportEdit.XML01Loss.AddressRS.getRowCount
                    LossReport.XML01Loss.AddressRS.addRows lAddNewRowPos
                    sUpdateMess = sUpdateMess & "Added Record" & vbCrLf
                End If
                If bRSIsNothing Then
                    LossReport.XML01Loss.AddressRS.setField lRSPos, CStr(saryColNames(lColNamesPos)), sNewValue
                Else
                    If sOldUniqueID = vbNullString And lAddNewRowPos = 0 Then
                        LossReport.XML01Loss.AddressRS.addRows 1
                        lAddNewRowPos = LossReport.XML01Loss.AddressRS.getRowCount
                        LossReport.XML01Loss.AddressRS.setField lAddNewRowPos, CStr(saryColNames(lColNamesPos)), sNewValue
                        sUpdateMess = sUpdateMess & "Added Record" & vbCrLf
                    ElseIf lAddNewRowPos > 0 Then
                        LossReport.XML01Loss.AddressRS.setField lAddNewRowPos, CStr(saryColNames(lColNamesPos)), sNewValue
                    Else
                        LossReport.XML01Loss.AddressRS.setField lEditRowPos, CStr(saryColNames(lColNamesPos)), sNewValue
                    End If
                End If
                'See if a Mailing Address Update For Homeowners is Required
                If StrComp(sOldValue, sNewValue, vbBinaryCompare) <> 0 Then
                    sTemp = CStr(saryColNames(lColNamesPos))
                    sUpdateMess = sUpdateMess & BuildUpdateMess(sTemp, sOldValue, sNewValue, CBool(lAddNewRowPos))
                    If Not mbMaillingAddressUpdateRequired Then
                        sTemp2 = oXML01Edit.GetAddressItem(sUniqueID, "Type")
                        If StrComp(sTemp2, "Mailing", vbTextCompare) = 0 Then
                            mbMaillingAddressUpdateRequired = True
                        End If
                    End If
                End If
            Next
        Next
        sUpdateMess = BuildUpdateMessEndRS(sUpdateMess)
    End If
    If Not LossReport.XML01Loss.AddressRS Is Nothing Then
        oWddxStruc.setProp "AddressRS", LossReport.XML01Loss.AddressRS
    End If
    '******************************END AdressRS***********************
    
    
    '********************************PolicyDetailRS**********************************************
    If Not LossReportEdit.XML01Loss.PolicyDetailRS Is Nothing Then
        sUpdateMess = BuildUpdateMessBeginRS(sUpdateMess, "Policy Detail")
        bRSIsNothing = False
        'Policy Details should only ever have 1 Row !!!
        lRSPos = 1
        'No Unique ID
        saryColNames() = LossReportEdit.XML01Loss.PolicyDetailRS.getColumnNames
        For lColNamesPos = LBound(saryColNames, 1) To UBound(saryColNames, 1)
            If Not LossReport.XML01Loss.PolicyDetailRS Is Nothing Then
                sOldValue = oXML01.GetPolicyDetailItem(CStr(saryColNames(lColNamesPos)))
            Else
                sOldValue = vbNullString
            End If
            sNewValue = oXML01Edit.GetPolicyDetailItem(CStr(saryColNames(lColNamesPos)))
            If LossReport.XML01Loss.PolicyDetailRS Is Nothing Then
                bRSIsNothing = True
                Set LossReport.XML01Loss.PolicyDetailRS = New WDDXRecordset
                saryNewCols() = Split("PolicyNumber,Status,CoverageStatus,BalanceDue,CompanyCode,CompanyName,RenewalDate,CancellationDate,NewBusinessDate,PolicyDescription,PolicyEdition,MortgageeName,MortgageeAddress,LienHolderName,LienHolderAddress", ",")
                For lNewColsPos = LBound(saryNewCols, 1) To UBound(saryNewCols, 1)
                    LossReport.XML01Loss.PolicyDetailRS.addColumn saryNewCols(lNewColsPos)
                Next
                lAddNewRowPos = lRSPos 'LossReportEdit.XML01Loss.PolicyDetailRS.getRowCount
                LossReport.XML01Loss.PolicyDetailRS.addRows lAddNewRowPos
                sUpdateMess = sUpdateMess & "Added Record" & vbCrLf
            End If
            LossReport.XML01Loss.PolicyDetailRS.setField lRSPos, CStr(saryColNames(lColNamesPos)), sNewValue
            'See if an Assignment Field Update For Property is Required
            If StrComp(sOldValue, sNewValue, vbBinaryCompare) <> 0 Then
                sTemp = CStr(saryColNames(lColNamesPos))
                sUpdateMess = sUpdateMess & BuildUpdateMess(sTemp, sOldValue, sNewValue, CBool(lAddNewRowPos))
                If Not mbPolicyNoUpdateRequired Then
                    If StrComp(sTemp, "PolicyNumber", vbTextCompare) = 0 Then
                        mbPolicyNoUpdateRequired = True
                    End If
                End If
                If Not mbPolicyDescriptionUpdateRequired Then
                    If StrComp(sTemp, "PolicyDescription", vbTextCompare) = 0 Then
                        mbPolicyDescriptionUpdateRequired = True
                    ElseIf StrComp(sTemp, "PolicyEdition", vbTextCompare) = 0 Then
                        mbPolicyDescriptionUpdateRequired = True
                    End If
                End If
                If Not mbMortgageeNameUpdateRequired Then
                    If StrComp(sTemp, "MortgageeName", vbTextCompare) = 0 Then
                        mbMortgageeNameUpdateRequired = True
                    End If
                End If
            End If
        Next
        sUpdateMess = BuildUpdateMessEndRS(sUpdateMess)
    End If
    If Not LossReport.XML01Loss.PolicyDetailRS Is Nothing Then
        oWddxStruc.setProp "PolicyDetailRS", LossReport.XML01Loss.PolicyDetailRS
    End If
    '********************************END PolicyDetailRS**********************************************
    
    '********************************CoverageRS**********************************************
    If Not LossReportEdit.XML01Loss.CoverageRS Is Nothing Then
        sUpdateMess = BuildUpdateMessBeginRS(sUpdateMess, "Coverage")
        bRSIsNothing = False
        lAddNewRowPos = 0
        For lRSPos = 1 To LossReportEdit.XML01Loss.CoverageRS.getRowCount
            sUniqueID = LossReportEdit.XML01Loss.CoverageRS.getField(lRSPos, "CoverageRowID")
            saryColNames() = LossReportEdit.XML01Loss.CoverageRS.getColumnNames
            For lColNamesPos = LBound(saryColNames, 1) To UBound(saryColNames, 1)
                sOldUniqueID = oXML01.GetCoverageItem(sUniqueID, "CoverageRowID", lEditRowPos)
                If sOldUniqueID <> vbNullString Then
                    sOldValue = oXML01.GetCoverageItem(sUniqueID, CStr(saryColNames(lColNamesPos)))
                Else
                    sOldValue = vbNullString
                    lAddNewRowPos = 0
                End If
                sNewValue = oXML01Edit.GetCoverageItem(sUniqueID, CStr(saryColNames(lColNamesPos)))
                If LossReport.XML01Loss.CoverageRS Is Nothing Then
                    bRSIsNothing = True
                    Set LossReport.XML01Loss.CoverageRS = New WDDXRecordset
                    saryNewCols() = Split("CoverageRowID,Coverage,Limits,Deductible1,Deductible2,Deductible3,Deductible4", ",")
                    For lNewColsPos = LBound(saryNewCols, 1) To UBound(saryNewCols, 1)
                        LossReport.XML01Loss.CoverageRS.addColumn saryNewCols(lNewColsPos)
                    Next
                    lAddNewRowPos = LossReportEdit.XML01Loss.CoverageRS.getRowCount
                    LossReport.XML01Loss.CoverageRS.addRows lAddNewRowPos
                    sUpdateMess = sUpdateMess & "Added Record" & vbCrLf
                End If
                If bRSIsNothing Then
                    LossReport.XML01Loss.CoverageRS.setField lRSPos, CStr(saryColNames(lColNamesPos)), sNewValue
                Else
                    If sOldUniqueID = vbNullString And lAddNewRowPos = 0 Then
                        LossReport.XML01Loss.CoverageRS.addRows 1
                        lAddNewRowPos = LossReport.XML01Loss.CoverageRS.getRowCount
                        LossReport.XML01Loss.CoverageRS.setField lAddNewRowPos, CStr(saryColNames(lColNamesPos)), sNewValue
                        sUpdateMess = sUpdateMess & "Added Record" & vbCrLf
                    ElseIf lAddNewRowPos > 0 Then
                        LossReport.XML01Loss.CoverageRS.setField lAddNewRowPos, CStr(saryColNames(lColNamesPos)), sNewValue
                    Else
                        LossReport.XML01Loss.CoverageRS.setField lEditRowPos, CStr(saryColNames(lColNamesPos)), sNewValue
                    End If
                End If
                'See if an Assignment Field Update For Property is Required
                If StrComp(sOldValue, sNewValue, vbBinaryCompare) <> 0 Then
                    sTemp = CStr(saryColNames(lColNamesPos))
                    sUpdateMess = sUpdateMess & BuildUpdateMess(sTemp, sOldValue, sNewValue, CBool(lAddNewRowPos))
                    If Not mbDeductibleUpdateRequired Then
                        If StrComp(sTemp, "Deductible1", vbTextCompare) = 0 Then
                            mbDeductibleUpdateRequired = True
                        End If
                    End If
                End If
            Next
        Next
        sUpdateMess = BuildUpdateMessEndRS(sUpdateMess)
    End If
    If Not LossReport.XML01Loss.CoverageRS Is Nothing Then
        oWddxStruc.setProp "CoverageRS", LossReport.XML01Loss.CoverageRS
    End If
    '********************************END CoverageRS**********************************************
    
    
    '********************************EndorsementRS**********************************************
    If Not LossReportEdit.XML01Loss.EndorsementRS Is Nothing Then
        sUpdateMess = BuildUpdateMessBeginRS(sUpdateMess, "Endorsement")
        bRSIsNothing = False
        lAddNewRowPos = 0
        For lRSPos = 1 To LossReportEdit.XML01Loss.EndorsementRS.getRowCount
            sUniqueID = LossReportEdit.XML01Loss.EndorsementRS.getField(lRSPos, "EndorsementsRowID")
            saryColNames() = LossReportEdit.XML01Loss.EndorsementRS.getColumnNames
            For lColNamesPos = LBound(saryColNames, 1) To UBound(saryColNames, 1)
                sOldUniqueID = oXML01.GetEndorsementItem(sUniqueID, "EndorsementsRowID", lEditRowPos)
                If sOldUniqueID <> vbNullString Then
                    sOldValue = oXML01.GetEndorsementItem(sUniqueID, CStr(saryColNames(lColNamesPos)))
                Else
                    sOldValue = vbNullString
                    lAddNewRowPos = 0
                End If
                sNewValue = oXML01Edit.GetEndorsementItem(sUniqueID, CStr(saryColNames(lColNamesPos)))
                If LossReport.XML01Loss.EndorsementRS Is Nothing Then
                    bRSIsNothing = True
                    Set LossReport.XML01Loss.EndorsementRS = New WDDXRecordset
                    saryNewCols() = Split("EndorsementsRowID,EndorsementNumber,EndorsementDescription", ",")
                    For lNewColsPos = LBound(saryNewCols, 1) To UBound(saryNewCols, 1)
                        LossReport.XML01Loss.EndorsementRS.addColumn saryNewCols(lNewColsPos)
                    Next
                    lAddNewRowPos = LossReportEdit.XML01Loss.EndorsementRS.getRowCount
                    LossReport.XML01Loss.EndorsementRS.addRows lAddNewRowPos
                    sUpdateMess = sUpdateMess & "Added Record" & vbCrLf
                End If
                If bRSIsNothing Then
                    LossReport.XML01Loss.EndorsementRS.setField lRSPos, CStr(saryColNames(lColNamesPos)), sNewValue
                Else
                    If sOldUniqueID = vbNullString And lAddNewRowPos = 0 Then
                        LossReport.XML01Loss.EndorsementRS.addRows 1
                        lAddNewRowPos = LossReport.XML01Loss.EndorsementRS.getRowCount
                        LossReport.XML01Loss.EndorsementRS.setField lAddNewRowPos, CStr(saryColNames(lColNamesPos)), sNewValue
                        sUpdateMess = sUpdateMess & "Added Record" & vbCrLf
                    ElseIf lAddNewRowPos > 0 Then
                        LossReport.XML01Loss.EndorsementRS.setField lAddNewRowPos, CStr(saryColNames(lColNamesPos)), sNewValue
                    Else
                        LossReport.XML01Loss.EndorsementRS.setField lEditRowPos, CStr(saryColNames(lColNamesPos)), sNewValue
                    End If
                End If
                If StrComp(sOldValue, sNewValue, vbBinaryCompare) <> 0 Then
                    sTemp = CStr(saryColNames(lColNamesPos))
                    sUpdateMess = sUpdateMess & BuildUpdateMess(sTemp, sOldValue, sNewValue, CBool(lAddNewRowPos))
                End If
            Next
        Next
        sUpdateMess = BuildUpdateMessEndRS(sUpdateMess)
    End If
    If Not LossReport.XML01Loss.EndorsementRS Is Nothing Then
        oWddxStruc.setProp "EndorsementRS", LossReport.XML01Loss.EndorsementRS
    End If
    '********************************END EndorsementRS**********************************************
    
     '********************************PaymentDetailRS**********************************************
    If Not LossReportEdit.XML01Loss.PaymentDetailRS Is Nothing Then
        sUpdateMess = BuildUpdateMessBeginRS(sUpdateMess, "Payment Detail")
        bRSIsNothing = False
        lAddNewRowPos = 0
        For lRSPos = 1 To LossReportEdit.XML01Loss.PaymentDetailRS.getRowCount
            sUniqueID = LossReportEdit.XML01Loss.PaymentDetailRS.getField(lRSPos, "PaymentDetailRowID")
            saryColNames() = LossReportEdit.XML01Loss.PaymentDetailRS.getColumnNames
            For lColNamesPos = LBound(saryColNames, 1) To UBound(saryColNames, 1)
                sOldUniqueID = oXML01.GetPaymentDetailItem(sUniqueID, "PaymentDetailRowID", lEditRowPos)
                If sOldUniqueID <> vbNullString Then
                    sOldValue = oXML01.GetPaymentDetailItem(sUniqueID, CStr(saryColNames(lColNamesPos)))
                Else
                    sOldValue = vbNullString
                    lAddNewRowPos = 0
                End If
                sNewValue = oXML01Edit.GetPaymentDetailItem(sUniqueID, CStr(saryColNames(lColNamesPos)))
                If LossReport.XML01Loss.PaymentDetailRS Is Nothing Then
                    bRSIsNothing = True
                    Set LossReport.XML01Loss.PaymentDetailRS = New WDDXRecordset
                    saryNewCols() = Split("PaymentDetailRowID,DateIssued,PayeeLineOne,PayeeLineTwo,PayeeLineThree,PayeeLineFour,AccountType,PaymentClass,PaymentAmount", ",")
                    For lNewColsPos = LBound(saryNewCols, 1) To UBound(saryNewCols, 1)
                        LossReport.XML01Loss.PaymentDetailRS.addColumn saryNewCols(lNewColsPos)
                    Next
                    lAddNewRowPos = LossReportEdit.XML01Loss.PaymentDetailRS.getRowCount
                    LossReport.XML01Loss.PaymentDetailRS.addRows lAddNewRowPos
                    sUpdateMess = sUpdateMess & "Added Record" & vbCrLf
                End If
                If bRSIsNothing Then
                    LossReport.XML01Loss.PaymentDetailRS.setField lRSPos, CStr(saryColNames(lColNamesPos)), sNewValue
                Else
                    If sOldUniqueID = vbNullString And lAddNewRowPos = 0 Then
                        LossReport.XML01Loss.PaymentDetailRS.addRows 1
                        lAddNewRowPos = LossReport.XML01Loss.PaymentDetailRS.getRowCount
                        LossReport.XML01Loss.PaymentDetailRS.setField lAddNewRowPos, CStr(saryColNames(lColNamesPos)), sNewValue
                        sUpdateMess = sUpdateMess & "Added Record" & vbCrLf
                    ElseIf lAddNewRowPos > 0 Then
                        LossReport.XML01Loss.PaymentDetailRS.setField lAddNewRowPos, CStr(saryColNames(lColNamesPos)), sNewValue
                    Else
                        LossReport.XML01Loss.PaymentDetailRS.setField lEditRowPos, CStr(saryColNames(lColNamesPos)), sNewValue
                    End If
                End If
                If StrComp(sOldValue, sNewValue, vbBinaryCompare) <> 0 Then
                    sTemp = CStr(saryColNames(lColNamesPos))
                    sUpdateMess = sUpdateMess & BuildUpdateMess(sTemp, sOldValue, sNewValue, CBool(lAddNewRowPos))
                End If
            Next
        Next
        sUpdateMess = BuildUpdateMessEndRS(sUpdateMess)
    End If
    If Not LossReport.XML01Loss.PaymentDetailRS Is Nothing Then
        oWddxStruc.setProp "PaymentDetailRS", LossReport.XML01Loss.PaymentDetailRS
    End If
    '********************************END PaymentDetailRS**********************************************
    
    '********************************PriorLossDetailRS**********************************************
    If Not LossReportEdit.XML01Loss.PriorLossDetailRS Is Nothing Then
        sUpdateMess = BuildUpdateMessBeginRS(sUpdateMess, "Prior Loss Detail")
        bRSIsNothing = False
        lAddNewRowPos = 0
        For lRSPos = 1 To LossReportEdit.XML01Loss.PriorLossDetailRS.getRowCount
            sUniqueID = LossReportEdit.XML01Loss.PriorLossDetailRS.getField(lRSPos, "PriorLossRowID")
            saryColNames() = LossReportEdit.XML01Loss.PriorLossDetailRS.getColumnNames
            For lColNamesPos = LBound(saryColNames, 1) To UBound(saryColNames, 1)
                sOldUniqueID = oXML01.GetPriorLossDetailItem(sUniqueID, "PriorLossRowID", lEditRowPos)
                If sOldUniqueID <> vbNullString Then
                    sOldValue = oXML01.GetPriorLossDetailItem(sUniqueID, CStr(saryColNames(lColNamesPos)))
                Else
                    sOldValue = vbNullString
                    lAddNewRowPos = 0
                End If
                sNewValue = oXML01Edit.GetPriorLossDetailItem(sUniqueID, CStr(saryColNames(lColNamesPos)))
                If LossReport.XML01Loss.PriorLossDetailRS Is Nothing Then
                    bRSIsNothing = True
                    Set LossReport.XML01Loss.PriorLossDetailRS = New WDDXRecordset
                    saryNewCols() = Split("PriorLossRowID,SALN,ClaimSegmentNumber,PolicyNumber,LossCause,ClaimClass,LossDate,SummaryAmount", ",")
                    For lNewColsPos = LBound(saryNewCols, 1) To UBound(saryNewCols, 1)
                        LossReport.XML01Loss.PriorLossDetailRS.addColumn saryNewCols(lNewColsPos)
                    Next
                    lAddNewRowPos = LossReportEdit.XML01Loss.PriorLossDetailRS.getRowCount
                    LossReport.XML01Loss.PriorLossDetailRS.addRows lAddNewRowPos
                    sUpdateMess = sUpdateMess & "Added Record" & vbCrLf
                End If
                If bRSIsNothing Then
                    LossReport.XML01Loss.PriorLossDetailRS.setField lRSPos, CStr(saryColNames(lColNamesPos)), sNewValue
                Else
                    If sOldUniqueID = vbNullString And lAddNewRowPos = 0 Then
                        LossReport.XML01Loss.PriorLossDetailRS.addRows 1
                        lAddNewRowPos = LossReport.XML01Loss.PriorLossDetailRS.getRowCount
                        LossReport.XML01Loss.PriorLossDetailRS.setField lAddNewRowPos, CStr(saryColNames(lColNamesPos)), sNewValue
                        sUpdateMess = sUpdateMess & "Added Record" & vbCrLf
                    ElseIf lAddNewRowPos > 0 Then
                        LossReport.XML01Loss.PriorLossDetailRS.setField lAddNewRowPos, CStr(saryColNames(lColNamesPos)), sNewValue
                    Else
                        LossReport.XML01Loss.PriorLossDetailRS.setField lEditRowPos, CStr(saryColNames(lColNamesPos)), sNewValue
                    End If
                End If
                If StrComp(sOldValue, sNewValue, vbBinaryCompare) <> 0 Then
                    sTemp = CStr(saryColNames(lColNamesPos))
                    sUpdateMess = sUpdateMess & BuildUpdateMess(sTemp, sOldValue, sNewValue, CBool(lAddNewRowPos))
                End If
            Next
        Next
        sUpdateMess = BuildUpdateMessEndRS(sUpdateMess)
    End If
    If Not LossReport.XML01Loss.PriorLossDetailRS Is Nothing Then
        oWddxStruc.setProp "PriorLossDetailRS", LossReport.XML01Loss.PriorLossDetailRS
    End If
    '********************************END PriorLossDetailRS**********************************************
    
    '********************************ActivitiesRS**********************************************
    If Not LossReportEdit.XML01Loss.ActivitiesRS Is Nothing Then
        sUpdateMess = BuildUpdateMessBeginRS(sUpdateMess, "Activities")
        bRSIsNothing = False
        lAddNewRowPos = 0
        For lRSPos = 1 To LossReportEdit.XML01Loss.ActivitiesRS.getRowCount
            sUniqueID = LossReportEdit.XML01Loss.ActivitiesRS.getField(lRSPos, "ActivityRowID")
            saryColNames() = LossReportEdit.XML01Loss.ActivitiesRS.getColumnNames
            For lColNamesPos = LBound(saryColNames, 1) To UBound(saryColNames, 1)
                sOldUniqueID = oXML01.GetActivitiesItem(sUniqueID, "ActivityRowID", lEditRowPos)
                If sOldUniqueID <> vbNullString Then
                    sOldValue = oXML01.GetActivitiesItem(sUniqueID, CStr(saryColNames(lColNamesPos)))
                Else
                    sOldValue = vbNullString
                    lAddNewRowPos = 0
                End If
                sNewValue = oXML01Edit.GetActivitiesItem(sUniqueID, CStr(saryColNames(lColNamesPos)))
                If LossReport.XML01Loss.ActivitiesRS Is Nothing Then
                    bRSIsNothing = True
                    Set LossReport.XML01Loss.ActivitiesRS = New WDDXRecordset
                    saryNewCols() = Split("ActivityRowID,GMTCreated,CreatedBy,Type,Description,Comment", ",")
                    For lNewColsPos = LBound(saryNewCols, 1) To UBound(saryNewCols, 1)
                        LossReport.XML01Loss.ActivitiesRS.addColumn saryNewCols(lNewColsPos)
                    Next
                    lAddNewRowPos = LossReportEdit.XML01Loss.ActivitiesRS.getRowCount
                    LossReport.XML01Loss.ActivitiesRS.addRows lAddNewRowPos
                    sUpdateMess = sUpdateMess & "Added Record" & vbCrLf
                End If
                If bRSIsNothing Then
                    LossReport.XML01Loss.ActivitiesRS.setField lRSPos, CStr(saryColNames(lColNamesPos)), sNewValue
                Else
                    If sOldUniqueID = vbNullString And lAddNewRowPos = 0 Then
                        LossReport.XML01Loss.ActivitiesRS.addRows 1
                        lAddNewRowPos = LossReport.XML01Loss.ActivitiesRS.getRowCount
                        LossReport.XML01Loss.ActivitiesRS.setField lAddNewRowPos, CStr(saryColNames(lColNamesPos)), sNewValue
                        sUpdateMess = sUpdateMess & "Added Record" & vbCrLf
                    ElseIf lAddNewRowPos > 0 Then
                        LossReport.XML01Loss.ActivitiesRS.setField lAddNewRowPos, CStr(saryColNames(lColNamesPos)), sNewValue
                    Else
                        LossReport.XML01Loss.ActivitiesRS.setField lEditRowPos, CStr(saryColNames(lColNamesPos)), sNewValue
                    End If
                End If
                If StrComp(sOldValue, sNewValue, vbBinaryCompare) <> 0 Then
                    sTemp = CStr(saryColNames(lColNamesPos))
                    sUpdateMess = sUpdateMess & BuildUpdateMess(sTemp, sOldValue, sNewValue, CBool(lAddNewRowPos))
                End If
            Next
        Next
        sUpdateMess = BuildUpdateMessEndRS(sUpdateMess)
    End If
    If Not LossReport.XML01Loss.ActivitiesRS Is Nothing Then
        oWddxStruc.setProp "ActivitiesRS", LossReport.XML01Loss.ActivitiesRS
    End If
    '********************************END ActivitiesRS**********************************************
    
    
'******************************************************************************************************
'*
'*                                        END PROPERTY
'*
'*
'******************************************************************************************************
    
    sLRData = oWddxSer.serialize(oWddxStruc)
    'Take out the new header
    sLRData = Mid(sLRData, InStr(1, sLRData, "<header>", vbTextCompare))
    sLRData = sOldHeader & sLRData
    
    'DEBUG *************************************************************************
    
    goUtil.utSaveFileData "c:\temp\" & sTickCount & "_Property_02AFTERUPDATE.xml", sLRData
    goUtil.utSaveFileData "c:\temp\" & sTickCount & "_Property_03UpdateReport.txt", sUpdateMess
    'DEBUG***************************************************************************
    
    psUpdateUnitReport = sUpdateMess
    UpdateUnitPropertyXML = sLRData
    
    'Cleanup
    Set oWddxSer = Nothing
    Set oWddxStruc = Nothing
    Set oXML01 = Nothing
    Set oXML01Edit = Nothing

    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
  
    Set oXML01 = Nothing
    Set oXML01Edit = Nothing
    Set oWddxSer = Nothing
    Set oWddxStruc = Nothing
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLR_ClassName & vbCrLf & "Public Function UpdateUnitPropertyXML"
End Function

Private Function BuildUpdateMess(psColName As String, psOldValue As String, psNewValue As String, pbAddedRow As Boolean) As String
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim sMess As String
    
    If pbAddedRow Then
        sMess = psColName & ":" & String(50 - Len(psColName), Chr(32)) & String(1, vbTab) & psNewValue & vbCrLf
    Else
        sMess = psColName & ":" & String(50 - Len(psColName), Chr(32)) & String(1, vbTab) & "Changed" & vbCrLf
        sMess = sMess & String(50, Chr(32)) & String(1, vbTab) & "From:" & String(1, vbTab) & psOldValue & vbCrLf
        sMess = sMess & String(50, Chr(32)) & String(1, vbTab) & "to:" & String(1, vbTab) & psNewValue & vbCrLf
    End If
    
    BuildUpdateMess = sMess
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLR_ClassName & vbCrLf & "Private Function BuildUpdateMess"
End Function

Private Function BuildUpdateMessEndRS(psUpdateMess As String) As String
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim sMess As String
    
    sMess = psUpdateMess
    
    
    sMess = sMess & vbCrLf & vbCrLf & String(50, "_") & vbCrLf & vbCrLf
    
    BuildUpdateMessEndRS = sMess
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLR_ClassName & vbCrLf & "Private Function BuildUpdateMessEndRS"
End Function

Private Function BuildUpdateMessBeginRS(psUpdateMess As String, psRSName As String) As String
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim sMess As String
    
    sMess = psUpdateMess
    
    sMess = sMess & vbCrLf & vbCrLf & String(50, "-") & vbCrLf & psRSName & vbCrLf & String(50, "-") & vbCrLf & vbCrLf
    
    BuildUpdateMessBeginRS = sMess
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLR_ClassName & vbCrLf & "Private Function BuildUpdateMessBeginRS"
End Function


Public Function CloseUnit(psUniqueID As String, _
                            Optional pbReassignment As Boolean, _
                            Optional psAssignedTo As String) As Boolean
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim sUserID As String
    Dim sPassword As String
    Dim sConnString As String
    Dim oConn As New ADODB.Connection
    Dim adoRS As ADODB.Recordset
    Dim sSQL As String
    Dim lRecordsAffected As Long
    
    Set oConn = New ADODB.Connection
    
    sUserID = GetECSCryptSetting("V2WebControl", "DBConn", "USERID")
    sPassword = GetECSCryptSetting("V2WebControl", "DBConn", "PASSWORD")
    sConnString = GetSetting("V2WebControl", "DSN", "NAME", vbNullString)
    
    Set oConn = New ADODB.Connection
    oConn.Open sConnString, sUserID, sPassword
    
    sSQL = "z_spuUpdateFarmersUnit "
    sSQL = sSQL & "'" & sUserID & "', "          '@UserName       varchar(20),
    sSQL = sSQL & "'" & psUniqueID & "', "       '@FarmersUnitRowID   varchar(20),
    sSQL = sSQL & "1, "                          '@IsDeleted      bit,
    If pbReassignment Then
        sSQL = sSQL & "'Unit Closed by client.  CRN Reassigned to " & psAssignedTo & ".' "     '@AdminComments      varchar(255)
    Else
        sSQL = sSQL & "'Unit Closed by client.' "     '@AdminComments      varchar(255)
    End If
    
    
    oConn.Execute sSQL, lRecordsAffected
    
    CloseUnit = CBool(lRecordsAffected)
    
    'Cleanup
    Set oConn = Nothing
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    'If the Error is No Records Affected then that means this is a
    'Reassignment from Outside IA to INSIDE IA
    If InStr(1, sErrDesc, "No Records Affected", vbTextCompare) > 0 Then
        Err.Clear
        CloseUnit = False
    Else
        Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLR_ClassName & vbCrLf & "Public Function CloseUnit"
    End If
End Function

Private Function GetECSCryptSetting(psAPP As String, psSECTION As String, psKEY As String, _
                                   Optional pvDefault As Variant = vbNullString) As Variant
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim sCryptSetting As String
    Dim oUtil As V2ECKeyBoard.clsUtil
    
    Set oUtil = New V2ECKeyBoard.clsUtil
    
    sCryptSetting = GetSetting(psAPP, psSECTION, psKEY, vbNullString)
    
    If sCryptSetting <> vbNullString Then
        GetECSCryptSetting = CStr(oUtil.Decode(sCryptSetting))
    Else
        GetECSCryptSetting = pvDefault
    End If
    
    Set oUtil = Nothing
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLR_ClassName & vbCrLf & "Public Function CloseUnit"
End Function

Private Function ActiveFiles(psActiveFileDir As String, _
                             Optional psWildCard = "*.*") As Boolean
    '10.3.2002 Active files will return true for 2 different reasons
    '1. There are files in the directory and 1 or many of them are Actively being written to disk
    '2. There are Zero, that means nada, none, no files in the directory at all.
    
    On Error GoTo EH
    Dim sMess As String
    Dim colActiveFiles As Collection
    Dim vActiveFile As Variant
    Dim sActiveFile As String
    Dim lCount As Long
    Dim iFFile As Integer
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    sActiveFile = Dir(psActiveFileDir & "\" & psWildCard, vbNormal)
    
    'If there are no files to check then set to true and bail
    If sActiveFile = vbNullString Then
        ActiveFiles = True
        Exit Function
    End If
    
    For lCount = 1 To 2
        sActiveFile = Dir(psActiveFileDir & "\" & psWildCard, vbNormal)
        
        '1. Add all the existing files in the Active Files collection
        Do Until sActiveFile = vbNullString
            If colActiveFiles Is Nothing Then
                Set colActiveFiles = New Collection
            End If
            colActiveFiles.Add psActiveFileDir & "\" & sActiveFile, psActiveFileDir & "\" & sActiveFile
            sActiveFile = Dir
        Loop
        
        '2. Loop through the collection and open them read lock
        'then close file.  If there is an error while opening, the file
        'is still being written to disk, ie is "Active"
        If Not colActiveFiles Is Nothing Then
            For Each vActiveFile In colActiveFiles
                sActiveFile = vActiveFile
                iFFile = FreeFile
                Open sActiveFile For Binary Access Read Lock Read As #iFFile
                Close #iFFile
            Next
        End If
        
        Set colActiveFiles = Nothing
        'Wait half a second
        DoEvents '
        Sleep 500
    Next
    
    Set colActiveFiles = Nothing

    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    
    ActiveFiles = True
    Set colActiveFiles = Nothing
    If lErrNum <> 70 Then
        Err.Raise lErrNum, , sErrDesc
    End If
End Function

Public Function GetXMLUnitVendorUpdates(poVendorUpdatesRS As ADODB.Recordset, ByRef psClaimId As String) As String
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim sXML As String
    'Claim and Unit
    Dim sUnitId As String
    Dim sClaimId As String
    'Activities
    Dim oActRS As ADODB.Recordset
    Dim sType As String
    Dim sDescription As String
    Dim sComment As String
    Dim sCompletedDate As String
    Dim sCreatedBy As String
    Dim sGMTCreated As String
    'Property
    Dim oPropRS As ADODB.Recordset
    Dim sContactRowId As String
    Dim sPropertyRowId As String
    Dim sPropertyDescription As String
    Dim sVehicleMake As String
    Dim sVehicleModel As String
    Dim sVehicleYear As String
    Dim sInspectedVIN As String
    Dim sLocationType As String
    Dim sLocationName As String
    Dim sLocationAddress As String
    Dim sLocationCity As String
    Dim sLocationCountry As String
    Dim sLocationState As String
    Dim sLocationPostalCode As String
    Dim sLocationPhoneNumber As String
    Dim bExit As Boolean
    Dim sGMTUpdated As String
    
    'First set the update recordsets respectively...
    
    'PropertyUpdates
    Set oPropRS = poVendorUpdatesRS
    'Activity Updates
    Set oActRS = poVendorUpdatesRS.NextRecordset
    
    bExit = True
    'First need to get the highlevel ID
    If Not oPropRS.EOF Then
        bExit = False
        sUnitId = oPropRS.Fields("UnitId").Value
        sClaimId = oPropRS.Fields("ClaimId").Value
    ElseIf Not oActRS.EOF Then
        bExit = False
        sUnitId = oActRS.Fields("UnitId").Value
        sClaimId = oActRS.Fields("ClaimId").Value
    End If
    
    If bExit Then
        GoTo CLEAN_UP
    End If
    sXML = "<?xml version=""1.0"" encoding=""utf-8"" standalone=""yes""?>"
    sXML = sXML & "<VendorUpdates UnitId=""" & sUnitId & """ ClaimId=""" & sClaimId & """>"
        'if there are activity updates then Include them :)
        If Not oActRS.EOF Then
            oActRS.MoveFirst
            sXML = sXML & "<Activities>"
                Do Until oActRS.EOF
                
                    'first set the vars
                    sType = oActRS.Fields("Type").Value
                    sDescription = oActRS.Fields("Description").Value
                    sComment = oActRS.Fields("Comment").Value
                    sCreatedBy = oActRS.Fields("CreatedBy").Value
                    sCompletedDate = oActRS.Fields("CompletedDate").Value
                    sCompletedDate = Format(sCompletedDate, "MM/DD/YYYY HH:mm:ss")
                    sGMTCreated = oActRS.Fields("GMTCreated").Value
                    sGMTCreated = Format(sGMTCreated, "MM/DD/YYYY HH:mm:ss")
                    
                    'Clean the muthers
                    sType = CleanXML(sType)
                    sDescription = CleanXML(sDescription)
                    sComment = CleanXML(sComment)
                    sCreatedBy = CleanXML(sCreatedBy)
                    sCompletedDate = CleanXML(sCompletedDate)
                    sGMTCreated = CleanXML(sGMTCreated)
                    
                    'Then plug em in... plug em in...
                    sXML = sXML & "<Activity>"
                        sXML = sXML & "<Type>" & sType & "</Type>"
                        sXML = sXML & "<Description>" & sDescription & "</Description>"
                        sXML = sXML & "<Comment>" & sComment & "</Comment>"
                        sXML = sXML & "<CreatedBy>" & sCreatedBy & "</CreatedBy>"
                        sXML = sXML & "<CompletedDate>" & sCompletedDate & "</CompletedDate>"
                        sXML = sXML & "<GMTCreated>" & sGMTCreated & "</GMTCreated>"
                    sXML = sXML & "</Activity>"
                    oActRS.MoveNext
                Loop
            sXML = sXML & "</Activities>"
        End If
        'if there are Property Updates then include them too!
        If Not oPropRS.EOF Then
            'This will only have 1 record !!!
            oPropRS.MoveFirst
            'set the vars
            sContactRowId = oPropRS.Fields("ContactRowId").Value
            sPropertyRowId = oPropRS.Fields("PropertyRowId").Value
            sPropertyDescription = oPropRS.Fields("PropertyDescription").Value
            sVehicleMake = oPropRS.Fields("VehicleMake").Value
            sVehicleModel = oPropRS.Fields("VehicleModel").Value
            sVehicleYear = oPropRS.Fields("VehicleYear").Value
            sInspectedVIN = oPropRS.Fields("InspectedVIN").Value
            sLocationType = oPropRS.Fields("LocationType").Value
            sLocationName = oPropRS.Fields("LocationName").Value
            sLocationAddress = oPropRS.Fields("LocationAddress").Value
            sLocationCity = oPropRS.Fields("LocationCity").Value
            sLocationCountry = oPropRS.Fields("LocationCountry").Value
            sLocationState = oPropRS.Fields("LocationState").Value
            sLocationPostalCode = oPropRS.Fields("LocationPostalCode").Value
            sLocationPhoneNumber = oPropRS.Fields("LocationPhoneNumber").Value
            sGMTUpdated = oPropRS.Fields("GMTUpdated").Value
            sGMTUpdated = Format(sGMTUpdated, "MM/DD/YYYY HH:mm:ss")
            
            'Clean the Muthers
            sContactRowId = CleanXML(sContactRowId)
            sPropertyRowId = CleanXML(sPropertyRowId)
            sPropertyDescription = CleanXML(sPropertyDescription)
            sVehicleMake = CleanXML(sVehicleMake)
            sVehicleModel = CleanXML(sVehicleModel)
            sVehicleYear = CleanXML(sVehicleYear)
            sInspectedVIN = CleanXML(sInspectedVIN)
            sLocationType = CleanXML(sLocationType)
            sLocationName = CleanXML(sLocationName)
            sLocationAddress = CleanXML(sLocationAddress)
            sLocationCity = CleanXML(sLocationCity)
            sLocationCountry = CleanXML(sLocationCountry)
            sLocationState = CleanXML(sLocationState)
            sLocationPostalCode = CleanXML(sLocationPostalCode)
            sLocationPhoneNumber = CleanXML(sLocationPhoneNumber)
            sGMTUpdated = CleanXML(sGMTUpdated)
            
            'Then plug em in... plug em in...
            sXML = sXML & "<Property ContactRowId=""" & sContactRowId & """ PropertyRowId=""" & sPropertyRowId & """>"
                sXML = sXML & "<PropertyDescription>" & sPropertyDescription & "</PropertyDescription>"
                sXML = sXML & "<VehicleMake>" & sVehicleMake & "</VehicleMake>"
                sXML = sXML & "<VehicleModel>" & sVehicleModel & "</VehicleModel>"
                sXML = sXML & "<VehicleYear>" & sVehicleYear & "</VehicleYear>"
                sXML = sXML & "<InspectedVIN>" & sInspectedVIN & "</InspectedVIN>"
                sXML = sXML & "<LocationType>" & sLocationType & "</LocationType>"
                sXML = sXML & "<LocationName>" & sLocationName & "</LocationName>"
                sXML = sXML & "<LocationAddress>" & sLocationAddress & "</LocationAddress>"
                sXML = sXML & "<LocationCity>" & sLocationCity & "</LocationCity>"
                sXML = sXML & "<LocationCountry>" & sLocationCountry & "</LocationCountry>"
                sXML = sXML & "<LocationState>" & sLocationState & "</LocationState>"
                sXML = sXML & "<LocationPostalCode>" & sLocationPostalCode & "</LocationPostalCode>"
                sXML = sXML & "<LocationPhoneNumber>" & sLocationPhoneNumber & "</LocationPhoneNumber>"
                sXML = sXML & "<GMTUpdated>" & sGMTUpdated & "</GMTUpdated>"
            sXML = sXML & "</Property>"
        End If
    sXML = sXML & "</VendorUpdates>"

    'Update the passed in ClaimId Param...
    psClaimId = sClaimId
    GetXMLUnitVendorUpdates = sXML
CLEAN_UP:
    'cleanup
    Set oPropRS = Nothing
    Set oActRS = Nothing
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Set oPropRS = Nothing
    Set oActRS = Nothing
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLR_ClassName & vbCrLf & "Public Function GetXMLUnitVendorUpdates"
End Function

Private Function CleanXML(psXML As String) As String
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim sXML As String
    Dim oSer As WDDXSerializer
    
    Set oSer = New WDDXSerializer
    sXML = psXML
    'First get rid of Yucky chars
    sXML = Replace(sXML, Chr(160), Chr(32), , , vbBinaryCompare)
    sXML = Replace(sXML, vbCrLf, "&#13;&#10;", , , vbBinaryCompare)
    sXML = Replace(sXML, vbTab, "&#9;", , , vbBinaryCompare)
    sXML = oSer.serialize(sXML)
   
    sXML = Mid(sXML, InStr(1, sXML, "<string>", vbBinaryCompare) + 8)
    sXML = left(sXML, InStr(1, sXML, "</string>", vbBinaryCompare) - 1)
    sXML = Replace(sXML, "&amp;#", "&#", , , vbBinaryCompare)
    CleanXML = sXML
    
    Set oSer = Nothing
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Set oSer = Nothing
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLR_ClassName & vbCrLf & "Private Function CleanXML"
End Function

Private Function ProcessExportTransactions(psDataPath As String, _
                               psOutPath As String, _
                               poPB As Object, _
                               potxtProgMess As Object, _
                               pbDebugMode As Boolean) As Boolean
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim oAssignmentsIDRS As ADODB.Recordset
    Dim oVendorUpdatesRS As ADODB.Recordset
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    Dim sAssignmentsID As String
    Dim sUnitId As String
    Dim sClaimId As String
    Dim sXMLUnitVendorUpdates As String
    Dim sTickCount As String
    Dim sFile As String
    Dim lCount As Long
    
    Set oConn = New ADODB.Connection
    OpenConnection oConn
    
    'Need to Call SP that will return list of AssignmentsID and UnitId
    'Update Message screen
    potxtProgMess.Text = "Checking for Export Updates..."
    potxtProgMess.Refresh
    Sleep 500
    
    sSQL = "Exec z_spsGetECSFarmersXML01VendorUpdates "
    
    'Use Disconnected Recordset
    Set oAssignmentsIDRS = New ADODB.Recordset
    oAssignmentsIDRS.CursorLocation = adUseClient
    oAssignmentsIDRS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set oAssignmentsIDRS.ActiveConnection = Nothing
    
    If Not oAssignmentsIDRS.EOF Then
        'Update Message screen
        potxtProgMess.Text = oAssignmentsIDRS.RecordCount & " Export Update Items(s) Found!"
        potxtProgMess.Refresh
        Sleep 500
        Do Until oAssignmentsIDRS.EOF
            lCount = lCount + 1
            potxtProgMess.Text = "Processing " & CStr(lCount) & " of " & oAssignmentsIDRS.RecordCount & " Export Record(s)"
            potxtProgMess.Refresh
            Sleep 500
            'Set the ID vars
            sAssignmentsID = oAssignmentsIDRS.Fields("AssignmentsID").Value
            sUnitId = oAssignmentsIDRS.Fields("UnitId").Value
            
            'Now need to call the same SP but pass in the AssignmentsID and The UnitId
            sSQL = "Exec z_spsGetECSFarmersXML01VendorUpdates "
            sSQL = sSQL & sAssignmentsID & ", "     '@AssignmentsID     int=null,
            sSQL = sSQL & "'" & sUnitId & "' "            '@UnitId            varchar(20)=null,
            
            'Can't Use Disconnected Recordset for mult RS Return !!!
            Set oVendorUpdatesRS = New ADODB.Recordset
            oVendorUpdatesRS.CursorLocation = adUseClient
            oVendorUpdatesRS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
            
            'Also pass in the claimId to be populated as well !!!
            sXMLUnitVendorUpdates = GetXMLUnitVendorUpdates(oVendorUpdatesRS, sClaimId)
            
            If sXMLUnitVendorUpdates = vbNullString Then
                GoTo NEXT_Record
            End If
            'Use tickount to help build unique File name...
            sTickCount = goUtil.utGetTickCount
            sFile = "{" & sAssignmentsID & "_" & sUnitId & "_" & sClaimId & "_" & sTickCount & "}.xml"
            'Now that the Unit Vendorupdates xml is populated...
            'Need to save it to proper directory
            potxtProgMess.Text = "Sending Export " & sFile
            potxtProgMess.Refresh
            Sleep 500
            If Not pbDebugMode Then
                goUtil.utSaveFileData psDataPath & "\" & sFile, sXMLUnitVendorUpdates
            End If
            goUtil.utSaveFileData psDataPath & "\BackUp\" & sFile, sXMLUnitVendorUpdates
            
            'Now that the Export was done.... need to update the rows that were just updated
            sSQL = "Exec z_spsGetECSFarmersXML01VendorUpdates "
            sSQL = sSQL & sAssignmentsID & ", "     '@AssignmentsID     int=null,
            sSQL = sSQL & "'" & sUnitId & "', "            '@UnitId            varchar(20)=null,
            sSQL = sSQL & "'" & sClaimId & "', "           '@ClaimId           varchar(20)=null,
            sSQL = sSQL & "0 "                      '@SetSendToCRN      bit=null
            oConn.Execute sSQL
            'Move to the next record
NEXT_Record:
            oAssignmentsIDRS.MoveNext
        Loop
    Else
        potxtProgMess.Text = "Exports not found! " & Now()
        potxtProgMess.Refresh
        Sleep 500
        GoTo CLEAN_UP
    End If
    
    potxtProgMess.Text = "Export Complete! " & Now()
    potxtProgMess.Refresh
    Sleep 500
    
CLEAN_UP:
    'Cleanup
    Set oVendorUpdatesRS = Nothing
    Set oAssignmentsIDRS = Nothing
    Set oConn = Nothing
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Set oVendorUpdatesRS = Nothing
    Set oAssignmentsIDRS = Nothing
    Set oConn = Nothing
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLR_ClassName & vbCrLf & "Private Function ProcessExportTransactions"
                               
End Function


Public Function clsCarLR_ProcessRawData(psDataPath As String, _
                               psOutPath As String, _
                               Optional poPB As Object, _
                               Optional potxtProgMess As Object) As Boolean
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim aryImportTransType(1 To 6) As String
    Dim lCount As Long
    Dim sDataPath As String
    Dim bProcessImportAssignments As Boolean
    
    aryImportTransType(1) = "Import_Assignments"
    aryImportTransType(2) = "Import_Reassignments"
    aryImportTransType(3) = "Import_Updates"
    aryImportTransType(4) = "Import_Closes"
    aryImportTransType(5) = "Export_Updates"
    'This is Farmers Accept, Reject, Delete of Documents we originallly Exported to Farmers
    aryImportTransType(6) = "Export_Documents_Results"
    
    
    For lCount = LBound(aryImportTransType, 1) To UBound(aryImportTransType, 1)
        sDataPath = psDataPath & aryImportTransType(lCount)
        bProcessImportAssignments = ProcessTransTypes(sDataPath, psOutPath, poPB, potxtProgMess)
        If StrComp("Import_Assignments", aryImportTransType(lCount), vbTextCompare) = 0 Then
            clsCarLR_ProcessRawData = bProcessImportAssignments
        End If
    Next
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLR_ClassName & vbCrLf & "Public Function ProcessRawData"
    
End Function

Private Function ProcessTransTypes(psDataPath As String, _
                               psOutPath As String, _
                               Optional poPB As Object, _
                               Optional potxtProgMess As Object) As Boolean
    Dim otxtProgMess As TextBox
    Dim oLR As V2ECKeyBoard.clsCarLR
    Dim colFiles As Collection
    Dim sFile As String
    Dim vFile As Variant
    Dim sData As String
    Dim sPrnData As String 'Will hold each loss PrnData
    Dim varyData As Variant
    Dim varyLoss As Variant
    Dim colLoss As Collection
    Dim vLoss As Variant
    Dim varyLossData As Variant
    Dim sSALN As String 'ClaimNo
    Dim sFact As String 'Fax ID Use to look up against Adjuster for CRID
    Dim sAssignedToFirstName As String
    Dim sAssignedToLastName As String
    Dim sRawFileName As String
    Dim bPB As Boolean 'True if using it
    Dim MyAdjuster As udtAdjuster
    Dim bMissing As Boolean
    Dim sTemp As String
    
    
    
    'Wddx Objects
    Dim lRetryDeserialize As Long
    Dim sErrDescPrefix As String
    Dim sUniqueID As String
    Dim sClaimId As String
    Dim oDeser As WDDXDeserializer
    Dim oMyStruct As WDDXStruct
    Dim sXMLErrorResult As String
    Dim sXMLResultAssignment As String
    Dim sXMLResultUpdate As String
    Dim sXMLResultClose As String
    Dim sXMLResultDocResponse As String
    Dim oAssignmentDetailRS As WDDXRecordset
    Dim sTransType As String
    Dim ThisLoss As XML01LossReport
    Dim vThisLoss As Variant
    Dim lPropNamesCount As Long
    Dim saryPropNames() As Variant
    Dim saryError() As String
    Dim bIsReassignment As Boolean
    Dim sTypeOfReassignment As String
    Dim bClosedPrevUnitReassignment As Boolean
    Dim bIsCloseTransaction As Boolean
    Dim bCheckExports As Boolean
    Dim bDocResponses  As Boolean
    Dim bDebugMode As Boolean
    
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    'SET DEBUG MODE
    'If False then result files will be written
    'If True then Only backup Result files will be written
    bDebugMode = False
    'SET DEBUG MODE
    
    
    If Not goUtil.utFileExists(psDataPath, True) Then
        Exit Function 'BAIL
    Else
        'If this is a Document Response Transaction Set the Flag
        If InStr(1, psDataPath, "Export_Documents_Results", vbTextCompare) > 0 Then
            bDocResponses = True
            'Check for Error Folder for Doc Responses
            If Not goUtil.utFileExists(psDataPath & "\Errors", True) Then
                goUtil.utMakeDir psDataPath & "\Errors"
            End If
            'Check for Results Folder for Doc Responses
            If Not goUtil.utFileExists(psDataPath & "\Results", True) Then
                goUtil.utMakeDir psDataPath & "\Results"
            End If
        End If
        
        'If this is Close Transaction then set the flag
        If InStr(1, psDataPath, "Import_Closes", vbTextCompare) > 0 Then
            bIsCloseTransaction = True
        End If
        
        'check for Backup folder
        If Not goUtil.utFileExists(psDataPath & "\BackUp", True) Then
            goUtil.utMakeDir psDataPath & "\BackUp"
        End If
        
        'Check for RESULT Folder
        'Don't create a results folder for Export_Documents_Results
        If Not bDocResponses Then
            If Not goUtil.utFileExists(psDataPath & "_Results", True) Then
                goUtil.utMakeDir psDataPath & "_Results"
            End If
        
            'Check for RESULT Folder Backup
            If Not goUtil.utFileExists(psDataPath & "_Results\BackUp", True) Then
                goUtil.utMakeDir psDataPath & "_Results\BackUp"
            End If
        End If
    End If
    
    'Set ProgBar flag
    If Not poPB Is Nothing Then
        bPB = True
    End If
    'Set Progress text Mess
    If Not potxtProgMess Is Nothing Then
        If TypeOf potxtProgMess Is TextBox Then
            Set otxtProgMess = potxtProgMess
        End If
    End If
    
    
    'If the process concerns Exporting
    If InStr(1, psDataPath, "Export_", vbTextCompare) > 0 And Not bDocResponses Then
        'Export Updates needs an extra folder for Our errors.
        If Not goUtil.utFileExists(psDataPath & "\Errors", True) Then
            goUtil.utMakeDir psDataPath & "\Errors"
        End If
        'First check registry flag that indicates Exports
        'Somewhere are waiting for this Class!!!
        'This will minimize the need for opening and closing SQL Server
        'Data Connections !  This flag is set under the V2AutoImport Process
        'Private Function VerifyDSNConn()
        bCheckExports = CBool(GetSetting("V2WebControl", "V2ECcarFarmers.clsLossXML01", "Exports", False))
        If bCheckExports Then
            ProcessTransTypes = ProcessExportTransactions(psDataPath, psOutPath, poPB, potxtProgMess, bDebugMode)
            'Now set the Reg setting back to False
            SaveSetting "V2WebControl", "V2ECcarFarmers.clsLossXML01", "Exports", "False"
            If Err.Number <> 0 Then
                sXMLErrorResult = GetXMLResultError("INTERNAL_ERROR", Err.Number, "Error on ProcessExportTransactions: " & Err.Description)
                sFile = "Error_ProcessExportTransactions_" & goUtil.utGetTickCount & ".xml"
                goUtil.utDeleteFile psDataPath & "\Errors\" & sFile
                goUtil.utSaveFileData psDataPath & "\Errors\" & sFile, sXMLErrorResult
            End If
        End If
        GoTo CLEAN_UP
    End If
    
    'Load XML01 format files
    sFile = Dir(psDataPath & "\" & "*.xml")
    
    If sFile <> vbNullString Then
        Set colFiles = New Collection
        Do
            colFiles.Add psDataPath & "\" & sFile
            sFile = Dir
        Loop Until sFile = vbNullString
    Else
        Exit Function 'BAIL
    End If
    
    'Check for Active Files first !
    otxtProgMess.Text = "Checking for Active Files..."
    otxtProgMess.Refresh
    Sleep 500
    If ActiveFiles(psDataPath, "*.xml") Then
        otxtProgMess.Text = "File(s) still active. Exiting..."
        otxtProgMess.Refresh
        Set colFiles = Nothing
        Sleep 500
        otxtProgMess.Text = vbNullString
        otxtProgMess.Refresh
        Exit Function
    Else
        otxtProgMess.Text = vbNullString
        otxtProgMess.Refresh
    End If
    
    'Progress bar
    If bPB Then
        poPB.Max = colFiles.Count
        poPB.Value = 0
        If Not otxtProgMess Is Nothing Then
            otxtProgMess.Text = poPB.Value & " Of " & poPB.Max
            otxtProgMess.Refresh
        End If
    End If
    
    'Loop through the collection of Files
    For Each vFile In colFiles
        'Be sure to free up Windows tasks before processing this
        DoEvents
        Sleep 100
        
        If goUtil.utFileExists(CStr(vFile)) Then
            If Not otxtProgMess Is Nothing Then
                sTemp = CStr(vFile)
                sTemp = Mid(sTemp, InStrRev(sTemp, "\") + 1)
                otxtProgMess.Text = "Loading: " & sTemp
                otxtProgMess.Refresh
            End If
            
            'Load XML01 Multi Print Job data
            sData = goUtil.utGetFileData(CStr(vFile))
            sData = Replace(sData, BOF_UTF_HEADER, vbNullString)
            
            'Check for Document Responses
            If bDocResponses Then
                sTransType = "IA_CRN_DOCRESPONSE"
                On Error Resume Next
                sUniqueID = Mid(sData, InStr(1, sData, "UnitNumber=", vbTextCompare) + Len("UnitNumber=") + 1)
                sUniqueID = left(sUniqueID, InStr(1, sUniqueID, Chr(34), vbBinaryCompare) - 1)
                'If there was an error trying to get the Unique ID
                'need to create Error off the File Name
                sFile = CStr(vFile)
                sFile = Mid(sFile, InStrRev(sFile, "\", , vbBinaryCompare) + 1)
                'Make sure that attribute exists !
                If InStr(1, left(sData, 100), "UnitNumber", vbBinaryCompare) = 0 Then
                    Err.Raise -999, , "UnitNumber is missing!"
                End If
                If Err.Number <> 0 Then
                    sUniqueID = CStr(vFile)
                    sUniqueID = Mid(sUniqueID, InStrRev(sUniqueID, "\", , vbBinaryCompare) + 1)
                    sErrDescPrefix = "Error on attribute ""UnitNumber"", "
                    GoTo ERROR_FILE
                End If
                       
                sXMLResultDocResponse = GetXMLResultDocResponse(sData)
                If Not bDebugMode Then
                    goUtil.utDeleteFile psDataPath & "\Results\" & sFile
                    goUtil.utSaveFileData psDataPath & "\Results\" & sFile, sXMLResultDocResponse
                End If
                GoTo NEXT_FILE
            End If
            
            sPrnData = sData
            'If the Deserialize does not work need to create Error Result
            'First need to get the Unique ID out
            On Error Resume Next
            sUniqueID = Mid(sData, InStr(1, sData, "UniqueID=", vbTextCompare) + Len("UniqueID=") + 1)
            sUniqueID = left(sUniqueID, InStr(1, sUniqueID, Chr(34), vbBinaryCompare) - 1)
            'If there was an error trying to get the Unique ID
            'need to create Error off the File Name
            sFile = CStr(vFile)
            sFile = Mid(sFile, InStrRev(sFile, "\", , vbBinaryCompare) + 1)
            'Make sure that attribute exists !
            If InStr(1, left(sData, 100), "UniqueID", vbBinaryCompare) = 0 Then
                Err.Raise -999, , "UniqueID is missing!"
            End If
            If Err.Number <> 0 Then
                sUniqueID = CStr(vFile)
                sUniqueID = Mid(sUniqueID, InStrRev(sUniqueID, "\", , vbBinaryCompare) + 1)
                sErrDescPrefix = "Error on attribute ""UniqueID"", "
                GoTo ERROR_FILE
            End If
            
            If bIsCloseTransaction Then
                'close trasaction does not have a claimID attribute
                GoTo SKIP_CLAIMID
            End If
            'Then Get the ClaimID
            On Error Resume Next
            sClaimId = Mid(sData, InStr(1, sData, "ClaimId=", vbTextCompare) + Len("ClaimId=") + 1)
            sClaimId = left(sClaimId, InStr(1, sClaimId, Chr(34), vbBinaryCompare) - 1)
            'If there was an error trying to get the ClaimID
            'need to create Error off the File Name
            'Make sure that attribute exists !
            If InStr(1, left(sData, 100), "ClaimId=", vbBinaryCompare) = 0 Then
                Err.Raise -999, , "ClaimId is missing!"
            End If
            If Err.Number <> 0 Then
                sClaimId = CStr(vFile)
                sClaimId = Mid(sClaimId, InStrRev(sClaimId, "\", , vbBinaryCompare) + 1)
                sErrDescPrefix = "Error on attribute ""ClaimId"", "
                GoTo ERROR_FILE
            End If
            'set the member variable
SKIP_CLAIMID:
            msUniqueID = sUniqueID
            If Not bIsCloseTransaction Then
                msClaimId = sClaimId
            End If
            
            lRetryDeserialize = 0
RETRY_DESERIALIZE:
            Set oDeser = Nothing
            Set oDeser = New WDDXDeserializer
            Set oMyStruct = oDeser.deserialize(sData)
            If Err.Number <> 0 Then
                'If there is an error here... let the system
                'do events and try again up to 3 times
                lRetryDeserialize = lRetryDeserialize + 1
                If lRetryDeserialize <= 3 Then
                    On Error GoTo 0 'Disables any enabled error handler in the current procedure
                    On Error Resume Next
                    DoEvents
                    Sleep 1000
                    GoTo RETRY_DESERIALIZE
                End If
                sErrDescPrefix = "Wddx Error on deserialize: "
                GoTo ERROR_FILE
            End If
            
            'Check to see if this Transaction Item is of the same class
            'if it is not then need to error and skip it.
            sTemp = oMyStruct.getProp("ClassName")
            
            If Err.Number <> 0 Then
                sXMLErrorResult = GetXMLResultError(sUniqueID, Err.Number, "Wddx Error on ""ClassName"": " & Err.Description)
                sErrDescPrefix = ""
                GoTo ERROR_FILE
            End If

            If StrComp(sTemp, clsCarLR_ClassName, vbTextCompare) <> 0 Then
                sXMLErrorResult = GetXMLResultError(sUniqueID, "Invalid ClassName", sTemp & " is not the correct ClassName.  Expecting ClassName=""" & clsCarLR_ClassName & """")
                sErrDescPrefix = ""
                GoTo ERROR_FILE
            End If

            'Add Class as prefix
            sData = clsCarLR_ClassName & vbCrLf
            
            'Validate the Transaction Type
            sTemp = oMyStruct.getProp("TransType")
            If Err.Number <> 0 Then
                sXMLErrorResult = GetXMLResultError(sUniqueID, Err.Number, "Wddx Error on ""TransType"": " & Err.Description)
                sErrDescPrefix = ""
                GoTo ERROR_FILE
            End If
            bIsReassignment = False
            sTypeOfReassignment = vbNullString
            bClosedPrevUnitReassignment = False
            If StrComp(sTemp, "IA_CRN_ASSIGN", vbTextCompare) = 0 Then
                sTransType = "IA_CRN_ASSIGN"
                If InStr(1, psDataPath, "Import_Reassignments", vbTextCompare) > 0 Then
                    bIsReassignment = True
                End If
            ElseIf StrComp(sTemp, "IA_CRN_UPDATE", vbTextCompare) = 0 Then
                sTransType = "IA_CRN_UPDATE"
                sXMLResultUpdate = GetXMLResultUpdate(sUniqueID, UpdateUnit(sPrnData))
                If Err.Number <> 0 Then
                    sXMLErrorResult = GetXMLResultError(sUniqueID, Err.Number, "Error on UpdateUnit: " & Err.Description)
                    sErrDescPrefix = ""
                    GoTo ERROR_FILE
                End If
                If Not bDebugMode Then
                    goUtil.utDeleteFile psDataPath & "_Results\" & sFile
                    goUtil.utSaveFileData psDataPath & "_Results\" & sFile, sXMLResultUpdate
                End If
                goUtil.utDeleteFile psDataPath & "_Results\BackUp\" & sFile
                goUtil.utSaveFileData psDataPath & "_Results\BackUp\" & sFile, sXMLResultUpdate
                GoTo NEXT_FILE
            ElseIf StrComp(sTemp, "IA_CRN_CLOSE", vbTextCompare) = 0 Then
                sTransType = "IA_CRN_CLOSE"
                sXMLResultClose = GetXMLResultClose(sUniqueID, CloseUnit(sUniqueID))
                If Err.Number <> 0 Then
                    sErrDescPrefix = "Error on TransType IA_CRN_CLOSE "
                    GoTo ERROR_FILE
                End If
                If Not bDebugMode Then
                    goUtil.utDeleteFile psDataPath & "_Results\" & sFile
                    goUtil.utSaveFileData psDataPath & "_Results\" & sFile, sXMLResultClose
                End If
                goUtil.utDeleteFile psDataPath & "_Results\BackUp\" & sFile
                goUtil.utSaveFileData psDataPath & "_Results\BackUp\" & sFile, sXMLResultClose
                GoTo NEXT_FILE
            Else
                'If the TransType Is Not found then give error Result
                sXMLErrorResult = GetXMLResultError(sUniqueID, "Invalid TransType", sTemp & " is not defined!")
                sErrDescPrefix = ""
                GoTo ERROR_FILE
            End If
            
            'Check to See if this a Property Or Auto Loss
            sTemp = oMyStruct.getProp("LossType")
            If Err.Number <> 0 Then
                sXMLErrorResult = GetXMLResultError(sUniqueID, Err.Number, "Wddx Error on ""LossType"": " & Err.Description)
                sErrDescPrefix = ""
                GoTo ERROR_FILE
            End If
            
            If StrComp(sTemp, "Property", vbTextCompare) = 0 Then
                clsCarLR_LossType = XML01Pro
            ElseIf StrComp(sTemp, "Homeowners", vbTextCompare) = 0 Then
                clsCarLR_LossType = XML01Pro
            ElseIf StrComp(sTemp, "Auto", vbTextCompare) = 0 Then
                clsCarLR_LossType = XML01Apd
            ElseIf StrComp(sTemp, "Commercial Auto", vbTextCompare) = 0 Then
                clsCarLR_LossType = XML01Apd
            Else
                'If the LossType Is Not found then give error Result
                sXMLErrorResult = GetXMLResultError(sUniqueID, "Invalid Loss Type", sTemp & " is not defined!")
                sErrDescPrefix = ""
                GoTo ERROR_FILE
            End If
            
            'Validate the Assignments Details Section.
            Set oAssignmentDetailRS = oMyStruct.getProp("AssignmentDetailRS")
            If Err.Number <> 0 Then
                sXMLErrorResult = GetXMLResultError(sUniqueID, Err.Number, "Wddx Error on ""AssignmentDetailRS"": " & Err.Description)
                sErrDescPrefix = ""
                GoTo ERROR_FILE
            End If
            
            'FACT or FAX id same as AssignedTo
            sTemp = oAssignmentDetailRS.getField(1, "AssignedTo")
            If Err.Number <> 0 Then
                sXMLErrorResult = GetXMLResultError(sUniqueID, Err.Number, "Wddx Error on AssignmentDetailRS.getField ""AssignedTo"": " & Err.Description)
                sErrDescPrefix = ""
                GoTo ERROR_FILE
            End If
            sFact = sTemp
            
            'Farmers version data FirstName Of Adjuster
            sTemp = oAssignmentDetailRS.getField(1, "AssignedToFirstName")
            If Err.Number <> 0 Then
                sXMLErrorResult = GetXMLResultError(sUniqueID, Err.Number, "Wddx Error on AssignmentDetailRS.getField ""AssignedToFirstName"": " & Err.Description)
                sErrDescPrefix = ""
                GoTo ERROR_FILE
            End If
            sAssignedToFirstName = sTemp
            
            'Farmers version data Lastname Of Adjuster
            sTemp = oAssignmentDetailRS.getField(1, "AssignedToLastName")
            If Err.Number <> 0 Then
                sXMLErrorResult = GetXMLResultError(sUniqueID, Err.Number, "Wddx Error on AssignmentDetailRS.getField ""AssignedToLastName"": " & Err.Description)
                sErrDescPrefix = ""
                GoTo ERROR_FILE
            End If
            sAssignedToLastName = sTemp
            
            'Check for CRN Reassignment
            'The CRN Unit Reassignment means 1 of 2 possibilities
            '1. This is a NEW Unit Assignment being assigned to our Company that was
            'originally assigned to an outside IA Firm.
            '2. This is a Close Unit and NEW Assignment from One of Our Adjusters to Another
            'one of Our Adjusters.
            'In Any event, The Close Unit Function must be Called to Close any existence of
            'this Unit that may previously exist.
            If bIsReassignment Then
                sTemp = sFact & " (" & sAssignedToFirstName & " " & sAssignedToLastName & ")"
                bClosedPrevUnitReassignment = CloseUnit(sUniqueID, bIsReassignment, sTemp)
                If Err.Number <> 0 Then
                    'Need to Trap for No Records Affected
                    sXMLErrorResult = GetXMLResultError(sUniqueID, Err.Number, "Error on Closing Previous Unit for Reassignment: " & Err.Description)
                    sErrDescPrefix = ""
                    GoTo ERROR_FILE
                End If
                If bClosedPrevUnitReassignment Then
                    sTypeOfReassignment = "Internal IA Reassignment"
                Else
                    sTypeOfReassignment = "External IA to Internal IA Reassignment"
                End If
            End If
            
            'Claim Number is Claim Segment Number in CRN = Unit number - #
            '5.24.2005 BGS If this is an Auto Claim... Use the UnitNumber as  Claim number !!!
            'This will create a separate Assignment if there are multiple units assigned under
            'the same claim in CRN System... which is exactly how we want to handle them!
            If clsCarLR_LossType = XML01Apd Then
                sTemp = oAssignmentDetailRS.getField(1, "UnitNumber")
            Else
'                'Property Claims are a different story since we want to "merge" separate units
'                'into the same assignment.
'                 sTemp = oAssignmentDetailRS.getField(1, "ClaimNumber")
                'BGS 6.15.2005 Need to Handle the HomeOwners just like the Auto
                'Since it makes it impossible to return Document Unique Ids when Merging UNits
                'into a sinlge assignment.  This will have some significant front end user
                'issues, such as how to apply deductibles accross multiple units ???
                'Deductibles will have to be manually verified by the adjusters!!!
                sTemp = oAssignmentDetailRS.getField(1, "UnitNumber")
            End If
           
            If Err.Number <> 0 Then
                sXMLErrorResult = GetXMLResultError(sUniqueID, Err.Number, "Wddx Error on AssignmentDetailRS.getField ""ClaimNumber"": " & Err.Description)
                sErrDescPrefix = ""
                GoTo ERROR_FILE
            End If
            'Need to Strip the "-" dashes out for Easy Claim Billing Sake.
            sSALN = Replace(sTemp, "-", vbNullString, , , vbBinaryCompare)
            'Farmers uses FACT ID for XML01 Property Reports which is THE ACID (Adjuster Client ID)
            msACID = sFact
            msCLIENTNUM = sSALN
            
            'Now update the actual Wddx RS ClaimNumber
            oAssignmentDetailRS.setField 1, "ClaimNumber", msCLIENTNUM
            
            'Get the Available Property Names into a comma separated List
            
            saryPropNames() = oMyStruct.getPropNames
            
            Set mudtXML01LossReport.XML01Loss.AssignmentDetailRS = Nothing
            Set mudtXML01LossReport.XML01Loss.ActivitiesRS = Nothing
            Set mudtXML01LossReport.XML01Loss.ContactDetailRS = Nothing
            Set mudtXML01LossReport.XML01Loss.ContactsRS = Nothing
            Set mudtXML01LossReport.XML01Loss.CoverageRS = Nothing
            Set mudtXML01LossReport.XML01Loss.EndorsementRS = Nothing
            Set mudtXML01LossReport.XML01Loss.LossDetailRS = Nothing
            Set mudtXML01LossReport.XML01Loss.PaymentDetailRS = Nothing
            Set mudtXML01LossReport.XML01Loss.PriorLossDetailRS = Nothing
            Set mudtXML01LossReport.XML01Loss.PolicyDetailRS = Nothing
            Set mudtXML01LossReport.XML01Loss.AddressRS = Nothing
            Set mudtXML01LossReport.XML01Loss.VehicleDetailRS = Nothing
            
            For lPropNamesCount = LBound(saryPropNames, 1) To UBound(saryPropNames, 1)
                If StrComp(saryPropNames(lPropNamesCount), "AssignmentDetailRS", vbTextCompare) = 0 Then
                    Set mudtXML01LossReport.XML01Loss.AssignmentDetailRS = oMyStruct.getProp(saryPropNames(lPropNamesCount))
                ElseIf StrComp(saryPropNames(lPropNamesCount), "ActivitiesRS", vbTextCompare) = 0 Then
                    Set mudtXML01LossReport.XML01Loss.ActivitiesRS = oMyStruct.getProp(saryPropNames(lPropNamesCount))
                ElseIf StrComp(saryPropNames(lPropNamesCount), "ContactDetailsRS", vbTextCompare) = 0 Then
                    Set mudtXML01LossReport.XML01Loss.ContactDetailRS = oMyStruct.getProp(saryPropNames(lPropNamesCount))
                ElseIf StrComp(saryPropNames(lPropNamesCount), "ContactRS", vbTextCompare) = 0 Then
                    Set mudtXML01LossReport.XML01Loss.ContactsRS = oMyStruct.getProp(saryPropNames(lPropNamesCount))
                ElseIf StrComp(saryPropNames(lPropNamesCount), "CoverageRS", vbTextCompare) = 0 Then
                    Set mudtXML01LossReport.XML01Loss.CoverageRS = oMyStruct.getProp(saryPropNames(lPropNamesCount))
                ElseIf StrComp(saryPropNames(lPropNamesCount), "EndorsementRS", vbTextCompare) = 0 Then
                    Set mudtXML01LossReport.XML01Loss.EndorsementRS = oMyStruct.getProp(saryPropNames(lPropNamesCount))
                ElseIf StrComp(saryPropNames(lPropNamesCount), "LossDetailRS", vbTextCompare) = 0 Then
                    Set mudtXML01LossReport.XML01Loss.LossDetailRS = oMyStruct.getProp(saryPropNames(lPropNamesCount))
                ElseIf StrComp(saryPropNames(lPropNamesCount), "PaymentDetailRS", vbTextCompare) = 0 Then
                    Set mudtXML01LossReport.XML01Loss.PaymentDetailRS = oMyStruct.getProp(saryPropNames(lPropNamesCount))
                ElseIf StrComp(saryPropNames(lPropNamesCount), "PriorLossDetailRS", vbTextCompare) = 0 Then
                    Set mudtXML01LossReport.XML01Loss.PriorLossDetailRS = oMyStruct.getProp(saryPropNames(lPropNamesCount))
                ElseIf StrComp(saryPropNames(lPropNamesCount), "PolicyDetailRS", vbTextCompare) = 0 Then
                    Set mudtXML01LossReport.XML01Loss.PolicyDetailRS = oMyStruct.getProp(saryPropNames(lPropNamesCount))
                ElseIf StrComp(saryPropNames(lPropNamesCount), "AddressRS", vbTextCompare) = 0 Then
                    Set mudtXML01LossReport.XML01Loss.AddressRS = oMyStruct.getProp(saryPropNames(lPropNamesCount))
                ElseIf StrComp(saryPropNames(lPropNamesCount), "VehicleDetailRS", vbTextCompare) = 0 Then
                    Set mudtXML01LossReport.XML01Loss.VehicleDetailRS = oMyStruct.getProp(saryPropNames(lPropNamesCount))
                End If
                If Err.Number <> 0 Then
                    sErrDescPrefix = "Wddx Error on Struct.getProp """ & saryPropNames(lPropNamesCount) & """: "
                    GoTo ERROR_FILE
                End If
            Next
            
            'Now that we have everything we can add the loss object
            Set oLR = Me
            
            With oLR
                'Set the Various Wddx RS in the Loss Report if they Exist
                Set .LRs = moLRs
                'Remove the class marker frm the prn data
                sPrnData = Replace(sPrnData, clsCarLR_ClassName & vbCrLf, vbNullString, 1)
                .PrnData = sPrnData
                .PopulateMe
                If Err.Number <> 0 Then
                    sErrDescPrefix = "Wddx Error on PopulateMe "
                    GoTo ERROR_FILE
                End If
            End With
            
            
            moLRs.AddRawLossReport oLR
            sXMLErrorResult = vbNullString
            If InStr(1, oLR.Status, "|", vbTextCompare) > 0 Then
                saryError = Split(oLR.Status, "|")
                'Check for certain types of errors
                'If Duplicate loss Report, actually reject the Transaction instead of Error
                If InStr(1, oLR.Status, "Duplicate Unit Assignment to the same adjuster not allowed! UniqueID = ", vbTextCompare) > 0 Then
                    sTemp = saryError(1) & " AssignedTo = " & sFact
                    sXMLResultAssignment = GetXMLResultAssignment(sUniqueID, False, sTemp)
                    If Not bDebugMode Then
                        goUtil.utDeleteFile psDataPath & "_Results\" & sFile
                        goUtil.utSaveFileData psDataPath & "_Results\" & sFile, sXMLResultAssignment
                    End If
                    goUtil.utDeleteFile psDataPath & "_Results\BackUp\" & sFile
                    goUtil.utSaveFileData psDataPath & "_Results\BackUp\" & sFile, sXMLResultAssignment
                ElseIf InStr(1, oLR.Status, ", User is not defined!", vbTextCompare) > 0 Then
                    'If the Adjuster User Is not defined, then send False Accept Message to Farmers
                    sXMLResultAssignment = GetXMLResultAssignment(sUniqueID, False, saryError(1))
                    If Not bDebugMode Then
                        goUtil.utDeleteFile psDataPath & "_Results\" & sFile
                        goUtil.utSaveFileData psDataPath & "_Results\" & sFile, sXMLResultAssignment
                    End If
                    goUtil.utDeleteFile psDataPath & "_Results\BackUp\" & sFile
                    goUtil.utSaveFileData psDataPath & "_Results\BackUp\" & sFile, sXMLResultAssignment
                ElseIf InStr(1, oLR.Status, " AssignedTo is blank! ", vbTextCompare) > 0 Then
                    'If the AssignedTo they sent us is blank, then send False Accept Message to Farmers
                    sXMLResultAssignment = GetXMLResultAssignment(sUniqueID, False, saryError(1))
                    If Not bDebugMode Then
                        goUtil.utDeleteFile psDataPath & "_Results\" & sFile
                        goUtil.utSaveFileData psDataPath & "_Results\" & sFile, sXMLResultAssignment
                    End If
                    goUtil.utDeleteFile psDataPath & "_Results\BackUp\" & sFile
                    goUtil.utSaveFileData psDataPath & "_Results\BackUp\" & sFile, sXMLResultAssignment
                ElseIf InStr(1, oLR.Status, " CatastropheCode is blank! ", vbTextCompare) > 0 Then
                    'If the Cat Code they sent us is blank, then send False Accept Message to Farmers
                    sXMLResultAssignment = GetXMLResultAssignment(sUniqueID, False, saryError(1))
                    If Not bDebugMode Then
                        goUtil.utDeleteFile psDataPath & "_Results\" & sFile
                        goUtil.utSaveFileData psDataPath & "_Results\" & sFile, sXMLResultAssignment
                    End If
                    goUtil.utDeleteFile psDataPath & "_Results\BackUp\" & sFile
                    goUtil.utSaveFileData psDataPath & "_Results\BackUp\" & sFile, sXMLResultAssignment
                ElseIf InStr(1, oLR.Status, "Command text was not set for the command object", vbTextCompare) > 0 Then
                    'Catch Error for Multi Unit Update still under development
                    sXMLResultAssignment = GetXMLResultAssignment(sUniqueID, True)
                    If Not bDebugMode Then
                        goUtil.utDeleteFile psDataPath & "_Results\" & sFile
                        goUtil.utSaveFileData psDataPath & "_Results\" & sFile, sXMLResultAssignment
                    End If
                    goUtil.utDeleteFile psDataPath & "_Results\BackUp\" & sFile
                    goUtil.utSaveFileData psDataPath & "_Results\BackUp\" & sFile, sXMLResultAssignment
                    'end Catch Error for Multi Unit Update still under development
                Else
                    sXMLErrorResult = GetXMLResultError(sUniqueID, saryError(0), saryError(1))
                    sErrDescPrefix = ""
                    GoTo ERROR_FILE
                End If
            Else
                sXMLResultAssignment = GetXMLResultAssignment(sUniqueID, True)
                If Not bDebugMode Then
                    goUtil.utDeleteFile psDataPath & "_Results\" & sFile
                    goUtil.utSaveFileData psDataPath & "_Results\" & sFile, sXMLResultAssignment
                End If
                goUtil.utDeleteFile psDataPath & "_Results\BackUp\" & sFile
                goUtil.utSaveFileData psDataPath & "_Results\BackUp\" & sFile, sXMLResultAssignment
            End If
            
            'Check to see if Process of Raw Data was Aborted...
            If oLR.AbortProcessRawData Then
                GoTo CLEAN_UP
            End If
            
            'Free local memory
            Set oLR = Nothing
            
            If bPB Then
                poPB.Value = poPB.Value + 1
                If Not otxtProgMess Is Nothing Then
                    otxtProgMess.Text = poPB.Value & " Of " & poPB.Max
                    otxtProgMess.Refresh
                End If
            End If

        End If
        If Err.Number = 0 Then
            GoTo NEXT_FILE
        End If
ERROR_FILE:
        If sXMLErrorResult = vbNullString Then
            sXMLErrorResult = GetXMLResultError(sUniqueID, Err.Number, sErrDescPrefix & Err.Description)
        End If
        If Not bDebugMode Then
            If Not bDocResponses Then
                goUtil.utDeleteFile psDataPath & "_Results\" & sFile
                goUtil.utSaveFileData psDataPath & "_Results\" & sFile, sXMLErrorResult
            Else
                'Document Responses
                goUtil.utDeleteFile psDataPath & "\Results\" & sFile
                goUtil.utSaveFileData psDataPath & "\Results\" & sFile, sXMLErrorResult
            End If
        End If
        If Not bDocResponses Then
            goUtil.utDeleteFile psDataPath & "_Results\BackUp\" & sFile
            goUtil.utSaveFileData psDataPath & "_Results\BackUp\" & sFile, sXMLErrorResult
        End If
        On Error GoTo EH
NEXT_FILE:
    If goUtil.utFileExists(CStr(vFile)) Then
        'Make a Backup copy of the Raw File
        'Check to see if the Raw File is already in the
        'BackUp folder first
        sFile = CStr(vFile)
        sFile = Mid(sFile, InStrRev(sFile, "\") + 1)
        sRawFileName = psDataPath & "\BackUp\" & sFile
        
        If goUtil.utFileExists(sRawFileName) Then
            SetAttr CStr(sRawFileName), vbNormal
            goUtil.utDeleteFile CStr(sRawFileName)
        End If
        goUtil.utCopyFile CStr(vFile), sRawFileName
        'Get rid of the original Raw File since it is copied into
        'the backup folder
        SetAttr CStr(vFile), vbNormal
        goUtil.utDeleteFile CStr(vFile)
    End If
    Set oDeser = Nothing
    Set oMyStruct = Nothing
    Next
    
    'Inform message that processing is done
    otxtProgMess.Text = "Process Completed " & Now()
    otxtProgMess.Refresh
CLEAN_UP:
    'CleanUp
    Set otxtProgMess = Nothing
    Set colFiles = Nothing
    Set colLoss = Nothing
    Set oLR = Nothing
    Set oDeser = Nothing
    Set oMyStruct = Nothing
    
    ProcessTransTypes = True
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Set colFiles = Nothing
    Set colLoss = Nothing
    Set oLR = Nothing
    sXMLErrorResult = GetXMLResultError(sUniqueID, CStr(lErrNum), sErrDesc)
    If Not bDebugMode Then
        If Not bDocResponses Then
            goUtil.utDeleteFile psDataPath & "_Results\" & sFile
            goUtil.utSaveFileData psDataPath & "_Results\" & sFile, sXMLErrorResult
        Else
            'Document Resposnes
            goUtil.utDeleteFile psDataPath & "\Results\" & sFile
            goUtil.utSaveFileData psDataPath & "\Results\" & sFile, sXMLErrorResult
        End If
    End If
    If Not bDocResponses Then
        goUtil.utDeleteFile psDataPath & "_Results\BackUp\" & sFile
        goUtil.utSaveFileData psDataPath & "_Results\BackUp\" & sFile, sXMLErrorResult
    End If
End Function

Public Function clsCarLR_AdditmX(poitmX As Object, polvw As Object) As Boolean
    On Error GoTo EH
    Dim sName As String
    
    Set poitmX = polvw.ListItems.Add(, , Format(msDateAssign, "MM/DD/YY"), , LRPic.lrFORMAT02)
    
    poitmX.SubItems(LossReports.DateAsgnSort - 1) = Format(msDateAssign, "YYYY/MM/DD")
    
    'AssignmentType
    poitmX.SubItems(LossReports.AssignmentType - 1) = msAssignmentType
    
    'Status
    poitmX.SubItems(LossReports.Status - 1) = msStatus
    
    'Cat Name
    poitmX.SubItems(LossReports.CatName - 1) = msCatName
    
    'Cat Code
    poitmX.SubItems(LossReports.CatCode - 1) = msCatCode
    
    'ACID
    poitmX.SubItems(LossReports.ACID - 1) = msACID
    
    'Adjuster
    poitmX.SubItems(LossReports.Adjuster - 1) = msAdjuster
    
    'CLIENTNUM
    poitmX.SubItems(LossReports.CLIENTNUM - 1) = msCLIENTNUM
    
    'IBNUM
    poitmX.SubItems(LossReports.IBNUM - 1) = msIBNUM
    
    'Insured name
    'Use MainFileInsuredName if possible use Insured name if Mainfile is null string
    poitmX.SubItems(LossReports.InsuredName - 1) = msInsuredName
    
    'Home Phone
    poitmX.SubItems(LossReports.HPhone - 1) = msHomePhone
    
    'Work Phone
    poitmX.SubItems(LossReports.WPhone - 1) = msWorkPhone
    
    'Format use Class Name
    poitmX.SubItems(LossReports.RFormat - 1) = clsCarLR_ClassName
    
    'Sort by Date Assigned and Ajuster and Insured Name
    'Format the Date so that it will sort by year first
    poitmX.SubItems(LossReports.RSort - 1) = Format(msDateAssign, "YY/MM/DD") & clsCarLR_PrnKey
    
    'Rememeber that the File path is used as the Key.
    'This allows for quick reference to this particular Report
    'when it is selcted from the listview
    poitmX.SubItems(LossReports.RKey - 1) = clsCarLR_PrnKey
    
    poitmX.Selected = False
    
    clsCarLR_AdditmX = True
    Exit Function
EH:
    Err.Clear
    
End Function

Public Function clsCarLR_PrintMe(psPrnDeviceName As String, plFormat As Long, pbPreview As Boolean, pbPDF As Boolean, _
                        pofrmIncludeDocs As Object, plHwnd As Long, plChainType As Long) As Boolean
    On Error GoTo EH
    Dim oChk As Object
    Dim bFireEvent As Boolean
    Dim PrnFormat As PrintFormat
    Dim MyAppDoc As V2ECKeyBoard.udtAppDoc
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    PrnFormat = plFormat
    'Print Loss Report
    
    'Check which format we will be using
    Select Case PrnFormat
        Case PrintFormat.RawText
            moLRs.PrintToSpool clsCarLR_PrnKey, mudtXML01LossReport.PrnData
        'Translation we will be using Active Reports
        Case PrintFormat.Translated
            PrintXML01 psPrnDeviceName, pbPreview, pbPDF, plHwnd, plChainType
    End Select
    
    'We will print Appending documents After the Loss report
    'if they are selected
    If moLRs.AppDocFlag And plChainType = ChainType.NotChain And Not pbPreview And Not pbPDF Then
        For Each oChk In pofrmIncludeDocs.Controls
            If TypeOf oChk Is CheckBox Then
                If oChk.Enabled And oChk.Tag = "AppDoc" Then
                    If oChk.Value = vbChecked Then
                        'If we have at least one item checked then set this flag
                        bFireEvent = True
                        MyAppDoc.DocName = oChk.Caption
                        MyAppDoc.Selected = True
                        moLRs.AppDocsCol.Remove MyAppDoc.DocName
                        moLRs.AppDocsCol.Add MyAppDoc, MyAppDoc.DocName
                    End If
                End If
            End If
        Next
        If bFireEvent Then
            moLRs.FirePrintAppDoc mudtXML01LossReport.XML01Loss.AssignmentDetailRS.getField(1, "ClaimNumber")
        End If
    End If
    clsCarLR_PrintMe = True
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
   Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLR_ClassName & vbCrLf & "PrintMe"
End Function


    
Public Function clsCarLR_CLEANUP() As Boolean
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    If Not moLRs Is Nothing Then
        Set moLRs = Nothing
    End If
    clsCarLR_CLEANUP = True
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLR_ClassName & vbCrLf & "Public Function CleanUp"
End Function

Public Sub clsCarLR_PopulateMe()
    'Here is where we populate the Main udt for the XML01 Format
    On Error GoTo EH
    Dim oDeser As WDDXDeserializer
    Dim oMyStruct As WDDXStruct
    Dim saryPropNames() As Variant
    Dim lPropNamesCount As Long
    Dim sTemp As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim sUniqueID As String
    Dim sClaimId As String
    
    If mudtXML01LossReport.XML01Loss.AssignmentDetailRS Is Nothing Then
        
        sUniqueID = Mid(clsCarLR_PrnData, InStr(1, clsCarLR_PrnData, "UniqueID=", vbTextCompare) + Len("UniqueID=") + 1)
        sUniqueID = left(sUniqueID, InStr(1, sUniqueID, Chr(34), vbBinaryCompare) - 1)
        
        If InStr(1, left(clsCarLR_PrnData, 100), "UniqueID", vbBinaryCompare) = 0 Then
            Err.Raise -999, , "UniqueID is missing!"
        End If
        msUniqueID = sUniqueID
        
        sClaimId = Mid(clsCarLR_PrnData, InStr(1, clsCarLR_PrnData, "ClaimId=", vbTextCompare) + Len("ClaimId=") + 1)
        sClaimId = left(sClaimId, InStr(1, sClaimId, Chr(34), vbBinaryCompare) - 1)
        
        If InStr(1, left(clsCarLR_PrnData, 100), "ClaimId", vbBinaryCompare) = 0 Then
            Err.Raise -999, , "ClaimId is missing!"
        End If
        
        msClaimId = sClaimId
        
        Set oDeser = New WDDXDeserializer
        Set oMyStruct = oDeser.deserialize(clsCarLR_PrnData)
        'Check to See if this a Property Or Auto Loss
        sTemp = Trim(oMyStruct.getProp("LossType"))
        
        If StrComp(sTemp, "Property", vbTextCompare) = 0 Then
            clsCarLR_LossType = XML01Pro
        ElseIf StrComp(sTemp, "Homeowners", vbTextCompare) = 0 Then
            clsCarLR_LossType = XML01Pro
        ElseIf StrComp(sTemp, "Auto", vbTextCompare) = 0 Then
            clsCarLR_LossType = XML01Apd
        End If
        
        saryPropNames() = oMyStruct.getPropNames
        
        Set mudtXML01LossReport.XML01Loss.AssignmentDetailRS = Nothing
        Set mudtXML01LossReport.XML01Loss.ActivitiesRS = Nothing
        Set mudtXML01LossReport.XML01Loss.ContactDetailRS = Nothing
        Set mudtXML01LossReport.XML01Loss.ContactsRS = Nothing
        Set mudtXML01LossReport.XML01Loss.CoverageRS = Nothing
        Set mudtXML01LossReport.XML01Loss.EndorsementRS = Nothing
        Set mudtXML01LossReport.XML01Loss.LossDetailRS = Nothing
        Set mudtXML01LossReport.XML01Loss.PaymentDetailRS = Nothing
        Set mudtXML01LossReport.XML01Loss.PriorLossDetailRS = Nothing
        Set mudtXML01LossReport.XML01Loss.PolicyDetailRS = Nothing
        Set mudtXML01LossReport.XML01Loss.AddressRS = Nothing
        Set mudtXML01LossReport.XML01Loss.VehicleDetailRS = Nothing
        
        For lPropNamesCount = LBound(saryPropNames, 1) To UBound(saryPropNames, 1)
            If StrComp(saryPropNames(lPropNamesCount), "AssignmentDetailRS", vbTextCompare) = 0 Then
                Set mudtXML01LossReport.XML01Loss.AssignmentDetailRS = oMyStruct.getProp(saryPropNames(lPropNamesCount))
            ElseIf StrComp(saryPropNames(lPropNamesCount), "ActivitiesRS", vbTextCompare) = 0 Then
                Set mudtXML01LossReport.XML01Loss.ActivitiesRS = oMyStruct.getProp(saryPropNames(lPropNamesCount))
            ElseIf StrComp(saryPropNames(lPropNamesCount), "ContactDetailsRS", vbTextCompare) = 0 Then
                Set mudtXML01LossReport.XML01Loss.ContactDetailRS = oMyStruct.getProp(saryPropNames(lPropNamesCount))
            ElseIf StrComp(saryPropNames(lPropNamesCount), "ContactRS", vbTextCompare) = 0 Then
                Set mudtXML01LossReport.XML01Loss.ContactsRS = oMyStruct.getProp(saryPropNames(lPropNamesCount))
            ElseIf StrComp(saryPropNames(lPropNamesCount), "CoverageRS", vbTextCompare) = 0 Then
                Set mudtXML01LossReport.XML01Loss.CoverageRS = oMyStruct.getProp(saryPropNames(lPropNamesCount))
            ElseIf StrComp(saryPropNames(lPropNamesCount), "EndorsementRS", vbTextCompare) = 0 Then
                Set mudtXML01LossReport.XML01Loss.EndorsementRS = oMyStruct.getProp(saryPropNames(lPropNamesCount))
            ElseIf StrComp(saryPropNames(lPropNamesCount), "LossDetailRS", vbTextCompare) = 0 Then
                Set mudtXML01LossReport.XML01Loss.LossDetailRS = oMyStruct.getProp(saryPropNames(lPropNamesCount))
            ElseIf StrComp(saryPropNames(lPropNamesCount), "PaymentDetailRS", vbTextCompare) = 0 Then
                Set mudtXML01LossReport.XML01Loss.PaymentDetailRS = oMyStruct.getProp(saryPropNames(lPropNamesCount))
            ElseIf StrComp(saryPropNames(lPropNamesCount), "PriorLossDetailRS", vbTextCompare) = 0 Then
                Set mudtXML01LossReport.XML01Loss.PriorLossDetailRS = oMyStruct.getProp(saryPropNames(lPropNamesCount))
            ElseIf StrComp(saryPropNames(lPropNamesCount), "PolicyDetailRS", vbTextCompare) = 0 Then
                Set mudtXML01LossReport.XML01Loss.PolicyDetailRS = oMyStruct.getProp(saryPropNames(lPropNamesCount))
            ElseIf StrComp(saryPropNames(lPropNamesCount), "AddressRS", vbTextCompare) = 0 Then
                Set mudtXML01LossReport.XML01Loss.AddressRS = oMyStruct.getProp(saryPropNames(lPropNamesCount))
            ElseIf StrComp(saryPropNames(lPropNamesCount), "VehicleDetailRS", vbTextCompare) = 0 Then
                Set mudtXML01LossReport.XML01Loss.VehicleDetailRS = oMyStruct.getProp(saryPropNames(lPropNamesCount))
            End If
        Next

    End If
    
    If Not mudtXML01LossReport.XML01Loss.AssignmentDetailRS Is Nothing Then
        On Error Resume Next
        msInsuredName = mudtXML01LossReport.XML01Loss.AssignmentDetailRS.getField(1, "FirstName") & " " & mudtXML01LossReport.XML01Loss.AssignmentDetailRS.getField(1, "LastName")
        msDateAssign = mudtXML01LossReport.XML01Loss.AssignmentDetailRS.getField(1, "DateAssigned")
        msACID = mudtXML01LossReport.XML01Loss.AssignmentDetailRS.getField(1, "AssignedTo")
        msCLIENTNUM = mudtXML01LossReport.XML01Loss.AssignmentDetailRS.getField(1, "UnitNumber")
        msCLIENTNUM = Replace(msCLIENTNUM, "-", vbNullString, , , vbBinaryCompare)
                
        'Now update the actual Wddx RS ClaimNumber
       mudtXML01LossReport.XML01Loss.AssignmentDetailRS.setField 1, "ClaimNumber", msCLIENTNUM
        If Err.Number <> 0 Then
            Err.Clear
        End If
        On Error GoTo EH
    End If
    If Not mudtXML01LossReport.XML01Loss.ContactDetailRS Is Nothing Then
        On Error Resume Next
        msWorkPhone = GetContactItem(mudtXML01LossReport.XML01Loss.ContactDetailRS.getField(1, "ContactRowID"), "WorkPhoneNumber")
        msHomePhone = GetContactItem(mudtXML01LossReport.XML01Loss.ContactDetailRS.getField(1, "ContactRowID"), "HomePhoneNumber")
        If Err.Number <> 0 Then
            Err.Clear
        End If
        On Error GoTo EH
    End If
    
    msStatus = "PENDING"
    msTypeOfACID = "FACT"
    msAdjuster = "*ADJUSTERUSERNAME*"
    
    'If this is A file path leav it alone... For Export
    If InStr(1, mudtXML01LossReport.PrnKey, ":\") = 0 Then
        mudtXML01LossReport.PrnKey = "*ADJUSTERUSERNAME*" & "_" & msACID & "_" & "*IBNUMBER*" & "_" & msCLIENTNUM
    End If
    
CLEAN_UP:

    Set oDeser = Nothing
    Set oMyStruct = Nothing
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLR_ClassName & vbCrLf & "Private Sub PopulateMe"
End Sub

Public Function GetCoverageItem(psCoverageRowID As String, psCoverageItem As String, Optional plEditRowPos As Long) As String
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim lCount As Long
    
    With mudtXML01LossReport.XML01Loss
        For lCount = 1 To .CoverageRS.getRowCount
            If StrComp(.CoverageRS.getField(lCount, "CoverageRowID"), psCoverageRowID, vbTextCompare) = 0 Then
                plEditRowPos = lCount
                GetCoverageItem = IIf(IsNull(.CoverageRS.getField(lCount, psCoverageItem)), vbNullString, .CoverageRS.getField(lCount, psCoverageItem))
                Exit Function
            End If
        Next
    End With
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLR_ClassName & vbCrLf & "Public Function GetCoverageItem"
End Function

Public Function GetCoverageItemByCoverage(psCoverage As String, psCoverageItem As String) As String
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim lCount As Long
    
    With mudtXML01LossReport.XML01Loss
        For lCount = 1 To .CoverageRS.getRowCount
            If StrComp(.CoverageRS.getField(lCount, "Coverage"), psCoverage, vbTextCompare) = 0 Then
                GetCoverageItemByCoverage = IIf(IsNull(.CoverageRS.getField(lCount, psCoverageItem)), vbNullString, .CoverageRS.getField(lCount, psCoverageItem))
                Exit Function
            End If
        Next
    End With
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLR_ClassName & vbCrLf & "Public Function GetCoverageItemByCoverage"
End Function

Public Function GetEndorsementItem(psEndorsementsRowID As String, psEndorsementItem As String, Optional plEditRowPos As Long) As String
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim lCount As Long
    
    With mudtXML01LossReport.XML01Loss
        For lCount = 1 To .EndorsementRS.getRowCount
            If StrComp(.EndorsementRS.getField(lCount, "EndorsementsRowID"), psEndorsementsRowID, vbTextCompare) = 0 Then
                plEditRowPos = lCount
                GetEndorsementItem = IIf(IsNull(.EndorsementRS.getField(lCount, psEndorsementItem)), vbNullString, .EndorsementRS.getField(lCount, psEndorsementItem))
                Exit Function
            End If
        Next
    End With
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLR_ClassName & vbCrLf & "Public Function GetEndorsementItem"
End Function

Public Function GetPaymentDetailItem(psPaymentDetailRowID As String, psPaymentDetailItem As String, Optional plEditRowPos As Long) As String
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim lCount As Long
    
    With mudtXML01LossReport.XML01Loss
        For lCount = 1 To .PaymentDetailRS.getRowCount
            If StrComp(.PaymentDetailRS.getField(lCount, "PaymentDetailRowID"), psPaymentDetailRowID, vbTextCompare) = 0 Then
                plEditRowPos = lCount
                GetPaymentDetailItem = IIf(IsNull(.PaymentDetailRS.getField(lCount, psPaymentDetailItem)), vbNullString, .PaymentDetailRS.getField(lCount, psPaymentDetailItem))
                Exit Function
            End If
        Next
    End With
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLR_ClassName & vbCrLf & "Public Function GetPaymentDetailItem"
End Function

Public Function GetPriorLossDetailItem(psPriorLossRowID As String, psPriorLossDetailItem As String, Optional plEditRowPos As Long) As String
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim lCount As Long
    
    With mudtXML01LossReport.XML01Loss
        For lCount = 1 To .PriorLossDetailRS.getRowCount
            If StrComp(.PriorLossDetailRS.getField(lCount, "PriorLossRowID"), psPriorLossRowID, vbTextCompare) = 0 Then
                plEditRowPos = lCount
                GetPriorLossDetailItem = IIf(IsNull(.PriorLossDetailRS.getField(lCount, psPriorLossDetailItem)), vbNullString, .PriorLossDetailRS.getField(lCount, psPriorLossDetailItem))
                Exit Function
            End If
        Next
    End With
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLR_ClassName & vbCrLf & "Public Function GetPriorLossDetailItem"
End Function

Public Function GetActivitiesItem(psActivityRowID As String, psActivitiesItem As String, Optional plEditRowPos As Long) As String
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim lCount As Long
    
    With mudtXML01LossReport.XML01Loss
        For lCount = 1 To .ActivitiesRS.getRowCount
            If StrComp(.ActivitiesRS.getField(lCount, "ActivityRowID"), psActivityRowID, vbTextCompare) = 0 Then
                plEditRowPos = lCount
                GetActivitiesItem = IIf(IsNull(.ActivitiesRS.getField(lCount, psActivitiesItem)), vbNullString, .ActivitiesRS.getField(lCount, psActivitiesItem))
                Exit Function
            End If
        Next
    End With
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLR_ClassName & vbCrLf & "Public Function GetActivitiesItem"
End Function


Public Function GetContactItem(psContactRowID As String, psContactItemName As String, Optional plEditRowPos As Long) As String
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim lCount As Long
'    Dim ary() As Variant
'
'    ary = mudtXML01LossReport.XML01Loss.ContactsRS.getColumnNames
    With mudtXML01LossReport.XML01Loss
        For lCount = 1 To .ContactsRS.getRowCount
            If StrComp(.ContactsRS.getField(lCount, "ContactRowID"), psContactRowID, vbTextCompare) = 0 Then
                plEditRowPos = lCount
                GetContactItem = IIf(IsNull(.ContactsRS.getField(lCount, psContactItemName)), vbNullString, .ContactsRS.getField(lCount, psContactItemName))
                Exit Function
            End If
        Next
    End With
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLR_ClassName & vbCrLf & "Public Function GetContactItem"
End Function

Public Function GetAddressItem(psAddressID As String, psAddressItemName As String, Optional plEditRowPos As Long) As String
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim lCount As Long
    Dim myvar As Variant
    Dim sTemp As String

    With mudtXML01LossReport.XML01Loss
        For lCount = 1 To .AddressRS.getRowCount
            If StrComp(.AddressRS.getField(lCount, "AddressID"), psAddressID, vbTextCompare) = 0 Then
                plEditRowPos = lCount
                GetAddressItem = IIf(IsNull(.AddressRS.getField(lCount, psAddressItemName)), vbNullString, .AddressRS.getField(lCount, psAddressItemName))
                Exit Function
            End If
        Next
    End With
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLR_ClassName & vbCrLf & "Public Function GetAddressItem"
End Function

Public Function GetAssignmentDetailItem(psUniqueID As String, psAssignmentDetailItemName As String, Optional plEditRowPos As Long) As String
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim lCount As Long
    Dim myvar As Variant
    Dim sTemp As String

    With mudtXML01LossReport.XML01Loss
        For lCount = 1 To .AssignmentDetailRS.getRowCount
            If StrComp(.AssignmentDetailRS.getField(lCount, "UniqueID"), psUniqueID, vbTextCompare) = 0 Then
                plEditRowPos = lCount
                GetAssignmentDetailItem = IIf(IsNull(.AssignmentDetailRS.getField(lCount, psAssignmentDetailItemName)), vbNullString, .AssignmentDetailRS.getField(lCount, psAssignmentDetailItemName))
                Exit Function
            End If
        Next
    End With
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLR_ClassName & vbCrLf & "Public Function GetAssignmentDetailItem"
End Function

Public Function GetLossDetailItem(psUniqueID As String, psLossDetailItemName As String, Optional plEditRowPos As Long) As String
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim lCount As Long
    Dim myvar As Variant
    Dim sTemp As String

    With mudtXML01LossReport.XML01Loss
        For lCount = 1 To .LossDetailRS.getRowCount
            If StrComp(.LossDetailRS.getField(lCount, "UniqueID"), psUniqueID, vbTextCompare) = 0 Then
                plEditRowPos = lCount
                GetLossDetailItem = IIf(IsNull(.LossDetailRS.getField(lCount, psLossDetailItemName)), vbNullString, .LossDetailRS.getField(lCount, psLossDetailItemName))
                Exit Function
            End If
        Next
    End With
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLR_ClassName & vbCrLf & "Public Function GetLossDetailItem"
End Function

Public Function GetVehicleDetailItem(psUniqueID As String, psVehicleDetailName As String, Optional plEditRowPos As Long) As String
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim lCount As Long
    Dim myvar As Variant
    Dim sTemp As String

    With mudtXML01LossReport.XML01Loss
        For lCount = 1 To .VehicleDetailRS.getRowCount
            If StrComp(.VehicleDetailRS.getField(lCount, "UniqueID"), psUniqueID, vbTextCompare) = 0 Then
                plEditRowPos = lCount
                GetVehicleDetailItem = IIf(IsNull(.VehicleDetailRS.getField(lCount, psVehicleDetailName)), vbNullString, .VehicleDetailRS.getField(lCount, psVehicleDetailName))
                Exit Function
            End If
        Next
    End With
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLR_ClassName & vbCrLf & "Public Function GetVehicleDetailItem"
End Function

Public Function GetContactDetailsItem(psUniqueID As String, psContactDetailsItemName As String, Optional plEditRowPos As Long) As String
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim lCount As Long
    Dim myvar As Variant
    Dim sTemp As String

    With mudtXML01LossReport.XML01Loss
        For lCount = 1 To .ContactDetailRS.getRowCount
            If StrComp(.ContactDetailRS.getField(lCount, "UniqueID"), psUniqueID, vbTextCompare) = 0 Then
                plEditRowPos = lCount
                GetContactDetailsItem = IIf(IsNull(.ContactDetailRS.getField(lCount, psContactDetailsItemName)), vbNullString, .ContactDetailRS.getField(lCount, psContactDetailsItemName))
                Exit Function
            End If
        Next
    End With
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLR_ClassName & vbCrLf & "Public Function GetContactDetailsItem"
End Function

Public Function GetPolicyDetailItem(psPolicyDetailItemName As String) As String
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim lCount As Long
    Dim myvar As Variant
    Dim sTemp As String

    With mudtXML01LossReport.XML01Loss
        lCount = 1
        GetPolicyDetailItem = IIf(IsNull(.PolicyDetailRS.getField(lCount, psPolicyDetailItemName)), vbNullString, .PolicyDetailRS.getField(lCount, psPolicyDetailItemName))
    End With
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLR_ClassName & vbCrLf & "Public Function GetPolicyDetailItem"
End Function


Private Function PrintXML01(psPrnDeviceName As String, pbPreview As Boolean, _
                           pbPDF As Boolean, plHwnd As Long, plChainType As Long) As Boolean
    On Error GoTo EH
    Dim MyXML01 As Object
    Dim sPDFPath As String
    Dim sCaption As String
    Dim dDays As Double
    Dim lErrNum As Long
    Dim sErrDesc As String

    'Build Caption
    sCaption = "Loss Report " & Chr(160) & " (" & mudtXML01LossReport.PrnKey & ")"

    'Set the Report Here
    If clsCarLR_LossType = XML01Pro Then
        Set MyXML01 = New arLossXML01
    ElseIf clsCarLR_LossType = XML01Apd Then
        Set MyXML01 = New arLossXML01Apd
    End If

    With MyXML01
        .Printer.Orientation = ddOPortrait
        On Error Resume Next
        .Printer.DeviceName = psPrnDeviceName
        If Err.Number > 0 Then
            Err.Clear
            On Error GoTo EH
        End If
        On Error GoTo EH
        .LossReport = mudtXML01LossReport
        Set .LossXML01 = Me
    End With

    'Check to see if we are bulding a Chain
    If plChainType = ChainType.FirstInChain Or plChainType = ChainType.NextLink Then
        moLRs.BuildChain MyXML01, plChainType, sCaption, clsCarLR_Adjuster
        GoTo CLEANUP
    End If
    'If we are previewing then we need to use ARV object
    If pbPreview Then
        If goUtil.gARV Is Nothing Then
            Set goUtil.gARV = New V2ARViewer.clsARViewer
        End If

        With goUtil.gARV
            .SetUtilObject goUtil
            MyXML01.Run 'Run it before it gets to ARV object
            .objARvReport = MyXML01
            .sRptTitle = sCaption
            .ShowReport vbModeless
        End With
        Screen.MousePointer = vbDefault
    Else 'If we not preview then check to see if this is going to PDF
        If pbPDF Then
            'Set PDF File
            sPDFPath = Trim(moLRs.LossReportsForm.txtPDFPath.Text)
            sPDFPath = sPDFPath & sCaption
            'Debuging
            sPDFPath = left(sPDFPath, InStrRev(sPDFPath, "\", , vbBinaryCompare))
            If Me.clsCarLR_LossType = TypeXML01.XML01Apd Then
                sPDFPath = sPDFPath & mudtXML01LossReport.XML01Loss.AssignmentDetailRS.getField(1, "ClaimNumber") & "_APD" & ".pdf"
            Else
                sPDFPath = sPDFPath & mudtXML01LossReport.XML01Loss.AssignmentDetailRS.getField(1, "ClaimNumber") & "_PRO" & ".pdf"
            End If
            
            'debugging
            moLRs.ExportFile MyXML01, sPDFPath, ExportType.ARPdf 'Export to PDF
        Else
            MyXML01.PrintReport False 'Print without showing print dialog
        End If
    End If

CLEANUP:
    PrintXML01 = True

    'Clean up
    Set MyXML01 = Nothing
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLR_ClassName & vbCrLf & "Private Function PrintXML01"
End Function

Public Function clsCarLR_ExportME(poLRs As V2ECKeyBoard.clsLossReports, pXportType As ExportType) As Boolean
    On Error GoTo EH
    Dim MyXML01 As Object
    Dim sXportPath As String
    Dim lErrNum As Long
    Dim sErrDesc As String

    'Set the Report Here
    If clsCarLR_LossType = XML01Pro Then
        Set MyXML01 = New arLossXML01
    ElseIf clsCarLR_LossType = XML01Apd Then
        Set MyXML01 = New arLossXML01Apd
    End If

    With MyXML01
        
        .Printer.Orientation = ddOPortrait
        .LossReport = mudtXML01LossReport
        Set .LossXML01 = Me
    End With

    sXportPath = clsCarLR_PrnKey
    If poLRs.ExportFile(MyXML01, sXportPath, pXportType) Then
        clsCarLR_ExportME = True
    End If

    'Clean up
    Set MyXML01 = Nothing

    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLR_ClassName & vbCrLf & "Public Function ExportME"
End Function

Private Function SQL_WEBCONTROL_20000(Optional psSQLError As String) As String
    On Error GoTo EH
    Dim sSQL As String
    Dim sUserName As String
    Dim sTemp As String
    Dim sTemp2 As String
    Dim sCovType As String
    Dim sUnitClass As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim bAssignByZip As Boolean
    Dim ContactDetail As udtXML01ContactDetail
    Dim Contact As udtXML01Contacts
    Dim oLRXML01 As clsLossXML01
    Dim oLRXML01Edit As clsLossXML01
        
    moLRs.FireGetAssignByZip bAssignByZip
       
    sUserName = goUtil.utGetECSCryptSetting("V2WebControl", "DBConn", "USERID")
        
    With mudtXML01LossReport.XML01Loss
        If clsCarLR_LossType = XML01Pro Then
        
            'Check the psSQLError for Already Exisitng Claim Key
            'This Can happen for Property since Claims are being
            'split up into multiple units.
            If psSQLError = vbNullString Then
                sSQL = "z_spuInsertECSFarmersXML01Property "
                sSQL = sSQL & S_z & sUserName & z_S                 'USERNAME (DB Update Priv.)
                sSQL = sSQL & S_z & clsCarLR_ClassName & z_S        'LRFormat Loss Report Format
                sSQL = sSQL & S_z & clsCarLR_PrnData & z_S          'Loss Report printer Data (ntext)
                sSQL = sSQL & "default" & ", "                      'Company Code (Use Default Value)
                sSQL = sSQL & "default" & ", "                      'Client Company Code (Use Default Value)
                sSQL = sSQL & S_z & clsCarLR_TypeOfACID & z_S       'Type Of ACID (Adjuster CLient ID)
                sSQL = sSQL & "default" & ", "                      'Default Fee Schedule name (Use Default Value)
                sSQL = sSQL & "default" & ", "                      'Assignment Type (Use Default Value)
                                                                    'bAssignByZip
                If bAssignByZip Then
                     sSQL = sSQL & "1" & ", "
                Else
                     sSQL = sSQL & "0" & ", "
                End If
                sSQL = sSQL & S_z & IIf(IsNull(.PolicyDetailRS.getField(1, "PolicyNumber")), vbNullString, .PolicyDetailRS.getField(1, "PolicyNumber")) & z_S
                sSQL = sSQL & S_z & IIf(IsNull(.PolicyDetailRS.getField(1, "Status")), vbNullString, .PolicyDetailRS.getField(1, "Status")) & z_S
                sSQL = sSQL & S_z & IIf(IsNull(.ContactDetailRS.getField(1, "AgentPrimaryPhone")), vbNullString, .ContactDetailRS.getField(1, "AgentPrimaryPhone")) & z_S
                ContactDetail.ContactRowID = IIf(IsNull(.ContactDetailRS.getField(1, "ContactRowID")), vbNullString, .ContactDetailRS.getField(1, "ContactRowID"))
                sSQL = sSQL & S_z & GetContactItem(ContactDetail.ContactRowID, "HomePhoneNumber") & z_S
                sSQL = sSQL & S_z & GetContactItem(ContactDetail.ContactRowID, "WorkPhoneNumber") & z_S
                sSQL = sSQL & S_z & IIf(IsNull(.PolicyDetailRS.getField(1, "MortgageeName")), vbNullString, .PolicyDetailRS.getField(1, "MortgageeName")) & z_S
                sSQL = sSQL & S_z & IIf(IsNull(.PolicyDetailRS.getField(1, "CompanyCode")), vbNullString, .PolicyDetailRS.getField(1, "CompanyCode")) & z_S
                sTemp = IIf(IsNull(.PolicyDetailRS.getField(1, "PolicyDescription")), vbNullString, .PolicyDetailRS.getField(1, "PolicyDescription"))
                sTemp = sTemp & " " & IIf(IsNull(.PolicyDetailRS.getField(1, "PolicyEdition")), vbNullString, .PolicyDetailRS.getField(1, "PolicyEdition"))
                sSQL = sSQL & S_z & sTemp & z_S '@FRPolicyDescription
                sSQL = sSQL & S_z & IIf(IsNull(.PolicyDetailRS.getField(1, "NewBusinessDate")), vbNullString, Format(.PolicyDetailRS.getField(1, "NewBusinessDate"), "MM/DD/YYYY")) & z_S
                sSQL = sSQL & S_z & IIf(IsNull(.PolicyDetailRS.getField(1, "RenewalDate")), vbNullString, Format(.PolicyDetailRS.getField(1, "RenewalDate"), "MM/DD/YYYY")) & z_S
                sSQL = sSQL & S_z & IIf(IsNull(.PolicyDetailRS.getField(1, "CancellationDate")), vbNullString, Format(.PolicyDetailRS.getField(1, "CancellationDate"), "MM/DD/YYYY")) & z_S
                sTemp = IIf(IsNull(.AssignmentDetailRS.getField(1, "FirstName")), vbNullString, .AssignmentDetailRS.getField(1, "FirstName")) & " "
                sTemp = sTemp & IIf(IsNull(.AssignmentDetailRS.getField(1, "LastName")), vbNullString, .AssignmentDetailRS.getField(1, "LastName"))
                sSQL = sSQL & S_z & sTemp & z_S ' @NamedInsured
                Contact.PrimaryAddressID = GetContactItem(ContactDetail.ContactRowID, "PrimaryAddressID")
                sTemp = GetAddressItem(Contact.PrimaryAddressID, "StreetAddress")
                sTemp = sTemp & " " & GetAddressItem(Contact.PrimaryAddressID, "StreetAddress2")
                sSQL = sSQL & S_z & sTemp & z_S '@MailAddress1
                sTemp = GetAddressItem(Contact.PrimaryAddressID, "City") & ", "
                sTemp = sTemp & GetAddressItem(Contact.PrimaryAddressID, "State") & " "
                sTemp = sTemp & GetAddressItem(Contact.PrimaryAddressID, "PostalCode")
                sSQL = sSQL & S_z & sTemp & z_S '@MailAddress2
                sTemp = IIf(IsNull(.LossDetailRS.getField(1, "LossLocationAddress")), vbNullString, .LossDetailRS.getField(1, "LossLocationAddress"))
                sTemp2 = IIf(IsNull(.LossDetailRS.getField(1, "LossLocationAddress2")), vbNullString, .LossDetailRS.getField(1, "LossLocationAddress2"))
                If Trim(sTemp2) <> vbNullString Then
                    sTemp = sTemp & " " & sTemp2
                End If
                sSQL = sSQL & S_z & sTemp & z_S '@PropAddress1
                sTemp = IIf(IsNull(.LossDetailRS.getField(1, "LossLocationCity")), vbNullString, .LossDetailRS.getField(1, "LossLocationCity")) & ", "
                sTemp = sTemp & IIf(IsNull(.LossDetailRS.getField(1, "LossLocationState")), vbNullString, .LossDetailRS.getField(1, "LossLocationState")) & " "
                sTemp = sTemp & IIf(IsNull(.LossDetailRS.getField(1, "LossLocationZip")), vbNullString, .LossDetailRS.getField(1, "LossLocationZip"))
                sSQL = sSQL & S_z & sTemp & z_S '@PropAddress2
                'Always Use Building for Default Deductible
                sCovType = "Building"
                sTemp = LTrim(Right(GetCoverageItemByCoverage(sCovType, "Deductible1"), 10))
                sSQL = sSQL & S_z & Val(goUtil.utCleanValString(sTemp)) & z_S '@Building_Deductible1
                sSQL = sSQL & S_z & IIf(IsNull(.AssignmentDetailRS.getField(1, "CatastropheCode")), vbNullString, .AssignmentDetailRS.getField(1, "CatastropheCode")) & z_S
                sSQL = sSQL & S_z & IIf(IsNull(.AssignmentDetailRS.getField(1, "LossDate")), vbNullString, Format(.AssignmentDetailRS.getField(1, "LossDate"), "MM/DD/YYYY")) & z_S
                sSQL = sSQL & S_z & IIf(IsNull(.AssignmentDetailRS.getField(1, "AssignedTo")), vbNullString, .AssignmentDetailRS.getField(1, "AssignedTo")) & z_S
        '        sSQL = sSQL & S_z & IIf(IsNull(.AssignmentDetailRS.getField(1, "DateAssigned")), vbNullString, .AssignmentDetailRS.getField(1, "DateAssigned")) & z_S
                sSQL = sSQL & S_z & Format(Now(), "MM/DD/YYYY") & z_S '@DateAssigned
                sSQL = sSQL & S_z & vbNullString & z_S '@DateClsd
                sSQL = sSQL & S_z & IIf(IsNull(.AssignmentDetailRS.getField(1, "ClaimNumber")), vbNullString, .AssignmentDetailRS.getField(1, "ClaimNumber")) & z_S
                sSQL = sSQL & S_z & msUniqueID & z_S ' @UnitUniqueID
                sSQL = sSQL & S_z & msClaimId & z_S ' @ClaimId
                
                sCovType = IIf(IsNull(.AssignmentDetailRS.getField(1, "Type")), vbNullString, .AssignmentDetailRS.getField(1, "Type"))
                sTemp = GetCoverageItemByCoverage(sCovType, "Limits")
                sTemp = Val(goUtil.utCleanValString(sTemp))
                sSQL = sSQL & S_z & sTemp & z_S '@UnitLimit
                If sTemp = "0" Then
                    sTemp = GetCoverageItemByCoverage(sCovType, "Limits")
                Else
                    sTemp = vbNullString
                End If
                If sTemp = vbNullString Then
                    sTemp = " "
                End If
                sSQL = sSQL & S_z & sTemp & z_S '@UnitLimitOther
                If StrComp(sCovType, "Building", vbTextCompare) = 0 Then
                    sUnitClass = "A"
                ElseIf StrComp(sCovType, "Contents", vbTextCompare) = 0 Then
                    sUnitClass = "C"
                ElseIf StrComp(sCovType, "ALE", vbTextCompare) = 0 Then
                    sUnitClass = "D"
                Else
                    sUnitClass = "OTHER"
                End If
                sSQL = sSQL & S_z & sUnitClass & z_S
                sSQL = sSQL & S_z & .AssignmentDetailRS.getField(1, "CauseOfLoss") & z_S '@UnitCauseOfLoss
                '**********************************************************
                '*     BEGIN   FarmersXML01_LocationUpdateQue For Auto Only
                '*     5.10.2005 BGS Need to Insert Update Record for Location updates.
                '*     This Record will exist FarmersXML01_LocationUpdateQue table
                '*     Triggers on Assignment Updates will set SendToCRN bit column
                '*     Every hour WebControl will query anything needing to be sent
                '*     to CRN.
                '****************************************************************
                'Property Types are not updating these items at this time
                'Just sending place holders just incase later we need them...
'                sSQL = sSQL & S_z & .LossDetailRS.getField(1, "ContactRowId") & z_S          '@ContactRowID            varchar(20)=Null,
'                sSQL = sSQL & S_z & .LossDetailRS.getField(1, "UniqueID") & z_S              '@PropertyRowId           varchar(20)=Null, **** USE UniqueID****
'                sSQL = sSQL & S_z & .LossDetailRS.getField(1, "PropertyDescription") & z_S   '@PropertyDescription     varchar(50)='',
'                sSQL = sSQL & S_z & .LossDetailRS.getField(1, "VehicleMake") & z_S           '@VehicleMake             varchar(30)='',
'                sSQL = sSQL & S_z & .LossDetailRS.getField(1, "VehicleModel") & z_S          '@VehicleModel            varchar(30)='',
'                sSQL = sSQL & S_z & .LossDetailRS.getField(1, "VehicleYear") & z_S           '@VehicleYear             varchar(4)='',
'                sSQL = sSQL & S_z & .LossDetailRS.getField(1, "InspectedVIN") & z_S          '@InspectedVIN            varchar(30)=''
'                sSQL = sSQL & S_z & .LossDetailRS.getField(1, "LocationType") & z_S          '@LocationType            varchar(50)='',
'                sSQL = sSQL & S_z & .LossDetailRS.getField(1, "LocationName") & z_S          '@LocationName            varchar(50)='',
'                sSQL = sSQL & S_z & .LossDetailRS.getField(1, "LocationAddress") & z_S       '@LocationAddress         varchar(200)='',
'                sSQL = sSQL & S_z & .LossDetailRS.getField(1, "LocationCity") & z_S          '@LocationCity            varchar(50)='',
'                sSQL = sSQL & S_z & .LossDetailRS.getField(1, "LocationCountry") & z_S       '@LocationCountry         varchar(50)='',
'                sSQL = sSQL & S_z & .LossDetailRS.getField(1, "LocationState") & z_S         '@LocationState           varchar(50)='',
'                sSQL = sSQL & S_z & .LossDetailRS.getField(1, "LocationPostalCode") & z_S    '@LocationPostalCode      varchar(30)='',
'                sSQL = sSQL & S_z & .LossDetailRS.getField(1, "LocationPhoneNumber") & z_S   '@LocationPhoneNumber     varchar(50)='',
'                sSQL = sSQL & S_z & .LossDetailRS.getField(1, "GMTUpdated") & S_z            '@GMTUpdated              datetime=null

                sSQL = sSQL & S_z & vbNullString & z_S  '@ContactRowID            varchar(20)=Null,
                sSQL = sSQL & S_z & vbNullString & z_S  '@PropertyRowId           varchar(20)=Null, **** USE UniqueID****
                sSQL = sSQL & S_z & vbNullString & z_S  '@PropertyDescription     varchar(50)='',
                sSQL = sSQL & S_z & vbNullString & z_S  '@VehicleMake             varchar(30)='',
                sSQL = sSQL & S_z & vbNullString & z_S  '@VehicleModel            varchar(30)='',
                sSQL = sSQL & S_z & vbNullString & z_S  '@VehicleYear             varchar(4)='',
                sSQL = sSQL & S_z & vbNullString & z_S  '@InspectedVIN            varchar(30)=''
                sSQL = sSQL & S_z & vbNullString & z_S  '@LocationType            varchar(50)='',
                sSQL = sSQL & S_z & vbNullString & z_S  '@LocationName            varchar(50)='',
                sSQL = sSQL & S_z & vbNullString & z_S  '@LocationAddress         varchar(200)='',
                sSQL = sSQL & S_z & vbNullString & z_S  '@LocationCity            varchar(50)='',
                sSQL = sSQL & S_z & vbNullString & z_S  '@LocationCountry         varchar(50)='',
                sSQL = sSQL & S_z & vbNullString & z_S  '@LocationState           varchar(50)='',
                sSQL = sSQL & S_z & vbNullString & z_S  '@LocationPostalCode      varchar(30)='',
                sSQL = sSQL & S_z & vbNullString & z_S  '@LocationPhoneNumber     varchar(50)='',
                sSQL = sSQL & S_z & vbNullString & S_z  '@GMTUpdated              datetime=null

                '****************************************************************
                '*     END     FarmersXML01_LocationUpdateQue
                '****************************************************************
            Else
                'Check for Error Message
                If InStr(1, psSQLError, "*ADJUSTERUSERNAME*", vbTextCompare) > 0 Then
'                    'Will Get here if Adding UNIT to Exisitng Loss from memory...
'                    'Means there are subsequent unit Assign transactions within
'                    'the same Transaction Process
                   UpdateUnit Me.clsCarLR_PrnData, True
                   msStatus = "AddToExistingLoss"
                ElseIf InStr(1, psSQLError, "V2ECCarFarmers.clsLossXML01|", vbTextCompare) > 0 Then
'                    'Will get here if adding UNIT to Exisiting Loss After previous UNIT(s)
'                    'were already added and inserted to the DB
                    UpdateUnit Me.clsCarLR_PrnData, True
                    msStatus = "AddToExistingLoss"
                End If
            End If
        ElseIf clsCarLR_LossType = XML01Apd Then
            If psSQLError = vbNullString Then
                sSQL = "z_spuInsertECSFarmersXML01Auto "
                sSQL = sSQL & S_z & sUserName & z_S             'USERNAME (DB Update Priv.)
                sSQL = sSQL & S_z & clsCarLR_ClassName & z_S    'LRFormat
                sSQL = sSQL & S_z & clsCarLR_PrnData & z_S      'Loss Report Data ntext
                sSQL = sSQL & "default" & ", "                  'Company Code (Use Default Value)
                sSQL = sSQL & "default" & ", "                  'Client Company Code (Use Default Value)
                sSQL = sSQL & S_z & clsCarLR_TypeOfACID & z_S   'Type Of ACID (Adjuster CLient ID)
                sSQL = sSQL & "default" & ", "                  'Default Fee Schedule name (Use Default Value)
                sSQL = sSQL & "default" & ", "                  'Assignment Type (Use Default Value)
                'bAssignByZip
                If bAssignByZip Then
                     sSQL = sSQL & "1" & ", "
                Else
                     sSQL = sSQL & "0" & ", "
                End If
                sSQL = sSQL & S_z & IIf(IsNull(.PolicyDetailRS.getField(1, "PolicyNumber")), vbNullString, .PolicyDetailRS.getField(1, "PolicyNumber")) & z_S
                sSQL = sSQL & S_z & IIf(IsNull(.PolicyDetailRS.getField(1, "Status")), vbNullString, .PolicyDetailRS.getField(1, "Status")) & z_S
                sSQL = sSQL & S_z & IIf(IsNull(.ContactDetailRS.getField(1, "AgentPrimaryPhone")), vbNullString, .ContactDetailRS.getField(1, "AgentPrimaryPhone")) & z_S
                ContactDetail.ContactRowID = IIf(IsNull(.ContactDetailRS.getField(1, "ContactRowID")), vbNullString, .ContactDetailRS.getField(1, "ContactRowID"))
                sSQL = sSQL & S_z & GetContactItem(ContactDetail.ContactRowID, "HomePhoneNumber") & z_S
                sSQL = sSQL & S_z & GetContactItem(ContactDetail.ContactRowID, "WorkPhoneNumber") & z_S
                sSQL = sSQL & S_z & IIf(IsNull(.PolicyDetailRS.getField(1, "LienHolderName")), vbNullString, .PolicyDetailRS.getField(1, "LienHolderName")) & z_S
                sSQL = sSQL & S_z & IIf(IsNull(.PolicyDetailRS.getField(1, "CompanyCode")), vbNullString, .PolicyDetailRS.getField(1, "CompanyCode")) & z_S
                sTemp = IIf(IsNull(.PolicyDetailRS.getField(1, "PolicyDescription")), vbNullString, .PolicyDetailRS.getField(1, "PolicyDescription"))
                sTemp = sTemp & " " & IIf(IsNull(.PolicyDetailRS.getField(1, "PolicyEdition")), vbNullString, .PolicyDetailRS.getField(1, "PolicyEdition"))
                sSQL = sSQL & S_z & sTemp & z_S '@FRPolicyDescription
                sSQL = sSQL & S_z & IIf(IsNull(.PolicyDetailRS.getField(1, "NewBusinessDate")), vbNullString, Format(.PolicyDetailRS.getField(1, "NewBusinessDate"), "MM/DD/YYYY")) & z_S
                sSQL = sSQL & S_z & IIf(IsNull(.PolicyDetailRS.getField(1, "RenewalDate")), vbNullString, Format(.PolicyDetailRS.getField(1, "RenewalDate"), "MM/DD/YYYY")) & z_S
                sSQL = sSQL & S_z & IIf(IsNull(.PolicyDetailRS.getField(1, "CancellationDate")), vbNullString, Format(.PolicyDetailRS.getField(1, "CancellationDate"), "MM/DD/YYYY")) & z_S
                sTemp = IIf(IsNull(.AssignmentDetailRS.getField(1, "FirstName")), vbNullString, .AssignmentDetailRS.getField(1, "FirstName")) & " "
                sTemp = sTemp & IIf(IsNull(.AssignmentDetailRS.getField(1, "LastName")), vbNullString, .AssignmentDetailRS.getField(1, "LastName"))
                sSQL = sSQL & S_z & sTemp & z_S ' @NamedInsured
                Contact.PrimaryAddressID = GetContactItem(ContactDetail.ContactRowID, "PrimaryAddressID")
                sTemp = GetAddressItem(Contact.PrimaryAddressID, "StreetAddress")
                sTemp = sTemp & " " & GetAddressItem(Contact.PrimaryAddressID, "StreetAddress2")
                sSQL = sSQL & S_z & sTemp & z_S '@MailAddress1
                sTemp = GetAddressItem(Contact.PrimaryAddressID, "City") & ", "
                sTemp = sTemp & GetAddressItem(Contact.PrimaryAddressID, "State") & " "
                sTemp = sTemp & GetAddressItem(Contact.PrimaryAddressID, "PostalCode")
                sSQL = sSQL & S_z & sTemp & z_S '@MailAddress2
                sTemp = IIf(IsNull(.VehicleDetailRS.getField(1, "LocationAddress")), vbNullString, .VehicleDetailRS.getField(1, "LocationAddress"))
                sSQL = sSQL & S_z & sTemp & z_S '@PropAddress1
                sTemp = IIf(IsNull(.VehicleDetailRS.getField(1, "LocationCity")), vbNullString, .VehicleDetailRS.getField(1, "LocationCity")) & ", "
                sTemp = sTemp & IIf(IsNull(.VehicleDetailRS.getField(1, "LocationState")), vbNullString, .VehicleDetailRS.getField(1, "LocationState")) & " "
                sTemp = sTemp & IIf(IsNull(.VehicleDetailRS.getField(1, "LocationPostalCode")), vbNullString, .VehicleDetailRS.getField(1, "LocationPostalCode"))
                sSQL = sSQL & S_z & sTemp & z_S '@PropAddress2
                sTemp = IIf(IsNull(.VehicleDetailRS.getField(1, "VehicleMake")), vbNullString, .VehicleDetailRS.getField(1, "VehicleMake")) & " "
                sTemp = sTemp & IIf(IsNull(.VehicleDetailRS.getField(1, "VehicleModel")), vbNullString, .VehicleDetailRS.getField(1, "VehicleModel")) & " "
                sTemp = sTemp & IIf(IsNull(.VehicleDetailRS.getField(1, "VehicleYear")), vbNullString, .VehicleDetailRS.getField(1, "VehicleYear"))
                sSQL = sSQL & S_z & sTemp & z_S '@VehicleDescription
                sCovType = IIf(IsNull(.AssignmentDetailRS.getField(1, "Type")), vbNullString, .AssignmentDetailRS.getField(1, "Type"))
                If StrComp(sCovType, "Comp", vbTextCompare) = 0 Then
                    sCovType = "Comprehensive"
                End If
                sTemp = LTrim(Right(GetCoverageItemByCoverage(sCovType, "Deductible1"), 10))
                sSQL = sSQL & S_z & Val(goUtil.utCleanValString(sTemp)) & z_S '@CompDed
                sTemp = IIf(IsNull(.VehicleDetailRS.getField(1, "LocationType")), vbNullString, .VehicleDetailRS.getField(1, "LocationType")) & " "
                sTemp = sTemp & IIf(IsNull(.VehicleDetailRS.getField(1, "LocationName")), vbNullString, .VehicleDetailRS.getField(1, "LocationName"))
                sSQL = sSQL & S_z & sTemp & z_S '@VehicleLocation
                
                sSQL = sSQL & S_z & IIf(IsNull(.AssignmentDetailRS.getField(1, "CatastropheCode")), vbNullString, .AssignmentDetailRS.getField(1, "CatastropheCode")) & z_S
                sSQL = sSQL & S_z & IIf(IsNull(.AssignmentDetailRS.getField(1, "LossDate")), vbNullString, Format(.AssignmentDetailRS.getField(1, "LossDate"), "MM/DD/YYYY")) & z_S
                sSQL = sSQL & S_z & IIf(IsNull(.AssignmentDetailRS.getField(1, "AssignedTo")), vbNullString, .AssignmentDetailRS.getField(1, "AssignedTo")) & z_S
        '        sSQL = sSQL & S_z & IIf(IsNull(.AssignmentDetailRS.getField(1, "DateAssigned")), vbNullString, .AssignmentDetailRS.getField(1, "DateAssigned")) & z_S
                sSQL = sSQL & S_z & Format(Now(), "MM/DD/YYYY") & z_S '@DateAssigned
                sSQL = sSQL & S_z & vbNullString & z_S '@DateClsd
                sSQL = sSQL & S_z & IIf(IsNull(.AssignmentDetailRS.getField(1, "ClaimNumber")), vbNullString, .AssignmentDetailRS.getField(1, "ClaimNumber")) & z_S
                sSQL = sSQL & S_z & msUniqueID & z_S ' @UnitUniqueID
                sSQL = sSQL & S_z & msClaimId & z_S ' @ClaimId
                
                sTemp = GetCoverageItemByCoverage(sCovType, "Limits")
                sTemp = Val(goUtil.utCleanValString(sTemp))
                sSQL = sSQL & S_z & sTemp & z_S '@UnitLimit
                If sTemp = "0" Then
                    sTemp = GetCoverageItemByCoverage(sCovType, "Limits")
                Else
                    sTemp = vbNullString
                End If
                If sTemp = vbNullString Then
                    sTemp = " "
                End If
                sSQL = sSQL & S_z & sTemp & z_S '@UnitLimitOther
                If StrComp(sCovType, "Comprehensive", vbTextCompare) = 0 Then
                    sUnitClass = "Z"
                Else
                    sUnitClass = "OTHER"
                End If
                sSQL = sSQL & S_z & sUnitClass & z_S
                sSQL = sSQL & S_z & .AssignmentDetailRS.getField(1, "CauseOfLoss") & z_S '@UnitCauseOfLoss
                '**********************************************************
                '*     BEGIN   FarmersXML01_LocationUpdateQue For Auto Only
                '*     5.10.2005 BGS Need to Insert Update Record for Location updates.
                '*     This Record will exist FarmersXML01_LocationUpdateQue table
                '*     Triggers on Assignment Updates will set SendToCRN bit column
                '*     Every hour WebControl will query anything needing to be sent
                '*     to CRN.
                '****************************************************************
                sSQL = sSQL & S_z & .VehicleDetailRS.getField(1, "ContactRowId") & z_S          '@ContactRowID            varchar(20)=Null,
                sSQL = sSQL & S_z & .VehicleDetailRS.getField(1, "UniqueID") & z_S              '@PropertyRowId           varchar(20)=Null, **** USE UniqueID****
                sSQL = sSQL & S_z & .VehicleDetailRS.getField(1, "PropertyDescription") & z_S   '@PropertyDescription     varchar(50)='',
                sSQL = sSQL & S_z & .VehicleDetailRS.getField(1, "VehicleMake") & z_S           '@VehicleMake             varchar(30)='',
                sSQL = sSQL & S_z & .VehicleDetailRS.getField(1, "VehicleModel") & z_S          '@VehicleModel            varchar(30)='',
                sSQL = sSQL & S_z & .VehicleDetailRS.getField(1, "VehicleYear") & z_S           '@VehicleYear             varchar(4)='',
                sSQL = sSQL & S_z & .VehicleDetailRS.getField(1, "InspectedVIN") & z_S          '@InspectedVIN            varchar(30)=''
                sSQL = sSQL & S_z & .VehicleDetailRS.getField(1, "LocationType") & z_S          '@LocationType            varchar(50)='',
                sSQL = sSQL & S_z & .VehicleDetailRS.getField(1, "LocationName") & z_S          '@LocationName            varchar(50)='',
                sSQL = sSQL & S_z & .VehicleDetailRS.getField(1, "LocationAddress") & z_S       '@LocationAddress         varchar(200)='',
                sSQL = sSQL & S_z & .VehicleDetailRS.getField(1, "LocationCity") & z_S          '@LocationCity            varchar(50)='',
                sSQL = sSQL & S_z & .VehicleDetailRS.getField(1, "LocationCountry") & z_S       '@LocationCountry         varchar(50)='',
                sSQL = sSQL & S_z & .VehicleDetailRS.getField(1, "LocationState") & z_S         '@LocationState           varchar(50)='',
                sSQL = sSQL & S_z & .VehicleDetailRS.getField(1, "LocationPostalCode") & z_S    '@LocationPostalCode      varchar(30)='',
                sSQL = sSQL & S_z & .VehicleDetailRS.getField(1, "LocationPhoneNumber") & z_S   '@LocationPhoneNumber     varchar(50)='',
                sSQL = sSQL & S_z & .VehicleDetailRS.getField(1, "GMTUpdated") & S_z           '@GMTUpdated         datetime=null
                '****************************************************************
                '*     END     FarmersXML01_LocationUpdateQue
                '****************************************************************
            Else
                'Check for Error Message
                If InStr(1, psSQLError, "*ADJUSTERUSERNAME*", vbTextCompare) > 0 Then
                        '******************AUTOMOBILE*************************
                        '               5/25/2005
                        '               SHOULD NEVER GET HERE
                        '               SINCE UNITNUMBER = CLIENTNUM
                        'Auto is being handled 1 to 1 concerning AssignmentsID
                        '*****************************************************
'                    'Will Get here if Adding UNIT to Exisitng Loss from memory...
'                    'Means there are subsequent unit Assign transactions within
'                    'the same Transaction Process
                    Set oLRXML01 = moLRs.RawLossReportsCol(psSQLError)
                    'Check first to see if this is a DUPLICATE Unique ID found in memory
                    'IN other words during the file processing... two Identical transactions...
                    'meaning one or more transactions contain the same UniqueID and the Same AdjusterID
                    If StrComp(oLRXML01.UniqueID, msUniqueID, vbTextCompare) = 0 And StrComp(oLRXML01.clsCarLR_ACID, msACID, vbTextCompare) = 0 Then
                        Set oLRXML01 = Nothing
                        lErrNum = -999
                        sErrDesc = "Duplicate Unit Assignment to the same adjuster not allowed! UniqueID = " & msUniqueID & " "
                        Err.Raise lErrNum, , sErrDesc
                    End If
                ElseIf InStr(1, psSQLError, "V2ECCarFarmers.clsLossXML01|", vbTextCompare) > 0 Then
                        '******************AUTOMOBILE*************************
                        '               5/25/2005
                        '               SHOULD NEVER GET HERE
                        '               SINCE UNITNUMBER = CLIENTNUM
                        'Auto is being handled 1 to 1 concerning AssignmentsID
                        '*****************************************************
'                    'Will get here if adding UNIT to Exisiting Loss After previous UNIT(s)
'                    'were already added and inserted to the DB
                End If
            End If
        End If
        
    End With
    
    sSQL = Replace(sSQL, Chr(39), Chr(39) & Chr(39))
    sSQL = Replace(sSQL, S_z, Chr(39))
    sSQL = Replace(sSQL, z_S, "', ")
    
    SQL_WEBCONTROL_20000 = sSQL
    
    'CleanUp
    Set oLRXML01 = Nothing
    Set oLRXML01Edit = Nothing
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLR_ClassName & vbCrLf & "Private Function SQL_WEBCONTROL_20000"
End Function

Private Function OpenConnection(poConn As ADODB.Connection, Optional plOptions As Long = -1) As Boolean
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim sUserID As String
    Dim sPassword As String
    Dim sMess As String
    Dim sConnString As String

    sConnString = GetSetting("V2WebControl", "DSN", "NAME", vbNullString)

    If Not poConn Is Nothing Then
        If InStr(1, poConn.ConnectionString, sConnString, vbTextCompare) > 0 Then
            If poConn.State = ADODB.adStateOpen Then
                Exit Function
            End If
        End If
        Set poConn = Nothing
    End If

    sUserID = GetECSCryptSetting("V2WebControl", "DBConn", "USERID")
    sPassword = GetECSCryptSetting("V2WebControl", "DBConn", "PASSWORD")

    Set poConn = New ADODB.Connection
    poConn.Open sConnString, sUserID, sPassword, plOptions

    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLR_ClassName & vbCrLf & "Private Function OpenConnection"
End Function

Private Function clsCarLR_SetUtilObject(oUtilObject As Object) As Boolean

    On Error Resume Next
    Set goUtil = oUtilObject
    If Err.Number = 0 Then
        clsCarLR_SetUtilObject = True
    Else
        clsCarLR_SetUtilObject = False
    End If

End Function

Private Property Let clsCarLR_WorkPhone(psWorkPhone As String)
    msWorkPhone = psWorkPhone
End Property

Private Property Get clsCarLR_WorkPhone() As String
    clsCarLR_WorkPhone = msWorkPhone
End Property

Private Function BOF_UTF_HEADER() As String
    BOF_UTF_HEADER = "﻿<?xml version=""1.0"" encoding=""utf-8""?>"
End Function
