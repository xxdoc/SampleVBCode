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

Public Type udtXML01AdminLossInfo
    ali0004_DateTimePrinted As String
    ali0005_PrintedBy As String
    ali0052_ReportedBy As String
    ali0053_RBPhone As String
    ali0054_PolicyNum As String
    ali0055_SC As String
    ali0056_AgentNum As String
    ali0057_HomePhone As String
    ali0058_BusPhone As String
    ali0059_NewBusDate As String
    ali0060_RenewalDate As String
    ali0061_LastCancDate As String
    ali0062_NamedInsured As String
    ali0063_MailAddress1 As String
    ali0064_MailAddress2 As String
    ali0065_MainFileInsuredName As String
    ali0066_MortgageHolder As String
    ali0067_2ndMort As String
    ali0068_CompCode As String
    ali0069_PolicyDescription As String
    ali0070_BldgLimit As String
    ali0071_ContLimit As String
    ali0072_Deductible1 As String
    ali0073_Deductible2 As String
    ali0074_Deductible3 As String
    ali0075_Deductible4 As String
    ali0076_AddlCoverage1 As String
    ali0077_AddlCoverage2 As String
    ali0078_AddlCoverage3 As String
    ali0079_AddlCoverage4 As String
    ali0080_LossLocAddress1 As String
    ali0081_LossLocAddress2 As String
End Type

Public Type udtPolicyUnits
    UnitType As String
    UnitNumber As String
    UnitLimit As Currency
    UnitComments As String
    IsReassigned As Boolean
    ReassignedACID As String
End Type

Public Type udtContact
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
    PrimaryAddressId As String
    CreatedUpdated As Date
End Type

Public Type udtXML01AdminLossInfoApd
    ali0004_DateTimePrinted As String
    ali0005_PrintedBy As String
    '6 - 48 Labels and header
    ali0049_ReportedBy As String
    ali0050_RBPhone As String
    ali0051_PolicyNum As String
    ali0052_SC As String
    ali0053_AgentNum As String
    ali0054_HomePhone As String
    ali0055_BusPhone As String
    '56 Spacer For where Mainfile insured name
    ali0057_MortgageHolder As String
    ali0058_CompCode As String
    ali0059_PolicyType As String
    ali0060_NewBusDate As String
    ali0061_RenewalDate As String
    ali0062_LastCancDate As String
    '63 This should be with CLI
    cli0063_VehicleLocation As String
    ali0064_NamedInsured As String
    ali0065_MailAddress1 As String
    ali0066_MailAddress2 As String
    ali0067_VehicleDescription As String
    ali0068_CompDed As String
    ali0069_VIN As String
    '
    '70 - 73 will usually be part of Endorsements
    'Must do a check for Pos of SALN to be sure.
End Type

'Deductibles
Public Type udtDeductibles
    DeductibleDescription As String
End Type

'Additional Coverages
Public Type udtAdditionalCoverages
    CoverageDescription As String
End Type

'Endorsements UDTs will be Stored in varyEndorsements
'since the max is 10 but there could be anywhere from
'4 to 10 EDs. which the XML01 Loss report does not print
'null strings for the remainaing 10 items.
Public Type udtXML01Endorsement
    EDCode As String
    EDDescription As String
End Type

Public Type udtXML01CurrentLossInfo
    cli01_CAT As String
    cli02_LossDate As String
    cli03_Adjuster As String
    cli04_DateAsgn As String
    cli05_DateClsd As String
    cli06_SALN As String
    cli07_AdjusterOrigInfo As String
    cli08_DateAsgnOrigInfo As String
    cli09_DateClsdOrigInfo As String
    'I have never seen cli10 populated ever.
    'Every time the cli information will be the 1st 9 items.
    'We will have to check the transistion from CLI to PLH and CAL,
    'To be sure we check for PaymentsMade this claim. when we are
    'scanning the original "laser jet 4" prn file.
    cli10_PaymentsMadeThisClaim As String
End Type

'PriorLossHist UDTs will be stored in colPLH
'Since there may be multiple PriorLoss entries.
'As well there can be continuation pages that have
'Prior Loss entries continued.
Public Type udtXML01PriorLossHist '---------->PLH
    plh01_SALN As String
    plh02_LossDate As String
    plh03_Adjuster As String
    plh04_DateAsgn As String
    plh05_DateClsd As String
    plh06_AmtPaid As String
End Type

'PriorLossHist UDTs will be stored in colPLHCont
'The Cont PLH is slight diff than PLH
'Since there may be multiple PriorLoss entries.
'As well there can be continuation pages that have
'Prior Loss entries continued.
Public Type udtXML01PriorLossHistCont '---------->PLHCont
    plh01_CAT As String
    plh02_LossDate As String
    plh03_Adjuster As String
    plh04_DateAsgn As String
    plh05_DateClsd As String
    plh06_AmtPaid As String
    plh07_SALN As String
End Type

'CommentsActLog UDTs will be stored in colCAL
'Since there may be multiple Comments Act Log entries.
'As well, there can be Continuation pages that have
'Comments Act log entries continued.
Public Type udtXML01CommentsActLog '---------->CAL
    cal01_CAT As String
    cal02_Date As String
    cal03_Time As String
    cal04_Action As String
    cal05_User As String
    cal06_Comments As String
End Type

'This is the entire Loss Report (Main page)
Public Type udtXML01Loss
    AdminLossInfo As udtXML01AdminLossInfo
    AdminLossInfoApd As udtXML01AdminLossInfoApd
    colEndorsements As Collection
    CurrentLossInfo As udtXML01CurrentLossInfo
    colPLH As Collection
    colCAL As Collection
End Type

'This is a Continued Loss report (Continued page)
Public Type udtXML01LossCont
    colPLHCont As Collection
    colCAL As Collection
End Type

Public Type XML01LossReport
    XML01Loss As udtXML01Loss
    colXML01LossCont As Collection 'Usually if there is a Continuation page theres only 1 page but you never know
    PrnData As String 'Contains Page1 and all continued page data
    PrnKey As String 'File path for this Loss Report
End Type

Public Enum TypeXML01
    XML01Apd = 0
    XML01Pro
End Enum


Private mudtXML01LossReport As XML01LossReport
Private moLRs As V2ECKeyBoard.clsLossReports
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

Private Sub Class_Initialize()
    Set goUtil = New V2ECKeyBoard.clsUtil
End Sub

Private Sub Class_Terminate()
    Set goUtil = Nothing
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
                    sSQL = SQL_WEBCONTROL_20000
                'iF Newer versions Require Changes to SQL ...
                'They will Have  go on top in Select Case
                Case Is >= 108
                    sSQL = SQL_WEBCONTROL_108
                    
            End Select
            
'<--------------------------------------EASY CLAIM-------------------------------------->
        Case "EBERLS", "EASYCLAIM"
            Select Case plVersion
                Case Is >= 200
                    sSQL = SQL_EASYCLAIM_20000
                'iF Newer versions Require Changes to SQL ...
                'They will Have  go on top in Select Case
                Case Is >= 10135
                    sSQL = SQL_EASYCLAIM_10135
                
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

Public Property Let clsCarLR_PrnKey(psKey As String)
    mudtXML01LossReport.PrnKey = psKey
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

Public Function clsCarLR_ProcessRawData(psDataPath As String, _
                               psOutPath As String, _
                               Optional poPB As Object, _
                               Optional potxtProgMess As Object) As Boolean
    'Process the single file we get from Farmers.  This file will
    'contain multiple HP4000 Laser print jobs.  We can intially Split these up into
    'an array using special zEOJ (End of Job) characters as the delimeter. We also must make the
    'code intelligent enough after making the inital Split to look for "Continued"
    'print jobs.  Any print jobs that are continued must be added to the previous
    'Loss Report print job.  Usually there should only be one continued if there is
    'going to be one at all, but it is possible we could run into multiple continuations
    'after the initial Loss Report print job.  In any event all continuations should
    'be stored in the udtXML01LossCont user defined type.  This udt will then be added to
    'the colXML01LossCont collection which is part of the main XML01LossReport user defined type.
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
    Dim sRawFileName As String
    Dim bPB As Boolean 'True if using it
    Dim MyAdjuster As udtAdjuster
    Dim bMissing As Boolean
    Dim bCorruptFile As Boolean 'true if the XML01 file contains corrupt data
    Dim bOtherXML01 As Boolean 'True if XML01 is other than Property IE APD (Auto)
    Dim sTemp As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    If Not goUtil.utFileExists(psDataPath, True) Then
        Exit Function 'BAIL
    Else
        'check for Backup folder
        If Not goUtil.utFileExists(psDataPath & "\BackUp", True) Then
            goUtil.utMakeDir psDataPath & "\BackUp"
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
    
    'Load XML01 format files
    sFile = Dir(psDataPath & "\" & "*.*")
    
    If sFile <> vbNullString Then
        Set colFiles = New Collection
        Do
            colFiles.Add psDataPath & "\" & sFile
            sFile = Dir
        Loop Until sFile = vbNullString
    Else
        Exit Function 'BAIL
    End If
    
    'Loop through the collection of Files
    'There should only be 1 XML01 file per day
    'because it consists of many print jobs.
    'But we possibly could be processing multiple days.
    For Each vFile In colFiles
    
        If goUtil.utFileExists(CStr(vFile)) Then
            If Not otxtProgMess Is Nothing Then
                sTemp = CStr(vFile)
                sTemp = Mid(sTemp, InStrRev(sTemp, "\") + 1)
                otxtProgMess.Text = "Loading: " & sTemp
                otxtProgMess.Refresh
            End If
            'Load XML01 Multi Print Job data
            sData = goUtil.utGetFileData(CStr(vFile))
            'If for some reason this XML01 file was labled
            'incorreclty then skip it
            'CLean out "
            If InStr(1, sData, zEOJ) = 0 Then
                GoTo NEXT_FILE
            End If
            
            'Split up each print job delim by zEOJ characters
            varyData = Split(sData, zEOJ)
            sData = vbNullString 'Clear it
            
            'We will put each Loss (with its continued pages)
            'into this collection, GetLossCol will also prefix ClassName to each Loss
            bCorruptFile = False
            Set colLoss = GetLossCol(varyData, bCorruptFile)
            
            'check to see if this job is corrupt or not
            If colLoss Is Nothing Then
                GoTo CORRUPT_FILE
            End If
            
            Erase varyData 'Clear this array
            
            'Progress bar
            If bPB Then
                poPB.Max = colLoss.Count
                poPB.Value = 0
                If Not otxtProgMess Is Nothing Then
                    otxtProgMess.Text = poPB.Value & " Of " & poPB.Max
                    otxtProgMess.Refresh
                End If
            End If
            
            For Each vLoss In colLoss
                'Need to get the PrnData, SALN, FACT, and Print Job continuations
                'out of the XML01 print job. The SALN and FACT will be located in the first element
                'of varyData when we split it. If there are any additional elements, they will be continuation jobs.
                sData = CStr(vLoss)
                
                'Check to see if this XML01 is other than Property
                If StrComp(Left(sData, 14), "clsLossUnknown", vbTextCompare) = 0 Then
                    bOtherXML01 = True
                Else
                    bOtherXML01 = False
                End If
                
                'PRN DATA
                sPrnData = sData
                
                varyData = Split(sData, zEOJ)
                sData = varyData(LBound(varyData, 1))
            
                varyLoss = Split(sData, ";LB") 'Use ';LB' as delimeter
                
                'FACT or FAX id will always be the UBOUND
                sData = varyLoss(UBound(varyLoss, 1))
                'Re format the data getting rid of Laser print job formating crudola
                sData = Left(sData, InStr(1, sData, Chr(180)) - 1)
                sFact = sData
                
                'SALN or Claim number will always be 4 less than UBOUND
                sData = varyLoss(UBound(varyLoss, 1) - 4)
                'Re format the data getting rid of Laser print job formating crudola
                sData = Left(sData, InStr(1, sData, Chr(180)) - 1)
                sSALN = sData
                sSALN = Replace(sSALN, "-", vbNullString, , , vbTextCompare)
                'Farmers uses FACT ID for XML01 Property Reports which is THE ACID (Adjuster Client ID)
                msACID = sFact
                msCLIENTNUM = sSALN
               
                'Now that we have everything we can add the loss object
                
                'Need to create Loss Report Object and add it to
                'The Raw Loss Report collection
                If bOtherXML01 Then
                    Set oLR = New clsLossUnknown
                Else
                    Set oLR = New clsLossXML01
                End If
                
                With oLR
                    Set .LRs = moLRs
                    'Remove the class marker frm the prn data
                    If bOtherXML01 Then
                        sPrnData = Replace(sPrnData, "clsLossUnknown" & vbCrLf, vbNullString, 1)
                    Else
                        sPrnData = Replace(sPrnData, clsCarLR_ClassName & vbCrLf, vbNullString, 1)
                    End If
                    .PrnData = sPrnData
                    .PopulateMe
                End With
                
                moLRs.AddRawLossReport oLR
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
                
            Next
CORRUPT_FILE:
            'Make a Backup copy of the Raw File
            'Check to see if the Raw File is already in the
            'BackUp folder first
            sFile = CStr(vFile)
            sFile = Mid(sFile, InStrRev(sFile, "\") + 1)
            If bCorruptFile Then
                sRawFileName = psDataPath & "\BackUp\" & "CorruptFile_" & sFile
            Else
                sRawFileName = psDataPath & "\BackUp\" & sFile
            End If
            
            If goUtil.utFileExists(sRawFileName) Then
                SetAttr CStr(sRawFileName), vbNormal
                Kill CStr(sRawFileName)
            End If
            FileCopy CStr(vFile), sRawFileName
            'Get rid of the original Raw File since it is copied into
            'the backup folder
            SetAttr CStr(vFile), vbNormal
            Kill CStr(vFile)
        End If
NEXT_FILE:
    Next
CLEAN_UP:
    'CleanUp
    Set otxtProgMess = Nothing
    Set colFiles = Nothing
    Set colLoss = Nothing
    Set oLR = Nothing
    clsCarLR_ProcessRawData = True
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Set colFiles = Nothing
    Set colLoss = Nothing
    Set oLR = Nothing
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLR_ClassName & vbCrLf & "Public Function ProcessRawData"
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
            moLRs.FirePrintAppDoc mudtXML01LossReport.XML01Loss.CurrentLossInfo.cli06_SALN
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
    Dim vLossP1 As Variant 'Contains the first page of Loss Report
    Dim colCont As Collection 'contains collection of Loss continuation pages
    Dim lPos As Long  'Used to flag what element Pos to beging looking in vLossP1
    Dim pbPLH As Boolean 'Flags that we have Prior Loss History data
    Dim pbCal As Boolean 'Flags that we Have Comments Activity Log Data
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    If Not GetLossData(vLossP1, colCont) Then
        GoTo CLEAN_UP
    End If
    
    'Get the SALN number so we can use it to Pos the end & start of endorsements, CLI, PLH, CAL.
    'SALN or Claim number will always be 4 less than UBOUND
    msCLIENTNUM = vLossP1(UBound(vLossP1, 1) - 4)
    
    '1. Populate the admin Loss information
    PopulateALI vLossP1, lPos
    
    '2. Populate the Endorsements collection
    PopulateEDCol vLossP1, lPos
    
    '3. Populate the Current Loss information
    PopulateCLI vLossP1, lPos, pbPLH, pbCal
    
    '4. Populate the Prior Loss History Collection
    If pbPLH Then
        PopulatePLHCol vLossP1, lPos, mudtXML01LossReport.XML01Loss.CurrentLossInfo.cli01_CAT, pbCal
    End If
    
    '5. Populate the Comments Activity Log Collection
    If pbCal Then
        PopulateCALCol vLossP1, lPos, mudtXML01LossReport.XML01Loss.CurrentLossInfo.cli01_CAT
    End If
    
    '6. Populate the Loss continuation collection with Loss continuations udts
    If Not colCont Is Nothing Then
        PopulateLossContCol colCont
    End If
    If mLossType = XML01Pro Then
        msInsuredName = mudtXML01LossReport.XML01Loss.AdminLossInfo.ali0065_MainFileInsuredName
        msWorkPhone = mudtXML01LossReport.XML01Loss.AdminLossInfo.ali0058_BusPhone
        msHomePhone = mudtXML01LossReport.XML01Loss.AdminLossInfo.ali0057_HomePhone
    ElseIf mLossType = XML01Apd Then
        msInsuredName = mudtXML01LossReport.XML01Loss.AdminLossInfoApd.ali0064_NamedInsured
        msWorkPhone = mudtXML01LossReport.XML01Loss.AdminLossInfoApd.ali0055_BusPhone
        msHomePhone = mudtXML01LossReport.XML01Loss.AdminLossInfoApd.ali0054_HomePhone
    End If
    msDateAssign = mudtXML01LossReport.XML01Loss.CurrentLossInfo.cli04_DateAsgn
    msStatus = "PENDING"
    msTypeOfACID = "FACT"
    msAdjuster = "*ADJUSTERUSERNAME*"
    msACID = mudtXML01LossReport.XML01Loss.CurrentLossInfo.cli03_Adjuster
    msCLIENTNUM = mudtXML01LossReport.XML01Loss.CurrentLossInfo.cli06_SALN
    
    'If this is A file path leav it alone... For Export
    If InStr(1, mudtXML01LossReport.PrnKey, ":\") = 0 Then
        mudtXML01LossReport.PrnKey = "*ADJUSTERUSERNAME*" & "_" & msACID & "_" & "*IBNUMBER*" & "_" & msCLIENTNUM
    End If
    
CLEAN_UP:
    Set colCont = Nothing
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLR_ClassName & vbCrLf & "Private Sub PopulateMe"
End Sub

Private Function PopulateALI(pvLoss As Variant, plPos As Long) As Boolean
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    'Elements 1 - 51 include HP4000 Laser Header info, and Label info
    'Start of information will always begin with element 4 =Date and Time Printed
    plPos = 4
    If clsCarLR_LossType = XML01Pro Then
        With mudtXML01LossReport.XML01Loss.AdminLossInfo
            .ali0004_DateTimePrinted = pvLoss(plPos)
            .ali0005_PrintedBy = pvLoss(plPos + 1)
            plPos = 52 'Set pos to 52
            .ali0052_ReportedBy = pvLoss(plPos)
            .ali0053_RBPhone = pvLoss(plPos + 1)
            .ali0054_PolicyNum = pvLoss(plPos + 2)
            .ali0055_SC = pvLoss(plPos + 3)
            .ali0056_AgentNum = pvLoss(plPos + 4)
            .ali0057_HomePhone = pvLoss(plPos + 5)
            .ali0058_BusPhone = pvLoss(plPos + 6)
            .ali0059_NewBusDate = pvLoss(plPos + 7)
            .ali0060_RenewalDate = pvLoss(plPos + 8)
            .ali0061_LastCancDate = pvLoss(plPos + 9)
            .ali0062_NamedInsured = pvLoss(plPos + 10)
            .ali0063_MailAddress1 = pvLoss(plPos + 11)
            .ali0064_MailAddress2 = pvLoss(plPos + 12)
            .ali0065_MainFileInsuredName = pvLoss(plPos + 13)
            .ali0066_MortgageHolder = pvLoss(plPos + 14)
            .ali0067_2ndMort = pvLoss(plPos + 15)
            .ali0068_CompCode = pvLoss(plPos + 16)
            .ali0069_PolicyDescription = pvLoss(plPos + 17)
            .ali0070_BldgLimit = pvLoss(plPos + 18)
            .ali0071_ContLimit = pvLoss(plPos + 19)
            .ali0072_Deductible1 = pvLoss(plPos + 20)
            .ali0073_Deductible2 = pvLoss(plPos + 21)
            .ali0074_Deductible3 = pvLoss(plPos + 22)
            .ali0075_Deductible4 = pvLoss(plPos + 23)
            .ali0076_AddlCoverage1 = pvLoss(plPos + 24)
            .ali0077_AddlCoverage2 = pvLoss(plPos + 25)
            .ali0078_AddlCoverage3 = pvLoss(plPos + 26)
            .ali0079_AddlCoverage4 = pvLoss(plPos + 27)
            .ali0080_LossLocAddress1 = pvLoss(plPos + 28)
            .ali0081_LossLocAddress2 = pvLoss(plPos + 29)
        End With
        'Set lPos to next element 82= Start of ENdorsements if any
        plPos = 82
    ElseIf clsCarLR_LossType = XML01Apd Then
        With mudtXML01LossReport.XML01Loss.AdminLossInfoApd
            .ali0004_DateTimePrinted = pvLoss(plPos)
            .ali0005_PrintedBy = pvLoss(plPos + 1)
            '6 - 48 Labels and header
            plPos = 49 'Set pos to 49
            .ali0049_ReportedBy = pvLoss(plPos)
            .ali0050_RBPhone = pvLoss(plPos + 1)
            .ali0051_PolicyNum = pvLoss(plPos + 2)
            .ali0052_SC = pvLoss(plPos + 3)
            .ali0053_AgentNum = pvLoss(plPos + 4)
            .ali0054_HomePhone = pvLoss(plPos + 5)
            .ali0055_BusPhone = pvLoss(plPos + 6)
            '56 Spacer For where Mainfile insured name is in XML01 Property
            plPos = 57 'Set pos to 57
            .ali0057_MortgageHolder = pvLoss(plPos)
            .ali0058_CompCode = pvLoss(plPos + 1)
            .ali0059_PolicyType = pvLoss(plPos + 2)
            .ali0060_NewBusDate = pvLoss(plPos + 3)
            .ali0061_RenewalDate = pvLoss(plPos + 4)
            .ali0062_LastCancDate = pvLoss(plPos + 5)
            '63 This should be with CLI
            .cli0063_VehicleLocation = pvLoss(plPos + 6)
            .ali0064_NamedInsured = pvLoss(plPos + 7)
            .ali0065_MailAddress1 = pvLoss(plPos + 8)
            .ali0066_MailAddress2 = pvLoss(plPos + 9)
            .ali0067_VehicleDescription = pvLoss(plPos + 10)
            .ali0068_CompDed = pvLoss(plPos + 11)
            .ali0069_VIN = pvLoss(plPos + 12)
            'Set lPos to next element 70= Start of AddCov\ENdorsements for APD
            plPos = 70
        End With
    End If
    
    PopulateALI = True
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLR_ClassName & vbCrLf & "Private Function PopulateALI"
End Function

Private Function PopulateEDCol(pvLoss As Variant, plPos As Long) As Boolean
    On Error GoTo EH
    Dim Endorsement As udtXML01Endorsement
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    Do Until StrComp(pvLoss(plPos + 5), msCLIENTNUM, vbTextCompare) = 0
        If clsCarLR_LossType = XML01Pro Then
            'Set ED Code
            Endorsement.EDCode = pvLoss(plPos)
            
            'Set ED Description
            Endorsement.EDDescription = pvLoss(plPos + 1)
            
            plPos = plPos + 2 'increment lPos to next set of endorsement elements
        ElseIf clsCarLR_LossType = XML01Apd Then
            'Use Description for APD Endorsement
            Endorsement.EDDescription = pvLoss(plPos)
            plPos = plPos + 1 'increment lPos to next endorsement element
        End If
        
        'Add the Endorsement to the XML01 endorsement collection
        If mudtXML01LossReport.XML01Loss.colEndorsements Is Nothing Then
            Set mudtXML01LossReport.XML01Loss.colEndorsements = New Collection
        End If
        mudtXML01LossReport.XML01Loss.colEndorsements.Add Endorsement
    Loop
    
    PopulateEDCol = True
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLR_ClassName & vbCrLf & "Private Function PopulateEDCol"
End Function

Private Function PopulateCLI(pvLoss As Variant, plPos As Long, _
                             pbPLH As Boolean, _
                             pbCal As Boolean) As Boolean
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    With mudtXML01LossReport.XML01Loss.CurrentLossInfo
        .cli01_CAT = pvLoss(plPos)
        .cli02_LossDate = pvLoss(plPos + 1)
        .cli03_Adjuster = pvLoss(plPos + 2)
        .cli04_DateAsgn = pvLoss(plPos + 3)
        .cli05_DateClsd = pvLoss(plPos + 4)
        .cli06_SALN = pvLoss(plPos + 5)
        .cli06_SALN = goUtil.utScrubAlphaNumeric(.cli06_SALN)
        If clsCarLR_LossType = XML01Pro Then
            .cli07_AdjusterOrigInfo = pvLoss(plPos + 6)
            .cli08_DateAsgnOrigInfo = pvLoss(plPos + 7)
            .cli09_DateClsdOrigInfo = pvLoss(plPos + 8)
            
            'increment to Payments made this claim if it exists
            'Or
            'to either PLH if any PLH data exists
            'Or to CAL if any CAL data exists.
            plPos = plPos + 9
        ElseIf clsCarLR_LossType = XML01Apd Then
            plPos = plPos + 6
        End If
    
        'If there is '-' for prior SALN and '.' for Prior Paid amount Then we are ready for PLH
        'Or
        'If the Cat Code is Next then we are ready for CAL
        Do Until FoundPLHorCAL(pvLoss, plPos, .cli01_CAT, pbPLH, pbCal)
            'This field should be rarely if ever populated
            .cli10_PaymentsMadeThisClaim = .cli10_PaymentsMadeThisClaim & pvLoss(plPos) & vbTab
            plPos = plPos + 1
        Loop
    End With
    
    PopulateCLI = True
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLR_ClassName & vbCrLf & "Private Function PopulateCLI"
End Function

Private Function PopulatePLHCol(pvLoss As Variant, plPos As Long, psCAT As String, pbCal As Boolean, _
                                Optional pbCont As Boolean) As Boolean
    On Error GoTo EH
    Dim PLH As udtXML01PriorLossHist
    Dim pbPLH As Boolean
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    Do Until Not FoundPLH(pvLoss, plPos, psCAT, pbPLH, pbCal)
        PLH.plh01_SALN = pvLoss(plPos)
        PLH.plh02_LossDate = pvLoss(plPos + 1)
        PLH.plh03_Adjuster = pvLoss(plPos + 2)
        PLH.plh04_DateAsgn = pvLoss(plPos + 3)
        PLH.plh05_DateClsd = pvLoss(plPos + 4)
        PLH.plh06_AmtPaid = pvLoss(plPos + 5)

        plPos = plPos + 6 'increment lPos to next set of PLH elements

        'Add the PLH to the XML01 PLH collection
        If mudtXML01LossReport.XML01Loss.colPLH Is Nothing Then
            Set mudtXML01LossReport.XML01Loss.colPLH = New Collection
        End If
        mudtXML01LossReport.XML01Loss.colPLH.Add PLH
    Loop

    PopulatePLHCol = True
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLR_ClassName & vbCrLf & "Private Function PopulatePLHCol"
End Function

Private Function PopulateCALCol(pvLoss As Variant, plPos As Long, psCAT As String, _
                                Optional pbCont As Boolean) As Boolean
    On Error GoTo EH
    Dim CAL As udtXML01CommentsActLog
    Dim pbPLH As Boolean
    Dim pbCal As Boolean
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    Do Until Not FoundCAL(pvLoss, plPos, psCAT, pbPLH, pbCal)
        CAL.cal01_CAT = pvLoss(plPos)
        CAL.cal02_Date = pvLoss(plPos + 1)
        CAL.cal03_Time = pvLoss(plPos + 2)
        CAL.cal04_Action = pvLoss(plPos + 3)
        CAL.cal05_User = pvLoss(plPos + 4)
        CAL.cal06_Comments = vbNullString 'Clear it first
        plPos = plPos + 5 'increment lPos to first element of comments
        
        'Comments can take up more than 1 element of data
        Do Until StrComp(pvLoss(plPos), psCAT, vbTextCompare) = 0 Or StrComp(pvLoss(plPos), zCAL_CAT_ALL, vbTextCompare) = 0 Or StrComp(pvLoss(plPos), zLBL_SALN, vbTextCompare) = 0
            CAL.cal06_Comments = CAL.cal06_Comments & pvLoss(plPos) & vbCrLf
            plPos = plPos + 1 'increment lPos to next element
        Loop

        'Add the CAL to the XML01 CAL collection
        If mudtXML01LossReport.XML01Loss.colCAL Is Nothing Then
            Set mudtXML01LossReport.XML01Loss.colCAL = New Collection
        End If
        mudtXML01LossReport.XML01Loss.colCAL.Add CAL
        
    Loop
    
    PopulateCALCol = True
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLR_ClassName & vbCrLf & "Private Function PopulateCALCol"
End Function

Private Function PopulateLossContCol(pcolCont As Collection) As Boolean
    On Error GoTo EH
    Dim LossCont As udtXML01LossCont
    Dim vCont As Variant 'use this to loop through pcolCont
    Dim PLHCont As udtXML01PriorLossHistCont
    Dim CAL As udtXML01CommentsActLog
    Dim lPos As Long
    Dim sCAT As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    'The Loss continuations always start off on pos 19
    lPos = 19
    
    'Set the CAT
    sCAT = mudtXML01LossReport.XML01Loss.CurrentLossInfo.cli01_CAT
    
    For Each vCont In pcolCont
        Do Until Not FoundPLHCont(vCont, lPos)
            PLHCont.plh01_CAT = vCont(lPos)
            PLHCont.plh02_LossDate = vCont(lPos + 1)
            PLHCont.plh03_Adjuster = vCont(lPos + 2)
            PLHCont.plh04_DateAsgn = vCont(lPos + 3)
            PLHCont.plh05_DateClsd = vCont(lPos + 4)
            PLHCont.plh06_AmtPaid = vCont(lPos + 5)
            PLHCont.plh07_SALN = vCont(lPos + 6)
            
            lPos = lPos + 7 'increment lPos to next element of PLH
    
            'Add the PLHCont to the PLHCont collection
            If LossCont.colPLHCont Is Nothing Then
                Set LossCont.colPLHCont = New Collection
            End If
            LossCont.colPLHCont.Add PLHCont
    
        Loop
        
        Do Until Not FoundCALCont(vCont, lPos, sCAT)
            CAL.cal01_CAT = vCont(lPos)
            CAL.cal02_Date = vCont(lPos + 1)
            CAL.cal03_Time = vCont(lPos + 2)
            CAL.cal04_Action = vCont(lPos + 3)
            CAL.cal05_User = vCont(lPos + 4)
            CAL.cal06_Comments = vbNullString 'Clear it first
            
            lPos = lPos + 5 'increment lPos to first element of comments
            
            'Comments can take up more than 1 element of data
            If lPos <= UBound(vCont, 1) Then
                Do Until StrComp(vCont(lPos), sCAT, vbTextCompare) = 0 Or StrComp(vCont(lPos), zCAL_CAT_ALL, vbTextCompare) = 0
                    CAL.cal06_Comments = CAL.cal06_Comments & vCont(lPos) & vbCrLf
                    'Bail if we exceed the elements of data
                    If lPos + 1 > UBound(vCont, 1) Then
                        Exit Do
                    End If
                    lPos = lPos + 1 'increment lPos to next element
                Loop
            End If
    
            'Add the CALCont to the CAL collection
            If LossCont.colCAL Is Nothing Then
                Set LossCont.colCAL = New Collection
            End If
            LossCont.colCAL.Add CAL
            
        Loop
        
        'Don't print blank Continuation pages even if Farmers HP4000 Laser print job does
        If lPos <= UBound(vCont, 1) Then
            'Add the LossCont to the XML01 Cont collection
            If mudtXML01LossReport.colXML01LossCont Is Nothing Then
                Set mudtXML01LossReport.colXML01LossCont = New Collection
            End If
            mudtXML01LossReport.colXML01LossCont.Add LossCont
        End If
        
    Next
    
    PopulateLossContCol = True
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLR_ClassName & vbCrLf & "Private Function PopulateLossContCol"
End Function

Private Function FoundPLH(pvLoss As Variant, plPos As Long, psCAT As String, _
                          pbPLH As Boolean, _
                          pbCal As Boolean) As Boolean
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    pbPLH = False
    pbCal = False
    FoundPLHorCAL pvLoss, plPos, psCAT, pbPLH, pbCal
    FoundPLH = pbPLH
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLR_ClassName & vbCrLf & "Private Function FoundPLH"
End Function

Private Function FoundCAL(pvLoss As Variant, plPos As Long, psCAT As String, _
                          pbPLH As Boolean, _
                          pbCal As Boolean) As Boolean
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    pbPLH = False
    pbCal = False
    FoundPLHorCAL pvLoss, plPos, psCAT, pbPLH, pbCal
    FoundCAL = pbCal
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLR_ClassName & vbCrLf & "Private Function FoundCAL"
End Function

Private Function FoundPLHorCAL(pvLoss As Variant, plPos As Long, psCAT As String, _
                               pbPLH As Boolean, _
                               pbCal As Boolean) As Boolean
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    'Check for Prior Loss History
    'If there is '-' for prior SALN and '.' for Prior Paid amount
    'Then we are ready for PLH
    
    
    If (InStr(1, pvLoss(plPos), "-", vbTextCompare) > 0 Or InStr(1, pvLoss(plPos + 5), ".", vbTextCompare) > 0) _
    And InStr(1, pvLoss(plPos + 2), ":", vbTextCompare) = 0 Then
        pbPLH = True
        FoundPLHorCAL = pbPLH
    End If
                          
    'Or
    
    'If the Cat Code is Next then we are ready for CAL
    If StrComp(pvLoss(plPos), psCAT, vbTextCompare) = 0 Or StrComp(pvLoss(plPos), zCAL_CAT_ALL, vbTextCompare) = 0 Then
        pbCal = True
        FoundPLHorCAL = pbCal
        pbPLH = False 'If we find CAL then Flag PLH as false
    End If
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLR_ClassName & vbCrLf & "Private Function FoundPLHorCAL"
End Function

Private Function FoundPLHCont(pvCont As Variant, plPos As Long) As Boolean
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    'Check for Prior Loss History for Continuation page
    'If there is '-' for prior SALN or '.' for Prior Amt Paid
    'Then we are ready for PLHCont
    If plPos + 6 > UBound(pvCont, 1) Then
        'Bail if we exceed data elements
        Exit Function
    End If
    If (InStr(1, pvCont(plPos + 5), ".", vbTextCompare) > 0 Or InStr(1, pvCont(plPos + 6), "-", vbTextCompare) > 0) _
    And InStr(1, pvCont(plPos + 2), ":", vbTextCompare) = 0 Then
        FoundPLHCont = True
    End If
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLR_ClassName & vbCrLf & "Private Function FoundPLHCont"
End Function

Private Function FoundCALCont(pvCont As Variant, plPos As Long, psCAT As String) As Boolean
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    If plPos > UBound(pvCont, 1) Then
        'Bail if we exceed data elements
        Exit Function
    End If
    'If the Cat Code is Next then we are ready for CAL
    If StrComp(pvCont(plPos), psCAT, vbTextCompare) = 0 Or StrComp(pvCont(plPos), zCAL_CAT_ALL, vbTextCompare) = 0 Then
        FoundCALCont = True
    End If
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLR_ClassName & vbCrLf & "Private Function FoundPLHCont"
End Function

Private Function GetLossData(pvLossP1 As Variant, pcolCont As Collection) As Boolean
    On Error GoTo EH
    Dim vData As Variant
    Dim vLoss As Variant
    Dim lcount As Long
    Dim lCount2 As Long
    Dim bContFlag As Boolean 'true if we are adding Continuation page
    Dim colCont As Collection 'contains collection of Loss continuation pages
    Dim vLossP1 As Variant
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    'Need to determine type Of XML01
    If InStr(1, Me.clsCarLR_PrnData, zSTART_JOB, vbTextCompare) > 0 Then
        clsCarLR_LossType = TypeXML01.XML01Pro
    ElseIf InStr(1, Me.clsCarLR_PrnData, zSTART_JOBApd, vbTextCompare) > 0 Then
        clsCarLR_LossType = TypeXML01.XML01Apd
    End If
    
    vData = Split(Me.clsCarLR_PrnData, zEOJ)
    
    For lcount = LBound(vData, 1) To UBound(vData, 1)
        If vData(lcount) = vbNullString Then
            GoTo NEXT_LOSS
        End If
        If bContFlag Then
            'BGS 2.3.2004 Fixed Bug
            'Was seting the collection new even if it was already set
            'If there were more than 1 continuation pages only the last one would
            'Show up.
            If colCont Is Nothing Then
                Set colCont = New Collection
            End If
        End If
        
        vLoss = Split(vData(lcount), zDM)
        'Get rid of HP 4000 Laser Printer format data just leave Text Data
        For lCount2 = LBound(vLoss, 1) To UBound(vLoss, 1)
            If InStr(1, vLoss(lCount2), Chr(180)) > 0 Then
                vLoss(lCount2) = Left(vLoss(lCount2), InStr(1, vLoss(lCount2), Chr(180)) - 1)
            End If
        Next
        
        If bContFlag Then
            colCont.Add vLoss 'Add to Cont collection
        Else
            vLossP1 = vLoss 'Set the Page 1 variant
        End If
        bContFlag = True 'After we process first element any others will be Continuation pages
NEXT_LOSS:
    Next
    
    Set pcolCont = colCont
    pvLossP1 = vLossP1
    
    GetLossData = True
    
    'Clean up
    Set colCont = Nothing
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Set colCont = Nothing
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLR_ClassName & vbCrLf & "Private Function GetLossData"
End Function

Private Function GetLossCol(pvaryData As Variant, pbCorruptFile As Boolean) As Collection

    On Error GoTo EH
    Dim sTemp As String
    Dim sData As String
    Dim bAddJob As Boolean
    Dim lcount As Long
    Dim colLoss As Collection
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    'Need to check for start of job and continuations of the
    'start jobs to be included in the Loss collection.
    
    For lcount = LBound(pvaryData, 1) To UBound(pvaryData, 1)
        'Be sure we append zEOJ to each Job
        'We need to do this because we used
        'zEOJ to DELIM into Array.
        sTemp = pvaryData(lcount)
        If sTemp = vbNullString Then
            GoTo NEXT_LOSS
        Else
            sTemp = sTemp & zEOJ
        End If
        
        If InStr(1, sTemp, zCONTINUED_JOB, vbTextCompare) > 0 Then
            'Append Continued job
            sData = sData & sTemp
        ElseIf InStr(1, sTemp, zSTART_JOB, vbTextCompare) > 0 Then
            'be sure we add ClassName to top line of data
            sData = clsCarLR_ClassName & vbCrLf & sTemp
            clsCarLR_LossType = XML01Pro
        ElseIf InStr(1, sTemp, zSTART_JOBApd, vbTextCompare) > 0 Then
            'be sure we add ClassName to top line of data
            sData = clsCarLR_ClassName & vbCrLf & sTemp
            clsCarLR_LossType = XML01Apd
        Else
            'If we get here there is some other XML01 format other than
            'Property.
            sData = "clsLossUnknown" & vbCrLf & sTemp
        End If
        
        'Check the next element to be sure it is not
        'a continuation job.  If it is we will not
        'add this loss inside the collection until
        'we have all of its continuations added on.
        If lcount + 1 <= UBound(pvaryData, 1) Then
            sTemp = pvaryData(lcount + 1)
            If InStr(1, sTemp, zCONTINUED_JOB) > 0 Then
                bAddJob = False
            Else
                bAddJob = True
            End If
        Else
            bAddJob = True
        End If
        
        If bAddJob Then
            If colLoss Is Nothing Then
                Set colLoss = New Collection
            End If
            colLoss.Add sData
            sData = vbNullString
        End If
NEXT_LOSS:
    Next
    If colLoss Is Nothing Then
        pbCorruptFile = True
    End If
    Set GetLossCol = colLoss
    
    'clean up local memory
    Set colLoss = Nothing
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLR_ClassName & vbCrLf & "Private Function GetLossCol"
End Function

Private Function PrintXML01(psPrnDeviceName As String, pbPreview As Boolean, _
                           pbPDF As Boolean, plHwnd As Long, plChainType As Long) As Boolean
'    On Error GoTo EH
'    Dim MyXML01 As Object
'    Dim sPDFPath As String
'    Dim sCaption As String
'    Dim dDays As Double
'    Dim lErrNum As Long
'    Dim sErrDesc As String
'
'    'Build Caption
'    sCaption = "Loss Report " & Chr(160) & " (" & mudtXML01LossReport.PrnKey & ")"
'
'    'Set the Report Here
'    If clsCarLR_LossType = XML01Pro Then
'        Set MyXML01 = New arLossXML01
'    ElseIf clsCarLR_LossType = XML01Apd Then
'        Set MyXML01 = New arLossXML01Apd
'    End If
'
'    With MyXML01
'        .Printer.Orientation = ddOPortrait
'        On Error Resume Next
'        .Printer.DeviceName = psPrnDeviceName
'        If Err.Number > 0 Then
'            Err.Clear
'            On Error GoTo EH
'        End If
'        .LossReport = mudtXML01LossReport
'    End With
'
'    'Check to see if we are bulding a Chain
'    If plChainType = ChainType.FirstInChain Or plChainType = ChainType.NextLink Then
'        moLRs.BuildChain MyXML01, plChainType, sCaption, clsCarLR_Adjuster
'        GoTo CLEANUP
'    End If
'    'If we are previewing then we need to use ARV object
'    If pbPreview Then
'        If goUtil.gARV Is Nothing Then
'            Set goUtil.gARV = New V2ARViewer.clsARViewer
'        End If
'
'        With goUtil.gARV
'            .SetUtilObject goUtil
'            MyXML01.Run 'Run it before it gets to ARV object
'            .objARvReport = MyXML01
'            .sRptTitle = sCaption
'            .ShowReport vbModeless
'        End With
'        Screen.MousePointer = vbDefault
'    Else 'If we not preview then check to see if this is going to PDF
'        If pbPDF Then
'            'Set PDF File
'            sPDFPath = Trim(moLRs.LossReportsForm.txtPDFPath.Text)
'            sPDFPath = sPDFPath & sCaption
'            moLRs.ExportFile MyXML01, sPDFPath, ExportType.ARPdf 'Export to PDF
'        Else
'            MyXML01.PrintReport False 'Print without showing print dialog
'        End If
'    End If
'
'CLEANUP:
'    PrintXML01 = True
'
'    'Clean up
'    Set MyXML01 = Nothing
'    Exit Function
'EH:
'    lErrNum = Err.Number
'    sErrDesc = Err.Description
'    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLR_ClassName & vbCrLf & "Private Function PrintXML01"
End Function
Public Function clsCarLR_ExportME(poLRs As V2ECKeyBoard.clsLossReports, pXportType As ExportType) As Boolean
'    On Error GoTo EH
'    Dim MyXML01 As Object
'    Dim sXportPath As String
'    Dim lErrNum As Long
'    Dim sErrDesc As String
'
'    'Set the Report Here
'    If clsCarLR_LossType = XML01Pro Then
'        Set MyXML01 = New arLossXML01
'    ElseIf clsCarLR_LossType = XML01Apd Then
'        Set MyXML01 = New arLossXML01Apd
'    End If
'
'    With MyXML01
'        .Printer.Orientation = ddOPortrait
'        .LossReport = mudtXML01LossReport
'    End With
'
'    sXportPath = clsCarLR_PrnKey
'    If poLRs.ExportFile(MyXML01, sXportPath, pXportType) Then
'        clsCarLR_ExportME = True
'    End If
'
'    'Clean up
'    Set MyXML01 = Nothing
'
'    Exit Function
'EH:
'    lErrNum = Err.Number
'    sErrDesc = Err.Description
'    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLR_ClassName & vbCrLf & "Public Function ExportME"
End Function

Private Function SQL_WEBCONTROL_108() As String
    Dim sSQL As String
    Dim sTemp As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    'ContactDate
    'CloseDate
    sSQL = "INSERT INTO ASSIGNMENTS (Client, "      '01Client
    sSQL = sSQL & "ClaimNoSaln, "                   '02ClaimNoSaln
    sSQL = sSQL & "DateAssigned, "        'Date/Time'03DateAssigned
    sSQL = sSQL & "Adjuster, "                      '04Adjuster
    sSQL = sSQL & "Company, "                       '05Company
    sSQL = sSQL & "TypeOfLoss, "                    '06TypeOfLoss
    sSQL = sSQL & "StateNo, "                       '07StateNo
    sSQL = sSQL & "PolicyNo, "                      '08PolicyNo
    sSQL = sSQL & "TexasSuffix, "                   '09TexasSuffix
    sSQL = sSQL & "CatCode, "                       '10CatCode
    sSQL = sSQL & "PolicyDescription, "             '11PolicyDescription
    sSQL = sSQL & "Insured, "                       '12Insured
    sSQL = sSQL & "Address, "                       '13Address
    sSQL = sSQL & "HomePhone, "                     '14HomePhone
    sSQL = sSQL & "BusinessPhone, "                 '14aBusinessPhone
    sSQL = sSQL & "PropertyAddress, "               '15PropertyAddress
    sSQL = sSQL & "MortgageeName, "                 '16MortgageeName
    sSQL = sSQL & "LoanNo, "                        '17LoanNo
    sSQL = sSQL & "MtgAddress, "                    '18MtgAddress
    sSQL = sSQL & "MtgCode, "                       '19MtgCode
    sSQL = sSQL & "AgentLR, "                       '20AgentLR
    sSQL = sSQL & "StateCD, "                       '21StateCD
    sSQL = sSQL & "District, "                      '22District
    sSQL = sSQL & "AgentNo, "                       '23AgentNo
    sSQL = sSQL & "ReportedBy, "                    '24ReportedBy
    sSQL = sSQL & "ReportedByPhone, "               '25ReportedByPhone
    sSQL = sSQL & "DateReportedToAgent, " 'Date/Time'26DateReportedToAgent
    sSQL = sSQL & "DateReportedByAgent, " 'Date/Time'27DateReportedByAgent
    sSQL = sSQL & "LossDate, "            'Date/Time'28LossDate
    sSQL = sSQL & "LossLocation, "                  '29LossLocation
    sSQL = sSQL & "BalanceDue, "                    '30BalanceDue
    sSQL = sSQL & "MFRec, "                         '31MFRec
    sSQL = sSQL & "RenewalDate, "         'Date/Time'32RenewalDate
    sSQL = sSQL & "NewBusReinDt, "        'Date/Time'33NewBusReinDt
    sSQL = sSQL & "BuildingLimits, "            'CUR'34BuildingLimits
    sSQL = sSQL & "ContentsLimits, "            'CUR'35ContentsLimits
    sSQL = sSQL & "Deductibles, "                   '36Deductibles
    sSQL = sSQL & "Format, "
    sSQL = sSQL & "LossReport ) "               'Memo'37LossReport
    With mudtXML01LossReport.XML01Loss
        If clsCarLR_LossType = XML01Pro Then
            'Use Main file Insured if there
            sSQL = sSQL & "VALUES (" & S_z & IIf(.AdminLossInfo.ali0065_MainFileInsuredName = vbNullString, .AdminLossInfo.ali0062_NamedInsured, .AdminLossInfo.ali0065_MainFileInsuredName) & z_S '01Client
            sSQL = sSQL & S_z & .CurrentLossInfo.cli06_SALN & z_S                                                   '02ClaimNoSaln
            sSQL = sSQL & IIf(IsDate(.CurrentLossInfo.cli04_DateAsgn), "#" & .CurrentLossInfo.cli04_DateAsgn & "#", "Null") & ", "  '03DateAssigned
            sSQL = sSQL & S_z & .CurrentLossInfo.cli03_Adjuster & z_S                                           '04Adjuster
            sSQL = sSQL & S_z & .AdminLossInfo.ali0068_CompCode & z_S                                           '05Company
            sSQL = sSQL & S_z & vbNullString & z_S                                                              '06TypeOfLoss
            sSQL = sSQL & S_z & .AdminLossInfo.ali0055_SC & z_S                                                 '07StateNo
            sSQL = sSQL & S_z & .AdminLossInfo.ali0054_PolicyNum & z_S                                          '08PolicyNo
            sSQL = sSQL & S_z & vbNullString & z_S                                                              '09TexasSuffix
            sSQL = sSQL & S_z & .CurrentLossInfo.cli01_CAT & z_S                                                '10CatCode
            sSQL = sSQL & S_z & .AdminLossInfo.ali0069_PolicyDescription & z_S                                  '11PolicyDescription
            sSQL = sSQL & S_z & .AdminLossInfo.ali0062_NamedInsured & z_S                                       '12Insured
            sSQL = sSQL & S_z & .AdminLossInfo.ali0063_MailAddress1 & " " & .AdminLossInfo.ali0064_MailAddress2 & z_S '13Address
            sSQL = sSQL & S_z & .AdminLossInfo.ali0057_HomePhone & z_S                                          '14HomePhone
            sSQL = sSQL & S_z & .AdminLossInfo.ali0058_BusPhone & z_S                                           '14aBusinessPhone
            sSQL = sSQL & S_z & .AdminLossInfo.ali0080_LossLocAddress1 & " " & .AdminLossInfo.ali0081_LossLocAddress2 & z_S '15PropertyAddress
            sSQL = sSQL & S_z & .AdminLossInfo.ali0066_MortgageHolder & z_S                                     '16MortgageeName
            sSQL = sSQL & S_z & .AdminLossInfo.ali0067_2ndMort & z_S                                            '17LoanNo
            sSQL = sSQL & S_z & vbNullString & z_S                                                              '18MtgAddress
            sSQL = sSQL & S_z & vbNullString & z_S                                                              '19MtgCode
            sSQL = sSQL & S_z & vbNullString & z_S                                                              '20AgentLR
            sSQL = sSQL & S_z & .AdminLossInfo.ali0055_SC & z_S                                                 '21StateCD
            sSQL = sSQL & S_z & vbNullString & z_S                                                              '22District
            sSQL = sSQL & S_z & .AdminLossInfo.ali0056_AgentNum & z_S                                           '23AgentNo
            sSQL = sSQL & S_z & .AdminLossInfo.ali0052_ReportedBy & z_S                                          '24ReportedBy
            sSQL = sSQL & S_z & .AdminLossInfo.ali0053_RBPhone & z_S                                            '25ReportedByPhone
            sSQL = sSQL & "Null, "                                                                              '26DateReportedToAgent
            sSQL = sSQL & "Null, "                                                                              '27DateReportedByAgent
            sSQL = sSQL & IIf(IsDate(.CurrentLossInfo.cli02_LossDate), "#" & .CurrentLossInfo.cli02_LossDate & "#", "Null") & ", "  '28LossDate
            sSQL = sSQL & S_z & .AdminLossInfo.ali0080_LossLocAddress1 & " " & .AdminLossInfo.ali0081_LossLocAddress2 & z_S '29LossLocation
            sSQL = sSQL & S_z & vbNullString & z_S                                                              '30BalanceDue
            sSQL = sSQL & S_z & vbNullString & z_S                                                              '31MFRec
            sSQL = sSQL & IIf(IsDate(.AdminLossInfo.ali0060_RenewalDate), "#" & .AdminLossInfo.ali0060_RenewalDate & "#", "Null") & ", " '32RenewalDate
            sSQL = sSQL & IIf(IsDate(.AdminLossInfo.ali0059_NewBusDate), "#" & .AdminLossInfo.ali0059_NewBusDate & "#", "Null") & ", " '33NewBusReinDt
            'Val function Bug in VB6
            'http://msdn.microsoft.com/vbasic/productinfo/previous/vb6/tips/01pasttips.asp
            'Need to parse out both % and !  because these trailing equate to Double and Single
            'and Val bugs because it can't convert Double or single into integer
            sTemp = .AdminLossInfo.ali0070_BldgLimit
            sSQL = sSQL & Val(goUtil.utCleanValString(sTemp)) & ", "                                          '34BuildingLimits
            sTemp = .AdminLossInfo.ali0071_ContLimit
            sSQL = sSQL & Val(goUtil.utCleanValString(sTemp)) & ", "                                          '35ContentsLimits
            'Use the First Deductible by default
            sSQL = sSQL & S_z & LTrim(Right(.AdminLossInfo.ali0072_Deductible1, 10)) & z_S                      '36Deductibles
            sSQL = sSQL & S_z & clsCarLR_ClassName & z_S
            sSQL = sSQL & S_z & clsCarLR_ClassName & vbCrLf & clsCarLR_PrnData & S_z & " ) "                                '37LossReport
        ElseIf clsCarLR_LossType = XML01Apd Then
            sSQL = sSQL & "VALUES (" & S_z & .AdminLossInfoApd.ali0064_NamedInsured & z_S '01Client
            sSQL = sSQL & S_z & .CurrentLossInfo.cli06_SALN & z_S                                                   '02ClaimNoSaln
            sSQL = sSQL & IIf(IsDate(.CurrentLossInfo.cli04_DateAsgn), "#" & .CurrentLossInfo.cli04_DateAsgn & "#", "Null") & ", "  '03DateAssigned
            sSQL = sSQL & S_z & .CurrentLossInfo.cli03_Adjuster & z_S                                           '04Adjuster
            sSQL = sSQL & S_z & .AdminLossInfoApd.ali0058_CompCode & z_S                                           '05Company
            sSQL = sSQL & S_z & vbNullString & z_S                                                              '06TypeOfLoss
            sSQL = sSQL & S_z & .AdminLossInfoApd.ali0052_SC & z_S                                                 '07StateNo
            sSQL = sSQL & S_z & .AdminLossInfoApd.ali0051_PolicyNum & z_S                                          '08PolicyNo
            sSQL = sSQL & S_z & vbNullString & z_S                                                              '09TexasSuffix
            sSQL = sSQL & S_z & .CurrentLossInfo.cli01_CAT & z_S                                                '10CatCode
            sSQL = sSQL & S_z & .AdminLossInfoApd.ali0059_PolicyType & z_S                                  '11PolicyDescription
            sSQL = sSQL & S_z & .AdminLossInfoApd.ali0064_NamedInsured & z_S                                       '12Insured
            sSQL = sSQL & S_z & .AdminLossInfoApd.ali0065_MailAddress1 & " " & .AdminLossInfoApd.ali0066_MailAddress2 & z_S '13Address
            sSQL = sSQL & S_z & .AdminLossInfoApd.ali0054_HomePhone & z_S                                          '14HomePhone
            sSQL = sSQL & S_z & .AdminLossInfoApd.ali0055_BusPhone & z_S                                           '14aBusinessPhone
            sSQL = sSQL & S_z & .AdminLossInfoApd.cli0063_VehicleLocation & z_S                                 '15PropertyAddress
            sSQL = sSQL & S_z & .AdminLossInfoApd.ali0057_MortgageHolder & z_S                                     '16MortgageeName
            sSQL = sSQL & S_z & vbNullString & z_S                                            '17LoanNo
            sSQL = sSQL & S_z & vbNullString & z_S                                                              '18MtgAddress
            sSQL = sSQL & S_z & vbNullString & z_S                                                              '19MtgCode
            sSQL = sSQL & S_z & vbNullString & z_S                                                              '20AgentLR
            sSQL = sSQL & S_z & .AdminLossInfoApd.ali0052_SC & z_S                                                  '21StateCD
            sSQL = sSQL & S_z & vbNullString & z_S                                                              '22District
            sSQL = sSQL & S_z & .AdminLossInfoApd.ali0053_AgentNum & z_S                                           '23AgentNo
            sSQL = sSQL & S_z & .AdminLossInfoApd.ali0049_ReportedBy & z_S                                          '24ReportedBy
            sSQL = sSQL & S_z & .AdminLossInfoApd.ali0050_RBPhone & z_S                                            '25ReportedByPhone
            sSQL = sSQL & "Null, "                                                                              '26DateReportedToAgent
            sSQL = sSQL & "Null, "                                                                              '27DateReportedByAgent
            sSQL = sSQL & IIf(IsDate(.CurrentLossInfo.cli02_LossDate), "#" & .CurrentLossInfo.cli02_LossDate & "#", "Null") & ", "  '28LossDate
            sSQL = sSQL & S_z & .AdminLossInfoApd.cli0063_VehicleLocation & z_S  '29LossLocation
            sSQL = sSQL & S_z & vbNullString & z_S                                                              '30BalanceDue
            sSQL = sSQL & S_z & vbNullString & z_S                                                              '31MFRec
            sSQL = sSQL & IIf(IsDate(.AdminLossInfoApd.ali0061_RenewalDate), "#" & .AdminLossInfoApd.ali0061_RenewalDate & "#", "Null") & ", " '32RenewalDate
            sSQL = sSQL & IIf(IsDate(.AdminLossInfoApd.ali0060_NewBusDate), "#" & .AdminLossInfoApd.ali0060_NewBusDate & "#", "Null") & ", " '33NewBusReinDt
            'Val function Bug in VB6
            'http://msdn.microsoft.com/vbasic/productinfo/previous/vb6/tips/01pasttips.asp
            'Need to parse out both % and !  because these trailing equate to Double and Single
            'and Val bugs because it can't convert Double or single into integer
            sTemp = 0
            sSQL = sSQL & Val(goUtil.utCleanValString(sTemp)) & ", "                                          '34BuildingLimits
            sTemp = 0
            sSQL = sSQL & Val(goUtil.utCleanValString(sTemp)) & ", "                                          '35ContentsLimits
            'Use the First Deductible by default
            sSQL = sSQL & S_z & LTrim(Right(.AdminLossInfoApd.ali0068_CompDed, 10)) & z_S                      '36Deductibles
            sSQL = sSQL & S_z & clsCarLR_ClassName & z_S
            sSQL = sSQL & S_z & clsCarLR_ClassName & vbCrLf & clsCarLR_PrnData & S_z & " ) "
        End If
    End With
    
    SQL_WEBCONTROL_108 = sSQL
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    sErrDesc = mudtXML01LossReport.XML01Loss.CurrentLossInfo.cli06_SALN & vbCrLf & sErrDesc
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLR_ClassName & vbCrLf & "Private Function SQL_WEBCONTROL_108"
End Function

Private Function SQL_WEBCONTROL_20000() As String
    Dim sSQL As String
    Dim sUserName As String
    Dim sTemp As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim bAssignByZip As Boolean
    
    moLRs.FireGetAssignByZip bAssignByZip
       
    sUserName = goUtil.utGetECSCryptSetting("V2WebControl", "DBConn", "USERID")
        
    With mudtXML01LossReport.XML01Loss
        If clsCarLR_LossType = XML01Pro Then
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
            sSQL = sSQL & S_z & .AdminLossInfo.ali0004_DateTimePrinted & z_S
            sSQL = sSQL & S_z & .AdminLossInfo.ali0005_PrintedBy & z_S
            sSQL = sSQL & S_z & .AdminLossInfo.ali0052_ReportedBy & z_S
            sSQL = sSQL & S_z & .AdminLossInfo.ali0053_RBPhone & z_S
            sSQL = sSQL & S_z & .AdminLossInfo.ali0054_PolicyNum & z_S
            sSQL = sSQL & S_z & .AdminLossInfo.ali0055_SC & z_S
            sSQL = sSQL & S_z & .AdminLossInfo.ali0056_AgentNum & z_S
            sSQL = sSQL & S_z & .AdminLossInfo.ali0057_HomePhone & z_S
            sSQL = sSQL & S_z & .AdminLossInfo.ali0058_BusPhone & z_S
            sSQL = sSQL & S_z & .AdminLossInfo.ali0059_NewBusDate & z_S
            sSQL = sSQL & S_z & .AdminLossInfo.ali0060_RenewalDate & z_S
            sSQL = sSQL & S_z & .AdminLossInfo.ali0061_LastCancDate & z_S
            sSQL = sSQL & S_z & .AdminLossInfo.ali0062_NamedInsured & z_S
            sSQL = sSQL & S_z & .AdminLossInfo.ali0063_MailAddress1 & z_S
            sSQL = sSQL & S_z & .AdminLossInfo.ali0064_MailAddress2 & z_S
            sSQL = sSQL & S_z & .AdminLossInfo.ali0065_MainFileInsuredName & z_S
            sSQL = sSQL & S_z & .AdminLossInfo.ali0066_MortgageHolder & z_S
            sSQL = sSQL & S_z & .AdminLossInfo.ali0067_2ndMort & z_S
            sSQL = sSQL & S_z & .AdminLossInfo.ali0068_CompCode & z_S
            sSQL = sSQL & S_z & .AdminLossInfo.ali0069_PolicyDescription & z_S
            sTemp = .AdminLossInfo.ali0070_BldgLimit
            sSQL = sSQL & S_z & Val(goUtil.utCleanValString(sTemp)) & z_S
            sTemp = .AdminLossInfo.ali0071_ContLimit
            sSQL = sSQL & S_z & Val(goUtil.utCleanValString(sTemp)) & z_S
            sTemp = LTrim(Right(.AdminLossInfo.ali0072_Deductible1, 10))
            sSQL = sSQL & S_z & Val(goUtil.utCleanValString(sTemp)) & z_S
            sTemp = LTrim(Right(.AdminLossInfo.ali0073_Deductible2, 10))
            sSQL = sSQL & S_z & Val(goUtil.utCleanValString(sTemp)) & z_S
            sTemp = LTrim(Right(.AdminLossInfo.ali0074_Deductible3, 10))
            sSQL = sSQL & S_z & Val(goUtil.utCleanValString(sTemp)) & z_S
            sTemp = LTrim(Right(.AdminLossInfo.ali0075_Deductible4, 10))
            sSQL = sSQL & S_z & Val(goUtil.utCleanValString(sTemp)) & z_S
            sSQL = sSQL & S_z & .AdminLossInfo.ali0076_AddlCoverage1 & z_S
            sSQL = sSQL & S_z & .AdminLossInfo.ali0077_AddlCoverage2 & z_S
            sSQL = sSQL & S_z & .AdminLossInfo.ali0078_AddlCoverage3 & z_S
            sSQL = sSQL & S_z & .AdminLossInfo.ali0079_AddlCoverage4 & z_S
            sSQL = sSQL & S_z & .AdminLossInfo.ali0080_LossLocAddress1 & z_S
            sSQL = sSQL & S_z & .AdminLossInfo.ali0081_LossLocAddress2 & z_S
        ElseIf clsCarLR_LossType = XML01Apd Then
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
            sSQL = sSQL & S_z & .AdminLossInfoApd.ali0004_DateTimePrinted & z_S
            sSQL = sSQL & S_z & .AdminLossInfoApd.ali0005_PrintedBy & z_S
            sSQL = sSQL & S_z & .AdminLossInfoApd.ali0049_ReportedBy & z_S
            sSQL = sSQL & S_z & .AdminLossInfoApd.ali0050_RBPhone & z_S
            sSQL = sSQL & S_z & .AdminLossInfoApd.ali0051_PolicyNum & z_S
            sSQL = sSQL & S_z & .AdminLossInfoApd.ali0052_SC & z_S
            sSQL = sSQL & S_z & .AdminLossInfoApd.ali0053_AgentNum & z_S
            sSQL = sSQL & S_z & .AdminLossInfoApd.ali0054_HomePhone & z_S
            sSQL = sSQL & S_z & .AdminLossInfoApd.ali0055_BusPhone & z_S
            sSQL = sSQL & S_z & .AdminLossInfoApd.ali0057_MortgageHolder & z_S
            sSQL = sSQL & S_z & .AdminLossInfoApd.ali0058_CompCode & z_S
            sSQL = sSQL & S_z & .AdminLossInfoApd.ali0059_PolicyType & z_S
            sSQL = sSQL & S_z & .AdminLossInfoApd.ali0060_NewBusDate & z_S
            sSQL = sSQL & S_z & .AdminLossInfoApd.ali0061_RenewalDate & z_S
            sSQL = sSQL & S_z & .AdminLossInfoApd.ali0062_LastCancDate & z_S
            sSQL = sSQL & S_z & .AdminLossInfoApd.ali0064_NamedInsured & z_S
            sSQL = sSQL & S_z & .AdminLossInfoApd.ali0065_MailAddress1 & z_S
            sSQL = sSQL & S_z & .AdminLossInfoApd.ali0066_MailAddress2 & z_S
            sSQL = sSQL & S_z & .AdminLossInfoApd.ali0067_VehicleDescription & z_S
            sTemp = LTrim(Right(.AdminLossInfoApd.ali0068_CompDed, 10))
            sSQL = sSQL & S_z & Val(goUtil.utCleanValString(sTemp)) & z_S
            sSQL = sSQL & S_z & .AdminLossInfoApd.ali0069_VIN & z_S
            sSQL = sSQL & S_z & .AdminLossInfoApd.cli0063_VehicleLocation & z_S
        End If
        sSQL = sSQL & S_z & .CurrentLossInfo.cli01_CAT & z_S
        sSQL = sSQL & S_z & .CurrentLossInfo.cli02_LossDate & z_S
        sSQL = sSQL & S_z & .CurrentLossInfo.cli03_Adjuster & z_S
        sSQL = sSQL & S_z & .CurrentLossInfo.cli04_DateAsgn & z_S
        sSQL = sSQL & S_z & .CurrentLossInfo.cli05_DateClsd & z_S
        sSQL = sSQL & S_z & .CurrentLossInfo.cli06_SALN & z_S
        sSQL = sSQL & S_z & .CurrentLossInfo.cli07_AdjusterOrigInfo & z_S
        sSQL = sSQL & S_z & .CurrentLossInfo.cli08_DateAsgnOrigInfo & z_S
        sSQL = sSQL & S_z & .CurrentLossInfo.cli09_DateClsdOrigInfo & z_S
        sSQL = sSQL & S_z & .CurrentLossInfo.cli10_PaymentsMadeThisClaim & S_z
    End With
    
    sSQL = Replace(sSQL, Chr(39), Chr(39) & Chr(39))
    sSQL = Replace(sSQL, S_z, Chr(39))
    sSQL = Replace(sSQL, z_S, "', ")
    
    
    
    SQL_WEBCONTROL_20000 = sSQL
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    sErrDesc = mudtXML01LossReport.XML01Loss.CurrentLossInfo.cli06_SALN & vbCrLf & sErrDesc
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLR_ClassName & vbCrLf & "Private Function SQL_WEBCONTROL_20000"
End Function

Private Function SQL_EASYCLAIM_10135() As String
    On Error GoTo EH
    'SQL Variables
    Dim sSQL As String
    Dim sMailAddress As String
    Dim sMailStreet As String
    Dim sMailCity As String
    Dim sMailState As String
    Dim sMailZip As String
    Dim sLossAddress As String
    Dim sLossStreet As String
    Dim sLossCity As String
    Dim sLossState As String
    Dim sLossZip As String
    Dim sDate As String 'Temp date
    Dim sInsuredNames As String
    Dim sIBNum As String
    Dim sTemp As String 'Help build sql
    'Parameter Variables
    Dim Param As V2ECKeyBoard.udtParameter
    Dim sIBPREFIX As String
    Dim sFName As String 'First Name Adjuster
    Dim sLName As String 'Last Name Adjuster
    Dim sCRID As String
    Dim dTaxPercent As Double
    Dim sTypeOfLoss As String
    Dim sSiteAddress As String
    Dim sSiteState As String
    Dim sSiteCity As String
    Dim sSiteZip As String
    Dim lNextID As Long 'Loss Reports Class will use parameter collection
    Dim lErrNum As Long
    Dim sErrDesc As String
   
    If moLRs.ParametersCol Is Nothing Then
        Exit Function
    End If
    'Get Parameters
    sIBPREFIX = moLRs.FindParameter("IB_PREFIX").ParamValue
    sFName = moLRs.FindParameter("ADJ_FIRST_NAME").ParamValue
    sLName = moLRs.FindParameter("ADJ_LAST_NAME").ParamValue
    sCRID = moLRs.FindParameter("CRID").ParamValue
    dTaxPercent = moLRs.FindParameter("TAX_PERCENT").ParamValue
    sTypeOfLoss = moLRs.FindParameter("TYPE_OF_LOSS").ParamValue
    sSiteAddress = moLRs.FindParameter("SITE_ADDRESS").ParamValue
    sSiteState = moLRs.FindParameter("SITE_STATE").ParamValue
    sSiteCity = moLRs.FindParameter("SITE_CITY").ParamValue
    sSiteZip = moLRs.FindParameter("SITE_ZIP").ParamValue
    
    'NextID is stored in loss Reports class so it can
    'continue to add 1 for each loss reports mmmKayyy.
    lNextID = moLRs.NextID
    'Get IBNUM
    sIBNum = sIBPREFIX & moLRs.IBSuffix
    
    'Send Parameters
    'Then Add it to Param collection
    'IBNumber
    Param.ParamName = "IB_NUMBER"
    Param.ParamValue = sIBNum
    moLRs.AddParameter Param
    
    If clsCarLR_LossType = XML01Pro Then
        'Add Policy Desc, Build Lim, Cont Lim
        Param.ParamName = "POLICY_DESC"
        Param.ParamValue = mudtXML01LossReport.XML01Loss.AdminLossInfo.ali0069_PolicyDescription
        moLRs.AddParameter Param
        'Building Lim
        Param.ParamName = "BUILD_LIM"
        Param.ParamValue = mudtXML01LossReport.XML01Loss.AdminLossInfo.ali0070_BldgLimit
        moLRs.AddParameter Param
        'Contents Lim
        Param.ParamName = "CONT_LIM"
        Param.ParamValue = mudtXML01LossReport.XML01Loss.AdminLossInfo.ali0071_ContLimit
        moLRs.AddParameter Param
        'Deductible
        Param.ParamName = "DEDUCTIBLE"
        Param.ParamValue = mudtXML01LossReport.XML01Loss.AdminLossInfo.ali0072_Deductible1
        moLRs.AddParameter Param
    ElseIf clsCarLR_LossType = XML01Apd Then
        'Add Policy Desc, Build Lim, Cont Lim
        Param.ParamName = "POLICY_DESC"
        Param.ParamValue = mudtXML01LossReport.XML01Loss.AdminLossInfoApd.ali0059_PolicyType
        moLRs.AddParameter Param
        'Building Lim
        Param.ParamName = "BUILD_LIM"
        Param.ParamValue = 0
        moLRs.AddParameter Param
        'Contents Lim
        Param.ParamName = "CONT_LIM"
        Param.ParamValue = 0
        moLRs.AddParameter Param
        'Deductible
        Param.ParamName = "DEDUCTIBLE"
        Param.ParamValue = mudtXML01LossReport.XML01Loss.AdminLossInfoApd.ali0068_CompDed
        moLRs.AddParameter Param
    End If
    
    'Build SQL
    sSQL = "INSERT INTO ASSIGNMENTS (ID, "
    sSQL = sSQL & "Client, "
    sSQL = sSQL & "Status, "
    sSQL = sSQL & "AdjustorFirstName, "
    sSQL = sSQL & "AdjustorLastName, "
    sSQL = sSQL & "ClaimRepIDNO, "
    sSQL = sSQL & "Check1FieldHandeld, "
    sSQL = sSQL & "Check1Final, "
    sSQL = sSQL & "Check1ForProperty, "
    sSQL = sSQL & "NCCToIssueExpenseCheck, "
    sSQL = sSQL & "NCCFirstPayment, "
    sSQL = sSQL & "SiteAddress, "
    sSQL = sSQL & "SiteCity, "
    sSQL = sSQL & "SiteState, "
    sSQL = sSQL & "SiteZip, "
    sSQL = sSQL & "PolicyNumber, "
    sSQL = sSQL & "AssignedDate, "
    sSQL = sSQL & "ReceivedDate, "
    sSQL = sSQL & "ClientClaimNo, "
    sSQL = sSQL & "Check1TexasSuffix, "
    sSQL = sSQL & "Check1TexasSubCoverageCode, "
    sSQL = sSQL & "CatCode, "
    sSQL = sSQL & "NewInsuredNames, "
    sSQL = sSQL & "NewMailingAddress, "
    sSQL = sSQL & "MailingAddress, "  'STREET mmKay ;)
    sSQL = sSQL & "MailingZip, "
    sSQL = sSQL & "MailingState, "
    sSQL = sSQL & "MailingCity, "
    sSQL = sSQL & "NewLosslocation, "
    sSQL = sSQL & "LossLocAddress, " 'STREET mmKay ;)
    sSQL = sSQL & "LossLocZip, "
    sSQL = sSQL & "LossLocState, "
    sSQL = sSQL & "LossLocCity, "
    sSQL = sSQL & "MortgageCompany, "
    sSQL = sSQL & "Deductible, "
    sSQL = sSQL & "BuildingPolicyLimit, "
    sSQL = sSQL & "PersonalPropertyLimit, "
    sSQL = sSQL & "LossDate, "
    sSQL = sSQL & "ClaimNo, "
    sSQL = sSQL & "ConfirmMortgageeIsCorrect, "
    sSQL = sSQL & "ConfirmMortgCorrNA, "
    sSQL = sSQL & "ExplainedMortgageeChecks, "
    sSQL = sSQL & "ExplainMortgCheckNA, "
    sSQL = sSQL & "ExplainedAndGaveEstToCustomer, " ' = "Y"
    sSQL = sSQL & "ExplandGaveRCVFormToCust, " '= "Y"
    sSQL = sSQL & "ConfirmedCoverage, " ' = "Y"
    sSQL = sSQL & "PriorLoss, " ' = False
    sSQL = sSQL & "PriorLossNA, " ' = False
    sSQL = sSQL & "TaxPercent, "
    sSQL = sSQL & "TypeOfLoss ) "
    
    With mudtXML01LossReport.XML01Loss
        If clsCarLR_LossType = XML01Pro Then
            sSQL = sSQL & "VALUES (" & lNextID & ", "       'ID, "
            sSQL = sSQL & S_z & "Farmers" & z_S             'Client, " = Faramers
            sSQL = sSQL & S_z & "NEW" & z_S                 '"Status, " = New
            sSQL = sSQL & S_z & sFName & z_S                ' "AdjustorLastName, " = param
            sSQL = sSQL & S_z & sLName & z_S                ' "AdjustorFirstName, " = param
            sSQL = sSQL & S_z & sCRID & z_S                 ' "ClaimRepIDNO, " =param
            sSQL = sSQL & True & ", "                       ' "Check1FieldHandeld, " =True
            sSQL = sSQL & True & ", "                       ' "Check1Final, " =True
            sSQL = sSQL & True & ", "                       ' "Check1ForProperty, " =True
            sSQL = sSQL & True & ", "                       ' "NCCToIssueExpenseCheck, " = True
            sSQL = sSQL & True & ", "                       ' "NCCFirstPayment, " =True
            sSQL = sSQL & S_z & sSiteAddress & z_S          ' "SiteAddress, " =Param
            sSQL = sSQL & S_z & sSiteCity & z_S             ' "SiteCity, " = Param
            sSQL = sSQL & S_z & sSiteState & z_S            ' "SiteState, " =Param
            sSQL = sSQL & S_z & sSiteZip & z_S              ' "SiteZip, " =Param
            sSQL = sSQL & S_z & .AdminLossInfo.ali0054_PolicyNum & z_S  ' "PolicyNumber, "
            sDate = .CurrentLossInfo.cli04_DateAsgn
            If Not IsDate(sDate) Then
                sDate = Now()
            End If
            sSQL = sSQL & S_z & Format(CDate(sDate), "YYYYMMDD") & z_S      ' "AssignedDate, "
            sDate = Format(Now(), "YYYYMMDD")
            sSQL = sSQL & S_z & sDate & z_S                                 ' "ReceivedDate, "
            sSQL = sSQL & S_z & .CurrentLossInfo.cli06_SALN & z_S           ' "ClientClaimNo, "
            sSQL = sSQL & S_z & vbNullString & z_S                          ' "Check1TexasSuffix, "
            sSQL = sSQL & S_z & vbNullString & z_S                          ' "Check1TexasSubCoverageCode, "
            sSQL = sSQL & S_z & .CurrentLossInfo.cli01_CAT & z_S            ' "CatCode, "
            
            sInsuredNames = .AdminLossInfo.ali0065_MainFileInsuredName
            If sInsuredNames = vbNullString Then
                sInsuredNames = .AdminLossInfo.ali0062_NamedInsured
            End If
            sSQL = sSQL & S_z & sInsuredNames & z_S                         ' "NewInsuredNames, "
            
            'Mailing Address
            sMailAddress = .AdminLossInfo.ali0063_MailAddress1 & String(2, Chr(32)) & IIf(.AdminLossInfo.ali0064_MailAddress2 = vbNullString, vbNullString, S_z & .AdminLossInfo.ali0064_MailAddress2)
            goUtil.utFillAddressFields sMailAddress, sMailZip, sMailState, sMailCity, sMailStreet
            goUtil.utUpdateAddress sMailAddress, sMailZip, sMailState, sMailCity, sMailStreet
            
            sSQL = sSQL & S_z & sMailAddress & z_S                          ' "NewMailingAddress, "
            sSQL = sSQL & S_z & sMailStreet & z_S                           ' "MailingAddress, "
            sSQL = sSQL & S_z & sMailZip & z_S                              ' "MailingZip, "
            sSQL = sSQL & S_z & sMailState & z_S                            ' "MailingState, "
            sSQL = sSQL & S_z & sMailCity & z_S                             ' "MailingCity, "
            
            'Loss Location Address
            sLossAddress = .AdminLossInfo.ali0080_LossLocAddress1 & String(2, Chr(32)) & IIf(.AdminLossInfo.ali0081_LossLocAddress2 = vbNullString, vbNullString, S_z & .AdminLossInfo.ali0081_LossLocAddress2)
            goUtil.utFillAddressFields sLossAddress, sLossZip, sLossState, sLossCity, sLossStreet
            'Check State against Mailing State
            If InStr(1, sLossAddress, sMailZip, vbTextCompare) > 0 Then
                sLossAddress = Replace(sLossAddress, sMailZip, vbNullString, , , vbTextCompare)
                sLossZip = sMailZip
            End If
            If InStr(1, sLossAddress, sMailState, vbTextCompare) > 0 Then
                sLossAddress = Replace(sLossAddress, " " & sMailState, vbNullString, , , vbTextCompare)
                sLossState = sMailState
            End If
            '-------------------------Debug ONLY------------------------------
    '        Dim lcount As Long
    '        For lcount = 1 To Len(sMailCity)
    '            Debug.Print Mid(sMailCity, lcount, 1) & "------->" & Asc(Mid(sMailCity, lcount, 1))
    '        Next
    '        For lcount = 1 To Len(sLossAddress)
    '            Debug.Print Mid(sLossAddress, lcount, 1) & "------->" & Asc(Mid(sLossAddress, lcount, 1))
    '        Next
            '-------------------------Debug Only-----------------------------
            If InStr(1, sLossAddress, Replace(sMailCity, Chr(160), Chr(32)), vbTextCompare) > 0 Then
                sLossAddress = Replace(sLossAddress, Replace(sMailCity, Chr(160), Chr(32)), vbNullString, , , vbTextCompare)
                sLossCity = sMailCity
            End If
            If InStr(1, sLossStreet, Replace(sMailCity, Chr(160), Chr(32)), vbTextCompare) > 0 Then
                sLossStreet = RTrim(Replace(sLossStreet, Replace(sMailCity, Chr(160), Chr(32)), vbNullString, , , vbTextCompare))
            End If
            
            goUtil.utUpdateAddress sLossAddress, sLossZip, sLossState, sLossCity, sLossStreet
            
            sSQL = sSQL & S_z & sLossAddress & z_S                          ' "NewLossLocation, "
            sSQL = sSQL & S_z & sLossStreet & z_S                           ' "LossAddress, "
            sSQL = sSQL & S_z & sLossZip & z_S                              ' "LossZip, "
            sSQL = sSQL & S_z & sLossState & z_S                            ' "LossState, "
            sSQL = sSQL & S_z & sLossCity & z_S                             ' "LossCity
    
            sSQL = sSQL & S_z & .AdminLossInfo.ali0066_MortgageHolder & z_S  ' "MortgageCompany, "
            'Val function Bug in VB6
            'http://msdn.microsoft.com/vbasic/productinfo/previous/vb6/tips/01pasttips.asp
            'Need to parse out both % and !  because these trailing equate to Double and Single
            'and Val bugs because it can't convert Double or single into integer
            sTemp = LTrim(Right(Replace(.AdminLossInfo.ali0072_Deductible1, ",", vbNullString), 10))
            sSQL = sSQL & Val(goUtil.utCleanValString(sTemp)) & ", " ' "Deductible, " 'CUR
            sTemp = Replace(.AdminLossInfo.ali0070_BldgLimit, ",", vbNullString)
            sSQL = sSQL & Val(goUtil.utCleanValString(sTemp)) & ", " ' "BuildingPolicyLimit, " 'CUR
            sTemp = Replace(.AdminLossInfo.ali0071_ContLimit, ",", vbNullString)
            sSQL = sSQL & Val(goUtil.utCleanValString(sTemp)) & ", " ' "PersonalPropertyLimit, " 'CUR
            sDate = .CurrentLossInfo.cli02_LossDate
            If Not IsDate(sDate) Then
                sDate = Now()
            End If
            sSQL = sSQL & S_z & Format(CDate(sDate), "YYYYMMDD") & z_S      'Loss Date
            sSQL = sSQL & S_z & sIBNum & z_S                                ' "ClaimNo, "
            If Trim(.AdminLossInfo.ali0066_MortgageHolder) = vbNullString Then
                sSQL = sSQL & S_z & vbNullString & z_S                      ' "ConfirmMortgageeIsCorrect, "
                sSQL = sSQL & True & ", "                                   ' "ConfirmMortgCorrNA, "
                sSQL = sSQL & S_z & vbNullString & z_S                      ' "ExplainedMortgageeChecks, "
                sSQL = sSQL & True & ", "                                   ' "ExplainMortgCheckNA, "
            Else
                sSQL = sSQL & S_z & "Y" & z_S                               ' "ConfirmMortgageeIsCorrect, "
                sSQL = sSQL & False & ", "                                  ' "ConfirmMortgCorrNA, "
                sSQL = sSQL & S_z & "Y" & z_S                               ' "ExplainedMortgageeChecks, "
                sSQL = sSQL & False & ", "
            End If
            
            sSQL = sSQL & S_z & "Y" & z_S                                   ' "ExplainedAndGaveEstToCustomer, " ' = "Y"
            sSQL = sSQL & S_z & "Y" & z_S                                   ' "ExplandGaveRCVFormToCust, " '= "Y"
            sSQL = sSQL & S_z & "Y" & z_S                                   ' "ConfirmedCoverage, " ' = "Y"
            sSQL = sSQL & False & ", "                                      ' "PriorLoss, " ' = False
            sSQL = sSQL & False & ", "                                      ' "PriorLossNA, " ' = False
            sSQL = sSQL & dTaxPercent & ", "                                ' "TaxPercent, "
            sSQL = sSQL & S_z & sTypeOfLoss & S_z & " ) "                   ' "TypeOfLoss, "
        'AUTO APD
        ElseIf clsCarLR_LossType = XML01Apd Then
            sSQL = sSQL & "VALUES (" & lNextID & ", "       'ID, "
            sSQL = sSQL & S_z & "Farmers" & z_S             'Client, " = Faramers
            sSQL = sSQL & S_z & "NEW" & z_S                 '"Status, " = New
            sSQL = sSQL & S_z & sFName & z_S                ' "AdjustorLastName, " = param
            sSQL = sSQL & S_z & sLName & z_S                ' "AdjustorFirstName, " = param
            sSQL = sSQL & S_z & sCRID & z_S                 ' "ClaimRepIDNO, " =param
            sSQL = sSQL & True & ", "                       ' "Check1FieldHandeld, " =True
            sSQL = sSQL & True & ", "                       ' "Check1Final, " =True
            sSQL = sSQL & True & ", "                       ' "Check1ForProperty, " =True
            sSQL = sSQL & True & ", "                       ' "NCCToIssueExpenseCheck, " = True
            sSQL = sSQL & True & ", "                       ' "NCCFirstPayment, " =True
            sSQL = sSQL & S_z & sSiteAddress & z_S          ' "SiteAddress, " =Param
            sSQL = sSQL & S_z & sSiteCity & z_S             ' "SiteCity, " = Param
            sSQL = sSQL & S_z & sSiteState & z_S            ' "SiteState, " =Param
            sSQL = sSQL & S_z & sSiteZip & z_S              ' "SiteZip, " =Param
            sSQL = sSQL & S_z & .AdminLossInfoApd.ali0051_PolicyNum & z_S  ' "PolicyNumber, "
            sDate = .CurrentLossInfo.cli04_DateAsgn
            If Not IsDate(sDate) Then
                sDate = Now()
            End If
            sSQL = sSQL & S_z & Format(CDate(sDate), "YYYYMMDD") & z_S      ' "AssignedDate, "
            sDate = Format(Now(), "YYYYMMDD")
            sSQL = sSQL & S_z & sDate & z_S                                 ' "ReceivedDate, "
            sSQL = sSQL & S_z & .CurrentLossInfo.cli06_SALN & z_S           ' "ClientClaimNo, "
            sSQL = sSQL & S_z & vbNullString & z_S                          ' "Check1TexasSuffix, "
            sSQL = sSQL & S_z & vbNullString & z_S                          ' "Check1TexasSubCoverageCode, "
            sSQL = sSQL & S_z & .CurrentLossInfo.cli01_CAT & z_S            ' "CatCode, "
            
            sInsuredNames = .AdminLossInfoApd.ali0064_NamedInsured
            sSQL = sSQL & S_z & sInsuredNames & z_S                         ' "NewInsuredNames, "
            
            'Mailing Address
            sMailAddress = .AdminLossInfoApd.ali0065_MailAddress1 & String(2, Chr(32)) & IIf(.AdminLossInfoApd.ali0066_MailAddress2 = vbNullString, vbNullString, S_z & .AdminLossInfoApd.ali0066_MailAddress2)
            goUtil.utFillAddressFields sMailAddress, sMailZip, sMailState, sMailCity, sMailStreet
            goUtil.utUpdateAddress sMailAddress, sMailZip, sMailState, sMailCity, sMailStreet
            
            sSQL = sSQL & S_z & sMailAddress & z_S                          ' "NewMailingAddress, "
            sSQL = sSQL & S_z & sMailStreet & z_S                           ' "MailingAddress, "
            sSQL = sSQL & S_z & sMailZip & z_S                              ' "MailingZip, "
            sSQL = sSQL & S_z & sMailState & z_S                            ' "MailingState, "
            sSQL = sSQL & S_z & sMailCity & z_S                             ' "MailingCity, "
            
            'Loss Location Address
            sLossAddress = .AdminLossInfoApd.cli0063_VehicleLocation
            goUtil.utFillAddressFields sLossAddress, sLossZip, sLossState, sLossCity, sLossStreet
            'Check State against Mailing State
            If InStr(1, sLossAddress, sMailZip, vbTextCompare) > 0 Then
                sLossAddress = Replace(sLossAddress, sMailZip, vbNullString, , , vbTextCompare)
                sLossZip = sMailZip
            End If
            If InStr(1, sLossAddress, sMailState, vbTextCompare) > 0 Then
                sLossAddress = Replace(sLossAddress, " " & sMailState, vbNullString, , , vbTextCompare)
                sLossState = sMailState
            End If
            '-------------------------Debug ONLY------------------------------
    '        Dim lcount As Long
    '        For lcount = 1 To Len(sMailCity)
    '            Debug.Print Mid(sMailCity, lcount, 1) & "------->" & Asc(Mid(sMailCity, lcount, 1))
    '        Next
    '        For lcount = 1 To Len(sLossAddress)
    '            Debug.Print Mid(sLossAddress, lcount, 1) & "------->" & Asc(Mid(sLossAddress, lcount, 1))
    '        Next
            '-------------------------Debug Only-----------------------------
            If InStr(1, sLossAddress, Replace(sMailCity, Chr(160), Chr(32)), vbTextCompare) > 0 Then
                sLossAddress = Replace(sLossAddress, Replace(sMailCity, Chr(160), Chr(32)), vbNullString, , , vbTextCompare)
                sLossCity = sMailCity
            End If
            If InStr(1, sLossStreet, Replace(sMailCity, Chr(160), Chr(32)), vbTextCompare) > 0 Then
                sLossStreet = RTrim(Replace(sLossStreet, Replace(sMailCity, Chr(160), Chr(32)), vbNullString, , , vbTextCompare))
            End If
            
            goUtil.utUpdateAddress sLossAddress, sLossZip, sLossState, sLossCity, sLossStreet
            
            sSQL = sSQL & S_z & sLossAddress & z_S                          ' "NewLossLocation, "
            sSQL = sSQL & S_z & sLossStreet & z_S                           ' "LossAddress, "
            sSQL = sSQL & S_z & sLossZip & z_S                              ' "LossZip, "
            sSQL = sSQL & S_z & sLossState & z_S                            ' "LossState, "
            sSQL = sSQL & S_z & sLossCity & z_S                             ' "LossCity
    
            sSQL = sSQL & S_z & .AdminLossInfoApd.ali0057_MortgageHolder & z_S  ' "MortgageCompany, "
            'Val function Bug in VB6
            'http://msdn.microsoft.com/vbasic/productinfo/previous/vb6/tips/01pasttips.asp
            'Need to parse out both % and !  because these trailing equate to Double and Single
            'and Val bugs because it can't convert Double or single into integer
            sTemp = LTrim(Right(Replace(.AdminLossInfoApd.ali0068_CompDed, ",", vbNullString), 10))
            sSQL = sSQL & Val(goUtil.utCleanValString(sTemp)) & ", " ' "Deductible, " 'CUR
            sTemp = 0
            sSQL = sSQL & Val(goUtil.utCleanValString(sTemp)) & ", " ' "BuildingPolicyLimit, " 'CUR
            sTemp = 0
            sSQL = sSQL & Val(goUtil.utCleanValString(sTemp)) & ", " ' "PersonalPropertyLimit, " 'CUR
            sDate = .CurrentLossInfo.cli02_LossDate
            If Not IsDate(sDate) Then
                sDate = Now()
            End If
            sSQL = sSQL & S_z & Format(CDate(sDate), "YYYYMMDD") & z_S      'Loss Date
            sSQL = sSQL & S_z & sIBNum & z_S                                ' "ClaimNo, "
            If Trim(.AdminLossInfoApd.ali0057_MortgageHolder) = vbNullString Then
                sSQL = sSQL & S_z & vbNullString & z_S                      ' "ConfirmMortgageeIsCorrect, "
                sSQL = sSQL & True & ", "                                   ' "ConfirmMortgCorrNA, "
                sSQL = sSQL & S_z & vbNullString & z_S                      ' "ExplainedMortgageeChecks, "
                sSQL = sSQL & True & ", "                                   ' "ExplainMortgCheckNA, "
            Else
                sSQL = sSQL & S_z & "Y" & z_S                               ' "ConfirmMortgageeIsCorrect, "
                sSQL = sSQL & False & ", "                                  ' "ConfirmMortgCorrNA, "
                sSQL = sSQL & S_z & "Y" & z_S                               ' "ExplainedMortgageeChecks, "
                sSQL = sSQL & False & ", "
            End If
            
            sSQL = sSQL & S_z & "Y" & z_S                                   ' "ExplainedAndGaveEstToCustomer, " ' = "Y"
            sSQL = sSQL & S_z & "Y" & z_S                                   ' "ExplandGaveRCVFormToCust, " '= "Y"
            sSQL = sSQL & S_z & "Y" & z_S                                   ' "ConfirmedCoverage, " ' = "Y"
            sSQL = sSQL & False & ", "                                      ' "PriorLoss, " ' = False
            sSQL = sSQL & False & ", "                                      ' "PriorLossNA, " ' = False
            sSQL = sSQL & dTaxPercent & ", "                                ' "TaxPercent, "
            sSQL = sSQL & S_z & sTypeOfLoss & S_z & " ) "                   ' "TypeOfLoss, "
        End If
    End With
    
    SQL_EASYCLAIM_10135 = sSQL
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    sErrDesc = mudtXML01LossReport.XML01Loss.CurrentLossInfo.cli06_SALN & vbCrLf & sErrDesc
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLR_ClassName & vbCrLf & "Private Function SQL_EASYCLAIM_10135"
End Function

Private Function SQL_EASYCLAIM_20000() As String
    On Error GoTo EH
    'SQL Variables
    Dim sSQL As String
    Dim sMailAddress As String
    Dim sMailStreet As String
    Dim sMailCity As String
    Dim sMailState As String
    Dim sMailZip As String
    Dim sLossAddress As String
    Dim sLossStreet As String
    Dim sLossCity As String
    Dim sLossState As String
    Dim sLossZip As String
    Dim sDate As String 'Temp date
    Dim sInsuredNames As String
    Dim sIBNum As String
    Dim sTemp As String 'Help build sql
    'Parameter Variables
    Dim Param As V2ECKeyBoard.udtParameter
    Dim sIBPREFIX As String
    Dim sFName As String 'First Name Adjuster
    Dim sLName As String 'Last Name Adjuster
    Dim sCRID As String
    Dim dTaxPercent As Double
    Dim sTypeOfLoss As String
    Dim sSiteAddress As String
    Dim sSiteState As String
    Dim sSiteCity As String
    Dim sSiteZip As String
    Dim lNextID As Long 'Loss Reports Class will use parameter collection
    Dim lErrNum As Long
    Dim sErrDesc As String
   
    If moLRs.ParametersCol Is Nothing Then
        Exit Function
    End If
    'Get Parameters
    sIBPREFIX = moLRs.FindParameter("IB_PREFIX").ParamValue
    sFName = moLRs.FindParameter("ADJ_FIRST_NAME").ParamValue
    sLName = moLRs.FindParameter("ADJ_LAST_NAME").ParamValue
    sCRID = moLRs.FindParameter("CRID").ParamValue
    dTaxPercent = moLRs.FindParameter("TAX_PERCENT").ParamValue
    sTypeOfLoss = moLRs.FindParameter("TYPE_OF_LOSS").ParamValue
    sSiteAddress = moLRs.FindParameter("SITE_ADDRESS").ParamValue
    sSiteState = moLRs.FindParameter("SITE_STATE").ParamValue
    sSiteCity = moLRs.FindParameter("SITE_CITY").ParamValue
    sSiteZip = moLRs.FindParameter("SITE_ZIP").ParamValue
    
    'NextID is stored in loss Reports class so it can
    'continue to add 1 for each loss reports mmmKayyy.
    lNextID = moLRs.NextID
    'Get IBNUM
    sIBNum = sIBPREFIX & moLRs.IBSuffix
    
    'Send Parameters
    'Then Add it to Param collection
    'IBNumber
    Param.ParamName = "IB_NUMBER"
    Param.ParamValue = sIBNum
    moLRs.AddParameter Param
    
    If clsCarLR_LossType = XML01Pro Then
        'Add Policy Desc, Build Lim, Cont Lim
        Param.ParamName = "POLICY_DESC"
        Param.ParamValue = mudtXML01LossReport.XML01Loss.AdminLossInfo.ali0069_PolicyDescription
        moLRs.AddParameter Param
        'Building Lim
        Param.ParamName = "BUILD_LIM"
        Param.ParamValue = mudtXML01LossReport.XML01Loss.AdminLossInfo.ali0070_BldgLimit
        moLRs.AddParameter Param
        'Contents Lim
        Param.ParamName = "CONT_LIM"
        Param.ParamValue = mudtXML01LossReport.XML01Loss.AdminLossInfo.ali0071_ContLimit
        moLRs.AddParameter Param
        'Deductible
        Param.ParamName = "DEDUCTIBLE"
        Param.ParamValue = mudtXML01LossReport.XML01Loss.AdminLossInfo.ali0072_Deductible1
        moLRs.AddParameter Param
    ElseIf clsCarLR_LossType = XML01Apd Then
        'Add Policy Desc, Build Lim, Cont Lim
        Param.ParamName = "POLICY_DESC"
        Param.ParamValue = mudtXML01LossReport.XML01Loss.AdminLossInfoApd.ali0059_PolicyType
        moLRs.AddParameter Param
        'Building Lim
        Param.ParamName = "BUILD_LIM"
        Param.ParamValue = 0
        moLRs.AddParameter Param
        'Contents Lim
        Param.ParamName = "CONT_LIM"
        Param.ParamValue = 0
        moLRs.AddParameter Param
        'Deductible
        Param.ParamName = "DEDUCTIBLE"
        Param.ParamValue = mudtXML01LossReport.XML01Loss.AdminLossInfoApd.ali0068_CompDed
        moLRs.AddParameter Param
    End If
    
    'Build SQL
    sSQL = "INSERT INTO ASSIGNMENTS (ID, "
    sSQL = sSQL & "Client, "
    sSQL = sSQL & "Status, "
    sSQL = sSQL & "AdjustorFirstName, "
    sSQL = sSQL & "AdjustorLastName, "
    sSQL = sSQL & "ClaimRepIDNO, "
    sSQL = sSQL & "Check1FieldHandeld, "
    sSQL = sSQL & "Check1Final, "
    sSQL = sSQL & "Check1ForProperty, "
    sSQL = sSQL & "NCCToIssueExpenseCheck, "
    sSQL = sSQL & "NCCFirstPayment, "
    sSQL = sSQL & "SiteAddress, "
    sSQL = sSQL & "SiteCity, "
    sSQL = sSQL & "SiteState, "
    sSQL = sSQL & "SiteZip, "
    sSQL = sSQL & "PolicyNumber, "
    sSQL = sSQL & "AssignedDate, "
    sSQL = sSQL & "ReceivedDate, "
    sSQL = sSQL & "ClientClaimNo, "
    sSQL = sSQL & "Check1TexasSuffix, "
    sSQL = sSQL & "Check1TexasSubCoverageCode, "
    sSQL = sSQL & "CatCode, "
    sSQL = sSQL & "NewInsuredNames, "
    sSQL = sSQL & "NewMailingAddress, "
    sSQL = sSQL & "MailingAddress, "  'STREET mmKay ;)
    sSQL = sSQL & "MailingZip, "
    sSQL = sSQL & "MailingState, "
    sSQL = sSQL & "MailingCity, "
    sSQL = sSQL & "NewLosslocation, "
    sSQL = sSQL & "LossLocAddress, " 'STREET mmKay ;)
    sSQL = sSQL & "LossLocZip, "
    sSQL = sSQL & "LossLocState, "
    sSQL = sSQL & "LossLocCity, "
    sSQL = sSQL & "MortgageCompany, "
    sSQL = sSQL & "Deductible, "
    sSQL = sSQL & "BuildingPolicyLimit, "
    sSQL = sSQL & "PersonalPropertyLimit, "
    sSQL = sSQL & "LossDate, "
    sSQL = sSQL & "ClaimNo, "
    sSQL = sSQL & "ConfirmMortgageeIsCorrect, "
    sSQL = sSQL & "ConfirmMortgCorrNA, "
    sSQL = sSQL & "ExplainedMortgageeChecks, "
    sSQL = sSQL & "ExplainMortgCheckNA, "
    sSQL = sSQL & "ExplainedAndGaveEstToCustomer, " ' = "Y"
    sSQL = sSQL & "ExplandGaveRCVFormToCust, " '= "Y"
    sSQL = sSQL & "ConfirmedCoverage, " ' = "Y"
    sSQL = sSQL & "PriorLoss, " ' = False
    sSQL = sSQL & "PriorLossNA, " ' = False
    sSQL = sSQL & "TaxPercent, "
    sSQL = sSQL & "TypeOfLoss ) "
    
    With mudtXML01LossReport.XML01Loss
        If clsCarLR_LossType = XML01Pro Then
            sSQL = sSQL & "VALUES (" & lNextID & ", "       'ID, "
            sSQL = sSQL & S_z & "Farmers" & z_S             'Client, " = Faramers
            sSQL = sSQL & S_z & "NEW" & z_S                 '"Status, " = New
            sSQL = sSQL & S_z & sFName & z_S                ' "AdjustorLastName, " = param
            sSQL = sSQL & S_z & sLName & z_S                ' "AdjustorFirstName, " = param
            sSQL = sSQL & S_z & sCRID & z_S                 ' "ClaimRepIDNO, " =param
            sSQL = sSQL & True & ", "                       ' "Check1FieldHandeld, " =True
            sSQL = sSQL & True & ", "                       ' "Check1Final, " =True
            sSQL = sSQL & True & ", "                       ' "Check1ForProperty, " =True
            sSQL = sSQL & True & ", "                       ' "NCCToIssueExpenseCheck, " = True
            sSQL = sSQL & True & ", "                       ' "NCCFirstPayment, " =True
            sSQL = sSQL & S_z & sSiteAddress & z_S          ' "SiteAddress, " =Param
            sSQL = sSQL & S_z & sSiteCity & z_S             ' "SiteCity, " = Param
            sSQL = sSQL & S_z & sSiteState & z_S            ' "SiteState, " =Param
            sSQL = sSQL & S_z & sSiteZip & z_S              ' "SiteZip, " =Param
            sSQL = sSQL & S_z & .AdminLossInfo.ali0054_PolicyNum & z_S  ' "PolicyNumber, "
            sDate = .CurrentLossInfo.cli04_DateAsgn
            If Not IsDate(sDate) Then
                sDate = Now()
            End If
            sSQL = sSQL & S_z & Format(CDate(sDate), "YYYYMMDD") & z_S      ' "AssignedDate, "
            sDate = Format(Now(), "YYYYMMDD")
            sSQL = sSQL & S_z & sDate & z_S                                 ' "ReceivedDate, "
            sSQL = sSQL & S_z & .CurrentLossInfo.cli06_SALN & z_S           ' "ClientClaimNo, "
            sSQL = sSQL & S_z & vbNullString & z_S                          ' "Check1TexasSuffix, "
            sSQL = sSQL & S_z & vbNullString & z_S                          ' "Check1TexasSubCoverageCode, "
            sSQL = sSQL & S_z & .CurrentLossInfo.cli01_CAT & z_S            ' "CatCode, "
            
            sInsuredNames = .AdminLossInfo.ali0065_MainFileInsuredName
            If sInsuredNames = vbNullString Then
                sInsuredNames = .AdminLossInfo.ali0062_NamedInsured
            End If
            sSQL = sSQL & S_z & sInsuredNames & z_S                         ' "NewInsuredNames, "
            
            'Mailing Address
            sMailAddress = .AdminLossInfo.ali0063_MailAddress1 & String(2, Chr(32)) & IIf(.AdminLossInfo.ali0064_MailAddress2 = vbNullString, vbNullString, S_z & .AdminLossInfo.ali0064_MailAddress2)
            goUtil.utFillAddressFields sMailAddress, sMailZip, sMailState, sMailCity, sMailStreet
            goUtil.utUpdateAddress sMailAddress, sMailZip, sMailState, sMailCity, sMailStreet
            
            sSQL = sSQL & S_z & sMailAddress & z_S                          ' "NewMailingAddress, "
            sSQL = sSQL & S_z & sMailStreet & z_S                           ' "MailingAddress, "
            sSQL = sSQL & S_z & sMailZip & z_S                              ' "MailingZip, "
            sSQL = sSQL & S_z & sMailState & z_S                            ' "MailingState, "
            sSQL = sSQL & S_z & sMailCity & z_S                             ' "MailingCity, "
            
            'Loss Location Address
            sLossAddress = .AdminLossInfo.ali0080_LossLocAddress1 & String(2, Chr(32)) & IIf(.AdminLossInfo.ali0081_LossLocAddress2 = vbNullString, vbNullString, S_z & .AdminLossInfo.ali0081_LossLocAddress2)
            goUtil.utFillAddressFields sLossAddress, sLossZip, sLossState, sLossCity, sLossStreet
            'Check State against Mailing State
            If InStr(1, sLossAddress, sMailZip, vbTextCompare) > 0 Then
                sLossAddress = Replace(sLossAddress, sMailZip, vbNullString, , , vbTextCompare)
                sLossZip = sMailZip
            End If
            If InStr(1, sLossAddress, sMailState, vbTextCompare) > 0 Then
                sLossAddress = Replace(sLossAddress, " " & sMailState, vbNullString, , , vbTextCompare)
                sLossState = sMailState
            End If
            '-------------------------Debug ONLY------------------------------
    '        Dim lcount As Long
    '        For lcount = 1 To Len(sMailCity)
    '            Debug.Print Mid(sMailCity, lcount, 1) & "------->" & Asc(Mid(sMailCity, lcount, 1))
    '        Next
    '        For lcount = 1 To Len(sLossAddress)
    '            Debug.Print Mid(sLossAddress, lcount, 1) & "------->" & Asc(Mid(sLossAddress, lcount, 1))
    '        Next
            '-------------------------Debug Only-----------------------------
            If InStr(1, sLossAddress, Replace(sMailCity, Chr(160), Chr(32)), vbTextCompare) > 0 Then
                sLossAddress = Replace(sLossAddress, Replace(sMailCity, Chr(160), Chr(32)), vbNullString, , , vbTextCompare)
                sLossCity = sMailCity
            End If
            If InStr(1, sLossStreet, Replace(sMailCity, Chr(160), Chr(32)), vbTextCompare) > 0 Then
                sLossStreet = RTrim(Replace(sLossStreet, Replace(sMailCity, Chr(160), Chr(32)), vbNullString, , , vbTextCompare))
            End If
            
            goUtil.utUpdateAddress sLossAddress, sLossZip, sLossState, sLossCity, sLossStreet
            
            sSQL = sSQL & S_z & sLossAddress & z_S                          ' "NewLossLocation, "
            sSQL = sSQL & S_z & sLossStreet & z_S                           ' "LossAddress, "
            sSQL = sSQL & S_z & sLossZip & z_S                              ' "LossZip, "
            sSQL = sSQL & S_z & sLossState & z_S                            ' "LossState, "
            sSQL = sSQL & S_z & sLossCity & z_S                             ' "LossCity
    
            sSQL = sSQL & S_z & .AdminLossInfo.ali0066_MortgageHolder & z_S  ' "MortgageCompany, "
            'Val function Bug in VB6
            'http://msdn.microsoft.com/vbasic/productinfo/previous/vb6/tips/01pasttips.asp
            'Need to parse out both % and !  because these trailing equate to Double and Single
            'and Val bugs because it can't convert Double or single into integer
            sTemp = LTrim(Right(Replace(.AdminLossInfo.ali0072_Deductible1, ",", vbNullString), 10))
            sSQL = sSQL & Val(goUtil.utCleanValString(sTemp)) & ", " ' "Deductible, " 'CUR
            sTemp = Replace(.AdminLossInfo.ali0070_BldgLimit, ",", vbNullString)
            sSQL = sSQL & Val(goUtil.utCleanValString(sTemp)) & ", " ' "BuildingPolicyLimit, " 'CUR
            sTemp = Replace(.AdminLossInfo.ali0071_ContLimit, ",", vbNullString)
            sSQL = sSQL & Val(goUtil.utCleanValString(sTemp)) & ", " ' "PersonalPropertyLimit, " 'CUR
            sDate = .CurrentLossInfo.cli02_LossDate
            If Not IsDate(sDate) Then
                sDate = Now()
            End If
            sSQL = sSQL & S_z & Format(CDate(sDate), "YYYYMMDD") & z_S      'Loss Date
            sSQL = sSQL & S_z & sIBNum & z_S                                ' "ClaimNo, "
            If Trim(.AdminLossInfo.ali0066_MortgageHolder) = vbNullString Then
                sSQL = sSQL & S_z & vbNullString & z_S                      ' "ConfirmMortgageeIsCorrect, "
                sSQL = sSQL & True & ", "                                   ' "ConfirmMortgCorrNA, "
                sSQL = sSQL & S_z & vbNullString & z_S                      ' "ExplainedMortgageeChecks, "
                sSQL = sSQL & True & ", "                                   ' "ExplainMortgCheckNA, "
            Else
                sSQL = sSQL & S_z & "Y" & z_S                               ' "ConfirmMortgageeIsCorrect, "
                sSQL = sSQL & False & ", "                                  ' "ConfirmMortgCorrNA, "
                sSQL = sSQL & S_z & "Y" & z_S                               ' "ExplainedMortgageeChecks, "
                sSQL = sSQL & False & ", "
            End If
            
            sSQL = sSQL & S_z & "Y" & z_S                                   ' "ExplainedAndGaveEstToCustomer, " ' = "Y"
            sSQL = sSQL & S_z & "Y" & z_S                                   ' "ExplandGaveRCVFormToCust, " '= "Y"
            sSQL = sSQL & S_z & "Y" & z_S                                   ' "ConfirmedCoverage, " ' = "Y"
            sSQL = sSQL & False & ", "                                      ' "PriorLoss, " ' = False
            sSQL = sSQL & False & ", "                                      ' "PriorLossNA, " ' = False
            sSQL = sSQL & dTaxPercent & ", "                                ' "TaxPercent, "
            sSQL = sSQL & S_z & sTypeOfLoss & S_z & " ) "                   ' "TypeOfLoss, "
        'AUTO APD
        ElseIf clsCarLR_LossType = XML01Apd Then
            sSQL = sSQL & "VALUES (" & lNextID & ", "       'ID, "
            sSQL = sSQL & S_z & "Farmers" & z_S             'Client, " = Faramers
            sSQL = sSQL & S_z & "NEW" & z_S                 '"Status, " = New
            sSQL = sSQL & S_z & sFName & z_S                ' "AdjustorLastName, " = param
            sSQL = sSQL & S_z & sLName & z_S                ' "AdjustorFirstName, " = param
            sSQL = sSQL & S_z & sCRID & z_S                 ' "ClaimRepIDNO, " =param
            sSQL = sSQL & True & ", "                       ' "Check1FieldHandeld, " =True
            sSQL = sSQL & True & ", "                       ' "Check1Final, " =True
            sSQL = sSQL & True & ", "                       ' "Check1ForProperty, " =True
            sSQL = sSQL & True & ", "                       ' "NCCToIssueExpenseCheck, " = True
            sSQL = sSQL & True & ", "                       ' "NCCFirstPayment, " =True
            sSQL = sSQL & S_z & sSiteAddress & z_S          ' "SiteAddress, " =Param
            sSQL = sSQL & S_z & sSiteCity & z_S             ' "SiteCity, " = Param
            sSQL = sSQL & S_z & sSiteState & z_S            ' "SiteState, " =Param
            sSQL = sSQL & S_z & sSiteZip & z_S              ' "SiteZip, " =Param
            sSQL = sSQL & S_z & .AdminLossInfoApd.ali0051_PolicyNum & z_S  ' "PolicyNumber, "
            sDate = .CurrentLossInfo.cli04_DateAsgn
            If Not IsDate(sDate) Then
                sDate = Now()
            End If
            sSQL = sSQL & S_z & Format(CDate(sDate), "YYYYMMDD") & z_S      ' "AssignedDate, "
            sDate = Format(Now(), "YYYYMMDD")
            sSQL = sSQL & S_z & sDate & z_S                                 ' "ReceivedDate, "
            sSQL = sSQL & S_z & .CurrentLossInfo.cli06_SALN & z_S           ' "ClientClaimNo, "
            sSQL = sSQL & S_z & vbNullString & z_S                          ' "Check1TexasSuffix, "
            sSQL = sSQL & S_z & vbNullString & z_S                          ' "Check1TexasSubCoverageCode, "
            sSQL = sSQL & S_z & .CurrentLossInfo.cli01_CAT & z_S            ' "CatCode, "
            
            sInsuredNames = .AdminLossInfoApd.ali0064_NamedInsured
            sSQL = sSQL & S_z & sInsuredNames & z_S                         ' "NewInsuredNames, "
            
            'Mailing Address
            sMailAddress = .AdminLossInfoApd.ali0065_MailAddress1 & String(2, Chr(32)) & IIf(.AdminLossInfoApd.ali0066_MailAddress2 = vbNullString, vbNullString, S_z & .AdminLossInfoApd.ali0066_MailAddress2)
            goUtil.utFillAddressFields sMailAddress, sMailZip, sMailState, sMailCity, sMailStreet
            goUtil.utUpdateAddress sMailAddress, sMailZip, sMailState, sMailCity, sMailStreet
            
            sSQL = sSQL & S_z & sMailAddress & z_S                          ' "NewMailingAddress, "
            sSQL = sSQL & S_z & sMailStreet & z_S                           ' "MailingAddress, "
            sSQL = sSQL & S_z & sMailZip & z_S                              ' "MailingZip, "
            sSQL = sSQL & S_z & sMailState & z_S                            ' "MailingState, "
            sSQL = sSQL & S_z & sMailCity & z_S                             ' "MailingCity, "
            
            'Loss Location Address
            sLossAddress = .AdminLossInfoApd.cli0063_VehicleLocation
            goUtil.utFillAddressFields sLossAddress, sLossZip, sLossState, sLossCity, sLossStreet
            'Check State against Mailing State
            If InStr(1, sLossAddress, sMailZip, vbTextCompare) > 0 Then
                sLossAddress = Replace(sLossAddress, sMailZip, vbNullString, , , vbTextCompare)
                sLossZip = sMailZip
            End If
            If InStr(1, sLossAddress, sMailState, vbTextCompare) > 0 Then
                sLossAddress = Replace(sLossAddress, " " & sMailState, vbNullString, , , vbTextCompare)
                sLossState = sMailState
            End If
            '-------------------------Debug ONLY------------------------------
    '        Dim lcount As Long
    '        For lcount = 1 To Len(sMailCity)
    '            Debug.Print Mid(sMailCity, lcount, 1) & "------->" & Asc(Mid(sMailCity, lcount, 1))
    '        Next
    '        For lcount = 1 To Len(sLossAddress)
    '            Debug.Print Mid(sLossAddress, lcount, 1) & "------->" & Asc(Mid(sLossAddress, lcount, 1))
    '        Next
            '-------------------------Debug Only-----------------------------
            If InStr(1, sLossAddress, Replace(sMailCity, Chr(160), Chr(32)), vbTextCompare) > 0 Then
                sLossAddress = Replace(sLossAddress, Replace(sMailCity, Chr(160), Chr(32)), vbNullString, , , vbTextCompare)
                sLossCity = sMailCity
            End If
            If InStr(1, sLossStreet, Replace(sMailCity, Chr(160), Chr(32)), vbTextCompare) > 0 Then
                sLossStreet = RTrim(Replace(sLossStreet, Replace(sMailCity, Chr(160), Chr(32)), vbNullString, , , vbTextCompare))
            End If
            
            goUtil.utUpdateAddress sLossAddress, sLossZip, sLossState, sLossCity, sLossStreet
            
            sSQL = sSQL & S_z & sLossAddress & z_S                          ' "NewLossLocation, "
            sSQL = sSQL & S_z & sLossStreet & z_S                           ' "LossAddress, "
            sSQL = sSQL & S_z & sLossZip & z_S                              ' "LossZip, "
            sSQL = sSQL & S_z & sLossState & z_S                            ' "LossState, "
            sSQL = sSQL & S_z & sLossCity & z_S                             ' "LossCity
    
            sSQL = sSQL & S_z & .AdminLossInfoApd.ali0057_MortgageHolder & z_S  ' "MortgageCompany, "
            'Val function Bug in VB6
            'http://msdn.microsoft.com/vbasic/productinfo/previous/vb6/tips/01pasttips.asp
            'Need to parse out both % and !  because these trailing equate to Double and Single
            'and Val bugs because it can't convert Double or single into integer
            sTemp = LTrim(Right(Replace(.AdminLossInfoApd.ali0068_CompDed, ",", vbNullString), 10))
            sSQL = sSQL & Val(goUtil.utCleanValString(sTemp)) & ", " ' "Deductible, " 'CUR
            sTemp = 0
            sSQL = sSQL & Val(goUtil.utCleanValString(sTemp)) & ", " ' "BuildingPolicyLimit, " 'CUR
            sTemp = 0
            sSQL = sSQL & Val(goUtil.utCleanValString(sTemp)) & ", " ' "PersonalPropertyLimit, " 'CUR
            sDate = .CurrentLossInfo.cli02_LossDate
            If Not IsDate(sDate) Then
                sDate = Now()
            End If
            sSQL = sSQL & S_z & Format(CDate(sDate), "YYYYMMDD") & z_S      'Loss Date
            sSQL = sSQL & S_z & sIBNum & z_S                                ' "ClaimNo, "
            If Trim(.AdminLossInfoApd.ali0057_MortgageHolder) = vbNullString Then
                sSQL = sSQL & S_z & vbNullString & z_S                      ' "ConfirmMortgageeIsCorrect, "
                sSQL = sSQL & True & ", "                                   ' "ConfirmMortgCorrNA, "
                sSQL = sSQL & S_z & vbNullString & z_S                      ' "ExplainedMortgageeChecks, "
                sSQL = sSQL & True & ", "                                   ' "ExplainMortgCheckNA, "
            Else
                sSQL = sSQL & S_z & "Y" & z_S                               ' "ConfirmMortgageeIsCorrect, "
                sSQL = sSQL & False & ", "                                  ' "ConfirmMortgCorrNA, "
                sSQL = sSQL & S_z & "Y" & z_S                               ' "ExplainedMortgageeChecks, "
                sSQL = sSQL & False & ", "
            End If
            
            sSQL = sSQL & S_z & "Y" & z_S                                   ' "ExplainedAndGaveEstToCustomer, " ' = "Y"
            sSQL = sSQL & S_z & "Y" & z_S                                   ' "ExplandGaveRCVFormToCust, " '= "Y"
            sSQL = sSQL & S_z & "Y" & z_S                                   ' "ConfirmedCoverage, " ' = "Y"
            sSQL = sSQL & False & ", "                                      ' "PriorLoss, " ' = False
            sSQL = sSQL & False & ", "                                      ' "PriorLossNA, " ' = False
            sSQL = sSQL & dTaxPercent & ", "                                ' "TaxPercent, "
            sSQL = sSQL & S_z & sTypeOfLoss & S_z & " ) "                   ' "TypeOfLoss, "
        End If
    End With
    
    SQL_EASYCLAIM_20000 = sSQL
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    sErrDesc = mudtXML01LossReport.XML01Loss.CurrentLossInfo.cli06_SALN & vbCrLf & sErrDesc
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLR_ClassName & vbCrLf & "Private Function SQL_EASYCLAIM_20000"
End Function

Private Function zEOJ() As String
    'HP4000 Laser End Of Print Job Special Characters
    Dim sEOJ As String
    sEOJ = Chr(10) & Chr(64) & Chr(80) & Chr(74) & Chr(76) & Chr(32) & Chr(69) & Chr(79) & Chr(74) & Chr(32)
    sEOJ = sEOJ & Chr(10) & Chr(27) & Chr(37) & Chr(45) & Chr(49) & Chr(50) & Chr(51) & Chr(52) & Chr(53) & Chr(88)
    zEOJ = sEOJ
End Function

Private Function zSTART_JOB() As String
    'HP4000 Laser Print Job Header
    zSTART_JOB = "@PJL JOB NAME = ""PRO LOSS REPORT"
End Function

Private Function zSTART_JOBApd() As String
    'HP4000 Laser Print Job Header
    zSTART_JOBApd = "@PJL JOB NAME = ""APD LOSS REPORT"
End Function

Private Function zCONTINUED_JOB() As String
    'HP4000 Laser Continued page Header
    zCONTINUED_JOB = "@PJL JOB NAME = ""PAGE"
End Function

Private Function zDM() As String
    'HP4000 Laser Data Delimeter
    zDM = ";LB"
End Function

Private Function zED() As String
    'Endorsement Code
    zED = "ED"
End Function

Private Function zUD_ED() As String
    'Endorse ment code is not defined
    zUD_ED = "UNDEFINED ENDORSEMENT"
End Function

Private Function zCAL_CAT_ALL() As String
    'Under comments Activity Log ALL can be used in place of
    'Cat code. Why ALL is Used ???
    zCAL_CAT_ALL = "ALL"
End Function

Private Function zLBL_SALN() As String
    'SALN lable that appers at the very end of First Page Loss Report
    zLBL_SALN = "SALN:"
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


