VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTestWddx 
   Caption         =   "Form1"
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11130
   LinkTopic       =   "Form1"
   ScaleHeight     =   8310
   ScaleWidth      =   11130
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClean 
      Caption         =   "Clean"
      Height          =   375
      Left            =   6720
      TabIndex        =   9
      Top             =   4440
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000008&
      ForeColor       =   &H80000005&
      Height          =   1335
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   600
      Width           =   8775
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   2520
      TabIndex        =   7
      Top             =   120
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdProcess 
      Caption         =   "&Process Raw Data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   120
      Width           =   1935
   End
   Begin VB.TextBox txtRS 
      Height          =   2535
      Left            =   600
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   4920
      Width           =   8775
   End
   Begin VB.TextBox txtStruct 
      Height          =   1335
      Left            =   600
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   3120
      Width           =   8775
   End
   Begin VB.TextBox txtXmlpath 
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   2280
      Width           =   6855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Deserialize"
      Height          =   375
      Left            =   7680
      TabIndex        =   1
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "RS"
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   4680
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "Structure"
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   2760
      Width           =   4215
   End
End
Attribute VB_Name = "frmTestWddx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private moDeser As WDDXDeserializer
Private moMyRS As WDDXRecordset
Private moMyStruct As WDDXStruct


'loss Report stuff
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
Private WithEvents moLoss As V2ECKeyBoard.clsLossReports
Attribute moLoss.VB_VarHelpID = -1

Private Sub cmdClean_Click()
    Dim sTemp As String
    
    sTemp = goUtil.utGetFileData(txtXmlpath.Text)
    
    sTemp = Replace(sTemp, vbCrLf, vbNullString, , , vbBinaryCompare)
    sTemp = Replace(sTemp, vbTab, vbNullString, , , vbBinaryCompare)
    
    txtRS.Text = sTemp
End Sub

Private Sub Command1_Click()
    Dim sXML As String
    Dim varyNames() As Variant
    Dim varyColumns() As Variant
    Dim lCount As Long
    Dim lCount2 As Long
    Dim sPropName As String
    Dim sColName As String
    Dim oMySer As WDDXSerializer
    Dim oMyStruct As WDDXStruct
    sXML = GetFileData(txtXmlpath.Text)
'    If InStr(1, sXML, Chr(60), vbBinaryCompare) > 0 Then
'        MsgBox "Hey theres a VBCRLF in there!"
'    End If
    'Since the example wddx xml contains Structure and Rs
'    Debug.Print Err.Description & vbCrLf & Err.Number
    Set moMyStruct = moDeser.deserialize(sXML)
    
'    Set oMySer = New WDDXSerializer
'    Set oMyStruct = New WDDXStruct
'
'    oMyStruct.setProp "TESTTHIS", "Testthis:"
'
'    sXML = oMySer.serialize(oMyStruct)
'
'    MsgBox sXML
'    Exit Sub
    'set variant array to get Names in the Structure
    varyNames() = moMyStruct.getPropNames
    For lCount = LBound(varyNames(), 1) To UBound(varyNames(), 1)
        txtStruct.Text = txtStruct.Text & varyNames(lCount) & vbCrLf
    Next
    
    For lCount2 = LBound(varyNames(), 1) To UBound(varyNames(), 1)
        sPropName = varyNames(lCount2)
        If StrComp(sPropName, "ClassName", vbTextCompare) = 0 Then
            txtRS.Text = txtRS.Text & vbCrLf & String(100, "*")
            txtRS.Text = txtRS.Text & vbCrLf & sPropName & " = " & moMyStruct.getProp(sPropName) & " "
            txtRS.Text = txtRS.Text & vbCrLf & String(100, "*")
        ElseIf StrComp(sPropName, "TransType", vbTextCompare) = 0 Then
            txtRS.Text = txtRS.Text & vbCrLf & String(100, "*")
            txtRS.Text = txtRS.Text & vbCrLf & sPropName & " = " & moMyStruct.getProp(sPropName) & " "
            txtRS.Text = txtRS.Text & vbCrLf & String(100, "*")
        ElseIf StrComp(sPropName, "LossType", vbTextCompare) = 0 Then
            txtRS.Text = txtRS.Text & vbCrLf & String(100, "*")
            txtRS.Text = txtRS.Text & vbCrLf & sPropName & " = " & moMyStruct.getProp(sPropName) & " "
            txtRS.Text = txtRS.Text & vbCrLf & String(100, "*")
        Else
            Set moMyRS = moMyStruct.getProp(sPropName)
            txtRS.Text = txtRS.Text & vbCrLf & String(100, "*")
            txtRS.Text = txtRS.Text & vbCrLf & sPropName & " Records Row(1)  Row Count = (" & moMyRS.getRowCount & ") "
            txtRS.Text = txtRS.Text & vbCrLf & String(100, "*") & vbCrLf & vbCrLf
            varyColumns() = moMyRS.getColumnNames
            For lCount = LBound(varyColumns(), 1) To UBound(varyColumns(), 1)
                sColName = varyColumns(lCount)
                txtRS.Text = txtRS.Text & varyColumns(lCount) & " = " & moMyRS.getField(1, sColName) & vbCrLf
            Next
        End If
        
    Next
    
    
    

End Sub

Private Sub Form_Load()
    Set moDeser = New WDDXDeserializer
    Set moMyRS = New WDDXRecordset
    Set moMyStruct = New WDDXStruct
    Set goUtil = New V2ECKeyBoard.clsUtil
    txtXmlpath.Text = App.Path & "\FixedTransformed.xml"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set moDeser = Nothing
    Set moMyRS = Nothing
    Set moMyStruct = Nothing
    Set goUtil = Nothing
End Sub

Public Function GetFileData(psFilePath As String, Optional pbLock As Boolean = False, Optional piFFile As Integer, Optional pbSkipMess As Boolean = True) As String
    On Error GoTo EH
    Dim lMyFileLen As Long
    Dim iFFile As Integer
    
    iFFile = FreeFile
    piFFile = iFFile
    If pbLock Then
        Open psFilePath For Binary Access Read Lock Read As #iFFile
    Else
        Open psFilePath For Binary Access Read As #iFFile
    End If
    lMyFileLen = FileLen(psFilePath) + 2
    GetFileData = Input(lMyFileLen, #iFFile)
    If Not pbLock Then
        Close #iFFile
    End If
    
    Exit Function
EH:
    Close #iFFile
    If Not pbSkipMess Then
        If MsgBox("Could not read file... " & vbCrLf & psFilePath & vbCrLf & "(" & Err.Description & ")" & vbCrLf & vbCrLf & _
                  "The network or file is busy." & vbCrLf & "Press ""Yes"" to try again." & vbCrLf & "Press ""No"" to abort this process", vbYesNo, "File is Busy") = vbYes Then
            Resume
        End If
    End If
    
End Function

Private Sub cmdProcess_Click()
    On Error GoTo EH
    Dim sMess As String
    Dim sclsLoss As String
    Dim sAppEXEName As String
    Dim sFormat As String
    Dim sRawDataPath As String
    Dim sFTPPath As String
    
    If Not moLoss Is Nothing Then
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    sAppEXEName = "XMLTypes"

    sFormat = "clsLossXML01"
    
    sRawDataPath = App.Path & "\RawData\"
    sFTPPath = App.Path & "\XMLData\"
    
    If Not goUtil.utFileExists(sRawDataPath, True) Then
        goUtil.utMakeDir sRawDataPath
    End If
    If Not goUtil.utFileExists(sFTPPath, True) Then
        goUtil.utMakeDir sFTPPath
    End If
    
    Set moLoss = New V2ECKeyBoard.clsLossReports
    moLoss.SetUtilObject goUtil
    moLoss.IgnoreProcessRawDataErrors = False
    'Very Important we set the Application we want ECWeb to access
    'FARMERS in this case
    moLoss.APPEXEName = sAppEXEName
    goUtil.gsAppEXEName = sAppEXEName
    goUtil.gsMainAppEXEName = sAppEXEName
    goUtil.SetUtilObject goUtil
    
    
    '1.5.2004 Not Applicable in V2
'    SendAdjusterTable
   
    'Current version Web Control will process ASN and CCMS formats
    'as well, we handle processing but not DB update of Unknown Text only formats
    'ASN
    
    If Not moLoss.ProcessRawData(sFormat, sRawDataPath, sFTPPath, ProgressBar1, Text1) Then
        MsgBox "Nothing to process!"
    Else
        MsgBox "Process Success!"
'        moLoss.ShowLossReports , , False
'        Do
'            DoEvents
'            Sleep 100
'        Loop
    End If
    
    
    Screen.MousePointer = vbDefault
CLEAN_UP:
    moLoss.CLEANUP
    Set moLoss = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
EH:
    MsgBox "error num: " & Err.Number & vbCrLf & Err.Description, vbCritical
    Me.MousePointer = vbDefault
    Set moLoss = Nothing
End Sub

Private Sub moLoss_ErrorMess(ByVal Mess As String)
    MsgBox Mess, vbCritical
End Sub

Private Sub moLoss_UpdateDB(ByVal oLossReport As V2ECKeyBoard.clsCarLR)
    On Error GoTo EH
    Dim sMess As String
    Dim sXML As String
    Dim sType As String
    Dim sSaln As String
    
    
    sXML = GetXMLLoss(oLossReport)
    
    If oLossReport.LossType = TypeXML01.XML01Apd Then
        sType = "APD"
    ElseIf oLossReport.LossType = TypeXML01.XML01Pro Then
        sType = "PRO"
    End If
    sSaln = oLossReport.CLIENTNUM
    
    goUtil.utSaveFileData App.Path & "\XMLData\" & sSaln & "_" & sType & ".xml", sXML
    
    
    Exit Sub
EH:
    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
    sMess = sMess & "ERROR # " & Err.Number & vbCrLf
    sMess = sMess & Err.Description & vbCrLf & vbCrLf
    sMess = sMess & "Private Sub moLoss_UpdateDB" & vbCrLf
    sMess = sMess & Me.Name & vbCrLf
    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
    MsgBox sMess, vbCritical
    Err.Clear
    Resume Next
End Sub


Public Function GetXMLLoss(ByVal oLossReport As V2ECKeyBoard.clsCarLR) As String
    On Error GoTo EH

    'Export Report Collection Items
    Dim oMySer As WDDXSerializer        'Allaire's WDDX serializer
    Dim oMyStruct As WDDXStruct         'Allaire's WDDX Structure (Cold Fusion Strucuture type)
    
    'Loss Report
    Dim sLossType As String
    Dim MyXML01LossReport As XMLTypes.XML01LossReport
    Dim MyudtXML01Loss As XMLTypes.udtXML01Loss
    'Current Loss Info
    Dim oAssignmentDetailRS As WDDXRecordset
    'Property
    Dim oLossDetailRS As WDDXRecordset
    'Auto
    Dim oVehicleDetailRS As WDDXRecordset
    'Policy Units
    Dim oPolicyDetailRS  As WDDXRecordset
    'ContactDetails
    Dim oContactDetailRS As WDDXRecordset
    'Additional Coverages
    Dim oCoverageRS As WDDXRecordset
    'Endorsements
    Dim oEndorsementRS As WDDXRecordset
    'Payment Detail
    Dim oPaymentDetailRS As WDDXRecordset
    'PriorLossDetail
    Dim oPriorLossDetailRS As WDDXRecordset
    'Comments Act Log
    Dim oActivitiesRS As WDDXRecordset
    
    MyXML01LossReport = oLossReport.LossReport
    MyudtXML01Loss = MyXML01LossReport.XML01Loss
       
    'Current Loss Info
    Set oAssignmentDetailRS = GetAssignmentDetailRS(oLossReport)
    'Check for Auto Or Property
    If oLossReport.LossType = TypeXML01.XML01Apd Then
        Set oVehicleDetailRS = GetVehicleDetailRS(oLossReport)
    ElseIf oLossReport.LossType = TypeXML01.XML01Pro Then
        Set oLossDetailRS = GetLossDetailRS(oLossReport)
    End If
     'ContactDetails
    Set oContactDetailRS = GetContactDetailRS(oLossReport)
    'Policy Unit
    Set oPolicyDetailRS = GetPolicyDetailRS(oLossReport)
    'Additional COverages
    Set oCoverageRS = GetCoverageRS(oLossReport)
    'Endorsements
    Set oEndorsementRS = GetEndorsementRS(oLossReport)
    'payment Detail
    Set oPaymentDetailRS = GetPaymentDetailRS(oLossReport)
    'PriorLossDetail
    Set oPriorLossDetailRS = GetPriorLossDetailRS(oLossReport)
    'Comments Act Log
    Set oActivitiesRS = GetActivitiesRS(oLossReport)

    'Create WDDX Structure
    Set oMyStruct = New WDDXStruct
    
    oMyStruct.setProp "ClassName", "V2ECcarFarmers.clsLossXML01"
    oMyStruct.setProp "TransType", "IA_CRN_ASSIGN"
    'Auto Or Property
    If oLossReport.LossType = TypeXML01.XML01Apd Then
        'Set Type and Admin Info for Auto
        oMyStruct.setProp "LossType", "Auto"
        'Set AssignmentDetail RS
        If Not oAssignmentDetailRS Is Nothing Then
            oMyStruct.setProp "AssignmentDetailRS", oAssignmentDetailRS
        End If
        oMyStruct.setProp "VehicleDetailRS", oVehicleDetailRS
    ElseIf oLossReport.LossType = TypeXML01.XML01Pro Then
        'Set Type and Admin Info for Property
        oMyStruct.setProp "LossType", "Property"
        'Set AssignmentDetail RS
        If Not oAssignmentDetailRS Is Nothing Then
            oMyStruct.setProp "AssignmentDetailRS", oAssignmentDetailRS
        End If
        oMyStruct.setProp "LossDetailRS", oLossDetailRS
    End If
    'Set the ContactDetailRS
    If Not oContactDetailRS Is Nothing Then
        oMyStruct.setProp "ContactDetailRS", oContactDetailRS
    End If
    'Set the PolicyUnit RS
    If Not oPolicyDetailRS Is Nothing Then
        oMyStruct.setProp "PolicyDetailRS", oPolicyDetailRS
    End If
    'Set Additional Coverages RS
    If Not oCoverageRS Is Nothing Then
        oMyStruct.setProp "CoverageRS", oCoverageRS
    End If
    'Set Endorsements RS
    If Not oEndorsementRS Is Nothing Then
        oMyStruct.setProp "EndorsementRS", oEndorsementRS
    End If
    'Set Payment Detail
    If Not oPaymentDetailRS Is Nothing Then
        oMyStruct.setProp "PaymentDetailRS", oPaymentDetailRS
    End If
    'Set Prior Loss History
    If Not oPriorLossDetailRS Is Nothing Then
        oMyStruct.setProp "PriorLossDetailRS", oPriorLossDetailRS
    End If
    'Set Comments Act Log RS
    If Not oActivitiesRS Is Nothing Then
        oMyStruct.setProp "ActivitiesRS", oActivitiesRS
    End If
    
    Set oMySer = New WDDXSerializer
    
    GetXMLLoss = oMySer.serialize(oMyStruct)
    
    'CLEANUP
    Set oAssignmentDetailRS = Nothing
    Set oVehicleDetailRS = Nothing
    Set oLossDetailRS = Nothing
    Set oPolicyDetailRS = Nothing
    Set oContactDetailRS = Nothing
    Set oCoverageRS = Nothing
    Set oEndorsementRS = Nothing
    Set oPriorLossDetailRS = Nothing
    Set oPaymentDetailRS = Nothing
    Set oActivitiesRS = Nothing
    Set oMyStruct = Nothing
    Set oMySer = Nothing
    
    Exit Function
EH:
    GetXMLLoss = "Version: " & "CRN_IA_000001" & vbCrLf
    GetXMLLoss = GetXMLLoss & "Error # " & Err.Number & vbCrLf
    GetXMLLoss = GetXMLLoss & "Description: " & vbCrLf
    GetXMLLoss = GetXMLLoss & Err.Description
End Function

Public Function GetAssignmentDetailRS(ByVal oLossReport As V2ECKeyBoard.clsCarLR) As WDDXRecordset
    Dim MyXML01LossReport As XMLTypes.XML01LossReport
    Dim MyudtXML01Loss As XMLTypes.udtXML01Loss
    Dim MyudtXML01AssignmentDetail As XMLTypes.udtXML01CurrentLossInfo
    Dim vValue As Variant 'Use Variant to Get Variant Equiv Data Type.
    Dim sTemp As String
    
    MyXML01LossReport = oLossReport.LossReport
    MyudtXML01Loss = MyXML01LossReport.XML01Loss
    
    MyudtXML01AssignmentDetail = MyXML01LossReport.XML01Loss.CurrentLossInfo
        
    Set GetAssignmentDetailRS = New WDDXRecordset
    GetAssignmentDetailRS.addColumn "UnitNumber"
    GetAssignmentDetailRS.addColumn "CatastropheCode"
    GetAssignmentDetailRS.addColumn "LossDate"
    GetAssignmentDetailRS.addColumn "Type"
    GetAssignmentDetailRS.addColumn "CauseOfLoss"
    GetAssignmentDetailRS.addColumn "FirstName"
    GetAssignmentDetailRS.addColumn "LastName"
    GetAssignmentDetailRS.addColumn "AssignedTo"
    GetAssignmentDetailRS.addColumn "AssignedToFirstName"
    GetAssignmentDetailRS.addColumn "AssignedToLastName"
    
    'Put one row for the Data RS
    GetAssignmentDetailRS.addRows 1
        
    With MyudtXML01AssignmentDetail
        vValue = "test unit number 123"
        GetAssignmentDetailRS.setField 1, "UnitNumber", vValue
        vValue = .cli01_CAT
        GetAssignmentDetailRS.setField 1, "CatastropheCode", vValue
        vValue = .cli02_LossDate
        GetAssignmentDetailRS.setField 1, "LossDate", vValue
        vValue = "test type Building"
        GetAssignmentDetailRS.setField 1, "Type", vValue
        vValue = "WIND"
        GetAssignmentDetailRS.setField 1, "CauseOfLoss", vValue
        If oLossReport.LossType = TypeXML01.XML01Apd Then
            sTemp = Trim(MyudtXML01Loss.AdminLossInfoApd.ali0064_NamedInsured)
        ElseIf oLossReport.LossType = TypeXML01.XML01Pro Then
            sTemp = Trim(MyudtXML01Loss.AdminLossInfo.ali0065_MainFileInsuredName)
        End If
        vValue = Trim(Left(sTemp, InStr(1, sTemp, " ", vbTextCompare)))
        GetAssignmentDetailRS.setField 1, "FirstName", vValue
        If oLossReport.LossType = TypeXML01.XML01Apd Then
            sTemp = Trim(MyudtXML01Loss.AdminLossInfoApd.ali0064_NamedInsured)
        ElseIf oLossReport.LossType = TypeXML01.XML01Pro Then
            sTemp = Trim(MyudtXML01Loss.AdminLossInfo.ali0065_MainFileInsuredName)
        End If
        vValue = Trim(Mid(sTemp, InStrRev(sTemp, " ", , vbBinaryCompare)))
        GetAssignmentDetailRS.setField 1, "LastName", vValue
        vValue = .cli03_Adjuster
        GetAssignmentDetailRS.setField 1, "AssignedTo", vValue
        vValue = "Adjuster First Name"
        GetAssignmentDetailRS.setField 1, "AssignedToFirstName", vValue
        vValue = "Adjuster Last Name"
        GetAssignmentDetailRS.setField 1, "AssignedToLastName", vValue
    End With
        
End Function

Public Function GetContactDetailRS(ByVal oLossReport As V2ECKeyBoard.clsCarLR) As WDDXRecordset
    Dim MyXML01LossReport As XMLTypes.XML01LossReport
    Dim MyudtXML01Loss As XMLTypes.udtXML01Loss
    Dim MyudtContactDetail As XMLTypes.udtContact
    Dim vValue As Variant 'Use Variant to Get Variant Equiv Data Type.
    
    MyXML01LossReport = oLossReport.LossReport
    MyudtXML01Loss = MyXML01LossReport.XML01Loss
        
    Set GetContactDetailRS = New WDDXRecordset
    
    GetContactDetailRS.addColumn "AgentFirstName"
    GetContactDetailRS.addColumn "AgentLastName"
    GetContactDetailRS.addColumn "AgentPrimaryPhone"
    GetContactDetailRS.addColumn "ContactRowID"
    GetContactDetailRS.addColumn "FirstName"
    GetContactDetailRS.addColumn "LastName"
    GetContactDetailRS.addColumn "RelationshipToInsured"
    GetContactDetailRS.addColumn "PrimaryPhoneNumber"
    GetContactDetailRS.addColumn "HomePhoneNumber"
    GetContactDetailRS.addColumn "CellularPhoneNumber"
    GetContactDetailRS.addColumn "WorkPhoneNumber"
    GetContactDetailRS.addColumn "Type"
    GetContactDetailRS.addColumn "StreetAddress"
    GetContactDetailRS.addColumn "StreetAddress2"
    GetContactDetailRS.addColumn "City"
    GetContactDetailRS.addColumn "State"
    GetContactDetailRS.addColumn "PostalCode"
    
    'Put a couple of  rows for the Data RS
    GetContactDetailRS.addRows 1
    
    '1 Get the Insured and use it as the first row
    With MyudtContactDetail
        .AgentFirstName = "AgentFirstName"
        .AgentLastName = "AgentLastName"
        .AgentPrimaryPhone = "AgentPrimaryPhone"
        .ContactRowID = "ContactRowID"
        If oLossReport.LossType = XML01Pro Then
            .FirstName = MyXML01LossReport.XML01Loss.AdminLossInfo.ali0062_NamedInsured
            .RelationshipToInsured = "RelationshipToInsured"
            .LastName = vbNullString
            .PrimaryPhoneNumber = "PrimaryPhoneNumber"
            .HomePhoneNumber = MyXML01LossReport.XML01Loss.AdminLossInfo.ali0057_HomePhone
            .WorkPhoneNumber = MyXML01LossReport.XML01Loss.AdminLossInfo.ali0058_BusPhone
            .CellularPhoneNumber = "CellularPhoneNumber"
        ElseIf oLossReport.LossType = XML01Apd Then
            .FirstName = MyXML01LossReport.XML01Loss.AdminLossInfoApd.ali0064_NamedInsured
            .RelationshipToInsured = "RelationshipToInsured"
            .LastName = vbNullString
            .PrimaryPhoneNumber = "PrimaryPhoneNumber"
            .HomePhoneNumber = MyXML01LossReport.XML01Loss.AdminLossInfoApd.ali0054_HomePhone
            .WorkPhoneNumber = MyXML01LossReport.XML01Loss.AdminLossInfoApd.ali0055_BusPhone
            .CellularPhoneNumber = "CellularPhoneNumber"
        End If
        .Type = "Type"
        .StreetAddress = "StreetAddress"
        .StreetAddress2 = "StreetAddress2"
        .City = "City"
        .State = "State"
        .PostalCode = "PostalCode"
    End With
        
    With MyudtContactDetail
        vValue = .AgentFirstName
        GetContactDetailRS.setField 1, "AgentFirstName", vValue
        vValue = .AgentLastName
        GetContactDetailRS.setField 1, "AgentLastName", vValue
        vValue = .AgentPrimaryPhone
        GetContactDetailRS.setField 1, "AgentPrimaryPhone", vValue
        vValue = .ContactRowID
        GetContactDetailRS.setField 1, "ContactRowID", vValue
        vValue = .FirstName
        GetContactDetailRS.setField 1, "FirstName", vValue
        vValue = .LastName
        GetContactDetailRS.setField 1, "LastName", vValue
        vValue = .RelationshipToInsured
        GetContactDetailRS.setField 1, "RelationshipToInsured", vValue
        vValue = .PrimaryPhoneNumber
        GetContactDetailRS.setField 1, "PrimaryPhoneNumber", vValue
        vValue = .HomePhoneNumber
        GetContactDetailRS.setField 1, "HomePhoneNumber", vValue
        vValue = .CellularPhoneNumber
        GetContactDetailRS.setField 1, "CellularPhoneNumber", vValue
        vValue = .WorkPhoneNumber
        GetContactDetailRS.setField 1, "WorkPhoneNumber", vValue
        vValue = .Type
        GetContactDetailRS.setField 1, "Type", vValue
        vValue = .StreetAddress
        GetContactDetailRS.setField 1, "StreetAddress", vValue
        vValue = .StreetAddress2
        GetContactDetailRS.setField 1, "StreetAddress2", vValue
        vValue = .City
        GetContactDetailRS.setField 1, "City", vValue
        vValue = .State
        GetContactDetailRS.setField 1, "State", vValue
        vValue = .PostalCode
        GetContactDetailRS.setField 1, "PostalCode", vValue
    End With
        
End Function

Public Function GetDeductiblesRS(ByVal oLossReport As V2ECKeyBoard.clsCarLR) As WDDXRecordset
    Dim MyXML01LossReport As XMLTypes.XML01LossReport
    Dim MyudtXML01Loss As XMLTypes.udtXML01Loss
    Dim MyudtDeductibles As XMLTypes.udtDeductibles
    Dim vValue As Variant 'Use Variant to Get Variant Equiv Data Type.
    
    MyXML01LossReport = oLossReport.LossReport
    MyudtXML01Loss = MyXML01LossReport.XML01Loss
    
    Set GetDeductiblesRS = New WDDXRecordset
    GetDeductiblesRS.addColumn "DeductibleDescription"
    
    If oLossReport.LossType = XML01Pro Then
        'Put 4 rows for the Data RS
        GetDeductiblesRS.addRows 4
        With MyudtXML01Loss.AdminLossInfo
            vValue = .ali0072_Deductible1
            GetDeductiblesRS.setField 1, "DeductibleDescription", vValue
            vValue = .ali0073_Deductible2
            GetDeductiblesRS.setField 2, "DeductibleDescription", vValue
            vValue = .ali0074_Deductible3
            GetDeductiblesRS.setField 3, "DeductibleDescription", vValue
            vValue = .ali0075_Deductible4
            GetDeductiblesRS.setField 4, "DeductibleDescription", vValue
        End With
    ElseIf oLossReport.LossType = XML01Apd Then
        'Put 1 row for the Data RS
        GetDeductiblesRS.addRows 1
        With MyudtXML01Loss.AdminLossInfoApd
            vValue = .ali0068_CompDed
            GetDeductiblesRS.setField 1, "DeductibleDescription", vValue
        End With
    End If
        
End Function

Public Function GetCoverageRS(ByVal oLossReport As V2ECKeyBoard.clsCarLR) As WDDXRecordset
    Dim MyXML01LossReport As XMLTypes.XML01LossReport
    Dim MyudtXML01Loss As XMLTypes.udtXML01Loss
    Dim MyudtAdditionalCoverages As XMLTypes.udtAdditionalCoverages
    Dim MyudtCoverage As XMLTypes.udtCoverage
    Dim vValue As Variant 'Use Variant to Get Variant Equiv Data Type.
    
    MyXML01LossReport = oLossReport.LossReport
    MyudtXML01Loss = MyXML01LossReport.XML01Loss
    
    Set GetCoverageRS = New WDDXRecordset
    GetCoverageRS.addColumn "Coverage"
    GetCoverageRS.addColumn "Limits"
    GetCoverageRS.addColumn "Deductible1"
    GetCoverageRS.addColumn "Deductible2"
    GetCoverageRS.addColumn "Deductible3"
    GetCoverageRS.addColumn "Deductible4"
   
    If oLossReport.LossType = XML01Pro Then
        'Put 4 rows for the Data RS
        GetCoverageRS.addRows 2
        'building
        With MyudtCoverage
            .Coverage = "Building"
            vValue = MyXML01LossReport.XML01Loss.AdminLossInfo.ali0070_BldgLimit
            .Limits = CheckForNullCurrency(vValue)
            .Deductible1 = MyXML01LossReport.XML01Loss.AdminLossInfo.ali0072_Deductible1
            .Deductible2 = MyXML01LossReport.XML01Loss.AdminLossInfo.ali0073_Deductible2
            .Deductible3 = MyXML01LossReport.XML01Loss.AdminLossInfo.ali0074_Deductible3
            .Deductible4 = MyXML01LossReport.XML01Loss.AdminLossInfo.ali0075_Deductible4
            vValue = .Coverage
            GetCoverageRS.setField 1, "Coverage", vValue
            vValue = .Limits
            GetCoverageRS.setField 1, "Limits", vValue
            vValue = .Deductible1
            GetCoverageRS.setField 1, "Deductible1", vValue
            vValue = .Deductible2
            GetCoverageRS.setField 1, "Deductible2", vValue
            vValue = .Deductible3
            GetCoverageRS.setField 1, "Deductible3", vValue
            vValue = .Deductible4
            GetCoverageRS.setField 1, "Deductible4", vValue
        End With
        'Contents
        With MyudtCoverage
            .Coverage = "Contents"
            vValue = MyXML01LossReport.XML01Loss.AdminLossInfo.ali0071_ContLimit
            .Limits = CheckForNullCurrency(vValue)
            .Deductible1 = MyXML01LossReport.XML01Loss.AdminLossInfo.ali0072_Deductible1
            .Deductible2 = MyXML01LossReport.XML01Loss.AdminLossInfo.ali0073_Deductible2
            .Deductible3 = MyXML01LossReport.XML01Loss.AdminLossInfo.ali0074_Deductible3
            .Deductible4 = MyXML01LossReport.XML01Loss.AdminLossInfo.ali0075_Deductible4
            vValue = .Coverage
            GetCoverageRS.setField 2, "Coverage", vValue
            vValue = .Limits
            GetCoverageRS.setField 2, "Limits", vValue
            vValue = .Deductible1
            GetCoverageRS.setField 2, "Deductible1", vValue
            vValue = .Deductible2
            GetCoverageRS.setField 2, "Deductible2", vValue
            vValue = .Deductible3
            GetCoverageRS.setField 2, "Deductible3", vValue
            vValue = .Deductible4
            GetCoverageRS.setField 2, "Deductible4", vValue
        End With
       
    ElseIf oLossReport.LossType = XML01Apd Then
        'Put 1 row for the Data RS
        GetCoverageRS.addRows 1
        'Contents
        With MyudtCoverage
            .Coverage = "Comp"
            vValue = 15000
            .Limits = CheckForNullCurrency(vValue)
            .Deductible1 = MyXML01LossReport.XML01Loss.AdminLossInfoApd.ali0068_CompDed
            vValue = .Coverage
            GetCoverageRS.setField 1, "Coverage", vValue
            vValue = .Limits
            GetCoverageRS.setField 1, "Limits", vValue
            vValue = .Deductible1
            GetCoverageRS.setField 1, "Deductible1", vValue
        End With
    End If
    
     
        
End Function

Public Function GetEndorsementRS(ByVal oLossReport As V2ECKeyBoard.clsCarLR) As WDDXRecordset
    Dim MyXML01LossReport As XMLTypes.XML01LossReport
    Dim MyudtXML01Loss As XMLTypes.udtXML01Loss
    Dim MyudtXML01Endorsement As XMLTypes.udtXML01Endorsement
    Dim lCount As Long
    Dim vValue As Variant 'Use Variant to Get Variant Equiv Data Type.
    
    MyXML01LossReport = oLossReport.LossReport
    MyudtXML01Loss = MyXML01LossReport.XML01Loss
    
    'if the endorsment collection does not exist then Bail
    If MyudtXML01Loss.colEndorsements Is Nothing Then
        Exit Function
    ElseIf MyudtXML01Loss.colEndorsements.Count = 0 Then
        Exit Function
    End If
    
    MyXML01LossReport = oLossReport.LossReport
    MyudtXML01Loss = MyXML01LossReport.XML01Loss
    
    Set GetEndorsementRS = New WDDXRecordset
    
    GetEndorsementRS.addColumn "EndorsementNumber"
    GetEndorsementRS.addColumn "EndorsementDescription"
    'Put as many rows as items in the Endorsement Collection
    GetEndorsementRS.addRows MyudtXML01Loss.colEndorsements.Count
    
    
    For lCount = 1 To MyudtXML01Loss.colEndorsements.Count
        MyudtXML01Endorsement = MyudtXML01Loss.colEndorsements(lCount)
        With MyudtXML01Endorsement
            vValue = .EDCode
            GetEndorsementRS.setField lCount, "EndorsementNumber", vValue
            vValue = .EDDescription
            GetEndorsementRS.setField lCount, "EndorsementDescription", vValue
        End With
    Next
    
        
End Function

Public Function GetPaymentDetailRS(ByVal oLossReport As V2ECKeyBoard.clsCarLR) As WDDXRecordset
    Dim MyXML01LossReport As XMLTypes.XML01LossReport
    Dim MyudtXML01Loss As XMLTypes.udtXML01Loss
    Dim lCount As Long
    Dim vValue As Variant 'Use Variant to Get Variant Equiv Data Type.
    
    MyXML01LossReport = oLossReport.LossReport
    MyudtXML01Loss = MyXML01LossReport.XML01Loss
    
    Set GetPaymentDetailRS = New WDDXRecordset
    
    GetPaymentDetailRS.addColumn "DateIssued"
    GetPaymentDetailRS.addColumn "PayeeLineOne"
    GetPaymentDetailRS.addColumn "PayeeLineTwo"
    GetPaymentDetailRS.addColumn "PayeeLineThree"
    GetPaymentDetailRS.addColumn "PayeeLineFour"
    GetPaymentDetailRS.addColumn "AccountType"
    GetPaymentDetailRS.addColumn "PaymentClass"
    GetPaymentDetailRS.addColumn "PaymentAmount"

    'Put as many rows as items in the Collection
    GetPaymentDetailRS.addRows 1
    
    lCount = 1
    With MyXML01LossReport
        vValue = "3/6/2005"
        GetPaymentDetailRS.setField lCount, "DateIssued", vValue
        If oLossReport.LossType = TypeXML01.XML01Apd Then
           vValue = .XML01Loss.AdminLossInfoApd.ali0064_NamedInsured
        ElseIf oLossReport.LossType = TypeXML01.XML01Pro Then
            vValue = .XML01Loss.AdminLossInfo.ali0065_MainFileInsuredName
        End If
        GetPaymentDetailRS.setField lCount, "PayeeLineOne", vValue
        vValue = vbNullString
        GetPaymentDetailRS.setField lCount, "PayeeLineTwo", vValue
        If oLossReport.LossType = TypeXML01.XML01Apd Then
           vValue = .XML01Loss.AdminLossInfoApd.ali0065_MailAddress1
        ElseIf oLossReport.LossType = TypeXML01.XML01Pro Then
            vValue = .XML01Loss.AdminLossInfo.ali0063_MailAddress1
        End If
        GetPaymentDetailRS.setField lCount, "PayeeLineThree", vValue
         If oLossReport.LossType = TypeXML01.XML01Apd Then
           vValue = .XML01Loss.AdminLossInfoApd.ali0066_MailAddress2
        ElseIf oLossReport.LossType = TypeXML01.XML01Pro Then
            vValue = .XML01Loss.AdminLossInfo.ali0064_MailAddress2
        End If
        GetPaymentDetailRS.setField lCount, "PayeeLineFour", vValue
        vValue = "test Account Type EFT ??"
        GetPaymentDetailRS.setField lCount, "AccountType", vValue
        vValue = "test Payment Class 01"
        GetPaymentDetailRS.setField lCount, "PaymentClass", vValue
        vValue = "555.55"
        GetPaymentDetailRS.setField lCount, "PaymentAmount", vValue
    End With
    
End Function


Public Function GetPriorLossDetailRS(ByVal oLossReport As V2ECKeyBoard.clsCarLR) As WDDXRecordset
    Dim MyXML01LossReport As XMLTypes.XML01LossReport
    Dim MyudtXML01Loss As XMLTypes.udtXML01Loss
    Dim MyudtXML01PriorLossDetail As XMLTypes.udtXML01PriorLossHist
    Dim lCount As Long
    Dim vValue As Variant 'Use Variant to Get Variant Equiv Data Type.
    
    MyXML01LossReport = oLossReport.LossReport
    MyudtXML01Loss = MyXML01LossReport.XML01Loss
    
    'if the endorsment collection does not exist then Bail
    If MyudtXML01Loss.colPLH Is Nothing Then
        Exit Function
    ElseIf MyudtXML01Loss.colPLH.Count = 0 Then
        Exit Function
    End If
    
    MyXML01LossReport = oLossReport.LossReport
    MyudtXML01Loss = MyXML01LossReport.XML01Loss
    
    Set GetPriorLossDetailRS = New WDDXRecordset
    
    GetPriorLossDetailRS.addColumn "SALN"
    GetPriorLossDetailRS.addColumn "ClaimSegmentNumber"
    GetPriorLossDetailRS.addColumn "PolicyNumber"
    GetPriorLossDetailRS.addColumn "LossCause"
    GetPriorLossDetailRS.addColumn "ClaimClass"
    GetPriorLossDetailRS.addColumn "LossDate"
    GetPriorLossDetailRS.addColumn "SummaryAmount"

    
    'Put as many rows as items in the Collection
    GetPriorLossDetailRS.addRows MyudtXML01Loss.colPLH.Count
    
    For lCount = 1 To MyudtXML01Loss.colPLH.Count
        MyudtXML01PriorLossDetail = MyudtXML01Loss.colPLH(lCount)
        With MyudtXML01PriorLossDetail
            vValue = .plh01_SALN
            GetPriorLossDetailRS.setField lCount, "SALN", vValue
            vValue = "test ClaimSegmentNumber 123"
            GetPriorLossDetailRS.setField lCount, "ClaimSegmentNumber", vValue
            vValue = MyudtXML01Loss.AdminLossInfo.ali0054_PolicyNum
            GetPriorLossDetailRS.setField lCount, "PolicyNumber", vValue
            vValue = "Wind And HAIL"
            GetPriorLossDetailRS.setField lCount, "LossCause", vValue
            vValue = "01"
            GetPriorLossDetailRS.setField lCount, "ClaimClass", vValue
            vValue = .plh02_LossDate
            GetPriorLossDetailRS.setField lCount, "LossDate", CheckForNullDate(vValue)
            vValue = .plh06_AmtPaid
            GetPriorLossDetailRS.setField lCount, "SummaryAmount", CheckForNullCurrency(vValue)
        End With
    Next
        
End Function

Public Function GetActivitiesRS(ByVal oLossReport As V2ECKeyBoard.clsCarLR) As WDDXRecordset
    Dim MyXML01LossReport As XMLTypes.XML01LossReport
    Dim MyudtXML01Loss As XMLTypes.udtXML01Loss
    Dim MyudtXML01Activities As XMLTypes.udtXML01CommentsActLog
    Dim lCount As Long
    Dim vValue As Variant 'Use Variant to Get Variant Equiv Data Type.
    
    MyXML01LossReport = oLossReport.LossReport
    MyudtXML01Loss = MyXML01LossReport.XML01Loss
    
    'if the endorsment collection does not exist then Bail
    If MyudtXML01Loss.colCAL Is Nothing Then
        Exit Function
    ElseIf MyudtXML01Loss.colCAL.Count = 0 Then
        Exit Function
    End If
    
    MyXML01LossReport = oLossReport.LossReport
    MyudtXML01Loss = MyXML01LossReport.XML01Loss
    
    Set GetActivitiesRS = New WDDXRecordset
    
    GetActivitiesRS.addColumn "GMTCreated"
    GetActivitiesRS.addColumn "CreatedBy"
    GetActivitiesRS.addColumn "Type"
    GetActivitiesRS.addColumn "Description"
    GetActivitiesRS.addColumn "Comment"
    
    'Put as many rows as items in the Collection
    GetActivitiesRS.addRows MyudtXML01Loss.colCAL.Count
    
    For lCount = 1 To MyudtXML01Loss.colCAL.Count
        MyudtXML01Activities = MyudtXML01Loss.colCAL(lCount)
        With MyudtXML01Activities
            vValue = .cal02_Date & " " & .cal03_Time
            GetActivitiesRS.setField lCount, "GMTCreated", CheckForNullDate(vValue)
            vValue = .cal05_User
            GetActivitiesRS.setField lCount, "CreatedBy", vValue
            vValue = "Type"
            GetActivitiesRS.setField lCount, "Type", vValue
            vValue = .cal04_Action
            GetActivitiesRS.setField lCount, "Description", vValue
            vValue = .cal06_Comments
            GetActivitiesRS.setField lCount, "Comment", vValue
        End With
    Next
        
End Function


Public Function GetVehicleDetailRS(ByVal oLossReport As V2ECKeyBoard.clsCarLR) As WDDXRecordset
    
    Dim MyXML01LossReport As XMLTypes.XML01LossReport
    Dim MyudtXML01Loss As XMLTypes.udtXML01Loss
    Dim MyudtXML01AdminLossInfoApd As XMLTypes.udtXML01AdminLossInfoApd
    Dim vValue As Variant 'Use Variant to Get Variant Equiv Data Type.
    Dim sAddress As String
    Dim sZip As String
    Dim sState As String
    Dim sCity As String
    Dim sStreet As String
    
    MyXML01LossReport = oLossReport.LossReport
    MyudtXML01Loss = MyXML01LossReport.XML01Loss
    
    'Create a WDDX RS
        Set GetVehicleDetailRS = New WDDXRecordset
        GetVehicleDetailRS.addColumn "VehicleMake"
        GetVehicleDetailRS.addColumn "VehicleModel"
        GetVehicleDetailRS.addColumn "VehicleYear"
        GetVehicleDetailRS.addColumn "VIN"
        GetVehicleDetailRS.addColumn "PropertyItemName"
        GetVehicleDetailRS.addColumn "DamageDescription"
        GetVehicleDetailRS.addColumn "LossDescription"
        GetVehicleDetailRS.addColumn "LocationType"
        GetVehicleDetailRS.addColumn "LocationName"
        GetVehicleDetailRS.addColumn "LocationPhoneNumber"
        GetVehicleDetailRS.addColumn "LocationAddress"
        GetVehicleDetailRS.addColumn "LocationCity"
        GetVehicleDetailRS.addColumn "LocationState"
        GetVehicleDetailRS.addColumn "LocationPostalCode"
        
        'Only one row for the Data RS
        GetVehicleDetailRS.addRows 1
        
        'Set the Col values for this one row
        MyudtXML01AdminLossInfoApd = MyudtXML01Loss.AdminLossInfoApd
        With MyudtXML01AdminLossInfoApd
        vValue = .ali0067_VehicleDescription
        GetVehicleDetailRS.setField 1, "VehicleMake", vValue
        vValue = .ali0067_VehicleDescription
        GetVehicleDetailRS.setField 1, "VehicleModel", vValue
        vValue = .ali0067_VehicleDescription
        GetVehicleDetailRS.setField 1, "VehicleYear", vValue
        vValue = .ali0069_VIN
        GetVehicleDetailRS.setField 1, "VIN", vValue
        vValue = "PropertyItemName"
        GetVehicleDetailRS.setField 1, "PropertyItemName", vValue
        vValue = "DamageDescription"
        GetVehicleDetailRS.setField 1, "DamageDescription", vValue
        vValue = "LossDescription"
        GetVehicleDetailRS.setField 1, "LossDescription", vValue
        vValue = "LocationType"
        GetVehicleDetailRS.setField 1, "LocationType", vValue
        vValue = "LocationName"
        GetVehicleDetailRS.setField 1, "LocationName", vValue
        vValue = "LocationPhoneNumber"
        GetVehicleDetailRS.setField 1, "LocationPhoneNumber", vValue
        
        sAddress = .ali0065_MailAddress1 & String(2, Chr(32)) & IIf(.ali0066_MailAddress2 = vbNullString, vbNullString, S_z & .ali0066_MailAddress2)
        goUtil.utFillAddressFields sAddress, sZip, sState, sCity, sStreet
        goUtil.utUpdateAddress sAddress, sZip, sState, sCity, sStreet
        vValue = "LocationAddress"
        GetVehicleDetailRS.setField 1, "LocationAddress", sAddress
        vValue = ""
        GetVehicleDetailRS.setField 1, "LocationCity", sCity
        vValue = ""
        GetVehicleDetailRS.setField 1, "LocationState", sState
        vValue = ""
        GetVehicleDetailRS.setField 1, "LocationPostalCode", sZip
        
        End With
End Function

Public Function GetLossDetailRS(ByVal oLossReport As V2ECKeyBoard.clsCarLR) As WDDXRecordset
    Dim MyXML01LossReport As XMLTypes.XML01LossReport
    Dim MyudtXML01Loss As XMLTypes.udtXML01Loss
    Dim MyudtXML01AdminLossInfo As XMLTypes.udtXML01AdminLossInfo
    Dim vValue As Variant 'Use Variant to Get Variant Equiv Data Type.
    Dim sAddress As String
    Dim sZip As String
    Dim sState As String
    Dim sCity As String
    Dim sStreet As String
    
    MyXML01LossReport = oLossReport.LossReport
    MyudtXML01Loss = MyXML01LossReport.XML01Loss
    
    'Create a WDDX RS
        Set GetLossDetailRS = New WDDXRecordset
        GetLossDetailRS.addColumn "LossLocationAddress"
        GetLossDetailRS.addColumn "LossLocationAddress2"
        GetLossDetailRS.addColumn "LossLocationCity"
        GetLossDetailRS.addColumn "LossLocationState"
        GetLossDetailRS.addColumn "LossLocationZip"
        GetLossDetailRS.addColumn "PropertyAddress"
        GetLossDetailRS.addColumn "PropertyCity"
        GetLossDetailRS.addColumn "PropertyState"
        GetLossDetailRS.addColumn "PropertyZip"
        GetLossDetailRS.addColumn "AffectedAreas"
        GetLossDetailRS.addColumn "LossDescription"
        
        'Only one row for the Data RS
        GetLossDetailRS.addRows 1

        'Set the Col values for this one row
        MyudtXML01AdminLossInfo = MyudtXML01Loss.AdminLossInfo
        With MyudtXML01AdminLossInfo
            sAddress = .ali0080_LossLocAddress1 & String(2, Chr(32)) & IIf(.ali0081_LossLocAddress2 = vbNullString, vbNullString, S_z & .ali0081_LossLocAddress2)
            goUtil.utFillAddressFields sAddress, sZip, sState, sCity, sStreet
            goUtil.utUpdateAddress sAddress, sZip, sState, sCity, sStreet
            GetLossDetailRS.setField 1, "LossLocationAddress", sStreet
            GetLossDetailRS.setField 1, "LossLocationAddress2", vbNullString
            GetLossDetailRS.setField 1, "LossLocationCity", sCity
            GetLossDetailRS.setField 1, "LossLocationState", sState
            GetLossDetailRS.setField 1, "LossLocationZip", sZip
            sAddress = .ali0063_MailAddress1 & String(2, Chr(32)) & IIf(.ali0064_MailAddress2 = vbNullString, vbNullString, S_z & .ali0064_MailAddress2)
            goUtil.utFillAddressFields sAddress, sZip, sState, sCity, sStreet
            goUtil.utUpdateAddress sAddress, sZip, sState, sCity, sStreet
            GetLossDetailRS.setField 1, "PropertyAddress", sStreet
'            GetLossDetailRS.setField 1, "PropertyAddress2", vbNullString
            GetLossDetailRS.setField 1, "PropertyCity", sCity
            GetLossDetailRS.setField 1, "PropertyState", sState
            GetLossDetailRS.setField 1, "PropertyZip", sZip
            vValue = "Test Affected Areas"
            GetLossDetailRS.setField 1, "AffectedAreas", vValue
            vValue = "Test Loss Description"
            GetLossDetailRS.setField 1, "LossDescription", vValue
        End With
End Function


Public Function GetPolicyDetailRS(ByVal oLossReport As V2ECKeyBoard.clsCarLR) As WDDXRecordset
    Dim MyXML01LossReport As XMLTypes.XML01LossReport
    Dim MyudtXML01Loss As XMLTypes.udtXML01Loss
    Dim MyudtXML01AdminLossInfo As XMLTypes.udtXML01AdminLossInfo
    Dim vValue As Variant 'Use Variant to Get Variant Equiv Data Type.
    Dim sAddress As String
    Dim sZip As String
    Dim sState As String
    Dim sCity As String
    Dim sStreet As String
    
    MyXML01LossReport = oLossReport.LossReport
    MyudtXML01Loss = MyXML01LossReport.XML01Loss
    
    'Create a WDDX RS
    Set GetPolicyDetailRS = New WDDXRecordset
    GetPolicyDetailRS.addColumn "PolicyNumber"
    GetPolicyDetailRS.addColumn "Status"
    GetPolicyDetailRS.addColumn "CoverageStatus"
    GetPolicyDetailRS.addColumn "BalanceDue"
    GetPolicyDetailRS.addColumn "CompanyCode"
    GetPolicyDetailRS.addColumn "CompanyName"
    GetPolicyDetailRS.addColumn "RenewalDate"
    GetPolicyDetailRS.addColumn "CancellationDate"
    GetPolicyDetailRS.addColumn "NewBusinessDate"
    GetPolicyDetailRS.addColumn "PolicyDescription"
    GetPolicyDetailRS.addColumn "PolicyEdition"
    GetPolicyDetailRS.addColumn "MortgageeName"
    GetPolicyDetailRS.addColumn "MortgageeAddress"
    GetPolicyDetailRS.addColumn "LienHolderName"
    GetPolicyDetailRS.addColumn "LienHolderAddress"
    'Only one row for the Data RS
    GetPolicyDetailRS.addRows 1

    'Set the Col values for this one row
    
    If oLossReport.LossType = TypeXML01.XML01Apd Then
        With MyudtXML01Loss.AdminLossInfoApd
            vValue = .ali0051_PolicyNum
            GetPolicyDetailRS.setField 1, "PolicyNumber", vValue
             vValue = "Status"
            GetPolicyDetailRS.setField 1, "Status", vValue
             vValue = "CoverageStatus"
            GetPolicyDetailRS.setField 1, "CoverageStatus", vValue
             vValue = "BalanceDue"
            GetPolicyDetailRS.setField 1, "BalanceDue", vValue
             vValue = .ali0058_CompCode
            GetPolicyDetailRS.setField 1, "CompanyCode", vValue
             vValue = "CompanyName"
            GetPolicyDetailRS.setField 1, "CompanyName", vValue
             vValue = .ali0061_RenewalDate
            GetPolicyDetailRS.setField 1, "RenewalDate", vValue
             vValue = .ali0062_LastCancDate
            GetPolicyDetailRS.setField 1, "CancellationDate", vValue
             vValue = .ali0060_NewBusDate
            GetPolicyDetailRS.setField 1, "NewBusinessDate", vValue
             vValue = .ali0059_PolicyType
            GetPolicyDetailRS.setField 1, "PolicyDescription", vValue
            vValue = "PolicyEdition"
            GetPolicyDetailRS.setField 1, "PolicyEdition", vValue
            vValue = .ali0057_MortgageHolder
            GetPolicyDetailRS.setField 1, "LienHolderName", vValue
            vValue = "LienHolderAddress"
            GetPolicyDetailRS.setField 1, "LienHolderAddress", vValue
        End With
    ElseIf oLossReport.LossType = TypeXML01.XML01Pro Then
        With MyudtXML01Loss.AdminLossInfo
            vValue = .ali0054_PolicyNum
            GetPolicyDetailRS.setField 1, "PolicyNumber", vValue
             vValue = "Status"
            GetPolicyDetailRS.setField 1, "Status", vValue
             vValue = "CoverageStatus"
            GetPolicyDetailRS.setField 1, "CoverageStatus", vValue
             vValue = "BalanceDue"
            GetPolicyDetailRS.setField 1, "BalanceDue", vValue
             vValue = .ali0068_CompCode
            GetPolicyDetailRS.setField 1, "CompanyCode", vValue
             vValue = "CompanyName"
            GetPolicyDetailRS.setField 1, "CompanyName", vValue
             vValue = .ali0060_RenewalDate
            GetPolicyDetailRS.setField 1, "RenewalDate", vValue
             vValue = .ali0061_LastCancDate
            GetPolicyDetailRS.setField 1, "CancellationDate", vValue
             vValue = .ali0059_NewBusDate
            GetPolicyDetailRS.setField 1, "NewBusinessDate", vValue
             vValue = .ali0069_PolicyDescription
            GetPolicyDetailRS.setField 1, "PolicyDescription", vValue
            vValue = .ali0066_MortgageHolder
            GetPolicyDetailRS.setField 1, "MortgageeName", vValue
            vValue = "MortgageeAddress"
            GetPolicyDetailRS.setField 1, "MortgageeAddress", vValue
        End With
    End If
             
End Function

Public Function CheckForNullDate(pvValue As Variant) As Date
    If IsDate(pvValue) Then
        CheckForNullDate = CDate(pvValue)
    Else
        CheckForNullDate = NULL_DATE
    End If
End Function

Public Function CheckForNullCurrency(pvValue As Variant) As Currency
    If IsNumeric(pvValue) Then
        CheckForNullCurrency = CCur(pvValue)
    Else
        CheckForNullCurrency = 0
    End If
End Function

