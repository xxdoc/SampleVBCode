VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLossASN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Loss Report Polymorphism
'Loss Report InterFace Definiton
'Use Implements V2ECKeyBoard.clsCarLR in ALL Loss Report Classes For All Carrier Objects

Implements V2ECKeyBoard.clsCarLR

'Admin udt for ASN format
Public Type udtASNAdmin
    DateAsgn As String
    CRAsgn As String
    SALN As String
    Company As String
    TypeOfLoss As String
    sT As String
    PolicyNumber As String
    TexasSuffix As String
    CatCD As String
    PolicyDesription As String
    Insured As String
    ADDRESS As String
    HPhone As String
    WPhone As String
    WEXT As String
    PropertyLocation As String
    MTGName As String
    LoanNo As String
    MTGAddress As String
    MTGCode As String
    AgentLR As String
    State As String
    District As String
    Agent As String
    ReportedBy As String
    ReportedByPhone As String
    DateRPTDToAgent As String
    DateRPTDByAgent As String
    LossDate As String
    LossLocationType As String
End Type

'Policy udt for ASN format
Public Type udtASNPolicy
    Insured As String
    BalanceDue As String
    RenewalDate As String
    MFRec As String
    NewBusReinDT As String
    LastCancelDate As String
    BuildingLim As String
    ContentsLim As String
    Deductible As String
End Type

Public Type udtASNLoss
    AdminInfo As udtASNAdmin
    PolicyInfo As udtASNPolicy
End Type

Public Type ASNLossReport
    ASNLoss As udtASNLoss
    PrnData As String
    PrnKey As String
End Type

Public Enum TypeASN
    ASNApd = 0
    ASNPro
End Enum

Private mudtASNLossReport As ASNLossReport
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
Private mLossType As TypeASN
Private msOleType As String
Private mbAbortProcessRawData As Boolean


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
    Dim sAddress As String
    Dim sStreet As String
    Dim sCity As String
    Dim sState As String
    Dim sZip As String
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
        Case "EBERLS"
            Select Case plVersion
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
        mLossType = ASNPro
    ElseIf CStr(pType) = "Auto" Then
        mLossType = ASNApd
    Else
        mLossType = pType
    End If
End Property
Public Property Get clsCarLR_LossType() As Variant
    clsCarLR_LossType = mLossType
End Property

Public Property Get clsCarLR_LossReport() As Variant
    clsCarLR_LossReport = mudtASNLossReport
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
    clsCarLR_MaxAllowedInChain = 300
End Property

Private Property Let clsCarLR_OleType(psOleType As String)
    msOleType = psOleType
End Property
Private Property Get clsCarLR_OleType() As String
    clsCarLR_OleType = msOleType
End Property

Public Property Let clsCarLR_PrnData(psData As String)
    mudtASNLossReport.PrnData = psData
End Property
Public Property Get clsCarLR_PrnData() As String
    clsCarLR_PrnData = mudtASNLossReport.PrnData
End Property

Public Property Let clsCarLR_PrnKey(psKey As String)
    mudtASNLossReport.PrnKey = psKey
End Property
Public Property Get clsCarLR_PrnKey() As String
    clsCarLR_PrnKey = mudtASNLossReport.PrnKey
End Property

Public Property Get clsCarLR_ClassName() As String
    clsCarLR_ClassName = App.EXEName & ".clsLossASN"
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
    'ASN we will be processing 1 file to 1 Loss report. So all we have to do
    'Is loop though all the files in the Directory and extract the info we need
    'Then save to proper out path and add to the collection of Raw loss report objects
    On Error GoTo EH
    Dim otxtProgMess As TextBox
    Dim oLR As V2ECKeyBoard.clsCarLR
    Dim colFiles As Collection
    Dim sFile As String
    Dim vFile As Variant
    Dim sData As String
    Dim sPrnData As String 'Will hold each loss PrnData
    Dim varyData As Variant
    Dim sCRID As String
    Dim sSALN As String 'ClaimNo
    Dim sFact As String 'FACT
    Dim sKey As String
    Dim sRawFileName As String
    Dim bPB As Boolean 'True if using it
    'For Missing Adjusters
    Dim MyAdjuster As udtAdjuster
    Dim bMissing As Boolean
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    If Not goUtil.utFileExists(psDataPath, True) Then
        Exit Function 'BAIL
    Else
        'check for Backup folder
        If Not goUtil.utFileExists(psDataPath & "\BackUp", True) Then
            MkDir psDataPath & "\BackUp"
        End If
    End If
    
    'Set ProgBar flag
    If Not poPB Is Nothing Then
        bPB = True
    End If
    'Set Progress text Mess
    If TypeOf potxtProgMess Is TextBox Then
        Set otxtProgMess = potxtProgMess
    End If
    
    'Load ASN format files
    sFile = Dir(psDataPath & "\" & "*.ASN")
    
    If sFile <> vbNullString Then
        Set colFiles = New Collection
        Do
            colFiles.Add psDataPath & "\" & sFile
            sFile = Dir
        Loop Until sFile = vbNullString
    Else
        Exit Function 'BAIL
    End If
    
    'Progress bar
    If bPB Then
        poPB.Max = colFiles.Count
        poPB.Value = 0
        If Not otxtProgMess Is Nothing Then
            otxtProgMess.Visible = True
            otxtProgMess.Text = poPB.Value & " Of " & poPB.Max
            otxtProgMess.Refresh
        End If
    End If
    
    'Loop through the collection of Files
    For Each vFile In colFiles
        If goUtil.utFileExists(CStr(vFile)) Then
        
            'Get ASN data
            sData = goUtil.utGetFileData(CStr(vFile))
            
            'Skip header data that comes before the first CRLF
            sData = Mid(sData, InStr(1, sData, vbCrLf))
            
            'Clean out any Forced Page breaks
            sData = Replace(sData, zEOP, vbNullString, , , vbBinaryCompare)
            
            'To Parse out the CRID and SALAN, split into array using CRLF delim
            varyData = Split(sData, vbCrLf)
            
            
            
            'FAX ID FACT
            sFact = RTrim(Mid(varyData(1), InStrRev(varyData(1), "CR ASSIGNED: ") + Len("CR ASSIGNED: ")))
            
            sData = vbNullString 'Clear it
            
            'SALN (ClaimNumber)
            'It will also be in this footer data
            sSALN = varyData(UBound(varyData, 1))
            'SALN is the only thing immediately followed by 'A  0101'
            'Go back 14 chars and trim spaces out.
            sSALN = Trim(Mid(sSALN, InStr(1, sSALN, "A  0101") - 14, 14))
            'Replace any null chars with null strings
            sSALN = Replace(sSALN, vbNullChar, vbNullString, , , vbBinaryCompare)
            'Replace the space in the SALN with "-"
            sSALN = Replace(sSALN, " ", vbNullString, , , vbTextCompare)
            
            msCLIENTNUM = sSALN
            
            'Format Report Data
            'Get rid of the Footer Data, don't need it anymore
            varyData(UBound(varyData, 1)) = vbNullString
            'Add the class name to the header line which will be "clsLossASN"
            varyData(LBound(varyData, 1)) = clsCarLR_ClassName
            
            'Put the modified report data back together
            sData = Join(varyData, vbCrLf)
            
            'Make a Backup copy of the Raw File
            'Check to see if the Raw File is already in the
            'BackUp folder first
            sFile = CStr(vFile)
            sFile = Mid(sFile, InStrRev(sFile, "\") + 1)
            sRawFileName = psDataPath & "\BackUp\" & sFile
            If goUtil.utFileExists(sRawFileName) Then
                SetAttr CStr(sRawFileName), vbNormal
                Kill CStr(sRawFileName)
            End If
            FileCopy CStr(vFile), sRawFileName
            'Get rid of the original Raw File since it is copied into
            'the backup folder
            SetAttr CStr(vFile), vbNormal
            Kill CStr(vFile)
            
            'Need to create Loss Report Object and add it to
            'The Raw Loss Report collection
            Set oLR = New clsLossASN
            With oLR
                Set .LRs = moLRs
                'Remove the class marker frm the prn data
                sPrnData = sData
                sPrnData = Replace(sPrnData, clsCarLR_ClassName & vbCrLf, vbNullString, 1)
                .PrnData = sPrnData
                .PopulateMe
            End With
            
            moLRs.AddRawLossReport oLR
            'Check to see if Process of Raw Data was Aborted...
            If oLR.AbortProcessRawData Then
                GoTo CLEAN_UP
            End If
            
            'Free loacal memory
            Set oLR = Nothing
        End If
        If bPB Then
            poPB.Value = poPB.Value + 1
            If Not otxtProgMess Is Nothing Then
                otxtProgMess.Text = poPB.Value & " Of " & poPB.Max
                otxtProgMess.Refresh
            End If
        End If
    Next
CLEAN_UP:
    'CleanUp
    Set otxtProgMess = Nothing
    Set colFiles = Nothing
    Set oLR = Nothing
    
    clsCarLR_ProcessRawData = True
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Set colFiles = Nothing
    Set oLR = Nothing
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLR_ClassName & vbCrLf & "Public Function ProcessRawData"
End Function

Public Function clsCarLR_AdditmX(poitmX As Object, polvw As Object) As Boolean
    On Error GoTo EH
    
    Set poitmX = polvw.ListItems.Add(, , Format(msDateAssign, "MM/DD/YY"), , LRPic.lrFORMAT01)
    
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
            'Very important to put EOP End of Page characer ASCII 12 to force Page break.
            'If you don't put EOP the print jobs will run into the next and mess up
            'The printer spooler buffer even if you use EndPagePrinter spool command.
            'Actually if your printer is smart eneough it will honor EndPagePrinter but not all
            'printers are so smart.
            moLRs.PrintToSpool clsCarLR_PrnKey, clsCarLR_PrnData & zEOP
        'Translation we will be using Active Reports
        Case PrintFormat.Translated
            'Use the General text function in LossReports Class
            moLRs.PrintGenText Me, psPrnDeviceName, pbPreview, pbPDF, plHwnd, plChainType
            
    End Select
    
    'We will print Appending documents After the Loss report
    'if they are selected
    If moLRs.AppDocFlag And plChainType = ChainType.NotChain And Not pbPreview And Not pbPDF Then
        For Each oChk In pofrmIncludeDocs
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
            moLRs.FirePrintAppDoc mudtASNLossReport.ASNLoss.AdminInfo.SALN
        End If
    End If
    clsCarLR_PrintMe = True
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLR_ClassName & vbCrLf & "Public Function PrintMe"
End Function

Public Function clsCarLR_ExportME(poLRs As V2ECKeyBoard.clsLossReports, pXportType As ExportType) As Boolean
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    If poLRs.ExportGenText(Me, pXportType) Then
        clsCarLR_ExportME = True
    End If
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLR_ClassName & vbCrLf & "Public Function ExportME"
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
    'Here is where we populate the Main udt for the ASN Format
    On Error GoTo EH
    Dim vLoss As Variant
    Dim lcount As Long
    Dim bPCI As Boolean 'true if reached Policy Coverage information
    Dim bPAY As Boolean 'True if reached Payment informtion
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    'Type Of ASN (ONLY Property at this time)
    clsCarLR_LossType = TypeASN.ASNPro
    'For population purposes get rid of existing VBCRLF
    'and split into vLoss Array.
    vLoss = Split(Replace(Me.clsCarLR_PrnData, vbCrLf, vbNullString, , , vbBinaryCompare), zDM)
    
    For lcount = LBound(vLoss, 1) To UBound(vLoss, 1)
        
        '1. Populate Admin Info
        If Not bPCI Then
            PopulateALI vLoss, lcount, bPCI
        End If
    
        '2. Populate Policy Coverage Info
        If bPCI And Not bPAY Then
            PopulatePCI vLoss, lcount, bPAY
        End If
        
        If bPAY Then
            Exit For 'Bail
        End If
    Next
    
    'SET THE KEY for this Loss Report
    msInsuredName = mudtASNLossReport.ASNLoss.AdminInfo.Insured
    msWorkPhone = mudtASNLossReport.ASNLoss.AdminInfo.WPhone
    msHomePhone = mudtASNLossReport.ASNLoss.AdminInfo.HPhone
    msDateAssign = mudtASNLossReport.ASNLoss.AdminInfo.DateAsgn
    msStatus = "PENDING"
    msTypeOfACID = "FACT"
    msAdjuster = "*"
    msACID = mudtASNLossReport.ASNLoss.AdminInfo.CRAsgn
    msCLIENTNUM = mudtASNLossReport.ASNLoss.AdminInfo.SALN
    'If this is A file path leav it alone... For Export
    If InStr(1, mudtASNLossReport.PrnKey, ":\") = 0 Then
        mudtASNLossReport.PrnKey = "*ADJUSTERUSERNAME*" & "_" & msACID & "_" & "*IBNUMBER*" & "_" & msCLIENTNUM
    End If
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLR_ClassName & vbCrLf & "Private Sub PopulateMe"
End Sub

Private Function PopulateALI(pvLoss As Variant, plCount As Long, pbPCI As Boolean) As Boolean
    On Error GoTo EH
    Dim sName As String
    Dim sValue As String
    Dim sTemp As String 'Used for helping figure Value strings
    Dim lPos As Long    'Used for helping figure Value strings
    Dim lOffset As Long 'offset for plCount
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    If plCount + 1 > UBound(pvLoss, 1) Then
        Exit Function 'Bail
    End If
    
    'Check to see if we have Reached the Policy Coverage Info yet
    If InStr(1, pvLoss(plCount), zSTART_PCI, vbTextCompare) > 0 Then
        pbPCI = True
'        Debug.Print Me.PrnPath
'        Debug.Print START_PCI
        Exit Function
    End If
    
    'Use Switch to populate sName with the field name we want
    'using switch will minimize the amount of time it takes to locate Field name
    
    sName = Switch( _
    StrComp(Right(pvLoss(plCount), 13), "DATE ASSIGNED", vbTextCompare) = 0, "DATE ASSIGNED", StrComp(Right(pvLoss(plCount), 11), "CR ASSIGNED", vbTextCompare) = 0, "CR ASSIGNED", _
    StrComp(Right(pvLoss(plCount), 4), "SALN", vbTextCompare) = 0, "SALN", StrComp(Right(pvLoss(plCount), 7), "COMPANY", vbTextCompare) = 0, "COMPANY", _
    StrComp(Right(pvLoss(plCount), 12), "TYPE OF LOSS", vbTextCompare) = 0, "TYPE OF LOSS", StrComp(Right(pvLoss(plCount), 2), "ST", vbTextCompare) = 0, "ST", _
    StrComp(Right(pvLoss(plCount), 5), "POL.#", vbTextCompare) = 0, "POL.#", StrComp(Right(pvLoss(plCount), 6), "TX SUF", vbTextCompare) = 0, "TX SUF", _
    StrComp(Right(pvLoss(plCount), 6), "CAT CD", vbTextCompare) = 0, "CAT CD", StrComp(Right(pvLoss(plCount), 18), "POLICY DESCRIPTION", vbTextCompare) = 0, "POLICY DESCRIPTION", _
    StrComp(Right(pvLoss(plCount), 7), "INSURED", vbTextCompare) = 0, "INSURED", StrComp(Right(pvLoss(plCount), 12), "MTG. ADDRESS", vbTextCompare) = 0, "MTG. ADDRESS", _
    StrComp(Right(pvLoss(plCount), 10), "HOME PHONE", vbTextCompare) = 0, "HOME PHONE", StrComp(Right(pvLoss(plCount), 9), "BUS.PHONE", vbTextCompare) = 0, "BUS.PHONE", _
    StrComp(Right(pvLoss(plCount), 3), "EXT", vbTextCompare) = 0, "EXT", StrComp(Right(pvLoss(plCount), 17), "PROPERTY LOCATION", vbTextCompare) = 0, "PROPERTY LOCATION", _
    StrComp(Right(pvLoss(plCount), 9), "MTG. NAME", vbTextCompare) = 0, "MTG. NAME", StrComp(Right(pvLoss(plCount), 8), "LOAN NO.", vbTextCompare) = 0, "LOAN NO.", _
    StrComp(Right(pvLoss(plCount), 7), "ADDRESS", vbTextCompare) = 0, "ADDRESS", StrComp(Right(pvLoss(plCount), 9), "MTG. CODE", vbTextCompare) = 0, "MTG. CODE", _
    False = False, vbNullString)
    
    If sName <> vbNullString Then
        GoTo SET_VALUE
    End If
    
    sName = Switch( _
    StrComp(Right(pvLoss(plCount), 18), "DATE RPTD TO AGENT", vbTextCompare) = 0, "DATE RPTD TO AGENT", StrComp(Right(pvLoss(plCount), 17), "DATE RPTD BY AGNT", vbTextCompare) = 0, "DATE RPTD BY AGNT", _
    StrComp(Right(pvLoss(plCount), 8), "DISTRICT", vbTextCompare) = 0, "DISTRICT", StrComp(Right(pvLoss(plCount), 5), "AGENT", vbTextCompare) = 0, "AGENT", _
    StrComp(Right(pvLoss(plCount), 11), "REPORTED BY", vbTextCompare) = 0, "REPORTED BY", StrComp(Right(pvLoss(plCount), 5), "PHONE", vbTextCompare) = 0, "PHONE", _
    StrComp(Right(pvLoss(plCount), 8), "AGENT LR", vbTextCompare) = 0, "AGENT LR", StrComp(Right(pvLoss(plCount), 5), "STATE", vbTextCompare) = 0, "STATE", _
    StrComp(Right(pvLoss(plCount), 16), "LOSS DATE / TIME", vbTextCompare) = 0, "LOSS DATE / TIME", StrComp(Right(pvLoss(plCount), 13), "LOSS LOCATION", vbTextCompare) = 0, "LOSS LOCATION", _
    False = False, vbNullString)
SET_VALUE:

    lOffset = 1 'Since we are not skipping any elements the offset will always be +1
    
    sValue = pvLoss(plCount + lOffset)

    'When setting the Value be sure to remove the Field label of the
    'next data element
    With mudtASNLossReport.ASNLoss.AdminInfo
        Select Case sName
            Case "DATE ASSIGNED"
                lPos = InStr(1, sValue, "CR ASSIGNED", vbTextCompare) - 1
                sValue = Trim(left(sValue, lPos))
                .DateAsgn = sValue
            Case "CR ASSIGNED"
                lPos = InStr(1, sValue, "SALN", vbTextCompare) - 1
                sValue = Trim(left(sValue, lPos))
                .CRAsgn = sValue
            Case "SALN"
                lPos = InStr(1, sValue, "COMPANY", vbTextCompare) - 1
                sValue = Trim(left(sValue, lPos))
                sValue = Replace(sValue, " ", "-") '-Put dash in the SALN instead of space
                .SALN = goUtil.utScrubAlphaNumeric(sValue)
            Case "COMPANY"
                lPos = InStr(1, sValue, "TYPE OF LOSS", vbTextCompare) - 1
                sValue = Trim(left(sValue, lPos))
                .Company = sValue
            Case "TYPE OF LOSS"
                lPos = InStr(1, sValue, "ST", vbTextCompare) - 1
                sValue = Trim(left(sValue, lPos))
                .TypeOfLoss = sValue
            Case "ST"
                lPos = InStr(1, sValue, "POL.#", vbTextCompare) - 1
                sValue = Trim(left(sValue, lPos))
                .sT = sValue
            Case "POL.#"
                lPos = InStr(1, sValue, "TX SUF", vbTextCompare) - 1
                sValue = Trim(left(sValue, lPos))
                .PolicyNumber = sValue
            Case "TX SUF"
                lPos = InStr(1, sValue, "CAT CD", vbTextCompare) - 1
                sValue = Trim(left(sValue, lPos))
                .TexasSuffix = sValue
            Case "CAT CD"
                lPos = InStr(1, sValue, "POLICY DESCRIPTION", vbTextCompare) - 1
                sValue = Trim(left(sValue, lPos))
                .CatCD = sValue
            Case "POLICY DESCRIPTION"
                lPos = InStr(1, sValue, "INSURED", vbTextCompare) - 1
                sValue = Trim(left(sValue, lPos))
                .PolicyDesription = sValue
            Case "INSURED"
                lPos = InStr(1, sValue, "ADDRESS", vbTextCompare) - 1
                If lPos = -1 Then
                    lPos = InStr(1, sValue, "   ", vbTextCompare) + 1
                End If
                sValue = Trim(left(sValue, lPos))
                .Insured = sValue
            Case "ADDRESS"
                lPos = InStr(1, sValue, "HOME PHONE", vbTextCompare) - 1
                sValue = Trim(left(sValue, lPos))
                'Need to see if the CITY State zip is reapeated in address
                'This is crudy data we are cleaning up
                lPos = InStrRev(sValue, String(10, Chr(32)))
                If lPos > 0 Then
                    sTemp = Trim(Mid(sValue, lPos))
                    If InStr(1, left(sValue, lPos), sTemp, vbTextCompare) > 0 Then
                        sValue = RTrim(left(sValue, lPos))
                    End If
                End If
                
                .ADDRESS = sValue
            Case "HOME PHONE"
                lPos = InStr(1, sValue, "BUS.PHONE", vbTextCompare) - 1
                sValue = Trim(left(sValue, lPos))
                .HPhone = sValue
            Case "BUS.PHONE"
                lPos = InStr(1, sValue, "EXT", vbTextCompare) - 1
                sValue = Trim(left(sValue, lPos))
                .WPhone = sValue
            Case "EXT"
                lPos = InStr(1, sValue, "PROPERTY LOCATION", vbTextCompare) - 1
                sValue = Trim(left(sValue, lPos))
                .WEXT = sValue
            Case "PROPERTY LOCATION"
                lPos = InStr(1, sValue, "MTG. NAME", vbTextCompare) - 1
                sValue = Trim(left(sValue, lPos))
                'Need to see if the CITY State zip is reapeated in address
                'This is crudy data we are cleaning up
                lPos = InStrRev(sValue, String(10, Chr(32)))
                If lPos > 0 Then
                    sTemp = Trim(Mid(sValue, lPos))
                    If InStr(1, left(sValue, lPos), sTemp, vbTextCompare) > 0 Then
                        sValue = RTrim(left(sValue, lPos))
                    End If
                End If
                .PropertyLocation = sValue
            Case "MTG. NAME"
                lPos = InStr(1, sValue, "LOAN NO.", vbTextCompare) - 1
                sValue = Trim(left(sValue, lPos))
                .MTGName = sValue
            Case "LOAN NO."
                lPos = InStr(1, sValue, "MTG. ADDRESS", vbTextCompare) - 1
                sValue = Trim(left(sValue, lPos))
                .LoanNo = sValue
            Case "MTG. ADDRESS"
                lPos = InStr(1, sValue, "MTG. CODE", vbTextCompare) - 1
                sValue = Trim(left(sValue, lPos))
                .MTGAddress = sValue
            Case "MTG. CODE"
                lPos = InStr(1, sValue, "AGENT LR", vbTextCompare) - 1
                sTemp = Trim(left(sValue, lPos))
                sValue = Trim(left(sTemp, 15))
                .MTGCode = sValue
                'MTG. CODE may have the City State ZIP trailing it
                If Len(sTemp) > 16 Then
                    sTemp = Trim(Mid(sTemp, 16))
                    .MTGAddress = .MTGAddress & sTemp
                End If
            Case "AGENT LR"
                lPos = InStr(1, sValue, "STATE", vbTextCompare) - 1
                sValue = Trim(left(sValue, lPos))
                .AgentLR = sValue
            Case "STATE"
                lPos = InStr(1, sValue, "DISTRICT", vbTextCompare) - 1
                sValue = Trim(left(sValue, lPos))
                .State = sValue
            Case "DISTRICT"
                lPos = InStr(1, sValue, "AGENT", vbTextCompare) - 1
                sValue = Trim(left(sValue, lPos))
                .District = sValue
            Case "AGENT"
                lPos = InStr(1, sValue, "REPORTED BY", vbTextCompare) - 1
                sValue = Trim(left(sValue, lPos))
                .Agent = sValue
            Case "REPORTED BY"
                lPos = InStr(1, sValue, "PHONE", vbTextCompare) - 1
                sValue = Trim(left(sValue, lPos))
                .ReportedBy = sValue
            Case "PHONE"
                lPos = InStr(1, sValue, "DATE RPTD TO AGENT", vbTextCompare) - 1
                sValue = Trim(left(sValue, lPos))
                .ReportedByPhone = sValue
            Case "DATE RPTD TO AGENT"
                lPos = InStr(1, sValue, "DATE RPTD BY AGNT", vbTextCompare) - 1
                sValue = Trim(left(sValue, lPos))
                .DateRPTDToAgent = sValue
            Case "DATE RPTD BY AGNT"
                lPos = InStr(1, sValue, "LOSS DATE / TIME", vbTextCompare) - 1
                sValue = Trim(left(sValue, lPos))
                .DateRPTDByAgent = sValue
            Case "LOSS DATE / TIME"
                lPos = InStr(1, sValue, "LOSS LOCATION", vbTextCompare) - 1
                sValue = Trim(left(sValue, lPos))
                .LossDate = sValue
            Case "LOSS LOCATION"
                lPos = InStr(1, sValue, "APPT DATE", vbTextCompare) - 1
                sValue = Trim(left(sValue, lPos))
                .LossLocationType = sValue
        End Select
    End With
    
    If sName <> vbNullString Then
'        Debug.Print sName & ": " & sValue
        PopulateALI = True
    End If
    

    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLR_ClassName & vbCrLf & "Private Function PopulateALI"
End Function

Private Function PopulatePCI(pvLoss As Variant, plCount As Long, pbPAY) As Boolean
    On Error GoTo EH
    Dim sName As String
    Dim sValue As String
    Dim sTemp As String 'Used for helping figure Value strings
    Dim lPos As Long
    Dim lOffset As Long 'Offset for plcount
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    If plCount + 1 > UBound(pvLoss, 1) Then
        Exit Function 'Bail
    End If
    
     'Check to see if we have Reached the Payment information  yet
    If InStr(1, pvLoss(plCount), zSTART_PAY, vbTextCompare) > 0 Then
        pbPAY = True
'        Debug.Print Me.PrnPath
'        Debug.Print START_PCI
        Exit Function
    End If
    
    'Use Switch to populate sName with the field name we want
    'using switch will minimize the amount of time it takes to locate Field name
    sName = Switch( _
    StrComp(Right(pvLoss(plCount), 7), "INSURED", vbTextCompare) = 0, "INSURED", StrComp(Right(pvLoss(plCount), 11), "BALANCE DUE", vbTextCompare) = 0, "BALANCE DUE", _
    StrComp(Right(pvLoss(plCount), 12), "RENEWAL DATE", vbTextCompare) = 0, "RENEWAL DATE", StrComp(Right(pvLoss(plCount), 7), "M/F REC", vbTextCompare) = 0, "M/F REC", _
    StrComp(Right(pvLoss(plCount), 15), "NEW BUS/REIN DT", vbTextCompare) = 0, "NEW BUS/REIN DT", StrComp(Right(pvLoss(plCount), 16), "LAST CANCEL DATE", vbTextCompare) = 0, "LAST CANCEL DATE", _
    StrComp(Right(pvLoss(plCount), 15), "BUILDING LIMITS", vbTextCompare) = 0, "BUILDING LIMITS", StrComp(Right(pvLoss(plCount), 21), "CONTENTS/EQUIP LIMITS", vbTextCompare) = 0, "CONTENTS/EQUIP LIMITS", _
    StrComp(Right(pvLoss(plCount), 11), "DEDUCTIBLES", vbTextCompare) = 0, "DEDUCTIBLES", _
    False = False, vbNullString)
            
SET_VALUE:

    lOffset = 1 'Since we are not skipping any elements the offset will always be +1
    
    sValue = pvLoss(plCount + lOffset)
    
    'When setting the Value be sure to remove the Field label of the
    'next data element
    With mudtASNLossReport.ASNLoss.PolicyInfo
        Select Case sName
            Case "INSURED"
                lPos = InStr(1, sValue, "BALANCE DUE", vbTextCompare) - 1
                sValue = Trim(left(sValue, lPos))
                .Insured = sValue
            Case "BALANCE DUE"
                lPos = InStr(1, sValue, "RENEWAL DATE", vbTextCompare) - 1
                sValue = Trim(left(sValue, lPos))
                .BalanceDue = sValue
            Case "RENEWAL DATE"
                lPos = InStr(1, sValue, "M/F REC", vbTextCompare) - 1
                sValue = Trim(left(sValue, lPos))
                .RenewalDate = sValue
            Case "M/F REC"
                lPos = InStr(1, sValue, "NEW BUS/REIN DT", vbTextCompare) - 1
                sValue = Trim(left(sValue, lPos))
                .MFRec = sValue
            Case "NEW BUS/REIN DT"
                lPos = InStr(1, sValue, "LAST CANCEL DATE", vbTextCompare) - 1
                sValue = Trim(left(sValue, lPos))
                .NewBusReinDT = sValue
            Case "LAST CANCEL DATE"
                lPos = InStr(1, sValue, "PRINT DATE", vbTextCompare) - 1
                sValue = Trim(left(sValue, lPos))
                .LastCancelDate = sValue
            Case "BUILDING LIMITS"
                lPos = InStr(1, sValue, "ENDORSEMENTS", vbTextCompare) - 1
                sValue = Trim(left(sValue, lPos))
                .BuildingLim = sValue
            Case "CONTENTS/EQUIP LIMITS"
                lPos = InStr(1, sValue, "CPL LIMITS", vbTextCompare) - 1
                sValue = Trim(left(sValue, 15)) 'Go 15 digits
                .ContentsLim = sValue
            Case "DEDUCTIBLES"
                lPos = InStr(1, sValue, zSTART_PAY, vbTextCompare) - 1
                sTemp = Trim(left(sValue, lPos))
                'Default Deductible to be first one
                If InStr(1, sTemp, zDED_C1, vbTextCompare) > 0 Then
                    sTemp = LTrim(Replace(sTemp, zDED_C1, vbNullString))
                    sTemp = RTrim(left(sTemp, 15)) 'Go 15 digits , that would be a high deductible !
                Else
                    sTemp = LTrim(Right(sTemp, 15))
                End If
                sValue = sTemp
                .Deductible = sValue
        End Select
    End With
    
    If sName <> vbNullString Then
'        Debug.Print sName & ": " & sValue
        PopulatePCI = True
    End If
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLR_ClassName & vbCrLf & "Private Function PopulatePCI"
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
    With mudtASNLossReport.ASNLoss
        'Use Policy Insured if there
        sSQL = sSQL & "VALUES (" & S_z & IIf(.PolicyInfo.Insured = vbNullString, .AdminInfo.Insured, .PolicyInfo.Insured) & z_S '01Client
        sSQL = sSQL & S_z & .AdminInfo.SALN & z_S                                                           '02ClaimNoSaln
        sSQL = sSQL & IIf(IsDate(.AdminInfo.DateAsgn), "#" & .AdminInfo.DateAsgn & "#", "Null") & ", "      '03DateAssigned
        sSQL = sSQL & S_z & .AdminInfo.CRAsgn & z_S                                                         '04Adjuster
        sSQL = sSQL & S_z & .AdminInfo.Company & z_S                                                        '05Company
        sSQL = sSQL & S_z & .AdminInfo.TypeOfLoss & z_S                                                     '06TypeOfLoss
        sSQL = sSQL & S_z & .AdminInfo.sT & z_S                                                             '07StateNo
        sSQL = sSQL & S_z & .AdminInfo.PolicyNumber & z_S                                                   '08PolicyNo
        sSQL = sSQL & S_z & .AdminInfo.TexasSuffix & z_S                                                    '09TexasSuffix
        sSQL = sSQL & S_z & .AdminInfo.CatCD & z_S                                                          '10CatCode
        sSQL = sSQL & S_z & .AdminInfo.PolicyDesription & z_S                                               '11PolicyDescription
        sSQL = sSQL & S_z & .AdminInfo.Insured & z_S                                                        '12Insured
        sSQL = sSQL & S_z & .AdminInfo.ADDRESS & z_S                                                        '13Address
        sSQL = sSQL & S_z & .AdminInfo.HPhone & z_S                                                         '14HomePhone
        sSQL = sSQL & S_z & .AdminInfo.WPhone & " EXT:" & .AdminInfo.WEXT & z_S                             '14aBusinessPhone
        sSQL = sSQL & S_z & .AdminInfo.PropertyLocation & z_S                                               '15PropertyAddress
        sSQL = sSQL & S_z & .AdminInfo.MTGName & z_S                                                        '16MortgageeName
        sSQL = sSQL & S_z & .AdminInfo.LoanNo & z_S                                                         '17LoanNo
        sSQL = sSQL & S_z & .AdminInfo.MTGAddress & z_S                                                     '18MtgAddress
        sSQL = sSQL & S_z & .AdminInfo.MTGCode & z_S                                                        '19MtgCode
        sSQL = sSQL & S_z & .AdminInfo.AgentLR & z_S                                                        '20AgentLR
        sSQL = sSQL & S_z & .AdminInfo.State & z_S                                                          '21StateCD
        sSQL = sSQL & S_z & .AdminInfo.District & z_S                                                       '22District
        sSQL = sSQL & S_z & .AdminInfo.Agent & z_S                                                          '23AgentNo
        sSQL = sSQL & S_z & .AdminInfo.ReportedBy & z_S                                                      '24ReportedBy
        sSQL = sSQL & S_z & .AdminInfo.ReportedByPhone & z_S                                                '25ReportedByPhone
        sSQL = sSQL & IIf(IsDate(.AdminInfo.DateRPTDToAgent), "#" & .AdminInfo.DateRPTDToAgent & "#", "Null") & ", " '26DateReportedToAgent
        sSQL = sSQL & IIf(IsDate(.AdminInfo.DateRPTDByAgent), "#" & .AdminInfo.DateRPTDByAgent & "#", "Null") & ", " '27DateReportedByAgent
        sSQL = sSQL & IIf(IsDate(.AdminInfo.LossDate), "#" & .AdminInfo.LossDate & "#", "Null") & ", "      '28LossDate
        sSQL = sSQL & S_z & .AdminInfo.LossLocationType & z_S                                               '29LossLocation
        sSQL = sSQL & S_z & .PolicyInfo.BalanceDue & z_S                                                    '30BalanceDue
        sSQL = sSQL & S_z & .PolicyInfo.MFRec & z_S                                                         '31MFRec
        sSQL = sSQL & IIf(IsDate(.PolicyInfo.RenewalDate), "#" & .PolicyInfo.RenewalDate & "#", "Null") & ", "   '32RenewalDate
        sSQL = sSQL & IIf(IsDate(.PolicyInfo.NewBusReinDT), "#" & .PolicyInfo.NewBusReinDT & "#", "Null") & ", " '33NewBusReinDt
        'Val function Bug in VB6
        'http://msdn.microsoft.com/vbasic/productinfo/previous/vb6/tips/01pasttips.asp
        'Need to parse out both % and !  because these trailing equate to Double and Single
        'and Val bugs because it can't convert Double or single into integer
        sTemp = .PolicyInfo.BuildingLim
        sSQL = sSQL & Val(goUtil.utCleanValString(sTemp)) & ", "                                                   '34BuildingLimits
        sTemp = .PolicyInfo.ContentsLim
        sSQL = sSQL & Val(goUtil.utCleanValString(sTemp)) & ", "                                                   '35ContentsLimits
        sSQL = sSQL & S_z & .PolicyInfo.Deductible & z_S                                                    '36Deductibles
        sSQL = sSQL & S_z & clsCarLR_ClassName & z_S                                                             'Format
        sSQL = sSQL & S_z & clsCarLR_ClassName & vbCrLf & clsCarLR_PrnData & S_z & " ) "                                '37LossReport
    End With
    
    SQL_WEBCONTROL_108 = sSQL
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    sErrDesc = mudtASNLossReport.ASNLoss.AdminInfo.SALN & vbCrLf & sErrDesc
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
        
    With mudtASNLossReport.ASNLoss
        sSQL = "z_spuInsertECSFarmersASNProperty "
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
        'Admin Info
        sSQL = sSQL & S_z & .AdminInfo.ADDRESS & z_S
        sSQL = sSQL & S_z & .AdminInfo.Agent & z_S
        sSQL = sSQL & S_z & .AdminInfo.AgentLR & z_S
        sSQL = sSQL & S_z & .AdminInfo.CatCD & z_S
        sSQL = sSQL & S_z & .AdminInfo.Company & z_S
        sSQL = sSQL & S_z & .AdminInfo.CRAsgn & z_S
        sSQL = sSQL & S_z & .AdminInfo.DateAsgn & z_S
        sSQL = sSQL & S_z & .AdminInfo.DateRPTDByAgent & z_S
        sSQL = sSQL & S_z & .AdminInfo.DateRPTDToAgent & z_S
        sSQL = sSQL & S_z & .AdminInfo.District & z_S
        sSQL = sSQL & S_z & .AdminInfo.HPhone & z_S
        sSQL = sSQL & S_z & .AdminInfo.Insured & z_S
        sSQL = sSQL & S_z & .AdminInfo.LoanNo & z_S
        sSQL = sSQL & S_z & .AdminInfo.LossDate & z_S
        sSQL = sSQL & S_z & .AdminInfo.LossLocationType & z_S
        sSQL = sSQL & S_z & .AdminInfo.MTGAddress & z_S
        sSQL = sSQL & S_z & .AdminInfo.MTGCode & z_S
        sSQL = sSQL & S_z & .AdminInfo.MTGName & z_S
        sSQL = sSQL & S_z & .AdminInfo.PolicyDesription & z_S
        sSQL = sSQL & S_z & .AdminInfo.PolicyNumber & z_S
        sSQL = sSQL & S_z & .AdminInfo.PropertyLocation & z_S
        sSQL = sSQL & S_z & .AdminInfo.ReportedBy & z_S
        sSQL = sSQL & S_z & .AdminInfo.ReportedByPhone & z_S
        sSQL = sSQL & S_z & .AdminInfo.SALN & z_S
        sSQL = sSQL & S_z & .AdminInfo.sT & z_S
        sSQL = sSQL & S_z & .AdminInfo.State & z_S
        sSQL = sSQL & S_z & .AdminInfo.TexasSuffix & z_S
        sSQL = sSQL & S_z & .AdminInfo.TypeOfLoss & z_S
        sSQL = sSQL & S_z & .AdminInfo.WEXT & z_S
        sSQL = sSQL & S_z & .AdminInfo.WPhone & z_S
        'Policy Info
        sTemp = .PolicyInfo.BalanceDue
        sSQL = sSQL & S_z & Val(goUtil.utCleanValString(sTemp)) & z_S
        sTemp = .PolicyInfo.BuildingLim
        sSQL = sSQL & S_z & Val(goUtil.utCleanValString(sTemp)) & z_S
        sTemp = .PolicyInfo.ContentsLim
        sSQL = sSQL & S_z & Val(goUtil.utCleanValString(sTemp)) & z_S
        sTemp = .PolicyInfo.Deductible
        sSQL = sSQL & S_z & Val(goUtil.utCleanValString(sTemp)) & z_S
        sSQL = sSQL & S_z & .PolicyInfo.Insured & z_S
        sSQL = sSQL & S_z & .PolicyInfo.LastCancelDate & z_S
        sSQL = sSQL & S_z & .PolicyInfo.MFRec & z_S
        sSQL = sSQL & S_z & .PolicyInfo.NewBusReinDT & z_S
        sSQL = sSQL & S_z & .PolicyInfo.RenewalDate & S_z
    End With
    
    sSQL = Replace(sSQL, Chr(39), Chr(39) & Chr(39))
    sSQL = Replace(sSQL, S_z, Chr(39))
    sSQL = Replace(sSQL, z_S, "', ")
    
    SQL_WEBCONTROL_20000 = sSQL
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    
    sErrDesc = mudtASNLossReport.ASNLoss.AdminInfo.SALN & vbCrLf & sErrDesc
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
    Dim sTemp As String
    Dim sTempCity As String
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
    Dim sCATCode As String
    Dim lNextID As Long 'Loss Reports Class will use parameter collection
    Dim lErrNum As Long
    Dim sErrDesc As String
   
    If moLRs.ParametersCol Is Nothing Then
        Exit Function
    End If
    'Get Parameters
    sCATCode = moLRs.FindParameter("CAT_CODE").ParamValue
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
    
    'IBNUM
    Param.ParamName = "IB_NUMBER"
    Param.ParamValue = sIBNum
    moLRs.AddParameter Param
    'Add Policy Desc, Build Lim, Cont Lim
    Param.ParamName = "POLICY_DESC"
    Param.ParamValue = mudtASNLossReport.ASNLoss.AdminInfo.PolicyDesription
    moLRs.AddParameter Param
    'Building Lim
    Param.ParamName = "BUILD_LIM"
    Param.ParamValue = mudtASNLossReport.ASNLoss.PolicyInfo.BuildingLim
    moLRs.AddParameter Param
    'Contents Lim
    Param.ParamName = "CONT_LIM"
    Param.ParamValue = mudtASNLossReport.ASNLoss.PolicyInfo.ContentsLim
    moLRs.AddParameter Param
    'Deductible
    Param.ParamName = "DEDUCTIBLE"
    Param.ParamValue = mudtASNLossReport.ASNLoss.PolicyInfo.Deductible
    moLRs.AddParameter Param
    
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
    
    
    With mudtASNLossReport.ASNLoss
    
        sSQL = sSQL & "VALUES (" & lNextID & ", "           'ID, "
        sSQL = sSQL & S_z & "Farmers" & z_S                 'Client, " = Faramers
        sSQL = sSQL & S_z & "NEW" & z_S                     '"Status, " = New
        sSQL = sSQL & S_z & sFName & z_S                    ' "AdjustorLastName, " = param
        sSQL = sSQL & S_z & sLName & z_S                    ' "AdjustorFirstName, " = param
        sSQL = sSQL & S_z & sCRID & z_S                     ' "ClaimRepIDNO, " =param
        sSQL = sSQL & True & ", "                           ' "Check1FieldHandeld, " =True
        sSQL = sSQL & True & ", "                           ' "Check1Final, " =True
        sSQL = sSQL & True & ", "                           ' "Check1ForProperty, " =True
        sSQL = sSQL & True & ", "                           ' "NCCToIssueExpenseCheck, " = True
        sSQL = sSQL & True & ", "                           ' "NCCFirstPayment, " =True
        sSQL = sSQL & S_z & sSiteAddress & z_S              ' "SiteAddress, " =Param
        sSQL = sSQL & S_z & sSiteCity & z_S                 ' "SiteCity, " = Param
        sSQL = sSQL & S_z & sSiteState & z_S                ' "SiteState, " =Param
        sSQL = sSQL & S_z & sSiteZip & z_S                  ' "SiteZip, " =Param
        sSQL = sSQL & S_z & .AdminInfo.PolicyNumber & z_S   ' "PolicyNumber, "
        sDate = .AdminInfo.DateAsgn
        If Not IsDate(sDate) Then
            sDate = Now()
        End If
        sSQL = sSQL & S_z & Format(CDate(sDate), "YYYYMMDD") & z_S                                  ' "AssignedDate, "
        sDate = Format(Now(), "YYYYMMDD")
        sSQL = sSQL & S_z & sDate & z_S                                                             ' "ReceivedDate, "
        sSQL = sSQL & S_z & .AdminInfo.SALN & z_S                                                   ' "ClientClaimNo, "
        sSQL = sSQL & S_z & .AdminInfo.TexasSuffix & z_S                                            ' "Check1TexasSuffix, "
        sSQL = sSQL & S_z & vbNullString & z_S                                                      ' "Check1TexasSubCoverageCode, "
        sSQL = sSQL & S_z & IIf(.AdminInfo.CatCD = vbNullString, sCATCode, .AdminInfo.CatCD) & z_S  ' "CatCode, "
        
        If Len(.AdminInfo.Insured) > Len(.PolicyInfo.Insured) Then
            sInsuredNames = .AdminInfo.Insured
        Else
            sInsuredNames = .PolicyInfo.Insured
        End If
        sSQL = sSQL & S_z & sInsuredNames & z_S             ' "NewInsuredNames, "
        
        'Mailing Address
        sMailAddress = .AdminInfo.ADDRESS  '1 & String(2, Chr(32)) & 2
        On Error Resume Next
        sTemp = LTrim(Mid(sMailAddress, InStrRev(sMailAddress, String(10, Chr(32)))))
        If Err.Number > 0 Then
            Err.Clear
            sTemp = vbNullString
        End If
        If sTemp <> vbNullString Then
            If InStr(1, sTemp, ",", vbBinaryCompare) > 0 Then
                sTempCity = left(sTemp, InStr(1, sTemp, ",", vbBinaryCompare) - 1)
                sTempCity = Replace(sTempCity, Chr(32), Chr(160))
                sTemp = Replace(sTemp, Replace(sTempCity, Chr(160), Chr(32)), sTempCity)
            End If
            sMailAddress = left(sMailAddress, InStr(1, sMailAddress, String(10, Chr(32))))
            sMailAddress = sMailAddress & Chr(32) & S_z & sTemp
        End If
        goUtil.utFillAddressFields sMailAddress, sMailZip, sMailState, sMailCity, sMailStreet
        goUtil.utUpdateAddress sMailAddress, sMailZip, sMailState, sMailCity, sMailStreet
        
        sSQL = sSQL & S_z & sMailAddress & z_S              ' "NewMailingAddress, "
        sSQL = sSQL & S_z & sMailStreet & z_S               ' "MailingAddress, "
        sSQL = sSQL & S_z & sMailZip & z_S                  ' "MailingZip, "
        sSQL = sSQL & S_z & sMailState & z_S                ' "MailingState, "
        sSQL = sSQL & S_z & sMailCity & z_S                 ' "MailingCity, "
        
        'Loss Location Address
        sLossAddress = .AdminInfo.PropertyLocation '.AdminLossInfo.ali0080_LossLocAddress1 & String(2, Chr(32)) & IIf(.AdminLossInfo.ali0081_LossLocAddress2 = vbNullString, vbNullString, S_z & .AdminLossInfo.ali0081_LossLocAddress2)
        On Error Resume Next
        sTemp = LTrim(Mid(sLossAddress, InStrRev(sLossAddress, String(10, Chr(32)))))
        If Err.Number > 0 Then
            Err.Clear
            sTemp = vbNullString
        End If
        On Error GoTo EH
        If sTemp <> vbNullString Then
            If InStr(1, sTemp, ",", vbBinaryCompare) > 0 Then
                sTempCity = left(sTemp, InStr(1, sTemp, ",", vbBinaryCompare) - 1)
                sTempCity = Replace(sTempCity, Chr(32), Chr(160))
                sTemp = Replace(sTemp, Replace(sTempCity, Chr(160), Chr(32)), sTempCity)
            End If
            sLossAddress = left(sLossAddress, InStr(1, sLossAddress, String(10, Chr(32))))
            sLossAddress = sLossAddress & Chr(32) & S_z & sTemp
        End If
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
        
        sSQL = sSQL & S_z & sLossAddress & z_S              ' "NewLossLocation, "
        sSQL = sSQL & S_z & sLossStreet & z_S               ' "LossAddress, "
        sSQL = sSQL & S_z & sLossZip & z_S                  ' "LossZip, "
        sSQL = sSQL & S_z & sLossState & z_S                ' "LossState, "
        sSQL = sSQL & S_z & sLossCity & z_S                 ' "LossCity

        sSQL = sSQL & S_z & .AdminInfo.MTGName & z_S                                                    ' "MortgageCompany, "
        
        'Val function Bug in VB6
        'http://msdn.microsoft.com/vbasic/productinfo/previous/vb6/tips/01pasttips.asp
        'Need to parse out both % and !  because these trailing equate to Double and Single
        'and Val bugs because it can't convert Double or single into integer
        sTemp = LTrim(Right(Replace(.PolicyInfo.Deductible, ",", vbNullString), 10))
        sSQL = sSQL & Val(goUtil.utCleanValString(sTemp)) & ", "  ' "Deductible, " 'CUR
        sTemp = Replace(.PolicyInfo.BuildingLim, ",", vbNullString)
        sSQL = sSQL & Val(goUtil.utCleanValString(sTemp)) & ", "                   ' "BuildingPolicyLimit, " 'CUR
        sTemp = Replace(.PolicyInfo.ContentsLim, ",", vbNullString)
        sSQL = sSQL & Val(goUtil.utCleanValString(sTemp)) & ", "                   ' "PersonalPropertyLimit, " 'CUR
        sDate = .AdminInfo.LossDate
        If Not IsDate(sDate) Then
            sDate = Now()
        End If
        sSQL = sSQL & S_z & Format(CDate(sDate), "YYYYMMDD") & z_S  'Loss Date
        sSQL = sSQL & S_z & sIBNum & z_S                            ' "ClaimNo, "
        If Trim(.AdminInfo.MTGName) = vbNullString Then
            sSQL = sSQL & S_z & vbNullString & z_S                  ' "ConfirmMortgageeIsCorrect, "
            sSQL = sSQL & True & ", "                               ' "ConfirmMortgCorrNA, "
            sSQL = sSQL & S_z & vbNullString & z_S                  ' "ExplainedMortgageeChecks, "
            sSQL = sSQL & True & ", "                               ' "ExplainMortgCheckNA, "
        Else
            sSQL = sSQL & S_z & "Y" & z_S                           ' "ConfirmMortgageeIsCorrect, "
            sSQL = sSQL & False & ", "                              ' "ConfirmMortgCorrNA, "
            sSQL = sSQL & S_z & "Y" & z_S                           ' "ExplainedMortgageeChecks, "
            sSQL = sSQL & False & ", "
        End If
        
        sSQL = sSQL & S_z & "Y" & z_S                               ' "ExplainedAndGaveEstToCustomer, " ' = "Y"
        sSQL = sSQL & S_z & "Y" & z_S                               ' "ExplandGaveRCVFormToCust, " '= "Y"
        sSQL = sSQL & S_z & "Y" & z_S                               ' "ConfirmedCoverage, " ' = "Y"
        sSQL = sSQL & False & ", "                                  ' "PriorLoss, " ' = False
        sSQL = sSQL & False & ", "                                  ' "PriorLossNA, " ' = False
        sSQL = sSQL & dTaxPercent & ", "                            ' "TaxPercent, "
        sSQL = sSQL & S_z & sTypeOfLoss & S_z & " ) "               ' "TypeOfLoss, "
    End With
    
    SQL_EASYCLAIM_10135 = sSQL
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    sErrDesc = mudtASNLossReport.ASNLoss.AdminInfo.SALN & vbCrLf & sErrDesc
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLR_ClassName & vbCrLf & "Private Function SQL_EASYCLAIM_10135"
End Function


Private Function zDM() As String
    'ASN format Data Delimeter use ': ' Suffix the Labels of exactly what we need
    zDM = ": "
End Function

Private Function zSTART_PCI() As String
    'ASN format Header for start of Policy Coverage Information
    zSTART_PCI = "======================== POLICY COVERAGE INFORMATION ========================="
End Function

Private Function zSTART_PAY() As String
    'ASN format Header for start of Payment Information
    zSTART_PAY = "============================ PAYMENT INFORMATION ============================="
End Function
Private Function zDED_C1() As String
    'ASN Description prefix before Deductible value for clause 1
    zDED_C1 = "DEDUCTIBLE CLAUSE #1"
End Function

Private Function zDED_C2() As String
    'ASN Description prefix before Deductible value for clause 2
    zDED_C2 = "DEDUCTIBLE CLAUSE #2"
End Function

Private Function zEOP() As String
    'ASN tells printer that we are at End of Page
    zEOP = Chr(12)
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

