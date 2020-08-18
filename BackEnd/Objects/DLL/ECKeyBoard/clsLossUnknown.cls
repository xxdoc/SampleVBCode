VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLossUnknown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements V2ECKeyBoard.clsCarLR

Public Type UKLossReport
    DateCreated As String
    FileName As String
    PrnData As String
    PrnKey As String
End Type

Public Enum TypeUNKNOWN
    UNKNOWNApd = 0
    UNKNOWNPro
End Enum


Private mudtUKLossReport As UKLossReport
Private moLRs As clsLossReports
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
Private mLossType As TypeUNKNOWN
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
        mLossType = UNKNOWNPro
    ElseIf CStr(pType) = "Auto" Then
        mLossType = UNKNOWNApd
    Else
        mLossType = pType
    End If
End Property
Public Property Get clsCarLR_LossType() As Variant
    clsCarLR_LossType = mLossType
End Property

Public Property Get clsCarLR_LossReport() As Variant
    clsCarLR_LossReport = mudtUKLossReport
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
    clsCarLR_MaxAllowedInChain = 200
End Property

Private Property Let clsCarLR_OleType(psOleType As String)
    msOleType = psOleType
End Property
Private Property Get clsCarLR_OleType() As String
    clsCarLR_OleType = msOleType
End Property

Public Property Let clsCarLR_PrnData(psData As String)
    mudtUKLossReport.PrnData = psData
End Property
Public Property Get clsCarLR_PrnData() As String
    clsCarLR_PrnData = mudtUKLossReport.PrnData
End Property

Public Property Let clsCarLR_PrnKey(psKey As String)
    mudtUKLossReport.PrnKey = psKey
End Property
Public Property Get clsCarLR_PrnKey() As String
    clsCarLR_PrnKey = mudtUKLossReport.PrnKey
End Property

Public Property Get clsCarLR_ClassName() As String
    clsCarLR_ClassName = App.EXEName & ".clsLossUnknown"
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

Public Property Get clsCarLR_ACID() As String
    'ACID will be the File Name for UNKOWN CLASS
    clsCarLR_ACID = msACID
End Property
Public Property Let clsCarLR_ACID(psACID As String)
    msACID = psACID
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
    'Unknown format should only be used for brand new Loss report formats
    'that as of yet do not have a class to process them.  These files must contain only one Loss Report. Otherwise
    'clsLossUnknown will not process it correctly.  This is just a very basic class to get a
    'text only formatted report from here to there to be printed ONLY.
    'clsLossUnknown will NOT and can NOT be used to update DB Fields of anykind.
    
    'we will be processing 1 file to 1 Loss report. So all we have to do
    'Is loop though all the files in the Directory
    
    On Error GoTo EH
    Dim oLR As V2ECKeyBoard.clsCarLR
    Dim colFiles As Collection
    Dim sFile As String
    Dim vFile As Variant
    Dim sData As String
    Dim sPrnData As String
    Dim varyData As Variant
    Dim sFilename As String
    Dim sRawFileName As String
    Dim bPB As Boolean 'True if using it
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    If Not goUtil.utFileExists(psDataPath, True) Then
        Exit Function
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
    
    'Load Unknown format files
    sFile = Dir(psDataPath & "\" & "*.*")
    
    If sFile <> vbNullString Then
        If colFiles Is Nothing Then
            Set colFiles = New Collection
        End If
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
    End If
    
    'Loop through the collection of Files
    For Each vFile In colFiles
        If goUtil.utFileExists(CStr(vFile)) Then
        
            'Get Unknown data
            sData = goUtil.utGetFileData(CStr(vFile))
            
            'Clean out Forced page breaks
            sData = Replace(sData, zEOP, vbNullString, , , vbBinaryCompare)
            
            varyData = Split(sData, vbCrLf)
            sData = vbNullString 'Clear it
            
           'Format Report Data
            'Add the class name to the header line which will be "clsLossUnkown"
            varyData(LBound(varyData, 1)) = clsCarLR_ClassName & vbCrLf & varyData(LBound(varyData, 1))
            
            'Put the modified report data back together
            sData = Join(varyData, vbCrLf)
            
            'Save The Report File
            'Now we need to save the processed file data to proper place and File name
            sFile = CStr(vFile)
            'Get rid of any existing .txt
            sFile = Replace(sFile, ".txt", vbNullString, , , vbTextCompare)
            sFilename = Mid(sFile, InStrRev(sFile, "\") + 1) & ".txt"
            sFilename = psOutPath & "\" & sFilename
            'Check to see if this file was already saved here.
            If goUtil.utFileExists(sFilename) Then
                SetAttr sFilename, vbNormal
                Kill sFilename
            End If
            goUtil.utSaveFileData sFilename, sData
            
            'Make a Backup copy of the Raw File
            'Check to see if the Raw File is already in the
            'BackUp folder first
            sFile = CStr(vFile)
            sFile = Mid(sFile, InStrRev(sFile, "\") + 1)
            msACID = sFile 'Set memeber CRID Variable
            
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
            Set oLR = New clsLossUnknown
            With oLR
                Set .LRs = moLRs
                .PrnKey = sFilename
                'Remove the class marker frm the prn data
                sPrnData = sData
                sPrnData = Replace(sPrnData, clsCarLR_ClassName & vbCrLf, vbNullString, 1)
                .ACID = msACID
                .PrnData = sData
                .PopulateMe
            End With
            moLRs.AddRawLossReport oLR
            'Free loacal memory
            Set oLR = Nothing
        End If
        If bPB Then
            poPB.Value = poPB.Value + 1
        End If
    Next
    
   'CleanUp
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
    Dim iPic As Integer
    
    
    If StrComp(msOleType, "HTML", vbTextCompare) = 0 Then
        iPic = LRPic.lrHTMLHelp
    ElseIf StrComp(msOleType, "TEXT", vbTextCompare) = 0 Then
        iPic = LRPic.lrUnknown
    ElseIf InStr(1, msOleType, "OLEType_pdf", vbTextCompare) > 0 Then
        iPic = LRPic.lrFORMATpdf
    Else
         iPic = LRPic.lrUnknown
    End If
    
    With mudtUKLossReport
         
        'Assigned date Use File creation Date
        Set poitmX = polvw.ListItems.Add(, , Format(msDateAssign, "MM/DD/YY"), , iPic)
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
    
        
        'Use Class name for format
        poitmX.SubItems(LossReports.RFormat - 1) = clsCarLR_ClassName
        
        'Sort by File creation Date and File Name instead of by Assigned Date and Adjuster and Insured Name
        'Format the Date so that it will sort by year first
        poitmX.SubItems(LossReports.RSort - 1) = Format(msDateAssign, "YY/MM/DD") & clsCarLR_PrnKey
        
        'Rememeber that the File path is used as the Key.
        'This allows for quick reference to this particular Report
        'when it is selcted from the listview
        poitmX.SubItems(LossReports.RKey - 1) = clsCarLR_PrnKey
        
        
        poitmX.Selected = False
    End With
    
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
            'If there is an OLE Type of Report Use it, otherwise use Gen Text Print
            Select Case UCase(msOleType)
                Case "TEXT"
                    moLRs.PrintGenText Me, psPrnDeviceName, pbPreview, pbPDF, plHwnd, plChainType
                Case Else
                    If InStr(1, msOleType, "OLEType_pdf", vbTextCompare) > 0 Then
                         moLRs.PrintOlePDFDoc Me, psPrnDeviceName, pbPreview, pbPDF, plHwnd, plChainType
                    Else
                        'Use the General text function in LossReports Class
                        moLRs.PrintGenText Me, psPrnDeviceName, pbPreview, pbPDF, plHwnd, plChainType
                    End If
            End Select
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
            moLRs.FirePrintAppDoc Mid(clsCarLR_PrnKey, InStrRev(clsCarLR_PrnKey, "_", , vbBinaryCompare) + 1)
        End If
    End If
    clsCarLR_PrintMe = True
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLR_ClassName & vbCrLf & "Public Function PrintMe"
End Function

Public Function clsCarLR_ExportME(poLRs As V2ECKeyBoard.clsLossReports, pXportType As V2ECKeyBoard.ExportType) As Boolean
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

Public Function clsCarLR_GetLRSQL(psAppName As String, plVersion As Long, Optional psSQLError As String) As String
    On Error GoTo EH
    Dim sSQL As String
    Dim sFilename As String
    Dim sMess As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    'Need to ucase the appname because ME will and who knows
    'maybe other windows OS will UCase the app exe name why I have not a friggin clue
    Select Case UCase(psAppName)
        '<--------------------------------------WEBCONTROL-------------------------------------->
        Case "WEBCONTROL"
            Select Case plVersion
                'iF Newer versions Require Chnages to SQL ...
                'They will Have  go on top in Select Case
                Case Is >= 108
                    sFilename = clsCarLR_PrnKey
                    sMess = "<-------------->" & vbCrLf
                    sMess = sMess & "Loss Report Format """ & clsCarLR_ClassName & """" & vbCrLf
                    sMess = sMess & "File Path: " & sFilename & vbCrLf
                    sMess = sMess & "Has not been defined in " & App.EXEName & " for Data Base update." & vbCrLf
                    sFilename = Mid(sFilename, InStrRev(sFilename, "\", , vbBinaryCompare) + 1)
                    sMess = sMess & "File Name: " & sFilename & " will be saved under ClaimNoSaln field." & vbCrLf
                    sMess = sMess & "<-------------->" & vbCrLf
                    moLRs.FireErrorMess sMess
                     
                    sSQL = "INSERT INTO ASSIGNMENTS (ClaimNoSaln, "                 '04Adjuster
                    sSQL = sSQL & "Format, "                                        'Format
                    sSQL = sSQL & "LossReport ) "                                   'Memo'37LossReport
                    
                    sSQL = sSQL & "VALUES (" & S_z & sFilename & z_S                '04Adjuster
                    sSQL = sSQL & S_z & clsCarLR_ClassName & z_S                             'Format
                    sSQL = sSQL & S_z & clsCarLR_ClassName & vbCrLf & clsCarLR_PrnData & S_z & " ) "  'Memo'37LossReport
            End Select
            
        '<--------------------------------------EASY CLAIM-------------------------------------->
        Case "EBERLS"
            Select Case plVersion
                'iF Newer versions Require Chnages to SQL ...
                'They will Have  go on top in Select Case
                Case Is >= 10135
                'We do not update Easy Claim 10135 DB with unknown Class
            End Select
            
    End Select
    
    If sSQL = vbNullString Then
        psSQLError = psAppName & " VS " & plVersion & " Not supported." & vbCrLf
    Else
        'Need to clean out any " chr(34) that may appear in the text and
        'replace with "" Chr(34) & Chr(34) for SQL purposes
        sSQL = Replace(sSQL, Chr(34), Chr(34) & Chr(34), , , vbBinaryCompare)
        'Now Set the Begin and end String fields
        sSQL = Replace(sSQL, S_z, S_z_SET, , , vbBinaryCompare)
        sSQL = Replace(sSQL, z_S, z_S_SET, , , vbBinaryCompare)
        
        clsCarLR_GetLRSQL = sSQL
    End If
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLR_ClassName & vbCrLf & "Public Function GetLRSQL"
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
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLR_ClassName & vbCrLf & "Public Function CLEANUP"
End Function

Public Sub clsCarLR_PopulateMe()
    'Here is where we populate the Main udt for the Unknown Format
    On Error GoTo EH
    Dim sPrnFileName As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    With mudtUKLossReport
        sPrnFileName = .PrnKey
        .DateCreated = moLRs.GetCreateDate(sPrnFileName)
        sPrnFileName = Mid(sPrnFileName, InStrRev(sPrnFileName, "\") + 1)
        .FileName = sPrnFileName
    End With
    
        
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarLR_ClassName & vbCrLf & "Private Sub PopulateMe"
End Sub

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
