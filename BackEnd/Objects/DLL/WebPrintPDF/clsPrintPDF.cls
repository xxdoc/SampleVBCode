VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPrintPDF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function PrintPackageToFile(poReportItem As Object, _
                                Optional piMode As VBRUN.FormShowConstants = vbModeless, _
                                Optional psCopyName As String = vbNullString, _
                                Optional pbPrintPreview As Boolean = True, _
                                Optional psSaveToFilePath As String, _
                                Optional psSaveToFileName As String, _
                                Optional pbExportXML As Boolean, _
                                Optional pbExportXMLOnly As Boolean) As Boolean
    On Error GoTo EH
    Dim sParams As String
    Dim sReportName As String
    Dim sReportTitle As String
    Dim srptProjectName As String
    Dim srptClassName As String
    Dim lrptVersion As Long
    Dim sData As String
    Dim saryData() As String
    Dim ocboReport As ComboBox
    Dim itmXReport As ListItem
    Dim MyActReport As ActiveReport
    Dim oCarList As V2ECKeyBoard.clsCarLists
    'If using Adobe PDF Viewer
    Dim sPDFFilePath As String
    Dim sPrintPreview As String
    Dim bUseAdobeReader As Boolean
    'Some Reports need extra Params passed to them
    'Payments
    Dim sRTChecksID As String
    Dim sCheckNum As String
    'Internal Billing
    Dim sIBID As String
    Dim sSupplement As String
    'Photo Reports (Multi Report)
    Dim sPhotoReportNumber As String
    'Worksheet Diagram (Multi Report)
    Dim sDiagramNumber As String
    Dim sNumber As String
    'Loss Report
    Dim oLR As V2ECKeyBoard.clsLossReports
    Dim MyAssignmentsRS As ADODB.Recordset
    Dim sLRFormat As String
    Dim sLossReport As String
    Dim sLRData As String
    'Export to XML FileName
    Dim sXMLFilePath As String
    Dim sXMLFileName As String
    
    sPrintPreview = GetSetting(goUtil.gsMainAppEXEName, "GENERAL", "PRINT_PREVIEW", "USE_ADOBE")
    
    Select Case UCase(sPrintPreview)
        Case "USE_ADOBE"
            bUseAdobeReader = True
    End Select
    
    'if saving to file path always use adobe reader
    If psSaveToFilePath <> vbNullString Then
        bUseAdobeReader = True
    End If
    
    If TypeOf poReportItem Is ComboBox Then
        Set ocboReport = poReportItem
        sData = ocboReport.Text
    ElseIf TypeOf poReportItem Is ListItem Then
        Set itmXReport = poReportItem
'        sData = itmXReport.ListSubItems(GuiPackageItemListView.ReportFormat - 1)
    Else
        Exit Function
    End If
    
    If sData <> vbNullString Then
        sReportTitle = Trim(Left(sData, 200))
        goUtil.utCleanFileFolderName sReportTitle, False
        sData = Mid(sData, InStr(1, sData, String(200, " "), vbBinaryCompare))
        sData = Trim(sData)
        saryData() = Split(sData, "|", , vbBinaryCompare)
        If UBound(saryData, 1) <= 1 Then
            'Check for Loss Report
            sLRFormat = saryData(0)
            If StrComp(sLRFormat, "LRFormat", vbTextCompare) = 0 Then
                Me.SetadoRSAssignments msAssignmentsID
                Set MyAssignmentsRS = Me.adoRSAssignments
                sLRFormat = goUtil.IsNullIsVbNullString(MyAssignmentsRS.Fields("LRFormat"))
                sLossReport = goUtil.IsNullIsVbNullString(MyAssignmentsRS.Fields("LossReport"))
                
                If InStr(1, sLRFormat, "OLEType_pdf", vbTextCompare) > 0 Then
                    sPDFFilePath = goUtil.AttachReposPath & sLossReport
                    If psSaveToFilePath <> vbNullString Then
                        'Do not do Loss Report if Export to xml only is true
                        If pbExportXML And pbExportXMLOnly Then
                            PrintActiveReport = False
                            Set ocboReport = Nothing
                            Set itmXReport = Nothing
                            Set oLR = Nothing
                            Set MyAssignmentsRS = Nothing
                            MsgBox "Loss Reports can not be part of and XML ONLY Export!", vbExclamation
                            Exit Function
                        End If
                        goUtil.utCopyFile sPDFFilePath, psSaveToFilePath & psSaveToFileName
                    Else
                        If pbPrintPreview Then
                            goUtil.utShellExecute , "OPEN", sPDFFilePath, , , vbNormalFocus, True, True, True, sReportTitle
                        Else
                            goUtil.utShellExecute , "PRINT", sPDFFilePath, , , vbNormalFocus, True, True, True, sReportTitle
                        End If
                        DoEvents
                        Sleep 100
                    End If
                Else
                    sPDFFilePath = goUtil.gsInstallDir & "\TempLossReport" & goUtil.utGetTickCount & ".pdf"
                    Set oLR = New V2ECKeyBoard.clsLossReports
                    If StrComp(sLRFormat, "TEXT", vbTextCompare) <> 0 Then
                        sLRData = sLRFormat & vbCrLf & sLossReport
                    Else
                        sLRData = sLossReport
                    End If
                    oLR.CreateExport sLRData, sPDFFilePath, ARPdf
                    If psSaveToFilePath <> vbNullString Then
                        'Do not do Loss Report if Export to xml only is true
                        If pbExportXML And pbExportXMLOnly Then
                            PrintActiveReport = False
                            Set ocboReport = Nothing
                            Set itmXReport = Nothing
                            Set oLR = Nothing
                            Set MyAssignmentsRS = Nothing
                            MsgBox "Loss Reports can not be part of and XML ONLY Export!", vbExclamation
                            Exit Function
                        End If
                        goUtil.utCopyFile sPDFFilePath, psSaveToFilePath & psSaveToFileName
                    Else
                        If pbPrintPreview Then
                            goUtil.utShellExecute , "OPEN", sPDFFilePath, , , vbNormalFocus, True, True, True, sReportTitle
                        Else
                            goUtil.utShellExecute , "PRINT", sPDFFilePath, , , vbNormalFocus, True, True, True, sReportTitle
                        End If
                        DoEvents
                        Sleep 100
                    End If
                    goUtil.utDeleteFile sPDFFilePath
                End If
                PrintActiveReport = True
                GoTo CLEAN_UP
            End If
            
            sPDFFilePath = saryData(0)
            If InStr(1, sPDFFilePath, ".pdf", vbTextCompare) > 0 Then
                sPDFFilePath = goUtil.AttachReposPath & sPDFFilePath
                'Check for Pdf Attachment file
                If psSaveToFilePath <> vbNullString Then
                    'Do not do Attachments if Export to xml only is true
                    If pbExportXML And pbExportXMLOnly Then
                        PrintActiveReport = False
                        Set ocboReport = Nothing
                        Set itmXReport = Nothing
                        Set oLR = Nothing
                        Set MyAssignmentsRS = Nothing
                        MsgBox "Attachments can not be part of and XML ONLY Export!", vbExclamation
                        Exit Function
                    End If
                    goUtil.utCopyFile sPDFFilePath, psSaveToFilePath & psSaveToFileName
                Else
                    If pbPrintPreview Then
                        goUtil.utShellExecute , "OPEN", sPDFFilePath, , , vbNormalFocus, True, True, True, sReportTitle
                    Else
                        goUtil.utShellExecute , "PRINT", sPDFFilePath, , , vbNormalFocus, True, True, True, sReportTitle
                    End If
                    DoEvents
                    Sleep 100
                End If
            End If
            PrintActiveReport = True
            GoTo CLEAN_UP
        End If
        srptProjectName = saryData(0)
        srptClassName = saryData(1)
        lrptVersion = saryData(2)
        'Check For Multi Reports Here
        If TypeOf poReportItem Is ListItem Then
            If UBound(saryData, 1) >= 3 Then
                sNumber = saryData(3)
            End If
            
            'If this is coming from the Package Screen need to populate the Number for certain reports
            If InStr(1, srptProjectName, "_arRptPhotos", vbTextCompare) > 0 Then
                sPhotoReportNumber = sNumber
            ElseIf InStr(1, srptProjectName, "_arWorkSheetDiag", vbTextCompare) > 0 Then
                sDiagramNumber = sNumber
            ElseIf InStr(1, srptProjectName, "_arRptAddlChk", vbTextCompare) > 0 Then
                sCheckNum = sNumber
            ElseIf InStr(1, srptProjectName, "_arRptIB", vbTextCompare) > 0 Then
                sSupplement = sNumber
            End If
        ElseIf InStr(1, srptProjectName, "_arRptPhotos", vbTextCompare) > 0 Then
            'Photo Reports (Multi Report)
            sPhotoReportNumber = saryData(3)
        ElseIf InStr(1, srptProjectName, "_arWorkSheetDiag", vbTextCompare) > 0 Then
            'Worksheet Diagram (Multi Report)
            sDiagramNumber = saryData(3)
        End If
    Else
        Exit Function
    End If
    
    'Build Params List to be passed in to Create Report Object
    'This Object will have list of Report Parameters it requires
    
    sParams = vbNullString
    sParams = sParams & "psAssignmentsID=" & msAssignmentsID & "|"
    'If using Adobe PDF Viewer
    If bUseAdobeReader Then
        sPDFFilePath = goUtil.gsInstallDir & "\TempActiveReport" & goUtil.utGetTickCount & ".pdf"
        sParams = sParams & "psXportPath=" & sPDFFilePath & "|"
        sParams = sParams & "pPDFJPEGQuality=" & "50" & "|"
        sParams = sParams & "pXportType=" & ExportType.ARPdf & "|"
    Else
        sParams = sParams & "pbGetObjectOnly=" & "True" & "|"
    End If
    
    'Certain Reports Need to have some more Params Passed in
    If InStr(1, srptProjectName, "_arRptAddlChk", vbTextCompare) > 0 Then
        'Need to Get the ChecksID and Check Number
        If Not ocboReport Is Nothing Then
            sRTChecksID = CStr(ocboReport.ItemData(ocboReport.ListIndex))
            If Not GetPaymentsParams(sRTChecksID, sCheckNum) Then
                GoTo CLEAN_UP
            End If
        ElseIf Not itmXReport Is Nothing Then
            'the schecknum was already set above
            If Not GetPaymentsParams(sRTChecksID, sCheckNum) Then
                GoTo CLEAN_UP
            End If
        End If
        sParams = sParams & "pRTChecksID=" & sRTChecksID & "|"
        sParams = sParams & "psCheckNum=" & sCheckNum & "|"
    ElseIf InStr(1, srptProjectName, "_arRptIB", vbTextCompare) > 0 Then
        'If the IBID and Supplement Parameters already exist then use them
        'Otherwise have to do Data Call to get em.
        If InStr(1, sData, "pIBID=", vbTextCompare) > 0 And InStr(1, sData, "pSupplement=", vbTextCompare) > 0 Then
            sParams = sParams & saryData(3) & "|"
            sParams = sParams & saryData(4) & "|"
            'Check for Report Title As Well
            If InStr(1, sData, "psReportTitle=", vbTextCompare) > 0 Then
                sReportTitle = Mid(saryData(5), InStr(1, saryData(5), "=", vbTextCompare) + 1)
            End If
        Else
            If Not ocboReport Is Nothing Then
                sIBID = CStr(ocboReport.ItemData(ocboReport.ListIndex))
                If Not GetIBParams(sIBID, sSupplement) Then
                    GoTo CLEAN_UP
                End If
            ElseIf Not itmXReport Is Nothing Then
                'The supplement was already set above
                If Not GetIBParams(sIBID, sSupplement) Then
                    GoTo CLEAN_UP
                End If
            End If
            sParams = sParams & "pIBID=" & sIBID & "|"
            sParams = sParams & "pSupplement=" & sSupplement & "|"
        End If
        sParams = sParams & "pCopyName=" & psCopyName & "|"
    ElseIf InStr(1, srptProjectName, "_arRptPhotos", vbTextCompare) > 0 Then
        'Photo Reports (Multi Report)
        sParams = sParams & "pNumber=" & sPhotoReportNumber & "|"
    ElseIf InStr(1, srptProjectName, "_arWorkSheetDiag", vbTextCompare) > 0 Then
        'Worksheet Diagram (Multi Report)
        sParams = sParams & "pNumber=" & sDiagramNumber & "|"
    End If
    
    sReportName = srptProjectName & "." & srptClassName

    If StrComp(psCopyName, "(-ALL COPIES-)", vbTextCompare) = 0 Then
        'Do a recursive call until All Copies are printed
        'Client company Copy
        If Not PrintActiveReport(poReportItem, , goUtil.gsCurCarDBName & " Copy", pbPrintPreview, psSaveToFilePath, psSaveToFileName) Then
            GoTo CLEAN_UP
        End If
        'Company Copy
        If Not PrintActiveReport(poReportItem, , GetSetting(goUtil.gsAppEXEName, "GENERAL", "CURRENT_COMPANY_NAME", "Company") & " Copy", pbPrintPreview, psSaveToFilePath, psSaveToFileName) Then
            GoTo CLEAN_UP
        End If
        'Remit Copy
        If Not PrintActiveReport(poReportItem, , "Remit Copy", pbPrintPreview, psSaveToFilePath, psSaveToFileName) Then
            GoTo CLEAN_UP
        End If
        'Adjuster Copy
        If Not PrintActiveReport(poReportItem, , "Adjuster Copy", pbPrintPreview, psSaveToFilePath, psSaveToFileName) Then
            GoTo CLEAN_UP
        End If
    Else
        Set oCarList = CreateObject(goUtil.goCurCarList.ClassName)
        If bUseAdobeReader Then
            'Add Export XML Parameters here
            If pbExportXML Then
                sParams = sParams & "pbExportXML=True|"
                If pbExportXMLOnly Then
                    sParams = sParams & "pbExportXMLOnly=True|"
                End If
            End If
            Set MyActReport = oCarList.GetARReport(sReportName, lrptVersion, sParams)
            If goUtil.utFileExists(sPDFFilePath) Or (pbExportXML And pbExportXMLOnly) Then
                If psSaveToFilePath <> vbNullString Then
                    If pbExportXML Then
                        If Not pbExportXMLOnly Then
                            If psCopyName <> vbNullString Then
                                goUtil.utCopyFile sPDFFilePath, psSaveToFilePath & Replace(psSaveToFileName, ".pdf", "_" & psCopyName & ".pdf", , 1, vbTextCompare)
                            Else
                                goUtil.utCopyFile sPDFFilePath, psSaveToFilePath & psSaveToFileName
                            End If
                        End If
                    Else
                        If psCopyName <> vbNullString Then
                            goUtil.utCopyFile sPDFFilePath, psSaveToFilePath & Replace(psSaveToFileName, ".pdf", "_" & psCopyName & ".pdf", , 1, vbTextCompare)
                        Else
                            goUtil.utCopyFile sPDFFilePath, psSaveToFilePath & psSaveToFileName
                        End If
                    End If
                    
                    If pbExportXML Then
                        'change the pdffile path the XML
                        sXMLFilePath = sPDFFilePath
                        sXMLFilePath = Left(sXMLFilePath, InStrRev(sXMLFilePath, ".", , vbBinaryCompare))
                        sXMLFilePath = sXMLFilePath & "xml"
                        'Change the pdf to XML file path
                        sXMLFileName = psSaveToFileName
                        sXMLFileName = Left(sXMLFileName, InStrRev(sXMLFileName, ".", , vbBinaryCompare))
                        sXMLFileName = sXMLFileName & "xml"
                       If psCopyName <> vbNullString Then
                            goUtil.utCopyFile sXMLFilePath, psSaveToFilePath & Replace(sXMLFileName, ".xml", "_" & psCopyName & ".xml", , 1, vbTextCompare)
                        Else
                            goUtil.utCopyFile sXMLFilePath, psSaveToFilePath & sXMLFileName
                        End If
                    End If
                Else
                    If pbPrintPreview Then
                        goUtil.utShellExecute , "OPEN", sPDFFilePath, , , vbNormalFocus, True, True, True, sReportTitle
                    Else
                        goUtil.utShellExecute , "PRINT", sPDFFilePath, , , vbNormalFocus, True, True, True, sReportTitle
                    End If
                    
                End If
                ' DoEvents
                Sleep 100
                goUtil.utDeleteFile sPDFFilePath
                goUtil.utDeleteFile sXMLFilePath
                If Not MyActReport Is Nothing Then
                    Unload MyActReport
                    Set MyActReport = Nothing
                End If
'                oCarList.CLEANUP
                Set oCarList = Nothing
            End If
        Else
            'Using Active Report Viewer
            Set MyActReport = oCarList.GetARReport(sReportName, lrptVersion, sParams)
        
            If pbPrintPreview Then
                If mArv Is Nothing Then
                    Set mArv = New V2ARViewer.clsARViewer
                    mArv.SetUtilObject goUtil
                End If
'                If Not moForm Is Nothing Then
'                    If StrComp(psCopyName, "(-ALL COPIES-)", vbTextCompare) <> 0 Then
'                        Unload moForm
'                        Set moForm = Nothing
'                    End If
'                End If
                With mArv
                    'Pass in true to have Active reports process on separate thread.
                    'This will allow the viewer to load while the report is processing
                    'false will force the report to run on single thread
                    MyActReport.Run False 'True
                    .objARvReport = MyActReport
                    .sRptTitle = sReportTitle
                    .HidePrintButton = False
                    .ShowReportOnForm moForm, piMode
        
                    Unload .objARvReport
                    Set .objARvReport = Nothing
                End With
            Else
                MyActReport.PrintReport False
            End If
            Unload MyActReport
            Set MyActReport = Nothing
'            oCarList.CLEANUP
            Set oCarList = Nothing
        End If
    End If
    PrintActiveReport = True
CLEAN_UP:
    'Cleanup
    Set ocboReport = Nothing
    Set itmXReport = Nothing
    Set oLR = Nothing
    Set MyAssignmentsRS = Nothing
    
    'Clear the Local ref to this report object only
    'The actual cleanup of this active report object will occur within gARV
    PrintActiveReport = True
    Exit Function
EH:
    Screen.MousePointer = vbDefault
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function PrintActiveReport"
End Function


Private Sub Class_Initialize()
    Set goUtil = New V2ECKeyBoard.clsUtil
End Sub
