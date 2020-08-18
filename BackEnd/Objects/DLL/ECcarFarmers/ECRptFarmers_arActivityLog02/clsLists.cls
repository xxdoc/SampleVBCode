VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLists"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements V2ECKeyBoard.clsRptLists
Private moARReportObject As Object

Private Property Get clsRptLists_ARReportObject() As Object
    Set clsRptLists_ARReportObject = moARReportObject
End Property

Private Property Get clsRptLists_ClassName() As String
    clsRptLists_ClassName = App.EXEName & ".clsLists"
End Property

Private Function clsRptLists_GetARMiscDelimParamsCol(plVersion As Long, _
                                         Optional psDelimParams As String = vbNullString) As Collection
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim saryParamDelim() As String
    Dim lCount As Long
    Dim saryParam() As String
    Dim sParamName As String
    Dim sParamValue As String
    'Put all the RPTDB classes here
    Dim oRptDB As V2ECKeyBoard.clsRptDB
    'Include all versions
    Dim oLatestRptDB As ECRptFarmers_arActivityLog02.clsRptDB
    
    'Split the Params
    saryParamDelim = Split(psDelimParams, "|")
    
    'Get the Correct Version class
    Select Case UCase(plVersion)
        Case Else
            'Lastest Version Goes HERE
            Set oLatestRptDB = New ECRptFarmers_arActivityLog02.clsRptDB
            'populate specific Params here
            For lCount = LBound(saryParamDelim) To UBound(saryParamDelim)
                If saryParamDelim(lCount) <> vbNullString Then
                    saryParam = Split(saryParamDelim(lCount), "=")
                    sParamName = saryParam(0)
                    sParamValue = saryParam(1)
                    Select Case UCase(sParamName)
                        Case UCase("pbPrintBlankReport")
                            oLatestRptDB.PrintBlankReport = sParamValue
                        Case Else
                    End Select
                End If
            Next
            Set oRptDB = oLatestRptDB
    End Select
    
    oRptDB.SetUtilObject goUtil
   
    'Loop through the Params and set the Rpr DB properties
    For lCount = LBound(saryParamDelim) To UBound(saryParamDelim)
        If saryParamDelim(lCount) <> vbNullString Then
            saryParam = Split(saryParamDelim(lCount), "=")
            sParamName = saryParam(0)
            sParamValue = saryParam(1)
            Select Case UCase(sParamName)
                Case UCase("psAssignmentsID")
                   oRptDB.AssignmentsID = sParamValue
                Case UCase("psID")
                   oRptDB.ID = sParamValue
                Case UCase("psClientClaimNo")
                   oRptDB.ClientClaimNo = sParamValue
                Case UCase("psIBNo")
                    oRptDB.IBNo = sParamValue
                Case UCase("psXportPath")
                    oRptDB.sXportPath = sParamValue
                Case UCase("pXportType")
                    oRptDB.XportType = CLng(sParamValue)
                Case UCase("pPDFJPEGQuality")
                    oRptDB.PDFJPEGQuality = CLng(sParamValue)
                Case UCase("pbPreview")
                    oRptDB.bPreview = CBool(sParamValue)
                Case UCase("pbShowPrintButton")
                    oRptDB.bShowPrintButton = CBool(sParamValue)
                Case UCase("plPrintCopies")
                    oRptDB.PrintCopies = CLng(sParamValue)
                Case UCase("pbGetObjectOnly")
                    oRptDB.bGetObjectOnly = CBool(sParamValue)
            End Select
        End If
    Next
    
    Set clsRptLists_GetARMiscDelimParamsCol = oRptDB.MiscDelimParamsCol
    
    'cleanup
    Set oLatestRptDB = Nothing
    If Not oRptDB Is Nothing Then
        oRptDB.CLEANUP
        Set oRptDB = Nothing
    End If
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Set clsRptLists_GetARMiscDelimParamsCol = Nothing
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsRptLists_ClassName & vbCrLf & "Public Function GetARMiscDelimParamsCol"
End Function

Private Function clsRptLists_GetARReport(plVersion As Long, _
                                         Optional psDelimParams As String = vbNullString) As Object
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim saryParamDelim() As String
    Dim lCount As Long
    Dim saryParam() As String
    Dim sParamName As String
    Dim sParamValue As String
    'Put all the RPTDB classes here
    Dim oRptDB As V2ECKeyBoard.clsRptDB
    'Include all versions
    Dim oLatestRptDB As ECRptFarmers_arActivityLog02.clsRptDB
    
    'Split the Params
    saryParamDelim = Split(psDelimParams, "|")
    
    'Get the Correct Version class
    Select Case UCase(plVersion)
        Case Else
            'Lastest Version Goes HERE
            Set oLatestRptDB = New ECRptFarmers_arActivityLog02.clsRptDB
            'populate specific Params here
            For lCount = LBound(saryParamDelim) To UBound(saryParamDelim)
                If saryParamDelim(lCount) <> vbNullString Then
                    saryParam = Split(saryParamDelim(lCount), "=")
                    sParamName = saryParam(0)
                    sParamValue = saryParam(1)
                    Select Case UCase(sParamName)
                        Case UCase("pbPrintBlankReport")
                            oLatestRptDB.PrintBlankReport = sParamValue
                        Case UCase("pbExportXML")
                            oLatestRptDB.ExportXML = sParamValue
                        Case UCase("pbExportXMLOnly")
                            oLatestRptDB.ExportXMLOnly = sParamValue
                        Case Else
                    End Select
                End If
            Next
            Set oRptDB = oLatestRptDB
    End Select
    
    oRptDB.SetUtilObject goUtil
   
    'Loop through the Params and set the Rpr DB properties
    For lCount = LBound(saryParamDelim) To UBound(saryParamDelim)
        If saryParamDelim(lCount) <> vbNullString Then
            saryParam = Split(saryParamDelim(lCount), "=")
            sParamName = saryParam(0)
            sParamValue = saryParam(1)
            Select Case UCase(sParamName)
                Case UCase("psAssignmentsID")
                   oRptDB.AssignmentsID = sParamValue
                Case UCase("psID")
                   oRptDB.ID = sParamValue
                Case UCase("psClientClaimNo")
                   oRptDB.ClientClaimNo = sParamValue
                Case UCase("psIBNo")
                    oRptDB.IBNo = sParamValue
                Case UCase("psXportPath")
                    oRptDB.sXportPath = sParamValue
                Case UCase("pXportType")
                    oRptDB.XportType = CLng(sParamValue)
                Case UCase("pPDFJPEGQuality")
                    oRptDB.PDFJPEGQuality = CLng(sParamValue)
                Case UCase("pbPreview")
                    oRptDB.bPreview = CBool(sParamValue)
                Case UCase("pbShowPrintButton")
                    oRptDB.bShowPrintButton = CBool(sParamValue)
                Case UCase("plPrintCopies")
                    oRptDB.PrintCopies = CLng(sParamValue)
                Case UCase("pbGetObjectOnly")
                    oRptDB.bGetObjectOnly = CBool(sParamValue)
            End Select
        End If
    Next
    
    'Print the AR Report
    If oRptDB.PrintARReport() Then
        Set clsRptLists_GetARReport = oRptDB.ARObject
        Set moARReportObject = oRptDB.ARObject
    End If
    'cleanup
    Set oLatestRptDB = Nothing
    If Not oRptDB Is Nothing Then
        oRptDB.CLEANUP
        Set oRptDB = Nothing
    End If
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsRptLists_ClassName & vbCrLf & "Public Function GetARReport"
End Function

Private Function clsRptLists_SetUtilObject(oUtilObject As Object) As Boolean
    On Error Resume Next
    Set goUtil = oUtilObject
    If Err.Number = 0 Then
        clsRptLists_SetUtilObject = True
    Else
       clsRptLists_SetUtilObject = False
    End If
End Function

Private Function clsRptLists_CLEANUP() As Boolean
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    If Not moARReportObject Is Nothing Then
        Unload moARReportObject
        Set moARReportObject = Nothing
    End If
    Set goUtil = Nothing

    clsRptLists_CLEANUP = True
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsRptLists_ClassName & vbCrLf & "Public Function CLEANUP"
End Function
