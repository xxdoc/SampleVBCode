VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} arPDFDoc 
   Caption         =   "PDF Document"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19606
   SectionData     =   "arPDFDoc.dsx":0000
End
Attribute VB_Name = "arPDFDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Report Items
Private moLists As V2ECKeyBoard.clsLists
Private moLR As V2ECKeyBoard.clsCarLR


'BGS 12.27.2000 Need this to see if the Report is still active.
Private mbActiveFlag As Boolean
'Chain Reports
'Private mbChainPgBrk As Boolean
Private mbChainFlag As Boolean
Private mlChainCount As Long
Private mcolChainReports As Collection ' Contains Reports Chained to it to be added to the Sub Report object
Private moChainReport As Object
Private moPDFDoc As CAcroPDDoc 'Create using Late Binding ONLY!!!
Private moPDFPage As CAcroPDPage 'Create using moPDFDoc
Private moPDFRec As CAcroRect  'Create using Late Binding ONLY!!!
Private moPDFSize As Object  'Create using Late Binding ONLY!!!
Private mlPDFPageCount As Long
Private mlCurPDFPage As Long
Private msPDFDocPath As String
'Portrait Dims
Private Const ARDoc_Port_PrintWidth = 11700
Private Const ImgPDF_Port_H = 14800
Private Const ImgPDF_Port_W = 11700
Private Const Detail_Port_H = 14800

'landScape Dims
Private Const ARDoc_Land_PrintWidth = 15000
Private Const ImgPDF_Land_H = 11700
Private Const ImgPDF_Land_W = 14760
Private Const Detail_Land_H = 11700


'

Public Property Let ChainReport(poActReport As Object)
    Set moChainReport = poActReport
End Property
Public Property Set ChainReport(poActReport As Object)
    Set moChainReport = poActReport
End Property
Public Property Get ChainReport() As Object
    Set ChainReport = moChainReport
End Property

Public Property Let Lists(poLists As V2ECKeyBoard.clsLists)
    Set moLists = poLists
End Property
Public Property Set Lists(poLists As V2ECKeyBoard.clsLists)
    Set moLists = poLists
End Property
Public Property Get Lists() As V2ECKeyBoard.clsLists
    Set Lists = moLists
End Property
Public Property Let LR(poLR As V2ECKeyBoard.clsCarLR)
    Set moLR = poLR
End Property
Public Property Set LR(poLR As V2ECKeyBoard.clsCarLR)
    Set moLR = poLR
End Property
Public Property Get LR() As V2ECKeyBoard.clsCarLR
    Set LR = moLR
End Property

Public Property Get ClassName() As String
    ClassName = App.EXEName & "." & Me.Name
End Property

Public Property Get ActiveFlag() As Boolean
    ActiveFlag = mbActiveFlag
End Property
Public Property Let ActiveFlag(pbFlag As Boolean)
    mbActiveFlag = pbFlag
End Property

Public Function ExportME(psXportPath As String, pXportType As ExportType) As Boolean
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    ExportME = Lists.ExportARReport(Me, psXportPath, pXportType)
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & Me.Name & vbCrLf & "Public Function ExportME"
End Function

Private Sub ActiveReport_ReportStart()
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim lRet As Long
    
    '---Pdf-------------------------------
    msPDFDocPath = moLR.PrnData
    lblPdfFileError.Caption = msPDFDocPath
    Set moPDFDoc = CreateObject("AcroExch.PDDoc")
    lRet = moPDFDoc.Open(msPDFDocPath)
    If lRet Then
        mlPDFPageCount = moPDFDoc.GetNumPages
        'Start off on Page 1
        mlCurPDFPage = 1
    End If
   
    '--Pdf-----------------------------
    
    'Set the Chain flag if we have any
    If Not mcolChainReports Is Nothing Then
        If Not mbChainFlag Then
            mbChainFlag = True
            mlChainCount = 1
        End If
    Else
        mbChainFlag = False
    End If
    
    'If we have Chained Reports...
    If mbChainFlag Then
        Set moChainReport = mcolChainReports(mlChainCount)
        'Start the daisy linking here
        SetNextChainReport mlChainCount, mcolChainReports
        'Set the ref to sub reports in this Report
        Set subChain.Object = moChainReport
    Else
        If Not moChainReport Is Nothing Then
            'Set the ref to sub reports in this Report
            Set subChain.Object = moChainReport
        End If
        
    End If
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & Me.Name & vbCrLf & "Private Sub ActiveReport_ReportStart"
End Sub

Private Sub ActiveReport_ReportEnd()
    On Error Resume Next
    mbActiveFlag = True
    Dim oAR As Object
    'Clean up chain reports collection and objects
    If Not mcolChainReports Is Nothing Then
        For Each oAR In mcolChainReports
            Unload oAR
            Set oAR = Nothing
        Next
        Set mcolChainReports = Nothing
        Unload moChainReport
        Set moChainReport = Nothing
    End If
    If Not moPDFDoc Is Nothing Then
        moPDFDoc.Close
    End If
    Clipboard.Clear
    imgPDF.Picture = LoadPicture()
    Set moPDFSize = Nothing
    Set moPDFPage = Nothing
    Set moPDFRec = Nothing
    Set moPDFDoc = Nothing
    
    Set moLR = Nothing
    
End Sub

'For Chained Reports
Public Sub AddChainReport(poActiveReport As Object)
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    If mcolChainReports Is Nothing Then
        Set mcolChainReports = New Collection
    End If
    
    mcolChainReports.Add poActiveReport
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & Me.Name & vbCrLf & "Public Sub AddChainReport"
End Sub
'For Chained Reports
Public Sub SetNextChainReport(plChainCount As Long, pcolChainReports As Collection)
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    If plChainCount + 1 <= pcolChainReports.Count Then
        Set pcolChainReports(plChainCount).ChainReport = pcolChainReports(plChainCount + 1)
        plChainCount = plChainCount + 1
        'Do daisy again
        pcolChainReports(plChainCount - 1).SetNextChainReport plChainCount, pcolChainReports
    End If
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & Me.Name & vbCrLf & "Public Sub SetNextChainReport"
End Sub

Private Sub Detail_Format()
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim lRet As Long

    'BGS 2.23.2004 ALWAYS SET the PDF PIcture to Nothing
    'This will clear out the memory from any previous Picture Load from Clipboard.
    imgPDF.Picture = LoadPicture()
    If mlCurPDFPage <= mlPDFPageCount And mlCurPDFPage <> 0 Then
        
        'Set the Current page
        'Base is 0 so -1 from Current Page Count
        Set moPDFPage = moPDFDoc.AcquirePage(mlCurPDFPage - 1)
        'Get the Size Object from Adobe
        Set moPDFSize = moPDFPage.GetSize
        'Create a PDF Rectangle
        Set moPDFRec = CreateObject("AcroExch.Rect")
        
        'Set the Size of the PDF Rec to the Document size
        moPDFRec.Bottom = moPDFSize.Y + (moPDFSize.Y * 0.25) '+ 25 %
        moPDFRec.left = 0
        moPDFRec.Right = moPDFSize.X + (moPDFSize.X * 0.25)  '+ 25 %
        moPDFRec.top = 0
        
        ' Switch between portrait and landscape
        If (moPDFSize.X < moPDFSize.Y) Then
            'Portrait
            arPDFDoc.Printer.Orientation = ddOPortrait
            arPDFDoc.PrintWidth = ARDoc_Port_PrintWidth
            imgPDF.Height = ImgPDF_Port_H
            imgPDF.Width = ImgPDF_Port_W
            Detail.Height = Detail_Port_H
        Else
            arPDFDoc.Printer.Orientation = ddOLandscape
            arPDFDoc.PrintWidth = ARDoc_Land_PrintWidth
            imgPDF.Height = ImgPDF_Land_H
            imgPDF.Width = ImgPDF_Land_W
            Detail.Height = Detail_Land_H
        End If
        'Clear the clipboard
        Clipboard.Clear
        'Send to the ClipBoard at 125%
        Call moPDFPage.CopyToClipboard(moPDFRec, 0, 0, 125)
        imgPDF.Picture = Clipboard.GetData(vbCFBitmap)
    Else
        If mlCurPDFPage = mlPDFPageCount + 1 Then
            If Not moChainReport Is Nothing Then
                subChain.Visible = True
                ReportFooter.Visible = True
            Else
                subChain.Visible = False
                ReportFooter.Visible = False
            End If
             Exit Sub 'Bail here!
        ElseIf mlCurPDFPage = 0 Then
            lblPdfFileError.Visible = True
            imgError.Visible = True
            If Not moChainReport Is Nothing Then
                subChain.Visible = True
                ReportFooter.Visible = True
            Else
                subChain.Visible = False
                ReportFooter.Visible = False
            End If
        Else
            Exit Sub
        End If
    End If
    
    mlCurPDFPage = mlCurPDFPage + 1
    Detail.PrintSection
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & Me.Name & vbCrLf & "Private Sub Detail_Format"
End Sub




