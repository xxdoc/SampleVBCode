VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} arGenText 
   Caption         =   "Text"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19606
   SectionData     =   "arGenText.dsx":0000
End
Attribute VB_Name = "arGenText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private moLists As V2ECKeyBoard.clsLists
Private mvData As Variant
Private moLR As V2ECKeyBoard.clsCarLR
Private mlPos As Long
'Chain Reports
'Private mbChainPgBrk As Boolean
Private mbChainFlag As Boolean
Private mlChainCount As Long
Private mcolChainReports As Collection ' Contains Reports Chained to it to be added to the Sub Report object
Private moChainReport As Object

Public Property Get ClassName() As String
    ClassName = App.EXEName & "." & Me.Name
End Property

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

Public Property Let TextData(psData As String)
    mvData = Split(psData, vbCrLf, , vbBinaryCompare)
    mlPos = 0
End Property
Public Property Get TextData() As String
    TextData = Join(mvData, vbCrLf)
End Property

Private Sub ActiveReport_ReportEnd()
    On Error Resume Next
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
    Set moLR = Nothing
End Sub

Private Sub ActiveReport_ReportStart()
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    If moLR Is Nothing Then
        Exit Sub
    End If
    mvData = Split(moLR.PrnData, vbCrLf, , vbBinaryCompare)
    mlPos = 0
    
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
        Set subLossASNChain.Object = moChainReport
    Else
        If Not moChainReport Is Nothing Then
            'Set the ref to sub reports in this Report
            Set subLossASNChain.Object = moChainReport
        End If
        
    End If
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & Me.Name & vbCrLf & "Private Sub ActiveReport_ReportStart"
End Sub

Private Sub ActiveReport_Terminate()
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    Set moLists = Nothing
    Set moLR = Nothing
    Set mcolChainReports = Nothing
    Set moChainReport = Nothing
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Private Sub ActiveReport_Terminate"
End Sub

Private Sub Detail_Format()
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    If mlPos <= UBound(mvData, 1) Then
        If InStr(1, CStr(mvData(mlPos)), INSERT_PAGE_BREAK, vbBinaryCompare) > 0 Then
            fText.Text = vbNullString
            Detail.NewPage = ddNPAfter
        Else
            fText.Text = mvData(mlPos)
            Detail.NewPage = ddNPNone
        End If
        
    Else
        If mlPos = UBound(mvData, 1) + 1 Then
            If Not moChainReport Is Nothing Then
'                Me.Detail.NewPage = ddNPBefore
                subLossASNChain.Visible = True
                ReportFooter.Visible = True
                fText.Visible = False
            Else
                subLossASNChain.Visible = False
                ReportFooter.Visible = False
                Exit Sub 'Bail here!
            End If
        Else
            Exit Sub 'Bail here!
        End If
    End If
    
    mlPos = mlPos + 1
    Detail.PrintSection
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & Me.Name & vbCrLf & "Private Sub Detail_Format"
End Sub

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

