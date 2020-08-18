VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} arRptAddlChkFarmers 
   Caption         =   "AddlCheck"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19606
   SectionData     =   "arRptAddlChkFarmers.dsx":0000
End
Attribute VB_Name = "arRptAddlChkFarmers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Report Items
Private moLists As V2ECKeyBoard.clsCarLists
Private mcolProperty As Collection

'Chain Reports
'Private mbChainPgBrk As Boolean
Private mbChainFlag As Boolean
Private mlChainCount As Long
Private mcolChainReports As Collection ' Contains Reports Chained to it to be added to the Sub Report object
Private moChainReport As Object

Public Property Let ChainReport(poActReport As Object)
    Set moChainReport = poActReport
End Property
Public Property Set ChainReport(poActReport As Object)
    Set moChainReport = poActReport
End Property
Public Property Get ChainReport() As Object
    Set ChainReport = moChainReport
End Property

Public Property Let Lists(poLists As V2ECKeyBoard.clsCarLists)
    Set moLists = poLists
End Property
Public Property Set Lists(poLists As V2ECKeyBoard.clsCarLists)
    Set moLists = poLists
End Property
Public Property Get Lists() As V2ECKeyBoard.clsCarLists
    Set Lists = moLists
End Property

Public Property Get ClassName() As String
    ClassName = App.EXEName & "." & Me.Name
End Property

Public Sub SetProperty(psName As String, pvValue As Variant, pType As VbVarType)
    On Error GoTo EH
    Dim sValue As String
    Dim vNewValue As Variant
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    sValue = RTrim(CStr(pvValue))
    'Replace any carriage return flags with vbCrLf
    sValue = Replace(sValue, F_VBCRLF, vbCrLf)
    
    Select Case pType
        Case VbVarType.vbDate
            If IsDate(sValue) Then
                vNewValue = CDate(sValue)
            Else
                vNewValue = CDate(NULL_DATE)
            End If
        Case VbVarType.vbCurrency
            vNewValue = CCur(sValue)
        Case VbVarType.vbString
            vNewValue = CStr(sValue)
        Case VbVarType.vbInteger
            vNewValue = CInt(sValue)
        Case VbVarType.vbBoolean
            vNewValue = CBool(sValue)
        Case VbVarType.vbLong
            vNewValue = CLng(sValue)
        Case VbVarType.vbDouble
            vNewValue = CDbl(sValue)
        Case VbVarType.vbSingle
            vNewValue = CSng(sValue)
    End Select
    
    If mcolProperty Is Nothing Then
        Set mcolProperty = New Collection
    End If
    RemoveProperty psName
    mcolProperty.Add vNewValue, psName
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Public Sub SetProperty"
End Sub

Public Function GetProperty(psName As String) As Variant
    On Error GoTo EH
    
    If Not mcolProperty Is Nothing Then
        GetProperty = mcolProperty(psName)
    Else
        GetProperty = psName & ": Property collection not set!"
    End If
    
    Exit Function
EH:
    GetProperty = vbNullString
    Err.Clear
End Function

Public Function RemoveProperty(psName As String) As Boolean
    On Error GoTo EH
    If Not mcolProperty Is Nothing Then
        mcolProperty.Remove psName
        RemoveProperty = True
    End If
    Exit Function
EH:
    Err.Clear
End Function

Public Function ExportME(psXportPath As String, pXportType As ExportType) As Boolean
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    ExportME = Lists.ExportARReport(Me, psXportPath, pXportType)
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Public Function ExportME"
End Function

Private Sub ActiveReport_ReportStart()
    On Error GoTo EH
    Dim oField As Object
    Dim sTemp As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    For Each oField In Me.Detail.Controls
        If Mid(CStr(oField.Name), 2, 1) = "_" Then
            sTemp = GetProperty(CStr(oField.Name))
            Select Case UCase(sTemp)
                Case "YES"
                    sTemp = "X"
                Case "NO"
                    sTemp = vbNullString
            End Select
            Select Case oField.Name
                Case f_RT63_dtDate.Name, f_RT65_dtApproveDate.Name
                    If sTemp = NULL_DATE Then
                        sTemp = vbNullString
                    End If
                Case f_RT53_cAmountOfCheck.Name, f_RT46_cTexasRoofDepreciation.Name
                    sTemp = Format(sTemp, "$##,##0.00")
            End Select
            oField.Text = sTemp
        Else
            If GetProperty(CStr(oField.Name)) <> vbNullString Then
                oField.Text = GetProperty(CStr(oField.Name))
            End If
        End If
    Next
    
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
        Set subChain.object = moChainReport
    Else
        If Not moChainReport Is Nothing Then
            'Set the ref to sub reports in this Report
            Set subChain.object = moChainReport
        End If
        
    End If
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Private Sub ActiveReport_ReportStart"
End Sub

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
    
    If Not mcolProperty Is Nothing Then
        Set mcolProperty = Nothing
    End If
    
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
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Public Sub AddChainReport"
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
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Public Sub SetNextChainReport"
End Sub

Private Sub Detail_Format()
CHAINED_REPORTS:
    If Not moChainReport Is Nothing Then
        subChain.Visible = True
        ReportFooter.Visible = True
    Else
        subChain.Visible = False
        ReportFooter.Visible = False
    End If
End Sub


