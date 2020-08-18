VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} arRptPaymentReq 
   Caption         =   "Payment Request"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19606
   SectionData     =   "arRptPaymentReq.dsx":0000
End
Attribute VB_Name = "arRptPaymentReq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Report Items
Private moLists As V2ECKeyBoard.clsCarLists
Private mcolProperty As Collection

'BGS 12.27.2000 Need this to see if the Report is still active.
Private mbActiveFlag As Boolean
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

Public Property Get ActiveFlag() As Boolean
    ActiveFlag = mbActiveFlag
End Property
Public Property Let ActiveFlag(pbFlag As Boolean)
    mbActiveFlag = pbFlag
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
    Err.Raise lErrNum, , sErrDesc & vbCrLf & Me.Name & vbCrLf & "Public Sub SetProperty"
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
    Err.Raise lErrNum, , sErrDesc & vbCrLf & Me.Name & vbCrLf & "Public Function ExportME"
End Function

Private Sub ActiveReport_ReportStart()
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    With Me
        'I Payment Info
        .fCheckNo.Text = GetProperty("fCheckNo")
        .fCheckofA.Text = GetProperty("fCheckofA")
        .fCheckofB.Text = GetProperty("fCheckofB")
        .fClientClaimNo.Text = GetProperty("fClientClaimNo")
        If GetProperty("fDatereq") <> NULL_DATE Then
            .fDatereq.Text = Format(GetProperty("fDatereq"), "mm/dd/yy")
        Else
            .fDatereq.Text = vbNullString
        End If
        .fDraft.Text = GetProperty("fDraft")
        .fPayee.Text = GetProperty("fPayee")
        .fPayeeAddress.Text = GetProperty("fPayeeAddress")
        .fPolicyNo.Text = GetProperty("fPolicyNo")
        .fReqBy.Text = GetProperty("fReqBy")
        
        'II Property
        
        'Dwelling
        .fDwellLessDed.Text = Format(GetProperty("fDwellLessDed"), "$##,##0.00")
        .fDwellReq.Text = Format(GetProperty("fDwellReq"), "$##,##0.00")
        .fDwellReqRes.Text = Format(GetProperty("fDwellReqRes"), "$##,##0.00")
        .fDwellTotReq.Text = Format(GetProperty("fDwellTotReq"), "$##,##0.00")
        
        'Other struc
        .fOtherStrucLessDed.Text = Format(GetProperty("fOtherStrucLessDed"), "$##,##0.00")
        .fOtherStrucReq.Text = Format(GetProperty(""), "$##,##0.00")
        .fOtherStrucReqRes.Text = Format(GetProperty("fOtherStrucReqRes"), "$##,##0.00")
        .fOtherStrucTotReq.Text = Format(GetProperty("fOtherStrucTotReq"), "$##,##0.00")
        
        'Person prop
        .fPersonLessDed.Text = Format(GetProperty("fPersonLessDed"), "$##,##0.00")
        .fPersonReq.Text = Format(GetProperty("fPersonReq"), "$##,##0.00")
        .fPersonReqRes.Text = Format(GetProperty("fPersonReqRes"), "$##,##0.00")
        .fPersonTotReq.Text = Format(GetProperty("fPersonTotReq"), "$##,##0.00")
        
        'Loss of use
        .fLossOfUseLessDed.Text = Format(GetProperty("fLossOfUseLessDed"), "$##,##0.00")
        .fLossOfUseReq.Text = Format(GetProperty("fLossOfUseReq"), "$##,##0.00")
        .fLossOfUseReqRes.Text = Format(GetProperty("fLossOfUseReqRes"), "$##,##0.00")
        .fLossOfUseTotReq.Text = Format(GetProperty("fLossOfUseTotReq"), "$##,##0.00")
        
        'Building
        .fBuildingLessDed.Text = Format(GetProperty("fBuildingLessDed"), "$##,##0.00")
        .fBuildingReq.Text = Format(GetProperty("fBuildingReq"), "$##,##0.00")
        .fBuildingReqRes.Text = Format(GetProperty("fBuildingReqRes"), "$##,##0.00")
        .fBuildingTotReq.Text = Format(GetProperty("fBuildingTotReq"), "$##,##0.00")
        
        'Bus-Interup
        .fBusIntLessDed.Text = Format(GetProperty("fBusIntLessDed"), "$##,##0.00")
        .fBusIntReq.Text = Format(GetProperty("fBusIntReq"), "$##,##0.00")
        .fBusIntReqRes.Text = Format(GetProperty("fBusIntReqRes"), "$##,##0.00")
        .fBusIntTotReq.Text = Format(GetProperty("fBusIntTotReq"), "$##,##0.00")
        
        'III Auto
        
        'CPR
        .fCPRLessDed.Text = Format(GetProperty("fCPRLessDed"), "$##,##0.00")
        .fCPRReq.Text = Format(GetProperty("fCPRReq"), "$##,##0.00")
        .fCPRReqRes.Text = Format(GetProperty("fCPRReqRes"), "$##,##0.00")
        .fCPRTotReq.Text = Format(GetProperty("fCPRTotReq"), "$##,##0.00")
        
        'LOU
        .fLOULessDed.Text = Format(GetProperty("fLOULessDed"), "$##,##0.00")
        .fLOUReq.Text = Format(GetProperty("fLOUReq"), "$##,##0.00")
        .fLOUReqRes.Text = Format(GetProperty("fLOUReqRes"), "$##,##0.00")
        .fLOUTotReq.Text = Format(GetProperty("fLOUTotReq"), "$##,##0.00")
        
        'Other
        .fOtherLessDed.Text = Format(GetProperty("fOtherLessDed"), "$##,##0.00")
        .fOtherReq.Text = Format(GetProperty("fOtherReq"), "$##,##0.00")
        .fOtherReqRes.Text = Format(GetProperty("fOtherReqRes"), "$##,##0.00")
        .fOtherTotReq.Text = Format(GetProperty("fOtherTotReq"), "$##,##0.00")
        
        'IV Comment Info
        .fComments.Text = GetProperty("fComments")
        .fFileClosePayAppNo.Text = GetProperty("fFileClosePayAppNo")
        .fFileClosePayAppYes.Text = GetProperty("fFileClosePayAppYes")
        .fInPayOfDescription.Text = GetProperty("fInPayOfDescription")
        .fPresGuarIssuesNo.Text = GetProperty("fPresGuarIssuesNo")
        .fPresGuarIssuesYes.Text = GetProperty("fPresGuarIssuesYes")
    End With
        
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
CHAINED_REPORTS:
    If Not moChainReport Is Nothing Then
        subChain.Visible = True
    Else
        subChain.Visible = False
    End If
End Sub


