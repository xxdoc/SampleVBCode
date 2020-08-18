VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} arLossXML01 
   Caption         =   "ActiveReport1"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19606
   SectionData     =   "arLossXML01.dsx":0000
End
Attribute VB_Name = "arLossXML01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private moLists As V2ECKeyBoard.clsCarLists
Private moLossXML01 As V2ECcarFarmers.clsLossXML01

Private msubAssDetails As arsubLossXML01AssDetails  ' Assignment Details
Private msubLossDetails As arsubLossXML01LossDetails   ' Loss Details
Private msubContactDetails As arsubLossXML01ContactDetails   ' Contact Details
Private msubCOV As arsubLossXML01cov ' Coverageges, Under Policy Detail
Private msubED As arsubLossXML01ed 'Endorsemnts, Under Policy detail
Private msubPAY As arsubLossXML01pay ' Payment Detail
Private msubPLH As arsubLossXML01plh 'Prior Loss Detail
Private msubCAL As arsubLossXML01cal 'Activities

Private mXML01LossReport As XML01LossReport 'XML01 Loss Repport user defined type

'Chain Reports
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

Public Property Let LossXML01(poLossXML01 As V2ECcarFarmers.clsLossXML01)
    Set moLossXML01 = poLossXML01
End Property
Public Property Set LossXML01(poLossXML01 As V2ECcarFarmers.clsLossXML01)
    Set moLossXML01 = poLossXML01
End Property
Public Property Get LossXML01() As V2ECcarFarmers.clsLossXML01
    Set LossXML01 = moLossXML01
End Property


Public Property Let LossReport(pLossReport As V2ECcarFarmers.XML01LossReport)
    mXML01LossReport = pLossReport
End Property

Public Property Get ClassName() As String
    ClassName = App.EXEName & "." & Me.Name
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
        Unload subLossXML01Chain
        Set subLossXML01Chain = Nothing
        Set mcolChainReports = Nothing
        Unload moChainReport
        Set moChainReport = Nothing
    End If
    
    'Clean up the Sub Report Objects
    Unload subLossXML01AssDetails
    Set subLossXML01AssDetails.object = Nothing
    Unload subLossXML01LossDetails
    Set subLossXML01LossDetails.object = Nothing
    Unload subLossXML01ContactDetails
    Set subLossXML01ContactDetails.object = Nothing
    Unload subLossXML01cov.object
    Set subLossXML01cov.object = Nothing
    Unload subLossXML01ed.object
    Set subLossXML01ed.object = Nothing
    Unload subLossXML01pay.object
    Set subLossXML01pay.object = Nothing
    Unload subLossXML01plh.object
    Set subLossXML01plh.object = Nothing
    Unload subLossXML01cal.object
    Set subLossXML01cal.object = Nothing
    
    
    Set msubAssDetails = Nothing
    Set msubLossDetails = Nothing
    Set msubContactDetails = Nothing
    Set msubCOV = Nothing
    Set msubED = Nothing
    Set msubPAY = Nothing
    Set msubPLH = Nothing
    Set msubCAL = Nothing
        
End Sub

Private Sub ActiveReport_ReportStart()
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    'Instance sub reports
    Set msubAssDetails = New arsubLossXML01AssDetails
    Set msubLossDetails = New arsubLossXML01LossDetails
    Set msubContactDetails = New arsubLossXML01ContactDetails
    Set msubCOV = New arsubLossXML01cov
    Set msubED = New arsubLossXML01ed
    Set msubPAY = New arsubLossXML01pay
    Set msubPLH = New arsubLossXML01plh
    Set msubCAL = New arsubLossXML01cal
    
    'Set their data collections
    Set msubCOV.CoverageRS = mXML01LossReport.XML01Loss.CoverageRS
    msubCOV.DedType = "Property"
    
    Set msubAssDetails.AssDetailRS = mXML01LossReport.XML01Loss.AssignmentDetailRS
    Set msubLossDetails.LossDetailRS = mXML01LossReport.XML01Loss.LossDetailRS
    Set msubContactDetails.LossXML01 = moLossXML01
    Set msubContactDetails.ContactDetailRS = mXML01LossReport.XML01Loss.ContactDetailRS
    Set msubContactDetails.ContactRS = mXML01LossReport.XML01Loss.ContactsRS
    Set msubContactDetails.AddressRS = mXML01LossReport.XML01Loss.AddressRS
    
    Set msubED.EndorsementRS = mXML01LossReport.XML01Loss.EndorsementRS
    Set msubPAY.PaymentDetailRS = mXML01LossReport.XML01Loss.PaymentDetailRS
    Set msubPLH.PriorLossDetailRS = mXML01LossReport.XML01Loss.PriorLossDetailRS
    Set msubCAL.ActivitiesRS = mXML01LossReport.XML01Loss.ActivitiesRS
    
    'Set the ref to sub reports in this Report
    subLossXML01AssDetails.object = msubAssDetails.object
    subLossXML01LossDetails.object = msubLossDetails.object
    subLossXML01ContactDetails.object = msubContactDetails.object
    Set subLossXML01cov.object = msubCOV.object
    Set subLossXML01ed.object = msubED.object
    Set subLossXML01pay.object = msubPAY.object
    Set subLossXML01plh.object = msubPLH.object
    Set subLossXML01cal.object = msubCAL.object
    
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
        'Instance Continued Loss report
        Set moChainReport = mcolChainReports(mlChainCount)
        'Start the daisy linking here
        SetNextChainReport mlChainCount, mcolChainReports
  
        'Set the ref to sub reports in this Report
        Set subLossXML01Chain.object = moChainReport.object
        
    Else
        If Not moChainReport Is Nothing Then
            'Set the ref to sub reports in this Report
            Set subLossXML01Chain.object = moChainReport.object
        End If
    End If
    
       
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Private Sub ActiveReport_ReportStart"
End Sub

Private Sub ActiveReport_Terminate()
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    Set moLists = Nothing
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
    Dim sToc As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim PolicyDetail As udtXML01PolicyDetail
    
    '1 Sub Report Items
    subLossXML01AssDetails.Visible = True
    subLossXML01LossDetails.Visible = True
    subLossXML01ContactDetails.Visible = True
   
    '2 Policy Detail
    With mXML01LossReport.XML01Loss
        PolicyDetail.PolicyNumber = IIf(IsNull(.PolicyDetailRS.getField(1, "PolicyNumber")), vbNullString, .PolicyDetailRS.getField(1, "PolicyNumber"))
        PolicyDetail.Status = IIf(IsNull(.PolicyDetailRS.getField(1, "Status")), vbNullString, .PolicyDetailRS.getField(1, "Status"))
        PolicyDetail.CoverageStatus = IIf(IsNull(.PolicyDetailRS.getField(1, "CoverageStatus")), vbNullString, .PolicyDetailRS.getField(1, "CoverageStatus"))
        PolicyDetail.BalanceDue = IIf(IsNull(.PolicyDetailRS.getField(1, "BalanceDue")), vbNullString, .PolicyDetailRS.getField(1, "BalanceDue"))
        PolicyDetail.CompanyCode = IIf(IsNull(.PolicyDetailRS.getField(1, "CompanyCode")), vbNullString, .PolicyDetailRS.getField(1, "CompanyCode"))
        PolicyDetail.CompanyName = IIf(IsNull(.PolicyDetailRS.getField(1, "CompanyName")), vbNullString, .PolicyDetailRS.getField(1, "CompanyName"))
        PolicyDetail.RenewalDate = IIf(IsNull(.PolicyDetailRS.getField(1, "RenewalDate")), vbNullString, .PolicyDetailRS.getField(1, "RenewalDate"))
        PolicyDetail.CancellationDate = IIf(IsNull(.PolicyDetailRS.getField(1, "CancellationDate")), vbNullString, .PolicyDetailRS.getField(1, "CancellationDate"))
        PolicyDetail.NewBusinessDate = IIf(IsNull(.PolicyDetailRS.getField(1, "NewBusinessDate")), vbNullString, .PolicyDetailRS.getField(1, "NewBusinessDate"))
        PolicyDetail.PolicyDescription = IIf(IsNull(.PolicyDetailRS.getField(1, "PolicyDescription")), vbNullString, .PolicyDetailRS.getField(1, "PolicyDescription"))
        PolicyDetail.MortgageeName = IIf(IsNull(.PolicyDetailRS.getField(1, "MortgageeName")), vbNullString, .PolicyDetailRS.getField(1, "MortgageeName"))
        PolicyDetail.MortgageeAddress = IIf(IsNull(.PolicyDetailRS.getField(1, "MortgageeAddress")), vbNullString, .PolicyDetailRS.getField(1, "MortgageeAddress"))
        
        f_PolicyNumber.Text = PolicyDetail.PolicyNumber
        f_Status.Text = PolicyDetail.Status
        f_CoverageStatus.Text = PolicyDetail.CoverageStatus
        f_BalanceDue.Text = PolicyDetail.BalanceDue
        f_CompanyCode.Text = PolicyDetail.CompanyCode
        f_CompanyName.Text = PolicyDetail.CompanyName
        f_RenewalDate.Text = PolicyDetail.RenewalDate
        f_CancellationDate.Text = PolicyDetail.CancellationDate
        f_NewBusinessDate.Text = PolicyDetail.NewBusinessDate
        f_PolicyDescription.Text = PolicyDetail.PolicyDescription
        f_MortgageeName.Text = PolicyDetail.MortgageeName
        f_MortgageeAddress.Text = PolicyDetail.MortgageeAddress
        
        
    End With

    '3 Sub Report Items
    With mXML01LossReport.XML01Loss
        'Coverage
        If Not .CoverageRS Is Nothing Then
            subLossXML01cov.Visible = True
        End If
        'Endorsements
        If Not .EndorsementRS Is Nothing Then
            subLossXML01ed.Visible = True
        End If
        'Payment Detail
        If Not .PaymentDetailRS Is Nothing Then
            subLossXML01pay.Visible = True
        End If
        'Prior Loss Detail
        If Not .PriorLossDetailRS Is Nothing Then
            subLossXML01plh.Visible = True
        End If
        'Activities
        If Not .ActivitiesRS Is Nothing Then
            subLossXML01cal.Visible = True
        End If
    End With
    
    
CHAINED_REPORTS:
    If Not moChainReport Is Nothing Then
        subLossXML01Chain.Visible = True
        ReportFooter.Visible = True
    Else
        subLossXML01Chain.Visible = False
    End If
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Private Sub Detail_Format"
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
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Public Function ExportME"
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

'For Chained Reports
Public Sub SetNextContReport(plContCount As Long, pcolContReports As Collection)
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    If plContCount + 1 <= pcolContReports.Count Then
        Set pcolContReports(plContCount).ContReport = pcolContReports(plContCount + 1)
        plContCount = plContCount + 1
        'Do daisy again
        pcolContReports(plContCount - 1).SetNextContReport plContCount, pcolContReports
    End If
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Public Sub SetNextContReport"
End Sub



