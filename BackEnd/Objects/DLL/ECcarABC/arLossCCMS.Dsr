VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} arLossCCMS 
   Caption         =   "CCMS Loss Report"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19606
   SectionData     =   "arLossCCMS.dsx":0000
End
Attribute VB_Name = "arLossCCMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private moLists As V2ECKeyBoard.clsCarLists
Private msubED As arsubLossCCMSed 'Endorsemnt sub report
Private msubPLH As arsubLossCCMSplh 'Prior Loss Hist sub report
Private msubCAL As arsubLossCCMScal 'Comments Activity Log sub report

'Continuation Reports
Private mbContFlag As Boolean
Private mlContCount As Long
Private msubCCMSCont As arLossCCMSCont 'CCMS CONTINUED PAGE
Private mCCMSLossReport As CCMSLossReport 'CCMS Loss Repport user defined type

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

Public Property Let ContReport(pvCCMSLossCont As Variant)
    Dim CCMSLossCont As udtCCMSLossCont
    CCMSLossCont = pvCCMSLossCont
    Set msubCCMSCont = New arLossCCMSCont
    msubCCMSCont.CCMSLossCont = CCMSLossCont
End Property

Public Property Let LossReport(pLossReport As V2ECcarABC.CCMSLossReport)
    mCCMSLossReport = pLossReport
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
        Unload subLossCCMSChain
        Set subLossCCMSChain = Nothing
        Set mcolChainReports = Nothing
        Unload moChainReport
        Set moChainReport = Nothing
    End If
    
    'Clean up the Sub Report Objects
    Unload subLossCCMSed.object
    Set subLossCCMSed.object = Nothing
    Unload subLossCCMSplh.object
    Set subLossCCMSplh.object = Nothing
    Unload subLossCCMScal.object
    Set subLossCCMScal.object = Nothing
    Unload subLossCCMSCont.object
    Set subLossCCMSCont.object = Nothing
    Set msubED = Nothing
    Set msubPLH = Nothing
    Set msubCAL = Nothing
    Set msubCCMSCont = Nothing
End Sub

Private Sub ActiveReport_ReportStart()
    On Error GoTo EH
    Dim CCMSLossCont As udtCCMSLossCont
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    'Instance sub reports
    Set msubED = New arsubLossCCMSed
    'Type is Property
    msubED.LossType = CCMSPro
    Set msubPLH = New arsubLossCCMSplh
    Set msubCAL = New arsubLossCCMScal
    'Set their data collections
    Set msubED.EDcol = mCCMSLossReport.CCMSLoss.colEndorsements
    Set msubPLH.PLHcol = mCCMSLossReport.CCMSLoss.colPLH
    Set msubCAL.CALcol = mCCMSLossReport.CCMSLoss.colCAL
    'Set the ref to sub reports in this Report
    Set subLossCCMSed.object = msubED.object
    Set subLossCCMSplh.object = msubPLH.object
    Set subLossCCMScal.object = msubCAL.object
    
    'Set the Continuaton flag if we have any
    If Not mCCMSLossReport.colCCMSLossCont Is Nothing Then
        mbContFlag = True
        mlContCount = 1
    Else
        mbContFlag = False
    End If
    
    'Set the Chain flag if we have any
    If Not mcolChainReports Is Nothing Then
        If Not mbChainFlag Then
            mbChainFlag = True
            mlChainCount = 1
        End If
    Else
        mbChainFlag = False
    End If
    
    'If we have Cont Reports...
    If mbContFlag Then
        'Instance Continued Loss report
        Set msubCCMSCont = New arLossCCMSCont
        CCMSLossCont = mCCMSLossReport.colCCMSLossCont(mlContCount)
        msubCCMSCont.CCMSLossCont = CCMSLossCont
        
        'Start the daisy linking here
        SetNextContReport mlContCount, mCCMSLossReport.colCCMSLossCont
        
        'Set the ref to sub reports in this Report
        Set subLossCCMSCont.object = msubCCMSCont
    Else
        If Not msubCCMSCont Is Nothing Then
            'Set the ref to sub reports in this Report
            Set subLossCCMSCont.object = msubCCMSCont
        End If
    End If
    
    'If we have Chained Reports...
    If mbChainFlag Then
        'Instance Continued Loss report
        Set moChainReport = mcolChainReports(mlChainCount)
        'Start the daisy linking here
        SetNextChainReport mlChainCount, mcolChainReports
  
        'Set the ref to sub reports in this Report
        Set subLossCCMSChain.object = moChainReport
        
    Else
        If Not moChainReport Is Nothing Then
            'Set the ref to sub reports in this Report
            Set subLossCCMSChain.object = moChainReport
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
    
    'Populate all the text fields with the main udt
    '1. Populate ALI
    With mCCMSLossReport.CCMSLoss.AdminLossInfo
        fali0004_DateTimePrinted.Text = .ali0004_DateTimePrinted
        fali0005_PrintedBy.Text = .ali0005_PrintedBy
        fali0052_ReportedBy.Text = .ali0052_ReportedBy
        fali0053_RBPhone.Text = .ali0053_RBPhone
        fali0054_PolicyNum.Text = .ali0054_PolicyNum
        fali0055_SC.Text = .ali0055_SC
        fali0056_AgentNum.Text = .ali0056_AgentNum
        fali0057_HomePhone.Text = .ali0057_HomePhone
        fali0058_BusPhone.Text = .ali0058_BusPhone
        fali0059_NewBusDate.Text = .ali0059_NewBusDate
        fali0060_RenewalDate.Text = .ali0060_RenewalDate
        fali0061_LastCancDate.Text = .ali0061_LastCancDate
        fali0062_NamedInsured.Text = .ali0062_NamedInsured
        fali0063_MailAddress1.Text = .ali0063_MailAddress1
        fali0064_MailAddress2.Text = .ali0064_MailAddress2
        fali0065_MainFileInsuredName.Text = .ali0065_MainFileInsuredName
        fali0066_MortgageHolder.Text = .ali0066_MortgageHolder
        fali0067_2ndMort.Text = .ali0067_2ndMort
        fali0068_CompCode.Text = .ali0068_CompCode
        fali0069_PolicyDescription.Text = .ali0069_PolicyDescription
        fali0070_BldgLimit.Text = .ali0070_BldgLimit
        fali0071_ContLimit.Text = .ali0071_ContLimit
        fali0072_Deductible1.Text = .ali0072_Deductible1
        fali0073_Deductible2.Text = .ali0073_Deductible2
        fali0074_Deductible3.Text = .ali0074_Deductible3
        fali0075_Deductible4.Text = .ali0075_Deductible4
        fali0076_AddlCoverage1.Text = .ali0076_AddlCoverage1
        fali0077_AddlCoverage2.Text = .ali0077_AddlCoverage2
        fali0078_AddlCoverage3.Text = .ali0078_AddlCoverage3
        fali0079_AddlCoverage4.Text = .ali0079_AddlCoverage4
        fali0080_LossLocAddress1.Text = .ali0080_LossLocAddress1
        fali0081_LossLocAddress2.Text = .ali0081_LossLocAddress2
    End With
    
    '2. Populate CLI
    With mCCMSLossReport.CCMSLoss.CurrentLossInfo
        fcli01_CAT.Text = .cli01_CAT
        fcli02_LossDate.Text = .cli02_LossDate
        fcli03_Adjuster.Text = .cli03_Adjuster
        fcli04_DateAsgn.Text = .cli04_DateAsgn
        fcli05_DateClsd.Text = .cli05_DateClsd
        fcli06_SALN.Text = .cli06_SALN
        fcli07_AdjusterOrigInfo.Text = .cli07_AdjusterOrigInfo
        fcli08_DateAsgnOrigInfo = .cli08_DateAsgnOrigInfo
        fcli09_DateClsdOrigInfo.Text = .cli09_DateClsdOrigInfo
        fcli10_PaymentsMadeThisClaim.Text = .cli10_PaymentsMadeThisClaim
        
    '3. Populate Verticle Tab
        lblTabSalnValue.Caption = .cli06_SALN
        lblTabCATValue.Caption = .cli01_CAT
        lblTabADJValue.Caption = .cli03_Adjuster
        
    End With

    
CHAINED_REPORTS:
    If Not msubCCMSCont Is Nothing Then
        subLossCCMSCont.Visible = True
        ReportFooter.Visible = True
    Else
        subLossCCMSCont.Visible = False
        ReportFooter.Visible = False
    End If
    If Not moChainReport Is Nothing Then
        subLossCCMSChain.Visible = True
        ReportFooter.Visible = True
    Else
        subLossCCMSChain.Visible = False
        If msubCCMSCont Is Nothing Then
            ReportFooter.Visible = False
        End If
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


