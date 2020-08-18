VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPaymentRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Type PaymentReq_Info
    fCheckNo As String
    fCheckofA As String
    fCheckofB As String
    fClientClaimNo  As String    'Ass.clientclaimno
    fDatereq As Date          'Print Date
    fDraft As String
    fPayee As String            'If{Checks.IncludeMorgOnDraft}= true then {Assignments.NewInsuredNames} + " " + {Assignments.MortgageCompany} else {Assignments.NewInsuredNames}
    fPayeeAddress As String      'if instr({Assignments.NewMailingAddress},chr(13)) > 1 then left({Assignments.NewMailingAddress},instr({Assignments.NewMailingAddress},chr(13))-1) + " " + mid({Assignments.NewMailingAddress},instr({Assignments.NewMailingAddress},chr(10))+1) else {Assignments.NewMailingAddress}
    fPolicyNo As String          'ass.policynumber
    fReqBy As String             'trim({Assignments.AdjustorFirstName}) + " " + trim({Assignments.AdjustorLastName})
End Type

Private Type PaymentReq_Property
    'Dwelling
    '@LessDedDwelling   'if {@ReqDwellingF} = "" then 0  else
                        'if {?ReqDwelling} - {?Deductible} >= 0 then {?Deductible} else
                        '{?ReqDwelling}
    '@ReqDwellingF      'if {?ReqDwelling} = 0 then "" else totext({?ReqDwelling})
    fDwellLessDed As String      'if {@LessDedDwelling} > 0 then totext({@LessDedDwelling}) else ""
    fDwellReq As String          'if {?ReqDwelling} = 0 then "" else totext({?ReqDwelling})
    fDwellReqRes As String      'if {@ReqDwellingF} = "" then "" else totext({?ReqResDwelling})
    fDwellTotReq As String      'if {?ReqDwelling} - {@LessDedDwelling} > 0 then totext({?ReqDwelling} - {@LessDedDwelling})else ""
    'Other struc
    fOtherStrucLessDed As String
    fOtherStrucReq As String
    fOtherStrucReqRes As String
    fOtherStrucTotReq As String
    'Person prop
    '@LessDedPersonProperty 'if {@ReqPersonPropertyF} = "" then
                            '   0
                            'Else
                            '   if  {@LessDedDwelling} < {?Deductible}  then
                            '       if {?ReqPersonProperty} - ({?Deductible} - {@LessDedDwelling}) >=0 then
                            '           {?Deductible} - {@LessDedDwelling}
                            '       Else
                            '           {?ReqPersonProperty}
                            '   Else
                            '      0
                                    
    '@ReqPersonPropertyF    'if {?ReqPersonProperty} = 0 then "" else totext({?ReqPersonProperty})
    fPersonLessDed As String          'if {@LessDedPersonProperty} > 0 then totext({@LessDedPersonProperty}) else ""
    fPersonReq As String              'if {?ReqPersonProperty} = 0 then "" else totext({?ReqPersonProperty})
    fPersonReqRes As String           'if {@ReqPersonPropertyF} = "" then "" else totext({?ReqResPersonProperty})
    fPersonTotReq As String           'if {?ReqPersonProperty} - {@LessDedPersonProperty} > 0 then totext({?ReqPersonProperty} - {@LessDedPersonProperty}) else ""
    'Loss of use

    '@LessDedLossOfUse      'if {@ReqLossOfUseF} = "" then 0 else
                            'if ({@LessDedPersonProperty} + {@LessDedDwelling}) < {?Deductible}  then {?Deductible} - ({@LessDedPersonProperty} + {@LessDedDwelling}) else 0
    '@ReqLossOfUseF         'if {?ReqLossOfUse} = 0 then "" else totext({?ReqLossOfUse})
    fLossOfUseLessDed As String       'if {@LessDedLossOfUse} > 0 then totext({@LessDedLossOfUse}) else ""
    fLossOfUseReq As String           'if {?ReqLossOfUse} = 0 then "" else totext({?ReqLossOfUse})
    fLossOfUseReqRes As String        'if {@ReqLossOfUseF} = "" then "" else totext({?ReqResLossOfUse})
    fLossOfUseTotReq As String        'if {?ReqLossOfUse} - {@LessDedLossOfUse} > 0 then totext({?ReqLossOfUse} - {@LessDedLossOfUse}) else ""
    'Building
    fBuildingLessDed As String
    fBuildingReq As String
    fBuildingReqRes As String
    fBuildingTotReq As String
    'Bus-Interup
    fBusIntLessDed As String
    fBusIntReq As String
    fBusIntReqRes As String
    fBusIntTotReq As String
End Type

Private Type PaymentReq_Auto
    'CPR
    fCPRLessDed As String
    fCPRReq As String
    fCPRReqRes As String
    fCPRTotReq As String
    'LOU
    fLOULessDed As String
    fLOUReq As String
    fLOUReqRes As String
    fLOUTotReq As String
    'Other
    fOtherLessDed As String
    fOtherReq As String
    fOtherReqRes As String
    fOtherTotReq As String
End Type

Private Type PaymentReq_CommentInfo
    fComments As String               'Checks.Instructions
    fFileClosePayAppNo As String
    fFileClosePayAppYes As String
    fInPayOfDescription As String    'Indemnity.Description
    fPresGuarIssuesNo As String
    fPresGuarIssuesYes As String
End Type

Private Type udtPaymentReq_Safeco
    'I. Payment Information
    PaymentInfo As PaymentReq_Info
    'II Property
    Property As PaymentReq_Property
    'III Auto
    Auto As PaymentReq_Auto
    'IV Comment Info
    CommentInfo As PaymentReq_CommentInfo
End Type

Private mCurDB As DAO.Database
Private mConn As ADODB.Connection
Private mRS As ADODB.Recordset
Private mARV As V2ARViewer.clsARViewer
Private mbApplyDeductibleToLossOfUse As Boolean

Public Property Let CurDB(pDB As DAO.Database)
    Set mCurDB = pDB
End Property
Public Property Set CurDB(pDB As DAO.Database)
    Set mCurDB = pDB
End Property
Public Property Get CurDB() As DAO.Database
    Set CurDB = mCurDB
End Property

Public Property Let ARV(poARV As V2ARViewer.clsARViewer)
    Set mARV = poARV
End Property
Public Property Set ARV(poARV As V2ARViewer.clsARViewer)
    Set mARV = poARV
End Property
Public Property Get ARV() As V2ARViewer.clsARViewer
    Set ARV = mARV
End Property

Public Property Let ApplyDeducibleToLossOfUse(pbFlag As Boolean)
    mbApplyDeductibleToLossOfUse = pbFlag
End Property
Public Property Get ApplyDeducibleToLossOfUse() As Boolean
    ApplyDeducibleToLossOfUse = mbApplyDeductibleToLossOfUse
End Property

Public Property Get ClassName() As String
    ClassName = "V2ECcarSafeco.clsPaymentRequest"
End Property

Public Function PrintPaymentReq(psClaimNo As String, psCheckNum As String, _
                                Optional psCopyName As String, _
                                Optional pbPreview As Boolean, _
                                Optional pbShowPrintButton As Boolean, _
                                Optional psXportPath As String, _
                                Optional pXportType As ExportType) As Boolean
    On Error GoTo EH
    Dim MyPaymentReq As V2ECcarSafeco.arRptPaymentReq
    Dim CarrierLists As Object
    Dim sCarrier As String
    Dim udtPaymentReq As udtPaymentReq_Safeco
    Dim oLRs As V2ECKeyBoard.clsLossReports
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    Set MyPaymentReq = New V2ECcarSafeco.arRptPaymentReq
    'See if we are not using using Adjuster DB need
    'to open connection to WEB DB
    If mCurDB Is Nothing Then
        OpenConnection
    End If
    
    udtPaymentReq = PopulateUDTPaymentReq_Safeco(psClaimNo, psCheckNum)
    
    With udtPaymentReq.PaymentInfo
        MyPaymentReq.SetProperty "fClientClaimNo", .fClientClaimNo, vbString
        MyPaymentReq.SetProperty "fDatereq", .fDatereq, vbDate
        MyPaymentReq.SetProperty "fPayee", .fPayee, vbString
        MyPaymentReq.SetProperty "fPayeeAddress", .fPayeeAddress, vbString
        MyPaymentReq.SetProperty "fPolicyNo", .fPolicyNo, vbString
        MyPaymentReq.SetProperty "fReqBy", .fReqBy, vbString
    End With
    
    With udtPaymentReq.Property
        If .fDwellLessDed = vbNullString Then
            MyPaymentReq.SetProperty "fDwellLessDed", .fDwellLessDed, vbString
        Else
            MyPaymentReq.SetProperty "fDwellLessDed", CCur(.fDwellLessDed), vbCurrency
        End If
        If .fDwellReq = vbNullString Then
            MyPaymentReq.SetProperty "fDwellReq", .fDwellReq, vbString
        Else
            MyPaymentReq.SetProperty "fDwellReq", CCur(.fDwellReq), vbCurrency
        End If
        If .fDwellReqRes = vbNullString Then
            MyPaymentReq.SetProperty "fDwellReqRes", .fDwellReqRes, vbString
        Else
            MyPaymentReq.SetProperty "fDwellReqRes", CCur(.fDwellReqRes), vbCurrency
        End If
        If .fBusIntTotReq = vbNullString Then
            MyPaymentReq.SetProperty "fDwellTotReq", .fDwellTotReq, vbString
        Else
            MyPaymentReq.SetProperty "fDwellTotReq", CCur(.fDwellTotReq), vbCurrency
        End If
        
        If .fPersonLessDed = vbNullString Then
            MyPaymentReq.SetProperty "fPersonLessDed", .fPersonLessDed, vbString
        Else
            MyPaymentReq.SetProperty "fPersonLessDed", CCur(.fPersonLessDed), vbCurrency
        End If
        If .fPersonReq = vbNullString Then
            MyPaymentReq.SetProperty "fPersonReq", .fPersonReq, vbString
        Else
            MyPaymentReq.SetProperty "fPersonReq", CCur(.fPersonReq), vbCurrency
        End If
        If .fPersonReqRes = vbNullString Then
            MyPaymentReq.SetProperty "fPersonReqRes", .fPersonReqRes, vbString
        Else
            MyPaymentReq.SetProperty "fPersonReqRes", CCur(.fPersonReqRes), vbCurrency
        End If
        If .fPersonTotReq = vbNullString Then
            MyPaymentReq.SetProperty "fPersonTotReq", .fPersonTotReq, vbString
        Else
            MyPaymentReq.SetProperty "fPersonTotReq", CCur(.fPersonTotReq), vbCurrency
        End If
        
        If .fLossOfUseLessDed = vbNullString Then
            MyPaymentReq.SetProperty "fLossOfUseLessDed", .fLossOfUseLessDed, vbString
        Else
            MyPaymentReq.SetProperty "fLossOfUseLessDed", CCur(.fLossOfUseLessDed), vbCurrency
        End If
        If .fLossOfUseReq = vbNullString Then
            MyPaymentReq.SetProperty "fLossOfUseReq", .fLossOfUseReq, vbString
        Else
            MyPaymentReq.SetProperty "fLossOfUseReq", CCur(.fLossOfUseReq), vbCurrency
        End If
        If .fLossOfUseReqRes = vbNullString Then
            MyPaymentReq.SetProperty "fLossOfUseReqRes", .fLossOfUseReqRes, vbString
        Else
            MyPaymentReq.SetProperty "fLossOfUseReqRes", CCur(.fLossOfUseReqRes), vbCurrency
        End If
        If .fLossOfUseTotReq = vbNullString Then
            MyPaymentReq.SetProperty "fLossOfUseTotReq", .fLossOfUseTotReq, vbString
        Else
            MyPaymentReq.SetProperty "fLossOfUseTotReq", CCur(.fLossOfUseTotReq), vbCurrency
        End If
    End With
    '8.30.2002 not updating anything in Auto at this time
'    With udtPaymentReq.Auto
'
'    End With
    
    With udtPaymentReq.CommentInfo
        MyPaymentReq.SetProperty "fComments", .fComments, vbString
        MyPaymentReq.SetProperty "fInPayOfDescription", .fInPayOfDescription, vbString
    End With
    
    With MyPaymentReq
        If Not mCurDB Is Nothing Then
            '190  Printing on HP1218  5.15.2002
            'Need to use WIN default printer setting if it is selected under the menu
            On Error Resume Next
            .Printer.DeviceName = GetSetting(goUtil.gsAppEXEName, "PRINTER", "PRINTER_NAME", vbNullString)
            If Err.Number > 0 Then
                Err.Clear
            End If
            .Printer.Port = GetSetting(goUtil.gsAppEXEName, "PRINTER", "PRINTER_PORT", vbNullString)
            If Err.Number > 0 Then
                Err.Clear
            End If
        End If
        On Error GoTo EH
        .Printer.Orientation = ddOPortrait
    End With
    
    'Adjuster DB
    If Not mCurDB Is Nothing Then
        If pbPreview Then
            If mARV Is Nothing Then
                Set mARV = New V2ARViewer.clsARViewer
            End If
            
            With mARV
                'Pass in true to have Active reports process on separate thread.
                'This will allow the viewer to load while the report is processing
                MyPaymentReq.Run True
                .objARvReport = MyPaymentReq
                .sRptTitle = psCheckNum
                .HidePrintButton = Not pbShowPrintButton
                .ShowReport vbModeless
            End With
        Else
            MyPaymentReq.Run True
        End If
    
    
        'Need to wait for the report to becom active
        Do
            If MyPaymentReq.ActiveFlag Then
                If Not pbPreview Then
                    MyPaymentReq.PrintReport False 'Don't show print setup dialog
                End If
                Exit Do
            End If
            DoEvents
            Sleep 500
        Loop
        PrintPaymentReq = True
    Else   'Web DB
        Set oLRs = New V2ECKeyBoard.clsLossReports
        If oLRs.ExportFile(MyPaymentReq, psXportPath, pXportType) Then
            PrintPaymentReq = True
        End If
        oLRs.CleanUp
        Set oLRs = Nothing
    End If
    
    'Clean up
    Unload MyPaymentReq
    Set MyPaymentReq = Nothing
    Set CarrierLists = Nothing
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Screen.MousePointer = vbDefault
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Public Function PrintPaymentReq"
End Function

Private Function PopulateUDTPaymentReq_Safeco(psClaimNo As String, psCheckNum As String) As udtPaymentReq_Safeco
    On Error GoTo EH
    Dim RS As DAO.Recordset
    Dim sSQL As String
    Dim sDescription As String
    Dim cDeductible As Currency
    Dim cAppliedDeductible As Currency
    Dim cReqDwelling As Currency
    Dim cReqResDwelling As Currency
    Dim cReqPersonProperty As Currency
    Dim cReqResPersonProperty As Currency
    Dim cReqLossOfUse As Currency
    Dim cReqResLossOfUse As Currency
    Dim cTemp As Currency
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    'Adjuster Access DB
    If Not mCurDB Is Nothing Then
        'BGS 10.12.2001 Need to populate the IB Safeco udt with records for the
        'Claimno and Checknum Provided
        sSQL = "SELECT A.ClientClaimNo, "
        sSQL = sSQL & "A.PolicyNumber, "
        sSQL = sSQL & "Trim(A.AdjustorFirstName) & ' ' & Trim(A.AdjustorLastName) As ReqBy, "
        sSQL = sSQL & "Trim(A.NewMailingAddress) As PayeeAddress, "
        sSQL = sSQL & "IIF(B.IncludeMorgOnDraft = True, A.NewInsuredNames & ' ' & A.MortgageCompany, A.NewInsuredNames) As Payee, "
        sSQL = sSQL & "B.Instructions As Comments, "
        sSQL = sSQL & "A.Deductible "
        sSQL = sSQL & "FROM Assignments As A LEFT JOIN Checks As B "
        sSQL = sSQL & "ON A.ClaimNo = B.ClaimNo "
        sSQL = sSQL & "WHERE A.ClaimNo = '" & goUtil.utCleanSQLString(psClaimNo) & "' "
        sSQL = sSQL & "AND B.CheckNum = '" & goUtil.utCleanSQLString(psCheckNum) & "' "
        
        Set RS = mCurDB.OpenRecordset(sSQL)
        
        With PopulateUDTPaymentReq_Safeco.PaymentInfo
            If Not RS.EOF Then
                RS.MoveFirst
                .fClientClaimNo = IIf(IsNull(RS!ClientClaimNo), vbNullString, RS!ClientClaimNo)
                .fDatereq = Format(Now, "MM/DD/YY")
                .fPayee = IIf(IsNull(RS!Payee), vbNullString, RS!Payee)
                .fPayeeAddress = IIf(IsNull(RS!PayeeAddress), vbNullString, RS!PayeeAddress)
                .fPolicyNo = IIf(IsNull(RS!PolicyNumber), vbNullString, RS!PolicyNumber)
                .fReqBy = IIf(IsNull(RS!ReqBy), vbNullString, RS!ReqBy)
                'Get Deductible here
                cDeductible = IIf(IsNull(RS!Deductible), 0, RS!Deductible)
                PopulateUDTPaymentReq_Safeco.CommentInfo.fComments = IIf(IsNull(RS!Comments), vbNullString, RS!Comments)
            End If
        End With
        RS.Close
        Set RS = Nothing
    Else ' Web DB
        sSQL = "SELECT RTSAFECO.SCLIENTCLA AS ClientClaimNo, "
        sSQL = sSQL & "RTSAFECO.RT09A_S AS PolicyNumber, "
        sSQL = sSQL & "RTSAFECO.RT62_S  AS ReqBy, "
        sSQL = sSQL & "RTSAFECO.RT52_S AS PayeeAddress, "
        sSQL = sSQL & "RTSAFECO.RT51_S + ' ' + RTSAFECO.RT50_S AS Payee, "
        sSQL = sSQL & "RTSAFECO.RT61_S AS Comments, "
        sSQL = sSQL & "RTSAFECO.RT15_C AS Deductible, "
        sSQL = sSQL & "RTSAFECO.RT63_DT AS DateReq "
        sSQL = sSQL & "FROM RTSAFECO "
        sSQL = sSQL & "WHERE LEFT(RTSAFECO.SIBNUMBER," & Len(psClaimNo) & ") = '" & goUtil.utCleanSQLString(psClaimNo) & "' "
        sSQL = sSQL & "AND RTSAFECO.RT03_S = '" & goUtil.utCleanSQLString(psCheckNum) & "' "
        Set mRS = New ADODB.Recordset
        mRS.CursorLocation = adUseClient
        mRS.Open sSQL, mConn, adOpenStatic, adLockReadOnly
        Set mRS.ActiveConnection = Nothing
        
        With PopulateUDTPaymentReq_Safeco.PaymentInfo
            If Not mRS.EOF Then
                mRS.MoveLast
                .fClientClaimNo = IIf(IsNull(mRS!ClientClaimNo), vbNullString, Trim(mRS!ClientClaimNo))
                .fDatereq = IIf(IsNull(mRS!DateReq), NULL_DATE, Format(mRS!DateReq, "MM/DD/YY"))
                .fPayee = IIf(IsNull(mRS!Payee), vbNullString, Trim(mRS!Payee))
                .fPayeeAddress = IIf(IsNull(mRS!PayeeAddress), vbNullString, Trim(mRS!PayeeAddress))
                .fPolicyNo = IIf(IsNull(mRS!PolicyNumber), vbNullString, Trim(mRS!PolicyNumber))
                .fReqBy = IIf(IsNull(mRS!ReqBy), vbNullString, Trim(mRS!ReqBy))
                'Get Deductible here
                cDeductible = IIf(IsNull(mRS!Deductible), 0, mRS!Deductible)
                PopulateUDTPaymentReq_Safeco.CommentInfo.fComments = IIf(IsNull(mRS!Comments), vbNullString, Trim(mRS!Comments))
            End If
        End With
        mRS.Close
        Set mRS = Nothing
    End If
    
    'Need to Subtract out any deductible already applied to previous
    'payment request, if the current payment request is a supplement
    If InStr(1, psCheckNum, "(01)", vbBinaryCompare) = 0 Then
        PopulateProperty PopulateUDTPaymentReq_Safeco, psClaimNo, psCheckNum, cDeductible, True, cAppliedDeductible
        cDeductible = cDeductible - cAppliedDeductible
        If cDeductible < 0 Then
            cDeductible = 0
        End If
    End If
        
    PopulateProperty PopulateUDTPaymentReq_Safeco, psClaimNo, psCheckNum, cDeductible, False, cAppliedDeductible
    
    'Adjuster Access DB
    If Not mCurDB Is Nothing Then
        'BGS 6.11.2002 Need to get the Description of Loss
        sSQL = "SELECT I.Description "
        sSQL = sSQL & "FROM Indemnity As I "
        sSQL = sSQL & "LEFT JOIN Assignments As A "
        sSQL = sSQL & "ON A.ClaimNo = I.ClaimNo "
        sSQL = sSQL & "WHERE A.ClaimNo = '" & goUtil.utCleanSQLString(psClaimNo) & "' "
        sSQL = sSQL & "AND I.IndemType = '" & goUtil.utCleanSQLString(psCheckNum) & "' "
        Set RS = mCurDB.OpenRecordset(sSQL)
        If Not RS.EOF Then
            Do Until RS.EOF
                If Not IsNull(RS!Description) Then
                    sDescription = sDescription & RS!Description & " "
                End If
                RS.MoveNext
            Loop
        End If
        Set RS = Nothing
    Else ' Web DB
        
        sSQL = "SELECT RTSAFECOINDEM.SDESCRIPTI AS Description "
        sSQL = sSQL & "FROM RTSAFECOINDEM "
        sSQL = sSQL & "WHERE RTSAFECOINDEM.SIBNUMBER = '" & goUtil.utCleanSQLString(psClaimNo) & "' "
        sSQL = sSQL & "AND RTSAFECOINDEM.SINDEMTYPE = '" & goUtil.utCleanSQLString(psCheckNum) & "' "
        Set mRS = New ADODB.Recordset
        mRS.CursorLocation = adUseClient
        mRS.Open sSQL, mConn, adOpenStatic, adLockReadOnly
        Set mRS.ActiveConnection = Nothing
        If Not mRS.EOF Then
            Do Until mRS.EOF
                If Not IsNull(mRS!Description) Then
                    sDescription = sDescription & Trim(mRS!Description) & " "
                End If
                mRS.MoveNext
            Loop
        End If
        mRS.Close
        Set mRS = Nothing
    End If
    
    With PopulateUDTPaymentReq_Safeco.CommentInfo
        .fInPayOfDescription = sDescription
    End With
    
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Set RS = Nothing
    Set mRS = Nothing
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Private Function PopulateUDTPaymentReq_Safeco"
End Function

Private Function PopulateProperty(PR_Safeco As udtPaymentReq_Safeco, psClaimNo As String, psCheckNum As String, pcDeductible As Currency, _
                                  pbGetAppliedDeductible As Boolean, pcAppliedDeductible As Currency) As Boolean
    On Error GoTo EH
    Dim RS As DAO.Recordset
    Dim sSQL As String
    Dim cReqDwelling As Currency
    Dim cReqResDwelling As Currency
    Dim cReqPersonProperty As Currency
    Dim cReqResPersonProperty As Currency
    Dim cReqLossOfUse As Currency
    Dim cReqResLossOfUse As Currency
    Dim cTemp As Currency
    Dim cUnderDeductible As Currency
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    PopulateProperty = True
    'Use Adjuster DB
    If Not mCurDB Is Nothing Then
        'Get Req/ReqRes Dwelling,PersonProperty,LossofUse
        sSQL = "SELECT I.RecoverableDepreciation As ReqRes, "
        sSQL = sSQL & "I.ACVLessExcessLimit As Req , "
        sSQL = sSQL & "I.Class "
        sSQL = sSQL & "FROM Indemnity As I "
        sSQL = sSQL & "LEFT JOIN Assignments As A "
        sSQL = sSQL & "ON A.ClaimNo = I.ClaimNo "
        sSQL = sSQL & "WHERE A.ClaimNo = '" & goUtil.utCleanSQLString(psClaimNo) & "' "
        If pbGetAppliedDeductible Then
            sSQL = sSQL & "AND I.IndemType <> '" & goUtil.utCleanSQLString(psCheckNum) & "' "
        Else
            sSQL = sSQL & "AND I.IndemType = '" & goUtil.utCleanSQLString(psCheckNum) & "' "
        End If
        Set RS = mCurDB.OpenRecordset(sSQL)
        
        If Not RS.EOF Then
            Do Until RS.EOF
                Select Case IIf(IsNull(RS!Class), vbNullString, Trim(RS!Class))
                    Case FC_BUILD & " - Dwelling"
                        cReqDwelling = cReqDwelling + RS!Req
                        cReqResDwelling = cReqResDwelling + RS!ReqRes
                    Case FC_CONTENTS & " - Personal Property"
                        cReqPersonProperty = cReqPersonProperty + RS!Req
                        cReqResPersonProperty = cReqResPersonProperty + RS!ReqRes
                    Case FC_ALE & " - Loss of Use"
                        cReqLossOfUse = cReqLossOfUse + RS!Req
                        cReqResLossOfUse = cReqResLossOfUse + RS!ReqRes
                End Select
                RS.MoveNext
            Loop
        End If
        Set RS = Nothing
    Else 'Use WEB DB
        'Get Req/ReqRes Dwelling,PersonProperty,LossofUse
        sSQL = "SELECT RTSAFECOINDEM.CRECOVERAB As ReqRes, "
        sSQL = sSQL & "RTSAFECOINDEM.CACVLESSEX As Req, "
        sSQL = sSQL & "RTRIM(RTSAFECOINDEM.SCLASS) AS Class "
        sSQL = sSQL & "FROM RTSAFECOINDEM "
        sSQL = sSQL & "WHERE RTSAFECOINDEM.SIBNUMBER = '" & goUtil.utCleanSQLString(psClaimNo) & "' "
        If pbGetAppliedDeductible Then
            sSQL = sSQL & "AND RTSAFECOINDEM.SINDEMTYPE <> '" & goUtil.utCleanSQLString(psCheckNum) & "' "
        Else
            sSQL = sSQL & "AND RTSAFECOINDEM.SINDEMTYPE = '" & goUtil.utCleanSQLString(psCheckNum) & "' "
        End If
        Set mRS = New ADODB.Recordset
        mRS.CursorLocation = adUseClient
        mRS.Open sSQL, mConn, adOpenStatic, adLockReadOnly
        Set mRS.ActiveConnection = Nothing
        
        If Not mRS.EOF Then
            Do Until mRS.EOF
                Select Case IIf(IsNull(mRS!Class), vbNullString, Trim(mRS!Class))
                    Case FC_BUILD & " - Dwelling"
                        cReqDwelling = cReqDwelling + mRS!Req
                        cReqResDwelling = cReqResDwelling + mRS!ReqRes
                    Case FC_CONTENTS & " - Personal Property"
                        cReqPersonProperty = cReqPersonProperty + mRS!Req
                        cReqResPersonProperty = cReqResPersonProperty + mRS!ReqRes
                    Case FC_ALE & " - Loss of Use"
                        cReqLossOfUse = cReqLossOfUse + mRS!Req
                        cReqResLossOfUse = cReqResLossOfUse + mRS!ReqRes
                End Select
                mRS.MoveNext
            Loop
        End If
        mRS.Close
        Set mRS = Nothing
    End If
    
    'Need to calculate the Dwelling,PersonProperty,LossofUse
    With PR_Safeco.Property
        'DWELLING
        'if {@LessDedDwelling} > 0 then totext({@LessDedDwelling}) else ""
        cTemp = LessDedDwelling(cReqDwelling, pcDeductible)
        If cTemp > 0 Then
            .fDwellLessDed = cTemp
        Else
            .fDwellLessDed = vbNullString
        End If
        'if {?ReqDwelling} = 0 then "" else totext({?ReqDwelling})
        If cReqDwelling = 0 Then
            .fDwellReq = vbNullString
        Else
            .fDwellReq = cReqDwelling
        End If
        'if {@ReqDwellingF} = "" then "" else totext({?ReqResDwelling})
        If ReqDwellingF(cReqDwelling) = vbNullString Then
            .fDwellReqRes = vbNullString
        Else
            .fDwellReqRes = cReqResDwelling
        End If
        'if {?ReqDwelling} - {@LessDedDwelling} > 0 then totext({?ReqDwelling} - {@LessDedDwelling})else ""
        cTemp = (cReqDwelling - LessDedDwelling(cReqDwelling, pcDeductible))
        If cTemp > 0 Then
            .fDwellTotReq = cTemp
        Else
            .fDwellTotReq = vbNullString
        End If
        
        'PERSONAL PROPERTY
        'if {@LessDedPersonProperty} > 0 then totext({@LessDedPersonProperty}) else ""
        cTemp = LessDedPersonProperty(pcDeductible, cReqPersonProperty, cReqDwelling, cUnderDeductible)
        If cTemp > 0 Then
            .fPersonLessDed = cTemp
        Else
            .fPersonLessDed = vbNullString
        End If
        'if {?ReqPersonProperty} = 0 then "" else totext({?ReqPersonProperty})
        If cReqPersonProperty = 0 Then
            .fPersonReq = vbNullString
        Else
            .fPersonReq = cReqPersonProperty
        End If
        'if {@ReqPersonPropertyF} = "" then "" else totext({?ReqResPersonProperty})
        If ReqPersonPropertyF(cReqPersonProperty) = vbNullString Then
            .fPersonReqRes = vbNullString
        Else
            .fPersonReqRes = cReqResPersonProperty
        End If
        'if {?ReqPersonProperty} - {@LessDedPersonProperty} > 0 then totext({?ReqPersonProperty} - {@LessDedPersonProperty}) else ""
        cTemp = (cReqPersonProperty - LessDedPersonProperty(pcDeductible, cReqPersonProperty, cReqDwelling))
        If cTemp > 0 Then
            .fPersonTotReq = cTemp
        Else
            .fPersonTotReq = vbNullString
        End If
        
        'LOSS OF USE
        'if {@LessDedLossOfUse} > 0 then totext({@LessDedLossOfUse}) else ""
        cTemp = LessDedLossOfUse(cReqLossOfUse, pcDeductible, cReqPersonProperty, cReqDwelling, cUnderDeductible)
        If cTemp > 0 Then
            .fLossOfUseLessDed = cTemp
        Else
            .fLossOfUseLessDed = vbNullString
        End If
        'if {?ReqLossOfUse} = 0 then "" else totext({?ReqLossOfUse})
        If cReqLossOfUse = 0 Then
            .fLossOfUseReq = vbNullString
        Else
            .fLossOfUseReq = cReqLossOfUse
        End If
        'if {@ReqLossOfUseF} = "" then "" else totext({?ReqResLossOfUse})
        If ReqLossOfUseF(cReqLossOfUse) = vbNullString Then
            .fLossOfUseReqRes = vbNullString
        Else
            .fLossOfUseReqRes = cReqResLossOfUse
        End If
        'if {?ReqLossOfUse} - {@LessDedLossOfUse} > 0 then totext({?ReqLossOfUse} - {@LessDedLossOfUse}) else ""
        cTemp = (cReqLossOfUse - LessDedLossOfUse(cReqLossOfUse, pcDeductible, cReqPersonProperty, cReqDwelling))
        If cTemp > 0 Then
            .fLossOfUseTotReq = cTemp
        Else
            .fLossOfUseTotReq = vbNullString
        End If
        
        'Check to see if we are populating Applied Deductible
        If pbGetAppliedDeductible Then
            If .fDwellLessDed = vbNullString Then
                .fDwellLessDed = 0
            End If
            If .fPersonLessDed = vbNullString Then
                .fPersonLessDed = 0
            End If
            If .fLossOfUseLessDed = vbNullString Then
                .fLossOfUseLessDed = 0
            End If
            pcAppliedDeductible = (CCur(.fDwellLessDed) + CCur(.fPersonLessDed) + CCur(.fLossOfUseLessDed)) - cUnderDeductible
        End If
        
    End With
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    PopulateProperty = False
    Set RS = Nothing
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Private Function PopulateProperty"
End Function

Private Function ReqDwellingF(pcReqDwelling As Currency) As String
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    '@ReqDwellingF      'if {?ReqDwelling} = 0 then "" else totext({?ReqDwelling})
    If pcReqDwelling = 0 Then
        ReqDwellingF = vbNullString
    Else
        ReqDwellingF = CStr(pcReqDwelling)
    End If
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Private Function ReqDwellingF"
End Function

Private Function LessDedDwelling(pcReqDwelling As Currency, pcDeductible As Currency) As Currency
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    '@LessDedDwelling   'if {@ReqDwellingF} = "" then 0  else
                        'if {?ReqDwelling} - {?Deductible} >= 0 then {?Deductible} else
                        '{?ReqDwelling}
    If ReqDwellingF(pcReqDwelling) = vbNullString Then
        LessDedDwelling = 0
    Else
        If pcReqDwelling - pcDeductible >= 0 Then
            LessDedDwelling = pcDeductible
        Else
            LessDedDwelling = pcReqDwelling
        End If
    End If
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Private Function LessDedDwelling"
End Function

Private Function ReqPersonPropertyF(pcReqPersonProperty As Currency) As String
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    '@ReqPersonPropertyF    'if {?ReqPersonProperty} = 0 then "" else totext({?ReqPersonProperty})
    If pcReqPersonProperty = 0 Then
        ReqPersonPropertyF = vbNullString
    Else
        ReqPersonPropertyF = CStr(pcReqPersonProperty)
    End If
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Private Function ReqPersonPropertyF"
End Function

Private Function LessDedPersonProperty(pcDeductible As Currency, pcReqPersonProperty As Currency, pcReqDwelling As Currency, Optional pcUnderDeductible As Currency = Empty) As Currency
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    '@LessDedPersonProperty 'if {@ReqPersonPropertyF} = "" then
                            '   0
                            'Else
                            '   if  {@LessDedDwelling} < {?Deductible}  then
                            '       if {?ReqPersonProperty} - ({?Deductible} - {@LessDedDwelling}) >=0 then
                            '           {?Deductible} - {@LessDedDwelling}
                            '       Else
                            '           {?ReqPersonProperty}
                            '   Else
                            '      0
    If ReqPersonPropertyF(pcReqPersonProperty) = vbNullString Then
        LessDedPersonProperty = 0
    Else
        If LessDedDwelling(pcReqDwelling, pcDeductible) < pcDeductible Then
            If pcReqPersonProperty - (pcDeductible - LessDedDwelling(pcReqDwelling, pcDeductible)) >= 0 Then
                LessDedPersonProperty = pcDeductible - LessDedDwelling(pcReqDwelling, pcDeductible)
                If Not IsEmpty(pcUnderDeductible) Then
                    If LessDedPersonProperty > pcReqPersonProperty Then
                        pcUnderDeductible = LessDedPersonProperty - pcReqPersonProperty
                    End If
                End If
            Else
                LessDedPersonProperty = pcReqPersonProperty
            End If
        Else
            LessDedPersonProperty = 0
        End If
    End If
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Private Function LessDedPersonProperty"
End Function

Private Function LessDedLossOfUse(pcReqLossOfUse As Currency, pcDeductible As Currency, pcReqPersonProperty As Currency, pcReqDwelling As Currency, Optional pcUnderDeductible As Currency = Empty) As Currency
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
'@LessDedLossOfUse      'if {@ReqLossOfUseF} = "" then 0 else
'                            'if ({@LessDedPersonProperty} + {@LessDedDwelling}) < {?Deductible}  then {?Deductible} - ({@LessDedPersonProperty} + {@LessDedDwelling}) else 0
    If ReqLossOfUseF(pcReqLossOfUse) = vbNullString Then
        LessDedLossOfUse = 0
    Else
        If (LessDedPersonProperty(pcDeductible, pcReqPersonProperty, pcReqDwelling) + LessDedDwelling(pcReqDwelling, pcDeductible)) < pcDeductible And ApplyDeducibleToLossOfUse Then
            LessDedLossOfUse = pcDeductible - (LessDedPersonProperty(pcDeductible, pcReqPersonProperty, pcReqDwelling) + LessDedDwelling(pcReqDwelling, pcDeductible))
            If Not IsEmpty(pcUnderDeductible) Then
                If LessDedLossOfUse > pcReqLossOfUse Then
                    pcUnderDeductible = LessDedLossOfUse - pcReqLossOfUse
                End If
            End If
        Else
            LessDedLossOfUse = 0
        End If
    End If
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Private Function LessDedLossOfUse"
End Function

Private Function ReqLossOfUseF(pcReqLossOfUse As Currency) As String
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
'    '@ReqLossOfUseF         'if {?ReqLossOfUse} = 0 then "" else totext({?ReqLossOfUse})
    If pcReqLossOfUse = 0 Then
        ReqLossOfUseF = vbNullString
    Else
        ReqLossOfUseF = CStr(pcReqLossOfUse)
    End If
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Private Function ReqLossOfUseF"
End Function

Private Function OpenConnection() As Boolean
    On Error GoTo EH
    Dim sUserID As String
    Dim sPassword As String
    Dim sProdDSN As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    OpenConnection = True
    
    'ADO Object
    If mConn Is Nothing Then
        sUserID = goUtil.utGetECSCryptSetting("V2WebControl", "DBConn", "USERID")
        sPassword = goUtil.utGetECSCryptSetting("V2WebControl", "DBConn", "PASSWORD")
        sProdDSN = GetSetting("V2WebControl", "DSN", "NAME", vbNullString)
        Set mConn = New ADODB.Connection
        mConn.Open sProdDSN, sUserID, sPassword
    End If
            
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    OpenConnection = False
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Private Function OpenConnection"
End Function

Public Function CleanUp() As Boolean
    On Error Resume Next
    If Not mRS Is Nothing Then
        mRS.Close
        Set mRS = Nothing
    End If
    If Not mConn Is Nothing Then
        mConn.Close
        Set mConn = Nothing
    End If
    'Only set this ref to nothing (Don't want to close Adjuster's current DB session)
    If Not mCurDB Is Nothing Then
        Set mCurDB = Nothing
    End If
    'Only set this ref to nothing (Don't want to close Adjuster's current gARV)
    If Not mARV Is Nothing Then
        Set mARV = Nothing
    End If
End Function

Public Function SetUtilObject(oUtilObject As Object) As Boolean
    On Error Resume Next
    Set goUtil = oUtilObject
    If Err.Number = 0 Then
        SetUtilObject = True
    Else
        SetUtilObject = False
    End If
End Function

