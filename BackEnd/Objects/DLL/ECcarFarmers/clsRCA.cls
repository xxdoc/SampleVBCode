VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRCA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


'Line Items
'Line 01 Policy Limits
Private mcDwellPolicyLim As Currency         '{Assignments.BuildingPolicyLimit}/{rtfarmer.CI14_cBu}
Private mcContPolicyLim As Currency          '{Assignments.PersonalPropertyLimit}/{rtfarmer.CI15_cCo}
'Line 02 Full Replacement Cost of the said property at the time of the loss was
Private mcDwellRCSaid As Currency            '{Assignments.RCSaidBuilding}/{rtfarmer.CI17_cBu}
Private mcContRCSaid As Currency             '{Assignments.RCSaidPersonal}/{rtfarmer.CI18_cCo}
'Line 03 The Full Cost of Repair or Replacement is
Private mcDwellFullRC As Currency            '{indemnity.ReplacementCost} where {left(indemnity.class,2) ='01'}/{rtfarmerindem.cReplaceme} where {left(rtfarmerindem.sClass,2) ='01'}
Private mcContFullRC As Currency             '{indemnity.ReplacementCost}where {left(indemnity.class,2) ='02'/{rtfarmerindem.cReplaceme} where {left(rtfarmerindem.sClass,2) ='02'}
'Line 04 Recov Depreciation
Private mcDwellRecDep As Currency            '{indemnity.RecoverableDepreciation} where {left(indemnity.class,2) ='01'}/{rtfarmerindem.cRecoverab}where {left(rtfarmerindem.sClass,2) ='01'}
Private mcContRecDep As Currency             '{indemnity.RecoverableDepreciation} where {left(indemnity.class,2) ='02'}/{rtfarmerindem.cRecoverab}where {left(rtfarmerindem.sClass,2) ='02'}
'Line 05 Non Recov Depreciation
Private mcDwellNonRecDep As Currency         '{indemnity.NonRecoverableDepreciation} where {left(indemnity.class,2) ='01'}/{rtfarmerindem.cNonRecove}where {left(rtfarmerindem.sClass,2) ='01'}
Private mcContNonRecDep As Currency          '{indemnity.NonRecoverableDepreciation} where {left(indemnity.class,2) ='02'}/{rtfarmerindem.cNonRecove}where {left(rtfarmerindem.sClass,2) ='02'}
'Line 06 ACV Actual Cash Value loss is (Line 3 minus Line 4 and Line 5)
Private mcDwellACV As Currency               '{indemnity.ACVClaim} where {left(indemnity.class,2) ='01'}/{rtfarmerindem.cACVClaim}where {left(rtfarmerindem.sClass,2) ='01'}
Private mcContACV As Currency                '{indemnity.ACVClaim} where {left(indemnity.class,2) ='02'}/{rtfarmerindem.cACVClaim}where {left(rtfarmerindem.sClass,2) ='02'}
'XS Limits
Private mcDwellXSLim As Currency             '{indemnity.LessExcessLimit} where {left(indemnity.class,2) ='01'}/{rtfarmerindem.cExcessLim}where {left(rtfarmerindem.sClass,2) ='01'}
Private mcContXSLim As Currency              '{indemnity.LessExcessLimit} where {left(indemnity.class,2) ='02'}/{rtfarmerindem.cExcessLim}where {left(rtfarmerindem.sClass,2) ='02'}
'Acv Less XS Limits (Deductible not applied yet)
Private mcDwellACVLessXSLim As Currency      '{indemnity.ACVLessExcessLimit} where {left(indemnity.class,2) ='01'}/{rtfarmerindem.cACVLessEx}where {left(rtfarmerindem.sClass,2) ='01'}
Private mcContACVLessXSLim As Currency       '{indemnity.ACVLessExcessLimit} where {left(indemnity.class,2) ='02'}/{rtfarmerindem.cACVLessEx}where {left(rtfarmerindem.sClass,2) ='02'}
'Line07 Dwelling Less Prior Payments
'<!--- Leave these 0 value for first version--->
Private mcDwellPrevPaymnet As Currency
Private mcContPrevPayment As Currency
'Line 08 Deductible
'See --->PopulateInitDeductibles
Private mcDeductible As Currency             '{Assignments.Deductible}/{rtfarmer.RT15_c}
Private mcDwellDed As Currency               'See --->PopulateInitDeductibles
Private mcContDed As Currency                'See --->PopulateInitDeductibles
'Line 09 Actual Cash Value Claim is (Line 6 minus Line 7 and Line 8)
Private mcDwellACVC As Currency
Private mcContACVC As Currency
'Line 10 Supplemental Claim, to be filed in accordance with the terms and conditions of the Replacement Cost Coverage
Private mcLineSuppClaim As Currency

Private Type udtRCAFarmers
    f_NC23_sFarmersInsExchange As String
    f_NC15_sFireInsuranceExchange As String
    f_NC22_sTruckInsExchange As String
    f_NC25_sMidCenturyInsCompany As String
    f_Other As String
    f_RCAOther As String
    f_RT10_sInsuredName As String
    f_RT07_sAdjusterName As String
    f_CI11sPolicyNumber As String
    f_RT09_sSALN As String
    f_RT12_dtDateOfLoss As Date
    f_RT11_sLossLocation As String
    f_TypeOfPropertyInvolved As String
    f_Dwell01 As Currency
    f_Dwell02 As Currency
    f_Dwell03 As Currency
    f_Dwell04 As Currency
    f_Dwell05 As Currency
    f_Dwell06 As Currency
    f_Dwell07 As Currency
    f_Dwell08 As Currency
    f_Dwell09 As Currency
    f_Cont01 As Currency
    f_Cont02 As Currency
    f_Cont03 As Currency
    f_Cont04 As Currency
    f_Cont05 As Currency
    f_Cont06 As Currency
    f_Cont07 As Currency
    f_Cont08 As Currency
    f_Cont09 As Currency
    f_Line10WithinDays As String
    f_Line10 As Currency
End Type

Private mCurDB As DAO.Database
Private mConn As ADODB.Connection
Private mRS As ADODB.Recordset
Private RS As DAO.Recordset
Private mARV As V2ARViewer.clsARViewer
Private marRCA As V2ECcarFarmers.arRptFarmersRCA
Private moLRs As V2ECKeyBoard.clsLossReports
Private msWithinDays As String
Private mlPrintCopies As Long
Private msCurIBNumber As String         'Current IBNumber (may have Rebill and or Supplement Suffix in Website Tables) (Used for Assignment and RTFarmers tables)
Private msIBNumber As String            'Original IBNumber, (used for Indemnity tables IBNumber)
Private msAssignmentsID As String   'Use String for SQL Server DB ID
Private msID As String              'Use String for USer DB ID
Private msClientClaimNo As String
Private msDwellClass As String
Private msContClass As String

Public Property Let DwellClass(psName As String)
    msDwellClass = psName
End Property
Public Property Get DwellClass() As String
    DwellClass = msDwellClass
End Property

Public Property Let ContClass(psName As String)
    msContClass = psName
End Property
Public Property Get ContClass() As String
    ContClass = msContClass
End Property

Public Property Let CurIBNumber(psName As String)
    msCurIBNumber = psName
End Property
Public Property Get CurIBNumber() As String
    CurIBNumber = msCurIBNumber
End Property

Public Property Let IBNumber(psName As String)
    msIBNumber = psName
End Property
Public Property Get IBNumber() As String
    IBNumber = msIBNumber
End Property

Public Property Let ClientClaimNo(psName As String)
    msClientClaimNo = psName
End Property
Public Property Get ClientClaimNo() As String
    ClientClaimNo = msClientClaimNo
End Property

Public Property Let AssignmentsID(psAssignmentsID As String)
    'SQL Server DB ID
    msAssignmentsID = psAssignmentsID
End Property
Public Property Get AssignmentsID() As String
 'SQL Server DB ID
    AssignmentsID = msAssignmentsID
End Property

Public Property Let ID(psID As String)
'User DB ID
    msID = psID
End Property
Public Property Get ID() As String
'User DB ID
    ID = msID
End Property

Public Property Let PrintCopies(plCopies As Long)
    mlPrintCopies = plCopies
End Property
Public Property Get PrintCopies() As Long
    PrintCopies = mlPrintCopies
End Property

Public Property Let WithinDays(psWithinDays As String)
    msWithinDays = psWithinDays
End Property
Public Property Get WithinDays() As String
    WithinDays = msWithinDays
End Property

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

Public Property Get ClassName() As String
    ClassName = App.EXEName & ".clsRCA"
End Property

Public Function PrintRCA(Optional psClaimNo As String, _
                                Optional psClientClaimNo As String, _
                                Optional pbPreview As Boolean, _
                                Optional pbShowPrintButton As Boolean = True, _
                                Optional psXportPath As String, _
                                Optional pXportType As ExportType) As Boolean
    On Error GoTo EH
    Dim udtRCA As udtRCAFarmers
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    Set marRCA = New V2ECcarFarmers.arRptFarmersRCA
    
    'See if we are not using using Adjuster DB need
    'to open connection to WEB DB
    If mCurDB Is Nothing Then
        OpenConnection
    End If
    
    'Check the ClaimNO (Ibnumber) and Client Claim no Params
    'If they are set then overdie the memeber variables
    If psClaimNo <> vbNullString Then
        msCurIBNumber = psClaimNo 'Set the CUrrent IBNumber
    End If
    If psClientClaimNo <> vbNullString Then
        msClientClaimNo = psClientClaimNo
    End If
    
    udtRCA = PopulateUDTRCAFarmers
    
    With udtRCA
        marRCA.SetProperty "f_NC23_sFarmersInsExchange", .f_NC23_sFarmersInsExchange, vbString
        marRCA.SetProperty "f_NC15_sFireInsuranceExchange", .f_NC15_sFireInsuranceExchange, vbString
        marRCA.SetProperty "f_NC22_sTruckInsExchange", .f_NC22_sTruckInsExchange, vbString
        marRCA.SetProperty "f_NC25_sMidCenturyInsCompany", .f_NC25_sMidCenturyInsCompany, vbString
        marRCA.SetProperty "f_Other", .f_Other, vbString
        marRCA.SetProperty "f_RCAOther", .f_RCAOther, vbString
        marRCA.SetProperty "f_RT10_sInsuredName", .f_RT10_sInsuredName, vbString
        marRCA.SetProperty "f_RT07_sAdjusterName", .f_RT07_sAdjusterName, vbString
        marRCA.SetProperty "f_CI11sPolicyNumber", .f_CI11sPolicyNumber, vbString
        marRCA.SetProperty "f_RT09_sSALN", .f_RT09_sSALN, vbString
        marRCA.SetProperty "f_RT12_dtDateOfLoss", .f_RT12_dtDateOfLoss, vbString
        marRCA.SetProperty "f_RT11_sLossLocation", .f_RT11_sLossLocation, vbString
        marRCA.SetProperty "f_TypeOfPropertyInvolved", .f_TypeOfPropertyInvolved, vbString
        marRCA.SetProperty "f_Dwell01", .f_Dwell01, vbCurrency
        marRCA.SetProperty "f_Cont01", .f_Cont01, vbCurrency
        marRCA.SetProperty "f_Dwell02", .f_Dwell02, vbCurrency
        marRCA.SetProperty "f_Cont02", .f_Cont02, vbCurrency
        marRCA.SetProperty "f_Dwell03", .f_Dwell03, vbCurrency
        marRCA.SetProperty "f_Cont03", .f_Cont03, vbCurrency
        marRCA.SetProperty "f_Dwell04", .f_Dwell04, vbCurrency
        marRCA.SetProperty "f_Cont04", .f_Cont04, vbCurrency
        marRCA.SetProperty "f_Dwell05", .f_Dwell05, vbCurrency
        marRCA.SetProperty "f_Cont05", .f_Cont05, vbCurrency
        marRCA.SetProperty "f_Dwell06", .f_Dwell06, vbCurrency
        marRCA.SetProperty "f_Cont06", .f_Cont06, vbCurrency
        marRCA.SetProperty "f_Dwell07", .f_Dwell07, vbCurrency
        marRCA.SetProperty "f_Cont07", .f_Cont07, vbCurrency
        marRCA.SetProperty "f_Dwell08", .f_Dwell08, vbCurrency
        marRCA.SetProperty "f_Cont08", .f_Cont08, vbCurrency
        marRCA.SetProperty "f_Dwell09", .f_Dwell09, vbCurrency
        marRCA.SetProperty "f_Cont09", .f_Cont09, vbCurrency
        marRCA.SetProperty "f_Line10WithinDays", .f_Line10WithinDays, vbString
        marRCA.SetProperty "f_Line10", .f_Line10, vbCurrency
    End With
    
    With marRCA
        If Not mCurDB Is Nothing Then
            '190  Printing on HP1218  5.15.2002
            'Need to use WIN default printer setting if it is selected under the menu
            On Error Resume Next
            .Printer.DeviceName = GetSetting("Eberls", "PRINTER", "PRINTER_NAME", vbNullString)
            If Err.Number > 0 Then
                Err.Clear
            End If
            .Printer.Port = GetSetting("Eberls", "PRINTER", "PRINTER_PORT", vbNullString)
            If Err.Number > 0 Then
                Err.Clear
            End If
        End If
        On Error GoTo EH
        .Printer.Orientation = ddOPortrait
        If mlPrintCopies > 0 Then
            .Printer.Copies = mlPrintCopies
        End If
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
                marRCA.Run True
                .objARvReport = marRCA
                .sRptTitle = "RCA"
                .HidePrintButton = Not pbShowPrintButton
                .ShowReport vbModeless
            End With
        Else
            marRCA.Run True
        End If
    
    
        'Need to wait for the report to becom active
        Do
            If marRCA.ActiveFlag Then
                If Not pbPreview Then
                    marRCA.PrintReport False 'Don't show print setup dialog
                End If
                Exit Do
            End If
            DoEvents
            Sleep 500
        Loop
        PrintRCA = True
    Else   'Web DB
        Set moLRs = New V2ECKeyBoard.clsLossReports
        If moLRs.ExportFile(marRCA, psXportPath, pXportType) Then
            PrintRCA = True
        End If
        moLRs.CLEANUP
        Set moLRs = Nothing
    End If
    
    'Clean up
    Unload marRCA
    Set marRCA = Nothing
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Screen.MousePointer = vbDefault
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Public Function PrintRCA"
End Function

Private Function PopulateUDTRCAFarmers() As udtRCAFarmers
    On Error GoTo EH
    Dim sSQL As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    'Adjuster Access DB
    If Not mCurDB Is Nothing Then
        sSQL = "z_spsGetNCCReport " & AssignmentsID & " " '<!---@AssignmentsID     int--->
        Set RS = mCurDB.OpenRecordset(sSQL)

        With PopulateUDTRCAFarmers
            If Not RS.EOF Then
                RS.MoveFirst
                .f_NC15_sFireInsuranceExchange = "X"
                    .f_NC22_sTruckInsExchange = "X"
                    .f_NC23_sFarmersInsExchange = "X"
                    .f_NC25_sMidCenturyInsCompany = "X"
                'See if the Other is populated
                If IIf(IsNull(RS!NCCTexasFarmersInsurance), False, RS!NCCTexasFarmersInsurance) Then
                    .f_RCAOther = "Texas Farmers Insurance Company"
                ElseIf IIf(IsNull(RS!NCCFarmersInsOfOregon), False, RS!NCCFarmersInsOfOregon) Then
                    .f_RCAOther = "Farmers Insurance Co. of Oregon"
                ElseIf IIf(IsNull(RS!NCCFarmersOfColumbus), False, RS!NCCFarmersOfColumbus) Then
                    .f_RCAOther = "Farmers Insurance CO. of Columbus, Inc."
                ElseIf IIf(IsNull(RS!NCCIllinoisFarmers), False, RS!NCCIllinoisFarmers) Then
                    .f_RCAOther = "Illinois Farmers Insurance Company"
                ElseIf IIf(IsNull(RS!NCCMidCenturyInsOfTexas), False, RS!NCCMidCenturyInsOfTexas) Then
                    .f_RCAOther = "Mid-Century Insurance Co. of Texas"
                ElseIf IIf(IsNull(RS!NCCFarmersInsExchangeLACA), False, RS!NCCFarmersInsExchangeLACA) Then
                    .f_RCAOther = "Farmers Insurance Exchange, LA CA"
                ElseIf IIf(IsNull(RS!NCCFarmersInsOfWashington), False, RS!NCCFarmersInsOfWashington) Then
                    .f_RCAOther = "Farmers Insurance of Washington"
                ElseIf IIf(IsNull(RS!NCCFarmersInsCompany), False, RS!NCCFarmersInsCompany) Then
                    .f_RCAOther = "Farmers Insurance Company, Inc."
                ElseIf IIf(IsNull(RS!NCCFarmersInsCoOfAZ), False, RS!NCCFarmersInsCoOfAZ) Then
                    .f_RCAOther = "Farmers Insurance Co. of Arizona"
                ElseIf IIf(IsNull(RS!NCCFarmersInsOfIdaho), False, RS!NCCFarmersInsOfIdaho) Then
                    .f_RCAOther = "Farmers Insurance Co. of Idaho"
                ElseIf IIf(IsNull(RS!NCCDomesticStockCo), False, RS!NCCDomesticStockCo) Then
                    .f_RCAOther = "Domestic Stock Company"
                ElseIf IIf(IsNull(RS!NCCFarmersTexasCounty), False, RS!NCCFarmersTexasCounty) Then
                    .f_RCAOther = "Farmers Texas County Mutual Insurance Company"
                End If
                If .f_RCAOther <> vbNullString Then
                    .f_Other = "X"
                End If
                'Rest of the Header info
                .f_RT10_sInsuredName = IIf(IsNull(RS!InsuredNames), vbNullString, RS!InsuredNames)
                .f_RT07_sAdjusterName = IIf(IsNull(RS!ADJName), vbNullString, RS!ADJName)
                .f_CI11sPolicyNumber = IIf(IsNull(RS!PolicyNumber), vbNullString, RS!PolicyNumber)
                .f_RT09_sSALN = msClientClaimNo
                .f_RT12_dtDateOfLoss = IIf(IsNull(RS!LossDate), NULL_DATE, IIf(IsDate(RS!LossDate), RS!LossDate, NULL_DATE))
                .f_RT11_sLossLocation = IIf(IsNull(RS!LossLocation), vbNullString, RS!LossLocation)
                .f_TypeOfPropertyInvolved = vbNullString 'Not Populated Yet ?
                
                'Start Line Items
                '1.     Full Amount of Insurance applicable to the
                '       property for which claim is presented was
                mcDwellPolicyLim = IIf(IsNull(RS!BuildingPolicyLimit), 0, RS!BuildingPolicyLimit)
                mcContPolicyLim = IIf(IsNull(RS!PersonalPropertyLimit), 0, RS!PersonalPropertyLimit)
                
                
                '2.     Full Replacement Cost of the said property
                '       at the time of the loss was
                mcDwellRCSaid = IIf(IsNull(RS!RCSaidBuilding), 0, RS!RCSaidBuilding)
                mcContRCSaid = IIf(IsNull(RS!RCSaidPersonal), 0, RS!RCSaidPersonal)
                '
                'Populate the initial Deductible
                mcDeductible = IIf(IsNull(RS!Deductible), 0, RS!Deductible)
            End If
        End With
        RS.Close
        Set RS = Nothing
        
    Else ' Web DB
        With PopulateUDTRCAFarmers
            sSQL = "SELECT RTFARMERNCC.NC15_SFI As NCCFireInsuranceExchange, "
            sSQL = sSQL & "RTFARMERNCC.NC22_STR As NCCTruckInsExchange, "
            sSQL = sSQL & "RTFARMERNCC.NC23_SFA As NCCFarmersInsExchange, "
            sSQL = sSQL & "RTFARMERNCC.NC25_SMI As NCCMidCenturyInsCompany, "
            sSQL = sSQL & "RTFARMERNCC.NC16_STE As NCCTexasFarmersInsurance, "
            sSQL = sSQL & "RTFARMERNCC.NC17_SFA As NCCFarmersInsOfOregon, "
            sSQL = sSQL & "RTFARMERNCC.NC18_SFA As NCCFarmersOfColumbus, "
            sSQL = sSQL & "RTFARMERNCC.NC19_SIL As NCCIllinoisFarmers, "
            sSQL = sSQL & "RTFARMERNCC.NC20_SMI As NCCMidCenturyInsOfTexas, "
            sSQL = sSQL & "RTFARMERNCC.NC21_SFA As NCCFarmersInsExchangeLACA, "
            sSQL = sSQL & "RTFARMERNCC.NC24_SFA As NCCFarmersInsOfWashington, "
            sSQL = sSQL & "RTFARMERNCC.NC26_SFA As NCCFarmersInsCompany, "
            sSQL = sSQL & "RTFARMERNCC.NC27_SFA As NCCFarmersInsCoOfAZ, "
            sSQL = sSQL & "RTFARMERNCC.NC28_SFA As NCCFarmersInsOfIdaho, "
            sSQL = sSQL & "RTFARMERNCC.NC29_SDO As NCCDomesticStockCo, "
            sSQL = sSQL & "RTFARMERNCC.NC30_SFA As NCCFarmersTexasCounty "
            sSQL = sSQL & "FROM RTFARMERNCC "
            sSQL = sSQL & "WHERE RTFARMERNCC.SIBNUMBER = '" & goUtil.utCleanSQLString(msCurIBNumber) & "' "
            sSQL = sSQL & "AND RTFARMERNCC.SCLIENTCLA = '" & goUtil.utCleanSQLString(msClientClaimNo) & "' "
        
            Set mRS = New ADODB.Recordset
            mRS.CursorLocation = adUseClient
            mRS.Open sSQL, mConn, adOpenStatic, adLockReadOnly
            Set mRS.ActiveConnection = Nothing
            
            If Not mRS.EOF Then
                mRS.MoveFirst
                If IIf(IsNull(mRS!NCCFireInsuranceExchange), False, IIf(RTrim(mRS!NCCFireInsuranceExchange) = "Y", True, False)) Then
                    .f_NC15_sFireInsuranceExchange = "X"
                End If
                If IIf(IsNull(mRS!NCCTruckInsExchange), False, IIf(RTrim(mRS!NCCTruckInsExchange) = "Y", True, False)) Then
                    .f_NC22_sTruckInsExchange = "X"
                End If
                If IIf(IsNull(mRS!NCCFarmersInsExchange), False, IIf(RTrim(mRS!NCCFarmersInsExchange) = "Y", True, False)) Then
                    .f_NC23_sFarmersInsExchange = "X"
                End If
                If IIf(IsNull(mRS!NCCMidCenturyInsCompany), False, IIf(RTrim(mRS!NCCMidCenturyInsCompany) = "Y", True, False)) Then
                    .f_NC25_sMidCenturyInsCompany = "X"
                End If
                'See if the Other is populated
                If IIf(IsNull(mRS!NCCTexasFarmersInsurance), False, IIf(RTrim(mRS!NCCTexasFarmersInsurance) = "Y", True, False)) Then
                    .f_RCAOther = "Texas Farmers Insurance Company"
                ElseIf IIf(IsNull(mRS!NCCFarmersInsOfOregon), False, IIf(RTrim(mRS!NCCFarmersInsOfOregon) = "Y", True, False)) Then
                    .f_RCAOther = "Farmers Insurance Co. of Oregon"
                ElseIf IIf(IsNull(mRS!NCCFarmersOfColumbus), False, IIf(RTrim(mRS!NCCFarmersOfColumbus) = "Y", True, False)) Then
                    .f_RCAOther = "Farmers Insurance CO. of Columbus, Inc."
                ElseIf IIf(IsNull(mRS!NCCIllinoisFarmers), False, IIf(RTrim(mRS!NCCIllinoisFarmers) = "Y", True, False)) Then
                    .f_RCAOther = "Illinois Farmers Insurance Company"
                ElseIf IIf(IsNull(mRS!NCCMidCenturyInsOfTexas), False, IIf(RTrim(mRS!NCCMidCenturyInsOfTexas) = "Y", True, False)) Then
                    .f_RCAOther = "Mid-Century Insurance Co. of Texas"
                ElseIf IIf(IsNull(mRS!NCCFarmersInsExchangeLACA), False, IIf(RTrim(mRS!NCCFarmersInsExchangeLACA) = "Y", True, False)) Then
                    .f_RCAOther = "Farmers Insurance Exchange, LA CA"
                ElseIf IIf(IsNull(mRS!NCCFarmersInsOfWashington), False, IIf(RTrim(mRS!NCCFarmersInsOfWashington) = "Y", True, False)) Then
                    .f_RCAOther = "Farmers Insurance of Washington"
                ElseIf IIf(IsNull(mRS!NCCFarmersInsCompany), False, IIf(RTrim(mRS!NCCFarmersInsCompany) = "Y", True, False)) Then
                    .f_RCAOther = "Farmers Insurance Company, Inc."
                ElseIf IIf(IsNull(mRS!NCCFarmersInsCoOfAZ), False, IIf(RTrim(mRS!NCCFarmersInsCoOfAZ) = "Y", True, False)) Then
                    .f_RCAOther = "Farmers Insurance Co. of Arizona"
                ElseIf IIf(IsNull(mRS!NCCFarmersInsOfIdaho), False, IIf(RTrim(mRS!NCCFarmersInsOfIdaho) = "Y", True, False)) Then
                    .f_RCAOther = "Farmers Insurance Co. of Idaho"
                ElseIf IIf(IsNull(mRS!NCCDomesticStockCo), False, IIf(RTrim(mRS!NCCDomesticStockCo) = "Y", True, False)) Then
                    .f_RCAOther = "Domestic Stock Company"
                ElseIf IIf(IsNull(mRS!NCCFarmersTexasCounty), False, IIf(RTrim(mRS!NCCFarmersTexasCounty) = "Y", True, False)) Then
                    .f_RCAOther = "Farmers Texas County Mutual Insurance Company"
                End If
                If .f_RCAOther <> vbNullString Then
                    .f_Other = "X"
                End If
            End If
            mRS.Close
            sSQL = "SELECT RTFARMER.RT50_S As InsuredNames, "
            sSQL = sSQL & "RTFARMER.RT07_S As ADJName, "
            sSQL = sSQL & "RTFARMER.CI11_SPO As PolicyNumber, "
            sSQL = sSQL & "RTFARMER.RT12_DT AS LossDate, "
            sSQL = sSQL & "RTFARMER.RT11_S As LossLocation, "
            sSQL = sSQL & "RTFARMER.CI14_CBU As BuildingPolicyLimit, "
            sSQL = sSQL & "RTFARMER.CI15_CCO As PersonalPropertyLimit, "
            sSQL = sSQL & "RTFARMER.CI17_CBU As RCSaidBuilding, "
            sSQL = sSQL & "RTFARMER.CI18_CCO As RCSaidPersonal, "
            sSQL = sSQL & "RTFARMER.RT15_C As Deductible "
            sSQL = sSQL & "FROM RTFARMER "
            sSQL = sSQL & "WHERE RTFARMER.SIBNUMBER = '" & goUtil.utCleanSQLString(msCurIBNumber) & "' "
            sSQL = sSQL & "AND RTFARMER.SCLIENTCLA = '" & goUtil.utCleanSQLString(msClientClaimNo) & "' "
            
            Set mRS = New ADODB.Recordset
            mRS.CursorLocation = adUseClient
            mRS.Open sSQL, mConn, adOpenStatic, adLockReadOnly
            Set mRS.ActiveConnection = Nothing
            
            If Not mRS.EOF Then
                'Rest of the Header info
                .f_RT10_sInsuredName = IIf(IsNull(mRS!InsuredNames), vbNullString, mRS!InsuredNames)
                .f_RT07_sAdjusterName = IIf(IsNull(mRS!ADJName), vbNullString, mRS!ADJName)
                .f_CI11sPolicyNumber = IIf(IsNull(mRS!PolicyNumber), vbNullString, mRS!PolicyNumber)
                .f_RT09_sSALN = msClientClaimNo
                .f_RT12_dtDateOfLoss = IIf(IsNull(mRS!LossDate), NULL_DATE, IIf(IsDate(mRS!LossDate), mRS!LossDate, NULL_DATE))
                .f_RT11_sLossLocation = IIf(IsNull(mRS!LossLocation), vbNullString, mRS!LossLocation)
                .f_TypeOfPropertyInvolved = vbNullString 'Not Populated Yet ?
                
                'Start Line Items
                '1.     Full Amount of Insurance applicable to the
                '       property for which claim is presented was
                mcDwellPolicyLim = IIf(IsNull(mRS!BuildingPolicyLimit), 0, mRS!BuildingPolicyLimit)
                mcContPolicyLim = IIf(IsNull(mRS!PersonalPropertyLimit), 0, mRS!PersonalPropertyLimit)
                
                '2.     Full Replacement Cost of the said property
                '       at the time of the loss was
                mcDwellRCSaid = IIf(IsNull(mRS!RCSaidBuilding), 0, mRS!RCSaidBuilding)
                mcContRCSaid = IIf(IsNull(mRS!RCSaidPersonal), 0, mRS!RCSaidPersonal)
                '
                'Populate the initial Deductible
                mcDeductible = IIf(IsNull(mRS!Deductible), 0, mRS!Deductible)
            End If
            mRS.Close
        End With
        Set mRS = Nothing
    End If
    
    With PopulateUDTRCAFarmers
        'Line Item 01
        .f_Dwell01 = mcDwellPolicyLim
        .f_Cont01 = mcContPolicyLim
        
        'Line Item 02
        .f_Dwell02 = mcDwellRCSaid
        .f_Cont02 = mcContRCSaid
        
        'Line Items 3 --> 6 get from PolulateIndemValues
        PopulateIndemValues
        .f_Dwell03 = mcDwellFullRC
        .f_Cont03 = mcContFullRC
        .f_Dwell04 = mcDwellRecDep
        .f_Cont04 = mcContRecDep
        .f_Dwell05 = mcDwellNonRecDep
        .f_Cont05 = mcContNonRecDep
        .f_Dwell06 = mcDwellACV
        .f_Cont06 = mcContACV
        
        '7.     Less Prior Payments
        'These are not poulated in 1st version
        mcDwellPrevPaymnet = 0
        mcContPrevPayment = 0
        .f_Dwell07 = mcDwellPrevPaymnet
        .f_Cont07 = mcContPrevPayment
        
        '8.     Less Deductible
        'First get the initial deductible values. Need these first to figure all deductibles.
        PopulateInitDeductibles
        .f_Dwell08 = GetDwellDed
        .f_Cont08 = GetContDed
        
        '9.     Actual Cash Value Claim is (Line 6 minus Line 7 and Line 8)
        .f_Dwell09 = GetDwellACV
        .f_Cont09 = GetContACV
        
        '10.        Supplemental Claim, to be filed in accordance with the
        '           terms and conditions of the Replacement Cost Coverage
        '           within ____ days of date of loss as shown above, will not
        '           exceed .... ______
        '           (This figure will be that portion of the amounts shown on
        '           Line 4 which is recoverable)
        .f_Line10WithinDays = msWithinDays
        .f_Line10 = GetSuppClaim
    End With
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Set RS = Nothing
    Set mRS = Nothing
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Private Function PopulateUDTRCAFarmers"
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

Public Function CLEANUP() As Boolean
    On Error Resume Next
    If Not RS Is Nothing Then
        RS.Close
        Set RS = Nothing
    End If
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

Private Sub PopulateIndemValues()
    '    'Line 03 The Full Cost of Repair or Replacement is
'    mcDwellFullRC             '{indemnity.ReplacementCost} where {left(indemnity.class,2) ='01'}/{rtfarmerindem.cReplaceme} where {left(rtfarmerindem.sClass,2) ='01'}
'    mcContFullRC              '{indemnity.ReplacementCost}where {left(indemnity.class,2) ='02'/{rtfarmerindem.cReplaceme} where {left(rtfarmerindem.sClass,2) ='02'}
'    'Line 04 Recov Depreciation
'    mcDwellRecDep            '{indemnity.RecoverableDepreciation} where {left(indemnity.class,2) ='01'}/{rtfarmerindem.cRecoverab}where {left(rtfarmerindem.sClass,2) ='01'}
'    mcContRecDep             '{indemnity.RecoverableDepreciation} where {left(indemnity.class,2) ='02'}/{rtfarmerindem.cRecoverab}where {left(rtfarmerindem.sClass,2) ='02'}
'    'Line 05 Non Recov Depreciation
'    mcDwellNonRecDep         '{indemnity.NonRecoverableDepreciation} where {left(indemnity.class,2) ='01'}/{rtfarmerindem.cNonRecove}where {left(rtfarmerindem.sClass,2) ='01'}
'    mcContNonRecDep          '{indemnity.NonRecoverableDepreciation} where {left(indemnity.class,2) ='02'}/{rtfarmerindem.cNonRecove}where {left(rtfarmerindem.sClass,2) ='02'}
'    'Line 06 ACV Actual Cash Value loss is (Line 3 minus Line 4 and Line 5)
'    mcDwellACV              '{indemnity.ACVClaim} where {left(indemnity.class,2) ='01'}/{rtfarmerindem.cACVClaim}where {left(rtfarmerindem.sClass,2) ='01'}
'    mcContACV               '{indemnity.ACVClaim} where {left(indemnity.class,2) ='02'}/{rtfarmerindem.cACVClaim}where {left(rtfarmerindem.sClass,2) ='02'}
'    'XS Limits
'    mcDwellXSLim              '{indemnity.LessExcessLimit} where {left(indemnity.class,2) ='01'}/{rtfarmerindem.cExcessLim}where {left(rtfarmerindem.sClass,2) ='01'}
'    mcContXSLim              '{indemnity.LessExcessLimit} where {left(indemnity.class,2) ='02'}/{rtfarmerindem.cExcessLim}where {left(rtfarmerindem.sClass,2) ='02'}
'    'Acv Less XS Limits (Deductible not applied yet)
'    mcDwellACVLessXSLim      '{indemnity.ACVLessExcessLimit} where {left(indemnity.class,2) ='01'}/{rtfarmerindem.cACVLessEx}where {left(rtfarmerindem.sClass,2) ='01'}
'    mcContACVLessXSLim        '{indemnity.ACVLessExcessLimit} where {left(indemnity.class,2) ='02'}/{rtfarmerindem.cACVLessEx}where {left(rtfarmerindem.sClass,2) ='02'}
    On Error GoTo EH
    Dim lClassCount As Long
    Dim sClass As String
    Dim RS As DAO.Recordset
    Dim sSQL As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    'Make sure these are clear
    mcDwellFullRC = 0
    mcDwellRecDep = 0
    mcDwellNonRecDep = 0
    mcDwellACV = 0
    mcDwellXSLim = 0
    mcDwellACVLessXSLim = 0
    mcContFullRC = 0
    mcContRecDep = 0
    mcContNonRecDep = 0
    mcContACV = 0
    mcContXSLim = 0
    mcContACVLessXSLim = 0
    
    'Adjuster Access DB
    If Not mCurDB Is Nothing Then
        For lClassCount = 1 To 2
            sClass = Switch(lClassCount = 1, msDwellClass, _
                            lClassCount = 2, msContClass, _
                            True, vbNullString)
            sSQL = "SELECT A.ReplacementCost, "
            sSQL = sSQL & "A.RecoverableDepreciation, "
            sSQL = sSQL & "A.NonRecoverableDepreciation, "
            sSQL = sSQL & "A.ACVClaim, "
            sSQL = sSQL & "A.LessExcessLimit, "
            sSQL = sSQL & "A.ACVLessExcessLimit "
            sSQL = sSQL & "FROM Indemnity As A "
            sSQL = sSQL & "WHERE A.ClaimNo = '" & goUtil.utCleanSQLString(msCurIBNumber) & "' " 'Use Current IB Number
            sSQL = sSQL & "AND Left(A.Class,2) IN(" & sClass & ") "
            Set RS = mCurDB.OpenRecordset(sSQL)
            If Not RS.EOF Then
                RS.MoveFirst
                Do Until RS.EOF
                    If lClassCount = 1 Then
                        mcDwellFullRC = mcDwellFullRC + IIf(IsNull(RS!ReplacementCost), 0, RS!ReplacementCost)
                        mcDwellRecDep = mcDwellRecDep + IIf(IsNull(RS!RecoverableDepreciation), 0, RS!RecoverableDepreciation)
                        mcDwellNonRecDep = mcDwellNonRecDep + IIf(IsNull(RS!NonRecoverableDepreciation), 0, RS!NonRecoverableDepreciation)
                        mcDwellACV = mcDwellACV + IIf(IsNull(RS!ACVClaim), 0, RS!ACVClaim)
                        mcDwellXSLim = mcDwellXSLim + IIf(IsNull(RS!LessExcessLimit), 0, RS!LessExcessLimit)
                        mcDwellACVLessXSLim = mcDwellACVLessXSLim + IIf(IsNull(RS!ACVLessExcessLimit), 0, RS!ACVLessExcessLimit)
                    ElseIf lClassCount = 2 Then
                        mcContFullRC = mcContFullRC + IIf(IsNull(RS!ReplacementCost), 0, RS!ReplacementCost)
                        mcContRecDep = mcContRecDep + IIf(IsNull(RS!RecoverableDepreciation), 0, RS!RecoverableDepreciation)
                        mcContNonRecDep = mcContNonRecDep + IIf(IsNull(RS!NonRecoverableDepreciation), 0, RS!NonRecoverableDepreciation)
                        mcContACV = mcContACV + IIf(IsNull(RS!ACVClaim), 0, RS!ACVClaim)
                        mcContXSLim = mcContXSLim + IIf(IsNull(RS!LessExcessLimit), 0, RS!LessExcessLimit)
                        mcContACVLessXSLim = mcContACVLessXSLim + IIf(IsNull(RS!ACVLessExcessLimit), 0, RS!ACVLessExcessLimit)
                    End If
                    RS.MoveNext
                Loop
            End If
            RS.Close
        Next
        Set RS = Nothing
    Else ' Web DB
        For lClassCount = 1 To 2
            sClass = Switch(lClassCount = 1, msDwellClass, _
                            lClassCount = 2, msContClass, _
                            True, vbNullString)
            sSQL = "SELECT RTFARMERINDEM.CREPLACEME As ReplacementCost, "
            sSQL = sSQL & "RTFARMERINDEM.CRECOVERAB As RecoverableDepreciation, "
            sSQL = sSQL & "RTFARMERINDEM.CNONRECOVE As NonRecoverableDepreciation, "
            sSQL = sSQL & "RTFARMERINDEM.CACVCLAIM As ACVClaim, "
            sSQL = sSQL & "RTFARMERINDEM.CEXCESSLIM As LessExcessLimit, "
            sSQL = sSQL & "RTFARMERINDEM.CACVLESSEX As ACVLessExcessLimit "
            sSQL = sSQL & "FROM RTFARMERINDEM "
            sSQL = sSQL & "WHERE RTFARMERINDEM.SIBNUMBER = '" & goUtil.utCleanSQLString(msIBNumber) & "' "      'Use Original IBNumber
            sSQL = sSQL & "AND RTFARMERINDEM.SCLIENTCLA  = '" & goUtil.utCleanSQLString(msClientClaimNo) & "' "
            sSQL = sSQL & "AND LEFT(RTFARMERINDEM.SCLASS,2) IN(" & sClass & ") "
            
            Set mRS = New ADODB.Recordset
            mRS.CursorLocation = adUseClient
            mRS.Open sSQL, mConn, adOpenStatic, adLockReadOnly
            Set mRS.ActiveConnection = Nothing
            
            If Not mRS.EOF Then
                mRS.MoveFirst
                Do Until mRS.EOF
                    If lClassCount = 1 Then
                        mcDwellFullRC = mcDwellFullRC + IIf(IsNull(mRS!ReplacementCost), 0, mRS!ReplacementCost)
                        mcDwellRecDep = mcDwellRecDep + IIf(IsNull(mRS!RecoverableDepreciation), 0, mRS!RecoverableDepreciation)
                        mcDwellNonRecDep = mcDwellNonRecDep + IIf(IsNull(mRS!NonRecoverableDepreciation), 0, mRS!NonRecoverableDepreciation)
                        mcDwellACV = mcDwellACV + IIf(IsNull(mRS!ACVClaim), 0, mRS!ACVClaim)
                        mcDwellXSLim = mcDwellXSLim + IIf(IsNull(mRS!LessExcessLimit), 0, mRS!LessExcessLimit)
                        mcDwellACVLessXSLim = mcDwellACVLessXSLim + IIf(IsNull(mRS!ACVLessExcessLimit), 0, mRS!ACVLessExcessLimit)
                    ElseIf lClassCount = 2 Then
                        mcContFullRC = mcContFullRC + IIf(IsNull(mRS!ReplacementCost), 0, mRS!ReplacementCost)
                        mcContRecDep = mcContRecDep + IIf(IsNull(mRS!RecoverableDepreciation), 0, mRS!RecoverableDepreciation)
                        mcContNonRecDep = mcContNonRecDep + IIf(IsNull(mRS!NonRecoverableDepreciation), 0, mRS!NonRecoverableDepreciation)
                        mcContACV = mcContACV + IIf(IsNull(mRS!ACVClaim), 0, mRS!ACVClaim)
                        mcContXSLim = mcContXSLim + IIf(IsNull(mRS!LessExcessLimit), 0, mRS!LessExcessLimit)
                        mcContACVLessXSLim = mcContACVLessXSLim + IIf(IsNull(mRS!ACVLessExcessLimit), 0, mRS!ACVLessExcessLimit)
                    End If
                    mRS.MoveNext
                Loop
            End If
            mRS.Close
        Next
        Set RS = Nothing
    End If
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Private Sub PopulateIndemValues"
End Sub

Private Sub PopulateInitDeductibles()
    On Error GoTo EH
    Dim cDed As Currency
    Dim cRemainDed As Currency
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    'Set the deductible to the initial deductible
    cDed = mcDeductible
    
    'Figure the Dwelling Deductible
    If mcDwellACVLessXSLim > cDed Then
        mcDwellDed = cDed - mcDwellXSLim
        If mcDwellDed < 0 Then
            mcDwellDed = 0
        End If
    Else
        cRemainDed = cDed - mcDwellACVLessXSLim
        mcDwellDed = cDed - cRemainDed
    End If
    
    'Set the deductible to the Remaining deductible
    cDed = cRemainDed
    
    'Figure the Contents Deductible
    If mcContACVLessXSLim > cDed Then
        mcContDed = cDed - mcContXSLim
        If mcContDed < 0 Then
            mcContDed = 0
        End If
    Else
        cRemainDed = cDed - mcContACVLessXSLim
        mcContDed = cDed - cRemainDed
    End If
    
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Private Sub PopulateInitDeductibles"
End Sub

Private Function GetDwellDed() As Currency
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    'Line 8

    If mcDwellFullRC = 0 Then
        GetDwellDed = 0
    ElseIf mcContACV = 0 Then
        GetDwellDed = mcDwellDed
    ElseIf GetDwellDed1 + GetContDed1 <= mcDeductible Then
        GetDwellDed = mcDwellDed
    End If
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Private Function GetDwellDed"
End Function

Private Function GetContDed() As Currency
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    'Line 8 Contents

    If mcContFullRC = 0 Then
        GetContDed = 0
    ElseIf GetDwellDed1 = mcDeductible Then
        GetContDed = 0
    ElseIf GetDwellDed1 = 0 Then
        GetContDed = mcContDed
    ElseIf GetDwellDed1 + GetContDed1 < mcDeductible Then
        GetContDed = mcDeductible - GetDwellDed1
    ElseIf GetDwellDed1 < mcDeductible Then
        GetContDed = GetContDed1
    End If
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Private Function GetContDed"
End Function

Private Function GetDwellDed1() As Currency
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    If mcDwellACV > mcDeductible Then
        GetDwellDed1 = mcDeductible
    Else
        GetDwellDed1 = mcDwellACV
    End If
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Private Function GetDwellDed1"
End Function

Private Function GetContDed1() As Currency
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    If mcContACV < mcDeductible - GetDwellDed1 Then
        GetContDed1 = GetACVClaimCont
    ElseIf GetDwellDed1 < mcDeductible Then
        GetContDed1 = mcDeductible - GetDwellDed1
    ElseIf GetDwellDed1 = mcDeductible Then
        GetContDed1 = 0
    ElseIf mcContACV < mcDeductible - GetDwellDed1 Then
        GetContDed1 = mcContACV
    End If
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Private Function GetContDed1"
End Function

Private Function GetACVClaimCont() As Currency
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    GetACVClaimCont = mcContACV - mcContDed
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Private Function GetACVClaimCont"
End Function

Private Function GetDwellACV() As Currency
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    'Line 9 Dwelling
    GetDwellACV = mcDwellACV - mcDwellDed
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Private Function GetDwellACV"
End Function

Private Function GetContACV() As Currency
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    'Line 9 Contents
    GetContACV = mcContACV - mcContDed
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Private Function GetContACV"
End Function

Private Function GetSuppClaim() As Currency
    On Error GoTo EH
    'Line 10
    Dim cTotRecDep As Currency
    Dim cUnusedDed As Currency
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    cTotRecDep = mcDwellRecDep + mcContRecDep
    cUnusedDed = mcDeductible - mcDwellDed - mcContDed
    
    GetSuppClaim = cTotRecDep - cUnusedDed
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Private Function GetSuppClaim"
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

