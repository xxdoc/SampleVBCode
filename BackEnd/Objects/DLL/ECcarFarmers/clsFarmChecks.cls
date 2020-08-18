VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFarmChecks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'BGS 9.26.2001 UDT for Farmers Indemnity form
Public Type FarmIndemnity
    a01ClaimNo As String  '= Trim(Text1(9))
    a02IndemType As String '= Trim(gCheckNUM(x))
    a03Class As String '= Trim(gClass(x))
    a04CauseOfLoss As String '= Trim(gSection(x))
    a05Description As String ' = Trim(gDescription(x))
    a06ReplacementCost As Currency '= Trim(gReplCost(x))
    a07RecoverableDepreciation As Currency ' = Trim(gRecDepr(x))
    a08NonRecoverableDepreciation As Currency '= Trim(gNonRecDepr(x))
    a09ACVClaim As Currency '= Trim(gACVClaim(x))
    a10LessExcessLimit As Currency  '= Trim(gExcessLimit(x))
    a10aExcessLimitDesc As String
End Type

Public Enum FChkReq
    a01ForProperty = 1
    a02ForAuto
    a03FinalPayment
    a04PartialPayment
    a05ForAddSupplement
    a06PaymentAssociated
    a07FieldHandled
    a08TotalLoss
    a09CashInLieu
    a10OwnerRetainsSalvage
    a11Subrogation
    a12Salvage
    a13IncludeMortgageeOnDraft
End Enum
'BGS 9.26.2001 UDT for Farmers Check Requests form
Public Type FarmCheckReq
    a01ForProperty As Boolean '= TrueFalse
    a02ForAuto As Boolean '= TrueFalse
    a03FinalPayment As Boolean '= TrueFalse
    a04PartialPayment As Boolean '= TrueFalse
    a05ForAdditionalSupplement As Boolean '= TrueFalse
    a06PaymentAssociated As Boolean '= TrueFalse
    a07FieldHandled As Boolean '= TrueFalse
    a08TotalLoss As Boolean
    a09CashInLieu As Boolean '= TrueFalse
    a10OwnerRetainedSalvage As Boolean ' = TrueFalse
    a11Subrogation As Boolean '= TrueFalse
    a12Salvage As Boolean '= TrueFalse
    a13IncludeMortgageeOnDraft As Boolean
    
    b01Class As String 'Same as Type of Loss "Code for Description"
    b02TypeOfLoss As String 'Same as Class "Description of Class"
    b03CauseOfLoss As String
    b04TexasSubCode As String
    b05TexasSuffix As String
    b06TexasRoofDepreciation As Currency
    b07AmountOfCheck As Currency
    b08Instructions As String
    z01ID As Long
    z02ClaimNo As String
    z03CheckNUM As String
    z04Rebilled As Boolean
    z05Supplement As Integer ' Integer Number
End Type

'BGS DB Values
Private mba01ForProperty As Boolean '= TrueFalse
Private mba02ForAuto As Boolean '= TrueFalse
Private mba03FinalPayment As Boolean '= TrueFalse
Private mba04PartialPayment As Boolean '= TrueFalse
Private mba05ForAdditionalSupplement As Boolean '= TrueFalse
Private mba06PaymentAssociated As Boolean '= TrueFalse
Private mba07FieldHandled As Boolean '= TrueFalse
Private mba08TotalLoss As Boolean
Private mba09CashInLieu As Boolean '= TrueFalse
Private mba10OwnerRetainedSalvage As Boolean ' = TrueFalse
Private mba11Subrogation As Boolean '= TrueFalse
Private mba12Salvage As Boolean '= TrueFalse
Private mba13IncludeMortgageeOnDraft As Boolean
Private msb01Class As String 'Same as Type of Loss "Code for Description"  'B
Private msb02TypeOfLoss As String 'Same as Class "Description of Class"
Private msb03CauseOfLoss As String
Private msb04TexasSubCode As String
Private msb05TexasSuffix As String
Private mcurb06TexasRoofDepreciation As Currency
Private mcurb07AmountOfCheck As Currency
Private msb08Instructions As String
Private mlz01ID As Long  'Z
Private msz02ClaimNo As String
Private msz03CheckNUM As String
Private mbz04Rebilled As Boolean
Private miz05Supplement As Integer ' Integer Number

'Other values
Private mbPopulate As Boolean 'True when Populating and not loading
Private mbForceMortgageeOnDraft As Boolean
Private mbSaveEnabled As Boolean
Private mfrmSaveForm As Form

Public Property Get SaveForm() As Object
    If mbSaveEnabled Then
        Set SaveForm = mfrmSaveForm
    Else
        Set SaveForm = Nothing
    End If
End Property

Public Property Get SaveEnabled() As Boolean
    SaveEnabled = mbSaveEnabled
End Property

Public Property Get ForceMortgageeOnDraft() As Boolean
    ForceMortgageeOnDraft = mbForceMortgageeOnDraft
End Property

Public Property Let ForceMortgageeOnDraft(psFlag As Boolean)
    mbForceMortgageeOnDraft = psFlag
End Property

Public Property Get Populate() As Boolean
    Populate = mbPopulate
End Property

Public Property Get IsDirty(pfrmEditChecks As Object) As Boolean
    On Error GoTo EH
    IsDirty = False
    With pfrmEditChecks
        If .chkRebill.Value <> IIf(mbz04Rebilled, vbChecked, vbUnchecked) Then
            IsDirty = True
            Exit Property
        End If
        If .ChkChekReq(FChkReq.a01ForProperty).Value <> IIf(mba01ForProperty, vbChecked, vbUnchecked) Then
            IsDirty = True
            Exit Property
        End If
        If .ChkChekReq(FChkReq.a02ForAuto).Value <> IIf(mba02ForAuto, vbChecked, vbUnchecked) Then
            IsDirty = True
            Exit Property
        End If
        If .ChkChekReq(FChkReq.a03FinalPayment).Value <> IIf(mba03FinalPayment, vbChecked, vbUnchecked) Then
            IsDirty = True
            Exit Property
        End If
        If .ChkChekReq(FChkReq.a04PartialPayment).Value <> IIf(mba04PartialPayment, vbChecked, vbUnchecked) Then
            IsDirty = True
            Exit Property
        End If
        If .ChkChekReq(FChkReq.a05ForAddSupplement).Value <> IIf(mba05ForAdditionalSupplement, vbChecked, vbUnchecked) Then
            IsDirty = True
            Exit Property
        End If
        If .ChkChekReq(FChkReq.a06PaymentAssociated).Value <> IIf(mba06PaymentAssociated, vbChecked, vbUnchecked) Then
            IsDirty = True
            Exit Property
        End If
        If .ChkChekReq(FChkReq.a07FieldHandled).Value <> IIf(mba07FieldHandled, vbChecked, vbUnchecked) Then
            IsDirty = True
            Exit Property
        End If
        If .ChkChekReq(FChkReq.a08TotalLoss).Value <> IIf(mba08TotalLoss, vbChecked, vbUnchecked) Then
            IsDirty = True
            Exit Property
        End If
        If .ChkChekReq(FChkReq.a09CashInLieu).Value <> IIf(mba09CashInLieu, vbChecked, vbUnchecked) Then
            IsDirty = True
            Exit Property
        End If
        If .ChkChekReq(FChkReq.a10OwnerRetainsSalvage).Value <> IIf(mba10OwnerRetainedSalvage, vbChecked, vbUnchecked) Then
            IsDirty = True
            Exit Property
        End If
        If .ChkChekReq(FChkReq.a11Subrogation).Value <> IIf(mba11Subrogation, vbChecked, vbUnchecked) Then
            IsDirty = True
            Exit Property
        End If
        If .ChkChekReq(FChkReq.a12Salvage).Value <> IIf(mba12Salvage, vbChecked, vbUnchecked) Then
            IsDirty = True
            Exit Property
        End If
        If .ChkChekReq(FChkReq.a13IncludeMortgageeOnDraft).Value <> IIf(mba13IncludeMortgageeOnDraft, vbChecked, vbUnchecked) Then
            IsDirty = True
            Exit Property
        End If
        If .cmbClassCheck.Text <> msb01Class Then
            IsDirty = True
            Exit Property
        End If
        If .cmbCauseOfLossCheck.Text <> msb03CauseOfLoss Then
            IsDirty = True
            Exit Property
        End If
        If .txtTexSubCovCode.Text <> msb04TexasSubCode Then
            IsDirty = True
            Exit Property
        End If
        If .txtTexSuffix.Text <> msb05TexasSuffix Then
            IsDirty = True
            Exit Property
        End If
        If .txtTexAmountOfRoofDepr.Text <> mcurb06TexasRoofDepreciation Then
            IsDirty = True
            Exit Property
        End If
        If .txtCheckAmount.Text <> mcurb07AmountOfCheck Then
            IsDirty = True
            Exit Property
        End If
        If .txtInstructions.Text <> msb08Instructions Then
            IsDirty = True
            Exit Property
        End If
    End With
    Exit Property
EH:
    ShowError Err, "Public Property Get IsDirty", , "ECWeb.clsFarmChecks.cls"
End Property

Private Sub LoadFarmUDT(pudtFarmCheck As FarmCheckReq)
    On Error GoTo EH
    With pudtFarmCheck
        mba01ForProperty = .a01ForProperty
        mba02ForAuto = .a02ForAuto
        mba03FinalPayment = .a03FinalPayment
        mba04PartialPayment = .a04PartialPayment
        mba05ForAdditionalSupplement = .a05ForAdditionalSupplement
        mba06PaymentAssociated = .a06PaymentAssociated
        mba07FieldHandled = .a07FieldHandled
        mba08TotalLoss = .a08TotalLoss
        mba09CashInLieu = .a09CashInLieu
        mba10OwnerRetainedSalvage = .a10OwnerRetainedSalvage
        mba11Subrogation = .a11Subrogation
        mba12Salvage = .a12Salvage
        mba13IncludeMortgageeOnDraft = .a13IncludeMortgageeOnDraft
        msb01Class = .b01Class
        msb02TypeOfLoss = .b02TypeOfLoss
        msb03CauseOfLoss = .b03CauseOfLoss
        msb04TexasSubCode = .b04TexasSubCode
        msb05TexasSuffix = .b05TexasSuffix
        mcurb06TexasRoofDepreciation = Format(.b06TexasRoofDepreciation, "0.00")
        mcurb07AmountOfCheck = Format(.b07AmountOfCheck, "0.00")
        msb08Instructions = .b08Instructions
        mlz01ID = .z01ID
        msz02ClaimNo = .z02ClaimNo
        msz03CheckNUM = .z03CheckNUM
        mbz04Rebilled = .z04Rebilled
        miz05Supplement = .z05Supplement
    End With
    
    Exit Sub
EH:
    ShowError Err, "Private Sub LoadFarmUDT", , "ECWeb.clsFarmChecks.cls"
End Sub

Public Sub LoadFarmChecksUDT(pfrmEditChecks As Object, psClaimNO As String, pRS As Recordset, pCurDB As Database)
    On Error GoTo EH
    Dim sSQL As String
    Dim udtFCheck As FarmCheckReq
    Dim sCheckNUM As String
    
    If mbPopulate Then
        Exit Sub
    End If
    
    With pfrmEditChecks
        sCheckNUM = .cmbCheckNum.Text
        sSQL = "SELECT * FROM Checks A "
        sSQL = sSQL & "WHERE A.ClaimNo = '" & psClaimNO & "' "
        sSQL = sSQL & "AND A.CheckNUM = '" & sCheckNUM & "' "
        
        Set pRS = pCurDB.OpenRecordset(sSQL)
        
        If Not pRS.EOF Then
            pRS.MoveFirst
            With udtFCheck
                .z01ID = pRS!ID
                .z02ClaimNo = pRS!ClaimNo
                .z03CheckNUM = IIf(IsNull(pRS!CheckNUM), vbNullString, pRS!CheckNUM)
                .z04Rebilled = pRS!Rebilled
                .z05Supplement = pRS!Supplement
                .a01ForProperty = pRS!ForProperty
                .a02ForAuto = pRS!ForAuto
                .a03FinalPayment = pRS!FinalPayment
                .a04PartialPayment = pRS!PartialPayment
                .a05ForAdditionalSupplement = pRS!ForAdditionalSupplement
                .a06PaymentAssociated = pRS!PaymentAssociated
                .a07FieldHandled = pRS!FieldHandled
                .a08TotalLoss = pRS!TotalLoss
                .a09CashInLieu = pRS!CashInLieu
                .a10OwnerRetainedSalvage = pRS!OwnerRetainedSalvage
                .a11Subrogation = pRS!Subrogation
                .a12Salvage = pRS!Salvage
                .a13IncludeMortgageeOnDraft = pRS!IncludeMorgOnDraft
                .b01Class = IIf(IsNull(pRS!Class), vbNullString, pRS!Class)
                .b02TypeOfLoss = IIf(IsNull(pRS!TypeOfLoss), vbNullString, pRS!TypeOfLoss)
                .b03CauseOfLoss = IIf(IsNull(pRS!CauseOfLoss), vbNullString, pRS!CauseOfLoss)
                .b04TexasSubCode = IIf(IsNull(pRS!TexasSubCode), vbNullString, pRS!TexasSubCode)
                .b05TexasSuffix = IIf(IsNull(pRS!TexasSuffix), vbNullString, pRS!TexasSuffix)
                .b06TexasRoofDepreciation = pRS!TexasRoofDepreciation
                .b07AmountOfCheck = pRS!AmountOfCheck
                .b08Instructions = IIf(IsNull(pRS!Instructions), vbNullString, pRS!Instructions)
            End With
        Else
            Exit Sub
        End If
        
        LoadFarmUDT udtFCheck
        'cleanup
        pRS.Close
        Set pRS = Nothing
    End With
   
    Exit Sub
EH:
    Set pRS = Nothing
    ShowError Err, "Public Sub LoadFarmChecksUDT", , "ECWeb.clsFarmChecks.cls"
End Sub

Public Sub PopulateFarmCheck(pfrmEditChecks As Object, Optional pbInitValues As Boolean = False, Optional pCurDir As String)
    On Error GoTo EH
    Dim iCount As Integer
    Dim sValue As String
    Dim MyText As Control
    Dim saryItems() As String
    
    If mbPopulate Then
        Exit Sub
    End If
    
    mbPopulate = True
    
    With pfrmEditChecks
        
        'BGS 10.2.2001 Load_Cause_Codes
        sValue = pCurDir & "\LossCauses.TXT"
        If Dir$(sValue) = vbNullString Then
            MsgBox "Can't find loss causes.  Please setup using File->Loss Causes"
        Else
            saryItems = Split(GetFileData(sValue), vbCrLf)
            For Each MyText In .Controls
                If InStr(1, MyText.Tag, "LossCauses", vbTextCompare) > 0 Then
                    MyText.Clear
                    For iCount = 0 To UBound(saryItems)
                        MyText.AddItem saryItems(iCount)
                    Next
                End If
            Next
        End If
        
        'BGS 10.2.2001 Load_Class_Codes
        For Each MyText In .Controls
            If InStr(1, MyText.Tag, "ClassCodes", vbTextCompare) > 0 Then
                With MyText
                    .Clear
                    .AddItem "01 - Building"
                    .AddItem "02 - Contents"
                    .AddItem "89 - ALE"
                    .AddItem "15 - Commercial Building"
                End With
            End If
        Next
        
        If pbInitValues Then
            .cmbCheckNum.Text = msz03CheckNUM
            .chkRebill.Value = vbUnchecked
            .ChkChekReq(FChkReq.a01ForProperty).Value = vbUnchecked
            .ChkChekReq(FChkReq.a02ForAuto).Value = vbUnchecked
            .ChkChekReq(FChkReq.a03FinalPayment).Value = vbUnchecked
            .ChkChekReq(FChkReq.a04PartialPayment).Value = vbUnchecked
            .ChkChekReq(FChkReq.a05ForAddSupplement).Value = vbUnchecked
            .ChkChekReq(FChkReq.a06PaymentAssociated).Value = vbUnchecked
            .ChkChekReq(FChkReq.a07FieldHandled).Value = vbUnchecked
            .ChkChekReq(FChkReq.a08TotalLoss).Value = vbUnchecked
            .ChkChekReq(FChkReq.a09CashInLieu).Value = vbUnchecked
            .ChkChekReq(FChkReq.a10OwnerRetainsSalvage).Value = vbUnchecked
            .ChkChekReq(FChkReq.a11Subrogation).Value = vbUnchecked
            .ChkChekReq(FChkReq.a12Salvage).Value = vbUnchecked
            .ChkChekReq(FChkReq.a13IncludeMortgageeOnDraft).Value = vbUnchecked
            .cmbClassCheck.ListIndex = -1
            .cmbCauseOfLossCheck.ListIndex = -1
            .txtTexSubCovCode.Text = vbNullString
            .txtTexSuffix.Text = vbNullString
            .txtTexAmountOfRoofDepr.Text = vbNullString
            .txtCheckAmount.Text = vbNullString
            .txtInstructions.Text = vbNullString
            'bgs show disbled colors
            .framEditCheckRequests.Enabled = False
            For Each MyText In .Controls
                If InStr(1, MyText.Tag, "ColorMe") > 0 Then
                    If MyText.Container.Name = "framEditCheckRequests" And Not TypeOf MyText Is Label And Not TypeOf MyText Is CheckBox Then
                        MyText.BackColor = &H80000016  ' Change Back color dark
                    End If
                    If MyText.Container.Name = "framEditCheckRequests" And TypeOf MyText Is CheckBox Then
                        MyText.Enabled = False
                    End If
                End If
            Next
        Else
            .cmbCheckNum.Text = msz03CheckNUM
            .chkRebill.Value = IIf(mbz04Rebilled, vbChecked, vbUnchecked)
            .ChkChekReq(FChkReq.a01ForProperty).Value = IIf(mba01ForProperty, vbChecked, vbUnchecked)
            .ChkChekReq(FChkReq.a02ForAuto).Value = IIf(mba02ForAuto, vbChecked, vbUnchecked)
            .ChkChekReq(FChkReq.a03FinalPayment).Value = IIf(mba03FinalPayment, vbChecked, vbUnchecked)
            .ChkChekReq(FChkReq.a04PartialPayment).Value = IIf(mba04PartialPayment, vbChecked, vbUnchecked)
            .ChkChekReq(FChkReq.a05ForAddSupplement).Value = IIf(mba05ForAdditionalSupplement, vbChecked, vbUnchecked)
            .ChkChekReq(FChkReq.a06PaymentAssociated).Value = IIf(mba06PaymentAssociated, vbChecked, vbUnchecked)
            .ChkChekReq(FChkReq.a07FieldHandled).Value = IIf(mba07FieldHandled, vbChecked, vbUnchecked)
            .ChkChekReq(FChkReq.a08TotalLoss).Value = IIf(mba08TotalLoss, vbChecked, vbUnchecked)
            .ChkChekReq(FChkReq.a09CashInLieu).Value = IIf(mba09CashInLieu, vbChecked, vbUnchecked)
            .ChkChekReq(FChkReq.a10OwnerRetainsSalvage).Value = IIf(mba10OwnerRetainedSalvage, vbChecked, vbUnchecked)
            .ChkChekReq(FChkReq.a11Subrogation).Value = IIf(mba11Subrogation, vbChecked, vbUnchecked)
            .ChkChekReq(FChkReq.a12Salvage).Value = IIf(mba12Salvage, vbChecked, vbUnchecked)
            .ChkChekReq(FChkReq.a13IncludeMortgageeOnDraft).Value = IIf(mba13IncludeMortgageeOnDraft, vbChecked, vbUnchecked)
            .cmbClassCheck.Text = msb01Class
            .cmbCauseOfLossCheck.Text = msb03CauseOfLoss
            .txtTexSubCovCode.Text = msb04TexasSubCode
            .txtTexSuffix.Text = msb05TexasSuffix
            .txtTexAmountOfRoofDepr.Text = mcurb06TexasRoofDepreciation
            .txtCheckAmount.Text = mcurb07AmountOfCheck
            .txtInstructions.Text = msb08Instructions
            'bgs show enabled colors
            .framEditCheckRequests.Enabled = True
            For Each MyText In .Controls
                If InStr(1, MyText.Tag, "ColorMe") > 0 Then
                    If MyText.Container.Name = "framEditCheckRequests" And Not TypeOf MyText Is Label And Not TypeOf MyText Is CheckBox Then
                        MyText.BackColor = &H80000005 ' Change Back color light
                    End If
                    If MyText.Container.Name = "framEditCheckRequests" And TypeOf MyText Is CheckBox Then
                        MyText.Enabled = True
                    End If
                End If
            Next
        End If
    End With
    
CLEANUP:
    mbPopulate = False
    Exit Sub
EH:
    If Err.Number = CHECK_NOT_CREATED_ERROR Then
        Resume Next
    End If
    Close #1
    ShowError Err, "Public Sub PopulateFarmCheck", , "ECWeb.clsFarmChecks.cls"
End Sub

Public Sub EnableSave(pfrmEditChecks As Object)
    On Error GoTo EH
    If IsDirty(pfrmEditChecks) And Not Populate Then
        pfrmEditChecks.cmdSaveCheck.Enabled = True
        mbSaveEnabled = True
        Set mfrmSaveForm = pfrmEditChecks
    Else
        pfrmEditChecks.cmdSaveCheck.Enabled = False
        mbSaveEnabled = False
        Set mfrmSaveForm = Nothing
    End If
    Exit Sub
EH:
    ShowError Err, "Public Sub EnableSave", , "ECWeb.clsFarmChecks"
End Sub

Public Function SaveCheck(pfrmEditChecks As Object, pRS As Recordset, pCurDB As Database) As Boolean
    On Error GoTo EH
    Dim sSQL As String
    
    SaveCheck = True
    
    sSQL = "SELECT * FROM Checks A "
    sSQL = sSQL & "WHERE A.ID = " & mlz01ID & " "
    
    Set pRS = pCurDB.OpenRecordset(sSQL)
    
    If Not pRS.EOF Then
        pRS.MoveFirst
        With pfrmEditChecks
            pRS.Edit
            pRS!Rebilled = IIf(.chkRebill.Value = vbUnchecked, False, True)
            pRS!ForProperty = IIf(.ChkChekReq(FChkReq.a01ForProperty).Value = vbUnchecked, False, True)
            pRS!ForAuto = IIf(.ChkChekReq(FChkReq.a02ForAuto).Value = vbUnchecked, False, True)
            pRS!FinalPayment = IIf(.ChkChekReq(FChkReq.a03FinalPayment).Value = vbUnchecked, False, True)
            pRS!PartialPayment = IIf(.ChkChekReq(FChkReq.a04PartialPayment).Value = vbUnchecked, False, True)
            pRS!ForAdditionalSupplement = IIf(.ChkChekReq(FChkReq.a05ForAddSupplement).Value = vbUnchecked, False, True)
            pRS!PaymentAssociated = IIf(.ChkChekReq(FChkReq.a06PaymentAssociated).Value = vbUnchecked, False, True)
            pRS!FieldHandled = IIf(.ChkChekReq(FChkReq.a07FieldHandled).Value = vbUnchecked, False, True)
            pRS!TotalLoss = IIf(.ChkChekReq(FChkReq.a08TotalLoss).Value = vbUnchecked, False, True)
            pRS!CashInLieu = IIf(.ChkChekReq(FChkReq.a09CashInLieu).Value = vbUnchecked, False, True)
            pRS!OwnerRetainedSalvage = IIf(.ChkChekReq(FChkReq.a10OwnerRetainsSalvage).Value = vbUnchecked, False, True)
            pRS!Subrogation = IIf(.ChkChekReq(FChkReq.a11Subrogation).Value = vbUnchecked, False, True)
            pRS!Salvage = IIf(.ChkChekReq(FChkReq.a12Salvage).Value = vbUnchecked, False, True)
            pRS!IncludeMorgOnDraft = IIf(.ChkChekReq(FChkReq.a13IncludeMortgageeOnDraft).Value = vbUnchecked, False, True)
            pRS!Class = .cmbClassCheck.Text
            pRS!TypeOfLoss = .cmbClassCheck.Text
            pRS!CauseOfLoss = .cmbCauseOfLossCheck.Text
            pRS!TexasSubCode = .txtTexSubCovCode.Text
            pRS!TexasSuffix = .txtTexSuffix.Text
            pRS!TexasRoofDepreciation = .txtTexAmountOfRoofDepr.Text
            pRS!AmountOfCheck = .txtCheckAmount.Text
            pRS!Instructions = .txtInstructions.Text
            pRS.Update
        End With
    Else
        SaveCheck = False
        Exit Function
    End If
    
    'clean up
    pRS.Close
    Set pRS = Nothing
    
    Exit Function
EH:
    Set pRS = Nothing
    SaveCheck = False
    ShowError Err, "Public Function SaveCheck", , "ECWeb.clsFarmChecks"
End Function

Private Sub Class_Initialize()
'    Set modMain.clsFarmChecks = Me
End Sub

