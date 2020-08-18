VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} arIndemnityReport 
   Caption         =   "ActiveReport2"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   14975
   SectionData     =   "arIndemnityReport.dsx":0000
End
Attribute VB_Name = "arIndemnityReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private moLists As V2ECKeyBoard.clsCarLists
Private mcolProperty As Collection
Private mcolPropertyKeys As Collection

'BGS 12.27.2000 Need this to see if the Report is still active.
Private mbActiveFlag As Boolean
Private mbStretch As Boolean


'Chain Reports
'Private mbChainPgBrk As Boolean
Private mbChainFlag As Boolean
Private mlChainCount As Long
Private mcolChainReports As Collection ' Contains Reports Chained to it to be added to the Sub Report object
Private moChainReport As Object
Private mcolIndemnity01 As Collection
Private mlIndemnityCount As Long
Private mlMaxCount As Long

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

Public Property Let Stretch(pbFlag As Boolean)
    mbStretch = pbFlag
End Property
Public Property Get Stretch() As Boolean
    Stretch = mbStretch
End Property

Public Property Get ActiveFlag() As Boolean
    ActiveFlag = mbActiveFlag
End Property
Public Property Let ActiveFlag(pbFlag As Boolean)
    mbActiveFlag = pbFlag
End Property

Public Property Get ClassName() As String
    ClassName = App.EXEName & "." & Me.Name
End Property

Private Sub ActiveReport_Initialize()
    mbActiveFlag = True
End Sub

Public Sub SetProperty(psName As String, pvValue As Variant, pType As VbVarType)
    On Error GoTo EH
    Dim sValue As String
    Dim vNewValue As Variant
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    If IsNull(pvValue) Then
        pvValue = vbNullString
    End If
    If Not IsObject(pvValue) Then
        sValue = RTrim(CStr(pvValue))
        'Replace any carriage return flags with vbCrLf
        sValue = Replace(sValue, F_VBCRLF, vbCrLf)
    End If
    
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
        Case VbVarType.vbObject
            Set vNewValue = pvValue
    End Select
    
    If mcolProperty Is Nothing Then
        Set mcolProperty = New Collection
        Set mcolPropertyKeys = New Collection
    End If
    RemoveProperty psName
    mcolProperty.Add vNewValue, psName
    mcolPropertyKeys.Add psName, psName
    
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
        mcolPropertyKeys.Remove psName
        RemoveProperty = True
    End If
    Exit Function
EH:
    Err.Clear
End Function

Private Sub ActiveReport_ReportStart()
    On Error GoTo EH
    Dim oField As Object
    Dim sTemp As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim sCurTemp As String
    Dim dtMyDate As Date
    
     'Populate Report Header controls
    For Each oField In Me.ReportHeader.Controls
         If GetProperty(CStr(oField.Name)) <> vbNullString Then
            If oField.Tag = "CurFormat" Then
                sCurTemp = GetProperty(CStr(oField.Name))
                sCurTemp = Format(sCurTemp, "#,###,###,##0.00")
                oField.Text = sCurTemp
            ElseIf oField.Tag = "NumFormat" Then
                sCurTemp = GetProperty(CStr(oField.Name))
                sCurTemp = Format(sCurTemp, "#,###,###,##0")
                oField.Text = sCurTemp
            ElseIf oField.Tag = "PctFormat" Then
                sCurTemp = GetProperty(CStr(oField.Name))
                sCurTemp = Format(sCurTemp, "0%")
                oField.Text = sCurTemp
            ElseIf oField.Tag = "DateFormat" Then
                dtMyDate = GetProperty(CStr(oField.Name))
                If dtMyDate = NULL_DATE Then
                    oField.Text = vbNullString
                Else
                    oField.Text = dtMyDate
                End If
            Else
                oField.Text = GetProperty(CStr(oField.Name))
            End If
        End If
    Next
    
    'Get the Collection for detail section
    Set mcolIndemnity01 = mcolProperty("coludtIndemnity01")
    mlIndemnityCount = 1
    If Not mcolIndemnity01 Is Nothing Then
        mlMaxCount = mcolIndemnity01.Count
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

Private Sub ActiveReport_Terminate()
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    Set moLists = Nothing
    Set mcolProperty = Nothing
    Set mcolPropertyKeys = Nothing
    Set mcolChainReports = Nothing
    Set moChainReport = Nothing
    Set mcolIndemnity01 = Nothing
    
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
    Dim MyIndemnity As udtIndemnity01
         
    If mlIndemnityCount <= mlMaxCount Then
        MyIndemnity = mcolIndemnity01(mlIndemnityCount)
        With MyIndemnity
            f_Payment.Text = .f_Payment
            f_IsPreviousPayment.Text = IIf(.f_IsPreviousPayment, "X", vbNullString)
            f_ClassOfLoss.Text = .f_ClassOfLoss
            f_ClassOfLossCode.Text = .f_ClassOfLossCode
            f_TypeOfLoss.Text = .f_TypeOfLoss
            f_typeOfLossCode.Text = .f_typeOfLossCode
            f_Description.Text = .f_Description
            f_ReplacementCost.Text = Format(.f_ReplacementCost, "#,###,###,##0.00")
            f_RecoverableDep.Text = Format(.f_RecoverableDep, "#,###,###,##0.00")
            f_NonRecoverableDep.Text = Format(.f_NonRecoverableDep, "#,###,###,##0.00")
            f_AcvClaim.Text = Format(.f_AcvClaim, "#,###,###,##0.00")
            f_SpecialLimits.Text = Format(.f_SpecialLimits, "#,###,###,##0.00")
            f_IsAddAmountOfInsurance.Text = IIf(.f_IsAddAmountOfInsurance, "X", vbNullString)
            f_ExcessLimits.Text = Format(.f_ExcessLimits, "#,###,###,##0.00")
            f_Miscellaneous.Text = Format(.f_Miscellaneous, "#,###,###,##0.00")
            f_ACVLessExcessLimits.Text = Format(.f_ACVLessExcessLimits, "#,###,###,##0.00")
        End With
        Detail.PrintSection
        mlIndemnityCount = mlIndemnityCount + 1
    Else
        If Not moChainReport Is Nothing Then
            Me.Detail.NewPage = ddNPBefore
            subChain.Visible = True
        Else
            subChain.Visible = False
            Exit Sub 'Bail here!
        End If
    End If
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Private Sub Detail_Format"
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
'    Set mcolIndemnity01 = Nothing
    'Cleanup
    mbActiveFlag = False
    
    Exit Sub

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

Public Function GetXMLExport() As String
    On Error GoTo EH

    'Export Report Collection Items
    Dim oMySer As WDDXSerializer        'Allaire's WDDX serializer
    Dim oMyRS As WDDXRecordset          'Allaire's WDDX Recordset
    Dim oMyStruct As WDDXStruct         'Allaire's WDDX Structure (Cold Fusion Strucuture type)
    Dim lCount As Long
    Dim sColName As String
    Dim vValue As Variant 'Use Variant to Get Variant Equiv Data Type.
    'Indemnity Needs to include Collection of Indemnity Entries
    Dim MyIndemnity As udtIndemnity01
    Dim oIndemRS As WDDXRecordset
    
    'Make sure the Collection items exist
    If mcolProperty Is Nothing Or mcolPropertyKeys Is Nothing Then
        Exit Function
    End If
    If mcolProperty.Count = 0 Or mcolPropertyKeys.Count = 0 Then
        Exit Function
    End If
    
    'Create a WDDX RS
    Set oMyRS = New WDDXRecordset
    For lCount = 1 To mcolPropertyKeys.Count
        'Use the Keys Collection to create Column Names
        sColName = mcolPropertyKeys(lCount)
        'Do not Add the Collection of Logs
        If StrComp(sColName, "coludtIndemnity01", vbTextCompare) <> 0 Then
            oMyRS.addColumn sColName
        End If
    Next
    
    'Only one row for the Data RS
    oMyRS.addRows 1
    'Set the Col values for this one row
    For lCount = 1 To mcolProperty.Count
        sColName = mcolPropertyKeys(lCount)
        'Use Variant Value to Get Data type
        If StrComp(sColName, "coludtIndemnity01", vbTextCompare) <> 0 Then
            vValue = mcolProperty(lCount)
            oMyRS.setField 1, sColName, vValue
        End If
    Next

    '****BEGIN***Indemnity Needs to include Collection of Indemnity Entries***
    Set mcolIndemnity01 = mcolProperty("coludtIndemnity01")
    If Not mcolIndemnity01 Is Nothing Then
        If mcolIndemnity01.Count > 0 Then
            Set oIndemRS = New WDDXRecordset
            'Add Colmn names
            oIndemRS.addColumn "f_AcvClaim"
            oIndemRS.addColumn "f_ACVLessExcessLimits"
            oIndemRS.addColumn "f_ClassOfLoss"
            oIndemRS.addColumn "f_ClassOfLossCode"
            oIndemRS.addColumn "f_Description"
            oIndemRS.addColumn "f_ExcessLimits"
            oIndemRS.addColumn "f_IsAddAmountOfInsurance"
            oIndemRS.addColumn "f_IsPreviousPayment"
            oIndemRS.addColumn "f_Miscellaneous"
            oIndemRS.addColumn "f_NonRecoverableDep"
            oIndemRS.addColumn "f_Payment"
            oIndemRS.addColumn "f_RecoverableDep"
            oIndemRS.addColumn "f_ReplacementCost"
            oIndemRS.addColumn "f_RTIndemnityID"
            oIndemRS.addColumn "f_SpecialLimits"
            oIndemRS.addColumn "f_TypeOfLoss"
            oIndemRS.addColumn "f_typeOfLossCode"
            'Add the same number in collection
            oIndemRS.addRows mcolIndemnity01.Count
        End If
    
    End If
         
    For lCount = 1 To mcolIndemnity01.Count
        MyIndemnity = mcolIndemnity01(lCount)
        With MyIndemnity
            vValue = .f_AcvClaim
            oIndemRS.setField lCount, "f_AcvClaim", vValue
            vValue = .f_ACVLessExcessLimits
            oIndemRS.setField lCount, "f_ACVLessExcessLimits", vValue
            vValue = .f_ClassOfLoss
            oIndemRS.setField lCount, "f_ClassOfLoss", vValue
            vValue = .f_ClassOfLossCode
            oIndemRS.setField lCount, "f_ClassOfLossCode", vValue
            vValue = .f_Description
            oIndemRS.setField lCount, "f_Description", vValue
            vValue = .f_ExcessLimits
            oIndemRS.setField lCount, "f_ExcessLimits", vValue
            vValue = .f_IsAddAmountOfInsurance
            oIndemRS.setField lCount, "f_IsAddAmountOfInsurance", vValue
            vValue = .f_IsPreviousPayment
            oIndemRS.setField lCount, "f_IsPreviousPayment", vValue
            vValue = .f_Miscellaneous
            oIndemRS.setField lCount, "f_Miscellaneous", vValue
            vValue = .f_NonRecoverableDep
            oIndemRS.setField lCount, "f_NonRecoverableDep", vValue
            vValue = .f_Payment
            oIndemRS.setField lCount, "f_Payment", vValue
            vValue = .f_RecoverableDep
            oIndemRS.setField lCount, "f_RecoverableDep", vValue
            vValue = .f_ReplacementCost
            oIndemRS.setField lCount, "f_ReplacementCost", vValue
            vValue = .f_RTIndemnityID
            oIndemRS.setField lCount, "f_RTIndemnityID", vValue
            vValue = .f_SpecialLimits
            oIndemRS.setField lCount, "f_SpecialLimits", vValue
            vValue = .f_TypeOfLoss
            oIndemRS.setField lCount, "f_TypeOfLoss", vValue
            vValue = .f_typeOfLossCode
            oIndemRS.setField lCount, "f_typeOfLossCode", vValue
        End With
    Next
    '****END***Indemnity Needs to include Collection of Indemnity Entries***
    
    'Create WDDX Structure
    Set oMyStruct = New WDDXStruct
    
    oMyStruct.setProp "ClassName", ClassName
    oMyStruct.setProp "DataRS", oMyRS
    If Not oIndemRS Is Nothing Then
        oMyStruct.setProp "IndemRS", oIndemRS
    End If
    
    Set oMySer = New WDDXSerializer
    
    GetXMLExport = oMySer.serialize(oMyStruct)
    
    'Cleanup
    Set oMyRS = Nothing
    Set oIndemRS = Nothing
    Set oMyStruct = Nothing
    Set oMySer = Nothing
    
    Exit Function
EH:
    GetXMLExport = "Class Name: " & ClassName & vbCrLf
    GetXMLExport = GetXMLExport & "Error # " & Err.Number & vbCrLf
    GetXMLExport = GetXMLExport & "Description: " & vbCrLf
    GetXMLExport = GetXMLExport & Err.Description
End Function





