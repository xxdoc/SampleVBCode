VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} arCommission 
   Caption         =   "ActiveReport1"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19606
   SectionData     =   "arCommission.dsx":0000
End
Attribute VB_Name = "arCommission"
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
Private mcolCommission01 As Collection
Private mlCommissionCount As Long
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
     'Populate Report Header controls
    For Each oField In Me.ReportHeader.Controls
        If GetProperty(CStr(oField.Name)) <> vbNullString Then
            oField.Text = GetProperty(CStr(oField.Name))
        End If
    Next
    
    'Populate Report Footer controls
    For Each oField In Me.ReportFooter.Controls
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
            Else
                oField.Text = GetProperty(CStr(oField.Name))
            End If
        End If
    Next
    
    'Get the Collection of Photos for detail section
    Set mcolCommission01 = mcolProperty("coludtCommission01")
    mlCommissionCount = 1
    If Not mcolCommission01 Is Nothing Then
        mlMaxCount = mcolCommission01.Count
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
    Set mcolChainReports = Nothing
    Set moChainReport = Nothing
    Set mcolCommission01 = Nothing
    
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
    Dim MyCommision As udtCommission01
         
    If mlCommissionCount <= mlMaxCount Then
        MyCommision = mcolCommission01(mlCommissionCount)
        With MyCommision
                f_CommissionCount.Text = Format(mlCommissionCount, "0.")
                f_CLIENTNUM.Text = .f_CLIENTNUM
                f_Insured.Text = .f_Insured
                f_MiscExp.Text = Format(.f_MiscExp, "#,###,###,##0.00")
                f_ServiceFee.Text = Format(.f_ServiceFee, "#,###,###,##0.00")
                f_OtherFees.Text = Format(.f_OtherFees, "#,###,###,##0.00")
                f_Tax.Text = Format(.f_Tax, "#,###,###,##0.00")
                f_Total.Text = Format(.f_Total, "#,###,###,##0.00")
        End With
        Detail.PrintSection
        mlCommissionCount = mlCommissionCount + 1
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
'    Set mcolCommission01 = Nothing
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
    'Commission Report Needs to include Collection of Entries
    Dim MyCommision As udtCommission01
    Dim oComissionRS As WDDXRecordset
    
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
        If StrComp(sColName, "coludtCommission01", vbTextCompare) <> 0 Then
            oMyRS.addColumn sColName
        End If
    Next
    
    'Only one row for the Data RS
    oMyRS.addRows 1
    'Set the Col values for this one row
    For lCount = 1 To mcolProperty.Count
        sColName = mcolPropertyKeys(lCount)
        'Use Variant Value to Get Data type
        If StrComp(sColName, "coludtCommission01", vbTextCompare) <> 0 Then
            vValue = mcolProperty(lCount)
            oMyRS.setField 1, sColName, vValue
        End If
    Next

    '****BEGIN***Commission Report Needs to include Collection of Entries***
    Set mcolCommission01 = mcolProperty("coludtCommission01")
    If Not mcolCommission01 Is Nothing Then
        If mcolCommission01.Count > 0 Then
            Set oComissionRS = New WDDXRecordset
            'Add Colmn names
            oComissionRS.addColumn "AssignmentsID"
            oComissionRS.addColumn "f_CLIENTNUM"
            oComissionRS.addColumn "f_Insured"
            oComissionRS.addColumn "f_Mileage"
            oComissionRS.addColumn "f_MiscExp"
            oComissionRS.addColumn "f_OtherFees"
            oComissionRS.addColumn "f_OutBldg"
            oComissionRS.addColumn "f_Photos"
            oComissionRS.addColumn "f_ServiceFee"
            oComissionRS.addColumn "f_Tax"
            oComissionRS.addColumn "f_Total"
            'Add the same number in collection
            oComissionRS.addRows mcolCommission01.Count
        End If
    
    End If
         
    For lCount = 1 To mcolCommission01.Count
        MyCommision = mcolCommission01(lCount)
        With MyCommision
            vValue = .AssignmentsID
            oComissionRS.setField lCount, "AssignmentsID", vValue
            vValue = .f_CLIENTNUM
            oComissionRS.setField lCount, "f_CLIENTNUM", vValue
            vValue = .f_Insured
            oComissionRS.setField lCount, "f_Insured", vValue
            vValue = .f_Mileage
            oComissionRS.setField lCount, "f_Mileage", vValue
            vValue = .f_MiscExp
            oComissionRS.setField lCount, "f_MiscExp", vValue
            vValue = .f_OtherFees
            oComissionRS.setField lCount, "f_OtherFees", vValue
            vValue = .f_OutBldg
            oComissionRS.setField lCount, "f_OutBldg", vValue
            vValue = .f_Photos
            oComissionRS.setField lCount, "f_Photos", vValue
            vValue = .f_ServiceFee
            oComissionRS.setField lCount, "f_ServiceFee", vValue
            vValue = .f_Tax
            oComissionRS.setField lCount, "f_Tax", vValue
            vValue = .f_Total
            oComissionRS.setField lCount, "f_Total", vValue
        End With
    Next
    '****END***Commission Report Needs to include Collection of Entries***
    
    'Create WDDX Structure
    Set oMyStruct = New WDDXStruct
    
    oMyStruct.setProp "ClassName", ClassName
    oMyStruct.setProp "DataRS", oMyRS
    If Not oComissionRS Is Nothing Then
        oMyStruct.setProp "ComissionRS", oComissionRS
    End If
    
    Set oMySer = New WDDXSerializer
    
    GetXMLExport = oMySer.serialize(oMyStruct)
    
    'Cleanup
    Set oMyRS = Nothing
    Set oComissionRS = Nothing
    Set oMyStruct = Nothing
    Set oMySer = Nothing
    
    Exit Function
EH:
    GetXMLExport = "Class Name: " & ClassName & vbCrLf
    GetXMLExport = GetXMLExport & "Error # " & Err.Number & vbCrLf
    GetXMLExport = GetXMLExport & "Description: " & vbCrLf
    GetXMLExport = GetXMLExport & Err.Description
End Function
