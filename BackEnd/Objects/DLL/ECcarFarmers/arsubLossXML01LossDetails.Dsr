VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} arsubLossXML01LossDetails 
   Caption         =   "ActiveReport2"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19606
   SectionData     =   "arsubLossXML01LossDetails.dsx":0000
End
Attribute VB_Name = "arsubLossXML01LossDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mLossDetailRS As WDDXRecordset
Private mlCount As Long

Public Property Let LossDetailRS(pLossDetailRS As WDDXRecordset)
    Set mLossDetailRS = pLossDetailRS
End Property
Public Property Set LossDetailRS(pLossDetailRS As WDDXRecordset)
    Set mLossDetailRS = pLossDetailRS
End Property

Private Sub ActiveReport_ReportEnd()
    Set mLossDetailRS = Nothing
End Sub

Private Sub ActiveReport_ReportStart()
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    mlCount = 1
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & Me.Name & vbCrLf & "Private Sub ActiveReport_ReportStart"
End Sub

Private Sub Detail_Format()
    On Error GoTo EH
    Dim LossDetail As udtXML01LossDetail
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    If mLossDetailRS Is Nothing Then
       Exit Sub
    End If
    
    If mlCount <= mLossDetailRS.getRowCount Then
        'Populate all the text fields with the main udt
        '2. Populate Loss Detail For Property
        With LossDetail
            .LossLocationAddress = IIf(IsNull(mLossDetailRS.getField(mlCount, "LossLocationAddress")), vbNullString, mLossDetailRS.getField(mlCount, "LossLocationAddress"))
            .LossLocationAddress2 = IIf(IsNull(mLossDetailRS.getField(mlCount, "LossLocationAddress2")), vbNullString, mLossDetailRS.getField(mlCount, "LossLocationAddress2"))
            .LossLocationCity = IIf(IsNull(mLossDetailRS.getField(mlCount, "LossLocationCity")), vbNullString, mLossDetailRS.getField(mlCount, "LossLocationCity"))
            .LossLocationState = IIf(IsNull(mLossDetailRS.getField(mlCount, "LossLocationState")), vbNullString, mLossDetailRS.getField(mlCount, "LossLocationState"))
            .LossLocationZip = IIf(IsNull(mLossDetailRS.getField(mlCount, "LossLocationZip")), vbNullString, mLossDetailRS.getField(mlCount, "LossLocationZip"))
            .PropertyAddress = IIf(IsNull(mLossDetailRS.getField(mlCount, "PropertyAddress")), vbNullString, mLossDetailRS.getField(mlCount, "PropertyAddress"))
            .PropertyCity = IIf(IsNull(mLossDetailRS.getField(mlCount, "PropertyCity")), vbNullString, mLossDetailRS.getField(mlCount, "PropertyCity"))
            .PropertyState = IIf(IsNull(mLossDetailRS.getField(mlCount, "PropertyState")), vbNullString, mLossDetailRS.getField(mlCount, "PropertyState"))
            .PropertyZip = IIf(IsNull(mLossDetailRS.getField(mlCount, "PropertyZip")), vbNullString, mLossDetailRS.getField(mlCount, "PropertyZip"))
            .AffectedAreas = IIf(IsNull(mLossDetailRS.getField(mlCount, "AffectedAreas")), vbNullString, mLossDetailRS.getField(mlCount, "AffectedAreas"))
            .LossDescription = IIf(IsNull(mLossDetailRS.getField(mlCount, "LossDescription")), vbNullString, mLossDetailRS.getField(mlCount, "LossDescription"))
            
            f_LossLocationAddress.Text = .LossLocationAddress
            f_LossLocationAddress2.Text = .LossLocationAddress2
            f_LossLocationCity.Text = .LossLocationCity
            f_LossLocationState.Text = .LossLocationState
            f_LossLocationZip.Text = .LossLocationZip
            f_PropertyAddress.Text = .PropertyAddress
            f_PropertyCity.Text = .PropertyCity
            f_PropertyState.Text = .PropertyState
            f_PropertyZip.Text = .PropertyZip
            f_AffectedAreas.Text = .AffectedAreas
            f_LossDescription.Text = .LossDescription
            
        End With
        
        mlCount = mlCount + 1
        If mlCount Mod 2 = 1 Then
            Detail.BackColor = &HE0E0E0
        Else
            Detail.BackColor = &HFFFFFF
        End If
        Detail.PrintSection
    End If
    
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & Me.Name & vbCrLf & "Private Sub Detail_Format"
End Sub






