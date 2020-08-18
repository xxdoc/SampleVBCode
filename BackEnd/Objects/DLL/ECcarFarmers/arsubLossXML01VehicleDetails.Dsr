VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} arsubLossXML01VehicleDetails 
   Caption         =   "ActiveReport3"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19606
   SectionData     =   "arsubLossXML01VehicleDetails.dsx":0000
End
Attribute VB_Name = "arsubLossXML01VehicleDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mVehicleDetailRS As WDDXRecordset
Private mlCount As Long

Public Property Let VehicleDetailRS(pVehicleDetailRS As WDDXRecordset)
    Set mVehicleDetailRS = pVehicleDetailRS
End Property
Public Property Set VehicleDetailRS(pVehicleDetailRS As WDDXRecordset)
    Set mVehicleDetailRS = pVehicleDetailRS
End Property

Private Sub ActiveReport_ReportEnd()
    Set mVehicleDetailRS = Nothing
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
    Dim VehicleDetail As udtXML01VehicleDetail
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    If mVehicleDetailRS Is Nothing Then
       Exit Sub
    End If
    
    If mlCount <= mVehicleDetailRS.getRowCount Then
        'Populate all the text fields with the main udt
        '2. Populate Vehicle Detail
        With VehicleDetail
            .VehicleMake = IIf(IsNull(mVehicleDetailRS.getField(mlCount, "VehicleMake")), vbNullString, mVehicleDetailRS.getField(mlCount, "VehicleMake"))
            .VehicleModel = IIf(IsNull(mVehicleDetailRS.getField(mlCount, "VehicleModel")), vbNullString, mVehicleDetailRS.getField(mlCount, "VehicleModel"))
            .VehicleYear = IIf(IsNull(mVehicleDetailRS.getField(mlCount, "VehicleYear")), vbNullString, mVehicleDetailRS.getField(mlCount, "VehicleYear"))
            .VIN = IIf(IsNull(mVehicleDetailRS.getField(mlCount, "VIN")), vbNullString, mVehicleDetailRS.getField(mlCount, "VIN"))
            .PropertyItemName = IIf(IsNull(mVehicleDetailRS.getField(mlCount, "PropertyItemName")), vbNullString, mVehicleDetailRS.getField(mlCount, "PropertyItemName"))
            .DamageDescription = IIf(IsNull(mVehicleDetailRS.getField(mlCount, "DamageDescription")), vbNullString, mVehicleDetailRS.getField(mlCount, "DamageDescription"))
            .LossDescription = IIf(IsNull(mVehicleDetailRS.getField(mlCount, "LossDescription")), vbNullString, mVehicleDetailRS.getField(mlCount, "LossDescription"))
            .LocationType = IIf(IsNull(mVehicleDetailRS.getField(mlCount, "LocationType")), vbNullString, mVehicleDetailRS.getField(mlCount, "LocationType"))
            .LocationName = IIf(IsNull(mVehicleDetailRS.getField(mlCount, "LocationName")), vbNullString, mVehicleDetailRS.getField(mlCount, "LocationName"))
            .LocationPhoneNumber = IIf(IsNull(mVehicleDetailRS.getField(mlCount, "LocationPhoneNumber")), vbNullString, mVehicleDetailRS.getField(mlCount, "LocationPhoneNumber"))
            .LocationAddress = IIf(IsNull(mVehicleDetailRS.getField(mlCount, "LocationAddress")), vbNullString, mVehicleDetailRS.getField(mlCount, "LocationAddress"))
            .LocationCity = IIf(IsNull(mVehicleDetailRS.getField(mlCount, "LocationCity")), vbNullString, mVehicleDetailRS.getField(mlCount, "LocationCity"))
            .LocationState = IIf(IsNull(mVehicleDetailRS.getField(mlCount, "LocationState")), vbNullString, mVehicleDetailRS.getField(mlCount, "LocationState"))
            .LocationPostalCode = IIf(IsNull(mVehicleDetailRS.getField(mlCount, "LocationPostalCode")), vbNullString, mVehicleDetailRS.getField(mlCount, "LocationPostalCode"))
            
            f_VehicleMake.Text = .VehicleMake
            f_VehicleModel.Text = .VehicleModel
            f_VehicleYear.Text = .VehicleYear
            f_VIN.Text = .VIN
            f_PropertyItemName.Text = .PropertyItemName
            f_DamageDescription.Text = .DamageDescription
            f_LossDescription.Text = .LossDescription
            f_LocationType.Text = .LocationType
            f_LocationName.Text = .LocationName
            f_LocationPhoneNumber.Text = .LocationPhoneNumber
            f_LocationAddress.Text = .LocationAddress
            f_LocationCity.Text = .LocationCity
            f_LocationState.Text = .LocationState
            f_LocationPostalCode.Text = .LocationPostalCode
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







