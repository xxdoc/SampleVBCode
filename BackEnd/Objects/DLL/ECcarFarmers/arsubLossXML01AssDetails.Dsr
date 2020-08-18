VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} arsubLossXML01AssDetails 
   Caption         =   "ActiveReport1"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19606
   SectionData     =   "arsubLossXML01AssDetails.dsx":0000
End
Attribute VB_Name = "arsubLossXML01AssDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mAssDetailRS As WDDXRecordset
Private mlCount As Long
' Detail Heights
Private Const SHOW_DETAIL_HEIGHT As Long = 660
Private Const HIDE_DETAIL_HEIGHT As Long = 270
' Detail Heights

Public Property Let AssDetailRS(pAssDetailRS As WDDXRecordset)
    Set mAssDetailRS = pAssDetailRS
End Property
Public Property Set AssDetailRS(pAssDetailRS As WDDXRecordset)
    Set mAssDetailRS = pAssDetailRS
End Property

Private Sub ActiveReport_ReportEnd()
    Set mAssDetailRS = Nothing
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
    Dim AssDetail As udtXML01AssignmentDetail
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    If mAssDetailRS Is Nothing Then
       Exit Sub
    End If
    
    If mlCount <= mAssDetailRS.getRowCount Then
        'Populate all the text fields with the main udt
        '1. Populate Assignments Detail
        With AssDetail
            .UnitNumber = IIf(IsNull(mAssDetailRS.getField(mlCount, "UnitNumber")), vbNullString, mAssDetailRS.getField(mlCount, "UnitNumber"))
            .CatastropheCode = IIf(IsNull(mAssDetailRS.getField(mlCount, "CatastropheCode")), vbNullString, mAssDetailRS.getField(mlCount, "CatastropheCode"))
            .LossDate = IIf(IsNull(mAssDetailRS.getField(mlCount, "LossDate")), vbNullString, mAssDetailRS.getField(mlCount, "LossDate"))
            .Type = IIf(IsNull(mAssDetailRS.getField(mlCount, "Type")), vbNullString, mAssDetailRS.getField(mlCount, "Type"))
            .CauseOfLoss = IIf(IsNull(mAssDetailRS.getField(mlCount, "CauseOfLoss")), vbNullString, mAssDetailRS.getField(mlCount, "CauseOfLoss"))
            .FirstName = IIf(IsNull(mAssDetailRS.getField(mlCount, "FirstName")), vbNullString, mAssDetailRS.getField(mlCount, "FirstName"))
            .LastName = IIf(IsNull(mAssDetailRS.getField(mlCount, "LastName")), vbNullString, mAssDetailRS.getField(mlCount, "LastName"))
            .AssignedTo = IIf(IsNull(mAssDetailRS.getField(mlCount, "AssignedTo")), vbNullString, mAssDetailRS.getField(mlCount, "AssignedTo"))
            .AssignedToFirstName = IIf(IsNull(mAssDetailRS.getField(mlCount, "AssignedToFirstName")), vbNullString, mAssDetailRS.getField(mlCount, "AssignedToFirstName"))
            .AssignedToLastName = IIf(IsNull(mAssDetailRS.getField(mlCount, "AssignedToLastName")), vbNullString, mAssDetailRS.getField(mlCount, "AssignedToLastName"))
            
            f_UnitNumber.Text = .UnitNumber
            f_CatastropheCode.Text = .CatastropheCode
            f_LossDate.Text = Format(.LossDate, "mm/dd/yy")
            f_Type.Text = AssDetail.Type
            f_CauseOfLoss.Text = .CauseOfLoss
            'First check to see if this record is a subsequent Assignment detail record (Another Farmers Unit)
            'Only show the first line since the second line will always be duplicate.
            Detail.Height = SHOW_DETAIL_HEIGHT
            lblcliCat.Visible = True
            f_CatastropheCode.Visible = True
            If mlCount > 1 Then
                Detail.Height = HIDE_DETAIL_HEIGHT
                lblcliCat.Visible = False
            f_CatastropheCode.Visible = False
            End If
            f_FirstName.Text = .FirstName
            f_LastName.Text = .LastName
            f_AssignedTo.Text = .AssignedTo
            f_AssignedToFirstName.Text = .AssignedToFirstName
            f_AssignedToLastName.Text = .AssignedToLastName
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




