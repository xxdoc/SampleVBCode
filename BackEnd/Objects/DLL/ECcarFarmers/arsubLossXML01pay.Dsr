VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} arsubLossXML01pay 
   Caption         =   "ActiveReport1"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19606
   SectionData     =   "arsubLossXML01pay.dsx":0000
End
Attribute VB_Name = "arsubLossXML01pay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mPaymentDetailRS As WDDXRecordset
Private mlCount As Long

Public Property Let PaymentDetailRS(pPaymentDetailRS As WDDXRecordset)
    Set mPaymentDetailRS = pPaymentDetailRS
End Property
Public Property Set PaymentDetailRS(pPaymentDetailRS As WDDXRecordset)
    Set mPaymentDetailRS = pPaymentDetailRS
End Property

Private Sub ActiveReport_ReportEnd()
    Set mPaymentDetailRS = Nothing
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
    Dim PAY As udtXML01PaymentDetail
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    If mPaymentDetailRS Is Nothing Then
       Exit Sub
    End If
    
    If mlCount <= mPaymentDetailRS.getRowCount Then
        With PAY
            .DateIssued = IIf(IsNull(mPaymentDetailRS.getField(mlCount, "DateIssued")), vbNullString, mPaymentDetailRS.getField(mlCount, "DateIssued"))
            .PayeeLineOne = IIf(IsNull(mPaymentDetailRS.getField(mlCount, "PayeeLineOne")), vbNullString, mPaymentDetailRS.getField(mlCount, "PayeeLineOne"))
            .PayeeLineTwo = IIf(IsNull(mPaymentDetailRS.getField(mlCount, "PayeeLineTwo")), vbNullString, mPaymentDetailRS.getField(mlCount, "PayeeLineTwo"))
            .PayeeLineThree = IIf(IsNull(mPaymentDetailRS.getField(mlCount, "PayeeLineThree")), vbNullString, mPaymentDetailRS.getField(mlCount, "PayeeLineThree"))
            .PayeeLineFour = IIf(IsNull(mPaymentDetailRS.getField(mlCount, "PayeeLineFour")), vbNullString, mPaymentDetailRS.getField(mlCount, "PayeeLineFour"))
            .AccountType = IIf(IsNull(mPaymentDetailRS.getField(mlCount, "AccountType")), vbNullString, mPaymentDetailRS.getField(mlCount, "AccountType"))
            .PaymentClass = IIf(IsNull(mPaymentDetailRS.getField(mlCount, "PaymentClass")), vbNullString, mPaymentDetailRS.getField(mlCount, "PaymentClass"))
            .PaymentAmount = IIf(IsNull(mPaymentDetailRS.getField(mlCount, "PaymentAmount")), vbNullString, mPaymentDetailRS.getField(mlCount, "PaymentAmount"))
        End With
        f_DateIssued.Text = PAY.DateIssued
        f_PayeeLineOne.Text = PAY.PayeeLineOne
        f_PayeeLineTwo.Text = PAY.PayeeLineTwo
        f_PayeeLineThree.Text = PAY.PayeeLineThree
        f_PayeeLineFour.Text = PAY.PayeeLineFour
        f_AccountType.Text = PAY.AccountType
        f_PaymentClass.Text = PAY.PaymentClass
        f_PaymentAmount.Text = Format(PAY.PaymentAmount, "0.00")
        
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


