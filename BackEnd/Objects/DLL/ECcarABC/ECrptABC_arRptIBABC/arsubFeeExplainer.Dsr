VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} arsubFeeExplainer 
   Caption         =   "Fee Explainer"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19606
   SectionData     =   "arsubFeeExplainer.dsx":0000
End
Attribute VB_Name = "arsubFeeExplainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mcoludtFeeItems01 As Collection
Private msServiceFeeComment As String
Private msMiscServiceFeeComment As String
Private msMiscExpenseFeeComment As String
Private mcServiceFee As Currency
Private mcMiscServiceFee As Currency
Private mcMiscExpenseFee As Currency
Private mlCount As Long
Private mcServiceFeeSubTtl As Currency
Private mcExpenseFeeSubTtl As Currency
Private mcTotalServiceAndExpense As Currency
Private mdblTaxPercent As Double
Private mcTaxAmount As Currency
Private mcTotalAdjustingFee As Currency
Private msAccountCode As String
Private msubSubTtlServiceFee As arSubTtlServiceFee
Private msubSubTtlExpenseFee As arSubTtlExpenseFee
Private Const DETAIL_H As Long = 480
Private Const DETAIL_H_SUBTTL = 690
Private mbShowedServiceFeeSubTtl As Boolean
Private mbShowedExpenseFeeSubTtl As Boolean

Public Property Let AccountCode(psValue As String)
    msAccountCode = psValue
End Property

Public Property Let TaxPercent(pdblValue As Double)
    mdblTaxPercent = pdblValue
End Property

Public Property Let TaxAmount(pcValue As Currency)
    mcTaxAmount = pcValue
End Property

Public Property Let TotalAdjustingFee(pcValue As Currency)
    mcTotalAdjustingFee = pcValue
End Property

Public Property Let FeeItemsCol(pcolFeeItems As Object)
    Set mcoludtFeeItems01 = pcolFeeItems
End Property
Public Property Set FeeItemsCol(pcolFeeItems As Object)
    Set mcoludtFeeItems01 = pcolFeeItems
End Property

Public Property Let ServiceFeeComment(psValue As String)
    msServiceFeeComment = psValue
End Property

Public Property Let MiscServiceFeeComment(psValue As String)
    msMiscServiceFeeComment = psValue
End Property

Public Property Let MiscExpenseFeeComment(psValue As String)
    msMiscExpenseFeeComment = psValue
End Property

Public Property Let ServiceFee(pcValue As String)
    mcServiceFee = pcValue
End Property

Public Property Let MiscServiceFee(pcValue As Currency)
    mcMiscServiceFee = pcValue
End Property

Public Property Let MiscExpenseFee(pcValue As Currency)
    mcMiscExpenseFee = pcValue
End Property

Private Sub ActiveReport_ReportStart()
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim lCount As Long
    Dim MyFeeItem  As udtFeeItems01
    
    'get the Subtotals first
    If mcoludtFeeItems01 Is Nothing Then
        Exit Sub
    End If
    If mcoludtFeeItems01.Count = 0 Then
        Exit Sub
    End If
    
    For lCount = 1 To mcoludtFeeItems01.Count
        MyFeeItem = mcoludtFeeItems01(lCount)
        If Not MyFeeItem.fsftIsExpense And InStr(1, MyFeeItem.fsftVBFormula, "OVERRIDES", vbTextCompare) = 0 Then
            mcServiceFeeSubTtl = mcServiceFeeSubTtl + MyFeeItem.Amount
        ElseIf MyFeeItem.fsftIsExpense Then
            mcExpenseFeeSubTtl = mcExpenseFeeSubTtl + MyFeeItem.Amount
        End If
    Next
    
    Set msubSubTtlServiceFee = New arSubTtlServiceFee
    Set msubSubTtlExpenseFee = New arSubTtlExpenseFee
    
    'Populate the Fields for the subtotals
    
    'Service Fee Subtotals
    With msubSubTtlServiceFee
        .f_AddServiceFee.Text = Format(mcServiceFeeSubTtl, "$#,###,###,##0.00")
        .f_MiscServiceFeeComment.Text = msMiscServiceFeeComment
        .f_MiscServiceFee.Text = Format(mcMiscServiceFee, "$#,###,###,##0.00")
        mcServiceFeeSubTtl = mcServiceFee + mcServiceFeeSubTtl + mcMiscServiceFee
        mcTotalServiceAndExpense = mcTotalServiceAndExpense + mcServiceFeeSubTtl
        .f_ServiceFeeSubTtl.Text = Format(mcServiceFeeSubTtl, "$#,###,###,##0.00")
    End With
    
    
    'Expesne Fee Subtotals
    With msubSubTtlExpenseFee
        .f_AddExpenseFee.Text = Format(mcExpenseFeeSubTtl, "$#,###,###,##0.00")
        .f_MiscExpenseFeeComment.Text = msMiscExpenseFeeComment
        .f_MiscExpenseFee.Text = Format(mcMiscExpenseFee, "$#,###,###,##0.00")
        mcExpenseFeeSubTtl = mcExpenseFeeSubTtl + mcMiscExpenseFee
        mcTotalServiceAndExpense = mcTotalServiceAndExpense + mcExpenseFeeSubTtl
        .f_ExpenseFeeSubTtl.Text = Format(mcExpenseFeeSubTtl, "$#,###,###,##0.00")
    End With
    
    mlCount = 1
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & Me.Name & vbCrLf & "Private Sub ActiveReport_ReportStart"
End Sub

Private Sub ActiveReport_ReportEnd()
    Set mcoludtFeeItems01 = Nothing
    Set subSubTtl = Nothing
    Set msubSubTtlServiceFee = Nothing
    Set msubSubTtlExpenseFee = Nothing
End Sub

Private Sub Detail_Format()
    On Error GoTo EH
    Dim MyFeeItem  As udtFeeItems01
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    If mcoludtFeeItems01 Is Nothing Then
        Exit Sub
    End If
    
    If mlCount <= mcoludtFeeItems01.Count Then
        MyFeeItem = mcoludtFeeItems01(mlCount)
        
        If Not MyFeeItem.fsftIsExpense And InStr(1, MyFeeItem.fsftVBFormula, "OVERRIDES", vbTextCompare) = 0 Then
            mcServiceFeeSubTtl = mcServiceFeeSubTtl + MyFeeItem.Amount
        ElseIf MyFeeItem.fsftIsExpense Then
            mcExpenseFeeSubTtl = mcExpenseFeeSubTtl + MyFeeItem.Amount
        End If
        f_fsftDescription.Text = MyFeeItem.fsftDescription
        f_Comment.Text = MyFeeItem.Comment
'        If CBool(MyFeeItem.fsftIsExpense) Then
'            f_fsftIsExpense.Text = "Is Expense"
'        Else
'            f_fsftIsExpense.Text = vbNullString
'        End If
        f_NumberOfItems.Text = MyFeeItem.NumberOfItems
        If f_NumberOfItems.Text = "0" Then
            f_NumberOfItems.Text = vbNullString
        End If
        f_fsftFeeAmount.Text = Format(MyFeeItem.fsftFeeAmount, "$#,###,###,##0.00")
        If f_fsftFeeAmount.Text = "$0.00" Then
            f_fsftFeeAmount.Text = vbNullString
        End If
        If f_NumberOfItems.Text = vbNullString And f_fsftFeeAmount.Text = vbNullString Then
            lblX.Caption = vbNullString
        Else
            lblX.Caption = "X"
        End If
        f_Amount.Text = Format(MyFeeItem.Amount, "$#,###,###,##0.00")
        mlCount = mlCount + 1
        'Check the next Fee item...
        'If its Expense Fee then need to Show the Service Fee SubTotal
        'If mlcount is over the Count then show the Expense Fee Subtotal
        If mlCount > mcoludtFeeItems01.Count And Not mbShowedExpenseFeeSubTtl Then
            mbShowedExpenseFeeSubTtl = True
            Detail.Height = DETAIL_H_SUBTTL
            Set subSubTtl.object = msubSubTtlExpenseFee.object
            subSubTtl.Visible = True
        Else
            MyFeeItem = mcoludtFeeItems01(mlCount)
            If MyFeeItem.fsftIsExpense And Not mbShowedServiceFeeSubTtl Then
                mbShowedServiceFeeSubTtl = True
                Detail.Height = DETAIL_H_SUBTTL
                Set subSubTtl.object = msubSubTtlServiceFee.object
                subSubTtl.Visible = True
            Else
                Detail.Height = DETAIL_H
                Set subSubTtl.object = Nothing
                subSubTtl.Visible = False
            End If
        End If
        Detail.PrintSection
    End If
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & Me.Name & vbCrLf & "Private Sub Detail_Format"
End Sub

Private Sub ReportFooter_Format()
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String

    
    f_TotalServiceAndExpense.Text = Format(mcTotalServiceAndExpense, "$#,###,###,##0.00")
    f_TaxPercent.Text = Format(mdblTaxPercent, "0.000")
    f_TaxAmount.Text = Format(mcTaxAmount, "$#,###,###,##0.00")
    f_TotalAdjustingFee.Text = Format(mcTotalAdjustingFee, "$#,###,###,##0.00")
    
Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & Me.Name & vbCrLf & "Private Sub Detail_Format"
End Sub

