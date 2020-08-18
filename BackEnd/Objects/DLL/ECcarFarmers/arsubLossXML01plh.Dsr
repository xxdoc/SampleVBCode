VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} arsubLossXML01plh 
   Caption         =   "ActiveReport4"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19606
   SectionData     =   "arsubLossXML01plh.dsx":0000
End
Attribute VB_Name = "arsubLossXML01plh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mPriorLossDetailRS As WDDXRecordset
Private mlCount As Long

Public Property Let PriorLossDetailRS(pPriorLossDetailRS As WDDXRecordset)
    Set mPriorLossDetailRS = pPriorLossDetailRS
End Property
Public Property Set PriorLossDetailRS(pPriorLossDetailRS As WDDXRecordset)
    Set mPriorLossDetailRS = pPriorLossDetailRS
End Property

Private Sub ActiveReport_ReportEnd()
    Set mPriorLossDetailRS = Nothing
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
    Dim PLH As udtXML01PriorLossDetail
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    If mPriorLossDetailRS Is Nothing Then
       Exit Sub
    End If
    
    If mlCount <= mPriorLossDetailRS.getRowCount Then
        With PLH
            .SALN = IIf(IsNull(mPriorLossDetailRS.getField(mlCount, "SALN")), vbNullString, mPriorLossDetailRS.getField(mlCount, "SALN"))
            .ClaimSegmentNumber = IIf(IsNull(mPriorLossDetailRS.getField(mlCount, "ClaimSegmentNumber")), vbNullString, mPriorLossDetailRS.getField(mlCount, "ClaimSegmentNumber"))
            .PolicyNumber = IIf(IsNull(mPriorLossDetailRS.getField(mlCount, "PolicyNumber")), vbNullString, mPriorLossDetailRS.getField(mlCount, "PolicyNumber"))
            .LossCause = IIf(IsNull(mPriorLossDetailRS.getField(mlCount, "LossCause")), vbNullString, mPriorLossDetailRS.getField(mlCount, "LossCause"))
            .ClaimClass = IIf(IsNull(mPriorLossDetailRS.getField(mlCount, "ClaimClass")), vbNullString, mPriorLossDetailRS.getField(mlCount, "ClaimClass"))
            .LossDate = IIf(IsNull(mPriorLossDetailRS.getField(mlCount, "LossDate")), vbNullString, mPriorLossDetailRS.getField(mlCount, "LossDate"))
            .SummaryAmount = IIf(IsNull(mPriorLossDetailRS.getField(mlCount, "SummaryAmount")), vbNullString, mPriorLossDetailRS.getField(mlCount, "SummaryAmount"))
        End With
        
        f_SALN.Text = PLH.SALN
        f_ClaimSegmentNumber.Text = PLH.ClaimSegmentNumber
        f_PolicyNumber.Text = PLH.PolicyNumber
        f_LossCause.Text = PLH.LossCause
        f_ClaimClass.Text = PLH.ClaimClass
        f_LossDate.Text = Format(PLH.LossDate, "MM/DD/YYYY")
        f_SummaryAmount.Text = Format(PLH.SummaryAmount, "0.00")
        
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

