VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} arsubLossXML01cov 
   Caption         =   "ActiveReport1"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19606
   SectionData     =   "arsubLossXML01cov.dsx":0000
End
Attribute VB_Name = "arsubLossXML01cov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mCoverageRS As WDDXRecordset
Private mlCount As Long
Private msDedType As String

Public Property Let DedType(psDedType As String)
    msDedType = psDedType
End Property
Public Property Get DedType() As String
    DedType = msDedType
End Property

Public Property Let CoverageRS(pCoverageRS As WDDXRecordset)
    Set mCoverageRS = pCoverageRS
End Property
Public Property Set CoverageRS(pCoverageRS As WDDXRecordset)
    Set mCoverageRS = pCoverageRS
End Property

Private Sub ActiveReport_ReportEnd()
    Set mCoverageRS = Nothing
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
    Dim COV As udtXML01Coverages
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    If mCoverageRS Is Nothing Then
        Exit Sub
    End If
NEXT_LOG:
    If mlCount <= mCoverageRS.getRowCount Then
        With COV
            .Coverage = IIf(IsNull(mCoverageRS.getField(mlCount, "Coverage")), vbNullString, mCoverageRS.getField(mlCount, "Coverage"))
            .Limits = IIf(IsNull(mCoverageRS.getField(mlCount, "Limits")), vbNullString, mCoverageRS.getField(mlCount, "Limits"))
            .Deductible1 = IIf(IsNull(mCoverageRS.getField(mlCount, "Deductible1")), vbNullString, mCoverageRS.getField(mlCount, "Deductible1"))
            If StrComp(msDedType, "Property", vbTextCompare) = 0 Then
                .Deductible2 = IIf(IsNull(mCoverageRS.getField(mlCount, "Deductible2")), vbNullString, mCoverageRS.getField(mlCount, "Deductible2"))
                .Deductible3 = IIf(IsNull(mCoverageRS.getField(mlCount, "Deductible3")), vbNullString, mCoverageRS.getField(mlCount, "Deductible3"))
                .Deductible4 = IIf(IsNull(mCoverageRS.getField(mlCount, "Deductible4")), vbNullString, mCoverageRS.getField(mlCount, "Deductible4"))
            End If
        End With
        f_Coverage.Text = COV.Coverage
        f_Limits.Text = COV.Limits
        f_Deductible1.Text = COV.Deductible1
        f_Deductible2.Text = COV.Deductible2
        f_Deductible3.Text = COV.Deductible3
        f_Deductible4.Text = COV.Deductible4
        
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




