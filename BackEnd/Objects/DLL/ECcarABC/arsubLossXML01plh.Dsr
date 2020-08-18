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
Private mcolPLH As Collection
Private mlCount As Long

Public Property Let PLHcol(pcolPLH As Object)
    Set mcolPLH = pcolPLH
End Property
Public Property Set PLHcol(pcolPLH As Object)
    Set mcolPLH = pcolPLH
End Property

Private Sub ActiveReport_ReportEnd()
    Set mcolPLH = Nothing
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
    Dim PLH As udtXML01PriorLossHist
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    If mcolPLH Is Nothing Then
       Exit Sub
    End If
    
    If mlCount <= mcolPLH.Count Then
        PLH = mcolPLH(mlCount)
        fplh01_SALN.Text = PLH.plh01_SALN
        fplh02_LossDate.Text = PLH.plh02_LossDate
        fplh03_Adjuster.Text = PLH.plh03_Adjuster
        fplh04_DateAsgn.Text = PLH.plh04_DateAsgn
        fplh05_DateClsd.Text = PLH.plh05_DateClsd
        fplh06_AmtPaid.Text = PLH.plh06_AmtPaid
        mlCount = mlCount + 1
        Detail.PrintSection
    End If
    
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & Me.Name & vbCrLf & "Private Sub Detail_Format"
End Sub

