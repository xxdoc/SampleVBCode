VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} arsubLossXML01plhCont 
   Caption         =   "ActiveReport5"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19606
   SectionData     =   "arsubLossXML01plhCont.dsx":0000
End
Attribute VB_Name = "arsubLossXML01plhCont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mcolPLHCont As Collection
Private mlCount As Long

Public Property Let PLHContcol(pcolPLHCont As Object)
    Set mcolPLHCont = pcolPLHCont
End Property
Public Property Set PLHContcol(pcolPLHCont As Object)
    Set mcolPLHCont = pcolPLHCont
End Property

Private Sub ActiveReport_ReportEnd()
    Set mcolPLHCont = Nothing
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
    Dim PLHCont As udtXML01PriorLossHistCont
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    If mcolPLHCont Is Nothing Then
        Exit Sub
    End If
    
    If mlCount <= mcolPLHCont.Count Then
        PLHCont = mcolPLHCont(mlCount)
        Me.fplh01_CAT.Text = PLHCont.plh01_CAT
        Me.fplh02_LossDate.Text = PLHCont.plh02_LossDate
        Me.fplh03_Adjuster.Text = PLHCont.plh03_Adjuster
        Me.fplh04_DateAsgn.Text = PLHCont.plh04_DateAsgn
        Me.fplh05_DateClsd.Text = PLHCont.plh05_DateClsd
        Me.fplh06_AmtPaid.Text = PLHCont.plh06_AmtPaid
        Me.fplh07_SALN.Text = PLHCont.plh07_SALN
        
        mlCount = mlCount + 1
        Detail.PrintSection
    End If
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & Me.Name & vbCrLf & "Private Sub Detail_Format"
End Sub


