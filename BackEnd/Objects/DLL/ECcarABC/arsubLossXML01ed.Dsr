VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} arsubLossXML01ed 
   Caption         =   "ActiveReport3"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19606
   SectionData     =   "arsubLossXML01ed.dsx":0000
End
Attribute VB_Name = "arsubLossXML01ed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mcolED As Collection
Private mlCount As Long
Private mLossType As TypeXML01

Public Property Let LossType(pType As TypeXML01)
    mLossType = pType
End Property
Public Property Get LossType() As TypeXML01
    LossType = mLossType
End Property

Public Property Let EDcol(pcolED As Object)
    Set mcolED = pcolED
End Property
Public Property Set EDcol(pcolED As Object)
    Set mcolED = pcolED
End Property

Private Sub ActiveReport_ReportEnd()
    Set mcolED = Nothing
End Sub

Private Sub ActiveReport_ReportStart()
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    mlCount = 1
    'Account for XML01 type
    If LossType = XML01Pro Then
        fedDescription.left = 1800
        fedDescription.Width = 9450
    ElseIf LossType = XML01Apd Then
        fedDescription.left = 0
        fedDescription.Width = 11250
    End If
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & Me.Name & vbCrLf & "Private Sub ActiveReport_ReportStart"
End Sub

Private Sub Detail_Format()
    On Error GoTo EH
    Dim ED As udtXML01Endorsement
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    If mcolED Is Nothing Then
        Exit Sub
    End If
    If mlCount <= mcolED.Count Then
        ED = mcolED(mlCount)
        fedCode.Text = ED.EDCode
        fedDescription.Text = ED.EDDescription
        
        mlCount = mlCount + 1
        Detail.PrintSection
    End If
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & Me.Name & vbCrLf & "Private Sub Detail_Format"
End Sub



