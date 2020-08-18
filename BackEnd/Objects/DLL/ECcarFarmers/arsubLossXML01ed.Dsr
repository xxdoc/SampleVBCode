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
Private mEndorsementRS As WDDXRecordset
Private mlCount As Long
Private mLossType As TypeXML01

Public Property Let EndorsementRS(pEndorsementRS As WDDXRecordset)
    Set mEndorsementRS = pEndorsementRS
End Property
Public Property Set EndorsementRS(pEndorsementRS As WDDXRecordset)
    Set mEndorsementRS = pEndorsementRS
End Property

Private Sub ActiveReport_ReportEnd()
    Set mEndorsementRS = Nothing
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
    Dim ED As udtXML01Endorsements
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    If mEndorsementRS Is Nothing Then
        Exit Sub
    End If
    If mlCount <= mEndorsementRS.getRowCount Then
        With ED
            .EndorsementNumber = IIf(IsNull(mEndorsementRS.getField(mlCount, "EndorsementNumber")), vbNullString, mEndorsementRS.getField(mlCount, "EndorsementNumber"))
            .EndorsementDescription = IIf(IsNull(mEndorsementRS.getField(mlCount, "EndorsementDescription")), vbNullString, mEndorsementRS.getField(mlCount, "EndorsementDescription"))
        End With
        f_EndorsementNumber.Text = ED.EndorsementNumber
        f_EndorsementDescription.Text = ED.EndorsementDescription
        
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



