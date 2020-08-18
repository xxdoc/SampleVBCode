VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} arLossXML01Cont 
   Caption         =   "ActiveReport1"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19606
   SectionData     =   "arLossXML01Cont.dsx":0000
End
Attribute VB_Name = "arLossXML01Cont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private msubPLHCont As arsubLossXML01plhCont 'Prior Loss Hist sub report
Private msubCALCont As arsubLossXML01cal 'Comments Activity Log sub report
Private mXML01LossCont As udtXML01LossCont  'XML01 Loss Repport user defined type
Private msDateTimePrinted As String

Public Property Let DateTimePrinted(psDateTime As String)
    msDateTimePrinted = psDateTime
End Property

Public Property Let XML01LossCont(pXML01LossCont As udtXML01LossCont)
    mXML01LossCont = pXML01LossCont
End Property

Public Property Get ClassName() As String
    ClassName = App.EXEName & "." & Me.Name
End Property

Private Sub ActiveReport_ReportEnd()
    Unload subLossXML01plhCont.object
    Set subLossXML01plhCont.object = Nothing
    Unload subLossXML01calCont.object
    Set subLossXML01calCont.object = Nothing
    Set msubPLHCont = Nothing
    Set msubCALCont = Nothing
End Sub

Private Sub ActiveReport_ReportStart()
    On Error GoTo EH
    Dim lErrNum  As Long
    Dim sErrDesc As String
    
    'Instance sub reports
    Set msubPLHCont = New arsubLossXML01plhCont
    Set msubCALCont = New arsubLossXML01cal
    'Set their data collections
    Set msubPLHCont.PLHContcol = mXML01LossCont.colPLHCont
    Set msubCALCont.CALcol = mXML01LossCont.colCAL
    'Set the ref to sub reports in this Report
    Set subLossXML01plhCont.object = msubPLHCont
    Set subLossXML01calCont.object = msubCALCont
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Private Sub ActiveReport_ReportStart"
End Sub

Private Sub Detail_Format()
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    fali0004_DateTimePrinted.Text = msDateTimePrinted
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Private Sub Detail_Format"
End Sub



