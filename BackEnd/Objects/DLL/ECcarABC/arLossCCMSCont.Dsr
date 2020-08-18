VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} arLossCCMSCont 
   Caption         =   "CCMS Loss Cont Report"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19606
   SectionData     =   "arLossCCMSCont.dsx":0000
End
Attribute VB_Name = "arLossCCMSCont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private msubPLHCont As arsubLossCCMSplhCont 'Prior Loss Hist sub report
Private msubCALCont As arsubLossCCMScal 'Comments Activity Log sub report
Private mCCMSLossCont As udtCCMSLossCont  'CCMS Loss Repport user defined type
Private msDateTimePrinted As String

Public Property Let DateTimePrinted(psDateTime As String)
    msDateTimePrinted = psDateTime
End Property

Public Property Let CCMSLossCont(pCCMSLossCont As udtCCMSLossCont)
    mCCMSLossCont = pCCMSLossCont
End Property

Public Property Get ClassName() As String
    ClassName = App.EXEName & "." & Me.Name
End Property

Private Sub ActiveReport_ReportEnd()
    Unload subLossCCMSplhCont.object
    Set subLossCCMSplhCont.object = Nothing
    Unload subLossCCMScalCont.object
    Set subLossCCMScalCont.object = Nothing
    Set msubPLHCont = Nothing
    Set msubCALCont = Nothing
End Sub

Private Sub ActiveReport_ReportStart()
    On Error GoTo EH
    Dim lErrNum  As Long
    Dim sErrDesc As String
    
    'Instance sub reports
    Set msubPLHCont = New arsubLossCCMSplhCont
    Set msubCALCont = New arsubLossCCMScal
    'Set their data collections
    Set msubPLHCont.PLHContcol = mCCMSLossCont.colPLHCont
    Set msubCALCont.CALcol = mCCMSLossCont.colCAL
    'Set the ref to sub reports in this Report
    Set subLossCCMSplhCont.object = msubPLHCont
    Set subLossCCMScalCont.object = msubCALCont
    
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


