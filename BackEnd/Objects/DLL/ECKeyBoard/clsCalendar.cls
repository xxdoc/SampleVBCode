VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mfrmCalendar As frmCalendar
Private mlCalTOP As Long
Private mlCalLEFT As Long
Private mdtCurrentDate As Date

Public Property Let CurrentDate(pdtDate As Date)
    mdtCurrentDate = pdtDate
End Property
Public Property Get CurrentDate() As Date
    CurrentDate = mdtCurrentDate
End Property

Public Property Let CalTOP(plValue As Long)
    mlCalTOP = plValue
End Property
Public Property Get CalTOP() As Long
    CalTOP = mlCalTOP
End Property

Public Property Let CalLEFT(plValue As Long)
    mlCalLEFT = plValue
End Property
Public Property Get CalLEFT() As Long
    CalLEFT = mlCalLEFT
End Property

Private Property Get msClassName() As String
    msClassName = App.EXEName & ".clsCalendar"
End Property

Public Sub ShowCalendar()
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    Set mfrmCalendar = New frmCalendar
    mfrmCalendar.CurrentDate = mdtCurrentDate
    
    Load mfrmCalendar
    
    mfrmCalendar.top = mlCalTOP
    mfrmCalendar.left = mlCalLEFT
    
    'Check for Top and Left...
    'If 0 then set the top and left realative to Mouse X y
    If mfrmCalendar.top = 0 And mfrmCalendar.left = 0 Then
         mfrmCalendar.SetTopAndLeft
    End If
    
    mfrmCalendar.Show vbModal
    'Set the CurrentDate after calendar form hides
    mdtCurrentDate = mfrmCalendar.CurrentDate
    
    Unload mfrmCalendar
    Set mfrmCalendar = Nothing
    
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & msClassName & vbCrLf & "Public Sub ShowCalendar"
End Sub



