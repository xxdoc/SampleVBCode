VERSION 5.00
Begin VB.Form frmTimer 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2505
   LinkTopic       =   "Form1"
   ScaleHeight     =   405
   ScaleWidth      =   2505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer TimerProcess 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "frmTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private moProcess As clsProcess
Private mlStartHwnd As Long

Private Property Get msClassName() As String
    msClassName = Me.Name
End Property

Public Property Let StartHwnd(plHwnd As Long)
    mlStartHwnd = plHwnd
End Property
Public Property Let Process(poProcess As clsProcess)
    Set moProcess = poProcess
End Property
Public Property Set Process(poProcess As clsProcess)
    Set moProcess = poProcess
End Property
Public Property Get Process() As clsProcess
    Set Process = moProcess
End Property

Private Sub TimerProcess_Timer()
    On Error GoTo EH
    Dim lhwndFound As Long
    If Not moProcess Is Nothing Then
        If moProcess.WaitForProgramToEnd() Then '(mlStartHwnd, lhwndFound) Then
            TimerProcess.Enabled = False
        End If
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub TimerProcess_Timer"
End Sub

Public Function CLEANUP() As Boolean
    On Error GoTo EH
    
    If Not moProcess Is Nothing Then
        Set moProcess = Nothing
    End If
    Unload Me
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function CleanUp"
End Function
