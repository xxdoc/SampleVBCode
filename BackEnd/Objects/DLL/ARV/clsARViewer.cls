VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsARViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Type tErr
    Description As String
    HelpContext As String
    HelpFile As String
    Number As Long
    Source As String
End Type

Private mobjARVReport As Object
Private msRptTitle As String
Private msOpenFile As String
Private mcolActhWnd As Collection
Private mbHidePrintButton As Boolean

Public Property Let HidePrintButton(pbFlag As Boolean)
    mbHidePrintButton = pbFlag
End Property
Public Property Get HidePrintButton() As Boolean
    HidePrintButton = mbHidePrintButton
End Property

Public Property Let objARvReport(pobjARVReport As Object)
    Set mobjARVReport = pobjARVReport
End Property
Public Property Set objARvReport(pobjARVReport As Object)
    Set mobjARVReport = pobjARVReport
End Property
Public Property Get objARvReport() As Object
    Set objARvReport = mobjARVReport
End Property

Public Property Let sRptTitle(psRptTitle As String)
    msRptTitle = psRptTitle
End Property
Public Property Get sRptTitle() As String
   sRptTitle = msRptTitle
End Property
Public Property Get sOpenFile() As String
   sOpenFile = msOpenFile
End Property
Public Property Let sOpenFile(psFilePath As String)
   msOpenFile = psFilePath
End Property
Public Property Let ADDhWnd(plHwnd As Long)
    If mcolActhWnd Is Nothing Then
        Set mcolActhWnd = New Collection
    End If
    mcolActhWnd.Add CStr(plHwnd), CStr(plHwnd)
End Property
Public Property Let POPhWnd(plHwnd As Long)
    mcolActhWnd.Remove CStr(plHwnd)
End Property
Public Property Get hWndCount() As Integer
    hWndCount = mcolActhWnd.Count
End Property

Public Property Get goForms() As Object
    Set goForms = Forms
End Property

Public Sub ShowReport(Optional piMode As VBRUN.FormShowConstants = vbModeless)
    Dim objPreview As frmPreview
    Dim MyErr As tErr
    
    Set objPreview = New frmPreview
    Set objPreview.ARViewer = Me
    Load objPreview
    With objPreview
        If mobjARVReport Is Nothing And msOpenFile <> vbNullString Then
            If .OpenFile(MyErr) Then
                .Show piMode
                .WindowState = vbNormal
                DoEvents
                Sleep 100
                .SetFocus
            Else
                With MyErr
                    Err.Raise .Number, .Source, .Description, .HelpFile, .HelpContext
                End With
            End If
        Else
            .RunReport mobjARVReport
            .Show piMode
            If piMode = vbModeless Then
                .WindowState = vbNormal
                If .Visible Then
                    .SetFocus
                End If
            End If
        End If
    End With
End Sub

Public Sub ShowReportOnForm(poForm As Object, Optional piMode As VBRUN.FormShowConstants = vbModeless)
    Dim MyErr As tErr
    
    Set poForm = New frmPreview
    Set poForm.ARViewer = Me
    Load poForm
    With poForm
        If mobjARVReport Is Nothing And msOpenFile <> vbNullString Then
            If .OpenFile(MyErr) Then
                .Show piMode
                .WindowState = vbNormal
                DoEvents
                Sleep 100
                .SetFocus
            Else
                With MyErr
                    Err.Raise .Number, .Source, .Description, .HelpFile, .HelpContext
                End With
            End If
        Else
            .RunReport mobjARVReport
            .Show piMode
            If piMode = vbModeless Then
                .WindowState = vbNormal
                If .Visible Then
                    .SetFocus
                End If
            End If
        End If
    End With
End Sub

Public Function CleanUp() As Boolean
    On Error GoTo EH
    Dim MyForm As frmPreview
    
    For Each MyForm In Forms
        Set MyForm.ARv.ReportSource = Nothing
        Unload MyForm
        Set MyForm = Nothing
    Next
    If Not mobjARVReport Is Nothing Then
        Set mobjARVReport = Nothing
    End If
    If Not mcolActhWnd Is Nothing Then
        Set mcolActhWnd = Nothing
    End If
    If Not goUtil Is Nothing Then
        Set goUtil = Nothing
    End If
    
    CleanUp = True
    Exit Function
EH:
    CleanUp = False
    
End Function

Public Function SetGlobalObjects(pcolObjects As Collection) As Boolean
    On Error Resume Next
    Set goUtil = pcolObjects("goUtil")
    If Err.Number = 0 Then
        SetGlobalObjects = True
    Else
        SetGlobalObjects = False
    End If
End Function

Public Function SetUtilObject(poUtil As V2ECKeyBoard.clsUtil) As Boolean
    On Error Resume Next
    Set goUtil = poUtil
    If Err.Number = 0 Then
        SetUtilObject = True
    Else
        SetUtilObject = False
    End If
End Function


