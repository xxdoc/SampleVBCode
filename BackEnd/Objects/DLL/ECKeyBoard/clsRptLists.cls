VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRptLists"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get ClassName() As String
'    clsRptLists_ClassName = App.EXEName & ".clsLists"
End Property

Public Property Get ARReportObject() As Object
'    clsRptLists_ARReportObject = moArReportObejct
End Property

Public Function GetARMiscDelimParamsCol(plVersion As Long, _
                                        Optional psDelimParams As String = vbNullString) As Collection
'    On Error GoTo EH
'    Dim lErrNum As Long
'    Dim sErrDesc As String

    
'
'    Exit Function
'EH:
'    lErrNum = Err.Number
'    sErrDesc = Err.Description
'    Err.Raise lerrnum, , serrdesc & vbCrLf & clsRptLists_ClassName & vbCrLf & "Public Function GetARMiscDelimParamsCol"
End Function

Public Function GetARReport(plVersion As Long, _
                            Optional psDelimParams As String = vbNullString) As Object
'    On Error GoTo EH
'    Dim lErrNum As Long
'    Dim sErrDesc As String

    
'
'    Select Case UCase(plVersion)
'        Case 1
'            Set clsRptLists_GetARReport = New ECRpt???.arRpt???ReportName
'    End Select
'
'    Exit Function
'EH:
'    lErrNum = Err.Number
'    sErrDesc = Err.Description
'    Err.Raise lerrnum, , serrdesc & vbCrLf & clsRptLists_ClassName & vbCrLf & "Public Function GetARReport"
End Function

Public Function SetUtilObject(oUtilObject As Object) As Boolean
'    On Error Resume Next
'    Set goUtil = oUtilObject
'    If Err.Number = 0 Then
'        SetUtilObject = True
'    Else
'        SetUtilObject = False
'    End If
End Function

Public Function CLEANUP() As Boolean
'    On Error GoTo EH
'    Dim lErrNum As Long
'    Dim sErrDesc As String

'    Set goUtil = Nothing
'
'    CLEANUP = True
'    Exit Function
'EH:
'    lErrNum = Err.Number
'    sErrDesc = Err.Description
'    Err.Raise lerrnum, , serrdesc & vbCrLf & clsCarLists_ClassName & vbCrLf & "Public Function CLEANUP"
End Function

