VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Carrier GUI Polymorphism
'Carrier GUI InterFace Definiton
'Use Implements V2ECKeyBoard.clsCarGUI in ALL clsGUI For All Carrier GUI Objects

Implements V2ECKeyBoard.clsCarGUI

Private mfrmClaimsList As frmClaimsList

Public Property Get clsCarGUI_ClassName() As String
    clsCarGUI_ClassName = App.EXEName & ".clsGUI"
End Property

Public Property Get clsCarGUI_goForms() As Object
    Set clsCarGUI_goForms = Forms
End Property

Private Function clsCarGUI_SetUtilObject(oUtilObject As Object) As Boolean

    On Error Resume Next
    Set goUtil = oUtilObject
    If Err.Number = 0 Then
        clsCarGUI_SetUtilObject = True
    Else
        clsCarGUI_SetUtilObject = False
    End If

End Function

Private Sub clsCarGUI_ShowForm(psFormName As String, Optional pMode As VBRUN.FormShowConstants = vbModal, Optional oMnuItem As Object)
    Dim oForm As Form
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    On Error GoTo EH
    
    If goUtil.utFindSetForm(Forms, psFormName, oForm) Then
        oForm.Show
        If oForm.WindowState <> vbMaximized Then
            oForm.WindowState = vbNormal
        End If
    Else
        Select Case psFormName
            Case "frmClaimsList"
                Set mfrmClaimsList = New frmClaimsList
                Load mfrmClaimsList
                mfrmClaimsList.Show pMode
                
        End Select
    End If
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & clsCarGUI_ClassName & vbCrLf & "Private Sub clsCarGUI_ShowForm" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Sub

Public Sub clsCarGUI_ShowAllForms(oECTray As Object, psNavPos As String, Optional psSkipForms As String)
    goUtil.utShowAllForms Forms, oECTray, psNavPos, psSkipForms
End Sub

Public Sub clsCarGUI_HideAllForms(Optional psSkipForms As String)
    goUtil.utHideAllForms Forms, psSkipForms
End Sub

Public Function clsCarGUI_SetGlobalObjects(pcolObjects As Collection) As Boolean
    On Error Resume Next
    
    Set goUtil = pcolObjects("goUtil")
    If Err.Number = 0 Then
        clsCarGUI_SetGlobalObjects = True
    Else
        clsCarGUI_SetGlobalObjects = False
    End If
End Function


Public Function clsCarGUI_CLEANUP() As Boolean
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    If Not mfrmClaimsList Is Nothing Then
        Unload mfrmClaimsList
        Set mfrmClaimsList = Nothing
    End If
    
    clsCarGUI_CLEANUP = True
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    clsCarGUI_CLEANUP = False
    Err.Raise lErrNum, , sErrDesc & vbCrLf & clsCarGUI_ClassName & vbCrLf & "Public Function CLEANUP"
End Function




