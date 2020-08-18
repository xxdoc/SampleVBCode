Attribute VB_Name = "modMain"
Option Explicit

'Pass this one Global Object between Apps
Public goUtil As V2ECKeyBoard.clsUtil
Public goECBill As V2EberlsBillings.clsSend2Billings

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long

Private Property Get msClassName() As String
    msClassName = "modMain"
End Property


Public Sub Main()
    On Error GoTo EH
    Dim sMess As String
    Dim colOBjects As Collection
    
    'If we already have this running then Bail
    If App.PrevInstance Then
        sMess = App.EXEName & " is already running. " & vbCrLf & vbCrLf
        sMess = sMess & "If you can't see " & App.EXEName & vbCrLf
        sMess = sMess & "that means Windows had trouble ending it's process.  If this is true, "
        sMess = sMess & "please manually end task on the previous " & App.EXEName & " session using the task manager." & vbCrLf & vbCrLf
        sMess = sMess & "Thank You!"
        MsgBox sMess, vbInformation
        End
        Exit Sub
    End If
    
    Set colOBjects = New Collection
    
    'Set Public Objects Here
    Set goUtil = New V2ECKeyBoard.clsUtil
    Set goUtil.goECKeyBoardList = New V2ECKeyBoard.clsLists
    Set goUtil.gARV = New V2ARViewer.clsARViewer
    
    goUtil.gsAppEXEName = App.EXEName 'Application Name
    
    goUtil.gsMainAppEXEName = App.EXEName
    
    goUtil.SetUtilObject goUtil
    
    Load frmMain
    frmMain.Show
    
    ShowSplash
    
    colOBjects.Add goUtil, "goUtil"
    colOBjects.Add frmMain, "frmMain"
    
    Set goECBill = New clsSend2Billings
    goECBill.SetGlobalObjects colOBjects
    goUtil.gARV.SetGlobalObjects colOBjects
    
    Set goECBill.progBar = frmMain.progBar
    goECBill.PopulateLookUp
    frmMain.PopulateLookupCbo
    Set colOBjects = Nothing
    
    HideSplash
Exit Sub
EH:
    goUtil.utErrorLog Err, goUtil.gsAppEXEName, msClassName, "Public Sub Main"
    frmMain.EnableLoadClaims
End Sub

Public Sub CleanUpAndExit()
    On Error GoTo EH
    
    goECBill.CleanUp
    Set goECBill = Nothing
    
    Unload frmMain
    Set frmMain = Nothing
    
    goUtil.CleanUp
    Set goUtil = Nothing

    Exit Sub
EH:
    MsgBox "Error #" & Err.Number & vbCrLf & Err.Description & vbCrLf & msClassName & vbCrLf & "Public Sub CleanUp"
End Sub

Public Sub ShowSplash()
    On Error GoTo EH
    frmMain.framSplash.Visible = True
    frmMain.framSplash.ZOrder
    Exit Sub
EH:
    goUtil.utErrorLog Err, goUtil.gsAppEXEName, msClassName, "Public Sub ShowSplash"
End Sub

Public Sub HideSplash()
    On Error GoTo EH
    
    frmMain.framSplash.Visible = False
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, goUtil.gsAppEXEName, msClassName, "Public Sub HideSplash"
End Sub

