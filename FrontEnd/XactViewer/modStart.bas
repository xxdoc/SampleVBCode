Attribute VB_Name = "modStart"
Option Explicit
Private mFrmCancel As frmCancel
Private moRegSetting As V2ECKeyBoard.clsRegSetting
Public goUtil As V2ECKeyBoard.clsUtil
Private moARV As V2ARViewer.clsARViewer

Private Const SHCNE_ASSOCCHANGED = &H8000000
Private Const SHCNF_IDLIST = &H0&
Private Declare Sub SHChangeNotify Lib "shell32.dll" (ByVal wEventId As Long, ByVal uFlags As Long, dwItem1 As Any, dwItem2 As Any)


Public Sub Main()
    On Error GoTo EH
    Dim sFile As String
    Dim lRet As Long
    Dim vXProj As Variant
    Dim colXProjects As Collection
    
    sFile = Replace(Command$, """", vbNullString)
    '2.3.2004 BGS Set registry Only when user clicks
    'on the XactViewer.exe
    If sFile = vbNullString Then
        SetFileAssociation
        GoTo CLEAN_UP
    End If
    
    Set goUtil = New V2ECKeyBoard.clsUtil
    goUtil.gsAppEXEName = App.EXEName
    goUtil.gsMainAppEXEName = App.EXEName
    
    
    Set moARV = New V2ARViewer.clsARViewer
    moARV.SetUtilObject goUtil
    Set goUtil.gARV = moARV
    
    Set goUtil.goXact = New V2ECKeyBoard.clsXact
    goUtil.goXact.SetUtilObject goUtil
    
    Set goUtil.goECKeyBoardList = New V2ECKeyBoard.clsLists
    goUtil.goECKeyBoardList.SetUtilObject goUtil
    
    goUtil.goXact.GetFromExport = True
    goUtil.goXact.ExportFilePath = sFile
    goUtil.goXact.PopulateFromXactExport
    'Show the Viewer
   If goUtil.goXact.ValidateXactProjects() Then
        'Display the Cancel Form
        Set mFrmCancel = New frmCancel
        Load mFrmCancel
        mFrmCancel.Visible = True
        mFrmCancel.lblMess.Caption = "Please wait! Sending files..."
        
        If Not goUtil Is Nothing Then
            If goUtil.goXact Is Nothing Then
               GoTo CLEAN_UP
            End If
        Else
            GoTo CLEAN_UP
        End If
        
        'If we are not skiping all then we need to SendToXactimate
        'SendToXact will launch Xactimate if it isn't already loaded and set focus to it
        If Not goUtil.goXact.SkipAll Then
            If goUtil.goXact.SendToXact() Then
                'OK we need a check here incase user tries to unload in middle of sending
                If Not goUtil Is Nothing Then
                    If goUtil.goXact Is Nothing Then
                       GoTo CLEAN_UP
                    End If
                Else
                    GoTo CLEAN_UP
                End If
                goUtil.goXact.LookForWindow Left(mFrmCancel.Caption, 30)
                If goUtil.goXact.SendToExport Then
                    MsgBox "Project(s) Exported to: " & vbCrLf & vbCrLf & goUtil.goXact.ExportFilePath, vbInformation + vbOKOnly, "Send To Xactimate Export File"
                Else
                    MsgBox "Project(s) Sent to Xactimate From Export File!", vbInformation + vbOKOnly, "Send To Xactimate From Export File"
                End If
                
            Else
                'OK we need a check here incase user tries to unload in middle of sending
                If Not goUtil Is Nothing Then
                    If goUtil.goXact Is Nothing Then
                        GoTo CLEAN_UP
                    End If
                Else
                   GoTo CLEAN_UP
                End If
                goUtil.goXact.LookForWindow Left(mFrmCancel.Caption, 30)
                If goUtil.goXact.SendToExport Then
                    MsgBox "Project(s) Failed to Export to: " & vbCrLf & vbCrLf & goUtil.goXact.ExportFilePath, vbExclamation + vbOKOnly, "Send To Xactimate Export File"
                Else
                    MsgBox "Project(s) Not Sent to Xactimate From Export File!", vbExclamation + vbOKOnly, "Send To Xactimate From Export File"
                End If
                
            End If
        Else
            goUtil.goXact.ExportData = vbNullString
            'Need to rebuild the Export file
            Set colXProjects = goUtil.goXact.XactProjects
            For Each vXProj In colXProjects
                goUtil.goXact.BuildXactExport vXProj
            Next
            
            If goUtil.utFileExists(goUtil.goXact.ExportFilePath) Then
                Kill goUtil.goXact.ExportFilePath
            End If
            goUtil.utSaveFileData goUtil.goXact.ExportFilePath, goUtil.goXact.ExportData
            goUtil.goXact.ExportData = vbNullString
        End If
    End If
    
CLEAN_UP:
    If Not mFrmCancel Is Nothing Then
        Unload mFrmCancel
        Set mFrmCancel = Nothing
    End If
    If Not moARV Is Nothing Then
         moARV.CleanUp
        Set moARV = Nothing
    End If
    If Not goUtil Is Nothing Then
        goUtil.CleanUp
        Set goUtil = Nothing
    End If
    Set moRegSetting = Nothing
    Set colXProjects = Nothing
    Exit Sub
EH:
   MsgBox "Could not open..." & vbCrLf & sFile & vbCrLf & Err.Description, vbInformation, "Xactimate Export File Viewer"
    If Not mFrmCancel Is Nothing Then
        Unload mFrmCancel
        Set mFrmCancel = Nothing
    End If
End Sub

Private Sub SetFileAssociation()
    '2.3.2004 BGS Need to be sure that the App path and Exe name
    'is Default EXE that opens .xef File Formats
    Set moRegSetting = New V2ECKeyBoard.clsRegSetting
    With moRegSetting
        '2.3.2004 BGS Set up registry to Auto Open .xef Extentions with the XactViewer.exe
        .SaveSetting HKEY_CLASSES_ROOT, ".xef", vbNullString, "xef_auto_file"
        .SaveSetting HKEY_CLASSES_ROOT, "xef_auto_file", vbNullString, "Xactimate Export File"
        .SaveSetting HKEY_CLASSES_ROOT, "xef_auto_file\Shell\Open", vbNullString, vbNullString
        '2.3.2004 BGS the & """%1""" allows for Command lines to be sent to it
        .SaveSetting HKEY_CLASSES_ROOT, "xef_auto_file\Shell\Open\Command", vbNullString, App.Path & "\" & App.EXEName & ".exe " & """%1"""
        
        .SaveSetting HKEY_LOCAL_MACHINE, "SOFTWARE\Classes\.xef", vbNullString, "xef_auto_file"
        .SaveSetting HKEY_LOCAL_MACHINE, "SOFTWARE\Classes\xef_auto_file", vbNullString, "Xactimate Export File"
        .SaveSetting HKEY_LOCAL_MACHINE, "SOFTWARE\Classes\xef_auto_file\Shell\Open", vbNullString, vbNullString
        .SaveSetting HKEY_LOCAL_MACHINE, "SOFTWARE\Classes\xef_auto_file\Shell\Open\Command", vbNullString, App.Path & "\" & App.EXEName & ".exe " & """%1"""
        
        '2.3.2004 BGS Set the Icon to be the EXE Icon
        .SaveSetting HKEY_CLASSES_ROOT, "xef_auto_file\DefaultIcon", vbNullString, App.Path & "\" & App.EXEName & ".exe,0"
        .SaveSetting HKEY_LOCAL_MACHINE, "SOFTWARE\Classes\xef_auto_file\DefaultIcon", vbNullString, App.Path & "\" & App.EXEName & ".exe,0"
        'BGS Refresh the Icons
        SHChangeNotify SHCNE_ASSOCCHANGED, SHCNF_IDLIST, 0, 0
    End With
    'clean up
    Set moRegSetting = Nothing
End Sub
