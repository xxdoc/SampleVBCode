Attribute VB_Name = "modMain"
Option Explicit
Public Const GW_HWNDNEXT = 2
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40

Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long

Private mFSO As Scripting.FileSystemObject

Public Sub Main()
    On Error GoTo EH
    Dim sCommand As String
    Dim saryList() As String
    Dim lCount As Long
    Dim bInstalled As Boolean
    Dim lSleep As Long
    Dim lSleepCount As Long
    Dim sDelFile As String
    Dim sMsg As String
    Dim sInstallLog As String
    Dim sMess As String
    
    
    sCommand = Command$
     
    'If we already have this running then Bail
    If App.PrevInstance Then
        sMess = App.EXEName & " is already running. " & vbCrLf & vbCrLf
        sMess = sMess & "If you can't see " & App.EXEName & ", "
        sMess = sMess & "that means Windows had trouble ending it's process.  If this is true, "
        sMess = sMess & "please manually end task on the previous " & App.EXEName & " session using the task manager." & vbCrLf & vbCrLf
        sMess = sMess & "Thank You!"
        MsgBox sMess, vbInformation
        DeleteECSPLListFile Command$
        End
        Exit Sub
    End If
   
    If sCommand <> vbNullString Then
       If Not FileExists(sCommand) Then
            GoTo BADCOMMANDLINE
        End If
        
        Load frmSPList
        frmSPList.lblMess = "Update Now!" & vbCrLf & "Please save your work first," & vbCrLf & "then click ""OK""."
        frmSPList.Show vbModal
        If frmSPList.CancelMe Then
            MsgBox "Sotware Update Aborted!", vbExclamation + vbOKOnly, "Software Update"
            Unload frmSPList
            Set frmSPList = Nothing
            'REMOVE the Install List
            DeleteECSPLListFile Command$
            End
        End If
        
        'Get the Install list...
        sDelFile = sCommand
        sCommand = GetFileData(sCommand)
        'REMOVE the Install List
        DeleteECSPLListFile Command$
        
        saryList = Split(sCommand, vbCrLf, , vbBinaryCompare)
        
        'Be sure any exisitng Log Files are Removed
        For lCount = LBound(saryList) To UBound(saryList)
            sCommand = saryList(lCount)
            If sCommand <> vbNullString Then
                'Reverse the Command string and replace the first instance
                'of .exe there may be another .exe in the install name
                'that's why a reverse string is needed then reverse it back.
                sInstallLog = StrReverse(sCommand)
                sInstallLog = Replace(sInstallLog, "exe.", "gol.", , 1, vbTextCompare)
                sInstallLog = StrReverse(sInstallLog)
                If FileExists(sInstallLog) Then
                    sMess = DeleteFile(sInstallLog)
                    If sMess <> vbNullString Then
                        Err.Raise CLng(Left(sMess, InStr(1, sMess, vbCrLf, vbBinaryCompare) - 1)), , sMess
                    End If
                End If
            End If
        Next
        For lCount = LBound(saryList) + 1 To UBound(saryList) + 1
            sCommand = saryList(lCount - 1)
            If sCommand <> vbNullString Then
ABORT_HERE:
                If frmSPList.CancelMe Then
                    MsgBox "Sotware Update Aborted!", vbExclamation + vbOKOnly, "Software Update"
                    GoTo BAIL_HERE
                End If
                frmSPList.lblMess = "Installed Package:" & vbCrLf & lCount & " of " & UBound(saryList) + 1 & vbCrLf & vbCrLf & "Please Wait..."
                'Reverse the Command string and replace the first instance
                'of .exe there may be another .exe in the install name
                'that's why a reverse string is needed then reverse it back.
                sInstallLog = StrReverse(sCommand)
                sInstallLog = Replace(sInstallLog, "exe.", "gol.", , 1, vbTextCompare)
                sInstallLog = StrReverse(sInstallLog)
                If Not FileExists(sCommand) Then
                    Err.Raise 53, , "Install File not Found" & vbCrLf & vbCrLf & sCommand
                End If
                If FileExists(App.Path & "\ShellTemp.bat") Then
                    DeleteFile App.Path & "\ShellTemp.bat"
                End If
                'Build a batch file that shells the Install exe.
                'Do this instead of directly shelling the install exe from VB
                'This will avoid VB hanging on the Shell command
                SaveFileData App.Path & "\ShellTemp.bat", """" & sCommand & """"
                ChDir App.Path
                Shell App.Path & "\ShellTemp.bat", vbHide
                bInstalled = False
TRY_INSTALL_AGAIN:
                lSleepCount = 0
                sDelFile = vbNullString
                
                For lSleep = 1 To 400
                    lSleepCount = lSleepCount + 1
                    DoEvents
                    Sleep 100
                    bInstalled = FileExists(sInstallLog)
                    If bInstalled Or lSleepCount > 300 Or frmSPList.CancelMe Then
                        Exit For
                    End If
                Next
                If frmSPList.CancelMe Then
                    GoTo ABORT_HERE
                End If
                If lSleepCount > 300 Then
                    'If the process takes too long then there is something
                    'wrong with the Install Package.  Ask to remove the package.
                    sMsg = "SP: " & sCommand & vbCrLf & vbCrLf & "Software update not responding!" & vbCrLf & vbCrLf
                    sMsg = sMsg & "Press ""Yes"" to continue to wait..." & vbCrLf & "Press ""No"" to abort the Service Pack" & vbCrLf & "(It will be downloaded again the next time you connect.)"
                    If MsgBox(sMsg, vbExclamation + vbYesNo, "Software update not responding!") = vbYes Then
                        ChDir App.Path
                        DoEvents
                        Sleep 500
                        Shell App.Path & "\ShellTemp.bat", vbHide
                        GoTo TRY_INSTALL_AGAIN:
                    End If
                    sMsg = "Installation aborted" & vbCrLf & vbCrLf & "SP: " & sCommand & vbCrLf & vbCrLf & "Software update not responding!"
                    Err.Raise 999, , sMsg
                End If
                DeleteFile App.Path & "\ShellTemp.bat"
                'Show the form again ... and wait
                frmSPList.Visible = True
                AppActivate frmSPList.Caption
                For lSleep = 1 To 15
                    DoEvents
                    Sleep 100
                Next
                frmSPList.Visible = False
            End If
        Next
        frmSPList.Visible = False
    Else
BADCOMMANDLINE:
        MsgBox "Error... Invalid Command Line !" & vbCrLf & vbCrLf & Command$, vbExclamation + vbOKOnly, "Install List Error"
        Set mFSO = Nothing
        End
    End If
    frmSPList.lblMess = "Update Successful!"
    frmSPList.Show vbModal
BAIL_HERE:
    Unload frmSPList
    Set frmSPList = Nothing
    Set mFSO = Nothing
    End
    Exit Sub
EH:
     MsgBox "Error Number " & Err.Number & vbCrLf & _
        Err.Description & vbCrLf & vbCrLf & "Public Sub Main", vbCritical
    Unload frmSPList
    Set frmSPList = Nothing
    Set mFSO = Nothing
    End
End Sub

Public Function FileExists(strFile As String, Optional pbDirOnly As Boolean) As Boolean
    '10.24.2002 Use the File System Object since it is Superior to Dir function.
    On Error GoTo EH
    If mFSO Is Nothing Then
        Set mFSO = New Scripting.FileSystemObject
    End If
    
    If strFile <> vbNullString Then
        If Not pbDirOnly Then
            FileExists = mFSO.FileExists(strFile)
        Else
            FileExists = mFSO.FolderExists(strFile)
        End If
    End If
    
    Exit Function
EH:
    FileExists = False
End Function

Public Function DeleteFile(psFilePath As String) As String
    On Error GoTo EH
    If mFSO Is Nothing Then
        Set mFSO = New Scripting.FileSystemObject
    End If
    
    mFSO.DeleteFile psFilePath, True
    
    Exit Function
EH:
    DeleteFile = Err.Number & vbCrLf & vbCrLf & Err.Description
End Function

Public Function GetFileData(psFilePath As String, Optional pbLock As Boolean = False, Optional piFFile As Integer, Optional pbSkipMess As Boolean = True) As String
'Purpose:

'Parameters :

'Returns:

'Author :

'Revision History:  SMR     Initials    Date        Description
    On Error GoTo EH
    Dim lMyFileLen As Long
    Dim iFFile As Integer
    
    iFFile = FreeFile
    piFFile = iFFile
    If pbLock Then
        Open psFilePath For Binary Access Read Lock Read As #iFFile
    Else
        Open psFilePath For Binary Access Read As #iFFile
    End If
    lMyFileLen = FileLen(psFilePath) + 2
    GetFileData = Input(lMyFileLen, #iFFile)
    If Not pbLock Then
        Close #iFFile
    End If
    
    Exit Function
EH:
    Close #iFFile
    If Not pbSkipMess Then
        If MsgBox("Could not read file... " & vbCrLf & psFilePath & vbCrLf & "(" & Err.Description & ")" & vbCrLf & vbCrLf & _
                  "The network or file is busy." & vbCrLf & "Press ""Yes"" to try again." & vbCrLf & "Press ""No"" to abort this process", vbYesNo, "File is Busy") = vbYes Then
            Resume
        End If
    End If
    
End Function

Public Sub SaveFileData(psFilePath As String, psFileData As String, Optional psDelimeter As String, Optional pbLock As Boolean = False, Optional piFFile As Integer)
'Purpose:

'Parameters :

'Returns:

'Author :

'Revision History:  SMR     Initials    Date        Description
    On Error GoTo EH
    Dim lMyFileLen As Long
    Dim iFFile As Integer
    
    iFFile = FreeFile
    piFFile = iFFile
    Open psFilePath For Binary Access Write As #iFFile
    Put #iFFile, 1, psFileData & psDelimeter
    If Not pbLock Then
        Close #iFFile
    End If
    Exit Sub
EH:
    Close #iFFile
    Err.Raise Err.Number, , App.EXEName & vbCrLf & "Public Sub SaveFileData" & vbCrLf & "Error # " & Err.Number & vbCrLf & Err.Description & vbCrLf
End Sub

Public Sub AlwaysOnTop(pForm As Form, pbSetOnTop As Boolean, Optional plHwnd As Long = 0)
    On Error GoTo EH
    Dim lFlag As Long
    Dim lhWnd As Long
    
    If plHwnd > 0 Then
        lhWnd = plHwnd
    Else
        lhWnd = pForm.hWnd
    End If
    
    If pbSetOnTop Then
        lFlag = HWND_TOPMOST
    Else
        lFlag = HWND_NOTOPMOST
    End If
    
    SetWindowPos lhWnd, lFlag, _
    pForm.Left / Screen.TwipsPerPixelX, _
    pForm.Top / Screen.TwipsPerPixelY, _
    pForm.Width / Screen.TwipsPerPixelX, _
    pForm.Height / Screen.TwipsPerPixelY, _
    SWP_NOACTIVATE Or SWP_SHOWWINDOW
    Exit Sub
EH:
    Err.Raise Err.Number, , App.EXEName & vbCrLf & "Public Sub AlwaysOnTop" & vbCrLf & "Error # " & Err.Number & vbCrLf & Err.Description & vbCrLf
End Sub


Public Sub DeleteECSPLListFile(psCommand As String)
    On Error GoTo EH
    Dim sDelFile As String
    Dim sCommand As String
    Dim sMsg As String
    
    sDelFile = psCommand
    'REMOVE the Install List
    DoEvents
    Sleep 100
    sDelFile = DeleteFile(sDelFile)
    If sDelFile <> vbNullString Then
        sMsg = sMsg & vbCrLf & vbCrLf & "Could not delete: " & vbCrLf & sDelFile & vbCrLf & sDelFile
        Err.Raise 999, , sMsg
    End If
    
    
    Exit Sub
EH:
    Err.Raise Err.Number, , App.EXEName & vbCrLf & "Public Sub DeleteECSPLListFile" & vbCrLf & "Error # " & Err.Number & vbCrLf & Err.Description & vbCrLf
End Sub

