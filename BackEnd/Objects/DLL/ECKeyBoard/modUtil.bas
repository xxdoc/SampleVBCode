Attribute VB_Name = "modUtil"
Option Explicit
'Window OS System Info
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Type OSVERSIONINFO
   dwOSVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long
   szCSDVersion As String * 128
End Type
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40
Private Const WM_CLOSE = &H10 'BGS used for Close window

Public Const NO_DB_VERSION_TABLE_ERROR As Integer = 3078
Public Const DB_VERSION_DUPLICATE_VALUE_ERROR As Integer = 3022
Public Const DB_VERSION_ITEM_NOT_FOUND_ERROR As Integer = 3265

Public Const SPI_GETWORKAREA = 48
Public Const SE_ERR_NOASSOC = 31

Public Type RECT
        left As Long
        top As Long
        Right As Long
        Bottom As Long
End Type

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Public Declare Function GetNextWindow Lib "user32" Alias "GetWindow" (ByVal hWnd As Long, ByVal wFlag As Long) As Long
Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function IsIconic Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Declare Function timeGetTime Lib "winmm.dll" () As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal sBuffer As String, lsize As Long) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Const SW_RESTORE = 9
'
' Private variables needed to support enumeration
'
Private m_hWnd As Long
Private m_Method As FindWindowPartialTypes
Private m_CaseSens As Boolean
Private m_Visible As Boolean
Private m_AppTitle As String
Public mbUCText As Boolean

'FLS File Life Span Enum
Public Enum FLSItem
    FLS0FileDir = 0
    FLS1LifeSpanDays
End Enum
Private mFSO As scripting.FileSystemObject

Private Property Get msClassName() As String
    msClassName = "modUtil"
End Property

Public Function GetTaskbarHeight() As Long
    On Error GoTo EH
    Dim lRes As Long
    Dim rectVal As RECT
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    lRes = SystemParametersInfo(SPI_GETWORKAREA, 0, rectVal, 0)
    GetTaskbarHeight = ((Screen.Height / Screen.TwipsPerPixelX) - rectVal.Bottom) * Screen.TwipsPerPixelX
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Function GetTaskbarHeight" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Function

Public Sub AlwaysOnTop(pForm As Form, pbSetOnTop As Boolean, Optional plHwnd As Long = 0)
    On Error GoTo EH
    Dim lFlag As Long
    Dim lHwnd As Long
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    If plHwnd > 0 Then
        lHwnd = plHwnd
    Else
        lHwnd = pForm.hWnd
    End If
    
    If pbSetOnTop Then
        lFlag = HWND_TOPMOST
    Else
        lFlag = HWND_NOTOPMOST
    End If
    
    SetWindowPos lHwnd, lFlag, _
    pForm.left / Screen.TwipsPerPixelX, _
    pForm.top / Screen.TwipsPerPixelY, _
    pForm.Width / Screen.TwipsPerPixelX, _
    pForm.Height / Screen.TwipsPerPixelY, _
    SWP_NOACTIVATE Or SWP_SHOWWINDOW
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Sub AlwaysOnTop" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Sub

Public Sub ShowError(psAppEXEName As String, pobjErr As ErrObject, psProc As String, Optional pFormOwner As Form, Optional psMod As String)
    
    Dim sErrMess As String

    sErrMess = Now() & vbCrLf
    sErrMess = sErrMess & "AppName: " & psAppEXEName & vbCrLf
    sErrMess = sErrMess & "ClassName: " & psMod & vbCrLf
    sErrMess = sErrMess & "ProcName: " & psProc & vbCrLf
    sErrMess = sErrMess & "ERROR # " & pobjErr.Number & vbCrLf
    sErrMess = sErrMess & pobjErr.Description & vbCrLf & vbCrLf
    If Not goUtil.goProgForm Is Nothing Then
        goUtil.goProgForm.ShowForm False
    End If
    MsgBox sErrMess, vbCritical + vbOKOnly, "ERROR in " & psProc
    
End Sub

Public Function FormWinRegPos(psAppEXEName As String, pMyForm As Form, Optional pbSave As Boolean, _
                              Optional pfrmOffset As Form, Optional pctrlOffset As Control, _
                              Optional pbUseFullCaption As Boolean = True, Optional pbUseFrmName As Boolean) As Boolean
'Purpose: This Procedure can be used by AnyForm to Get or Save the Form Position
'         from the Windows Registry using Save Setting and GetSetting :)

'Parameters : pMyForm As Form, Optional pbSave As Boolean

'Returns: FormWinRegPos Returns True Only if  Retrieving and Finds Stored  Values
'         FormWinRegPos Returns False if Retreieving and does not find Stored Values
'         FormWinRegPos Returns False When Saveing IE pbSave is Set to True.


'Author : BGS - 3/10/2000

'Revision History:  SMR     Initials    Date        Description
'                     1     BGS         3/21/2000   Added Optional pfrOffset incase you want to Offset the posn
'                                                   in realation to another form.
'                     2     BGS         10/23/2001  Changed the SECTION to enter ALL forms under FORM_POSN
'                                                   Also check for Borderstyle do Change width or height on non sizable windows :)
                    
    'This Procedure will Either Retrieve or Save Form Posn values
    'Best used on Form Load and Unload or QueryUnLoad
    Dim sCap As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    On Error GoTo EH
    With pMyForm
        If Not pbUseFullCaption Then
            If pbUseFrmName Then
                sCap = left(.Name, InStr(1, .Caption, " "))
            Else
                sCap = left(.Caption, InStr(1, .Caption, " "))
            End If
            
        Else
            If pbUseFrmName Then
                sCap = .Name & " "
            Else
                sCap = .Caption & " "
            End If
            
        End If
        If pbSave Then
            'If Saving then do this...
            'If Form was minimized or Maximized then Closed Need to Save Windowstate
            'THEN... set Back to Normal Or previous non Max or Min State then Save
            'Posn Parameters
            
            SaveSetting psAppEXEName, "FORM_POSN", .Name & sCap & "_WindowState", .WindowState
            
            If .WindowState = vbMinimized Or .WindowState = vbMaximized Then
                .Visible = True
                .WindowState = vbNormal
            End If
            
            'Save AppName...FrmName...KeyName...Value
            If pfrmOffset Is Nothing Then
                'Check to be sure windows didn't screw up and get set to something way off in la la land
                If .top < 0 Then .top = 0
                If .top > Screen.Height Then .top = Screen.Height - 100
                If .left < 0 Then .left = 0
                If .left > Screen.Width Then .left = Screen.Width - 100
                SaveSetting psAppEXEName, "FORM_POSN", .Name & sCap & "_Top", .top
                SaveSetting psAppEXEName, "FORM_POSN", .Name & sCap & "_Left", .left
                SaveSetting psAppEXEName, "FORM_POSN", .Name & sCap & "_Height", .Height
                SaveSetting psAppEXEName, "FORM_POSN", .Name & sCap & "_Width", .Width
            Else
                SaveSetting psAppEXEName, "FORM_POSN", .Name & sCap & "_Top", .top - pfrmOffset.top - pctrlOffset.top
                SaveSetting psAppEXEName, "FORM_POSN", .Name & sCap & "_Left", .left - pfrmOffset.left
                SaveSetting psAppEXEName, "FORM_POSN", .Name & sCap & "_Height", .Height
                SaveSetting psAppEXEName, "FORM_POSN", .Name & sCap & "_Width", .Width
            End If
        Else
            'If Not Saveing Must Be Getting ..
            'Need to ref AppName...FrmName...KeyName
            '(If nothing Stored Use The Exisiting Form value)
            If .WindowState = vbMinimized Or .WindowState = vbMaximized Then
                .WindowState = vbNormal
            End If
            If pfrmOffset Is Nothing Then
                .top = GetSetting(psAppEXEName, "FORM_POSN", .Name & sCap & "_Top", .top)
                .left = GetSetting(psAppEXEName, "FORM_POSN", .Name & sCap & "_Left", .left)
                If .BorderStyle = vbSizable Or .BorderStyle = vbSizableToolWindow Then
                    .Height = GetSetting(psAppEXEName, "FORM_POSN", .Name & sCap & "_Height", .Height)
                    .Width = GetSetting(psAppEXEName, "FORM_POSN", .Name & sCap & "_Width", .Width)
                End If
                'Be Sure WindowState is set last (Can't Change POSN if vbMinimized Or Maximized)
                .WindowState = GetSetting(psAppEXEName, "FORM_POSN", .Name & sCap & "_WindowState", .WindowState)
            Else
                .top = GetSetting(psAppEXEName, "FORM_POSN", .Name & sCap & "_Top", .top) + pfrmOffset.top + pctrlOffset.top
                .left = GetSetting(psAppEXEName, "FORM_POSN", .Name & sCap & "_Left", .left) + pfrmOffset.left
                If .BorderStyle = vbSizable Or .BorderStyle = vbSizableToolWindow Then
                    .Height = GetSetting(psAppEXEName, "FORM_POSN", .Name & sCap & "_Height", .Height)
                    .Width = GetSetting(psAppEXEName, "FORM_POSN", .Name & sCap & "_Width", .Width)
                End If
                'Be Sure WindowState is set last (Can't Change POSN if vbMinimized Or Maximized)
                .WindowState = GetSetting(psAppEXEName, "FORM_POSN", .Name & sCap & "_WindowState", .WindowState)
            End If
        End If
    End With

    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Function FormWinRegPos" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Function

Public Function CloseApplication(ByVal psAppCaption As String, plHwnd As Long) As Boolean
    On Error GoTo EH
    Dim lHwnd As Long
    Dim lRetVal As Long
    Dim sFoundCaption As String
    Dim sDoNotClose As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    lHwnd = plHwnd
LOOK_FOR_WINDOW:
    Do Until lHwnd = 0
        lHwnd = GetNextWindow(lHwnd, 2)
        sFoundCaption = GetCaption(lHwnd)
        If InStr(1, sFoundCaption, psAppCaption, vbTextCompare) > 0 Then
            If InStr(1, sDoNotClose, sFoundCaption, vbTextCompare) = 0 Then
                Exit Do
            End If
        End If
    Loop

    If lHwnd <> 0 Then
        If MsgBox("Do you want to close """ & sFoundCaption & """ ? ", vbQuestion + vbYesNo) = vbNo Then
            sDoNotClose = sDoNotClose & sFoundCaption
            GoTo LOOK_FOR_WINDOW
        End If
        lRetVal = PostMessage(lHwnd, WM_CLOSE, 0&, 0&)
        'BGS 10.9.2001 if we close the app have to wait until
        'it finishes closing before we allow any files to be copied
        'in the update code.
        If lRetVal <> 0 Then
            DoEvents
            Sleep 5000 '5 Seconds should be enough time if not then there comp is a DOG !
            DoEvents
        End If
        'BGS Need to reset the Starting point hwd to an Open window
        lHwnd = plHwnd
        GoTo LOOK_FOR_WINDOW
    End If
    
    If lRetVal <> 0 Then
        CloseApplication = True
    End If
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Function CloseApplication" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Function

Public Function GetCaption(ByVal lHwnd As Long) As String
    On Error GoTo EH
    Dim sInput As String
    Dim lLen As Long
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    lLen = GetWindowTextLength(lHwnd)
    sInput = String(lLen&, 0)
    Call GetWindowText(lHwnd&, sInput, lLen + 1)
    GetCaption = sInput
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Function GetCaption" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Function

Public Function FindNextWindow(psAppCaption As String, plhWndStart As Long, plhWndFound As Long) As Boolean
    On Error GoTo EH
    Dim lHwnd As Long
    Dim lRetVal As Long
    Dim sFoundCaption As String
    Dim sDoNotClose As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    lHwnd = plhWndStart
LOOK_FOR_WINDOW:
    Do Until lHwnd = 0
        lHwnd = GetNextWindow(lHwnd, 2)
        sFoundCaption = GetCaption(lHwnd)
        If InStr(1, sFoundCaption, psAppCaption, vbTextCompare) > 0 Then
            If InStr(1, sDoNotClose, sFoundCaption, vbTextCompare) = 0 Then
                plhWndFound = lHwnd
                FindNextWindow = True
                Exit Do
            End If
        End If
    Loop
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Function FindNextWindow" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Function

'************************Begin Note 1.28.2002 **************************

'*The following Block of code was obtained from http://www.mvps.org/vb/*
'The source code is free to use within any application as long as the actual
'uncompiled source code is not sold or distributed to other programmers.
' *********************************************************************
'  Copyright ©1995-2000 Karl E. Peterson, All Rights Reserved
' *********************************************************************
'  You are free to use this code within your own applications, but you
'  are expressly forbidden from selling or otherwise distributing this
'  source code without prior written consent.
' *********************************************************************

Public Function AppActivatePartial(AppTitle As String, Optional Method As FindWindowPartialTypes = FwpStartsWith, Optional CaseSensitive As Boolean = False) As Long
   Dim hWndApp As Long
   '
   ' Retrieve window handle for first top-level window
   ' that starts with or contains the passed string.
   '
   hWndApp = FindWindowPartial(AppTitle, Method, CaseSensitive, True)
   If hWndApp Then
      '
      ' Switch to it, restoring if need be.
      '
      If IsIconic(hWndApp) Then
         Call ShowWindow(hWndApp, SW_RESTORE)
      End If
      Call SetForegroundWindow(hWndApp)
      AppActivatePartial = hWndApp
   End If
End Function

Public Function FindWindowPartial(AppTitle As String, _
   Optional Method As FindWindowPartialTypes = FwpStartsWith, _
   Optional CaseSensitive As Boolean = False, _
   Optional MustBeVisible As Boolean = False) As Long
   '
   ' Reset all search parameters.
   '
   m_hWnd = 0
   m_Method = Method
   m_CaseSens = CaseSensitive
   m_AppTitle = AppTitle
   '
   ' Upper-case search string if case-insensitive.
   '
   If m_CaseSens = False Then
      m_AppTitle = UCase$(m_AppTitle)
   End If
   '
   ' Fire off enumeration, and return m_hWnd when done.
   '
   Call EnumWindows(AddressOf EnumWindowsProc, MustBeVisible)
   FindWindowPartial = m_hWnd
End Function

Private Function EnumWindowsProc(ByVal hWnd As Long, ByVal lParam As Long) As Long
   Static WindowText As String
   Static nRet As Long
   '
   ' Make sure we meet visibility requirements.
   '
   If lParam Then 'window must be visible
      If IsWindowVisible(hWnd) = False Then
         EnumWindowsProc = True
         Exit Function
      End If
   End If
   '
   ' Retrieve windowtext (caption)
   '
   WindowText = Space$(256)
   nRet = GetWindowText(hWnd, WindowText, Len(WindowText))
   If nRet Then
      '
      ' Clean up window text and prepare for comparison.
      '
      WindowText = left$(WindowText, nRet)
      If m_CaseSens = False Then
         WindowText = UCase$(WindowText)
      End If
      '
      ' Use appropriate method to determine if
      ' current window's caption either starts
      ' with, contains, or matches passed string.
      '
      Select Case m_Method
         Case FwpStartsWith
            If InStr(WindowText, m_AppTitle) = 1 Then
               m_hWnd = hWnd
            End If
         Case FwpContains
            If InStr(WindowText, m_AppTitle) <> 0 Then
               m_hWnd = hWnd
            End If
         Case FwpMatches
            If WindowText = m_AppTitle Then
               m_hWnd = hWnd
            End If
      End Select
   End If
   '
   ' Return True to continue enumeration if we haven't
   ' found what we're looking for.
   '
   EnumWindowsProc = (m_hWnd = 0)
End Function
'************************End Note 1.28.2002 **************************

'*The Block of code above was obtained from http://www.mvps.org/vb/*

Public Function FileExists(strFile As String, Optional pbDirOnly As Boolean) As Boolean
    '10.24.2002 Use the File System Object since it is Superior to Dir function.
    On Error GoTo EH
    If mFSO Is Nothing Then
        Set mFSO = New scripting.FileSystemObject
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
        Set mFSO = New scripting.FileSystemObject
    End If
    
    mFSO.DeleteFile psFilePath, True
    
    Exit Function
EH:
    DeleteFile = Err.Number & vbCrLf & vbCrLf & Err.Description
End Function

Public Function CopyFile(psSourceFilePath As String, psDestFilePath As String) As String
    On Error GoTo EH
    If mFSO Is Nothing Then
        Set mFSO = New scripting.FileSystemObject
    End If
    
    mFSO.CopyFile psSourceFilePath, psDestFilePath, True
    
    Exit Function
EH:
    CopyFile = Err.Number & vbCrLf & vbCrLf & Err.Description
End Function

Public Function DeleteDir(psDirPath As String) As String
    On Error GoTo EH
    If mFSO Is Nothing Then
        Set mFSO = New scripting.FileSystemObject
    End If
    
    mFSO.DeleteFolder psDirPath, True

    Exit Function
EH:
    DeleteDir = Err.Number & vbCrLf & vbCrLf & Err.Description
End Function

Public Function MakeDir(psDirPath As String) As String
    On Error GoTo EH
    If mFSO Is Nothing Then
        Set mFSO = New scripting.FileSystemObject
    End If
    
    mFSO.CreateFolder psDirPath
    
    Exit Function
EH:
    MakeDir = Err.Number & vbCrLf & vbCrLf & Err.Description
End Function

Public Function CopyDir(psSourceDirPath As String, psDestDirPath As String) As String
    On Error GoTo EH
    If mFSO Is Nothing Then
        Set mFSO = New scripting.FileSystemObject
    End If
    
    mFSO.CopyFolder psSourceDirPath, psDestDirPath, True
    
    Exit Function
EH:
    CopyDir = Err.Number & vbCrLf & vbCrLf & Err.Description
End Function

Public Function GetSystemDir() As String
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim sDir As String
    Dim lRet As Long
    
    sDir = Space(260)
    lRet = GetSystemDirectory(sDir, Len(sDir))
    sDir = left(sDir, lRet)
    GetSystemDir = sDir
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Function GetSystemDir" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
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
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    iFFile = FreeFile
    piFFile = iFFile
    Open psFilePath For Binary Access Write As #iFFile
    Put #iFFile, 1, psFileData & psDelimeter
    If Not pbLock Then
        Close #iFFile
    End If
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Close #iFFile
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Sub SaveFileData" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Sub

Public Function isControlArray(MyForm As Form, MyControl As Control) As Boolean
    
    'BGS 8/1/1999 Added this function to determin if a Control is part of
    'a control array or not. I had to do this because VB does not have a
    'function that figures this out. (IsArray does not work on Control Arrays)

    On Error GoTo EH
    Dim MyCount As Integer
    Dim CheckMyControl As Control
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    For Each CheckMyControl In MyForm.Controls
        If CheckMyControl.Name = MyControl.Name Then
            MyCount = MyCount + 1
            If MyCount > 1 Then
                Exit For
            End If
        End If
    Next
    
    isControlArray = MyCount - 1
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Function isControlArray" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Function

Public Sub SelText(pTextBox As Control)
'Purpose: Highlights All Text

'Parameters : TextBox

'Returns: Just Highlights TExt in the pTextBox

'Author : BGS - 3/10/2000

'Revision History:  SMR     Initials    Date        Description
    On Error GoTo EH
    With pTextBox
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
    Exit Sub
EH:
    Err.Clear
End Sub


Public Sub UCText(pControl As Control)
    On Error GoTo EH
    Dim iSelpos As Integer
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    If Not mbUCText Then
        mbUCText = True
        iSelpos = pControl.SelStart
        With pControl
            .Text = UCase(.Text)
            .SelStart = iSelpos
            .SelLength = 0
        End With
        mbUCText = False
    End If
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    mbUCText = False
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Sub UCText" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Sub

Public Function DBTblExists(psDBPath As String, psDBTbl As String) As Boolean
    On Error GoTo EH
    Dim WS As Workspace
    ' SourceVariables
    Dim dbSource As Database
    Dim tblSource As TableDef
    Dim sTemp As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    'BGS 1.10.2001 Need to be sure the default directory is
    'Ap.path or will get strange errors when creating WorkSpace
    ChDir App.Path
    Set WS = CreateWorkspace("", "admin", "", dbUseJet)
    Set dbSource = WS.OpenDatabase(psDBPath, False, True)
    
    For Each tblSource In dbSource.TableDefs
        sTemp = tblSource.Name
        If InStr(1, psDBTbl, sTemp, vbTextCompare) > 0 Then
            DBTblExists = True
            GoTo CLEANUP
        End If
    Next
    DBTblExists = False
CLEANUP:
    Set tblSource = Nothing
    dbSource.Close: Set dbSource = Nothing
    WS.Close: Set WS = Nothing
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Set tblSource = Nothing
    Set dbSource = Nothing
    Set WS = Nothing
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Function DBTblExists" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Function

Public Function DynamicArraySet(pVarArray As Variant) As Boolean
    'Purpose: To see if a Dynamic array has been set
    'Parameters : pVarArray As Variant: Send in any Dynamic array data type
    'Returns: True if has been set, false if not
    'Author : BGS-3/24/2000
    'Revision History:  SMR     Initials    Date    Description
    
    On Error GoTo NOT_SET
    Dim iRet As Integer
    
    If IsArray(pVarArray) Then
        iRet = LBound(pVarArray, 1)
        'if the Lbound call to the first dimension of
        'pVarArray does not error then the dynamic array must
        'be set so...
        DynamicArraySet = True
        Exit Function
    End If
    
NOT_SET:
    DynamicArraySet = False
End Function

Public Sub CompRepair(psAppEXEName As String, psSourcePath As String, Optional pbSkipBackup As Boolean)
    On Error GoTo EH
    Dim sTemp As String
    Dim sDBBackup As String
    Dim sDBName As String
    Dim lKey As Long
    Dim sKey As String
    Dim sPassword As String
    Dim lErrorNum As Long
    Dim sErrorDescription As String
    
    'BGS 12.15.2000 Make the Temp Same Path as the Original
    sTemp = left(psSourcePath, InStrRev(psSourcePath, "\") - 1) & "\DB.tmp"
    sDBBackup = left(psSourcePath, InStrRev(psSourcePath, ".") - 1)
    sDBBackup = sDBBackup & "_BackUp.db"
    sDBName = Mid(psSourcePath, InStrRev(psSourcePath, "\") + 1)
    'Kill Temp  if it there for some reason
    If FileExists(sTemp) Then
        If MsgBox("There were system problems the last time you Used Compact Repair" & vbCrLf & _
               "The " & sDBName & " can be restored with " & vbCrLf & sDBBackup & "." & vbCrLf & _
               "Press OK to Restore or Cancel to NOT Restore " & psAppEXEName & ".", vbOKCancel) = vbOK Then
               Kill sTemp
               goUtil.utCopyFile sDBBackup, psSourcePath
        Else
'            End 'BAIL !!!
        End If
    End If
    'Copy the DB to the Temp PAth
    goUtil.utCopyFile psSourcePath, sTemp
    
    'BGS 12.13.2000 Make Backup Data base
    'In case there is some problems
    If Not pbSkipBackup Then
        goUtil.utCopyFile psSourcePath, sDBBackup
    End If
    
    'Kill the Source since it is copied to Temp
    goUtil.utDeleteFile psSourcePath
    
    'Compact the Temp and send it back to where it was originally
    'copied from
Set_Key:
    On Error Resume Next
    lKey = lKey + 1
    sKey = CStr(lKey)
    sPassword = goUtil.DB_PASSWORD(sKey)
    sPassword = goUtil.Decode(sPassword)
    If Err.Number <> 0 Then
        lErrorNum = Err.Number
        sErrorDescription = Err.Description
        Err.Clear
        On Error GoTo EH
        Err.Raise lErrorNum, , sErrorDescription & vbCrLf & "Invalid Password!" & vbCrLf & msClassName & vbCrLf & "Public Sub CompRepair"
    End If
    'Always Set the Password to the Latest Version Which will always be
    'Found in Key 1 goUtil.DB_PASSWORD("1")... On the other hand
    'If the current password is not the latest keep checking until that password can be updated to
    'the Most current Password.
    CompactDatabase sTemp, psSourcePath, dbLangGeneral & ";PWD=" & goUtil.Decode(goUtil.DB_PASSWORD("1")), , ";PWD=" & sPassword
    If Err.Number <> 0 Then
        'Check the Error Message, if it is an invalid Password...
        'Keep Checking the Dictionary until one works.
        If Err.Number = 3031 Then
            Err.Clear
            GoTo Set_Key
        Else
            lErrorNum = Err.Number
            sErrorDescription = Err.Description
            Err.Clear
            On Error GoTo EH
            Err.Raise lErrorNum, , sErrorDescription & vbCrLf & msClassName & vbCrLf & "Public Sub CompRepair"
        End If
    End If
    'Finally kill the Temp since we done with it
    goUtil.utDeleteFile sTemp
    
    Exit Sub
EH:
    'if for some reason we errored while compacting and repairing
    'See if we can recover the Source back to what it was before
    'If this fails the Data Base we were trying to compact was really
    'Hosed up.
    lErrorNum = Err.Number
    sErrorDescription = Err.Description
    If Not FileExists(psSourcePath) Then
        If FileExists(sTemp) Then
            goUtil.utCopyFile sTemp, psSourcePath
            goUtil.utDeleteFile sTemp
        End If
    End If
    Err.Raise lErrorNum, , sErrorDescription
End Sub

Public Sub BubbleSort(pvArray As Variant, Optional psEndString As String, Optional pbRebillPresent As Boolean)
    On Error GoTo EH
    Dim lMainLoop As Long
    Dim lSubLoop As Long
    Dim sStringA As String
    Dim sStringB As String
    Dim sTempA As String
    Dim sTempB As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    If DynamicArraySet(pvArray) Then
        For lMainLoop = UBound(pvArray) To LBound(pvArray) Step -1
            For lSubLoop = LBound(pvArray) + 1 To lMainLoop
                'Only sort non nullstrings
                If pvArray(lSubLoop) <> vbNullString Then
                    'BGS 11.26.2001 check for Rebill
                    If InStr(1, pvArray(lSubLoop), "R", vbTextCompare) > 0 Then
                        pbRebillPresent = True
                    End If
                    If psEndString > vbNullString Then
                        'BGS 7.12.2001 need to sort with S as "A" because the
                        'Supplements have precendence over rebilling in the sort
                        sStringA = left(pvArray(lSubLoop - 1), InStr(1, pvArray(lSubLoop - 1), psEndString) - 1)
                        sTempA = Right(sStringA, 6)
                        sStringA = Replace(sStringA, sTempA, vbNullString)
                        sTempA = Replace(sTempA, "S", "A", , , vbTextCompare)
                        sStringA = sStringA & sTempA
                        
                        sStringB = left(pvArray(lSubLoop), InStr(1, pvArray(lSubLoop), psEndString) - 1)
                        sTempB = Right(sStringB, 6)
                        sStringB = Replace(sStringB, sTempB, vbNullString)
                        sTempB = Replace(sTempB, "S", "A", , , vbTextCompare)
                        sStringB = sStringB & sTempB
                        
                        If sStringA > sStringB Then
                            Call SwitchPlace(pvArray(lSubLoop - 1), pvArray(lSubLoop))
                        End If
                    Else
                        If pvArray(lSubLoop - 1) > pvArray(lSubLoop) Then
                            Call SwitchPlace(pvArray(lSubLoop - 1), pvArray(lSubLoop))
                        End If
                    End If
                End If
            Next
        Next
    End If
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Sub BubbleSort" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Sub

Public Sub SwitchPlace(a As Variant, b As Variant)
  Dim c As Variant
  c = a
  a = b
  b = c
End Sub

Public Function FormExists(goForms As Object, psFormName As String) As Boolean
    On Error GoTo EH
    Dim iCount As Integer
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    For iCount = 0 To goForms.Count - 1
        If UCase(goForms(iCount).Name) = UCase(psFormName) Then
            FormExists = True
            Exit For
        End If
    Next
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Function FormExists" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Function

Public Function FindSetForm(goForms As Object, psFormName As String, pForm As Form) As Boolean
    On Error GoTo EH
    Dim iCount As Integer
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    For iCount = 0 To goForms.Count - 1
        If UCase(goForms(iCount).Name) = UCase(psFormName) Then
            Set pForm = goForms(iCount)
            FindSetForm = True
            Exit For
        End If
    Next
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Function FindSetForm" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Function

Public Function CleanSQLString(pvText As Variant) As Variant
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    'BGS 11.21.2001 add some more cleaning
    If VarType(pvText) = vbString Then
        CleanSQLString = Replace(pvText, "'", "''")
    Else
        CleanSQLString = pvText
    End If
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    CleanSQLString = pvText
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Function CleanSQLString" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Function

Public Function CleanValString(psValText As String) As String
    'Val function Bug in VB6
    'http://msdn.microsoft.com/vbasic/productinfo/previous/vb6/tips/01pasttips.asp
    'Need to parse out both % and !  because these trailing equate to Double and Single
    'and Val() bugs because it can't convert Double or single into integer
    On Error GoTo EH
    Dim lCount As Long
    Dim sTemp As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    sTemp = psValText
    
    sTemp = Trim(sTemp)
    For lCount = 1 To Len(sTemp)
        If IsNumeric(Mid(sTemp, lCount, 1)) Then
            Exit For
        End If
    Next
    sTemp = Mid(sTemp, lCount)
    
    sTemp = Replace(sTemp, ",", vbNullString)
    sTemp = Replace(sTemp, "%", vbNullString)
    sTemp = Replace(sTemp, "!", vbNullString)
    CleanValString = sTemp
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Function CleanValString" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Function

Public Sub CleanValTextBox(pvText As Variant)
    On Error GoTo EH
    Dim lPos As Long
    Dim sTemp As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    'Issue 180 180 Error #76 Error in Function_ CreateCat; Form SetupNewCat
    'Clean out any Special chars \/:*?"<>|
    If IsObject(pvText) Then
        lPos = pvText.SelStart
        sTemp = pvText.Text
    Else
        sTemp = CStr(pvText)
    End If
    
    sTemp = Replace(sTemp, "%", vbNullString, , , vbBinaryCompare)
    sTemp = Replace(sTemp, "!", vbNullString, , , vbBinaryCompare)
    
    If IsObject(pvText) Then
        pvText.Text = sTemp
        pvText.SelStart = lPos
    Else
        pvText = sTemp
    End If
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Sub CleanValTextBox" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Sub

Public Function JoinToAddress(psStreet As String, psCity As String, psState As String, psZip As String) As String
    On Error GoTo EH
    Dim sTemp As String
    Dim sAddress As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    sTemp = Trim(Replace(psStreet, vbCrLf, vbNullString)) & String(2, " ") & vbCrLf
    sAddress = sTemp
    sTemp = Trim(Replace(psCity, vbCrLf, vbNullString))
    sTemp = Replace(sTemp, ",", vbNullString) & ", "
    sAddress = sAddress & sTemp
    sTemp = Trim(Replace(psState, vbCrLf, vbNullString)) & " "
    sAddress = sAddress & sTemp
    sTemp = Trim(Replace(psZip, vbCrLf, vbNullString))
    sAddress = sAddress & sTemp
    
    JoinToAddress = sAddress
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Private Function JoinToAddress" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Function

Public Sub UpdateAddress(psAddress As String, _
                         psZip As String, _
                         psState As String, _
                         psCity As String, _
                         psStreet As String)
    On Error GoTo EH
    Dim sTemp As String
    Dim sAddress As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    sTemp = Trim(Replace(psStreet, vbCrLf, vbNullString)) & String(2, " ") & vbCrLf
    sAddress = sTemp
    sTemp = Trim(Replace(psCity, vbCrLf, vbNullString))
    sTemp = Replace(sTemp, ",", vbNullString) & ", "
    sAddress = sAddress & sTemp
    sTemp = Trim(Replace(psState, vbCrLf, vbNullString)) & " "
    sAddress = sAddress & sTemp
    sTemp = Trim(Replace(psZip, vbCrLf, vbNullString))
    sAddress = sAddress & sTemp
    
     
    If Right(sAddress, 5) = vbCrLf & ", " & " " Then
        On Error Resume Next
        sAddress = RTrim(left(sAddress, InStrRev(sAddress, vbCrLf) - 1))
        Dim l As Long
'        For l = 1 To Len(sAddress)
'            Debug.Print Mid(sAddress, l, 1) & " ---->" & Asc(Mid(sAddress, l, 1))
'        Next
    End If
    
    psAddress = sAddress
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Sub UpdateAddress" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Sub

Public Sub FillAddressFields(psAddress As String, _
                             psZip As String, _
                             psState As String, _
                             psCity As String, _
                             psStreet As String)
    On Error GoTo EH
    Dim sTemp As String
    Dim sAddress As String
    Dim sValTemp As String
    Dim lPos As Long
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    sAddress = Trim(Replace(psAddress, vbCrLf, vbNullString))
    
    'Zip code
    If InStr(1, sAddress, " ", vbBinaryCompare) > 0 Then
        sTemp = Trim(Mid(sAddress, InStrRev(sAddress, " ", , vbBinaryCompare)))
        'Val function Bug in VB6
        'http://msdn.microsoft.com/vbasic/productinfo/previous/vb6/tips/01pasttips.asp
        'Need to parse out both % and !  because these trailing equate to Double and Single
        'and Val bugs because it can't convert Double or single into integer
        sValTemp = Replace(sTemp, "-", vbNullString)
        If Val(CleanValString(sValTemp)) > 0 Then
            If Len(sTemp) >= 5 Then
                'Issue 243 9.10.2002 Copy Button for Address Chops of Letters
                'Need to use string reverse to get proper Left length
                'Using Replace can not work here, must use right to left logic.
                lPos = InStrRev(sAddress, sTemp, , vbBinaryCompare) - 1
                If lPos >= 0 Then
                    sAddress = Trim(left(sAddress, lPos))
                End If
                sTemp = Replace(sTemp, ",", vbNullString)
                psZip = sTemp
            Else
                psZip = vbNullString
                psState = vbNullString
                psCity = vbNullString
                GoTo ADDRESS
            End If
        Else
            psZip = vbNullString
            psState = vbNullString
            psCity = vbNullString
            GoTo ADDRESS
        End If
    Else
        psZip = vbNullString
        psState = vbNullString
        psCity = vbNullString
        GoTo ADDRESS
    End If
    
    'State
    If Len(sAddress) > 2 Then
        sTemp = Right(sAddress, 2)
        If Val(CleanValString(sTemp)) = 0 Then
            'Issue 243 9.10.2002 Copy Button for Address Chops of Letters
            lPos = InStrRev(sAddress, sTemp, , vbBinaryCompare) - 1
            If lPos >= 0 Then
                sAddress = Trim(left(sAddress, lPos))
            End If
            sTemp = Replace(sTemp, ",", vbNullString)
            psState = sTemp
        Else
            psState = vbNullString
            psCity = vbNullString
            GoTo ADDRESS
        End If
    Else
        psState = vbNullString
        psCity = vbNullString
        GoTo ADDRESS
    End If
    
    'City
    If InStr(1, sAddress, S_z, vbBinaryCompare) > 0 Then
        sTemp = Trim(Mid(sAddress, InStrRev(sAddress, S_z, , vbBinaryCompare)))
        If Val(CleanValString(sTemp)) = 0 Then
            'Issue 243 9.10.2002 Copy Button for Address Chops of Letters
            lPos = InStrRev(sAddress, sTemp, , vbBinaryCompare) - 1
            If lPos >= 0 Then
                sAddress = Trim(left(sAddress, lPos))
            End If
            sTemp = Replace(sTemp, ",", vbNullString)
            sTemp = Replace(sTemp, S_z, vbNullString)
            sTemp = Replace(sTemp, Chr(32), Chr(160))
            psCity = sTemp
        Else
            psCity = vbNullString
        End If
    'City
    ElseIf InStr(1, sAddress, String(2, Chr(32)), vbBinaryCompare) > 0 Then
        sTemp = Trim(Mid(sAddress, InStrRev(sAddress, String(2, Chr(32)), , vbBinaryCompare)))
        If Val(CleanValString(sTemp)) = 0 Then
            'Issue 243 9.10.2002 Copy Button for Address Chops of Letters
            lPos = InStrRev(sAddress, sTemp, , vbBinaryCompare) - 1
            If lPos >= 0 Then
                sAddress = Trim(left(sAddress, lPos))
            End If
            sTemp = Replace(sTemp, ",", vbNullString)
            sTemp = Replace(sTemp, S_z, vbNullString)
            sTemp = Replace(sTemp, Chr(32), Chr(160))
            psCity = sTemp
        Else
            psCity = vbNullString
        End If
    ElseIf InStr(1, sAddress, String(1, Chr(32)), vbBinaryCompare) > 0 Then
        sTemp = Trim(Mid(sAddress, InStrRev(sAddress, String(1, Chr(32)), , vbBinaryCompare)))
        If Val(CleanValString(sTemp)) = 0 Then
            'Issue 243 9.10.2002 Copy Button for Address Chops of Letters
            lPos = InStrRev(sAddress, sTemp, , vbBinaryCompare) - 1
            If lPos >= 0 Then
                sAddress = Trim(left(sAddress, lPos))
            End If
            sTemp = Replace(sTemp, ",", vbNullString)
            sTemp = Replace(sTemp, S_z, vbNullString)
            sTemp = Replace(sTemp, Chr(32), Chr(160))
            psCity = sTemp
        Else
            psCity = vbNullString
        End If
    Else
        psCity = vbNullString
    End If
ADDRESS:
    'Address
    sAddress = Replace(sAddress, ",", vbNullString)
    sAddress = Replace(sAddress, S_z, vbNullString)
    psStreet = sAddress
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Sub FillAddressFields" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Sub

Public Function ValidSSN(psSSN As String) As String
    'Returns "Error" if invliad SSN
    On Error GoTo EH
    Dim sSSN As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    sSSN = psSSN

    sSSN = Val(CleanValString(sSSN))
    If Len(sSSN) < 8 Or Len(sSSN) > 9 Then
        sSSN = "Error"
    End If
    
    ValidSSN = sSSN
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Function ValidSSN" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Function
        
Public Function ValidDate(psDate As String) As String
    On Error GoTo EH
    'will retunrn "12:00:00 AM" if invalid
    Dim sDate As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    sDate = psDate
    
    If IsDate(sDate) Then
        'BGS 5.8.2002 Check to see if they entered like 12:00AM
        'it will default to year 1899 All dates enterd should be at least
        '1900 and above. Unless we have over 100 year old claims.. Hmm you think ?
        If Format(sDate, "YYYY") > 1899 Then
            sDate = Format(sDate, "MM/DD/YYYY")
        Else
            sDate = NULL_DATE
        End If
    Else
        sDate = NULL_DATE
    End If
    
    ValidDate = sDate
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Function ValidDate" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Function
    
Public Function GetAppVSInfo(psAppEXEName As String, psAppPath As String) As String
    On Error GoTo EH
    Dim vOPaths As Variant
    Dim lCount As Long
    Dim sFI As String
    Dim sText As String
    Dim sEXE As String
    Dim vEXE As Variant
    Dim colEXE As Collection
    Dim oFI As V2ECKeyBoard.clsFileVersion
    Dim FI As V2ECKeyBoard.FILE_INFORMATION
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    If psAppEXEName = vbNullString Then
        Set oFI = New V2ECKeyBoard.clsFileVersion
        FI = oFI.GetFileInformation(psAppPath)
        sText = sText & "(" & FI.cFilename & ") "
        sText = sText & "VS " & FI.nVerMajor & "." & FI.nVerMinor & "." & FI.nVerRevision & " "
        sText = sText & FI.dtLastModifyTime & vbCrLf
        GetAppVSInfo = sText
        GoTo CLEAN_UP
    End If
    
    sEXE = Dir(psAppPath & "\*.exe")
    Do Until sEXE = vbNullString
        If colEXE Is Nothing Then
            Set colEXE = New Collection
        End If
        colEXE.Add sEXE, sEXE
        sEXE = Dir
    Loop
    
    Set oFI = New V2ECKeyBoard.clsFileVersion
    
    If Not colEXE Is Nothing Then
        For Each vEXE In colEXE
            sEXE = vEXE
            FI = oFI.GetFileInformation(psAppPath & "\" & sEXE)
            sText = sText & "(" & FI.cFilename & ") "
            sText = sText & "VS " & FI.nVerMajor & "." & FI.nVerMinor & "." & FI.nVerRevision & " "
            sText = sText & FI.dtLastModifyTime & vbCrLf
        Next
    End If
    
    GetAppVSInfo = sText
    
    vOPaths = GetECSWinSysObjectsPaths
    If IsArray(vOPaths) Then
        For lCount = LBound(vOPaths, 1) To UBound(vOPaths, 1)
            FI = oFI.GetFileInformation(vOPaths(lCount))
            sFI = sFI & "(" & FI.cFilename & ") "
            sFI = sFI & "VS " & FI.nVerMajor & "." & FI.nVerMinor & "." & FI.nVerRevision & " "
            sFI = sFI & FI.dtLastModifyTime & vbCrLf
        Next
        GetAppVSInfo = GetAppVSInfo & sFI
    End If
CLEAN_UP:
    'Cleanup
    If Not oFI Is Nothing Then
        Set oFI = Nothing
    End If
    Set colEXE = Nothing
  Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    If Not oFI Is Nothing Then
        Set oFI = Nothing
    End If
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Function GetAppVSInfo" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Function
    
Public Function GetECSWinSysObjectsPaths() As Variant
    On Error GoTo EH
    Dim sBaseDLL As String
    Dim sDLL As String
    Dim sBaseOCX As String
    Dim sOCX As String
    Dim saryObjectPaths()
    Dim lCount As Long
    Dim bFound As Boolean
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim sSystemDir As String
    
    sSystemDir = GetSystemDir
    
    sBaseDLL = sSystemDir & "\ECS\DLL"
    sBaseOCX = sSystemDir & "\ECS\OCX"
   

    If FileExists(sBaseDLL, True) Then
        sDLL = Dir(sBaseDLL & "\*.dll")
        If sDLL > vbNullString Then
            bFound = True
        End If
        Do Until sDLL = vbNullString
            lCount = lCount + 1
            ReDim Preserve saryObjectPaths(1 To lCount)
            saryObjectPaths(lCount) = sBaseDLL & "\" & sDLL
            sDLL = Dir
        Loop
    End If
    
    If FileExists(sBaseOCX, True) Then
        sOCX = Dir(sBaseOCX & "\*.ocx")
        If sOCX > vbNullString Then
            bFound = True
        End If
        Do Until sOCX = vbNullString
            lCount = lCount + 1
            ReDim Preserve saryObjectPaths(1 To lCount)
            saryObjectPaths(lCount) = sBaseOCX & "\" & sOCX
            sOCX = Dir
        Loop
    End If
            
    GetECSWinSysObjectsPaths = saryObjectPaths
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Function GetECSWinSysObjectsPaths" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Function

Public Function FindCBOItem(psSearchText As String, pCBO As ComboBox, piPos As Integer) As String
    On Error GoTo EH
    Dim iCount As Integer
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    If pCBO.ListCount > 0 Then
        For iCount = 0 To pCBO.ListCount - 1
            If InStr(1, left(pCBO.List(iCount), piPos), psSearchText, vbTextCompare) > 0 Then
                FindCBOItem = pCBO.List(iCount)
                Exit Function
            End If
        Next
        'BGS if we did not find the search item then pass it back what was sent in
        FindCBOItem = psSearchText
    End If
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Function FindCBOItem" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Function

Public Function Validate(Optional poForm As Object, Optional poControl As Object) As Boolean
    On Error GoTo EH
    Dim MyControl As Control
    Dim sValidMess As String
    Dim lPos As Long
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    Validate = True
    If poControl Is Nothing Then
        For Each MyControl In poForm.Controls
            If TypeOf MyControl Is TextBox Then
                ValidControl MyControl, sValidMess, Validate
            End If
        Next
    Else
        ValidControl poControl, sValidMess, Validate
    End If
    
    If Not Validate Then
        MsgBox "Please fix the following value(s)..." & vbCrLf & vbCrLf & sValidMess, vbExclamation + vbOKOnly, "Validation"
    End If
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Function Validate" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Function

Private Sub ValidControl(pMyControl As Object, psValidMess As String, pbValidate As Boolean)
    On Error GoTo EH
    Dim lPos As Long
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim MyTextBox As TextBox
        
        Set MyTextBox = pMyControl
    
        'Be sure to trim the text box
        MyTextBox.Text = Trim(MyTextBox.Text)
        
        'Check for UCASE
        If InStr(1, MyTextBox.Tag, "UCASE", vbTextCompare) > 0 Then
            MyTextBox.Text = UCase(MyTextBox.Text)
        End If
        
        'Check for Alpha Numeric
        If InStr(1, MyTextBox.Tag, "ALPHANUM", vbTextCompare) > 0 Then
            MyTextBox.Text = goUtil.utScrubAlphaNumeric(MyTextBox.Text)
        End If
    
        'Numeric validation
        If InStr(1, MyTextBox.Tag, "Numeric", vbTextCompare) > 0 Then
            If Not IsNumeric(MyTextBox.Text) Then
                MyTextBox.Text = 0
            Else
                If CDbl(MyTextBox.Text) < 0 Then
                    MyTextBox.Text = 0
                End If
            End If
        End If
        
        'Hours in Decimal
        If InStr(1, MyTextBox.Tag, "HoursInDecimal", vbTextCompare) > 0 Then
            If Not IsNumeric(MyTextBox.Text) Then
                MyTextBox.Text = "0.00"
            Else
                If CDbl(MyTextBox.Text) < 0 Then
                    MyTextBox.Text = "0.00"
                Else
                    MyTextBox.Text = Format(CDbl(MyTextBox.Text), "0.00")
                End If
            End If
        End If
        
        'Currency
        If InStr(1, MyTextBox.Tag, "Currency", vbTextCompare) > 0 Then
            If Not IsNumeric(MyTextBox.Text) Then
                MyTextBox.Text = "0.00"
            Else
                MyTextBox.Text = Format(CCur(MyTextBox.Text), "#,###,###,##0.00")
            End If
        End If
        
        'Zip Code
        If InStr(1, MyTextBox.Tag, "ZipCode5", vbTextCompare) > 0 Then
            If Not IsNumeric(MyTextBox.Text) Then
                MyTextBox.Text = 0
            ElseIf CDbl(MyTextBox.Text) < 0 Then
                MyTextBox.Text = 0
            End If
            MyTextBox.Text = Format(MyTextBox.Text, "00000")
        ElseIf InStr(1, MyTextBox.Tag, "ZipCode4", vbTextCompare) > 0 Then
            If Not IsNumeric(MyTextBox.Text) Then
                MyTextBox.Text = 0
            ElseIf CDbl(MyTextBox.Text) < 0 Then
                MyTextBox.Text = 0
            End If
            MyTextBox.Text = Format(MyTextBox.Text, "0000")
        End If
        
        'Percent Validation
        If InStr(1, MyTextBox.Tag, "Percent", vbTextCompare) > 0 Then
            If Not IsNumeric(MyTextBox.Text) Then
                If MyTextBox.Text = vbNullString Then
                    MyTextBox.Text = "0.000"
                Else
                    psValidMess = psValidMess & MyTextBox.Text & " Is not a valid percent!" & vbCrLf
                    pbValidate = False
                End If
            Else
                lPos = InStr(1, MyTextBox.Text, ".", vbBinaryCompare)
                If CDbl(MyTextBox.Text) > 100 Or CDbl(MyTextBox.Text) < 0 Then
                    psValidMess = psValidMess & MyTextBox.Text & " Is not a valid percent!" & vbCrLf
                    pbValidate = False
                ElseIf lPos > 0 And InStr(1, MyTextBox.Tag, "TaxPercent", vbTextCompare) > 0 Then
                    'issue 178 Force Taxes for Texas, New Mexico, and West Virginia
                    'The percent will need to go 3 digits to right of decimal for all
                    'states not just Texas New Mexico West virginia.
                    If Len(Mid(MyTextBox.Text, lPos + 1)) < 3 Then
TAX_PERCENT:
                        psValidMess = psValidMess & MyTextBox.Text & " Tax Percent must have at least 3 decimal places!" & vbCrLf
                        pbValidate = False
                    End If
                ElseIf lPos = 0 And InStr(1, MyTextBox.Tag, "TaxPercent", vbTextCompare) > 0 Then
                    If Val(MyTextBox.Text) <> 0 Then
                        GoTo TAX_PERCENT
                    End If
                End If
            End If
        End If
        
        'Date Validation
        If InStr(1, MyTextBox.Tag, "Date", vbTextCompare) > 0 Then
            If ValidDate(MyTextBox.Text) = NULL_DATE Then
                MyTextBox.Text = vbNullString
            Else
                MyTextBox.Text = Format(MyTextBox.Text, "MM/DD/YYYY")
            End If
        End If
        
        'Directory Validation
        If InStr(1, MyTextBox.Tag, "Directory", vbTextCompare) > 0 Then
            If Not goUtil.utFileExists(MyTextBox.Text, True) Then
                MyTextBox.Text = vbNullString
            End If
        End If
        
        'FileOrFolderName
        If InStr(1, MyTextBox.Tag, "FileOrFolderName", vbTextCompare) > 0 Then
            CleanFileFolderName MyTextBox
        End If
        
        'cleanup
        Set MyTextBox = Nothing
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Private Sub ValidControl" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Sub

Public Sub EnterUserPass(psSection As String, psUserName As String, pvPass As Variant, _
                         Optional poForm As Object)
    On Error GoTo EH
    Dim sCryptUserName As String
    Dim sCryptPass As String
    Dim sRet As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    'If pass is control that has text then we are
    'prompting for confirmation of the old password first if it exists
    'and then re enter the new password
    If IsObject(pvPass) Then
        sCryptPass = GetECSCryptSetting("ECS", psSection, "PASSWORD")
        If sCryptPass <> vbNullString Then
            If Not poForm Is Nothing Then
                'PUt input box top left of form
                sRet = InputBox("Please enter old password.", "OLD PASSWORD", , poForm.left, poForm.top)
                If sRet = vbNullString Then
                    'they clicked on cancel
                    GoTo CLEANUP
                End If
            Else
                'Put input box default windows pos
                sRet = InputBox("Please enter old password.", "OLD PASSWORD")
                If sRet = vbNullString Then
                    'they clicked on cancel
                    GoTo CLEANUP
                End If
            End If
            'check sret against the old password
            If StrComp(sCryptPass, sRet, vbBinaryCompare) <> 0 Then
                MsgBox "The password you entered does not match.", vbOKOnly + vbExclamation, "INCORRECT PASSWORD"
                GoTo CLEANUP
            Else
                'Save the Old Pass word
                SaveECSCryptSetting "ECS", psSection, "OLD_PASSWORD", sCryptPass
                'Need to reset password on server when connecting to server
                SaveSetting "ECS", psSection, "RESET_PASSWORD", True
                GoTo ENTER_NEWPASS
            End If
            
        Else
            'if there is no password saved yet then just ask for
            'Password
ENTER_NEWPASS:
            sRet = InputBox("Please enter a new password.", "ENTER NEW PASSWORD", , poForm.left, poForm.top)
            If sRet = vbNullString Then
                'they clicked on cancel
                'Need to undo reset password on server when connecting to server
                SaveSetting "ECS", psSection, "RESET_PASSWORD", False
                GoTo CLEANUP
            Else
                sCryptPass = sRet
                'Ask them to double check the password they just entered
                sRet = InputBox("Please enter the same password again.", "ENTER PASSWORD AGAIN", , poForm.left, poForm.top)
                If sRet = vbNullString Then
                    'they clicked on cancel
                    'Need to undo reset password on server when connecting to server
                    SaveSetting "ECS", psSection, "RESET_PASSWORD", False
                     GoTo CLEANUP
                Else
                    If StrComp(sCryptPass, sRet, vbBinaryCompare) <> 0 Then
                        MsgBox "The password you entered does not match.", vbOKOnly + vbExclamation, "INCORRECT PASSWORD"
                        'Need to undo reset password on server when connecting to server
                        SaveSetting "ECS", psSection, "RESET_PASSWORD", False
                        GoTo CLEANUP
                    End If
                End If
            End If
            
            'Disaplay password
            pvPass.Text = sCryptPass
        End If
    Else
        sCryptPass = CStr(pvPass)
    End If
    'set the user Name
    sCryptUserName = psUserName
    
    If sCryptPass <> vbNullString Then
        SaveECSCryptSetting "ECS", psSection, "PASSWORD", sCryptPass
    End If
    If sCryptUserName <> vbNullString Then
        SaveECSCryptSetting "ECS", psSection, "USER_NAME", sCryptUserName
    End If
    
CLEANUP:

    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Sub EnterPass" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Sub

Public Function GetECSCryptSetting(psAPP As String, psSection As String, psKey As String, _
                                   Optional pvDefault As Variant = vbNullString) As Variant
    On Error GoTo EH
    Dim sCryptSetting As String
    Dim lErrNum As Long
    Dim sErrDesc As String
           
    sCryptSetting = GetSetting(psAPP, psSection, psKey, vbNullString)
    
    If sCryptSetting <> vbNullString Then
        GetECSCryptSetting = CStr(goUtil.Decode(sCryptSetting))
    Else
        GetECSCryptSetting = pvDefault
    End If
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Function GetECSCryptSetting" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Function

Public Sub SaveECSCryptSetting(psAPP As String, psSection As String, psKey As String, psSetting As String)
    On Error GoTo EH
    Dim sCryptSetting As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    sCryptSetting = psSetting
    
    
    sCryptSetting = goUtil.Encode(sCryptSetting)
    
    SaveSetting psAPP, psSection, psKey, sCryptSetting
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Sub SaveECSCryptSetting" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Sub

Public Function GetWinOSVersion() As String
    On Error GoTo EH
    Dim WinOS As OSVERSIONINFO
    Dim sVSInfo As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    WinOS.dwOSVersionInfoSize = Len(WinOS)
    GetVersionEx WinOS
    
    sVSInfo = "Build Number: " & WinOS.dwBuildNumber & vbCrLf
    sVSInfo = sVSInfo & "Major Version: " & WinOS.dwMajorVersion & vbCrLf
    sVSInfo = sVSInfo & "Minor Verison: " & WinOS.dwMinorVersion & vbCrLf
    sVSInfo = sVSInfo & "OS VS Info Size: " & WinOS.dwOSVersionInfoSize & vbCrLf
    sVSInfo = sVSInfo & "Platform ID: " & WinOS.dwPlatformId & vbCrLf
    sVSInfo = sVSInfo & "CSD Version: " & WinOS.szCSDVersion & vbCrLf
    
    GetWinOSVersion = sVSInfo
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Function GetWinOSVersion" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Function

Public Function SaveDefaultPrinterSettings(psAppEXEName As String, Optional psDefaultPrinterName As String) As Boolean
    On Error GoTo EH
    Dim prn As Printer
    Dim sDefaultPrinter As String
    Dim nRet As Integer
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    If psDefaultPrinterName = vbNullString Then
        'Get Default Printer Name
        sDefaultPrinter = Space(255)
        nRet = GetProfileString("Windows", ByVal "device", "", sDefaultPrinter, Len(sDefaultPrinter))
        'Trim it
        If nRet Then
            sDefaultPrinter = left(sDefaultPrinter, InStr(sDefaultPrinter, ",") - 1)
        End If
    Else
        sDefaultPrinter = psDefaultPrinterName
    End If
    'Loop until we find the Default prn object
    For Each prn In Printers
        If InStr(1, sDefaultPrinter, prn.DeviceName, vbTextCompare) = 1 Then
            Exit For
        End If
    Next prn
        
    If CBool(GetSetting(psAppEXEName, "PRINTER", "INIT_WITH_WIN_DEFAULT_PRINTER", True)) Then
        'Find the selected printer against printer list.
        'Do this because it is possible that the printer list
        'could have changed from the last time they loaded the form
        If Not prn Is Nothing Then
            'BGS 8.26.2002 when user Changes the Default printer.
            If prn.DeviceName <> GetSetting(psAppEXEName, "PRINTER", "PRINTER_NAME", vbNullString) Or _
                prn.DriverName <> GetSetting(psAppEXEName, "PRINTER", "PRINTER_DRIVER", vbNullString) Or _
                prn.Port <> GetSetting(psAppEXEName, "PRINTER", "PRINTER_PORT", vbNullString) Then
                SaveSetting psAppEXEName, "PRINTER", "PRINTER_NAME", prn.DeviceName
                SaveSetting psAppEXEName, "PRINTER", "PRINTER_PORT", prn.Port
                SaveSetting psAppEXEName, "PRINTER", "PRINTER_DRIVER", prn.DriverName
            End If
        Else
            SaveSetting psAppEXEName, "PRINTER", "PRINTER_NAME", vbNullString
            SaveSetting psAppEXEName, "PRINTER", "PRINTER_PORT", vbNullString
            SaveSetting psAppEXEName, "PRINTER", "PRINTER_DRIVER", vbNullString
        End If
    End If
    
    SaveDefaultPrinterSettings = True
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Function SaveDefaultPrinterSettings" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Function

Public Function GetDefaultPrinterSettings(psAppEXEName As String) As V2ECKeyBoard.ECPRINTER
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    GetDefaultPrinterSettings.PRINTER_NAME = GetSetting(psAppEXEName, "PRINTER", "PRINTER_NAME", vbNullString)
    GetDefaultPrinterSettings.PRINTER_PORT = GetSetting(psAppEXEName, "PRINTER", "PRINTER_PORT", vbNullString)
    GetDefaultPrinterSettings.PRINTER_DRIVER = GetSetting(psAppEXEName, "PRINTER", "PRINTER_DRIVER", vbNullString)
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Function GetDefaultPrinterSettings" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Function

Public Sub LoadSP(psInstallDir As String, poSP As V2ECKeyBoard.clsSP)
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    If poSP Is Nothing Then
        Set poSP = New V2ECKeyBoard.clsSP
        poSP.DictionaryPath = psInstallDir & "\Templates\Dictionary\ENGLISH.dic"
        poSP.CusDictionaryPath = psInstallDir & "\Templates\Dictionary\CUSTOM.dic"
        poSP.LoadDictionaries
    End If
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Sub LoadSP" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Sub

Public Function GetPath(psAppEXEName As String, psName As String, psMess As String, psFileMess As String, psDeFaultPath As String, plHwnd As Long, _
                         Optional psFilter As String = vbNullString, _
                         Optional psSelFile As String, _
                         Optional plFlags As Long, _
                         Optional pbCenterForm As Boolean = True, _
                         Optional pbShowOpen As Boolean = True) As String
    On Error GoTo EH
    Dim sOpen As SelectedFile
    Dim sDir As String
    Dim sMyFilter As String
    Dim sInitDir As String
    Dim sNewCatName As String
    Dim MyFileDialog As OPENFILENAME
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    FileDialog = MyFileDialog
    sMyFilter = psFilter
    FileDialog.sFilter = sMyFilter
    
    ' See Standard CommonDialog Flags for all options
    If plFlags > 0 Then
        FileDialog.flags = plFlags
    Else
        FileDialog.flags = OFN_HIDEREADONLY Or OFN_NOVALIDATE
    End If
    FileDialog.sDlgTitle = psMess
    FileDialog.sFile = psFileMess
    If FileExists(psDeFaultPath, False) Or FileExists(psDeFaultPath, True) Then
        sInitDir = psDeFaultPath
    Else
        sInitDir = GetSetting(psAppEXEName, "Dir", psName, "Error")
    End If
    
    If sInitDir <> "Error" And sInitDir <> vbNullString Then
        FileDialog.sInitDir = sInitDir
    Else
        FileDialog.sInitDir = "C:\"
    End If
    
    If pbShowOpen Then
        sOpen = ShowOpen(plHwnd, pbCenterForm)
    Else
        sOpen = ShowSave(plHwnd, pbCenterForm)
    End If
    If Err.Number <> 32755 And sOpen.bCanceled = False Then
        sDir = left(sOpen.sLastDirectory, InStrRev(sOpen.sLastDirectory, "\"))
        If sDir = vbNullString And sOpen.nFilesSelected = 1 Then
            sDir = left(sOpen.sFiles(1), InStrRev(sOpen.sFiles(1), "\"))
        End If
        'Set the selected file
        psSelFile = Replace(sOpen.sFiles(1), sDir, vbNullString)
        SaveSetting psAppEXEName, "Dir", psName, sDir
        GetPath = sDir
    Else
        GetPath = psDeFaultPath
    End If
    
   Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Function GetPath" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Function


Public Function GetSavePath(psAppEXEName As String, psName As String, psMess As String, psFileMess As String, psDeFaultPath As String, plHwnd As Long, _
                         Optional psFilter As String = vbNullString, _
                         Optional psSelFile As String, _
                         Optional plFlags As Long, _
                         Optional pbCenterForm As Boolean = True, _
                         Optional pbShowOpen As Boolean = True) As String
    Dim sSave As SelectedFile
    Dim sDir As String
    Dim sMyFilter As String
    Dim sSaveDir As String
    Dim sFilePath As String
    Dim bUseFilePath As Boolean
    Dim sExt As String
    Dim sSelFile As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    On Error GoTo EH
    
    sMyFilter = psFilter
    FileDialog.sFilter = sMyFilter
    
    ' See Standard CommonDialog Flags for all options
    FileDialog.flags = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_OVERWRITEPROMPT
    FileDialog.sDlgTitle = psMess
    sSaveDir = GetSetting(psAppEXEName, "Dir", "Save" & psName, "Error")
    If sSaveDir <> "Error" Then
        FileDialog.sInitDir = sSaveDir
    Else
        FileDialog.sInitDir = "C:\"
    End If
    
    sSave = ShowSave(plHwnd)
    If Err.Number <> 32755 And sSave.bCanceled = False Then
        sDir = sSave.sLastDirectory
        If goUtil.utFileExists(sDir, True) Then
            SaveSetting psAppEXEName, "Dir", "Save" & psName, sDir
            If Right(sDir, 1) <> "\" Then
                sDir = sDir & "\"
            End If
        Else
            If InStr(1, sSave.sFiles(1), "\", vbTextCompare) > 0 Then
                bUseFilePath = True
                sFilePath = left(sSave.sFiles(1), InStrRev(sSave.sFiles(1), "\"))
                SaveSetting psAppEXEName, "Dir", "Save" & psName, sFilePath
                sDir = sFilePath
            End If
        End If
        'Set the selected file
        sSelFile = sSave.sFiles(1)
        'Check for an ext added by the use rot get rif of it
        If InStr(1, sSelFile, ".", vbBinaryCompare) > 0 Then
            sExt = Mid(sSelFile, InStrRev(sSave.sFiles(1), "."))
            If Len(sExt) <> 4 Then
                sExt = vbNullString
            End If
        End If
        sSelFile = Replace(sSelFile, sDir, vbNullString)
        sSelFile = Replace(sSelFile, sExt, vbNullString)
        psSelFile = sSelFile
        GetSavePath = sDir
        
    End If
Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Function GetSavePath" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Function

Public Sub LoadCarriers(pcboList As ComboBox)
    On Error GoTo EH
    Dim sCar As String
    Dim sBaseCarDllPath As String
    Dim vCar As Variant
    Dim colCar As Collection
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    If FileExists("C:\WINDOWS\SYSTEM\ECS\DLL", True) Then
        sBaseCarDllPath = "C:\WINDOWS\SYSTEM\ECS\DLL"
    ElseIf FileExists("C:\WINDOWS\SYSTEM32\ECS\DLL", True) Then
        sBaseCarDllPath = "C:\WINDOWS\SYSTEM32\ECS\DLL"
    ElseIf FileExists("C:\WINNT\SYSTEM32\ECS\DLL", True) Then
        sBaseCarDllPath = "C:\WINNT\SYSTEM32\ECS\DLL"
    End If
    
    sCar = Dir(sBaseCarDllPath & "\" & goUtil.gsCarPrefix & "*.dll")
    Do Until sCar = vbNullString
        If colCar Is Nothing Then
            Set colCar = New Collection
        End If
        colCar.Add sCar, sCar
        sCar = Dir
    Loop
    
    pcboList.Clear
    If Not colCar Is Nothing Then
        For Each vCar In colCar
            sCar = vCar
            sCar = Replace(sCar, goUtil.gsCarPrefix, vbNullString)
            sCar = Replace(sCar, ".dll", vbNullString)
            pcboList.AddItem sCar
        Next
    End If
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Sub LoadCarriers" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Sub

Public Sub LoadTOL(pcboList As ComboBox)
    On Error GoTo EH
    Dim oXact As New clsXact
    Dim vTOL As Variant
    Dim lCount As Long
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    'Use the Xactimate class to steal the avail Type Of Loss from it.
    'If Xactimate is not installed or there is nothing set up yet
    'the adjuster will still be able to Type in TOL as usual.
    vTOL = oXact.TOL
    oXact.CLEANUP
    Set oXact = Nothing
    
    pcboList.Clear
    If IsArray(vTOL) Then
        For lCount = LBound(vTOL) To UBound(vTOL)
            If Trim(CStr(vTOL(lCount))) <> vbNullString Then
                pcboList.AddItem vTOL(lCount)
            End If
        Next
    End If
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Sub LoadTOL" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Sub

Public Sub LoadStates(pcboList As ComboBox)
    On Error GoTo EH
    Dim oXact As New clsXact
    Dim vStates As Variant
    Dim lCount As Long
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    'Use the Xactimate class to steal the avail Type Of Loss from it.
    'If Xactimate is not installed or there is nothing set up yet
    'the adjuster will still be able to Type in TOL as usual.
    vStates = oXact.States
    oXact.CLEANUP
    Set oXact = Nothing
    
    pcboList.Clear
    If IsArray(vStates) Then
        For lCount = LBound(vStates) To UBound(vStates)
            If Trim(CStr(vStates(lCount))) <> vbNullString Then
                pcboList.AddItem vStates(lCount)
            End If
        Next
    End If
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Sub LoadStates" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Sub

Public Function FetchFLSFile(psRawPathFileName As String, psCopyToPath As String, _
                             psLifeSpanServerPath As String, Optional plLifeSpanDays = 1, _
                             Optional pbDelRawFile As Boolean) As String
    'This Function purpose is to Copy File from a protected directory into
    'an unprotected dir for the purpose of Downloading Via the web.  The File will
    'An entry in the FLS.dat file will be made giving that File a Life Span whatever that may be
    'A separate process (DelFLSFiles) will monitor FLS.dat looking for Files that have Expired their
    'life span.  DelFLSFiles Can be processed from an EXE or a Service that calls DelFLSFiles
    On Error GoTo EH
    Dim sMess As String
    Dim sTemp As String
    Dim sFilename As String
    Dim vFLSData As Variant
    Dim vFLSItem As Variant
    Dim lCount As Long
    Dim bItemUpdated As Boolean
    Dim lPermCount As Long
    Dim iFile As Integer
    Dim lPos As Long
    Dim lEndPos As Long
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    'Check to be sure using UncPaths only !
    If left(psRawPathFileName, 2) <> "\\" Then
        sMess = "'" & psRawPathFileName & "'" & vbCrLf
        sMess = sMess & "Is not a valid UNC path \\{ServerName}\{ShareName}\{Path}"
        Err.Raise -999, , sMess
    End If
    If left(psCopyToPath, 2) <> "\\" Then
        sMess = "'" & psCopyToPath & "'" & vbCrLf
        sMess = sMess & "Is not a valid UNC path \\{ServerName}\{ShareName}\{Path}"
        Err.Raise -999, , sMess
    End If
    If left(psLifeSpanServerPath, 2) <> "\\" Then
        sMess = "'" & psLifeSpanServerPath & "'" & vbCrLf
        sMess = sMess & "Is not a valid UNC path \\{ServerName}\{ShareName}\{Path}"
        Err.Raise -999, , sMess
    End If
    'Validate that Raw File Exists
    If Not FileExists(psRawPathFileName) Then
        sMess = "'" & psRawPathFileName & "'" & vbCrLf
        sMess = sMess & "Raw File does not exist!"
        Err.Raise -999, , sMess
    End If
    'Validate that the desired Directory exists
    If Not FileExists(psCopyToPath, True) Then
        sMess = "'" & psCopyToPath & "'" & vbCrLf
        sMess = sMess & "Is not a valid path!"
        Err.Raise -999, , sMess
    End If
    'Validate that the LifeSpanServerPath Exists
    If Not FileExists(psLifeSpanServerPath, True) Then
        sMess = "'" & psLifeSpanServerPath & "'" & vbCrLf
        sMess = sMess & "Is not a valid Path!"
        Err.Raise -999, , sMess
    End If
    'Validate that the Data file Exists
    If Not FileExists(psLifeSpanServerPath & "\FLS.dat") Then
        'Make it Here with nothing in it
        SaveFileData psLifeSpanServerPath & "\FLS.dat", vbNullString
    End If
    
    'If we get to here we are ready to Copy the file over and Update the
    'FLS.dat(FileLifeSpan.dat) (This contains list of files, The date time created and Number of Days
    'they have from the Date Time Created before they need to be Deleted.
    
    '1. First Populate the FLS Array (Items Stored with sub Items Delim by ",", Items Delim By vbCrLf
    On Error Resume Next
GET_DATA:
    vFLSData = Split(GetFileData(psLifeSpanServerPath & "\FLS.dat", True, iFile), vbCrLf)
    'If permision error keep trying for 5 seconds
    If Err.Number > 0 Then
        If lPermCount > 10 Then
            On Error GoTo EH
            GoTo GET_DATA
        Else
            Err.Clear
            Sleep 500
            lPermCount = lPermCount + 1
            GoTo GET_DATA
        End If
    Else
        On Error GoTo EH
    End If
    
    '2. Copy over the File
    'Next check to see if the CopyToPath already Exists.
    'if it does then we will Copy Over it.
    lPos = InStrRev(psRawPathFileName, "\") + 1
    lEndPos = InStrRev(psRawPathFileName, ".")
    sTemp = Mid(psRawPathFileName, lEndPos)
    sFilename = Mid(psRawPathFileName, lPos, lEndPos - lPos)
    sFilename = sFilename & "_FLS_" & Format(Now, "MMDDYY") & "_FLS" & sTemp
    
    FetchFLSFile = sFilename
    
    If FileExists(psCopyToPath & "\" & sFilename) Then
        SetAttr psCopyToPath & "\" & sFilename, vbNormal
        FileCopy psRawPathFileName, psCopyToPath & "\" & sFilename
        If pbDelRawFile Then
            SetAttr psRawPathFileName, vbNormal
            Kill psRawPathFileName
        End If
        'Need to Check the Array of Files To See if we need to Update the Item
        If IsArray(vFLSData) Then
            For lCount = 0 To UBound(vFLSData, 1)
                sTemp = CStr(vFLSData(lCount))
                sTemp = left(sTemp, InStr(1, sTemp, ",") - 1)
                If StrComp(RTrim(psCopyToPath), RTrim(sTemp), vbTextCompare) = 0 Then
                    vFLSItem = Split(vFLSData(lCount), ",")
                    vFLSItem(FLSItem.FLS1LifeSpanDays) = plLifeSpanDays & String(10 - Len(plLifeSpanDays), Chr(32)) 'Put some space
                    vFLSData(lCount) = Join(vFLSItem, ",")
                    bItemUpdated = True
                    Exit For 'Bail Here since we found it
                End If
            Next
        End If
    End If
    
    '3. Update Item If it did not already Exist
    If Not bItemUpdated Then
        If FileExists(psRawPathFileName) Then
            FileCopy psRawPathFileName, psCopyToPath & "\" & sFilename
            If pbDelRawFile Then
                SetAttr psRawPathFileName, vbNormal
                Kill psRawPathFileName
            End If
        End If
        'Need to Check the Array of Files To See if we need to Update the Item
        If IsArray(vFLSData) Then
            For lCount = 0 To UBound(vFLSData, 1)
                sTemp = CStr(vFLSData(lCount))
                sTemp = left(sTemp, InStr(1, sTemp, ",") - 1)
                If StrComp(RTrim(psCopyToPath), RTrim(sTemp), vbTextCompare) = 0 Then
                    vFLSItem = Split(vFLSData(lCount), ",")
                    vFLSItem(FLSItem.FLS1LifeSpanDays) = plLifeSpanDays & String(10 - Len(plLifeSpanDays), Chr(32)) 'Put some space
                    vFLSData(lCount) = Join(vFLSItem, ",")
                    bItemUpdated = True
                    Exit For 'Bail Here since we found it
                End If
            Next
        End If
        If Not bItemUpdated Then
            ReDim vFLSItem(0 To 1)
            vFLSItem(FLSItem.FLS0FileDir) = psCopyToPath & String(255 - Len(psCopyToPath), Chr(32))
            vFLSItem(FLSItem.FLS1LifeSpanDays) = plLifeSpanDays & String(10 - Len(plLifeSpanDays), Chr(32))
            ReDim Preserve vFLSData(0 To UBound(vFLSData, 1) + 1)
            vFLSData(UBound(vFLSData, 1)) = Join(vFLSItem, ",")
        End If
    End If
    
    '4. Finally Save the FLSData file
    SaveFileData psLifeSpanServerPath & "\FLS.dat", Join(vFLSData, vbCrLf), , , iFile
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Close iFile
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Function FetchFLSFile" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Function

Public Sub DelFLSFiles(psLifeSpanServerPath As String)
    On Error GoTo EH
    Dim sFilePath As String
    Dim lLifeSpanDays As Long
    Dim sMess As String
    Dim vFLSData As Variant
    Dim vNewFLSData As Variant
    Dim vFLSItem As Variant
    Dim lCount As Long
    Dim lNewCount As Long
    Dim iFile As Integer
    Dim bAllItemsDeleted As Boolean
    Dim bUpdateFLSdat As Boolean
    Dim lPermCount As Long
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    'Validate that the LifeSpanServerPath Exists
    If Not FileExists(psLifeSpanServerPath, True) Then
        sMess = "'" & psLifeSpanServerPath & "'" & vbCrLf
        sMess = sMess & "Is not a valid Path!"
        Err.Raise -999, , sMess
    End If
    'Validate that the Data file Exists
    If Not FileExists(psLifeSpanServerPath & "\FLS.dat") Then
        'Make it Here with nothing in it
        SaveFileData psLifeSpanServerPath & "\FLS.dat", vbNullString
    End If
    
    '1. Need to retrieve the FLSData
    On Error Resume Next
    vFLSData = Split(GetFileData(psLifeSpanServerPath & "\FLS.dat", True, iFile), vbCrLf)
    'If permision error keep trying for 5 seconds
    If Err.Number > 0 Then
        'Bail if we get Permission error here
        'This is ok becuase this process will be called again
        Err.Clear
        Exit Sub
    Else
        On Error GoTo EH
    End If
    
    '2. Check each FLS Item
    If IsArray(vFLSData) Then
        For lCount = 0 To UBound(vFLSData, 1)
            vFLSItem = Split(vFLSData(lCount), ",")
            lLifeSpanDays = CLng(RTrim(vFLSItem(FLSItem.FLS1LifeSpanDays)))
            
            bAllItemsDeleted = DeleteAllFLSFiles(CStr(RTrim(vFLSItem(FLSItem.FLS0FileDir))), lLifeSpanDays)
            
            'If there are still some files with life span left need to add this to the NewArray
            If Not bAllItemsDeleted Then
                bUpdateFLSdat = True
                lNewCount = lNewCount + 1
                'This is the only array in FLS that starts with 1 as Lbound element
                If IsArray(vNewFLSData) Then
                    ReDim Preserve vNewFLSData(1 To lNewCount)
                Else
                    ReDim vNewFLSData(1 To lNewCount)
                End If
                vNewFLSData(lNewCount) = Join(vFLSItem, ",")
            End If
        Next
    End If
    
    '3. If we actually deleted a file, need to Save FLS.dat file
    If bUpdateFLSdat Then
        Close iFile
        On Error Resume Next
GET_DATA:
        SetAttr psLifeSpanServerPath & "\FLS.dat", vbNormal
        Kill psLifeSpanServerPath & "\FLS.dat"
        'If permision error keep trying for 5 seconds
        If Err.Number > 0 Then
            If lPermCount > 10 Then
                On Error GoTo EH
                GoTo GET_DATA
            Else
                Err.Clear
                Sleep 500
                lPermCount = lPermCount + 1
                GoTo GET_DATA
            End If
        Else
            On Error GoTo EH
        End If
        
        If IsArray(vNewFLSData) Then
            SaveFileData psLifeSpanServerPath & "\FLS.dat", Join(vNewFLSData, vbCrLf)
        End If
    Else
        Close iFile
    End If
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Close iFile
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Sub DelFLSFiles" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Sub

Public Sub SuffixLabels(plblArray As Object, Optional plLen As Long = 25)
    On Error Resume Next
    Dim lCount As Long

    For lCount = plblArray.LBound To plblArray.UBound
        plblArray(lCount).Caption = plblArray(lCount).Caption & String(plLen - Len(plblArray(lCount).Caption), ".")
'        Debug.Print plblArray(lCount).Caption
        If Err.Number > 0 Then
            Err.Clear
        End If
    Next

End Sub

Public Function DeleteAllFLSFiles(psFLSDir As String, plFLSDays As Long) As Boolean
    On Error GoTo EH
    Dim dFLSDate As Date
    Dim dNow As Date
    Dim colFLS As Collection
    Dim sFLS As String
    Dim vFLS As Variant
    Dim sTemp As String
    Dim bDelAll As Boolean
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    'sFileName & "_FLS_" & Format(Now, "DDYYMM") & "_FLS" & sTemp
    'Load ".TXT" files
    sFLS = Dir(psFLSDir & "\" & "*_FLS_*_FLS.*")
    
    If sFLS <> vbNullString Then
        If colFLS Is Nothing Then
            Set colFLS = New Collection
        End If
        'Load ".TXT" files
        sFLS = Dir(psFLSDir & "\" & "*_FLS_*_FLS.*")
    
        Do
            colFLS.Add psFLSDir & "\" & sFLS
            sFLS = Dir
        Loop Until sFLS = vbNullString
    End If
    
    sFLS = vbNullString
    
    If Not colFLS Is Nothing Then
        bDelAll = True
        For Each vFLS In colFLS
            sFLS = vFLS
            sTemp = Mid(sFLS, InStr(1, sFLS, "_FLS_", vbTextCompare) + 5, 6)
            sTemp = left(sTemp, 2) & "/" & Mid(sTemp, 3, 2) & "/" & Right(sTemp, 2)
            dFLSDate = CDate(sTemp)
            dNow = Format(Now(), "MM/DD/YY")
            If DateDiff("d", dFLSDate, dNow) >= plFLSDays Then
                On Error Resume Next
                SetAttr sFLS, vbNormal
                Kill sFLS
                If Err.Number > 0 Then
                    Err.Clear
                    bDelAll = False
                    On Error GoTo EH
                End If
            Else
                bDelAll = False
            End If
        Next
    End If
    
    DeleteAllFLSFiles = bDelAll
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Private Function DeleteAllFLSFiles" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Function

Public Sub CleanFileFolderName(pvText As Variant, Optional pbFilePath As Boolean)
    On Error GoTo EH
    Dim lPos As Long
    Dim sTemp As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    'Issue 180 180 Error #76 Error in Function_ CreateCat; Form SetupNewCat
    'Clean out any Special chars \/:*?"<>|
    If IsObject(pvText) Then
        lPos = pvText.SelStart
        sTemp = pvText.Text
    Else
        sTemp = CStr(pvText)
    End If
    If Not pbFilePath Then
        sTemp = Replace(sTemp, "\", vbNullString, , , vbBinaryCompare)
    End If
    sTemp = Replace(sTemp, "/", vbNullString, , , vbBinaryCompare)
    If Not pbFilePath Then
        sTemp = Replace(sTemp, ":", vbNullString, , , vbBinaryCompare)
    End If
    sTemp = Replace(sTemp, "*", vbNullString, , , vbBinaryCompare)
    sTemp = Replace(sTemp, "?", vbNullString, , , vbBinaryCompare)
    sTemp = Replace(sTemp, """", vbNullString, , , vbBinaryCompare)
    sTemp = Replace(sTemp, "<", vbNullString, , , vbBinaryCompare)
    sTemp = Replace(sTemp, ">", vbNullString, , , vbBinaryCompare)
    sTemp = Replace(sTemp, "|", vbNullString, , , vbBinaryCompare)
    
    If IsObject(pvText) Then
        pvText.Text = sTemp
        pvText.SelStart = lPos
    Else
        pvText = sTemp
    End If
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Sub CleanFileName" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Sub

Public Function CreateCat(psAppEXEName As String, psInstallDir As String, _
                          pcolCatPrefs As Collection, _
                          Optional pbHideError As Boolean) As Boolean
    On Error GoTo EH
    Dim sDestDir As String
    Dim sSourceDir As String
    Dim sCatPref As String
    Dim sCatName As String
    Dim sCarrier As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    'Get some stuff from Cat Pref Collecion
    If Not pcolCatPrefs Is Nothing Then
        sCatName = Replace(pcolCatPrefs.Item("CAT_NAME"), "CAT_NAME=", vbNullString)
        sCarrier = Replace(pcolCatPrefs.Item("CAT_CARRIER"), "CAT_CARRIER=", vbNullString)
    Else
        Exit Function
    End If
    
    'Build Dest directory
    sDestDir = psInstallDir & "\Cats\" & sCarrier & "\" & sCatName
    
    'Build Template Source Directory
    sSourceDir = psInstallDir & "\Templates\" & "Carriers\" & sCarrier
    
    If Not FileExists(sSourceDir, True) Then
        If Not pbHideError Then
            MsgBox "Invalid Directory " & sSourceDir, vbExclamation + vbOKOnly, "Can't Create CAT"
        End If
        Exit Function
    End If
    
    'Don't create if already there
    If FileExists(sDestDir, True) Then
        If Not pbHideError Then
            MsgBox "CAT already exists!", vbExclamation + vbOKOnly, "Can't Create CAT"
            Exit Function
        End If
    End If
      
    'Make the Cat
    
    '1. first check to be sure the Cat Folder is there and The Carrier Folder
    If Not FileExists(psInstallDir & "\Cats", True) Then
        MakeDir psInstallDir & "\Cats"
    End If
    If Not FileExists(psInstallDir & "\Cats\" & sCarrier, True) Then
        MakeDir psInstallDir & "\Cats\" & sCarrier
    End If
    
    '2. Make the New CAT Dir
    MakeDir sDestDir
    
    '3. Copy from Template Source into Dest
    CopyDir sSourceDir, sDestDir
    
    'Create the CAT PRef ini from the Param COllecttion
    sCatPref = CreateCatPref(pcolCatPrefs)
    SaveFileData sDestDir & "\CatPref.ini", sCatPref
    
    'Save the CAT Structure in the registry
    'This will be used to only allow Cats created with Easy Claim
    'to be Populated in the Tree. It Will also be used to indicate if
    'a cat has been sent to the recycle bin.
    SaveSetting psAppEXEName, "CAT_STRUCTURE\" & sCarrier, sCarrier, goUtil.Encode(sCarrier)
    SaveSetting psAppEXEName, "CAT_STRUCTURE\" & sCarrier & "\" & sCatName, sCatName, goUtil.Encode(sCatName)
    
    If Not pbHideError Then
        MsgBox "CAT: " & sCatName & " Setup complete!", vbInformation + vbOKOnly, "Cat Setup"
    End If
    
    CreateCat = True
  
  Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Function CreateCat" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Function

Public Function CreateCatPref(pcolCatPrefs As Collection) As String
    On Error GoTo EH
    Dim vPref As Variant
    Dim sPref As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    'This will create the CAT PREF INI FILE
    If Not pcolCatPrefs Is Nothing Then
        For Each vPref In pcolCatPrefs
            sPref = vPref
            CreateCatPref = CreateCatPref & sPref & vbCrLf
        Next
    End If
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Private Function CreatCatPref" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Function

Public Sub SaveLic(plDays As Long)
    On Error GoTo EH
    Dim dtLicDate As Date
    Dim dtTodayDate As Date
    Dim sTemp As String
    Dim sMess As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    'First check the Date. If it is the same as Today then do not subtract from
    'the Lic number
    sTemp = GetECSCryptSetting("ECS", "WEB_SECURITY", "DATE")
    
    If sTemp = vbNullString Then
        plDays = plDays - 1 'Chink one day off
        sTemp = Format(Now(), "MM/DD/YY")
        SaveECSCryptSetting "ECS", "WEB_SECURITY", "DATE", sTemp
    End If
    
    dtLicDate = CDate(sTemp)
    dtTodayDate = Format(Now(), "MM/DD/YY")
    
    If dtLicDate <> dtTodayDate Then
        plDays = plDays - 1 'Chink one day off
        SaveECSCryptSetting "ECS", "WEB_SECURITY", "DATE", CStr(dtTodayDate)
    End If
    
    SaveECSCryptSetting "ECS", "WEB_SECURITY", "LIC", CStr(plDays)
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Sub SaveLic" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Sub

'This function will look for the Application window by Title caption (First partial match)
'If it does not find it then it will loop for specified time
Public Function WindowFound(psName As String, Optional plSleepSeconds As Long = 1) As Boolean
    On Error GoTo EH
    Dim lSleep As Long
    Dim lRet As Long
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    'BGS 1.28.2002 Sleep and look for the window
    For lSleep = 1 To plSleepSeconds * 10
        DoEvents
        Sleep 100
        lRet = AppActivatePartial(psName)
        If lRet Then
            WindowFound = True
            Exit Function
        End If
    Next
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Private Function WindowFound" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Function

Public Function GetMyComputerName() As String
    On Error GoTo EH
    Dim lsize As Long
    Dim lRet As Long
    Dim lErrNum As Long
    Dim sErrDesc As String
    lsize = 255
    GetMyComputerName = Space(lsize)
    lRet = GetComputerName(GetMyComputerName, lsize)
    GetMyComputerName = left(GetMyComputerName, lsize)
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Function GetMyComputerName" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Function

Public Function GetMyUserName() As String
    On Error GoTo EH
    Dim lsize As Long
    Dim lRet As Long
    Dim lErrNum As Long
    Dim sErrDesc As String
    lsize = 255
    GetMyUserName = Space(lsize)
    lRet = GetUserName(GetMyUserName, lsize)
    GetMyUserName = left(GetMyUserName, lsize)
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Function GetMyUserName" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Function

Public Function CLEANUPModUtil() As Boolean
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    Set mFSO = Nothing
    
    CLEANUPModUtil = True
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    CLEANUPModUtil = False
    Err.Raise lErrNum, , sErrDesc & vbCrLf & msClassName & vbCrLf & "Public Function CLEANUP"
End Function
