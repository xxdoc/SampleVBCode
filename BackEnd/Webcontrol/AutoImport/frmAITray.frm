VERSION 5.00
Object = "{307C5043-76B3-11CE-BF00-0080AD0EF894}#1.0#0"; "MsgHoo32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAITray 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Auto Import (Web Control)"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7410
   Icon            =   "frmAITray.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   7410
   Begin VB.Timer Timer_Status 
      Interval        =   500
      Left            =   600
      Top             =   600
   End
   Begin VB.Timer Timer_VerifyDSNConn 
      Left            =   120
      Top             =   600
   End
   Begin MsghookLib.Msghook Msghook 
      Left            =   720
      Top             =   120
      _Version        =   65536
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   120
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAITray.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAITray.frx":0894
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAITray.frx":0CE6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar ProgBarLoss 
      Height          =   375
      Left            =   45
      TabIndex        =   0
      ToolTipText     =   "Right-Click for Options"
      Top             =   3840
      Width           =   7080
      _ExtentX        =   12488
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.TextBox txtMess 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4215
      Left            =   40
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   0
      Width           =   7335
   End
   Begin VB.Menu mPopUp 
      Caption         =   "Pop"
      Visible         =   0   'False
      Begin VB.Menu mPop 
         Caption         =   "&Show"
         Index           =   0
      End
      Begin VB.Menu mPop 
         Caption         =   "&Hide"
         Index           =   1
      End
      Begin VB.Menu mPop 
         Caption         =   "&Disable"
         Index           =   2
      End
      Begin VB.Menu mPop 
         Caption         =   "&Restart"
         Index           =   3
      End
      Begin VB.Menu mPop 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mPop 
         Caption         =   "&Interval"
         Index           =   5
      End
      Begin VB.Menu mPop 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mPop 
         Caption         =   "View &WebControl"
         Index           =   7
      End
      Begin VB.Menu mPop 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mPop 
         Caption         =   " &EC Update Batches (Force)"
         Index           =   9
      End
      Begin VB.Menu mPop 
         Caption         =   "-"
         Index           =   10
         Visible         =   0   'False
      End
      Begin VB.Menu mPop 
         Caption         =   "&Fail Logon VS <"
         Index           =   11
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmAITray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Enum PicList
   Idle = 1
   Busy
   Disabled
End Enum

Public Enum MenuList
    Show = 0
    Hide
    Disable
    Restart
    BarInt
    Interval
    BarViewWebCOntrol
    ViewWebControl
    BarECUpdateBatches
    ECUpdateBatches
    BarFailLogonVS
    FailLogonVS
End Enum

' User defined constant values
Private Const cbNotify As Long = &H4000
Private Const uID As Long = 61860

' Member variables
Private m_NID As NOTIFYICONDATA
Private m_TaskbarCreated As Long

'BGS 10.31.2001 Use this to process the finished claims
Private msLastTime As String
Private WithEvents moUL As V2ECKeyBoard.clsUpload
Attribute moUL.VB_VarHelpID = -1
Private mbHourly As Boolean
Private mbCheckingActiveFiles As Boolean
Private mbImporting As Boolean
'Help process security Tokens
Private mcolIBPFX As Collection

'True of current DSN Connection is Valid
Private mbValidDSNConn As Boolean


Private Sub Form_Load()
    On Error GoTo EH
    Dim sMess As String
    ' Don't want to be visible initially!
    Me.Visible = False
    Me.Caption = "Web Control-Auto Import (ON)"
    App.Title = Me.Caption
    FormWinRegPos Me
    SaveSetting "V2WebControl", "Msg", "Status", PicList.Idle
    'UNFlag Active command
    SaveSetting "V2WebControl", "WCService", "cmdActive", False
    ' Retrieve broadcast message sent by
    ' Windows when taskbar is created.
    m_TaskbarCreated = RegisterWindowMessage(TaskbarCreatedString)
    
    ' Setup MsgHook
    Msghook.HwndHook = Me.hWnd
    Msghook.Message(cbNotify) = True
    ' Msghook only accepts Integer-ranged values
    If m_TaskbarCreated > &H7FFF& Then
      Msghook.Message(m_TaskbarCreated - &H10000) = True
    Else
      Msghook.Message(m_TaskbarCreated) = True
    End If
    
    ' Setup icon notification from shell
    Call AddTrayIcon
        
   Exit Sub
EH:
    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
    sMess = sMess & "Private Sub Form_Load" & vbCrLf
    sMess = sMess & "ERROR # " & Err.Number & " " & Now & vbCrLf
    sMess = sMess & Err.Description & vbCrLf
    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
    ErrorLog sMess
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo EH
    Dim sMess As String
    If UnloadMode = vbFormControlMenu Then
        ' Just hide form if user presses Closes by X
        Me.Visible = False
        Cancel = True
    ElseIf UnloadMode = vbAppWindows Then
        SaveSetting "V2WebControl", "Msg", "Reset", True
        CloseConnection
        If Not goUtil Is Nothing Then
            goUtil.CLEANUP
            Set goUtil = Nothing
        End If
        DoEvents
        Sleep 1000
        End
    ElseIf UnloadMode = vbFormCode Then
        FormWinRegPos Me, True
        Call ShellNotifyIcon(NIM_DELETE, m_NID)
        CloseConnection
        If Not goUtil Is Nothing Then
            goUtil.CLEANUP
            Set goUtil = Nothing
        End If
        End
    End If
    Exit Sub
EH:
    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
    sMess = sMess & "Private Sub Form_QueryUnload" & vbCrLf
    sMess = sMess & "ERROR # " & Err.Number & " " & Now & vbCrLf
    sMess = sMess & Err.Description & vbCrLf
    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
    ErrorLog sMess
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo EH
    Dim sMess As String
    
    FormWinRegPos Me, True
    Call ShellNotifyIcon(NIM_DELETE, m_NID)
    CloseConnection
    Set mcolIBPFX = Nothing
    
    If Not goUtil Is Nothing Then
        goUtil.CLEANUP
        Set goUtil = Nothing
    End If
    
    Exit Sub
EH:
    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
    sMess = sMess & "Private Sub Form_Unload" & vbCrLf
    sMess = sMess & "ERROR # " & Err.Number & " " & Now & vbCrLf
    sMess = sMess & Err.Description & vbCrLf
    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
    ErrorLog sMess
End Sub


Private Sub ProgBarLoss_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error GoTo EH

    If Button = vbRightButton Then
        Me.PopupMenu mPopUp, vbPopupMenuRightButton, , , mPop(0)
    End If
    Exit Sub
EH:
    ShowError Err, "Private Sub ProgBarLoss_MouseUp", Me
End Sub


Private Sub ProcessTokens()
    On Error GoTo EH
    'Need to check for .tokin files uploaded from
    'users wanting to connect to the server to upload requests for info
    Dim sMess As String
    Dim sListADJFTP As String
    Dim vListADJFTP As Variant
    Dim colTokens As Collection
    Dim sToken As String
    Dim MyToken As TokenInfo
    Dim vToken As Variant
    Dim lPermissionErrorCount As Long
    Dim sAssignmentsPath As String
    Dim lPos As Long
    Dim iFFile As Integer
    Dim bCheckingTokins As Boolean
    Dim lErrorNumber As Long
    Dim sCommandLine As String
    Dim sUserName As String
    Dim bShowBusyIcon As Boolean
    
    '5.18.2005 BGS If the utility Object is nothing then
    'Bail this process
    If goUtil Is Nothing Then
        Exit Sub
    End If
    
    sListADJFTP = GetSetting("V2WebControl", "Dir", "ListADJFTP", vbNullString)
    
    'Need to look in Assignments token requests uploaded.
    lPos = InStr(1, sListADJFTP, "\Assignments\", vbTextCompare)
    If lPos > 0 Then
        sAssignmentsPath = Left(sListADJFTP, lPos) & "Assignments"
        If goUtil.utFileExists(sAssignmentsPath, True) Then
            'Be sure there is a User folders directory for this assignments path
            If Not goUtil.utFileExists(sAssignmentsPath & "\USER_FOLDERS", True) Then
                goUtil.utMakeDir (sAssignmentsPath & "\" & "\USER_FOLDERS")
            End If
            '.tokin for Token In from Adjuster Requesting Access
            sToken = Dir(sAssignmentsPath & "\*.tokin", vbNormal)
            Do Until sToken = vbNullString
                If Not bShowBusyIcon Then
                     ShowBusyIcon
                     bShowBusyIcon = True
                End If
                bCheckingTokins = True
                If colTokens Is Nothing Then
                    Set colTokens = New Collection
                End If
                lPermissionErrorCount = 0
                lErrorNumber = 0
                'Check to see if this file is still active
                iFFile = FreeFile
                On Error Resume Next
                Open sAssignmentsPath & "\" & sToken For Binary Access Read Lock Read As #iFFile
                'Error number 70 is Permissions Error, If the file is locked by another process
                'then Skip it and come back to it later
                If Err.Number = 70 Then
                    Err.Clear
                    On Error GoTo EH
                    GoTo NEXT_TOKIN
                End If
                On Error GoTo EH
                Close #iFFile
                MyToken.sToken = goUtil.utGetFileData(sAssignmentsPath & "\" & sToken)
                'once we have retrieved the token file we can kill it
                SetAttr sAssignmentsPath & "\" & sToken, vbNormal
                goUtil.utDeleteFile sAssignmentsPath & "\" & sToken
                'if the token is null sting then go to next tokin
                If Trim(MyToken.sToken) = vbNullString Then
                    GoTo NEXT_TOKIN:
                End If
                vToken = Split(MyToken.sToken, vbCrLf)
                'the token type is stored in the very first line of the file
                MyToken.iTokenType = vToken(0)
                'the Carrier is stored in the 2nd line of the file
                'This will Also Include Company Company\Carrier.
                MyToken.sCarrier = vToken(1)
                MyToken.sPath = sAssignmentsPath & "\" & sToken
                colTokens.Add MyToken, sToken
NEXT_TOKIN:
                bCheckingTokins = False
                sToken = Dir
            Loop
        End If
    End If
    
    If colTokens Is Nothing Then
        GoTo CLEANUP
    End If
    
    'Now that we have a collection of Tokens we need to process them
    'Shell ProcessUserTokin Via Batch Files
    txtMess.Text = "Processing Token(s) " & Now()
    For Each vToken In colTokens
        MyToken = vToken
        Select Case MyToken.iTokenType
            Case TokenType.Security
                'See if the file path exists for this user
                sUserName = Mid(MyToken.sPath, InStrRev(MyToken.sPath, "\", , vbBinaryCompare) + 1)
                sToken = sUserName
                sUserName = Left(sUserName, InStr(1, sUserName, "_", vbBinaryCompare) - 1)
                If Not goUtil.utFileExists(sAssignmentsPath & "\USER_FOLDERS\" & sUserName, True) Then
                    goUtil.utMakeDir (sAssignmentsPath & "\USER_FOLDERS\" & sUserName)
                End If
                'Include:
                '1. Dependency parameter
                '2. UserName
                '3. The Path to the User_Folder
                sCommandLine = "RunAsDepOfAutoImport|" & sUserName & "|" & sAssignmentsPath & "\USER_FOLDERS\" & sUserName & "\"
                'if the batch file already exists Get Rid of it
                If goUtil.utFileExists(sAssignmentsPath & "\USER_FOLDERS\" & sUserName & "\" & sUserName & ".bat") Then
                    goUtil.utDeleteFile sAssignmentsPath & "\USER_FOLDERS\" & sUserName & "\" & sUserName & ".bat"
                End If
                goUtil.utSaveFileData sAssignmentsPath & "\USER_FOLDERS\" & sUserName & "\" & sUserName & ".bat", """" & App.Path & "\ProcessUserTokin.exe "" """ & sCommandLine
                'Save the Tokin file to the User Folder
                goUtil.utSaveFileData sAssignmentsPath & "\USER_FOLDERS\" & sUserName & "\" & sToken, MyToken.sToken
                '##########################DEBUG Comment Out the Shell#########################
                Shell sAssignmentsPath & "\USER_FOLDERS\" & sUserName & "\" & sUserName & ".bat", vbHide
                '###################################End Debug##################################
        End Select
    Next
    txtMess.Text = txtMess.Text & vbCrLf & "Finished Processing " & Now()
    
    
CLEANUP:
    Set colTokens = Nothing
     If bShowBusyIcon Then
        ShowIdleIcon
    End If
    
    Exit Sub
EH:
    lErrorNumber = Err.Number
    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
    sMess = sMess & "Private Sub ProcessTokens" & vbCrLf
    sMess = sMess & "ERROR # " & lErrorNumber & " " & Now & vbCrLf
    sMess = sMess & Err.Description & vbCrLf
    sMess = sMess & sAssignmentsPath & "\" & sToken & vbCrLf
    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
    moUL_ErrorMess sMess
    Set colTokens = Nothing
    If bShowBusyIcon Then
       ShowIdleIcon
    End If
End Sub
    

Private Sub DeleteFLS()
    On Error GoTo EH
    '(DelFLSFiles) will monitor FLS.dat looking for Files that have Expired their
    'life span.  DelFLSFiles Can be processed from an EXE or a Service that calls DelFLSFiles
    Dim sMess As String
    Dim sFLSDatPath As String
        
    sFLSDatPath = GetSetting("V2WebControl", "Dir", "FLSdatPath", vbNullString)
    If sFLSDatPath <> vbNullString Then
        'Need to Add "\" if its not there
        If Right(sFLSDatPath, 1) <> "\" Then
            sFLSDatPath = sFLSDatPath & "\"
            SaveSetting "V2WebControl", "Dir", "FLSdatPath", sFLSDatPath
        End If
        goUtil.utDelFLSFiles sFLSDatPath
    End If
    
    Exit Sub
EH:
    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
    sMess = sMess & "Private Sub DeleteFLS" & vbCrLf
    sMess = sMess & "ERROR # " & Err.Number & " " & Now & vbCrLf
    sMess = sMess & Err.Description & vbCrLf
    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
    ErrorLog sMess
End Sub

Private Sub moUL_ErrorMess(ByVal Mess As String)
    On Error GoTo EH

    ErrorLog Mess
        
    Exit Sub
EH:
    Err.Clear
End Sub

Private Sub Msghook_Message(ByVal Msg As Long, ByVal wp As Long, ByVal lp As Long, result As Long)
    On Error GoTo EH
    Dim param As String
    Dim sMess As String
    Select Case Msg
        Case cbNotify
            If wp = uID Then
                Select Case lp
                    Case WM_MOUSEMOVE
                    Case WM_LBUTTONDOWN
                    Case WM_LBUTTONUP
                    Case WM_LBUTTONDBLCLK
                        ' Show form
                        Me.Visible = True
                        AppActivate Me.Caption
        
                    Case WM_RBUTTONDOWN
                    Case WM_RBUTTONUP
                    ' Display context menu
                    ' Highlight default (Open)
                    Call SetForegroundWindow(Me.hWnd)
                    Me.PopupMenu mPopUp, vbPopupMenuRightButton, , , mPop(0)
        
                    Case WM_RBUTTONDBLCLK
                    Case WM_MBUTTONDOWN
                    Case WM_MBUTTONUP
                    Case WM_MBUTTONDBLCLK
                    Case Else
                        param = "msg: " & Msg & ", wp: " & wp & ", lp: " & lp
                        Debug.Print "Message unknown!" & param
                End Select
            End If
        
        Case m_TaskbarCreated
            ' IE just (re)started the taskbar!
            Call AddTrayIcon
    End Select
   Exit Sub
EH:
    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
    sMess = sMess & "Private Sub Msghook_Message" & vbCrLf
    sMess = sMess & "ERROR # " & Err.Number & " " & Now & vbCrLf
    sMess = sMess & Err.Description & vbCrLf
    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
    ErrorLog sMess
End Sub

Private Sub mPop_Click(Index As Integer)
    On Error GoTo EH
    Dim sRet As String
    Dim sMess As String
    Dim sDefault As String
'<----------------------------VERY IMPORTANT NOTE---------------------------->
    'IMPORTANT !!! All Menu Tasks MUST GO IN HERE
    'Otherwise MESSAGE HOOK WILL MESS UP without this Post message Call
    ' Necessary to force task switch -- see Q135788
    Call PostMessage(Me.hWnd, WM_NULL, 0, 0)
'<----------------------------VERY IMPORTANT NOTE---------------------------->
   
    ' React to menu choice
    Select Case Index
        Case MenuList.Show  'Open (show form)
            Me.Visible = True
            AppActivate Me.Caption
      
        Case MenuList.Hide 'Hide
            Me.Visible = False
      
        Case MenuList.Disable  'Disable
            If Not mPop(2).Checked Then
                WCService NetPause
            Else
                WCService NetContinue
            End If
            m_NID.szTip = Me.Caption & vbNullChar
            Call ShellNotifyIcon(NIM_MODIFY, m_NID)
      
        Case MenuList.Restart 'ShutOff
            sMess = "Are you sure you really want to do that?"
            If MsgBox(sMess, vbInformation + vbYesNo, "Restart V2WebControl Service " & App.EXEName) = vbYes Then
                WCService NetRestart
            End If
        Case MenuList.Interval 'Interval
            sDefault = GetSetting("V2WebControl", "Msg", "UpdateHourly", 30000)
            If Val(sDefault) >= 1000 Then
                sDefault = Val(sDefault) / 1000
            End If
            '<-------------IMPORTANT NOTE----------------->
            'Must show the form before showing input box otherwise
            'it will mess up MESSAGE HOOK
            Me.Visible = True
            AppActivate Me.Caption
            '<-------------IMPORTANT NOTE----------------->
            sMess = "Please enter 10 to 60 seconds." & vbCrLf & vbCrLf
            sMess = sMess & "Or..." & vbCrLf & vbCrLf
            sMess = sMess & "Type the word 'Hourly' for an Hourly Interval."
            sRet = InputBox(sMess, "Update Interval", sDefault, Me.Left, Me.Top)
            
            If sRet = vbNullString Then
                sRet = sDefault
            End If
            
            If InStr(1, sRet, "Hourly", vbTextCompare) > 0 Then
                sRet = "Hourly"
            Else
                sRet = Val(sRet) * 1000
            End If
            
            SaveSetting "V2WebControl", "Msg", "UpdateHourly", sRet
        Case MenuList.ViewWebControl
            SaveSetting "V2WebControl", "Msg", "WebControlVisible", True
        Case MenuList.ECUpdateBatches
            Shell App.Path & "\V2ECUpdateBatches.exe RunAsDepOfV2AutoImport"
        Case MenuList.FailLogonVS 'Fail Logons
            sDefault = GetSetting("V2WebControl", "Msg", "FailLogonVS", vbNullString)
            '<-------------IMPORTANT NOTE----------------->
            'Must show the form before showing input box otherwise
            'it will mess up MESSAGE HOOK
            Me.Visible = True
            AppActivate Me.Caption
            '<-------------IMPORTANT NOTE----------------->
            sMess = "Enter Current Version Information." & vbCrLf & vbCrLf
            sMess = sMess & "IE..." & vbCrLf & vbCrLf
            sMess = sMess & "1.137.5"
            sRet = InputBox(sMess, "Current Version", sDefault, Me.Left, Me.Top)
            
            SaveSetting "V2WebControl", "Msg", "FailLogonVS", sRet
        End Select
    Exit Sub
EH:
    ShowError Err, "Private Sub mPop_Click", Me
End Sub

' *****************************************
'  Private Methods
' *****************************************
Private Sub AddTrayIcon()
    On Error GoTo EH
    Dim sMess As String
   ' Initialize NOTIFYICONDATA structure
   ' and add icon to tray.
   With m_NID
      .cbSize = Len(m_NID)
      .hWnd = Msghook.HwndHook
      .uID = uID
      .uFlags = NIF_MESSAGE Or NIF_TIP Or NIF_ICON
      .uCallbackMessage = cbNotify
      .hIcon = imgList.ListImages(PicList.Idle).Picture
      .szTip = Me.Caption & Chr(0)
   End With
   Call ShellNotifyIcon(NIM_ADD, m_NID)
   Exit Sub
EH:
    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
    sMess = sMess & "Private Sub AddTrayIcon" & vbCrLf
    sMess = sMess & "ERROR # " & Err.Number & " " & Now & vbCrLf
    sMess = sMess & Err.Description & vbCrLf
    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
    ErrorLog sMess
End Sub

Private Sub Timer_VerifyDSNConn_Timer()
    On Error GoTo EH
    Dim sMess As String
    
    'If the DSN Connection is Failing then can't allow
    'Process Tokens EXE to be launched
    ShowBusyIcon
    If VerifyDSNConn() Then
        mbValidDSNConn = True
        ShowIdleIcon
    Else
        mbValidDSNConn = False
    End If
    Timer_VerifyDSNConn.Enabled = True
    Exit Sub
EH:
    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
    sMess = sMess & "Private Sub Timer_Status_Timer" & vbCrLf
    sMess = sMess & "ERROR # " & Err.Number & " " & Now & vbCrLf
    sMess = sMess & Err.Description & vbCrLf
    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
    moUL_ErrorMess sMess
End Sub

Private Sub Timer_Status_Timer()
    On Error GoTo EH
    Dim iStatus As Integer
    Dim sUpdate As String
    Dim bEnabled As Boolean
    
    Dim sMess As String
    Dim sFLSDelTime As String
   
    
    'Update Status Here
    On Error Resume Next
    iStatus = GetSetting("V2WebControl", "Msg", "Status", 0)
    If Err.Number > 0 Then
        Err.Clear
        iStatus = 0
        SaveSetting "V2WebControl", "Msg", "Status", iStatus
    End If
    On Error GoTo EH
    If iStatus > 0 Then
        Select Case iStatus
            Case PicList.Disabled
                If m_NID.hIcon <> imgList.ListImages(PicList.Disabled).Picture Then
                    m_NID.hIcon = imgList.ListImages(PicList.Disabled).Picture
                    Timer_VerifyDSNConn.Enabled = False
                    mPop(2).Checked = True
                    Me.Caption = "Web Control-Auto Import (OFF)"
                    Me.Icon = imgList.ListImages(PicList.Disabled).Picture
                    m_NID.szTip = Me.Caption & vbNullChar
                    Call ShellNotifyIcon(NIM_MODIFY, m_NID)
                End If
            Case PicList.Idle
                If m_NID.hIcon <> imgList.ListImages(PicList.Idle).Picture Then
                    m_NID.hIcon = imgList.ListImages(PicList.Idle).Picture
                    Timer_VerifyDSNConn.Enabled = True
                    mPop(2).Checked = False
                    Me.Caption = "Web Control-Auto Import (ON)"
                    Me.Icon = imgList.ListImages(PicList.Idle).Picture
                    m_NID.szTip = Me.Caption & vbNullChar
                    Call ShellNotifyIcon(NIM_MODIFY, m_NID)
                End If
                
                'Enter other tasks here since we are idle
                '<---------------------Process Tokens HERE !!--------------------->
                'A Token request is a file FTP'd requesting an action by the
                'V2AutoImport server.  When the action has been accomplished the
                '.tokin file will be changed to .tokout.  The .tokout file will
                'have information concerning the .tokin file that was passed in.
                If mbValidDSNConn Then
                    ProcessTokens
                End If
                
                'File Life Span (Temp files that need to be deleted)
                sFLSDelTime = GetSetting("V2WebControl", "Msg", "FLSDelTime", "00:15")
                If StrComp(Format(Now(), "hh:mm"), sFLSDelTime, vbTextCompare) = 0 Then
                    DeleteFLS
                End If
            Case PicList.Busy
                If mbValidDSNConn Then
                    ProcessTokens
                End If
        End Select
    Else
        Unload Me
    End If
    
    'Update the Interval the server will check for Uploads
    sUpdate = GetSetting("V2WebControl", "Msg", "UpdateHourly", "False")
    If sUpdate = "Hourly" Then
        mbHourly = True
        'Still need to have the Timer event fire every
        '10 seconds to check for the hour
        If Timer_VerifyDSNConn.Interval <> 10000 Then
            bEnabled = Timer_VerifyDSNConn.Enabled
            Timer_VerifyDSNConn.Enabled = False
            Timer_VerifyDSNConn.Interval = 10000
            Timer_VerifyDSNConn.Enabled = bEnabled
        End If
    ElseIf sUpdate = "False" Then
DEFAULT_30:
        '30 seconds is the default
        mbHourly = False
        If Timer_VerifyDSNConn.Interval <> 30000 Then
            bEnabled = Timer_VerifyDSNConn.Enabled
            Timer_VerifyDSNConn.Enabled = False
            Timer_VerifyDSNConn.Interval = 30000
            Timer_VerifyDSNConn.Enabled = bEnabled
        End If
        SaveSetting "V2WebControl", "Msg", "UpdateHourly", "30000"
    Else
        If Val(sUpdate) >= 10000 And Val(sUpdate) <= 60000 Then
            mbHourly = False
            If Timer_VerifyDSNConn.Interval <> Val(sUpdate) Then
                bEnabled = Timer_VerifyDSNConn.Enabled
                Timer_VerifyDSNConn.Enabled = False
                Timer_VerifyDSNConn.Interval = Val(sUpdate)
                Timer_VerifyDSNConn.Enabled = bEnabled
            End If
         Else
            GoTo DEFAULT_30:
        End If
    End If
        
    Exit Sub
EH:
    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
    sMess = sMess & "Private Sub Timer_Status_Timer" & vbCrLf
    sMess = sMess & "ERROR # " & Err.Number & " " & Now & vbCrLf
    sMess = sMess & Err.Description & vbCrLf
    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
    moUL_ErrorMess sMess
End Sub

Private Sub ShowBusyIcon()
    On Error GoTo EH
    Dim sMess As String
    'Before we can import need to update
    'Icons and disable stuff
    Timer_VerifyDSNConn.Enabled = False
    Timer_Status.Enabled = False
    mPop(MenuList.Restart).Enabled = False
    
    m_NID.hIcon = imgList.ListImages(PicList.Busy).Picture
    SaveSetting "V2WebControl", "Msg", "Status", PicList.Busy
    Me.Icon = imgList.ListImages(PicList.Busy).Picture
    Me.Caption = "Web Control-Auto Import (Busy)"
    m_NID.szTip = Me.Caption & vbNullChar
    Call ShellNotifyIcon(NIM_MODIFY, m_NID)
    
    Exit Sub
EH:
    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
    sMess = sMess & "Private Sub ShowBusyIcon" & vbCrLf
    sMess = sMess & "ERROR # " & Err.Number & " " & Now & vbCrLf
    sMess = sMess & Err.Description & vbCrLf
    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
    moUL_ErrorMess sMess
End Sub

Private Sub ShowIdleIcon()
    On Error GoTo EH
    Dim sMess As String
    
    '1.22.2003 Check if V2WebControl Service was stopped before showing the
    ' Idle icon. In other words, only run showidleIcon if the current Status still says "Busy".
    'This will Allow V2AutoImport to unload if the Stop Service Event occurs in V2WebControl,
    'while at the same exact time V2AutoImport is in the middle of doing something "Busy" .
    'The new status (what ever that may be, most likely a stop service message) will be
    'able to be processed in Private Sub Timer_Status_Timer.
    If CLng(GetSetting("V2WebControl", "Msg", "Status", 0)) <> PicList.Busy Then
        Timer_VerifyDSNConn.Enabled = True
        Timer_Status.Enabled = True
        mPop(MenuList.Restart).Enabled = True
        Exit Sub
    End If
    
    'After import Update icons and reenable stuff
    m_NID.hIcon = imgList.ListImages(PicList.Idle).Picture
    SaveSetting "V2WebControl", "Msg", "Status", PicList.Idle
    Me.Icon = imgList.ListImages(PicList.Idle).Picture
    Me.Caption = "Web Control-Auto Import (ON)"
    m_NID.szTip = Me.Caption & vbNullChar
    Call ShellNotifyIcon(NIM_MODIFY, m_NID)
    
    ProgBarLoss.Value = 0
    
    Timer_VerifyDSNConn.Enabled = True
    Timer_Status.Enabled = True
    mPop(MenuList.Restart).Enabled = True
    
    Exit Sub
EH:
    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
    sMess = sMess & "Private Sub ShowIdleIcon" & vbCrLf
    sMess = sMess & "ERROR # " & Err.Number & " " & Now & vbCrLf
    sMess = sMess & Err.Description & vbCrLf
    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
    moUL_ErrorMess sMess
End Sub

Private Function CleanSQL(psSQL As String) As String
    Dim sMess As String
    Dim sSQL As String
    On Error GoTo EH
    
    sSQL = psSQL
    
    
    sSQL = Replace(sSQL, "'", "''", , , vbBinaryCompare)
    sSQL = Replace(sSQL, DT_z, "'", , , vbBinaryCompare)
    'Now Set the Begin and end String fields
    sSQL = Replace(sSQL, S_z, S_z_SET, , , vbBinaryCompare)
    sSQL = Replace(sSQL, z_S, z_S_SET, , , vbBinaryCompare)
    
    CleanSQL = sSQL
    
    Exit Function
EH:
    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
    sMess = sMess & "Private Function CleanSQL" & vbCrLf
    sMess = sMess & "ERROR # " & Err.Number & " " & Now & vbCrLf
    sMess = sMess & Err.Description & vbCrLf
    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
    moUL_ErrorMess sMess
End Function

Private Function VerifyDSNConn() As Boolean
    On Error GoTo EH
    Dim sUserID As String
    Dim sPassWord As String
    Dim sMess As String
    Dim Conn As ADODB.Connection
    Dim sDSN As String
    Dim sSQL As String
    Dim lSecStart As Long
    Dim lSecStop As Long
    Dim bSQLSVR As Boolean
    Dim RS As ADODB.Recordset
    Dim bV2ECcarFarmersclsLossXML01Exports As Boolean
    Dim lCheckV2ECcarFarmersclsLossXML01Exports As Long
    Dim RSV2ECcarFarmersclsLossXML01Exports As ADODB.Recordset
    'Process Packages vars
    Dim RSPackages As ADODB.Recordset
    'Need The CLient Company Name For
    'Reference incase of Any Errors...
    Dim sPackErrorMess As String
    Dim bPackError As Boolean
    Dim sClientCoName As String
    Dim sTemp As String
    Dim sThisUserName As String
    Dim sThisPassWord As String
    Dim sAdjUserName As String
    Dim sAssignmentsID As String
    Dim sPackageID As String
    Dim sCarListClassName As String
    Dim sPackageEmailQueueID As String
    
    
    '5.18.2005 BGS Check registry flag for XML exports
    'to see if it is currently set to true
    'If it is then no need to check for exports
    'Farmers Class for Eberls
    sTemp = GetSetting("V2WebControl", "V2ECcarFarmers.clsLossXML01", "Exports", False)
    bV2ECcarFarmersclsLossXML01Exports = CBool(sTemp)
    
    If Not bV2ECcarFarmersclsLossXML01Exports Then
        lCheckV2ECcarFarmersclsLossXML01Exports = 1
    End If
    
    
    'Get user ID And Pass Word
    sUserID = GetECSCryptSetting("V2WebControl", "DBConn", "USERID")
    sPassWord = GetECSCryptSetting("V2WebControl", "DBConn", "PASSWORD")
    
    'Check for same DSN for both, if true then using SQLSVR
    sDSN = GetSetting("V2WebControl", "DSN", "NAME", vbNullString)
    
    Set Conn = New ADODB.Connection
    
    'Time the Open call
    lSecStart = Timer
    Conn.Open sDSN, sUserID, sPassWord
    lSecStop = Timer
    If lSecStop - lSecStart > 2 Then
        Err.Raise -999, , sDSN & " Connection Problems.  Open took " & lSecStop - lSecStart & " seconds to execute. "
    End If
    
    'Check the SP that Verifies Connections to Linked servers as well...
    'Insert Code to Use Gary's SP that verifies
    'Connectivity to all linked tables.
    
    sSQL = "Exec spaCheckLinks 'ClaimProdV1', "                         '@Links  Varchar(500),
    sSQL = sSQL & CStr(lCheckV2ECcarFarmersclsLossXML01Exports) & " "   '@lCheckV2ECcarFarmersclsLossXML01Exports bit =0
    
    Set RS = New ADODB.Recordset
    RS.CursorLocation = adUseClient
    RS.Open sSQL, Conn, adOpenForwardOnly, adLockReadOnly

    If Not RS.EOF Then
        If Not CBool(RS.Fields("Valid").Value) Then
            RS.Close
            Set RS = Nothing
            Err.Raise -999, , sDSN & " Connection Problems. Linked Server Failed."
        End If
    End If
    
    'Again, Only check for the next recordset if needed !
    
    '5.18.2005 Farmers Class
    'if the current flag says there is none , then check it again
    'and set it to true if applicable.
    If Not bV2ECcarFarmersclsLossXML01Exports Then
        Set RSV2ECcarFarmersclsLossXML01Exports = RS.NextRecordset
        Set RSPackages = RS.NextRecordset
        If Not RSV2ECcarFarmersclsLossXML01Exports.EOF Then
            bV2ECcarFarmersclsLossXML01Exports = CBool(RSV2ECcarFarmersclsLossXML01Exports.Fields("bV2ECcarFarmersclsLossXML01Exports").Value)
            'ONLY SET THIS FLAG TO TRUE
            'V2ECcarFarmers.clsLossXML01 will be responsible for setting this
            'flag false !!!
            If bV2ECcarFarmersclsLossXML01Exports Then
                SaveSetting "V2WebControl", "V2ECcarFarmers.clsLossXML01", "Exports", True
            End If
        End If
    End If
    
    '6.14.2005 BGS Check for Packages to be processed
    sSQL = "Exec spsGetPackageToProcess "
    Set RSPackages = New ADODB.Recordset
    RSPackages.CursorLocation = adUseClient
    RSPackages.Open sSQL, Conn, adOpenForwardOnly, adLockReadOnly
    Set RSPackages.ActiveConnection = Nothing
    
    If Not RSPackages.EOF Then
        RSPackages.MoveFirst
        sClientCoName = goUtil.IsNullIsVbNullString(RSPackages.Fields("ClientCoName"))
        sThisUserName = sUserID
        sThisPassWord = sPassWord
        sAdjUserName = goUtil.IsNullIsVbNullString(RSPackages.Fields("AdjUserName"))
        sAssignmentsID = goUtil.IsNullIsVbNullString(RSPackages.Fields("AssignmentsID"))
        sPackageID = goUtil.IsNullIsVbNullString(RSPackages.Fields("PackageID"))
        sCarListClassName = goUtil.IsNullIsVbNullString(RSPackages.Fields("CarListClassName"))
        sPackageEmailQueueID = vbNullString
        'Now process the poop out of this!
        ProcessPackages sThisUserName, _
                        sThisPassWord, _
                        sAdjUserName, _
                        sAssignmentsID, _
                        sPackageID, _
                        sCarListClassName, _
                        sPackageEmailQueueID
    End If
    
    '7.28.2005 BGS Check for Packages to be Emailed
    sSQL = "Exec spsGetPackageEmailQueue "
    Set RSPackages = New ADODB.Recordset
    RSPackages.CursorLocation = adUseClient
    RSPackages.Open sSQL, Conn, adOpenForwardOnly, adLockReadOnly
    Set RSPackages.ActiveConnection = Nothing
    
    If Not RSPackages.EOF Then
        RSPackages.MoveFirst
        sClientCoName = goUtil.IsNullIsVbNullString(RSPackages.Fields("ClientCoName"))
        sThisUserName = sUserID
        sThisPassWord = sPassWord
        sAdjUserName = goUtil.IsNullIsVbNullString(RSPackages.Fields("AdjUserName"))
        sAssignmentsID = goUtil.IsNullIsVbNullString(RSPackages.Fields("AssignmentsID"))
        sPackageID = goUtil.IsNullIsVbNullString(RSPackages.Fields("PackageID"))
        sCarListClassName = goUtil.IsNullIsVbNullString(RSPackages.Fields("CarListClassName"))
        sPackageEmailQueueID = goUtil.IsNullIsVbNullString(RSPackages.Fields("PackageEmailQueueID"))
        'Now process the poop out of this!
        ProcessPackages sThisUserName, _
                        sThisPassWord, _
                        sAdjUserName, _
                        sAssignmentsID, _
                        sPackageID, _
                        sCarListClassName, _
                        sPackageEmailQueueID
    End If
    
    
    
    Set RSV2ECcarFarmersclsLossXML01Exports = Nothing
    Set RSPackages = Nothing
    Set RS = Nothing
    Set Conn = Nothing
    
    VerifyDSNConn = True
    sMess = "Data Connections Verified."
    txtMess.Text = sMess & " " & Now() & vbCrLf
    Exit Function
PACK_ERROR:
    'Raise an Error... however make sure the Verify Conn is True
    'Since this is a problem that is not realted to Connectivity
    VerifyDSNConn = True
    sPackErrorMess = "Can't Process Packages!" & vbCrLf & vbCrLf
    sPackErrorMess = sPackErrorMess & "Single File / Email Settings are invlaid for Client Company: " & sClientCoName & vbCrLf
    Err.Raise -999, , sPackErrorMess
    Exit Function
EH:
    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
    sMess = sMess & "Private Function VerifyDSNConn" & vbCrLf
    sMess = sMess & "ERROR # " & Err.Number & " " & Now & vbCrLf
    sMess = sMess & "DSN NAME: " & sDSN & vbCrLf
    sMess = sMess & Err.Description & vbCrLf
    sMess = sMess & "All DSN connections must validate before Adjuster upload files will be processed." & vbCrLf
    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
    Set RSV2ECcarFarmersclsLossXML01Exports = Nothing
    Set RSPackages = Nothing
    Set RS = Nothing
    Set Conn = Nothing
    moUL_ErrorMess sMess
    txtMess.Text = sMess & " " & Now() & vbCrLf
    txtMess.Refresh
    Sleep 1000
End Function

Private Function ProcessPackages(psUserName As String, _
                                 psPassWord As String, _
                                 psAdjUserName As String, _
                                 psAssignmentsID As String, _
                                 psPackageID As String, _
                                 Optional psCarListClassName As String, _
                                 Optional psPackageEmailQueueID As String) As Boolean
    On Error GoTo EH
    Dim sCommandLine As String
    Dim sAssignmentsPath As String
    Dim sFTPSitePath As String
    Dim lPos As Long
    Dim sTemp As String
    Dim sMess As String
    
    'Set the FTP Site path
    sFTPSitePath = GetSetting("V2WebControl", "Dir", "FTPSitePath", vbNullString)
    sAssignmentsPath = Replace(sFTPSitePath, "\Upload\", "\", , , vbTextCompare)
    
    'Include:
    '1. Dependency parameter
    sCommandLine = "RunAsDepOfAutoImport" & "|"
    sCommandLine = sCommandLine & psUserName & "|"
    sCommandLine = sCommandLine & psPassWord & "|"
    sCommandLine = sCommandLine & psAdjUserName & "|"
    sCommandLine = sCommandLine & psAssignmentsID & "|"
    sCommandLine = sCommandLine & psPackageID & "|"
    sCommandLine = sCommandLine & psCarListClassName & "|"
    sCommandLine = sCommandLine & psPackageEmailQueueID
    'if the batch file already exists Get Rid of it
    
    'first see if the actual User Folder Directory Exists... if not then
    'Create it !
    'if this top level Directory does not Exist then
    'This is BAD!  raise an error
    If Not goUtil.utFileExists(sAssignmentsPath & "\USER_FOLDERS", True) Then
        Err.Raise -999, , sAssignmentsPath & "\USER_FOLDERS\ " & vbCrLf & "Directory missing!"
    End If
    
    'Check for the UserName Folder
    If Not goUtil.utFileExists(sAssignmentsPath & "\USER_FOLDERS\" & psAdjUserName, True) Then
        sTemp = goUtil.utMakeDir(sAssignmentsPath & "\USER_FOLDERS\" & psAdjUserName)
        If sTemp <> vbNullString Then
            Err.Raise -999, , sTemp
        End If
    End If
    
    'Check for ProcessPackages Folder
    If Not goUtil.utFileExists(sAssignmentsPath & "\USER_FOLDERS\" & psAdjUserName & "\PROCESS_PACKAGES", True) Then
        sTemp = goUtil.utMakeDir(sAssignmentsPath & "\USER_FOLDERS\" & psAdjUserName & "\PROCESS_PACKAGES")
        If sTemp <> vbNullString Then
            Err.Raise -999, , sTemp
        End If
    End If
    
    'Check For BUILD Folder under the Process packages Folder
    If Not goUtil.utFileExists(sAssignmentsPath & "\USER_FOLDERS\" & psAdjUserName & "\PROCESS_PACKAGES\BUILD", True) Then
        sTemp = goUtil.utMakeDir(sAssignmentsPath & "\USER_FOLDERS\" & psAdjUserName & "\PROCESS_PACKAGES\BUILD")
        If sTemp <> vbNullString Then
            Err.Raise -999, , sTemp
        End If
    End If
    
    'Under the BUILD Folder CHeck for Assignments Folder
    If Not goUtil.utFileExists(sAssignmentsPath & "\USER_FOLDERS\" & psAdjUserName & "\PROCESS_PACKAGES\BUILD\ASSIGNMENTS", True) Then
        sTemp = goUtil.utMakeDir(sAssignmentsPath & "\USER_FOLDERS\" & psAdjUserName & "\PROCESS_PACKAGES\BUILD\ASSIGNMENTS")
        If sTemp <> vbNullString Then
            Err.Raise -999, , sTemp
        End If
    End If
    'Now Check for the AssignmentsID Folder Under the Assignments Folder
    If Not goUtil.utFileExists(sAssignmentsPath & "\USER_FOLDERS\" & psAdjUserName & "\PROCESS_PACKAGES\BUILD\ASSIGNMENTS\" & psAssignmentsID, True) Then
        sTemp = goUtil.utMakeDir(sAssignmentsPath & "\USER_FOLDERS\" & psAdjUserName & "\PROCESS_PACKAGES\BUILD\ASSIGNMENTS\" & psAssignmentsID)
        If sTemp <> vbNullString Then
            Err.Raise -999, , sTemp
        End If
    End If
    'Check for the packages Folder Under the AssignmentsID Folder
    If Not goUtil.utFileExists(sAssignmentsPath & "\USER_FOLDERS\" & psAdjUserName & "\PROCESS_PACKAGES\BUILD\ASSIGNMENTS\" & psAssignmentsID & "\PACKAGES", True) Then
        sTemp = goUtil.utMakeDir(sAssignmentsPath & "\USER_FOLDERS\" & psAdjUserName & "\PROCESS_PACKAGES\BUILD\ASSIGNMENTS\" & psAssignmentsID & "\PACKAGES")
        If sTemp <> vbNullString Then
            Err.Raise -999, , sTemp
        End If
    End If
    'Check for the PackageID folder under the Packages Folder
    If Not goUtil.utFileExists(sAssignmentsPath & "\USER_FOLDERS\" & psAdjUserName & "\PROCESS_PACKAGES\BUILD\ASSIGNMENTS\" & psAssignmentsID & "\PACKAGES\" & psPackageID, True) Then
        sTemp = goUtil.utMakeDir(sAssignmentsPath & "\USER_FOLDERS\" & psAdjUserName & "\PROCESS_PACKAGES\BUILD\ASSIGNMENTS\" & psAssignmentsID & "\PACKAGES\" & psPackageID)
        If sTemp <> vbNullString Then
            Err.Raise -999, , sTemp
        End If
    End If
    'Check For Error Folder under the PackageID packages Folder
    If Not goUtil.utFileExists(sAssignmentsPath & "\USER_FOLDERS\" & psAdjUserName & "\PROCESS_PACKAGES\BUILD\ASSIGNMENTS\" & psAssignmentsID & "\PACKAGES\" & psPackageID & "\ERRORS", True) Then
        sTemp = goUtil.utMakeDir(sAssignmentsPath & "\USER_FOLDERS\" & psAdjUserName & "\PROCESS_PACKAGES\BUILD\ASSIGNMENTS\" & psAssignmentsID & "\PACKAGES\" & psPackageID & "\ERRORS")
        If sTemp <> vbNullString Then
            Err.Raise -999, , sTemp
        End If
    End If
    
    
    If goUtil.utFileExists(sAssignmentsPath & "\USER_FOLDERS\" & psAdjUserName & "\PROCESS_PACKAGES\" & psAdjUserName & ".bat") Then
        goUtil.utDeleteFile sAssignmentsPath & "\USER_FOLDERS\" & psAdjUserName & "\PROCESS_PACKAGES\" & psAdjUserName & ".bat"
    End If
    If goUtil.utFileExists(sAssignmentsPath & "\USER_FOLDERS\" & psAdjUserName & "\PROCESS_PACKAGES\" & psAdjUserName & ".bat") Then
        goUtil.utDeleteFile (sAssignmentsPath & "\USER_FOLDERS\" & psAdjUserName & "\PROCESS_PACKAGES\" & psAdjUserName & ".bat")
    End If
    goUtil.utSaveFileData sAssignmentsPath & "\USER_FOLDERS\" & psAdjUserName & "\PROCESS_PACKAGES\" & psAdjUserName & ".bat", """" & App.Path & "\ProcessPackages.exe "" """ & sCommandLine
    '##########################DEBUG Comment Out the Shell#########################
    Shell sAssignmentsPath & "\USER_FOLDERS\" & psAdjUserName & "\PROCESS_PACKAGES\" & psAdjUserName & ".bat", vbHide
    
    Exit Function
EH:
    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
    sMess = sMess & "Private Function ProcessPackages" & vbCrLf
    sMess = sMess & "ERROR # " & Err.Number & " " & Now & vbCrLf
    sMess = sMess & Err.Description & vbCrLf
    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
    moUL_ErrorMess sMess
    txtMess.Text = sMess & " " & Now() & vbCrLf
    txtMess.Refresh
    Sleep 1000
End Function
