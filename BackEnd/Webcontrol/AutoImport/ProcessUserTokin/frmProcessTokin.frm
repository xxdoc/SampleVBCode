VERSION 5.00
Object = "{307C5043-76B3-11CE-BF00-0080AD0EF894}#1.0#0"; "MsgHoo32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProcessTokin 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Process User Tokin (User name)"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6330
   Icon            =   "frmProcessTokin.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   6330
   Begin VB.Timer Timer_Status 
      Interval        =   500
      Left            =   600
      Top             =   600
   End
   Begin VB.Timer Timer_Import 
      Interval        =   2000
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
            Picture         =   "frmProcessTokin.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProcessTokin.frx":0894
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProcessTokin.frx":0CE6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar ProgBarLoss 
      Height          =   375
      Left            =   40
      TabIndex        =   0
      ToolTipText     =   "Right-Click for Options"
      Top             =   2160
      Width           =   6000
      _ExtentX        =   10583
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
      Height          =   2535
      Left            =   40
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   0
      Width           =   6255
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
         Caption         =   "&Exit"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmProcessTokin"
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
    StopMe
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
Private mbExporting As Boolean
'Help process security Tokens
Private msUserName As String
Private msUsersID As String
Private msUserFolderPath As String
Private mbSHUTDOWN As Boolean
'Token info to be passed to Export Download (ExportDL)
Private msMyTokoutPath As String
Private msMyTokenData As String


Public Property Let UserFolderPath(psUserFolderPath As String)
    msUserFolderPath = psUserFolderPath
End Property

Public Property Get UserFolderPath() As String
    UserFolderPath = msUserFolderPath
End Property

Public Property Let UserName(psUserName As String)
    msUserName = psUserName
End Property

Public Property Get UserName() As String
    UserName = msUserName
End Property

Private Sub Form_Load()
    On Error GoTo EH
    Dim sMess As String
    Dim sTokOut As String
    Dim vTokOut As Variant
    Dim colTokout As Collection
    
    ' Don't want to be visible initially!
    Me.Visible = False
    Me.Caption = "Process User Tokin (" & msUserName & ")"
    App.Title = Me.Caption
    FormWinRegPos Me
'    SaveSetting "V2AutoImport", "Msg", "Status", PicList.Idle
    'UNFlag Active command
'    SaveSetting "V2AutoImport", "WCService", "cmdActive", False
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
    
    'Get rid of any .tokout files that may exist
    sTokOut = Dir(msUserFolderPath & "\*.tokout", vbNormal)
    Do Until sTokOut = vbNullString
        If colTokout Is Nothing Then
            Set colTokout = New Collection
        End If
        colTokout.Add sTokOut, sTokOut
        sTokOut = Dir
    Loop
    'Also get rid of any Flags that are hanging
    sTokOut = Dir(msUserFolderPath & "\*.flag", vbNormal)
    Do Until sTokOut = vbNullString
        If colTokout Is Nothing Then
            Set colTokout = New Collection
        End If
        colTokout.Add sTokOut, sTokOut
        sTokOut = Dir
    Loop
   
    sTokOut = Dir(msUserFolderPath & "\*.zdl", vbNormal)
    Do Until sTokOut = vbNullString
        If colTokout Is Nothing Then
            Set colTokout = New Collection
        End If
        colTokout.Add sTokOut, sTokOut
        sTokOut = Dir
    Loop
    sTokOut = Dir(msUserFolderPath & "\*.dl", vbNormal)
    Do Until sTokOut = vbNullString
        If colTokout Is Nothing Then
            Set colTokout = New Collection
        End If
        colTokout.Add sTokOut, sTokOut
        sTokOut = Dir
    Loop
    sTokOut = Dir(msUserFolderPath & "\*.zul", vbNormal)
    Do Until sTokOut = vbNullString
        If colTokout Is Nothing Then
            Set colTokout = New Collection
        End If
        colTokout.Add sTokOut, sTokOut
        sTokOut = Dir
    Loop
    sTokOut = Dir(msUserFolderPath & "\*.ul", vbNormal)
    Do Until sTokOut = vbNullString
        If colTokout Is Nothing Then
            Set colTokout = New Collection
        End If
        colTokout.Add sTokOut, sTokOut
        sTokOut = Dir
    Loop
    
    If Not colTokout Is Nothing Then
        For Each vTokOut In colTokout
            sTokOut = vTokOut
            goUtil.utDeleteFile msUserFolderPath & "\" & sTokOut
        Next
    End If
    
    Set colTokout = Nothing
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
    ElseIf UnloadMode = vbFormCode Then
        FormWinRegPos Me, True
        Call ShellNotifyIcon(NIM_DELETE, m_NID)
        CloseConnection
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
    
    Set gDB = Nothing
    Set gWS = Nothing
    CloseConnection
    
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

'Private Sub moUL_UpdateDBRT(ByVal vBatches As Variant, poBatch As V2ECKeyBoard.clsBatches, poULRT As V2ECKeyBoard.clsCarUL, poUL As V2ECKeyBoard.clsUpload)
'    On Error GoTo EH
'    Dim sMess As String
'    Dim MyBat As V2ECKeyBoard.udtBatchesRT
'
'    MyBat = vBatches
'    If UCase(MyBat.sStatus) = "DELETED" Then
'        UpdateDBRT_SQLServer vBatches, poBatch, poULRT, poUL, True
'    Else
'        UpdateDBRT_SQLServer vBatches, poBatch, poULRT, poUL
'    End If
'
'    Exit Sub
'EH:
'    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
'    sMess = sMess & "Private Sub moUL_UpdateDBRT" & vbCrLf
'    sMess = sMess & "ERROR # " & Err.Number & " " & Now & vbCrLf
'    sMess = sMess & Err.Description & vbCrLf
'    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
'    ErrorLog sMess
'End Sub

Private Sub ProgBarLoss_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error GoTo EH

    If Button = vbRightButton Then
        Me.PopupMenu mPopUp, vbPopupMenuRightButton, , , mPop(0)
    End If
    Exit Sub
EH:
    ShowError Err, "Private Sub ProgBarLoss_MouseUp", Me
End Sub

Private Sub Timer_Import_Timer()
    On Error GoTo EH
    Dim sTime As String
    Dim sMess As String
    
    ShowBusyIcon

    'Check for Shut Down Tokin Process flag
    If ClientFlagCheck("UploadUpdateReady.flag", , 1) Then
        ImportULUpdates
    ElseIf ClientFlagCheck("UploadReady.flag", , 1) Then
        ImportUL msMyTokoutPath, msMyTokenData
    ElseIf ClientFlagCheck("ShutDownTokinProcess.flag", , 1) Then
        'Need to check for UL Updates before exiting
        ImportULUpdates
        txtMess.Text = "Client disconnected... Exiting process. " & Now()
        txtMess.Refresh
        DoEvents
        Sleep 1000
        mbSHUTDOWN = True
    End If
    
    
    ShowIdleIcon
    
    
    Exit Sub
EH:
    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
    sMess = sMess & "Private Sub Timer_Import_Timer" & vbCrLf
    sMess = sMess & "ERROR # " & Err.Number & " " & Now & vbCrLf
    sMess = sMess & Err.Description & vbCrLf
    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
    moUL_ErrorMess sMess
    Timer_Status.Enabled = True
    Timer_Import.Enabled = True
End Sub

Private Sub ImportULUpdates()
    On Error GoTo EH
    Dim sMess As String
    
    mbImporting = True
    ImportULUpdates_SQLServer
    mbImporting = False
    
    Exit Sub
EH:
    mbImporting = False
    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
    sMess = sMess & "Private Sub ImportULUpdates" & vbCrLf
    sMess = sMess & "ERROR # " & Err.Number & " " & Now & vbCrLf
    sMess = sMess & Err.Description & vbCrLf
    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
    ErrorLog sMess
    Set moUL = Nothing
End Sub

Private Sub ImportUL(psMyTokoutPath As String, psMyTokenData As String)
    On Error GoTo EH
    Dim sMess As String
    
    mbImporting = True
    ImportUL_SQLServer psMyTokoutPath, psMyTokenData
    mbImporting = False
    
    Exit Sub
EH:
    mbImporting = False
    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
    sMess = sMess & "Private Sub ImportUL" & vbCrLf
    sMess = sMess & "ERROR # " & Err.Number & " " & Now & vbCrLf
    sMess = sMess & Err.Description & vbCrLf
    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
    ErrorLog sMess
    Set moUL = Nothing
End Sub

Private Sub ExportDL(psMyTokoutPath As String, psMyTokenData As String)
    On Error GoTo EH
    Dim sMess As String
    
    mbExporting = True
    ExportDL_SQLServer psMyTokoutPath, psMyTokenData
    mbExporting = False
    
    Exit Sub
EH:
    mbExporting = False
    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
    sMess = sMess & "Private Sub ExportDL" & vbCrLf
    sMess = sMess & "ERROR # " & Err.Number & " " & Now & vbCrLf
    sMess = sMess & Err.Description & vbCrLf
    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
    ErrorLog sMess
    Set moUL = Nothing
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
    
    sToken = Dir(msUserFolderPath & "\*.tokin", vbNormal)
    Do Until sToken = vbNullString
        bCheckingTokins = True
        If colTokens Is Nothing Then
            Set colTokens = New Collection
        End If
        lPermissionErrorCount = 0
        lErrorNumber = 0
        'Check to see if this file is still active
        iFFile = FreeFile
        On Error Resume Next
        Open msUserFolderPath & "\" & sToken For Binary Access Read Lock Read As #iFFile
        'Error number 70 is Permissions Error, If the file is locked by another process
        'then Skip it and come back to it later
        If Err.Number = 70 Then
            Err.Clear
            On Error GoTo EH
            GoTo NEXT_TOKIN
        End If
        On Error GoTo EH
        Close #iFFile
        MyToken.sToken = goUtil.utGetFileData(msUserFolderPath & "\" & sToken)
        'once we have retrieved the token file we can kill it
        SetAttr msUserFolderPath & "\" & sToken, vbNormal
        goUtil.utDeleteFile msUserFolderPath & "\" & sToken

        vToken = Split(MyToken.sToken, vbCrLf)
        'the token type is stored in the very first line of the file
        MyToken.iTokenType = vToken(0)
        'the Carrier is stored in the 2nd line of the file
        'This will Also Include Company Company\Carrier.
        MyToken.sCarrier = vToken(1)
        MyToken.sPath = msUserFolderPath & "\" & sToken
        colTokens.Add MyToken, sToken
NEXT_TOKIN:
        bCheckingTokins = False
        sToken = Dir
    Loop

    If colTokens Is Nothing Then
        GoTo CLEANUP
    End If

    'Now that we have a collection of Tokens we need to process them
    ShowBusyIcon
    txtMess.Text = "Processing Token(s) " & Now()
    For Each vToken In colTokens
        MyToken = vToken

        Select Case MyToken.iTokenType
            Case TokenType.Security
                ProcessSecurityToken_SQLServer MyToken
        End Select
    Next
    txtMess.Text = txtMess.Text & vbCrLf & "Waiting for Client to finish processing" & Now()
    ShowIdleIcon

CLEANUP:
    Set colTokens = Nothing

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
End Sub
    

Private Sub DeleteFLS()
'    On Error GoTo EH
'    '(DelFLSFiles) will monitor FLS.dat looking for Files that have Expired their
'    'life span.  DelFLSFiles Can be processed from an EXE or a Service that calls DelFLSFiles
'    Dim sMess As String
'    Dim sFLSDatPath As String
'    Dim oUtil As Object
'
'    sFLSDatPath = GetSetting("V2WebControl", "Dir", "FLSdatPath", vbNullString)
'    If sFLSDatPath <> vbNullString Then
'        'Need to Add "\" if its not there
'        If Right(sFLSDatPath, 1) <> "\" Then
'            sFLSDatPath = sFLSDatPath & "\"
''            SaveSetting "V2WebControl", "Dir", "FLSdatPath", sFLSDatPath
'        End If
'        Set oUtil = CreateObject("V2ECKeyBoard.clsUtil")
'        oUtil.DelFLSFiles sFLSDatPath
'    End If
'
'    Set oUtil = Nothing
'    Exit Sub
'EH:
'    Set oUtil = Nothing
'    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
'    sMess = sMess & "Private Sub DeleteFLS" & vbCrLf
'    sMess = sMess & "ERROR # " & Err.Number & " " & Now & vbCrLf
'    sMess = sMess & Err.Description & vbCrLf
'    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
'    ErrorLog sMess
End Sub

Private Function RollBackSYNCH(plRecordsAffected As Long, psMess As String) As Boolean
    On Error GoTo EH
    Dim sProdDSN As String
    Dim sSQL As String
    Dim lRecordsAffected As Long
    Dim lTotalRecordsAffected As Long
    Dim sMess As String
    Dim lCount As Long
    Dim sSynchFile As String ' *.syc
    Dim sULFile As String ' *.zul
    Dim dicULUpdateFiles As Scripting.Dictionary
    Dim sTemp As String
    Dim sPassedInMess As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim bBeginRollback As Boolean
    Dim bEndRollback As Boolean
    
    sPassedInMess = psMess
    psMess = vbNullString
    
    'Look for *.syc and *.zul
    'SYNCH *.syc
    sSynchFile = Dir(msUserFolderPath & "*.syc", vbNormal)
    Do Until sSynchFile = vbNullString
        If dicULUpdateFiles Is Nothing Then
            Set dicULUpdateFiles = New Scripting.Dictionary
        End If
        dicULUpdateFiles.Add sSynchFile, sSynchFile
        sSynchFile = Dir
    Loop
    'ULFILE *.zul
    sULFile = Dir(msUserFolderPath & "*.zul", vbNormal)
    Do Until sULFile = vbNullString
        If dicULUpdateFiles Is Nothing Then
            Set dicULUpdateFiles = New Scripting.Dictionary
        End If
        dicULUpdateFiles.Add sULFile, sULFile
        sSynchFile = Dir
    Loop
    
    If dicULUpdateFiles Is Nothing Then
        'Check the Passed in Message for anything that may require a rollback
        If InStr(1, sPassedInMess, "public function synchronizeultable", vbTextCompare) > 0 Then
            RollBackSYNCH = True
            GoTo ROLL_BACK
        Else
            'This means there were no SYNCH Problems
            GoTo CLEANUP
        End If
    Else
        'This means YUP there were Synch problems, which include the following
        'zero length files, or more than likely corrupted files = "garbage data partial data"
        RollBackSYNCH = True
        'clean up / remove all *.zul and *.syc files
        sTemp = goUtil.utDeleteFile(msUserFolderPath & "*.zul")
        If sTemp <> vbNullString Then
            sMess = sMess & "No *.zul files were cleaned, reason:" & vbCrLf & "Error #" & sTemp & vbCrLf
        End If
        sTemp = goUtil.utDeleteFile(msUserFolderPath & "*.syc")
        If sTemp <> vbNullString Then
            sMess = sMess & "No *.syc files were cleaned, reason:" & vbCrLf & "Error #" & sTemp & vbCrLf
        End If
    End If
    
ROLL_BACK:
    sProdDSN = GetSetting("V2WebControl", "DSN", "NAME", vbNullString)
    OpenConnection sProdDSN
    bBeginRollback = True
    'PackageItemHistory
    sSQL = "DELETE FROM PackageItemHistory "
    sSQL = sSQL & "WHERE PackageItemID IN ( "
    sSQL = sSQL & "SELECT PackageItemID FROM PackageItem "
    sSQL = sSQL & "WHERE ([ID] < 0 OR [IDRTAttachments] < 0) "
    sSQL = sSQL & "AND [UpdateByUserID] = " & msUsersID & " "
    sSQL = sSQL & ") "
    gConn.Execute sSQL, lRecordsAffected
    If lRecordsAffected > 0 Then
        lTotalRecordsAffected = lTotalRecordsAffected + lRecordsAffected
        sMess = sMess & "Removed " & lRecordsAffected & " From Table [PackageItemHistory] " & vbCrLf
    End If
    
    'PackageItem
    sSQL = "DELETE FROM PackageItem "
    sSQL = sSQL & "WHERE ([ID] < 0 or [IDRTAttachments] < 0) "
    sSQL = sSQL & "AND [UpdateByUserID] = " & msUsersID & " "
    gConn.Execute sSQL, lRecordsAffected
    If lRecordsAffected > 0 Then
        lTotalRecordsAffected = lTotalRecordsAffected + lRecordsAffected
        sMess = sMess & "Removed " & lRecordsAffected & " From Table [PackageItem] " & vbCrLf
    End If
    
    'PackageHistory
    sSQL = "DELETE FROM PackageHistory "
    sSQL = sSQL & "WHERE PackageID IN ( "
    sSQL = sSQL & "SELECT PackageID "
    sSQL = sSQL & "FROM Package WHERE [ID] < 0 "
    sSQL = sSQL & "AND [UpdateByUserID] = " & msUsersID & " "
    sSQL = sSQL & ") "
    gConn.Execute sSQL, lRecordsAffected
    If lRecordsAffected > 0 Then
        lTotalRecordsAffected = lTotalRecordsAffected + lRecordsAffected
        sMess = sMess & "Removed " & lRecordsAffected & " From Table [PackageHistory] " & vbCrLf
    End If
    
    'Package
    sSQL = "DELETE  FROM Package "
    sSQL = sSQL & "WHERE [ID] < 0 "
    sSQL = sSQL & "AND [UpdateByUserID] = " & msUsersID & " "
    gConn.Execute sSQL, lRecordsAffected
    If lRecordsAffected > 0 Then
        lTotalRecordsAffected = lTotalRecordsAffected + lRecordsAffected
        sMess = sMess & "Removed " & lRecordsAffected & " From Table [Package] " & vbCrLf
    End If
    
    'MiscReportParam
    sSQL = "DELETE FROM MiscReportParam "
    sSQL = sSQL & "WHERE [ID] < 0 "
    sSQL = sSQL & "AND [UpdateByUserID] = " & msUsersID & " "
    gConn.Execute sSQL, lRecordsAffected
    If lRecordsAffected > 0 Then
        lTotalRecordsAffected = lTotalRecordsAffected + lRecordsAffected
        sMess = sMess & "Removed " & lRecordsAffected & " From Table [MiscReportParam] " & vbCrLf
    End If
    
    'MiscReportParam01 to MiscReportParam30
    For lCount = 1 To 30
        sSQL = "DELETE FROM MiscReportParam" & Format(lCount, "00") & " "
        sSQL = sSQL & "WHERE [ID] < 0 "
        sSQL = sSQL & "AND [UpdateByUserID] = " & msUsersID & " "
        gConn.Execute sSQL, lRecordsAffected
        If lRecordsAffected > 0 Then
            lTotalRecordsAffected = lTotalRecordsAffected + lRecordsAffected
            sMess = sMess & "Removed " & lRecordsAffected & " From Table [MiscReportParam" & Format(lCount, "00") & "] " & vbCrLf
        End If
    Next
    
    'RTAttachmentsHistory
    sSQL = "DELETE FROM RTAttachmentsHistory "
    sSQL = sSQL & "WHERE RTAttachmentsID IN ( "
    sSQL = sSQL & "SELECT RTAttachmentsID FROM RTAttachments "
    sSQL = sSQL & "WHERE [ID] < 0 "
    sSQL = sSQL & "AND [UpdateByUserID] = " & msUsersID & " "
    sSQL = sSQL & ") "
    gConn.Execute sSQL, lRecordsAffected
    If lRecordsAffected > 0 Then
        lTotalRecordsAffected = lTotalRecordsAffected + lRecordsAffected
        sMess = sMess & "Removed " & lRecordsAffected & " From Table [RTAttachmentsHistory] " & vbCrLf
    End If
    
    'RTAttachments
    sSQL = "DELETE FROM RTAttachments "
    sSQL = sSQL & "WHERE [ID] < 0 "
    sSQL = sSQL & "AND [UpdateByUserID] = " & msUsersID & " "
    gConn.Execute sSQL, lRecordsAffected
    If lRecordsAffected > 0 Then
        lTotalRecordsAffected = lTotalRecordsAffected + lRecordsAffected
        sMess = sMess & "Removed " & lRecordsAffected & " From Table [RTAttachments] " & vbCrLf
    End If

    'RTPhotoLog
    
    'RTPhotolog (Null the Billingcount IDs)
    sSQL = "UPDATE RTPhotolog SET [BillingCountID] = Null, "
    sSQL = sSQL & "[IDBillingCount] = Null "
    sSQL = sSQL & "WHERE [IDBillingCount] < 0"
    sSQL = sSQL & "AND [UpdateByUserID] = " & msUsersID & " "
    gConn.Execute sSQL, lRecordsAffected
    If lRecordsAffected > 0 Then
        lTotalRecordsAffected = lTotalRecordsAffected + lRecordsAffected
        sMess = sMess & "BillingCountID Nulled " & lRecordsAffected & " ON Table [RTPhotolog] " & vbCrLf
    End If
    
    sSQL = "DELETE FROM RTPhotoLog "
    sSQL = sSQL & "WHERE ([ID] < 0 Or [IDRTPhotoReport] < 0 Or [IDBillingCount] < 0) "
    sSQL = sSQL & "AND [UpdateByUserID] = " & msUsersID & " "
    gConn.Execute sSQL, lRecordsAffected
    If lRecordsAffected > 0 Then
        lTotalRecordsAffected = lTotalRecordsAffected + lRecordsAffected
        sMess = sMess & "Removed " & lRecordsAffected & " From Table [RTPhotoLog] " & vbCrLf
    End If
    
    'RTPhotoReport
    sSQL = "DELETE FROM RTPhotoReport "
    sSQL = sSQL & "WHERE [ID] < 0 "
    sSQL = sSQL & "AND [UpdateByUserID] = " & msUsersID & " "
    gConn.Execute sSQL, lRecordsAffected
    If lRecordsAffected > 0 Then
        lTotalRecordsAffected = lTotalRecordsAffected + lRecordsAffected
        sMess = sMess & "Removed " & lRecordsAffected & " From Table [RTPhotoReport] " & vbCrLf
    End If
    
    'RTActivityLog
    'RTActivityLog (Null the Billingcount IDs)
    sSQL = "UPDATE RTActivityLog Set [BillingCountID] = Null, "
    sSQL = sSQL & "[IDBillingCount] = Null "
    sSQL = sSQL & "WHERE [IDBillingCount] < 0 "
    sSQL = sSQL & "AND [UpdateByUserID] = " & msUsersID & " "
    gConn.Execute sSQL, lRecordsAffected
    If lRecordsAffected > 0 Then
        lTotalRecordsAffected = lTotalRecordsAffected + lRecordsAffected
        sMess = sMess & "BillingCountID Nulled " & lRecordsAffected & " ON Table [RTActivityLog] " & vbCrLf
    End If
    
    sSQL = "DELETE FROM RTActivityLog "
    sSQL = sSQL & "WHERE [ID] < 0 "
    sSQL = sSQL & "AND [UpdateByUserID] = " & msUsersID & " "
    gConn.Execute sSQL, lRecordsAffected
    If lRecordsAffected > 0 Then
        lTotalRecordsAffected = lTotalRecordsAffected + lRecordsAffected
        sMess = sMess & "Removed " & lRecordsAffected & " From Table [RTActivityLog] " & vbCrLf
    End If
    
    'RTIndemnityHistory
    sSQL = "DELETE FROM RTIndemnityHistory "
    sSQL = sSQL & "WHERE [RTIndemnityID] IN ( "
    sSQL = sSQL & "SELECT [RTIndemnityID] "
    sSQL = sSQL & "FROM RTIndemnity "
    sSQL = sSQL & "WHERE [ID] < 0 "
    sSQL = sSQL & "AND [UpdateByUserID] = " & msUsersID & " "
    sSQL = sSQL & ") "
    gConn.Execute sSQL, lRecordsAffected
    If lRecordsAffected > 0 Then
        lTotalRecordsAffected = lTotalRecordsAffected + lRecordsAffected
        sMess = sMess & "Removed " & lRecordsAffected & " From Table [RTIndemnityHistory] " & vbCrLf
    End If
    
    'RTIndemnity
    'RTIndemnity (NUll RTChecksID)
    sSQL = "UPDATE RTIndemnity set [IDRTChecks] = Null, "
    sSQL = sSQL & "[RTChecksID] = Null "
    sSQL = sSQL & "WHERE  [IDRTChecks] < 0 "
    sSQL = sSQL & "AND [UpdateByUserID] = " & msUsersID & " "
    gConn.Execute sSQL, lRecordsAffected
    If lRecordsAffected > 0 Then
        lTotalRecordsAffected = lTotalRecordsAffected + lRecordsAffected
        sMess = sMess & "RTChecksID Nulled " & lRecordsAffected & " ON Table [RTIndemnity] " & vbCrLf
    End If
    
    sSQL = "DELETE FROM RTIndemnity "
    sSQL = sSQL & "WHERE [ID] < 0 "
    sSQL = sSQL & "AND [UpdateByUserID] = " & msUsersID & " "
    gConn.Execute sSQL, lRecordsAffected
    If lRecordsAffected > 0 Then
        lTotalRecordsAffected = lTotalRecordsAffected + lRecordsAffected
        sMess = sMess & "Removed " & lRecordsAffected & " From Table [RTIndemnity] " & vbCrLf
    End If
    
    
    
    'RTChecks
    sSQL = "DELETE FROM RTChecks "
    sSQL = sSQL & "WHERE ([ID] < 0 or [IDBillingCount] < 0) "
    sSQL = sSQL & "AND [UpdateByUserID] = " & msUsersID & " "
    gConn.Execute sSQL, lRecordsAffected
    If lRecordsAffected > 0 Then
        lTotalRecordsAffected = lTotalRecordsAffected + lRecordsAffected
        sMess = sMess & "Removed " & lRecordsAffected & " From Table [RTChecks] " & vbCrLf
    End If
    
    'IBFee
    sSQL = "DELETE FROM IBFee "
    sSQL = sSQL & "WHERE ([ID] < 0 Or [IDIB] < 0) "
    sSQL = sSQL & "AND [UpdateByUserID] = " & msUsersID & " "
    gConn.Execute sSQL, lRecordsAffected
    If lRecordsAffected > 0 Then
        lTotalRecordsAffected = lTotalRecordsAffected + lRecordsAffected
        sMess = sMess & "Removed " & lRecordsAffected & " From Table [IBFee] " & vbCrLf
    End If
    
    'IB
    sSQL = "DELETE FROM IB "
    sSQL = sSQL & "WHERE ([ID] < 0 or [IDBillingCount] < 0) "
    sSQL = sSQL & "AND [UpdateByUserID] = " & msUsersID & " "
    gConn.Execute sSQL, lRecordsAffected
    If lRecordsAffected > 0 Then
        lTotalRecordsAffected = lTotalRecordsAffected + lRecordsAffected
        sMess = sMess & "Removed " & lRecordsAffected & " From Table [IB] " & vbCrLf
    End If
    
    'RTIBFee
    sSQL = "DELETE FROM RTIBFee "
    sSQL = sSQL & "WHERE [ID] < 0 "
    sSQL = sSQL & "AND [UpdateByUserID] = " & msUsersID & " "
    gConn.Execute sSQL, lRecordsAffected
    If lRecordsAffected > 0 Then
        lTotalRecordsAffected = lTotalRecordsAffected + lRecordsAffected
        sMess = sMess & "Removed " & lRecordsAffected & " From Table [RTIBFee] " & vbCrLf
    End If
    
    'RTIB
    sSQL = "DELETE FROM RTIB "
    sSQL = sSQL & "WHERE [IDBillingCount] < 0 "
    sSQL = sSQL & "AND [UpdateByUserID] = " & msUsersID & " "
    gConn.Execute sSQL, lRecordsAffected
    If lRecordsAffected > 0 Then
        lTotalRecordsAffected = lTotalRecordsAffected + lRecordsAffected
        sMess = sMess & "Removed " & lRecordsAffected & " From Table [RTIB] " & vbCrLf
    End If
    
    'Billingcount
    sSQL = "DELETE FROM Billingcount "
    sSQL = sSQL & "WHERE [ID] < 0 "
    sSQL = sSQL & "AND [UpdateByUserID] = " & msUsersID & " "
    gConn.Execute sSQL, lRecordsAffected
    If lRecordsAffected > 0 Then
        lTotalRecordsAffected = lTotalRecordsAffected + lRecordsAffected
        sMess = sMess & "Removed " & lRecordsAffected & " From Table [Billingcount] " & vbCrLf
    End If
    bEndRollback = True
    If lTotalRecordsAffected > 0 Then
        plRecordsAffected = lTotalRecordsAffected
        sMess = sMess & "Total Records Affected = " & lTotalRecordsAffected & vbCrLf
        psMess = sMess
    End If
    
CLEANUP:
    'cleanup
    Set dicULUpdateFiles = Nothing
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    
    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
    sMess = sMess & "Private Function RollBackSYNCH" & vbCrLf
    sMess = sMess & "ERROR # " & lErrNum & " " & Now & vbCrLf
    sMess = sMess & sErrDesc & vbCrLf
    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
    sMess = sMess & "<------------SQL------------>" & vbCrLf
    sMess = sMess & sSQL & vbCrLf
    sMess = sMess & "<------------SQL------------>" & vbCrLf
    If bBeginRollback And Not bEndRollback Then
        Resume Next
    End If
    psMess = sMess
    Set dicULUpdateFiles = Nothing
End Function

Private Sub moUL_ErrorMess(ByVal Mess As String)
    On Error GoTo EH
    Dim lRecordsAffected As Long
    Dim sMess As String
    Dim sPassInMess As String
    
    'Check for SYNCH Roll Back
    'Pass in the current error message
    'Check for anything that needs to be rolled back
    sPassInMess = Mess
    sMess = sPassInMess
    If RollBackSYNCH(lRecordsAffected, sMess) Then
        If lRecordsAffected > 0 Then
            sMess = Mess & vbCrLf & vbCrLf & "Roll Back SYNCH " & vbCrLf & sMess
        Else
            sMess = Mess & vbCrLf & vbCrLf & "Roll Back SYNCH No Records Affected" & sMess
        End If
    Else
        sMess = Mess
    End If
    
    ErrorLog msUserName & vbCrLf & sMess
    txtMess.Text = msUserName & vbCrLf & sMess & " " & Now()
    txtMess.Refresh
    'Shut down after Error
    mbSHUTDOWN = True
    
    Exit Sub
EH:
    Err.Clear
End Sub

Private Sub moUL_UpdateDB(ByVal vBatches As Variant)
    On Error GoTo EH
    Dim sDSN As String
    Dim sMess As String
    
    sDSN = GetSetting(App.EXEName, "DSN", "NAME", "ACCESS_2000")
    
    If sDSN = "ACCESS_2000" Then
        'Not updating production table with IB, use RT instead
'        UpdateDB_Access2000 vBatches
    Else
        'CODE FOR SQL SERVER
    End If
    
    
    Exit Sub
EH:
    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
    sMess = sMess & "Private Sub moUL_UpdateDB" & vbCrLf
    sMess = sMess & "ERROR # " & Err.Number & " " & Now & vbCrLf
    sMess = sMess & Err.Description & vbCrLf
    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
    ErrorLog sMess
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
      
        Case MenuList.StopMe   'Exit
            If MsgBox("Are you sure you want to do that?", vbExclamation + vbYesNo, "Exit Process") = vbYes Then
                mbSHUTDOWN = True
            End If
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

Private Sub ImportULUpdates_SQLServer()
    On Error GoTo EH
    Dim sMess As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim sProdDSN As String
    Dim sSQL As String
    Dim lRecordsAffected As Long
    Dim oXZip As V2ECKeyBoard.clsXZip
    Dim sULUpdateFile As String
    Dim vULUPdateFile As Variant
    Dim dicULUpdateFiles As Scripting.Dictionary
    Dim lRetryTimeOut As Long
    
    
    Dim sULUpdateData As String
    
    'Check for Active Files for the current UserFolder for Upload Updates
    sULUpdateFile = Dir(msUserFolderPath & "*.zulud", vbNormal)
    
    Do Until sULUpdateFile = vbNullString
        If dicULUpdateFiles Is Nothing Then
            Set dicULUpdateFiles = New Scripting.Dictionary
        End If
        dicULUpdateFiles.Add sULUpdateFile, sULUpdateFile
        sULUpdateFile = Dir
    Loop
    
    'If the dictionary object is set that means there are some files to process
    If Not dicULUpdateFiles Is Nothing Then
        txtMess.Text = msUserName & " Table Updates Found " & Now()
        
        'Check for Adjuster currently FTP the file.
        'Active Will return true if the files are locked by the
        'FTP process.
        If Not ActiveFiles(msUserFolderPath, "*.zulud", msUserName & " Table Updates ") Then
            'Open Connection
            sProdDSN = GetSetting("V2WebControl", "DSN", "NAME", vbNullString)
            OpenConnection sProdDSN
            'Loop through each file and process it
            'check for zero length files first
            'That means connectivity issue caused a nothing file to be uploaded
            For Each vULUPdateFile In dicULUpdateFiles
                sULUpdateFile = vULUPdateFile
                 If goUtil.utGetFileData(msUserFolderPath & sULUpdateFile) = vbNullString Then
                    goUtil.utDeleteFile msUserFolderPath & "*.zulud"
                    sULUpdateFile = Left(sULUpdateFile, InStrRev(sULUpdateFile, ".", , vbBinaryCompare) - 1)
                    Err.Raise -999, , msUserName & vbCrLf & vbCrLf & sULUpdateFile & vbCrLf & vbCrLf & "Could not process ""Zero Length"" file."
                End If
            Next
            
            For Each vULUPdateFile In dicULUpdateFiles
                sULUpdateFile = vULUPdateFile
                If oXZip Is Nothing Then
                    Set oXZip = New V2ECKeyBoard.clsXZip
                End If
                'Check to be sure this is not a zero lenth file.
                'The file could be zero length if there were connection problems
                'experienced by the client while uploading the file
                'The file will be recreated by the client the next time
                'the client connects to uload download info.
                If goUtil.utGetFileData(msUserFolderPath & sULUpdateFile) = vbNullString Then
                    goUtil.utDeleteFile msUserFolderPath & "*.zulud"
                    sULUpdateFile = Left(sULUpdateFile, InStrRev(sULUpdateFile, ".", , vbBinaryCompare) - 1)
                    Err.Raise -999, , msUserName & vbCrLf & vbCrLf & sULUpdateFile & vbCrLf & vbCrLf & "Could not process ""Zero Length"" file."
                Else
                    'Unzip the file and remove the Zip file
                    oXZip.UNZipFiles msUserFolderPath, msUserFolderPath & sULUpdateFile, False
                    DoEvents
                    Sleep 500
                    goUtil.utDeleteFile msUserFolderPath & sULUpdateFile
                    
                    'change the ext for the file just unzipped
                    sULUpdateFile = Replace(sULUpdateFile, ".zulud", ".ulud", , , vbTextCompare)
                    sULUpdateData = goUtil.utGetFileData(msUserFolderPath & sULUpdateFile)
                    'Remove the File from server
                    goUtil.utDeleteFile msUserFolderPath & sULUpdateFile
                    sSQL = sULUpdateData
                    
                    'Now Execute this update againts the server
                    gConn.Execute sSQL, lRecordsAffected
                    
                    'Change Name of File for Message
                    sULUpdateFile = Left(sULUpdateFile, InStrRev(sULUpdateFile, ".", , vbBinaryCompare) - 1)
                    txtMess.Text = msUserName & vbCrLf & vbCrLf & sULUpdateFile & "... (" & lRecordsAffected & ") Records Affected " & Now()
                End If
            Next
        End If
    End If
    
    'cleanup
    Set oXZip = Nothing
    Set dicULUpdateFiles = Nothing
    Exit Sub
EH:
    'Check for TimeOut Error
    lErrNum = Err.Number
    sErrDesc = Err.Description
    'TimeOut error
    'Only try 5 times
    If lErrNum = -2147217871 Then
        lRetryTimeOut = lRetryTimeOut + 1
        If lRetryTimeOut < 5 Then
            Set gConn = Nothing
            DoEvents
            Sleep 1000
            sProdDSN = GetSetting("V2WebControl", "DSN", "NAME", vbNullString)
            OpenConnection sProdDSN
            Resume
        End If
    End If
    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
    sMess = sMess & "Private Sub ImportULUpdates_SQLServer" & vbCrLf
    sMess = sMess & "ERROR # " & lErrNum & " " & Now & vbCrLf
    sMess = sMess & sErrDesc & vbCrLf
    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
    moUL_ErrorMess sMess
    Err.Clear
    CloseConnection
End Sub


Private Sub ImportUL_SQLServer(psMyTokoutPath As String, psMyTokenData As String)
    On Error GoTo EH
    Dim sMess As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim sProdDSN As String
    Dim lRecordsAffected As Long
    Dim lRecordsAffected2 As Long
    Dim oXZip As V2ECKeyBoard.clsXZip
    Dim sULFile As String
    Dim vULFile As Variant
    Dim dicULFiles As Scripting.Dictionary
    Dim dicIDSynchServer As Scripting.Dictionary
    Dim sTickCount As String
    Dim sSYNCH As String
    Dim sTableName As String
    Dim sFileName As String
    Dim sFileNameZip As String
    Dim sULUpdateData As String
    Dim bDoSaveSynch As Boolean
    Dim saryUpdateData() As String
    Dim sSQL As String
    Dim vSQL As Variant
    Dim sTemp As String
    Dim sTemp2 As String
    Dim lPos As Long
    Dim bECUpdateBatches As Boolean
    Dim lRetryTimeOut As Long
    Dim lRetryCreateTempFile As Long
    Dim sCheckDupUL As String
    Dim sDuplicateULFile As String
    
    
    'Check for Active Files for the current UserFolder for Upload Updates
    sULFile = Dir(msUserFolderPath & "*.zul", vbNormal)
    
    Do Until sULFile = vbNullString
        If dicULFiles Is Nothing Then
            Set dicULFiles = New Scripting.Dictionary
        End If
        dicULFiles.Add sULFile, sULFile
        sULFile = Dir
    Loop
    
    'If the dictionary object is set that means there are some files to process
    If Not dicULFiles Is Nothing Then
        txtMess.Text = msUserName & " Table Upload Files Found " & Now()
        Sleep 500
        
        For Each vULFile In dicULFiles
            sULFile = vULFile
            '9.26.2005  need to check for duplicate Upload files.
            'This can happen if there was a previous session that terminated
            'on the adjusters end abnormally, or possibly the server
            sCheckDupUL = Left(sULFile, 4)
            If InStr(1, sDuplicateULFile, sCheckDupUL, vbBinaryCompare) > 0 Then
                'Besure to delete ALL other files as well
                'The entire transaction must be cleared
                goUtil.utDeleteFile msUserFolderPath & "*.zul"
                sULFile = Left(sULFile, InStrRev(sULFile, ".", , vbBinaryCompare) - 1)
                Err.Raise -999, , msUserName & vbCrLf & vbCrLf & sULFile & vbCrLf & vbCrLf & "Duplicate Upload Transaction detected.  Transaction aborted and cleared."
            End If
            sDuplicateULFile = sDuplicateULFile & sULFile & "|"
            If goUtil.utGetFileData(msUserFolderPath & sULFile) = vbNullString Then
                'Besure to delete ALL other files as well
                'The entire transaction must be cleared
                goUtil.utDeleteFile msUserFolderPath & "*.zul"
                sULFile = Left(sULFile, InStrRev(sULFile, ".", , vbBinaryCompare) - 1)
                Err.Raise -999, , msUserName & vbCrLf & vbCrLf & sULFile & vbCrLf & vbCrLf & "Could not process ""Zero Length"" file."
            End If
        Next
    
        'Check for Adjuster currently FTP the file.
        'Active Will return true if the files are locked by the
        'FTP process.
        'Open Connection
        sProdDSN = GetSetting("V2WebControl", "DSN", "NAME", vbNullString)
        OpenConnection sProdDSN
        'Loop through each file and process it
        sTickCount = goUtil.utGetTickCount
        For Each vULFile In dicULFiles
            sULFile = vULFile
            If oXZip Is Nothing Then
                Set oXZip = New V2ECKeyBoard.clsXZip
            End If
            'Check to be sure this is not a zero lenth file.
            'The file could be zero length if there were connection problems
            'experienced by the client while uploading the file
            'The file will be recreated by the client the next time
            'the client connects to upload info.
            If goUtil.utGetFileData(msUserFolderPath & sULFile) = vbNullString Then
                goUtil.utDeleteFile msUserFolderPath & "*.zul"
                sULFile = Left(sULFile, InStrRev(sULFile, ".", , vbBinaryCompare) - 1)
                Err.Raise -999, , msUserName & vbCrLf & vbCrLf & sULFile & vbCrLf & vbCrLf & "Could not process ""Zero Length"" file."
            Else
                'Unzip the file and remove the Zip file
                oXZip.UNZipFiles msUserFolderPath, msUserFolderPath & sULFile, False
'                    DoEvents
                Sleep 100
                goUtil.utDeleteFile msUserFolderPath & sULFile
                
                'change the ext for the file just unzipped
                sULFile = Replace(sULFile, ".zul", ".ul", , , vbTextCompare)
                sULUpdateData = goUtil.utGetFileData(msUserFolderPath & sULFile)
                'Remove the File from server
                goUtil.utDeleteFile msUserFolderPath & sULFile
                
                sSYNCH = vbNullString
                If Not SynchronizeULTable(sULFile, sULUpdateData, sSYNCH, lRecordsAffected, sTableName) Then
                    txtMess.Text = msUserName & " Synchronize Upload Table Failed! " & Now() & vbCrLf & sULUpdateData
                    txtMess.Refresh
                    Sleep 5000
                    GoTo CLEAN_UP
                ElseIf InStr(1, sULFile, "_IB_", vbTextCompare) = 0 Then
                    bECUpdateBatches = True
                End If
                
                '-------------------Save Synchro File------
                If sSYNCH <> vbNullString Then
                    'Save the Data to Upload Folder
                    sFileName = "SYNCH_" & sULFile & ".syc"
                    bDoSaveSynch = True
                    goUtil.utSaveFileData msUserFolderPath & sFileName, sSYNCH
                    'After Saving the Synch file need to update the Server with
                    'the [ID] Fields that will also be synched Client Side
                    Erase saryUpdateData()
                    saryUpdateData() = Split(sSYNCH, RECORD_DELIM, , vbBinaryCompare)
                    
                    For lPos = LBound(saryUpdateData, 1) To UBound(saryUpdateData, 1)
                        sSQL = saryUpdateData(lPos)
                        sSQL = Replace(sSQL, "#", "'", , , vbBinaryCompare)
                        If sSQL <> vbNullString Then
                            'Need to take the Unique ID and make it AND part of UPDATE
                            sTemp = Left(sSQL, InStr(1, sSQL, ",", vbBinaryCompare) - 1)
                            sTemp2 = Left(sSQL, InStr(1, sSQL, "SET ", vbTextCompare) + 3)
                            sTemp = Mid(sTemp, InStr(1, sTemp, "SET ", vbTextCompare) + 4)
                            sSQL = Mid(sSQL, InStr(1, sSQL, ",", vbBinaryCompare) + 1)
                            sSQL = sTemp2 & sSQL & " AND " & sTemp & " "
                        
                            'Add the Items to be Udpated After All Synch FIles
                            'have been Created ONLY !!!
                            If dicIDSynchServer Is Nothing Then
                                Set dicIDSynchServer = New Scripting.Dictionary
                            End If
                            dicIDSynchServer.Add sSQL, sFileName & "_" & CStr(lPos)
                        End If
                    Next
                End If
                '---------------------Save Synchro File------
                
                'Change Name of File for Message
                sULFile = Left(sULFile, InStrRev(sULFile, ".", , vbBinaryCompare) - 1)
                txtMess.Text = msUserName & " Table Upload" & vbCrLf & vbCrLf & sULFile & "... (" & lRecordsAffected & ") Records Affected " & Now()
            End If
            
        Next
    End If
    
    'Save the Valid Token to let client know to start Downloading
    'The Synchronize UID data
    If bDoSaveSynch Then
        'Save the multiple Sycnh files into Zip file
        sFileNameZip = "SYNCH_" & sTickCount & ".zsyc"
        oXZip.SaveZIPFiles msUserFolderPath, sFileNameZip, "*.syc", goUtil.DB_PASSWORD("1")
        SetAttr msUserFolderPath & sFileNameZip, vbNormal
        Sleep 100
    End If
    goUtil.utSaveFileData psMyTokoutPath, psMyTokenData
    
    'While the Client is Synching up , let the serve Synch too
    If Not dicIDSynchServer Is Nothing Then
        For Each vSQL In dicIDSynchServer
            sSQL = vSQL
            gConn.Execute sSQL, lRecordsAffected2
        Next
        
        'Also need to Update Batches
        If Not bECUpdateBatches Then
            'Need to check for Prev Instance flag.
            'tried to update again while working on a previous batch job.
            bECUpdateBatches = CBool(GetSetting("ECUpdateBatches", "Msg", "PrevInstance", False))
            If bECUpdateBatches Then
                SaveSetting "ECUpdateBatches", "Msg", "PrevInstance", False
            End If
        End If
        If bECUpdateBatches Then
            Shell App.Path & "\V2ECUpdateBatches.exe RunAsDepOfV2AutoImport"
        End If
    End If
    
CLEAN_UP:
    
    'cleanup
    Set oXZip = Nothing
    Set dicULFiles = Nothing
    Set dicIDSynchServer = Nothing
    Exit Sub
EH:
    'Check for TimeOut Error
    lErrNum = Err.Number
    sErrDesc = Err.Description
    'TimeOut error
    'Only try 5 times
    'Only try 5 times
    If lErrNum = -2147217871 Then
        lRetryTimeOut = lRetryTimeOut + 1
        If lRetryTimeOut < 5 Then
            Set gConn = Nothing
            DoEvents
            Sleep 1000
            sProdDSN = GetSetting("V2WebControl", "DSN", "NAME", vbNullString)
            OpenConnection sProdDSN
            Resume
        End If
    'Can't create temp file Error
    'This error will occur if there is concurrent connections competeing for the same temp file name
    'Only thing to do is try again when the same temp name the xceed sip uses is free
    'Wait 30 seconds
    ElseIf lErrNum = 504 Then
        lRetryCreateTempFile = lRetryCreateTempFile + 1
        If lRetryCreateTempFile < 30 Then
            DoEvents
            Sleep 1000
            Resume
        End If
    End If
    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
    sMess = sMess & "Private Sub ImportUL_SQLServer" & vbCrLf
    sMess = sMess & "ERROR # " & lErrNum & " " & Now & vbCrLf
    sMess = sMess & sErrDesc & vbCrLf
    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
    moUL_ErrorMess sMess
    Err.Clear
    CloseConnection
End Sub

Public Function SynchronizeULTable(psULFile As String, _
                                   psULUpdateData As String, _
                                   psSYNCH As String, _
                                   plRecordsAffected As Long, _
                                   psTableName As String) As Boolean
    'This function retruns true if SuccessFull Update of Download Info
    On Error GoTo EH
    Dim sMess As String
    Dim sData As String
    Dim saryRecords() As String
    Dim lRecordsCount As Long
    Dim lRecordsAffected As Long
    Dim lRecordsAffectedTotal As Long
    Dim saryFields() As String
    Dim lFieldsCount As Long
    Dim sFieldValue As String
    Dim sSQL As String
    Dim oField As ADODB.Field
    Dim dicDLFiles As Scripting.Dictionary
    Dim sTableName As String
    Dim RSTableDef As ADODB.Recordset
    'Need to Check for Certain Fields in
    'Assignments Table When Updating
    Dim bSkipThisRecord As Boolean
    'FK for multiple tables
    Dim lPosAssignmentsID As Long
    Dim lPosIDAssignments As Long
    Dim lPosBillingCountID As Long
    Dim lPosIDBillingCount As Long
    Dim lPosPolicyLimitsID As Long
    Dim lPosIDPolicyLimits As Long
    Dim lPosRTIBFeeID As Long
    Dim lPosIDRTIBFee As Long
    Dim lPosIBID As Long
    Dim lPosIDIB As Long
    Dim lPosIBFeeID As Long
    Dim lPosIDIBFee As Long
    Dim lPosRTChecksID As Long
    Dim lPosIDRTChecks As Long
    Dim lPosRTIndemnityID As Long
    Dim lPosIDRTIndemnity As Long
    Dim lPosRTActivityLogID As Long
    Dim lPosIDRTActivityLog As Long
    Dim lPosRTPhotoReportID As Long
    Dim lPosIDRTPhotoReport As Long
    Dim lPosRTPhotoLogID As Long
    Dim lPosIDRTPhotoLog As Long
    Dim lPosRTWSDiagramID As Long
    Dim lPosIDRTWSDiagram As Long
    Dim lPosRTAttachmentsID As Long
    Dim lPosIDRTAttachments As Long
    Dim lPosMiscReportParamID As Long
    Dim lPosIDMiscReportParam As Long
    Dim lPosPackageID As Long
    Dim lPosIDPackage As Long
    Dim lPosPackageItemID As Long
    Dim lPosIDPackageItem As Long
    'END FK for multiple tables
    Dim sSQLGetSynchFields As String
    Dim sSQLGetSynchFieldsWHERE As String
    Dim sMYIDName As String
    Dim RSGet As ADODB.Recordset
    Dim sSQLGet As String
    Dim sFieldData As String
    Dim sRecords As String
    Dim sTempData As String
    Dim RS As ADODB.Recordset
    Dim sSQLWHERE As String
    'Update SQL Server Flags and Assignments Status
    Dim sSYNCH As String
    Dim bSkipComma As Boolean
    
    Set RS = New ADODB.Recordset
    
    
    
    'Update the Table
    sTableName = Mid(psULFile, InStr(1, psULFile, "_", vbBinaryCompare) + 1)
    sTableName = Left(sTableName, InStrRev(sTableName, "_", , vbBinaryCompare) - 1)
    psTableName = sTableName
    'Set The Table Def
    sSQL = "SELECT TOP 1 * FROM " & sTableName & " "
    'Use Disconnected Record Set on asUseClient Cusor ONLY !
    Set RSTableDef = New ADODB.Recordset
    RSTableDef.CursorLocation = adUseClient
    RSTableDef.Open sSQL, gConn, adOpenForwardOnly, adLockReadOnly
    Set RSTableDef.ActiveConnection = Nothing
     
     
    'set table records to array
    sData = psULUpdateData
     
    saryRecords() = Split(sData, RECORD_DELIM, , vbBinaryCompare)
    
    'Get Pos for Unique IDs in the Current Table
    'INIT TO -1
    lPosAssignmentsID = -1
    lPosIDAssignments = -1
    lPosBillingCountID = -1
    lPosIDBillingCount = -1
    lPosPolicyLimitsID = -1
    lPosIDPolicyLimits = -1
    lPosRTIBFeeID = -1
    lPosIDRTIBFee = -1
    lPosIBID = -1
    lPosIDIB = -1
    lPosIBFeeID = -1
    lPosIDIBFee = -1
    lPosRTChecksID = -1
    lPosIDRTChecks = -1
    lPosRTIndemnityID = -1
    lPosIDRTIndemnity = -1
    lPosRTActivityLogID = -1
    lPosIDRTActivityLog = -1
    lPosRTPhotoReportID = -1
    lPosIDRTPhotoReport = -1
    lPosRTPhotoLogID = -1
    lPosIDRTPhotoLog = -1
    lPosRTWSDiagramID = -1
    lPosIDRTWSDiagram = -1
    lPosRTAttachmentsID = -1
    lPosIDRTAttachments = -1
    lPosMiscReportParamID = -1
    lPosIDMiscReportParam = -1
    lPosPackageID = -1
    lPosIDPackage = -1
    lPosPackageItemID = -1
    lPosIDPackageItem = -1
    
    lFieldsCount = 0
    For Each oField In RSTableDef.Fields
        If StrComp(oField.Name, "AssignmentsID", vbTextCompare) = 0 Then
            lPosAssignmentsID = lFieldsCount
        ElseIf StrComp(oField.Name, "IDAssignments", vbTextCompare) = 0 Then
            lPosIDAssignments = lFieldsCount
        ElseIf StrComp(oField.Name, "BillingCountID", vbTextCompare) = 0 Then
            lPosBillingCountID = lFieldsCount
        ElseIf StrComp(oField.Name, "IDBillingCount", vbTextCompare) = 0 Then
            lPosIDBillingCount = lFieldsCount
        ElseIf StrComp(oField.Name, "PolicyLimitsID", vbTextCompare) = 0 Then
            lPosPolicyLimitsID = lFieldsCount
         ElseIf StrComp(oField.Name, "IDPolicyLimits", vbTextCompare) = 0 Then
            lPosIDPolicyLimits = lFieldsCount
        ElseIf StrComp(oField.Name, "RTIBFeeID", vbTextCompare) = 0 Then
            lPosRTIBFeeID = lFieldsCount
        ElseIf StrComp(oField.Name, "IDRTIBFee", vbTextCompare) = 0 Then
            lPosIDRTIBFee = lFieldsCount
        ElseIf StrComp(oField.Name, "IBID", vbTextCompare) = 0 Then
            lPosIBID = lFieldsCount
        ElseIf StrComp(oField.Name, "IDIB", vbTextCompare) = 0 Then
            lPosIDIB = lFieldsCount
        ElseIf StrComp(oField.Name, "IBFeeID", vbTextCompare) = 0 Then
            lPosIBFeeID = lFieldsCount
        ElseIf StrComp(oField.Name, "IDIBFee", vbTextCompare) = 0 Then
            lPosIDIBFee = lFieldsCount
        ElseIf StrComp(oField.Name, "RTChecksID", vbTextCompare) = 0 Then
            lPosRTChecksID = lFieldsCount
        ElseIf StrComp(oField.Name, "IDRTChecks", vbTextCompare) = 0 Then
            lPosIDRTChecks = lFieldsCount
        ElseIf StrComp(oField.Name, "RTIndemnityID", vbTextCompare) = 0 Then
            lPosRTIndemnityID = lFieldsCount
        ElseIf StrComp(oField.Name, "IDRTIndemnity", vbTextCompare) = 0 Then
            lPosIDRTIndemnity = lFieldsCount
        ElseIf StrComp(oField.Name, "RTActivityLogID", vbTextCompare) = 0 Then
            lPosRTActivityLogID = lFieldsCount
        ElseIf StrComp(oField.Name, "IDRTActivityLog", vbTextCompare) = 0 Then
            lPosIDRTActivityLog = lFieldsCount
        ElseIf StrComp(oField.Name, "RTPhotoReportID", vbTextCompare) = 0 Then
            lPosRTPhotoReportID = lFieldsCount
        ElseIf StrComp(oField.Name, "IDRTPhotoReport", vbTextCompare) = 0 Then
            lPosIDRTPhotoReport = lFieldsCount
        ElseIf StrComp(oField.Name, "RTPhotoLogID", vbTextCompare) = 0 Then
            lPosRTPhotoLogID = lFieldsCount
        ElseIf StrComp(oField.Name, "IDRTPhotoLog", vbTextCompare) = 0 Then
            lPosIDRTPhotoLog = lFieldsCount
        ElseIf StrComp(oField.Name, "RTWSDiagramID", vbTextCompare) = 0 Then
            lPosRTWSDiagramID = lFieldsCount
        ElseIf StrComp(oField.Name, "IDRTWSDiagram", vbTextCompare) = 0 Then
            lPosIDRTWSDiagram = lFieldsCount
        ElseIf StrComp(oField.Name, "RTAttachmentsID", vbTextCompare) = 0 Then
            lPosRTAttachmentsID = lFieldsCount
        ElseIf StrComp(oField.Name, "IDRTAttachments", vbTextCompare) = 0 Then
            lPosIDRTAttachments = lFieldsCount
        ElseIf StrComp(oField.Name, "MiscReportParamID", vbTextCompare) = 0 Then
            lPosMiscReportParamID = lFieldsCount
         ElseIf StrComp(oField.Name, "IDMiscReportParam", vbTextCompare) = 0 Then
            lPosIDMiscReportParam = lFieldsCount
        ElseIf StrComp(oField.Name, "PackageID", vbTextCompare) = 0 Then
            lPosPackageID = lFieldsCount
        ElseIf StrComp(oField.Name, "IDPackage", vbTextCompare) = 0 Then
            lPosIDPackage = lFieldsCount
        ElseIf StrComp(oField.Name, "PackageItemID", vbTextCompare) = 0 Then
            lPosPackageItemID = lFieldsCount
        ElseIf StrComp(oField.Name, "IDPackageItem", vbTextCompare) = 0 Then
            lPosIDPackageItem = lFieldsCount
        End If
        lFieldsCount = lFieldsCount + 1
    Next
    
     
    For lRecordsCount = LBound(saryRecords, 1) To UBound(saryRecords, 1)
        sData = saryRecords(lRecordsCount)
         
        If sData = vbNullString Then
            GoTo SKIP_THIS_RECORD
        End If
        
        saryFields() = Split(sData, COLUMN_DELIM, , vbBinaryCompare)
        lFieldsCount = 0
        
        'Build Update SQL
        sSQL = "SELECT * FROM " & sTableName & " "
        sSQLWHERE = "WHERE "
        For Each oField In RSTableDef.Fields
            'Get the Unique ID... All tables have the Unique ID in the very
            'First Column.  In the future this could Change per table
            'So the Code is set up incase of this eventuality
            sFieldValue = saryFields(lFieldsCount)
            sFieldValue = Replace(sFieldValue, COLUMN_DELIM_REP, COLUMN_DELIM, , , vbBinaryCompare)
            sFieldValue = Replace(sFieldValue, RECORD_DELIM_REP, RECORD_DELIM, , , vbBinaryCompare)
            Select Case UCase(sTableName)
                Case UCase("")
                Case Else
                    If lFieldsCount > 0 Then
                        Exit For
                    End If
                    sSQLWHERE = sSQLWHERE & "[" & oField.Name & "] = " & sFieldValue & " "
            End Select
            lFieldsCount = lFieldsCount + 1
         Next
    
        sSQL = sSQL & sSQLWHERE
        
        Set RS = Nothing
        Set RS = New ADODB.Recordset
        'Use Disconnected Record Set on asUseClient Cusor ONLY !
        RS.CursorLocation = adUseClient
        RS.Open sSQL, gConn, adOpenForwardOnly, adLockReadOnly
        Set RS.ActiveConnection = Nothing
         
                   
        sSQLGetSynchFields = vbNullString
        sSQLGetSynchFieldsWHERE = vbNullString
        
        
         'If the record ID does not exist , then Insert the Upload Record
         
         If RS.RecordCount = 0 Then
            lFieldsCount = 0
            'Build Update SQL
            sSQL = "INSERT INTO " & sTableName & " "
            sSQL = sSQL & "SELECT "
            
            'Need to Build Sql to Return updated Uniue ID for Sych with Client
            sSQLGetSynchFields = "SELECT "
            
            For Each oField In RSTableDef.Fields
                If lFieldsCount = 0 Then
                    'The very first field will always be the
                    'Unique ID for this Table. Since inserting a new
                    'Record must leave this Field out inorder to
                    'Seen a unique ID
                    sSQLGetSynchFields = sSQLGetSynchFields & "[" & oField.Name & "], "
                    sMYIDName = oField.Name
                    If StrComp(sTableName, "Assignments", vbTextCompare) <> 0 Then
                        If StrComp(sTableName, "RTIB", vbTextCompare) <> 0 Then
                            If StrComp(sTableName, "RTActivityLogInfo", vbTextCompare) <> 0 Then
                                GoTo SKIP_THIS_FIELD
                            End If
                        End If
                    End If
                ElseIf lFieldsCount = 1 Then
                    If StrComp(sTableName, "Assignments", vbTextCompare) = 0 Then
                        sSQL = sSQL & ", "
                    ElseIf StrComp(sTableName, "RTIB", vbTextCompare) = 0 Then
                        sSQL = sSQL & ", "
                    ElseIf StrComp(sTableName, "RTActivityLogInfo", vbTextCompare) = 0 Then
                        sSQL = sSQL & ", "
                    Else
                        sSQL = sSQL
                    End If
                ElseIf lFieldsCount > 1 Then
                    'Add Comma to start new Column
                    If bSkipComma Then
                        bSkipComma = False
                    Else
                        sSQL = sSQL & ", "
                    End If
                End If
                
                sFieldValue = saryFields(lFieldsCount)
                sFieldValue = Replace(sFieldValue, COLUMN_DELIM_REP, COLUMN_DELIM, , , vbBinaryCompare)
                sFieldValue = Replace(sFieldValue, RECORD_DELIM_REP, RECORD_DELIM, , , vbBinaryCompare)
               
                'Snag and or Modify Field values Here if applicable
                'For All Tables
                Select Case UCase(oField.Name)
                    Case UCase("AssignmentsID")
                        If StrComp(sTableName, "Assignments", vbTextCompare) <> 0 Then
                            If StrComp(sTableName, "RTIB", vbTextCompare) <> 0 Then
                                If StrComp(sTableName, "RTActivityLogInfo", vbTextCompare) <> 0 Then
                                    sSQLGetSynchFields = sSQLGetSynchFields & "[AssignmentsID] As [AssignmentsID], "
                                End If
                            End If
                        End If
                    Case UCase("IDAssignments")
                        sSQLGetSynchFields = sSQLGetSynchFields & "[AssignmentsID] As [IDAssignments], "
                        If sSQLGetSynchFieldsWHERE = vbNullString Then
                            sSQLGetSynchFieldsWHERE = "WHERE [IDAssignments] = " & sFieldValue & " "
                        Else
                            sSQLGetSynchFieldsWHERE = sSQLGetSynchFieldsWHERE & "AND [IDAssignments] = " & sFieldValue & " "
                        End If
                    Case UCase("ID")
                        sSQLGetSynchFields = sSQLGetSynchFields & "[" & sMYIDName & "] As [ID], "
                        If sSQLGetSynchFieldsWHERE = vbNullString Then
                            sSQLGetSynchFieldsWHERE = "WHERE [ID] = " & sFieldValue & " "
                        Else
                            sSQLGetSynchFieldsWHERE = sSQLGetSynchFieldsWHERE & "AND [ID] = " & sFieldValue & " "
                        End If
                    Case UCase("UpLoadMe")
                        sSQLGetSynchFields = sSQLGetSynchFields & "0 As [UpLoadMe], "
                    Case UCase("DateLastUpdated")
                        'Need to change the Date lastupdaetd to use server Date
                        'This should always be the last field in the Get Synch Fields
                        sSQLGetSynchFields = sSQLGetSynchFields & "[DateLastUpdated] As [DateLastUpdated] "
                    
                End Select
                Select Case UCase(sTableName)
                    Case UCase("Assignments")
                        Select Case UCase(oField.Name)
                            Case UCase("UploadLossReport")
                                sFieldValue = "0" 'Set the UploadLossReport Flag to 0
                        End Select
                    Case UCase("BillingCount")
                        'Do Nothing here
                    Case UCase("PolicyLimits")
                        'Do Nothing here
                    Case UCase("RTIB"), UCase("IB"), UCase("RTChecks"), UCase("RTActivityLog"), UCase("RTPhotoLog")
                        'Need to Get the Updated BillingCountID
                        Select Case UCase(oField.Name)
                            Case UCase("BillingCountID")
                                sSQLGet = "SELECT [BillingCountID] "
                                sSQLGet = sSQLGet & "FROM BillingCount "
                                sSQLGet = sSQLGet & "WHERE [AssignmentsID] = " & saryFields(lPosAssignmentsID) & " "
                                If StrComp(saryFields(lPosIDBillingCount), "IS_NULL", vbTextCompare) = 0 Then
                                    sSQLGet = sSQLGet & "AND [ID] Is Null "
                                Else
                                    sSQLGet = sSQLGet & "AND [ID] = " & saryFields(lPosIDBillingCount) & " "
                                End If
                                
                                Set RSGet = New ADODB.Recordset
                                RSGet.CursorLocation = adUseClient
                                RSGet.Open sSQLGet, gConn, adOpenForwardOnly, adLockReadOnly
                                Set RSGet.ActiveConnection = Nothing
                                If RSGet.RecordCount = 1 Then
                                    RSGet.MoveFirst
                                    sFieldValue = RSGet.Fields("BillingCountID")
                                    sSQLGetSynchFields = sSQLGetSynchFields & "[BillingCountID] As [BillingCountID], "
                                    sSQLGetSynchFields = sSQLGetSynchFields & "[BillingCountID] As [IDBillingCount], "
                                Else
                                    sFieldValue = "IS_NULL"
                                    sSQLGetSynchFields = sSQLGetSynchFields & "Null As [BillingCountID], "
                                    sSQLGetSynchFields = sSQLGetSynchFields & "Null As [IDBillingCount], "
                                End If
                            Case UCase("RTPhotoReportID")
                                sSQLGet = "SELECT [RTPhotoReportID] "
                                sSQLGet = sSQLGet & "FROM RTPhotoReport "
                                sSQLGet = sSQLGet & "WHERE [AssignmentsID] = " & saryFields(lPosAssignmentsID) & " "
                                sSQLGet = sSQLGet & "AND [ID] = " & saryFields(lPosIDRTPhotoReport) & " "
                                Set RSGet = New ADODB.Recordset
                                RSGet.CursorLocation = adUseClient
                                RSGet.Open sSQLGet, gConn, adOpenForwardOnly, adLockReadOnly
                                Set RSGet.ActiveConnection = Nothing
                                If RSGet.RecordCount = 1 Then
                                    RSGet.MoveFirst
                                    sFieldValue = RSGet.Fields("RTPhotoReportID")
                                    sSQLGetSynchFields = sSQLGetSynchFields & "[RTPhotoReportID] As [RTPhotoReportID], "
                                    sSQLGetSynchFields = sSQLGetSynchFields & "[RTPhotoReportID] As [IDRTPhotoReport], "
                                End If
                        End Select
                    Case UCase("RTIBFee")
                        'Do Nothing here
                    Case UCase("IBFee")
                        Select Case UCase(oField.Name)
                            Case UCase("IBID")
                                sSQLGet = "SELECT [IBID] "
                                sSQLGet = sSQLGet & "FROM IB "
                                sSQLGet = sSQLGet & "WHERE [AssignmentsID] = " & saryFields(lPosAssignmentsID) & " "
                                sSQLGet = sSQLGet & "AND [ID] = " & saryFields(lPosIDIB) & " "
                                Set RSGet = New ADODB.Recordset
                                RSGet.CursorLocation = adUseClient
                                RSGet.Open sSQLGet, gConn, adOpenForwardOnly, adLockReadOnly
                                Set RSGet.ActiveConnection = Nothing
                                If RSGet.RecordCount = 1 Then
                                    RSGet.MoveFirst
                                    sFieldValue = RSGet.Fields("IBID")
                                    sSQLGetSynchFields = sSQLGetSynchFields & "[IBID] As [IBID], "
                                    sSQLGetSynchFields = sSQLGetSynchFields & "[IBID] As [IDIB], "
                                End If
                        End Select
                    Case UCase("RTIndemnity")
                        Select Case UCase(oField.Name)
                            Case UCase("RTChecksID")
                                sSQLGet = "SELECT [RTChecksID] "
                                sSQLGet = sSQLGet & "FROM RTChecks "
                                sSQLGet = sSQLGet & "WHERE [AssignmentsID] = " & saryFields(lPosAssignmentsID) & " "
                                If StrComp(saryFields(lPosIDRTChecks), "IS_NULL", vbTextCompare) = 0 Then
                                    sSQLGet = sSQLGet & "AND [ID] Is Null "
                                Else
                                    sSQLGet = sSQLGet & "AND [ID] = " & saryFields(lPosIDRTChecks) & " "
                                End If
                                
                                Set RSGet = New ADODB.Recordset
                                RSGet.CursorLocation = adUseClient
                                RSGet.Open sSQLGet, gConn, adOpenForwardOnly, adLockReadOnly
                                Set RSGet.ActiveConnection = Nothing
                                If RSGet.RecordCount = 1 Then
                                    RSGet.MoveFirst
                                    sFieldValue = RSGet.Fields("RTChecksID")
                                    sSQLGetSynchFields = sSQLGetSynchFields & "[RTChecksID] As [RTChecksID], "
                                    sSQLGetSynchFields = sSQLGetSynchFields & "[RTChecksID] As [IDRTChecks], "
                                Else
                                    sFieldValue = "IS_NULL"
                                    sSQLGetSynchFields = sSQLGetSynchFields & "Null As [RTChecksID], "
                                    sSQLGetSynchFields = sSQLGetSynchFields & "Null As [IDRTChecks], "
                                End If
                        End Select
                    Case UCase("RTActivityLogInfo")
                        'Do Nothing here
                    Case UCase("RTPhotoReport")
                        'Do Nothing here
                    Case UCase("RTWSDiagram")
                        'Do Nothing here
                    Case UCase("RTAttachments")
                        'Do Nothing here
                    Case UCase("MiscReportParam")
                        'Do Nothing here
                        '2/8/2004 MiscReportParam , and MiscReportParam01 to MiscReportParam30
                    Case UCase("Package")
                        'Do Nothing here
                    Case UCase("PackageItem")
                        Select Case UCase(oField.Name)
                            Case UCase("PackageID")
                                sSQLGet = "SELECT [PackageID] "
                                sSQLGet = sSQLGet & "FROM Package "
                                sSQLGet = sSQLGet & "WHERE [AssignmentsID] = " & saryFields(lPosAssignmentsID) & " "
                                sSQLGet = sSQLGet & "AND [ID] = " & saryFields(lPosIDPackage) & " "
                                Set RSGet = New ADODB.Recordset
                                RSGet.CursorLocation = adUseClient
                                RSGet.Open sSQLGet, gConn, adOpenForwardOnly, adLockReadOnly
                                Set RSGet.ActiveConnection = Nothing
                                If RSGet.RecordCount = 1 Then
                                    RSGet.MoveFirst
                                    sFieldValue = RSGet.Fields("PackageID")
                                    sSQLGetSynchFields = sSQLGetSynchFields & "[PackageID] As [PackageID], "
                                    sSQLGetSynchFields = sSQLGetSynchFields & "[PackageID] As [IDPackage], "
                                End If
                            Case UCase("RTAttachmentsID")
                                sSQLGet = "SELECT [RTAttachmentsID] "
                                sSQLGet = sSQLGet & "FROM RTAttachments "
                                sSQLGet = sSQLGet & "WHERE [AssignmentsID] = " & saryFields(lPosAssignmentsID) & " "
                                If StrComp(saryFields(lPosIDRTAttachments), "IS_NULL", vbTextCompare) = 0 Then
                                    sSQLGet = sSQLGet & "AND [ID] Is Null "
                                Else
                                    sSQLGet = sSQLGet & "AND [ID] = " & saryFields(lPosIDRTAttachments) & " "
                                End If
                                
                                Set RSGet = New ADODB.Recordset
                                RSGet.CursorLocation = adUseClient
                                RSGet.Open sSQLGet, gConn, adOpenForwardOnly, adLockReadOnly
                                Set RSGet.ActiveConnection = Nothing
                                If RSGet.RecordCount = 1 Then
                                    RSGet.MoveFirst
                                    sFieldValue = RSGet.Fields("RTAttachmentsID")
                                    sSQLGetSynchFields = sSQLGetSynchFields & "[RTAttachmentsID] As [RTAttachmentsID], "
                                    sSQLGetSynchFields = sSQLGetSynchFields & "[RTAttachmentsID] As [IDRTAttachments], "
                                Else
                                    sFieldValue = "IS_NULL"
                                    sSQLGetSynchFields = sSQLGetSynchFields & "Null As [RTAttachmentsID], "
                                    sSQLGetSynchFields = sSQLGetSynchFields & "Null As [IDRTAttachments], "
                                End If
                        End Select
                End Select
                
                If StrComp(sFieldValue, "IS_NULL", vbTextCompare) = 0 Then
                    sSQL = sSQL & "null "
                Else
                     Select Case oField.Type
                        Case ADODB.DataTypeEnum.adBinary
                            sFieldValue = Replace(sFieldValue, "'", "''", , , vbBinaryCompare)
                            If sFieldValue = vbNullString Then
                                sFieldValue = " "
                            End If
                            sSQL = sSQL & "'" & sFieldValue & "'"
                        Case ADODB.DataTypeEnum.adBoolean
                            'Reset UPLoadMe Flags
                            If StrComp(oField.Name, "UpLoadMe", vbTextCompare) = 0 Then
                                sSQL = sSQL & " 0"
                            ElseIf StrComp(oField.Name, "UpLoadAll", vbTextCompare) = 0 Then
                                sSQL = sSQL & " 0"
                            Else
                                If CBool(sFieldValue) Then
                                    sSQL = sSQL & " -1"
                                Else
                                    sSQL = sSQL & " 0"
                                End If
                            End If
                        Case ADODB.DataTypeEnum.adVarBinary, ADODB.DataTypeEnum.adLongVarBinary, ADODB.DataTypeEnum.adVarChar, ADODB.DataTypeEnum.adWChar, ADODB.DataTypeEnum.adChar, ADODB.DataTypeEnum.adLongVarChar, ADODB.DataTypeEnum.adLongVarWChar
                            'The PackageItemGUID is created on SQL Server Only
                            'This Item should never be inserted or Updated in this process
                            If StrComp(oField.Name, "PackageItemGUID", vbTextCompare) = 0 Then
                                sFieldValue = " (newid()) "
                                sSQL = sSQL & sFieldValue
                            Else
                                sFieldValue = Replace(sFieldValue, "'", "''", , , vbBinaryCompare)
                                If sFieldValue = vbNullString Then
                                    sFieldValue = " "
                                End If
                                sSQL = sSQL & "'" & sFieldValue & "'"
                            End If
                            
                        Case ADODB.DataTypeEnum.adDate, ADODB.DataTypeEnum.adDBDate, ADODB.DataTypeEnum.adDBTime, ADODB.DataTypeEnum.adDBTimeStamp
                            Select Case UCase(oField.Name)
                                Case UCase("DateLastUpdated")
                                    'Need to change the Date lastupdaetd to use server Date
                                    sSQL = sSQL & "GetDate()"
                                Case Else
                                    sSQL = sSQL & "'" & sFieldValue & "'"
                            End Select
                        Case ADODB.DataTypeEnum.adNumeric, ADODB.DataTypeEnum.adInteger, ADODB.DataTypeEnum.adDecimal, ADODB.DataTypeEnum.adSingle, ADODB.DataTypeEnum.adDouble, ADODB.DataTypeEnum.adCurrency, ADODB.DataTypeEnum.adBigInt
                            sSQL = sSQL & sFieldValue
                        Case ADODB.DataTypeEnum.adGUID
                            sFieldValue = Replace(sFieldValue, "'", "''", , , vbBinaryCompare)
                            If sFieldValue = vbNullString Then
                                sFieldValue = " "
                            End If
                            sSQL = sSQL & "'" & sFieldValue & "'"
                        Case Else
                            sSQL = sSQL & sFieldValue
                    End Select
                End If
                sSQL = sSQL & " As [" & oField.Name & "] "
SKIP_THIS_FIELD:
                lFieldsCount = lFieldsCount + 1
            Next 'oField In oTableDef.Fields
            'Insert this Record Into the Table
             'Get Synch Values
            
           gConn.Execute sSQL, lRecordsAffected
            lRecordsAffectedTotal = lRecordsAffectedTotal + lRecordsAffected
            
            If lRecordsAffected = 1 Then
                sSQLGetSynchFields = sSQLGetSynchFields & "FROM " & sTableName & " "
                sSQLGetSynchFields = sSQLGetSynchFields & sSQLGetSynchFieldsWHERE & " "
                Set RSGet = New ADODB.Recordset
                RSGet.CursorLocation = adUseClient
                RSGet.Open sSQLGetSynchFields, gConn, adOpenForwardOnly, adLockReadOnly
                Set RSGet.ActiveConnection = Nothing
                If RSGet.RecordCount = 1 Then
                    sRecords = vbNullString
                    RSGet.MoveFirst
                    Do Until RSGet.EOF
                        sFieldData = vbNullString
                        For Each oField In RSGet.Fields
                            If sFieldData = vbNullString Then
                                sFieldData = "UPDATE " & sTableName & " SET "
                            Else
                                sFieldData = sFieldData & ", "
                            End If
                            If IsNull(oField.Value) Then
                                sFieldData = sFieldData & "[" & oField.Name & "] = NULL "
                            Else
                                Select Case oField.Type
                                    Case ADODB.DataTypeEnum.adWChar, ADODB.DataTypeEnum.adVarWChar, ADODB.DataTypeEnum.adVarChar, ADODB.DataTypeEnum.adLongVarWChar, ADODB.DataTypeEnum.adLongVarChar, ADODB.DataTypeEnum.adChar, ADODB.DataTypeEnum.adBSTR
                                        sTempData = CStr(oField.Value)
                                        sTempData = Replace(sTempData, "'", "''", , , vbBinaryCompare)
                                        sTempData = Replace(sTempData, RECORD_DELIM, RECORD_DELIM_REP, , , vbBinaryCompare)
                                        sFieldData = sFieldData & "[" & oField.Name & "] = '" & sTempData & "' "
                                    Case ADODB.DataTypeEnum.adDate, ADODB.DataTypeEnum.adDBDate, ADODB.DataTypeEnum.adDBTime, ADODB.DataTypeEnum.adDBTimeStamp
                                        sFieldData = sFieldData & "[" & oField.Name & "] = #" & CStr(oField.Value) & "# "
                                    Case Else
                                        sFieldData = sFieldData & "[" & oField.Name & "] = " & CStr(oField.Value) & " "
                                End Select
                            End If
                        Next
                        RSGet.MoveNext
                        sFieldData = sFieldData & sSQLGetSynchFieldsWHERE
                        sFieldData = sFieldData & RECORD_DELIM
                        sRecords = sRecords & sFieldData
                    Loop
                    sSYNCH = sSYNCH & sRecords
                End If
            End If
         End If
         
         'If the Record Does Exists then Update the Download record
         If RS.RecordCount = 1 Then
            bSkipThisRecord = False
            'Skip the record if the Server has this item
            'flagged in the following Select case
            Select Case UCase(sTableName)
                Case UCase("Assignments")
                    'Reassigned
                    'Since the Client GUI makes the user unable to edit Reassigned items
                    'AFTER THEY DOWNLOAD THE REASSIGNED FLAG
                    'Taking this option out... Just incase the Server needs to do a
                    'FILE ATTACHMENT RESTORE, which would Require the Client to be able to
                    'Send updates to REASSIGNED ITEMS!
'                    sFieldValue = goUtil.IsNullIsVbNullString(RS.Fields("Reassigned"))
'                    If CBool(sFieldValue) Then
'                        bSkipThisRecord = True
'                    End If
                    'IsLocked
                    sFieldValue = goUtil.IsNullIsVbNullString(RS.Fields("IsLocked"))
                    If CBool(sFieldValue) Then
'                        bSkipThisRecord = True
                    End If
            End Select
            'Skip it if flagged
            If bSkipThisRecord Then
                GoTo SKIP_THIS_RECORD
            End If
        
            lFieldsCount = 0
            'Build Update SQL
            sSQL = "UPDATE " & sTableName & " SET "
            
            'Need to Build Sql to Return updated Uniue ID for Sych with Client
            sSQLGetSynchFields = "SELECT "
            
            For Each oField In RSTableDef.Fields
                If lFieldsCount = 0 Then
                    'The very first field will always be the
                    'Unique ID for this Table. Since inserting a new
                    'Record must leave this Field out inorder to
                    'Seen a unique ID
                    sSQLGetSynchFields = sSQLGetSynchFields & "[" & oField.Name & "], "
                    sMYIDName = oField.Name
                    If StrComp(sTableName, "Assignments", vbTextCompare) <> 0 Then
                        If StrComp(sTableName, "RTIB", vbTextCompare) <> 0 Then
                            If StrComp(sTableName, "RTActivityLogInfo", vbTextCompare) <> 0 Then
                                GoTo SKIP_THIS_FIELD2
                            End If
                        End If
                    End If
                ElseIf lFieldsCount = 1 Then
                    If StrComp(sTableName, "Assignments", vbTextCompare) = 0 Then
                        sSQL = sSQL & ", "
                    ElseIf StrComp(sTableName, "RTIB", vbTextCompare) = 0 Then
                        sSQL = sSQL & ", "
                    ElseIf StrComp(sTableName, "RTActivityLogInfo", vbTextCompare) = 0 Then
                        sSQL = sSQL & ", "
                    Else
                        sSQL = sSQL
                    End If
                ElseIf lFieldsCount > 1 Then
                    'Add Comma to start new Column
                    If bSkipComma Then
                        bSkipComma = False
                    Else
                        sSQL = sSQL & ", "
                    End If
                End If
                
                sSQL = sSQL & "[" & oField.Name & "] = "
                sFieldValue = saryFields(lFieldsCount)
                sFieldValue = Replace(sFieldValue, COLUMN_DELIM_REP, COLUMN_DELIM, , , vbBinaryCompare)
                sFieldValue = Replace(sFieldValue, RECORD_DELIM_REP, RECORD_DELIM, , , vbBinaryCompare)
                
                'Snag and or Modify Field values Here if applicable
                'For All Tables
                Select Case UCase(oField.Name)
                    Case UCase("AssignmentsID")
                        If StrComp(sTableName, "Assignments", vbTextCompare) <> 0 Then
                            If StrComp(sTableName, "RTIB", vbTextCompare) <> 0 Then
                                If StrComp(sTableName, "RTActivityLogInfo", vbTextCompare) <> 0 Then
                                    sSQLGetSynchFields = sSQLGetSynchFields & "[AssignmentsID] As [AssignmentsID], "
                                End If
                            End If
                        End If
                    Case UCase("IDAssignments")
                        sSQLGetSynchFields = sSQLGetSynchFields & "[AssignmentsID] As [IDAssignments], "
                        If sSQLGetSynchFieldsWHERE = vbNullString Then
                            sSQLGetSynchFieldsWHERE = "WHERE [IDAssignments] = " & sFieldValue & " "
                        Else
                            sSQLGetSynchFieldsWHERE = sSQLGetSynchFieldsWHERE & "AND [IDAssignments] = " & sFieldValue & " "
                        End If
                    Case UCase("ID")
                        sSQLGetSynchFields = sSQLGetSynchFields & "[" & sMYIDName & "] As [ID], "
                        If sSQLGetSynchFieldsWHERE = vbNullString Then
                            sSQLGetSynchFieldsWHERE = "WHERE [ID] = " & sFieldValue & " "
                        Else
                            sSQLGetSynchFieldsWHERE = sSQLGetSynchFieldsWHERE & "AND [ID] = " & sFieldValue & " "
                        End If
                    Case UCase("UpLoadMe")
                        sSQLGetSynchFields = sSQLGetSynchFields & "0 As [UpLoadMe], "
                    Case UCase("DateLastUpdated")
                        'Need to change the Date lastupdaetd to use server Date
                        'This should always be the last field in the Get Synch Fields
                        sSQLGetSynchFields = sSQLGetSynchFields & "[DateLastUpdated] As [DateLastUpdated] "
                    
                End Select
                Select Case UCase(sTableName)
                    Case UCase("Assignments")
                        Select Case UCase(oField.Name)
                            Case UCase("UploadLossReport")
                                sFieldValue = "0" 'Set the UploadLossReport Flag to 0
                        End Select
                    Case UCase("BillingCount")
                        'Do Nothing here
                    Case UCase("PolicyLimits")
                        'Do Nothing here
                    Case UCase("RTIB"), UCase("IB"), UCase("RTChecks"), UCase("RTActivityLog"), UCase("RTPhotoLog")
                        'Need to Get the Updated BillingCountID
                        Select Case UCase(oField.Name)
                            Case UCase("BillingCountID")
                                sSQLGet = "SELECT [BillingCountID] "
                                sSQLGet = sSQLGet & "FROM BillingCount "
                                sSQLGet = sSQLGet & "WHERE [AssignmentsID] = " & saryFields(lPosAssignmentsID) & " "
                                If StrComp(saryFields(lPosIDBillingCount), "IS_NULL", vbTextCompare) = 0 Then
                                    sSQLGet = sSQLGet & "AND [ID] Is Null "
                                Else
                                    sSQLGet = sSQLGet & "AND [ID] = " & saryFields(lPosIDBillingCount) & " "
                                End If
                                
                                Set RSGet = New ADODB.Recordset
                                RSGet.CursorLocation = adUseClient
                                RSGet.Open sSQLGet, gConn, adOpenForwardOnly, adLockReadOnly
                                Set RSGet.ActiveConnection = Nothing
                                If RSGet.RecordCount = 1 Then
                                    RSGet.MoveFirst
                                    sFieldValue = RSGet.Fields("BillingCountID")
                                    sSQLGetSynchFields = sSQLGetSynchFields & "[BillingCountID] As [BillingCountID], "
                                    sSQLGetSynchFields = sSQLGetSynchFields & "[BillingCountID] As [IDBillingCount], "
                                Else
                                    sFieldValue = "IS_NULL"
                                    sSQLGetSynchFields = sSQLGetSynchFields & "Null As [BillingCountID], "
                                    sSQLGetSynchFields = sSQLGetSynchFields & "Null As [IDBillingCount], "
                                End If
                            Case UCase("RTPhotoReportID")
                                sSQLGet = "SELECT [RTPhotoReportID] "
                                sSQLGet = sSQLGet & "FROM RTPhotoReport "
                                sSQLGet = sSQLGet & "WHERE [AssignmentsID] = " & saryFields(lPosAssignmentsID) & " "
                                sSQLGet = sSQLGet & "AND [ID] = " & saryFields(lPosIDRTPhotoReport) & " "
                                Set RSGet = New ADODB.Recordset
                                RSGet.CursorLocation = adUseClient
                                RSGet.Open sSQLGet, gConn, adOpenForwardOnly, adLockReadOnly
                                Set RSGet.ActiveConnection = Nothing
                                If RSGet.RecordCount = 1 Then
                                    RSGet.MoveFirst
                                    sFieldValue = RSGet.Fields("RTPhotoReportID")
                                    sSQLGetSynchFields = sSQLGetSynchFields & "[RTPhotoReportID] As [RTPhotoReportID], "
                                    sSQLGetSynchFields = sSQLGetSynchFields & "[RTPhotoReportID] As [IDRTPhotoReport], "
                                End If
                        End Select
                    Case UCase("RTIBFee")
                        'Do Nothing here
                    Case UCase("IBFee")
                        Select Case UCase(oField.Name)
                            Case UCase("IBID")
                                sSQLGet = "SELECT [IBID] "
                                sSQLGet = sSQLGet & "FROM IB "
                                sSQLGet = sSQLGet & "WHERE [AssignmentsID] = " & saryFields(lPosAssignmentsID) & " "
                                sSQLGet = sSQLGet & "AND [ID] = " & saryFields(lPosIDIB) & " "
                                Set RSGet = New ADODB.Recordset
                                RSGet.CursorLocation = adUseClient
                                RSGet.Open sSQLGet, gConn, adOpenForwardOnly, adLockReadOnly
                                Set RSGet.ActiveConnection = Nothing
                                If RSGet.RecordCount = 1 Then
                                    RSGet.MoveFirst
                                    sFieldValue = RSGet.Fields("IBID")
                                    sSQLGetSynchFields = sSQLGetSynchFields & "[IBID] As [IBID], "
                                    sSQLGetSynchFields = sSQLGetSynchFields & "[IBID] As [IDIB], "
                                End If
                        End Select
                    Case UCase("RTIndemnity")
                        Select Case UCase(oField.Name)
                            Case UCase("RTChecksID")
                                sSQLGet = "SELECT [RTChecksID] "
                                sSQLGet = sSQLGet & "FROM RTChecks "
                                sSQLGet = sSQLGet & "WHERE [AssignmentsID] = " & saryFields(lPosAssignmentsID) & " "
                                If StrComp(saryFields(lPosIDRTChecks), "IS_NULL", vbTextCompare) = 0 Then
                                    sSQLGet = sSQLGet & "AND [ID] Is Null "
                                Else
                                    sSQLGet = sSQLGet & "AND [ID] = " & saryFields(lPosIDRTChecks) & " "
                                End If
                                
                                Set RSGet = New ADODB.Recordset
                                RSGet.CursorLocation = adUseClient
                                RSGet.Open sSQLGet, gConn, adOpenForwardOnly, adLockReadOnly
                                Set RSGet.ActiveConnection = Nothing
                                If RSGet.RecordCount = 1 Then
                                    RSGet.MoveFirst
                                    sFieldValue = RSGet.Fields("RTChecksID")
                                    sSQLGetSynchFields = sSQLGetSynchFields & "[RTChecksID] As [RTChecksID], "
                                    sSQLGetSynchFields = sSQLGetSynchFields & "[RTChecksID] As [IDRTChecks], "
                                Else
                                    sFieldValue = "IS_NULL"
                                    sSQLGetSynchFields = sSQLGetSynchFields & "Null As [RTChecksID], "
                                    sSQLGetSynchFields = sSQLGetSynchFields & "Null As [IDRTChecks], "
                                End If
                        End Select
                    Case UCase("RTActivityLogInfo")
                        'Do Nothing here
                    Case UCase("RTPhotoReport")
                        'Do Nothing here
                    Case UCase("RTWSDiagram")
                        'Do Nothing here
                    Case UCase("RTAttachments")
                        'Do Nothing here
                    Case UCase("MiscReportParam")
                        '2/8/2004 MiscReportParam , and MiscReportParam01 to MiscReportParam30
                        'Do Nothing here
                    Case UCase("Package")
                        'Do Nothing here
                    Case UCase("PackageItem")
                        Select Case UCase(oField.Name)
                            Case UCase("PackageID")
                                sSQLGet = "SELECT [PackageID] "
                                sSQLGet = sSQLGet & "FROM Package "
                                sSQLGet = sSQLGet & "WHERE [AssignmentsID] = " & saryFields(lPosAssignmentsID) & " "
                                sSQLGet = sSQLGet & "AND [ID] = " & saryFields(lPosIDPackage) & " "
                                Set RSGet = New ADODB.Recordset
                                RSGet.CursorLocation = adUseClient
                                RSGet.Open sSQLGet, gConn, adOpenForwardOnly, adLockReadOnly
                                Set RSGet.ActiveConnection = Nothing
                                If RSGet.RecordCount = 1 Then
                                    RSGet.MoveFirst
                                    sFieldValue = RSGet.Fields("PackageID")
                                    sSQLGetSynchFields = sSQLGetSynchFields & "[PackageID] As [PackageID], "
                                    sSQLGetSynchFields = sSQLGetSynchFields & "[PackageID] As [IDPackage], "
                                End If
                            Case UCase("RTAttachmentsID")
                                sSQLGet = "SELECT [RTAttachmentsID] "
                                sSQLGet = sSQLGet & "FROM RTAttachments "
                                sSQLGet = sSQLGet & "WHERE [AssignmentsID] = " & saryFields(lPosAssignmentsID) & " "
                                If StrComp(saryFields(lPosIDRTAttachments), "IS_NULL", vbTextCompare) = 0 Then
                                    sSQLGet = sSQLGet & "AND [ID] Is Null "
                                Else
                                    sSQLGet = sSQLGet & "AND [ID] = " & saryFields(lPosIDRTAttachments) & " "
                                End If
                                
                                Set RSGet = New ADODB.Recordset
                                RSGet.CursorLocation = adUseClient
                                RSGet.Open sSQLGet, gConn, adOpenForwardOnly, adLockReadOnly
                                Set RSGet.ActiveConnection = Nothing
                                If RSGet.RecordCount = 1 Then
                                    RSGet.MoveFirst
                                    sFieldValue = RSGet.Fields("RTAttachmentsID")
                                    sSQLGetSynchFields = sSQLGetSynchFields & "[RTAttachmentsID] As [RTAttachmentsID], "
                                    sSQLGetSynchFields = sSQLGetSynchFields & "[RTAttachmentsID] As [IDRTAttachments], "
                                Else
                                    sFieldValue = "IS_NULL"
                                    sSQLGetSynchFields = sSQLGetSynchFields & "Null As [RTAttachmentsID], "
                                    sSQLGetSynchFields = sSQLGetSynchFields & "Null As [IDRTAttachments], "
                                End If
                        End Select
                End Select
                
                If StrComp(sFieldValue, "IS_NULL", vbTextCompare) = 0 Then
                    sSQL = sSQL & "null"
                Else
                    Select Case oField.Type
                        Case ADODB.DataTypeEnum.adBinary
                            sFieldValue = Replace(sFieldValue, "'", "''", , , vbBinaryCompare)
                            If sFieldValue = vbNullString Then
                                sFieldValue = " "
                            End If
                            sSQL = sSQL & "'" & sFieldValue & "'"
                        Case ADODB.DataTypeEnum.adBoolean
                            'Reset UPLoadMe Flags
                            If StrComp(oField.Name, "UpLoadMe", vbTextCompare) = 0 Then
                                sSQL = sSQL & " 0"
                            ElseIf StrComp(oField.Name, "UpLoadAll", vbTextCompare) = 0 Then
                                sSQL = sSQL & " 0"
                            Else
                                If CBool(sFieldValue) Then
                                    sSQL = sSQL & " -1"
                                Else
                                    sSQL = sSQL & " 0"
                                End If
                            End If
                        Case ADODB.DataTypeEnum.adVarBinary, ADODB.DataTypeEnum.adLongVarBinary, ADODB.DataTypeEnum.adVarChar, ADODB.DataTypeEnum.adWChar, ADODB.DataTypeEnum.adChar, ADODB.DataTypeEnum.adLongVarChar, ADODB.DataTypeEnum.adLongVarWChar
                            'The PackageItemGUID is created on SQL Server Only
                            'This Item should never be inserted or Updated in this process
                            If StrComp(oField.Name, "PackageItemGUID", vbTextCompare) = 0 Then
                                'Use the exisiting [PackageItemGUID] value on update statements !!!
                                sFieldValue = " ([PackageItemGUID]) "
                                sSQL = sSQL & sFieldValue
                            Else
                                sFieldValue = Replace(sFieldValue, "'", "''", , , vbBinaryCompare)
                                If sFieldValue = vbNullString Then
                                    sFieldValue = " "
                                End If
                                sSQL = sSQL & "'" & sFieldValue & "'"
                            End If
                            
                        Case ADODB.DataTypeEnum.adDate, ADODB.DataTypeEnum.adDBDate, ADODB.DataTypeEnum.adDBTime, ADODB.DataTypeEnum.adDBTimeStamp
                            Select Case UCase(oField.Name)
                                Case UCase("DateLastUpdated")
                                    'Need to change the Date lastupdaetd to use server Date
                                    sSQL = sSQL & "GetDate()"
                                Case Else
                                    sSQL = sSQL & "'" & sFieldValue & "'"
                            End Select
                        Case ADODB.DataTypeEnum.adNumeric, ADODB.DataTypeEnum.adInteger, ADODB.DataTypeEnum.adDecimal, ADODB.DataTypeEnum.adSingle, ADODB.DataTypeEnum.adDouble, ADODB.DataTypeEnum.adCurrency, ADODB.DataTypeEnum.adBigInt
                            sSQL = sSQL & sFieldValue
                        Case ADODB.DataTypeEnum.adGUID
                            sFieldValue = Replace(sFieldValue, "'", "''", , , vbBinaryCompare)
                            If sFieldValue = vbNullString Then
                                sFieldValue = " "
                            End If
                            sSQL = sSQL & "'" & sFieldValue & "'"
                        Case Else
                            sSQL = sSQL & sFieldValue
                    End Select
                End If
SKIP_THIS_FIELD2:
                lFieldsCount = lFieldsCount + 1
            Next 'oField In oTableDef.Fields
                'Insert this Record Into the Table
            sSQL = sSQL & " " & sSQLWHERE
                
            gConn.Execute sSQL, lRecordsAffected
            lRecordsAffectedTotal = lRecordsAffectedTotal + lRecordsAffected
            
            If lRecordsAffected = 1 Then
                sSQLGetSynchFields = sSQLGetSynchFields & "FROM " & sTableName & " "
                sSQLGetSynchFields = sSQLGetSynchFields & sSQLGetSynchFieldsWHERE & " "
                Set RSGet = New ADODB.Recordset
                RSGet.CursorLocation = adUseClient
                RSGet.Open sSQLGetSynchFields, gConn, adOpenForwardOnly, adLockReadOnly
                Set RSGet.ActiveConnection = Nothing
                If RSGet.RecordCount = 1 Then
                    sRecords = vbNullString
                    RSGet.MoveFirst
                    Do Until RSGet.EOF
                        sFieldData = vbNullString
                        For Each oField In RSGet.Fields
                            If sFieldData = vbNullString Then
                                sFieldData = "UPDATE " & sTableName & " SET "
                            Else
                                sFieldData = sFieldData & ", "
                            End If
                            If IsNull(oField.Value) Then
                                sFieldData = sFieldData & "[" & oField.Name & "] = NULL "
                            Else
                                Select Case oField.Type
                                    Case ADODB.DataTypeEnum.adWChar, ADODB.DataTypeEnum.adVarWChar, ADODB.DataTypeEnum.adVarChar, ADODB.DataTypeEnum.adLongVarWChar, ADODB.DataTypeEnum.adLongVarChar, ADODB.DataTypeEnum.adChar, ADODB.DataTypeEnum.adBSTR
                                        sTempData = CStr(oField.Value)
                                        sTempData = Replace(sTempData, "'", "''", , , vbBinaryCompare)
                                        sTempData = Replace(sTempData, RECORD_DELIM, RECORD_DELIM_REP, , , vbBinaryCompare)
                                        sFieldData = sFieldData & "[" & oField.Name & "] = '" & sTempData & "' "
                                    Case ADODB.DataTypeEnum.adDate, ADODB.DataTypeEnum.adDBDate, ADODB.DataTypeEnum.adDBTime, ADODB.DataTypeEnum.adDBTimeStamp
                                        sFieldData = sFieldData & "[" & oField.Name & "] = #" & CStr(oField.Value) & "# "
                                    Case Else
                                        sFieldData = sFieldData & "[" & oField.Name & "] = " & CStr(oField.Value) & " "
                                End Select
                            End If
                        Next
                        RSGet.MoveNext
                        sFieldData = sFieldData & sSQLGetSynchFieldsWHERE
                        sFieldData = sFieldData & RECORD_DELIM
                        sRecords = sRecords & sFieldData
                    Loop
                    sSYNCH = sSYNCH & sRecords
                End If
            End If
            
        End If
SKIP_THIS_RECORD:
    Next 'lRecordsCount
    
    
    If lRecordsAffectedTotal > 0 Then
        plRecordsAffected = lRecordsAffectedTotal
        psSYNCH = sSYNCH
        SynchronizeULTable = True
    End If
    
    
    'cleanup
    Set RSTableDef = Nothing
    Set oField = Nothing
    Set RS = Nothing
    Set RSGet = Nothing

    
    Exit Function
EH:
    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
    sMess = sMess & "Public Function SynchronizeULTable" & vbCrLf
    sMess = sMess & "ERROR # " & Err.Number & " " & Now & vbCrLf
    sMess = sMess & Err.Description & vbCrLf & vbCrLf
    sMess = sMess & "<---------------------SQL--------------------->" & vbCrLf
    sMess = sMess & sSQL & vbCrLf
    sMess = sMess & "<---------------------SQL--------------------->" & vbCrLf
    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
    moUL_ErrorMess sMess
    Err.Clear
    CloseConnection
End Function

Private Sub ExportDL_SQLServer(psMyTokoutPath As String, psMyTokenData As String)
    On Error GoTo EH
    Dim sSQL As String
    Dim sMess As String
    Dim sProdDSN As String
    Dim sTableRSData As String
    Dim sTickCount As String
    Dim sFileName As String
    Dim sFileNameZip As String
    Dim sDLFilePath As String
    Dim sTableName As String
    Dim saryDBLookUpTables(1 To 200) As String
    Dim lCountDBLookupTables As Long
    Dim saryDBTables(1 To 500) As String
    Dim lCountDBTables As Long
    Dim oXZip As V2ECKeyBoard.clsXZip
    Dim sDLFile As String
    Dim sListDLFile As String
    Dim lCheckCount As Long
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim lRetryCount As Long
    Dim sRetryErrDesc As String
    
    sProdDSN = GetSetting("V2WebControl", "DSN", "NAME", vbNullString)
    OpenConnection sProdDSN
    Set oXZip = New V2ECKeyBoard.clsXZip
    
    'A. Need to Create DL files that Easy Claim will Download to Update Look up
    'Info specific to User and User Assigned Cat Info.  This will include Software Information.
    'Add the Look up tables to the lookup table array
    saryDBLookUpTables(1) = "DB_VERSION"
    saryDBLookUpTables(2) = "SoftwarePackageRegSetting"
    saryDBLookUpTables(3) = "RegSetting"
    saryDBLookUpTables(4) = "RegSettingHistory"
    saryDBLookUpTables(5) = "SoftwarePackageDocument"
    saryDBLookUpTables(6) = "Document"
    saryDBLookUpTables(7) = "DocumentHistory"
    saryDBLookUpTables(8) = "SoftwarePackageApplication"
    saryDBLookUpTables(9) = "Application"
    saryDBLookUpTables(10) = "ApplicationHistory"
    saryDBLookUpTables(11) = "SoftwarePackage"
    saryDBLookUpTables(12) = "SoftwarePackageHistory"
    saryDBLookUpTables(13) = "ClientCompanyUsersCat"
    saryDBLookUpTables(14) = "ClientCompanyCat"
    saryDBLookUpTables(15) = "CAT"
    saryDBLookUpTables(16) = "ClientCompanyCatSpec"
    saryDBLookUpTables(17) = "Company"
    saryDBLookUpTables(18) = "CompanyUsers"
    saryDBLookUpTables(19) = "Users"
    saryDBLookUpTables(20) = "AdjusterUsersSoftware"
    saryDBLookUpTables(21) = "AdjusterUsersUpdates"
    saryDBLookUpTables(22) = "Adjuster"
    saryDBLookUpTables(23) = "ClientCoAdjusterSpec"
    saryDBLookUpTables(24) = "FeeSchedule"
    saryDBLookUpTables(25) = "FeeScheduleFeeTypes"
    saryDBLookUpTables(26) = "FeeScheduleLevels"
    saryDBLookUpTables(27) = "TypeOfLoss"
    saryDBLookUpTables(28) = "ClassType"
    saryDBLookUpTables(29) = "ClassOfLoss"
    saryDBLookUpTables(30) = "State"
    saryDBLookUpTables(31) = "AssignmentType"
    saryDBLookUpTables(32) = "Status"

    'A.Loop through the Lookup Tables
    
    For lCountDBLookupTables = LBound(saryDBLookUpTables, 1) To UBound(saryDBLookUpTables, 1)
        sTableName = saryDBLookUpTables(lCountDBLookupTables)
        If sTableName <> vbNullString Then
            sSQL = "SELECT * FROM " & sTableName & " "
            'Build the Where Statement
            
            Select Case sTableName
                Case "SoftwarePackageRegSetting", "RegSetting", "RegSettingHistory"
                    sSQL = sSQL & "WHERE IsDeleted = 0 "
                Case "SoftwarePackageDocument", "Document", "DocumentHistory"
                    sSQL = sSQL & "WHERE IsDeleted = 0 "
                Case "SoftwarePackageApplication", "Application", "ApplicationHistory"
                    sSQL = sSQL & "WHERE IsDeleted = 0 "
                Case "SoftwarePackage", "SoftwarePackageHistory"
                    sSQL = sSQL & "WHERE IsDeleted = 0 "
                Case "ClientCompanyUsersCat"
                    sSQL = sSQL & "WHERE UsersID = " & msUsersID & " "
                    sSQL = sSQL & "AND  Active = 1 "
                Case "ClientCompanyCat", "ClientCompanyCatSpec"
                    sSQL = sSQL & "WHERE ClientCompanyID IN ( "
                                                sSQL = sSQL & "SELECT   ClientCompanyID "
                                                sSQL = sSQL & "FROM     ClientCompanyUsersCat "
                                                sSQL = sSQL & "WHERE UsersID = " & msUsersID & " "
                                                sSQL = sSQL & ") "
                    sSQL = sSQL & "AND CATID IN ( "
                                                sSQL = sSQL & "SELECT   CATID "
                                                sSQL = sSQL & "FROM     ClientCompanyUsersCat "
                                                sSQL = sSQL & "WHERE UsersID = " & msUsersID & " "
                                                sSQL = sSQL & ") "
                Case "CAT"
                     sSQL = sSQL & "WHERE CATID IN ( "
                                                sSQL = sSQL & "SELECT   CATID "
                                                sSQL = sSQL & "FROM     ClientCompanyUsersCat "
                                                sSQL = sSQL & "WHERE UsersID = " & msUsersID & " "
                                                sSQL = sSQL & ") "
                Case "Company"
                    sSQL = sSQL & "WHERE Active = 1 "
                Case "CompanyUsers", "Users", "AdjusterUsersSoftware", "AdjusterUsersUpdates", "Adjuster", "ClientCoAdjusterSpec"
                    sSQL = sSQL & "WHERE UsersID = " & msUsersID & " "
                    If sTableName = "Users" Then
                        sSQL = sSQL & "AND  Active = 1 "
                    End If
                Case "FeeSchedule", "TypeOfLoss", "ClassOfLoss"
                    sSQL = sSQL & "WHERE ClientCompanyID IN ( "
                                            sSQL = sSQL & "SELECT   ClientCompanyID "
                                            sSQL = sSQL & "FROM     ClientCompanyUsersCat "
                                            sSQL = sSQL & "WHERE UsersID = " & msUsersID & " "
                                            sSQL = sSQL & ") "
                    sSQL = sSQL & "AND IsDeleted = 0 "
                Case "FeeScheduleFeeTypes", "FeeScheduleLevels"
                    sSQL = sSQL & "WHERE FeeScheduleID IN ( "
                                            sSQL = sSQL & "SELECT   FeeScheduleID "
                                            sSQL = sSQL & "FROM     FeeSchedule "
                                            sSQL = sSQL & "WHERE ClientCompanyID IN ( "
                                                                    sSQL = sSQL & "SELECT   ClientCompanyID "
                                                                    sSQL = sSQL & "FROM     ClientCompanyUsersCat "
                                                                    sSQL = sSQL & "WHERE UsersID = " & msUsersID & " "
                                                                    sSQL = sSQL & ") "
                                            sSQL = sSQL & "AND IsDeleted = 0 "
                                            sSQL = sSQL & ") "
                    sSQL = sSQL & "AND IsDeleted = 0 "
                Case "ClassType", "State", "AssignmentType", "Status"
                    
            End Select
            
            sFileName = sTableName & ".dllu"
            sFileNameZip = sTableName & ".zdllu"
            sDLFilePath = msUserFolderPath
            sTableRSData = GetExportDLTableRS(sSQL)
            
            'Besure that this particular Lookup File is Replaced with the latest
            goUtil.utDeleteFile sDLFilePath & sFileNameZip
            goUtil.utDeleteFile sDLFilePath & sFileName
            
            goUtil.utSaveFileData sDLFilePath & sFileName, sTableRSData
            If sTableName = "DB_VERSION" Then
                'Do not Zip and Encrypt the DB Version Table
                SetAttr sDLFilePath & sFileName, vbNormal
                'Wait for Flag from CLient that The DB Version is up to
                'Date on Client. if not then need to Exit this Process
                txtMess.Text = "Client is checking database version " & Now()
                txtMess.Refresh
                'Allow for the Downloading of DB_VErsion File
                lCheckCount = 0
                Sleep 2000
CHECK_DBFile:
                sListDLFile = vbNullString
                lCheckCount = lCheckCount + 1
                sDLFile = Dir(sDLFilePath & "DB_VERSION.dllu", vbNormal)
                Do While sDLFile <> vbNullString
                    sListDLFile = sListDLFile & sDLFile & vbCrLf
                    sDLFile = Dir
                Loop
                
                If sListDLFile <> vbNullString Then
                    sListDLFile = "Client is checking database version " & Now() & vbCrLf & sListDLFile
                    txtMess.Text = sListDLFile
                    txtMess.Refresh
                    DoEvents
                    Sleep 1000
                    ProcessTokens
                    'Only keep on checking this for a timeout 1 hour
                    If lCheckCount > 30 Or mbSHUTDOWN Then
                        txtMess.Text = "Timeout waiting for Client to download DB_VERSION files"
                        txtMess.Refresh
                        moUL_ErrorMess txtMess.Text
                        DoEvents
                        Sleep 1000
                        mbSHUTDOWN = True
                        GoTo CLEAN_UP
                    Else
                        GoTo CHECK_DBFile
                    End If
                End If
                If Not ClientFlagCheck("DBUpToDate.flag") Then
                    txtMess.Text = "Client database update required.  Exiting process. " & Now()
                    txtMess.Refresh
                    Sleep 1000
                    DoEvents
                    GoTo CLEAN_UP
                End If
                txtMess.Text = txtMess.Text & vbCrLf & "Client database is up to date. " & Now()
                txtMess.Refresh
                Sleep 1000
                DoEvents
            Else
                SetAttr sDLFilePath & sFileName, vbNormal
                '3.26.2005
                'Because Zip Utility will create a temporary file under the Temp directory
                'when it is generating a zip file under the specified directory...
                'It is possible that multiple instances of Process User Tokin will try to
                'Create a temp file with the same file name.  When this happens the SaveFiles
                'function will return an error , 504 cannot create temporary file.
                'This error can be used to indicate to the process receiving the error that it needs
                'to sleep, wait for the other process to finish up.  This will also help to
                'govern the process load on the server.
                lRetryCount = 0
                oXZip.SaveZIPFiles sDLFilePath, sFileNameZip, sFileName, goUtil.DB_PASSWORD("1")
                SetAttr sDLFilePath & sFileNameZip, vbNormal
            End If
            
        Else
            Exit For
        End If
    Next
    
     'Save the Valid Token to let client know
    'to start checking Software Versions
    goUtil.utSaveFileData psMyTokoutPath, psMyTokenData
    
    'Wait for Flag from CLient that Software Is Up to
    'Date on Client. if not then need to Exit this Process
    txtMess.Text = "Client is checking software " & Now()
    txtMess.Refresh
    'Allow for the Downloading of Lookup Files
    lCheckCount = 0
CHECK_DLFiles:
    sListDLFile = vbNullString
    lCheckCount = lCheckCount + 1
    sDLFile = Dir(sDLFilePath & "*.zdllu", vbNormal)
    Do While sDLFile <> vbNullString
        sListDLFile = sListDLFile & sDLFile & vbCrLf
        sDLFile = Dir
    Loop
    
    If sListDLFile <> vbNullString Then
        sListDLFile = "Client is checking software " & Now() & vbCrLf & sListDLFile
        txtMess.Text = sListDLFile
        txtMess.Refresh
        DoEvents
        Sleep 1000
        ProcessTokens
        'Only keep on checking this for a timeout 1 hour
        If lCheckCount > 1800 Or mbSHUTDOWN Then
            txtMess.Text = "Timeout waiting for Client to download Look Up files"
            txtMess.Refresh
            moUL_ErrorMess txtMess.Text
            DoEvents
            Sleep 1000
            mbSHUTDOWN = True
            GoTo CLEAN_UP
        Else
            GoTo CHECK_DLFiles
        End If
    End If
    
    'Check for Software Up to date flag
    If Not ClientFlagCheck("SoftwareUpToDate.flag") Then
        txtMess.Text = "Client software update required.  Exiting process. " & Now()
        txtMess.Refresh
        Sleep 1000
        DoEvents
        GoTo CLEAN_UP
    End If
    txtMess.Text = txtMess.Text & vbCrLf & "Client software is up to date. " & Now()
    txtMess.Refresh
    Sleep 1000
    DoEvents
    
   'B. Need to Create DL files that Easy Claim will Download to Update Look up
    'Info specific to User and User Assigned Cat Info.  This will include Software Information.
    'Add the Look up tables to the lookup table array
    saryDBTables(1) = "Assignments"
    saryDBTables(2) = "BillingCount"
    saryDBTables(3) = "PolicyLimits"
    saryDBTables(4) = "RTIB"
    saryDBTables(5) = "RTIBFee"
    saryDBTables(6) = "IB"
    saryDBTables(7) = "IBFee"
    saryDBTables(8) = "RTChecks"
    saryDBTables(9) = "RTIndemnity"
    saryDBTables(10) = "RTActivityLog"
    saryDBTables(11) = "RTActivityLogInfo"
    saryDBTables(12) = "RTPhotoReport"
    saryDBTables(13) = "RTPhotoLog"
    saryDBTables(14) = "RTWSDiagram"
    saryDBTables(15) = "RTAttachments"
    saryDBTables(16) = "Package"
    saryDBTables(17) = "PackageItem"
    '2/8/2004 MiscReportParam , and MiscReportParam01 to MiscReportParam30
    saryDBTables(18) = "MiscReportParam"
    saryDBTables(19) = "MiscReportParam01"
    saryDBTables(20) = "MiscReportParam02"
    saryDBTables(21) = "MiscReportParam03"
    saryDBTables(22) = "MiscReportParam04"
    saryDBTables(23) = "MiscReportParam05"
    saryDBTables(24) = "MiscReportParam06"
    saryDBTables(25) = "MiscReportParam07"
    saryDBTables(26) = "MiscReportParam08"
    saryDBTables(27) = "MiscReportParam09"
    saryDBTables(28) = "MiscReportParam10"
    saryDBTables(29) = "MiscReportParam11"
    saryDBTables(30) = "MiscReportParam12"
    saryDBTables(31) = "MiscReportParam13"
    saryDBTables(32) = "MiscReportParam14"
    saryDBTables(33) = "MiscReportParam15"
    saryDBTables(34) = "MiscReportParam16"
    saryDBTables(35) = "MiscReportParam17"
    saryDBTables(36) = "MiscReportParam18"
    saryDBTables(37) = "MiscReportParam19"
    saryDBTables(38) = "MiscReportParam20"
    saryDBTables(39) = "MiscReportParam21"
    saryDBTables(40) = "MiscReportParam22"
    saryDBTables(41) = "MiscReportParam23"
    saryDBTables(42) = "MiscReportParam24"
    saryDBTables(43) = "MiscReportParam25"
    saryDBTables(44) = "MiscReportParam26"
    saryDBTables(45) = "MiscReportParam27"
    saryDBTables(46) = "MiscReportParam28"
    saryDBTables(47) = "MiscReportParam29"
    saryDBTables(48) = "MiscReportParam30"
    'B. Get DataBase Updates Here
     
    'Use Tick Count on Non Look up info
    sTickCount = goUtil.utGetTickCount
    For lCountDBTables = LBound(saryDBTables, 1) To UBound(saryDBTables, 1)
        sTableName = saryDBTables(lCountDBTables)
        
        If sTableName = vbNullString Then
            Exit For
        End If
        
        sSQL = "SELECT * FROM " & sTableName & " "
        
        'Build the Where Statement
        Select Case sTableName
            Case "Assignments"
                sSQL = sSQL & "WHERE AdjusterSpecID IN( "
                                        sSQL = sSQL & "SELECT   ClientCoAdjusterSpecID "
                                        sSQL = sSQL & "FROM     ClientCoAdjusterSpec "
                                        sSQL = sSQL & "WHERE    UsersID = " & msUsersID & " "
                                        sSQL = sSQL & ") "
                sSQL = sSQL & "AND DownLoadMe = 1 "
                
            Case Else
                '"RTIB", "RTIBFee", "IB", "IBFee", "BillingCount",
                '"PolicyLimits", "Package", "PackageItem",
                '"RTChecks", "RTIndemnity", "RTActivityLog",
                '"RTActivityLogInfo", "RTPhotoLog", "RTPhotoReport",
                '"RTWSDiagram", "RTAttachments", "RTFarmerNCC",
                '2/8/2004 MiscReportParam , and MiscReportParam01 to MiscReportParam30
                '"MiscReportParam"
                sSQL = sSQL & "WHERE AssignmentsID IN ("
                                        sSQL = sSQL & "SELECT   AssignmentsID "
                                        sSQL = sSQL & "FROM     Assignments "
                                        sSQL = sSQL & "WHERE    AdjusterSpecID IN( "
                                                                sSQL = sSQL & "SELECT   ClientCoAdjusterSpecID "
                                                                sSQL = sSQL & "FROM     ClientCoAdjusterSpec "
                                                                sSQL = sSQL & "WHERE    UsersID = " & msUsersID & " "
                                                                sSQL = sSQL & ") "
                                         sSQL = sSQL & ") "
                sSQL = sSQL & "AND DownLoadMe = 1 "
                
        End Select
        
        sFileName = sTableName & "_" & sTickCount & ".dl"
        sFileNameZip = sTableName & "_" & sTickCount & ".zdl"
        sDLFilePath = msUserFolderPath
        sTableRSData = GetExportDLTableRS(sSQL)
        
        'If there is No data don't create Download
        If sTableRSData <> vbNullString Then
            goUtil.utSaveFileData sDLFilePath & sFileName, sTableRSData
            lRetryCount = 0
            oXZip.SaveZIPFiles sDLFilePath, sFileNameZip, sFileName, goUtil.DB_PASSWORD("1")
            SetAttr sDLFilePath & sFileNameZip, vbNormal
        End If
    Next
    
    
    'Save the Valid Token to let client know
    'to start Downloading Data
    goUtil.utSaveFileData psMyTokoutPath, psMyTokenData
    
    txtMess.Text = "Client is Downloading Data " & Now()
    txtMess.Refresh
    'Allow for the Downloading of Lookup Files
    lCheckCount = 0
CHECK_DLFiles2:
    sListDLFile = vbNullString
    lCheckCount = lCheckCount + 1
    sDLFile = Dir(sDLFilePath & "*.zdl", vbNormal)
    Do While sDLFile <> vbNullString
        sListDLFile = sListDLFile & sDLFile & vbCrLf
        sDLFile = Dir
    Loop
    
    If sListDLFile <> vbNullString Then
        sListDLFile = "Client is Downloading Data " & Now() & vbCrLf & sListDLFile
        txtMess.Text = sListDLFile
        txtMess.Refresh
        DoEvents
        Sleep 1000
        ProcessTokens
        'Only keep on checking this for a timeout 1 hour
        If lCheckCount > 3600 Or mbSHUTDOWN Then
            txtMess.Text = "Timeout waiting for Client to download Data files"
            txtMess.Refresh
            moUL_ErrorMess txtMess.Text
            DoEvents
            Sleep 1000
            mbSHUTDOWN = True
            GoTo CLEAN_UP
        Else
            GoTo CHECK_DLFiles2
        End If
    End If
    
CLEAN_UP:

    Set oXZip = Nothing
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    'Ceck for Can't create Temp file errors (504)...
    'and cannot move temporary file to target zip file(508)...
    'these are an indication of abusy server and Zip utility trying to
    'create duplicate temp files, for multiple adjusters connecting at the same time.
    'since the zip utility is apparently using the temp directory to create temp files internally.
    If lErrNum = 504 Or lErrNum = 508 Then
        If lRetryCount < 30 Then
            lRetryCount = lRetryCount + 1
            sRetryErrDesc = sRetryErrDesc & vbCrLf & "ERROR # " & lErrNum & " " & Now & vbCrLf & sErrDesc
            sRetryErrDesc = sRetryErrDesc & vbCrLf & "Retry Count = " & lRetryCount
            txtMess.Text = sErrDesc & vbCrLf & "Retry Count = " & lRetryCount
            DoEvents
            Sleep 1000
            Resume
        End If
    End If
    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
    sMess = sMess & "Private Sub ExportDL_SQLServer" & vbCrLf
    sMess = sMess & "ERROR # " & lErrNum & " " & Now & vbCrLf
    sMess = sMess & sErrDesc & vbCrLf
    If sRetryErrDesc <> vbNullString Then
        sMess = sMess & "Retry Errors: " & vbCrLf
        sMess = sMess & sRetryErrDesc & vbCrLf
    End If
    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
    moUL_ErrorMess sMess
    Err.Clear
    CloseConnection
End Sub

Private Function ClientFlagCheck(psFlagName As String, _
                                Optional plAttempts As Long, _
                                Optional plMaxAttempts As Long = 120) As Boolean
    On Error GoTo EH
    Dim sMess As String
    Dim sFlag As String
    Dim sFlagData As String
    Dim iFFile As Integer
    Dim lAttempts As Long
    
    lAttempts = plAttempts + 1
    'wait for 120 seconds (by default) , 120 attmepts at 1 second each
    If lAttempts > plMaxAttempts Then
        ClientFlagCheck = False
        Exit Function
    End If
    
    txtMess.Text = txtMess.Text & "."
    txtMess.Refresh
    DoEvents
    Sleep 1000
    
    sFlag = Dir(msUserFolderPath & "\" & psFlagName, vbNormal)
    
    'Be sure that sFlag is not null string
    'If it is it means that it is taking longer than a second for
    'the client to upload a 4 byte file !!!
    If sFlag = vbNullString Then
        GoTo ATTEMP_FLAG_CHECK
    Else
        sFlagData = goUtil.utGetFileData(msUserFolderPath & "\" & sFlag)
        If sFlagData = vbNullString Then
            GoTo ATTEMP_FLAG_CHECK
        End If
    End If
    
    'Check to see if this file is still active
    iFFile = FreeFile
    On Error Resume Next
    Open msUserFolderPath & "\" & sFlag For Binary Access Read Lock Read As #iFFile
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo EH
        If mbSHUTDOWN Then
            Exit Function
        End If
ATTEMP_FLAG_CHECK:
        ClientFlagCheck = ClientFlagCheck(psFlagName, lAttempts, plMaxAttempts)
        Exit Function
    End If
    On Error GoTo EH
    Close #iFFile
    
    'get the Flag Data
    sFlagData = goUtil.utGetFileData(msUserFolderPath & "\" & sFlag)
    'Delete the flag from the Server
    goUtil.utDeleteFile msUserFolderPath & "\" & sFlag
    
    ClientFlagCheck = CBool(sFlagData)
    
    Exit Function
EH:
    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
    sMess = sMess & "Private Function ClientFlagCheck" & vbCrLf
    sMess = sMess & "ERROR # " & Err.Number & " " & Now & vbCrLf
    sMess = sMess & Err.Description & vbCrLf
    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
    moUL_ErrorMess sMess
    Err.Clear
End Function

Private Function GetExportDLTableRS(psSQL As String) As String
    On Error GoTo EH
    Dim sMess As String
    Dim RS As ADODB.Recordset
    Dim oField As ADODB.Field
    Dim sFieldData As String
    Dim sRecords As String
    Dim sTempData As String
    
    Set RS = New ADODB.Recordset
    'Use Disconnected Record Set on asUseClient Cusor ONLY !
    
    RS.CursorLocation = adUseClient
    RS.Open psSQL, gConn, adOpenForwardOnly, adLockReadOnly
    Set RS.ActiveConnection = Nothing
    
    If Not RS.EOF Then
        RS.MoveFirst
        Do Until RS.EOF
            sFieldData = vbNullString
            For Each oField In RS.Fields
                'Clear previous Value
                If IsNull(oField.Value) Then
                    sFieldData = sFieldData & "IS_NULL"
                Else
                    Select Case oField.Type
                        Case ADODB.DataTypeEnum.adWChar, ADODB.DataTypeEnum.adVarWChar, ADODB.DataTypeEnum.adVarChar, ADODB.DataTypeEnum.adLongVarWChar, ADODB.DataTypeEnum.adLongVarChar, ADODB.DataTypeEnum.adChar, ADODB.DataTypeEnum.adBSTR
                            sTempData = CStr(oField.Value)
                            sTempData = Replace(sTempData, COLUMN_DELIM, COLUMN_DELIM_REP, , , vbBinaryCompare)
                            sTempData = Replace(sTempData, RECORD_DELIM, RECORD_DELIM_REP, , , vbBinaryCompare)
                            sFieldData = sFieldData & sTempData
                        Case Else
                            sFieldData = sFieldData & CStr(oField.Value)
                    End Select
                End If
                sFieldData = sFieldData & COLUMN_DELIM
            Next
            RS.MoveNext
            sFieldData = sFieldData & RECORD_DELIM
            sRecords = sRecords & sFieldData
        Loop
    End If
    
    GetExportDLTableRS = sRecords
    
    'cleanup
    Set RS = Nothing
    
    
    Exit Function
EH:
    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
    sMess = sMess & "Private Function GetExportDLTableRS" & vbCrLf
    sMess = sMess & "ERROR # " & Err.Number & " " & Now & vbCrLf
    sMess = sMess & Err.Description & vbCrLf
    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
    moUL_ErrorMess sMess
    Err.Clear
End Function

'Private Sub UpdateDBRT_SQLServer(ByVal pvBatches As Variant, poBatch As V2ECKeyBoard.clsBatches, poRTUL As Object, poUL As V2ECKeyBoard.clsUpload, _
'                                  Optional pbDeleted As Boolean)
''    On Error GoTo EH
''    Dim sSQL As String
''    Dim sMess As String
''    Dim sSQLError As String
''    Dim MyBat As V2ECKeyBoard.udtBatchesRT
''    Dim sDBName As String
''    Dim sTemp As String
''    Dim cRTTotalFee As Currency
''    Dim RS As ADODB.Recordset
''    Dim lID As Long
''    Dim sIBNumberPF As String
''    Dim lSFPos As Long  'Suffix "." posn
''    Dim lRSFPos As Long 'Suffix "R" Rebill posn
''    Dim lSSFPos As Long 'Suffic "S" Supplement posn
''    Dim sProdDSN As String
''    Dim lRecordsAffected As Long
''
''    sProdDSN = GetSetting("V2AutoImport", "DSN", "NAME", vbNullString)
''    OpenConnection sProdDSN
''    Set RS = New ADODB.Recordset
''
''    sDBName = gConn.DefaultDatabase
''
''    MyBat = pvBatches
''    'Get the ibnumber prefix (Excludes the Rebill and or supplement suffix)
''    sIBNumberPF = MyBat.sIBNumber
''    lSFPos = InStr(6, sIBNumberPF, ".")
''    If lSFPos > 6 Then
''        'Check for rebill suffix
''        lRSFPos = InStrRev(sIBNumberPF, "R", lSFPos)
''        'Don't want to count the Adjuster initial R as a rebill
''        If lRSFPos <= 6 Then
''            lRSFPos = 0
''        End If
''        'Check for Supplement suffix
''        lSSFPos = InStrRev(sIBNumberPF, "S", lSFPos)
''        'Don't want to count the Adjuster initial S as a Supplement
''        If lSSFPos <= 6 Then
''            lSSFPos = 0
''        End If
''    End If
''    If lRSFPos + lSSFPos >= 6 Then
''        sIBNumberPF = Left(sIBNumberPF, lRSFPos + lSSFPos - 1)
''    End If
''
''    'Issue 244  9.10.2002 Production Report should indicate Status Deleted
''    If pbDeleted Then
''        With MyBat
''            sSQL = "UPDATE Assignments SET "
''            sSQL = sSQL & "Assignments.Status = " & S_z & .sStatus & z_S
''            sSQL = sSQL & "Assignments.Updated = " & S_z & Now() & S_z & " "
''            sSQL = sSQL & "WHERE Assignments.ClaimNoSaln = " & S_z & .sClaimNumber & S_z & " "
''            sSQL = sSQL & "AND Assignments.IBNumber = " & S_z & sIBNumberPF & S_z & " "
''        End With
''        sSQL = CleanSQL(sSQL)
''        gConn.Execute sSQL
''        Exit Sub 'Since we are just updating Deleted Status Bail here.
''    End If
''
''    '1. list all same "duplicate" claim records that do not have an IB yet (IB = "")
''    sSQL = "SELECT Assignments.ID "
''    sSQL = sSQL & "FROM Assignments "
''    sSQL = sSQL & "WHERE LTrim(RTrim(Assignments.ClaimNoSaln)) = " & S_z & Trim(MyBat.sClaimNumber) & S_z & " "
''    '10.22.2002 Added Check for Null string as well
''    sSQL = sSQL & "AND (Assignments.IBNumber Is Null OR RTrim(Assignments.IBNumber) = ') "
''    sSQL = sSQL & "ORDER BY Assignments.ID "
''
''    sSQL = CleanSQL(sSQL)
''    RS.Open sSQL, gConn, adOpenForwardOnly, adLockReadOnly
''    If Not RS.EOF Then
''        RS.MoveFirst
''        lID = IIf(IsNull(RS!ID), 0, RS!ID)
''    End If
''    RS.Close
''
''    '2. update IB field for the very first record in that group with a flag "-1"
''    If lID > 0 Then
''        sSQL = "UPDATE Assignments SET "
''        sSQL = sSQL & "Assignments.IBNumber = " & S_z & "-1" & S_z & " "
''        sSQL = sSQL & "WHERE Assignments.ClaimNoSaln = " & S_z & MyBat.sClaimNumber & S_z & " "
''        sSQL = sSQL & "AND Assignments.ID = " & lID & " "
''
''        sSQL = CleanSQL(sSQL)
''        gConn.Execute sSQL
''    End If
''
''    '3. When updating the production table with adjuster info,
''    '   first try to update the record where Claimno and IB = the adjuster info.
''    With MyBat
''        cRTTotalFee = poRTUL.GetRTTotalFee(.sClaimNumber, sIBNumberPF, poBatch, poUL)
''        sSQL = "UPDATE Assignments SET "
''        sSQL = sSQL & "Assignments.Adjuster = " & S_z & .sAdjuster_I & z_S
''        sSQL = sSQL & "Assignments.TypeOfLoss = " & S_z & .sTypeOfLoss & z_S
''        sSQL = sSQL & "Assignments.PolicyNo = " & S_z & .sPolicyNumber & z_S
''        sSQL = sSQL & "Assignments.CatCode = " & S_z & .sCatCode & z_S
''        sSQL = sSQL & "Assignments.Insured = " & S_z & .sInsuredName & z_S
''        sSQL = sSQL & "Assignments.Address = " & S_z & .sMailAddress & z_S
''        sSQL = sSQL & "Assignments.PropertyAddress = " & S_z & .sLossLocation & z_S
''        sSQL = sSQL & "Assignments.MortgageeName = " & S_z & .sMortgageeName & z_S
''        sSQL = sSQL & "Assignments.BuildingLimits = " & .cBuildingLimits & ", "
''        sSQL = sSQL & "Assignments.ContentsLimits = " & .cContentsLimits & ", "
''        sSQL = sSQL & "Assignments.Deductibles = " & S_z & .sDeductibles & z_S
''        sSQL = sSQL & "Assignments.LossDate = " & S_z & .dtDateOfLoss & z_S
''        sSQL = sSQL & "Assignments.DateAssigned = " & S_z & .dtAssignedDate & z_S
''        sSQL = sSQL & "Assignments.ContactDate = " & S_z & .dtContactedDate & z_S
''        sSQL = sSQL & "Assignments.CloseDate = " & S_z & .dtDateClosed & z_S
''        sSQL = sSQL & "Assignments.GrossLoss = " & .cGrossLoss & ", "
''        sSQL = sSQL & "Assignments.TotalFee = " & cRTTotalFee & ", "
''        sSQL = sSQL & "Assignments.Status = " & S_z & .sStatus & z_S
''        sSQL = sSQL & "Assignments.Updated = " & S_z & Now() & S_z & " "
''        sSQL = sSQL & "WHERE Assignments.ClaimNoSaln = " & S_z & .sClaimNumber & S_z & " "
''        sSQL = sSQL & "AND Assignments.IBNumber = " & S_z & sIBNumberPF & S_z & " "
''    End With
''
''    sSQL = CleanSQL(sSQL)
''    gConn.Execute sSQL, lRecordsAffected
''    '   If no records are afected, then update the IB field to be the adjuster
''    '   IB where claimno = Adjuster claimno and IB = -1.
''    If lRecordsAffected = 0 Then
''        With MyBat
''            sSQL = "UPDATE Assignments SET "
''            sSQL = sSQL & "Assignments.IBNumber = " & S_z & sIBNumberPF & z_S
''            sSQL = sSQL & "Assignments.Adjuster = " & S_z & .sAdjuster_I & z_S
''            sSQL = sSQL & "Assignments.TypeOfLoss = " & S_z & .sTypeOfLoss & z_S
''            sSQL = sSQL & "Assignments.PolicyNo = " & S_z & .sPolicyNumber & z_S
''            sSQL = sSQL & "Assignments.CatCode = " & S_z & .sCatCode & z_S
''            sSQL = sSQL & "Assignments.Insured = " & S_z & .sInsuredName & z_S
''            sSQL = sSQL & "Assignments.Address = " & S_z & .sMailAddress & z_S
''            sSQL = sSQL & "Assignments.PropertyAddress = " & S_z & .sLossLocation & z_S
''            sSQL = sSQL & "Assignments.MortgageeName = " & S_z & .sMortgageeName & z_S
''            sSQL = sSQL & "Assignments.BuildingLimits = " & .cBuildingLimits & ", "
''            sSQL = sSQL & "Assignments.ContentsLimits = " & .cContentsLimits & ", "
''            sSQL = sSQL & "Assignments.Deductibles = " & S_z & .sDeductibles & z_S
''            sSQL = sSQL & "Assignments.LossDate = " & S_z & .dtDateOfLoss & z_S
''            sSQL = sSQL & "Assignments.DateAssigned = " & S_z & .dtAssignedDate & z_S
''            sSQL = sSQL & "Assignments.ContactDate = " & S_z & .dtContactedDate & z_S
''            sSQL = sSQL & "Assignments.CloseDate = " & S_z & .dtDateClosed & z_S
''            sSQL = sSQL & "Assignments.GrossLoss = " & .cGrossLoss & ", "
''            sSQL = sSQL & "Assignments.TotalFee = " & cRTTotalFee & ", "
''            sSQL = sSQL & "Assignments.Status = " & S_z & .sStatus & z_S
''            sSQL = sSQL & "Assignments.Updated = " & S_z & Now() & S_z & " "
''            sSQL = sSQL & "WHERE Assignments.ClaimNoSaln = " & S_z & .sClaimNumber & S_z & " "
''            sSQL = sSQL & "AND Assignments.IBNumber = " & S_z & "-1" & S_z & " "
''        End With
''
''        sSQL = CleanSQL(sSQL)
''        gConn.Execute sSQL, lRecordsAffected
''    End If
''
''    '4. If still no records are affect then do an insert into the Assignments table
''    '   with the Adjuster info
''    If lRecordsAffected = 0 Then
''        sSQL = "INSERT INTO ASSIGNMENTS (Client, "      '01Client
''        sSQL = sSQL & "ClaimNoSaln, "                   '02ClaimNoSaln
''        sSQL = sSQL & "IBNumber, "                      '02aIBNumber
''        sSQL = sSQL & "DateAssigned, "        'Date/Time'03DateAssigned
''        sSQL = sSQL & "Adjuster, "                      '04Adjuster
''        sSQL = sSQL & "Company, "                       '05Company
''        sSQL = sSQL & "TypeOfLoss, "                    '06TypeOfLoss
''        sSQL = sSQL & "StateNo, "                       '07StateNo
''        sSQL = sSQL & "PolicyNo, "                      '08PolicyNo
''        sSQL = sSQL & "TexasSuffix, "                   '09TexasSuffix
''        sSQL = sSQL & "CatCode, "                       '10CatCode
''        sSQL = sSQL & "PolicyDescription, "             '11PolicyDescription
''        sSQL = sSQL & "Insured, "                       '12Insured
''        sSQL = sSQL & "Address, "                       '13Address
''        sSQL = sSQL & "HomePhone, "                     '14HomePhone
''        sSQL = sSQL & "BusinessPhone, "                 '14aBusinessPhone
''        sSQL = sSQL & "PropertyAddress, "               '15PropertyAddress
''        sSQL = sSQL & "MortgageeName, "                 '16MortgageeName
''        sSQL = sSQL & "LoanNo, "                        '17LoanNo
''        sSQL = sSQL & "MtgAddress, "                    '18MtgAddress
''        sSQL = sSQL & "MtgCode, "                       '19MtgCode
''        sSQL = sSQL & "AgentLR, "                       '20AgentLR
''        sSQL = sSQL & "StateCD, "                       '21StateCD
''        sSQL = sSQL & "District, "                      '22District
''        sSQL = sSQL & "AgentNo, "                       '23AgentNo
''        sSQL = sSQL & "ReportedBy, "                    '24ReportedBy
''        sSQL = sSQL & "ReportedByPhone, "               '25ReportedByPhone
''        sSQL = sSQL & "DateReportedToAgent, " 'Date/Time'26DateReportedToAgent
''        sSQL = sSQL & "DateReportedByAgent, " 'Date/Time'27DateReportedByAgent
''        sSQL = sSQL & "LossDate, "            'Date/Time'28LossDate
''        sSQL = sSQL & "LossLocation, "                  '29LossLocation
''        sSQL = sSQL & "BalanceDue, "                    '30BalanceDue
''        sSQL = sSQL & "MFRec, "                         '31MFRec
''        sSQL = sSQL & "RenewalDate, "         'Date/Time'32RenewalDate
''        sSQL = sSQL & "NewBusReinDt, "        'Date/Time'33NewBusReinDt
''        sSQL = sSQL & "BuildingLimits, "            'CUR'34BuildingLimits
''        sSQL = sSQL & "ContentsLimits, "            'CUR'35ContentsLimits
''        sSQL = sSQL & "Deductibles, "                   '36Deductibles
''        sSQL = sSQL & "Format, "                        '37Format
''        sSQL = sSQL & "LossReport, "                'Memo'38LossReport
''        sSQL = sSQL & "ContactDate, "          'Date/Time'39ContactDate
''        sSQL = sSQL & "CloseDate, "            'Date/Time'40CloseDate
''        sSQL = sSQL & "GrossLoss, "                  'Cur'41GrossLoss
''        sSQL = sSQL & "TotalFee, "                  'Cur'42TotalFee
''        sSQL = sSQL & "Status, "                    'Status
''        sSQL = sSQL & "Updated ) "                  'Updated Time
''        With MyBat
''            'Use Policy Insured if there
''            sSQL = sSQL & "VALUES (" & S_z & .sInsuredName & z_S  '01Client
''            sSQL = sSQL & S_z & .sClaimNumber & z_S '02ClaimNoSaln
''            sSQL = sSQL & S_z & sIBNumberPF & z_S    '02aIBNumber
''            'BGS Use DT Assigned 9.11.2002
''            sSQL = sSQL & IIf(IsDate(.dtAssignedDate), DT_z & .dtAssignedDate & DT_z, "Null") & ", "    'Date/Time'03DateAssigned
''            sSQL = sSQL & S_z & .sAdjuster_I & z_S          '04Adjuster
''            sSQL = sSQL & S_z & vbNullString & z_S          '05Company
''            sSQL = sSQL & S_z & .sTypeOfLoss & z_S           '06TypeOfLoss
''            sSQL = sSQL & S_z & vbNullString & z_S          '07StateNo
''            sSQL = sSQL & S_z & .sPolicyNumber & z_S           '08PolicyNo
''            sSQL = sSQL & S_z & vbNullString & z_S          '09TexasSuffix
''            sSQL = sSQL & S_z & .sCatCode & z_S    '10CatCode
''            sSQL = sSQL & S_z & vbNullString & z_S          '11PolicyDescription
''            sSQL = sSQL & S_z & .sInsuredName & z_S         '12Insured
''            sSQL = sSQL & S_z & .sMailAddress & z_S         '13Address
''            sSQL = sSQL & S_z & vbNullString & z_S          '14HomePhone
''            sSQL = sSQL & S_z & vbNullString & z_S          '14aBusinessPhone
''            sSQL = sSQL & S_z & .sLossLocation & z_S           '15PropertyAddress
''            sSQL = sSQL & S_z & .sMortgageeName & z_S           '16MortgageeName
''            sSQL = sSQL & S_z & vbNullString & z_S          '17LoanNo
''            sSQL = sSQL & S_z & vbNullString & z_S          '18MtgAddress
''            sSQL = sSQL & S_z & vbNullString & z_S          '19MtgCode
''            sSQL = sSQL & S_z & vbNullString & z_S          '20AgentLR
''            sSQL = sSQL & S_z & vbNullString & z_S          '21StateCD
''            sSQL = sSQL & S_z & vbNullString & z_S          '22District
''            sSQL = sSQL & S_z & vbNullString & z_S          '23AgentNo
''            sSQL = sSQL & S_z & vbNullString & z_S          '24ReportedBy
''            sSQL = sSQL & S_z & vbNullString & z_S          '25ReportedByPhone
''            sSQL = sSQL & "Null" & ", "                     '26DateReportedToAgent
''            sSQL = sSQL & "Null" & ", "                     '27DateReportedByAgent
''            sSQL = sSQL & IIf(IsDate(.dtDateOfLoss), DT_z & .dtDateOfLoss & DT_z, "Null") & ", "  'Date/Time'28LossDate
''            sSQL = sSQL & S_z & vbNullString & z_S          '29LossLocation
''            sSQL = sSQL & S_z & vbNullString & z_S          '30BalanceDue
''            sSQL = sSQL & S_z & vbNullString & z_S          '31MFRec
''            sSQL = sSQL & "Null" & ", "                     '32RenewalDate
''            sSQL = sSQL & "Null" & ", "                     '33NewBusReinDt
''            sSQL = sSQL & .cBuildingLimits & ", "           '34BuildingLimits
''            sSQL = sSQL & .cContentsLimits & ", "           '35ContentsLimits
''            sSQL = sSQL & S_z & .sDeductibles & z_S           '36Deductibles
''            sSQL = sSQL & S_z & .sCarrierCode & z_S         '37 Format
''            sSQL = sSQL & S_z & vbNullString & z_S          '38 LossReport
''            sSQL = sSQL & IIf(IsDate(.dtContactedDate), DT_z & .dtContactedDate & DT_z, "Null") & ", "    'Date/Time'39ContactDate
''            sSQL = sSQL & IIf(IsDate(.dtDateClosed), DT_z & .dtDateClosed & DT_z, "Null") & ", "  'Date/Time'40CloseDate
''            sSQL = sSQL & .cGrossLoss & ", "                      'Cur'41GrossLoss
''            sSQL = sSQL & cRTTotalFee & " , "                     'Cur'42TotalFee
''            sSQL = sSQL & S_z & .sStatus & z_S                  'Status
''            sSQL = sSQL & S_z & Now() & S_z & " ) "         'Updated time
''
''        End With
''
''        sSQL = CleanSQL(sSQL)
''        gConn.Execute sSQL, lRecordsAffected
''
''        If lRecordsAffected = 0 Then
''            Err.Raise -999, , "No Records Affected."
''        End If
''    End If
''
''    'CleanUp
''    Set RS = Nothing
''    Exit Sub
''EH:
''    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
''    sMess = sMess & "Private Sub UpdateDBRT_SQLServer" & vbCrLf
''    sMess = sMess & "ERROR # " & Err.Number & " " & Now & vbCrLf
''    sMess = sMess & Err.Description & vbCrLf
''    sMess = sMess & "DB: " & sDBName & vbCrLf
''    sMess = sMess & "Could not update SALN '" & MyBat.sClaimNumber & "' " & vbCrLf
''    sMess = sMess & "Adjuster: " & MyBat.sAdjuster_N & vbCrLf
''    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
''    moUL_ErrorMess sMess
''    Err.Clear
''    Resume Next
'End Sub

Private Sub Timer_Status_Timer()
    On Error GoTo EH
    Dim iStatus As Integer
    Dim sUpdate As String
    Dim bEnabled As Boolean
    Dim sMsg As String
    Dim sMess As String
    Dim sFLSDelTime As String
    
    'Check for Commands sent to the Registry
    sMsg = GetSetting(App.EXEName, "MSG", "COMMAND", vbNullString)
    
    Select Case sMsg
        Case "SHUT_DOWN"
            mbSHUTDOWN = True
    End Select
   
    
    'Update Status Here
    On Error Resume Next
    iStatus = GetSetting(App.EXEName, "Msg", "Status", 1)
    If Err.Number > 0 Then
        Err.Clear
        iStatus = 1
        SaveSetting App.EXEName, "Msg", "Status", iStatus
    End If
    On Error GoTo EH
'    If iStatus > 0 Then
        Select Case iStatus
            Case PicList.Disabled
                If m_NID.hIcon <> imgList.ListImages(PicList.Disabled).Picture Then
                    m_NID.hIcon = imgList.ListImages(PicList.Disabled).Picture
                    Timer_Import.Enabled = False
                    mPop(2).Checked = True
'                    Me.Caption = "Process User Tokin (" & msUserName & ") - OFF"
                    Me.Icon = imgList.ListImages(PicList.Disabled).Picture
                    m_NID.szTip = Me.Caption & vbNullChar
                    Call ShellNotifyIcon(NIM_MODIFY, m_NID)
                End If
            Case PicList.Idle
                If m_NID.hIcon <> imgList.ListImages(PicList.Idle).Picture Then
                    m_NID.hIcon = imgList.ListImages(PicList.Idle).Picture
                    Timer_Import.Enabled = True
                    mPop(2).Checked = False
'                    Me.Caption = "Process User Tokin (" & msUserName & ") - ON"
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
                ProcessTokens
            Case PicList.Busy
                ProcessTokens
        End Select
    
    
    If mbSHUTDOWN Then
        If Not mbImporting And Not mbExporting Then
            Unload Me
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
    Timer_Import.Enabled = False
    Timer_Status.Enabled = False
    
    m_NID.hIcon = imgList.ListImages(PicList.Busy).Picture
'    SaveSetting "V2AutoImport", "Msg", "Status", PicList.Busy
    Me.Icon = imgList.ListImages(PicList.Busy).Picture
'    Me.Caption = "Process User Tokin (" & msUserName & ") - Busy"
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
    If CLng(GetSetting("V2AutoImport", "Msg", "Status", 0)) <> PicList.Busy Then
        Timer_Import.Enabled = True
        Timer_Status.Enabled = True
        Exit Sub
    End If
    
    'After import Update icons and reenable stuff
    m_NID.hIcon = imgList.ListImages(PicList.Idle).Picture
'    SaveSetting "V2AutoImport", "Msg", "Status", PicList.Idle
    Me.Icon = imgList.ListImages(PicList.Idle).Picture
'    Me.Caption = "Process User Tokin (" & msUserName & ") - ON"
    m_NID.szTip = Me.Caption & vbNullChar
    Call ShellNotifyIcon(NIM_MODIFY, m_NID)
    
    ProgBarLoss.Value = 0
    
    Timer_Import.Enabled = True
    Timer_Status.Enabled = True
    
    Exit Sub
EH:
    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
    sMess = sMess & "Private Sub ShowIdleIcon" & vbCrLf
    sMess = sMess & "ERROR # " & Err.Number & " " & Now & vbCrLf
    sMess = sMess & Err.Description & vbCrLf
    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
    moUL_ErrorMess sMess
End Sub



Private Sub ProcessSecurityToken_SQLServer(pToken As TokenInfo)
    On Error GoTo EH
    Dim sMess As String
    Dim RS As ADODB.Recordset
    Dim CheckDupRS As ADODB.Recordset
    Dim sSQL As String
    Dim vSecToken As Variant
    'Will be passed back to .tokout
    Dim bSSN As Boolean
    Dim bUserName As Boolean
    Dim bPass As Boolean
    Dim bOldPass As Boolean
    Dim bResetPass As Boolean
    Dim lLicDaysLeft As Long
    Dim bResetLic As Boolean
    Dim bSSNIsZero As Boolean 'If SSN was manually entered by V2WebControl
    Dim lPermissionErrorCount As Long
    '1.16.2003 Added the IBPrefix
    Dim sIBPrefix As String
    Dim sUserName As String
    Dim lSSN As Long
    Dim sNewTokOutPath As String
    Dim sProdDSN As String
    Dim lRecordsAffected As Long
    '6.20.2003 Fail Logon Attempts For those Who do not Have the Latest Version.
    Dim sLatestVS As String
    Dim bInvalidVS As Boolean
    Dim lUsersID As Long
    Dim sPassOnServer As String
    Dim sPassOnEasyClaim As String
    
    
    sProdDSN = GetSetting("V2WebControl", "DSN", "NAME", vbNullString)
    OpenConnection sProdDSN
    Set RS = New ADODB.Recordset
    Set CheckDupRS = New ADODB.Recordset
    
    vSecToken = Split(pToken.sToken, vbCrLf)
    'need to decrypt some strings
    vSecToken(SecurityToken.UserName) = goUtil.Decode(CStr(vSecToken(SecurityToken.UserName)))
    sUserName = vSecToken(SecurityToken.UserName)
    vSecToken(SecurityToken.SSN) = goUtil.Decode(CStr(vSecToken(SecurityToken.SSN)))
    lSSN = vSecToken(SecurityToken.SSN)
       
    vSecToken(SecurityToken.Pass) = goUtil.Decode(CStr(vSecToken(SecurityToken.Pass)))
    sPassOnEasyClaim = vSecToken(SecurityToken.Pass)
    'This will be Nullstring on very first Upload
    If vSecToken(SecurityToken.OldPass) <> vbNullString Then
        vSecToken(SecurityToken.OldPass) = goUtil.Decode(CStr(vSecToken(SecurityToken.OldPass)))
    End If
    vSecToken(SecurityToken.LicDaysLeft) = goUtil.Decode(CStr(vSecToken(SecurityToken.LicDaysLeft)))
       
    '10.2.2002 If the Adjuster uses a numeric CRID Skip to Fail this logon
    'This is the Opposite to how VS1 worked
    If IsNumeric(vSecToken(SecurityToken.UserName)) Then
        GoTo FAIL_LOGON
    End If
           
    'If we found the UserName but not SSN...
    'Need to check SSN against the UserName.  If the SSN is = 0 then we need to update the
    'SSN.  This needs to be done for those Records that have been manullay entered using
    'V2WebControl.  The CRID may be in there without the SSN set up.
    sSQL = "SELECT U.UserName, U.UsersID, U.SSN FROM Users U "
    sSQL = sSQL & "WHERE Upper(U.UserName) = '" & UCase(vSecToken(SecurityToken.UserName)) & "' "
    
    RS.Open sSQL, gConn, adOpenForwardOnly, adLockReadOnly
    
    If Not RS.EOF Then
        bUserName = True
        'Set the Memebr variables for UserName and UsersID
        msUserName = IIf(IsNull(RS!UserName), vbNullString, RS!UserName)
        msUsersID = IIf(IsNull(RS!UsersID), 0, CStr(RS!UsersID))
        Do Until RS.EOF
            If IIf(IsNull(RS!SSN), 0, RS!SSN) = 0 Then
                bSSNIsZero = True
                sSQL = "UPDATE Users SET Users.SSN = " & lSSN & ", "
                sSQL = sSQL & "Users.Comments = 'Easy Claim Update SSN', "
                sSQL = sSQL & "Users.DateLastUpdated = GetDate(), "
                sSQL = sSQL & "Users.UpdateByUserID = UsersID "
                sSQL = sSQL & "WHERE Users.UsersID = " & RS!UsersID & " "
                gConn.Execute sSQL, lRecordsAffected
                If lRecordsAffected > 0 Then
                    bUserName = True
                End If
            End If
            RS.MoveNext
        Loop
    End If
    
    RS.Close
    
    'Process the Token
    Select Case bUserName
        'UserName found on the Server
        Case True
            sSQL = "SELECT U.UsersID, "
            sSQL = sSQL & "U.UserName, "
            sSQL = sSQL & "U.FirstName, "
            sSQL = sSQL & "U.LastName, "
            sSQL = sSQL & "U.EMail, "
            sSQL = sSQL & "U.SSN, "
            sSQL = sSQL & "U.Password, "
            sSQL = sSQL & "U.ContactPhone, "
            sSQL = sSQL & "US.VersionInfo, "
            sSQL = sSQL & "US.LicenseDaysLeft, "
            sSQL = sSQL & "US.ResetLicense, "
            sSQL = sSQL & "US.IBPrefix, "
            sSQL = sSQL & "US.ResetIBPrefix, "
            sSQL = sSQL & "US.SingleFileSendAuthority "
            sSQL = sSQL & "FROM Users U INNER JOIN AdjusterUsersSoftware US ON U.UsersID = US.UsersID "
            sSQL = sSQL & "WHERE Upper(U.UserName) = '" & UCase(vSecToken(SecurityToken.UserName)) & "' "
            RS.Open sSQL, gConn, adOpenForwardOnly, adLockReadOnly

            If Not RS.EOF Then
                'Set the UsersID
                lUsersID = RS!UsersID
                'If we found UserName match up
                'then we can process from here
                'Check for Password reset by user
                If CBool(vSecToken(SecurityToken.ResetPass)) Then
                    If bSSNIsZero Or IsNull(RS!Password) Then
                        GoTo RESET_PASS
                    End If
                    '1.16.2003
                    If Trim(RS!Password) = vbNullString Then
                        GoTo RESET_PASS
                    End If
                    'Check to see if Password matches Old Password
                    If goUtil.Decode(CStr(RS!Password)) = vSecToken(SecurityToken.OldPass) Then
RESET_PASS:
                        'If it matches then we need to update all passwords for SSN
                        sSQL = "UPDATE Users SET Users.Password = '" & CleanString(goUtil.Encode(CStr(vSecToken(SecurityToken.Pass)))) & "', "
                        sSQL = sSQL & "Users.Comments = 'Easy Claim Changed Password', "
                        sSQL = sSQL & "Users.DateLastUpdated = GetDate(), "
                        sSQL = sSQL & "Users.UpdateByUserID = UsersID "
                        sSQL = sSQL & "WHERE Users.UsersID = " & RS!UsersID & " "
                        gConn.Execute sSQL
                        bPass = True
                        sPassOnServer = CStr(vSecToken(SecurityToken.Pass))
                        bOldPass = True
                    Else
                        'old password failed
                        'This should never happen unless someone on the server changes the
                        'password.  All password changes should be done by user only.
                        'If the user forgets their password they just won't be able to change it  unless they
                        'call server tech support to look up their password.
                        'Also we can get here if user changes there SSN
                        If goUtil.Decode(CStr(RS!Password)) = vSecToken(SecurityToken.Pass) Then
                            GoTo UPDATE_PASS
                        Else
                            bPass = False
                            bOldPass = False
                        End If
                    End If
                Else
                    'Check to see if Password matches Old Password
                    If bSSNIsZero Or IsNull(RS!Password) Then
                        GoTo UPDATE_PASS
                    End If
                    '1.16.2003
                    If Trim(RS!Password) = vbNullString Then
                        GoTo UPDATE_PASS
                    End If
                    sPassOnServer = goUtil.Decode(CStr(RS!Password))
                    If UCase(sPassOnServer) = UCase(vSecToken(SecurityToken.Pass)) Then
                        bPass = True
UPDATE_PASS:
                        
                        If Not bPass Then
                            sSQL = "UPDATE Users SET Users.Password = '" & CleanString(goUtil.Encode(CStr(vSecToken(SecurityToken.Pass)))) & "', "
                            sSQL = sSQL & "Users.Comments = 'Easy Claim Update PassWord', "
                            sSQL = sSQL & "Users.DateLastUpdated = GetDate(), "
                            sSQL = sSQL & "Users.UpdateByUserID = UsersID "
                            sSQL = sSQL & "WHERE Users.UsersID = " & RS!UsersID & " "
                            gConn.Execute sSQL, lRecordsAffected
                        End If
                        bPass = True
                        sPassOnServer = CStr(vSecToken(SecurityToken.Pass))
                        bOldPass = True
                    Else
                        'password failed
                        'This should never happen unless someone on the server changes the
                        'password.  All password changes should be done by user only.
                        'If the user forgets their password they just won't be able to change it  unless they
                        'call server tech support to look up their password.
                        bPass = False
                        bOldPass = False
                    End If
                End If
                
                If bPass Then
                    'If the password was verified then we need to update LicDaysLeft on the server
                    'unless there is * reset flag.  In this instance we will get rid of the * flag on the server
                    'and pass back to .tokout the new Lic amount.
                    sSQL = "SELECT US.UsersID, "
                    sSQL = sSQL & "US.LicenseDaysLeft, "
                    sSQL = sSQL & "US.ResetLicense, "
                    sSQL = sSQL & "US.IBPrefix, "
                    sSQL = sSQL & "US.ResetIBPrefix, "
                    sSQL = sSQL & "US.SingleFileSendAuthority "
                    sSQL = sSQL & "FROM AdjusterUsersSoftware US "
                    sSQL = sSQL & "WHERE US.UsersID = " & lUsersID & " "
                    RS.Close
                    
                    RS.Open sSQL, gConn, adOpenForwardOnly, adLockReadOnly
                    If Not RS.EOF Then
                        RS.MoveFirst
                        If CBool(RS!ResetLicense) Then
                            lLicDaysLeft = RS!LicenseDaysLeft
                        Else
                            lLicDaysLeft = CLng(vSecToken(SecurityToken.LicDaysLeft))
                        End If
                        '1.17.2003 Allow for backwards Compat(before added IBPrefix check)
                        If UBound(vSecToken, 1) >= SecurityToken.IBPrefix Then
                            'Do the same for the IB Prefix
                            If CBool(RS!ResetIBPrefix) Then
                                'Don't Verify if this is Dup IBPrefix, since the prefix is being
                                'forced , we must know already that it won't be a duplicate.
                                sIBPrefix = IIf(IsNull(RS!IBPrefix), vbNullString, RS!IBPrefix)
                            Else
                                sIBPrefix = vSecToken(SecurityToken.IBPrefix)
                                '1.17.2003 Change the IBPrefix to next avail prefix if this one
                                'Happens to be used by someone else Use the SQL Server Function
                                sSQL = "SELECT  dbo.VerifyNotDupIBPrefix('" & sIBPrefix & "', " & lUsersID & ") As IBPREFIX "
                                CheckDupRS.Open sSQL, gConn, adOpenForwardOnly, adLockReadOnly
                                If Not CheckDupRS.EOF Then
                                    CheckDupRS.MoveFirst
                                    If Not IsNull(CheckDupRS!IBPrefix) Then
                                        sIBPrefix = CheckDupRS!IBPrefix
                                    End If
                                End If
                                CheckDupRS.Close
                            End If
                        End If
                        
                        sSQL = "UPDATE AdjusterUsersSoftware SET VersionInfo = '" & Replace(vSecToken(SecurityToken.AppVSInfo), F_VBCRLF, vbCrLf) & "', "
                        sSQL = sSQL & "LicenseDaysLeft = " & lLicDaysLeft & ", "
                        sSQL = sSQL & "ResetLicense = 0, "
                        sSQL = sSQL & "IBPrefix = '" & CleanString(sIBPrefix) & "', "
                        sSQL = sSQL & "ResetIBPrefix = 0, "
                        sSQL = sSQL & "DateLastUpdated = GetDate(), "
                        sSQL = sSQL & "UpdateByUserID = UsersID "
                        sSQL = sSQL & "WHERE UsersID = " & lUsersID & " "
                        gConn.Execute sSQL, lRecordsAffected
                        
                        
                        sSQL = "UPDATE AdjusterUsersUpdates SET FirstName = '" & vSecToken(SecurityToken.FName) & "', "
                        sSQL = sSQL & "LastName = '" & vSecToken(SecurityToken.LName) & "', "
                        sSQL = sSQL & "SSN = " & lSSN & ", "
                        sSQL = sSQL & "Email = '" & vSecToken(SecurityToken.Email) & "', "
                        sSQL = sSQL & "ContactPhone = '" & vSecToken(SecurityToken.ContactPhone) & "', "
                        sSQL = sSQL & "EmergencyPhone = '" & vSecToken(SecurityToken.sEmergencyPhone) & "', "
                        sSQL = sSQL & "ADDRESS = '" & vSecToken(SecurityToken.sAddress) & "', "
                        sSQL = sSQL & "City = '" & vSecToken(SecurityToken.sCity) & "', "
                        sSQL = sSQL & "State = '" & vSecToken(SecurityToken.sState) & "', "
                        sSQL = sSQL & "Zip = " & IIf(vSecToken(SecurityToken.iZip) = vbNullString, 0, vSecToken(SecurityToken.iZip)) & ", "
                        sSQL = sSQL & "Zip4 = " & IIf(vSecToken(SecurityToken.iZip4) = vbNullString, 0, vSecToken(SecurityToken.iZip4)) & ", "
                        sSQL = sSQL & "OtherPostCode = '" & vSecToken(SecurityToken.sOtherPostCode) & "', "
                        sSQL = sSQL & "DateLastUpdated = GetDate(), "
                        sSQL = sSQL & "UpdateByUserID = UsersID "
                        sSQL = sSQL & "WHERE UsersID = " & lUsersID & " "
                        gConn.Execute sSQL, lRecordsAffected
                        
                        'Also Record these changes ON Users Table
                        sSQL = "UPDATE Users Set Comments = 'Easy Claim Update<BR>Adjuster Users Software<BR>Adjuster Users Updates', "
                        sSQL = sSQL & "DateLastUpdated = GetDate(), "
                        sSQL = sSQL & "UpdateByUserID = UsersID "
                        sSQL = sSQL & "WHERE UsersID = " & lUsersID & " "
                        gConn.Execute sSQL, lRecordsAffected
                    End If
                Else
                    RS.Close
                    GoTo FAIL_LOGON
                End If
            Else
                RS.Close
                GoTo FAIL_LOGON
            End If
            
            RS.Close
        Case False
FAIL_LOGON:
            'Can't assign this UserName
            bUserName = False
            '10.2.2002 also log this information to help us track failed logon attempts
            If sLatestVS <> vbNullString And bInvalidVS Then
                sMess = sMess & "<---INVALID VERSION!--->" & vbCrLf
            End If
            sMess = "<<<<<<<<<< FAILED LOGON ATTEMPT >>>>>>>>>>" & vbCrLf
            sMess = sMess & "Private Sub ProcessSecurityToken_SQLServer" & vbCrLf
            sMess = sMess & "CARRIER: (" & vSecToken(SecurityToken.Carrier) & ")" & vbCrLf
            sMess = sMess & "USER NAME: (" & vSecToken(SecurityToken.UserName) & ")" & vbCrLf
            sMess = sMess & "PASSWORD_ON_SERVER: (" & sPassOnServer & ")" & vbCrLf
            sMess = sMess & "PASSWORD_ON_EasyClaim: (" & sPassOnEasyClaim & ")" & vbCrLf
            sMess = sMess & "SSN: (" & vSecToken(SecurityToken.SSN) & ")" & vbCrLf
            sMess = sMess & "FIRST NAME: (" & vSecToken(SecurityToken.FName) & ")" & vbCrLf
            sMess = sMess & "LAST NAME: (" & vSecToken(SecurityToken.LName) & ")" & vbCrLf
            If UBound(vSecToken, 1) >= SecurityToken.LName + 1 Then
                sMess = sMess & "IB_PREFIX: (" & vSecToken(SecurityToken.LName + 1) & ")" & vbCrLf
            End If
            sMess = sMess & "E-MAIL: (" & vSecToken(SecurityToken.Email) & ")" & vbCrLf
            sMess = sMess & "CONTACT PHONE: (" & vSecToken(SecurityToken.ContactPhone) & ")" & vbCrLf
            sMess = sMess & "TEAM LEADER: (" & vSecToken(SecurityToken.TeamLeader) & ")" & vbCrLf
            sMess = sMess & "LICENSE DAYS LEFT: (" & vSecToken(SecurityToken.LicDaysLeft) & ")" & vbCrLf
            sMess = sMess & "APPLICATION VERSION INFO: " & vbCrLf
            sMess = sMess & Replace(vSecToken(SecurityToken.AppVSInfo), F_VBCRLF, vbCrLf) & vbCrLf
            If sLatestVS <> vbNullString And bInvalidVS Then
                sMess = sMess & "<---INVALID VERSION!--->" & vbCrLf
            End If
            sMess = sMess & "<<<<<<<<<< END FAILED LOGON ATTEMPT >>>>>>>>>>" & vbCrLf & vbCrLf
            moUL_ErrorMess sMess
    End Select
    
    'once were done processing the token we need to save it to .tokout path
    vSecToken(SecurityToken.SSN) = bSSN
    vSecToken(SecurityToken.UserName) = bUserName
    vSecToken(SecurityToken.Pass) = bPass
    vSecToken(SecurityToken.OldPass) = bOldPass
    vSecToken(SecurityToken.LicDaysLeft) = lLicDaysLeft
   
    If UBound(vSecToken, 1) >= SecurityToken.LName + 1 Then
        If sIBPrefix <> vbNullString Then
            vSecToken(SecurityToken.IBPrefix) = sIBPrefix
        End If
    End If
    '6.20.2003 Check for Invalid Version Info
    If sLatestVS <> vbNullString And bInvalidVS Then
        pToken.sToken = "<---INVALID VERSION! YOU NEED " & sLatestVS & " --->"
        mbSHUTDOWN = True
    Else
        pToken.sToken = Join(vSecToken, vbCrLf)
    End If
    
    'If the UserName\Password Checks Out Start creating Down Load Files
    msMyTokoutPath = Replace(pToken.sPath, ".tokin", ".tokout", , , vbTextCompare)
    msMyTokenData = pToken.sToken
    If bUserName And bPass Then
        'If nothing left in the License then don't process Download Files
        If lLicDaysLeft > 0 Then
            'Save the Valid Token to let client know they Pass muster
            'and to start checking Database Version
            goUtil.utSaveFileData msMyTokoutPath, msMyTokenData
            ExportDL msMyTokoutPath, msMyTokenData
            GoTo CLEANUP
        Else
            mbSHUTDOWN = True
        End If
    Else
        mbSHUTDOWN = True
    End If
    
    lPermissionErrorCount = 0
    
    SaveFileData Replace(pToken.sPath, ".tokin", ".tokout"), pToken.sToken
    
CLEANUP:

    Set RS = Nothing
    Set CheckDupRS = Nothing
    
    Exit Sub
EH:
    'If we get a persmisiion error means we tried to retrieve
    'the file while it was still being written to disk.
    'Give it a chance to release permissions.
    If Err.Number = 70 Then
        lPermissionErrorCount = lPermissionErrorCount + 1
        Sleep 500
        If lPermissionErrorCount <= 5 Then
            Err.Clear
            Resume
        End If
    End If
    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
    sMess = sMess & "Private Sub ProcessSecurityToken_SQLServer" & vbCrLf
    sMess = sMess & "ERROR # " & Err.Number & " " & Now & vbCrLf
    sMess = sMess & Err.Description & vbCrLf
    sMess = sMess & pToken.sPath & vbCrLf
    sMess = sMess & pToken.sToken & vbCrLf
    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
    moUL_ErrorMess sMess
    pToken.sToken = "<---INVALID SECURITY TOKEN!--->"
    SaveFileData Replace(pToken.sPath, ".tokin", ".tokout"), pToken.sToken
    mbSHUTDOWN = True
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

Private Function ActiveFiles(psActiveFileDir As String, _
                             Optional psWildCard = "*.ZOT", _
                             Optional psProcessDesc As String = "Adjuster Claim file upload", _
                             Optional pbProcessTokens As Boolean = True) As Boolean
    '10.3.2002 Active files will return true for 2 different reasons
    '1. There are files in the directory and 1 or many of them are Actively being written to disk
    '2. There are Zero, that means nada, none, no files in the directory at all.
    
    On Error GoTo EH
    Dim sMess As String
    Dim colActiveFiles As Collection
    Dim vActiveFile As Variant
    Dim sActiveFile As String
    Dim lCount As Long
    Dim iFFile As Integer
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    sActiveFile = Dir(psActiveFileDir & "\" & psWildCard, vbNormal)
    
    'If there are no files to check then set to true and bail
    If sActiveFile = vbNullString Then
        ActiveFiles = True
        Exit Function
    Else
        txtMess.Text = "Analyzing " & psProcessDesc & "... " & Now()
        txtMess.Refresh
    End If
    
    'Need to do this process for at least 2.5 seconds
    mbCheckingActiveFiles = True 'set this to true so we won't process Upload again until its done
    
    For lCount = 1 To 5
        sActiveFile = Dir(psActiveFileDir & "\" & psWildCard, vbNormal)
        
        '1. Add all the existing files in the Active Files collection
        Do Until sActiveFile = vbNullString
            If colActiveFiles Is Nothing Then
                Set colActiveFiles = New Collection
            End If
            colActiveFiles.Add psActiveFileDir & "\" & sActiveFile, psActiveFileDir & "\" & sActiveFile
            sActiveFile = Dir
        Loop
        
        '2. Loop through the collection and open them read lock
        'then close file.  If there is an error while opening, the file
        'is still being written to disk, ie is "Active"
        If Not colActiveFiles Is Nothing Then
            For Each vActiveFile In colActiveFiles
                sActiveFile = vActiveFile
                iFFile = FreeFile
                Open sActiveFile For Binary Access Read Lock Read As #iFFile
                Close #iFFile
            Next
        End If
        
        Set colActiveFiles = Nothing
        'Wait half a second
        DoEvents ' This should only allow for processing of Tokens... ONLY !
        Sleep 500
        If pbProcessTokens Then
            ProcessTokens
        End If
        txtMess.Text = "Analyzing " & psProcessDesc & "... " & Now()
        txtMess.Refresh
    Next
    txtMess.Text = vbNullString
    
    Set colActiveFiles = Nothing
    mbCheckingActiveFiles = False
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    mbCheckingActiveFiles = False
    ActiveFiles = True
    Set colActiveFiles = Nothing
    If lErrNum <> 70 Then
        sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
        sMess = sMess & "Private Function ActiveFiles" & vbCrLf
        sMess = sMess & "ERROR # " & lErrNum & " " & Now & vbCrLf
        sMess = sMess & sErrDesc & vbCrLf
        sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
        moUL_ErrorMess sMess
    Else
        txtMess.Text = txtMess.Text & vbCrLf & vbCrLf & psProcessDesc & " currently in progress, waiting to process... " & Now() & vbCrLf
        txtMess.Refresh
        Sleep 1000
    End If
End Function


